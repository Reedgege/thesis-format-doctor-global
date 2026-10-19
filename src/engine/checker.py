"""Rule Engine（确定性规则引擎）：把实测文档画像与 TargetProfile 比对。

原则：
- 全部为确定性判断，不调用任何大模型 / 网络。
- 每条 Finding 标注来源（spec / school / user / ai），供对比报告裁决。
- 页面排版来自「规范」时仅作参考（info），不强制 fail；
  页面排版来自学校/问卷/AI 时硬判。参考文献始终以规范硬判。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Optional

from .docx_reader import (DocumentProfile, read_docx, DocxReadError,
                          describe_paper_size, identify_paper_size)
from .questionnaire import TargetProfile, PAGE_NUMBER_LAYOUT, origin_of_source
from .specs import (MARGIN_DIMS, get_spec, resolve_spec,
                    resolve_heading_formats, describe_heading, size_for_font)


@dataclass
class FixAction:
    """checker 产出、fixer 消费的「修正动作」。

    这是 **check→fix 单一真相源** 的载体：checker 在发现某维度需要修正时，
    把「改哪个维度、改成什么值」封进 FixAction 挂到对应 Finding 上；fixer 只消费
    带 fix 的 Finding，绝不再独立重算「要改什么」。这样「检查出来的」与「实际改的」
    必然一致，也杜绝了 fixer 用另一套规范/逻辑重算导致的漂移。

    - dimension: 与 fixer 内部动作对应的维度名（margins/font/.../reference_numbered/heading）。
    - value: 该维度的目标值（供 preview/报告展示与断言；fixer 施加时仍读 target 的同名值，
      两者恒等，故 value 仅作展示与一致性核对）。
    - note: 可选的人类可读说明（编号制条目数等）。
    """
    dimension: str
    value: Any = None
    note: str = ""

# 容差（依据：1/16 英寸内视为排版误差；字号/行距给更宽 warn 带）
# 约定：所有维度的 warn 带必须**宽于**其 fail 带（WARN < FAIL），
#       否则 warn 段几乎不可达（曾出现 MARGIN_TOL_WARN=0.05 < FAIL=0.06 的笔误）。
MARGIN_TOL_FAIL = 0.10      # 英寸
MARGIN_TOL_WARN = 0.05      # 英寸（≈1.27mm）
SIZE_TOL_FAIL = 0.5         # 磅
SIZE_TOL_WARN = 0.3
SPACING_TOL_FAIL = 0.15     # 倍
SPACING_TOL_WARN = 0.1
INDENT_TOL_FAIL = 0.12      # 英寸（正文首行缩进）
INDENT_TOL_WARN = 0.06
HEADER_TOL_FAIL = 0.12      # 英寸（页眉距页顶）
HEADER_TOL_WARN = 0.06
SPACE_TOL_FAIL = 1.5        # 磅（段后间距）
SPACE_TOL_WARN = 0.8
REF_HANG_TOL_FAIL = 0.15    # 英寸（参考文献悬挂缩进；引用维度，容差略宽）
REF_HANG_TOL_WARN = 0.03

SEVERITY_ORDER = {"pass": 0, "info": 1, "warn": 2, "fail": 3}

# Alignment labels (also reused by report.py to avoid duplicate hard-coding)
ALIGN_LABEL = {"left": "Left", "justify": "Justified",
               "center": "Center", "right": "Right"}


def _tp_expected_text(vals: dict) -> str:
    """把标题页三态期望拼成一句英文说明（进报告 expected 列）。"""
    bits = []
    if vals.get("title_page_title_centered") is not None:
        bits.append("title centered" if vals["title_page_title_centered"]
                    else "title left-aligned")
    if vals.get("title_page_title_bold") is not None:
        bits.append("title bold" if vals["title_page_title_bold"] else "title not bold")
    if vals.get("title_page_author_centered") is not None:
        bits.append("author/affiliation centered"
                    if vals["title_page_author_centered"]
                    else "author/affiliation not centered")
    return ", ".join(bits)

# 已带顺序编号的形态：[1] / (1) / 1. / 1)  —— 与 fixer 的 _LEADING_NUMBER_RE 保持一致
_HAS_LEADING_NUMBER_RE = re.compile(r"^\s*(?:\[\s*\d+\s*\]|\(\s*\d+\s*\)|\d+\s*[.)])\s")


def _has_leading_number(text: str) -> bool:
    """条目文本是否已带顺序编号前缀（[n] / (n) / n. / n)）。"""
    return bool(_HAS_LEADING_NUMBER_RE.match(text or ""))


@dataclass
class Finding:
    dimension: str
    severity: str            # pass / warn / fail / info
    message: str
    expected: str = ""
    actual: str = ""
    source: str = ""         # spec / school / user / ai
    fix: Optional[FixAction] = None   # 非 None 表示 check 认为该项需要（且可被）自动修正


# ---------------------------------------------------------------------------
# 确定性比对基元
# ---------------------------------------------------------------------------

def _cmp_margin(measured: Optional[float], expected: float) -> tuple[str, str]:
    if measured is None:
        return "warn", "Could not read (document may have no section or be protected)"
    diff = abs(measured - expected)
    # Note: only the value itself is returned here; the report layer appends
    # "Expected: ... | Actual: ..." — so do NOT prepend "Actual" here
    # (otherwise it would read "Actual: Actual 2.0\"").
    if diff <= MARGIN_TOL_WARN:
        return "pass", f'{measured}"'
    if diff <= MARGIN_TOL_FAIL:
        return "warn", f'{measured}" (deviation {diff:.2f}")'
    return "fail", f'{measured}" (deviation {diff:.2f}")'


def _cmp_num(measured: Optional[float], expected: float,
             tol_fail: float, tol_warn: float, unit: str) -> tuple[str, str]:
    if measured is None:
        return "warn", "Could not read (unable to determine this dimension from the document)"
    diff = abs(measured - expected)
    if diff <= tol_warn:
        return "pass", f"{measured}{unit}"
    if diff <= tol_fail:
        return "warn", f"{measured}{unit} (deviation {diff:.2f})"
    return "fail", f"{measured}{unit} (deviation {diff:.2f})"


def _downgrade(sev: str, origin: str, is_page_or_heading: bool,
               dim: Optional[str] = None, spec_key: Optional[str] = None) -> str:
    """规范来源的页面/标题维度仅作参考（info），不强制 fail/warn。

    例外（Task B）：当规范为 APA 且其维度属于「不可被模板推翻的硬性正文规则」时，
    即便来自 spec，也至少保留 warn 下限。这些规则在 APA 7th 中明确且不可能由学校模板
    覆盖：
      - alignment（§2.22：不得两端对齐，须左对齐）
      - first_line_indent（§2.24：每段首行缩进 0.5″）
      - heading（§2.27：标题层级≤5，且应使用标题样式）
    其余页面维度（页边距/字体/字号/行距）仍按模板驱动原则降级为 info。
    """
    if is_page_or_heading and origin == "spec" and sev in ("warn", "fail"):
        if spec_key and spec_key.startswith("APA") and dim in (
            "alignment", "first_line_indent", "heading"
        ):
            return "warn"
        return "info"
    return sev


# ---------------------------------------------------------------------------
# 参考文献检查：数据驱动（依据 ReferenceRule.checks 分发）
# 每个子检查返回 (severity, message)。
# ---------------------------------------------------------------------------

def _chk_has_year_paren(entries, style) -> tuple:
    text = "\n".join(entries)
    if re.search(r"\(\s*\d{4}\s*\)", text):
        return "pass", f"[{style}] Detected author-year form with '(year)' ({len(entries)} entries)"
    return "warn", f"[{style}] No author-year '(year)' pattern detected; please verify reference format"


def _chk_author_year_order(entries, style) -> tuple:
    text = "\n".join(entries)
    if re.search(r"^[A-Z][a-z]+,\s+[A-Z]\.?", text.strip(), re.M):
        return "pass", f"[{style}] Entry start matches author-year form"
    return "info", f"[{style}] Entry does not start with typical 'Author, Initials'; please verify manually"


def _chk_author_title_order(entries, style) -> tuple:
    text = "\n".join(entries)
    if re.search(r"^[A-Z]", text.strip(), re.M):
        return "pass", f"[{style}] Works Cited format roughly matches ({len(entries)} entries)"
    return "warn", f"[{style}] Entries do not start with a capital letter; please verify Works Cited format"


def _chk_author_title_place_order(entries, style) -> tuple:
    return "info", f"[{style}] {len(entries)} bibliographic entries detected; please verify 'Author. Title. Place: Publisher, Year.'"


def _chk_numbered_ref(entries, style) -> tuple:
    numbered = sum(1 for e in entries if re.match(r"^\s*\[\s*\d+\s*\]", e))
    if numbered == 0:
        return "warn", f"[{style}] No numbered [n] form detected; please verify reference format"
    return "pass", f"[{style}] Numbered form matches ({numbered}/{len(entries)} entries start with [n])"


_REF_CHECKS = {
    "has_year_paren": _chk_has_year_paren,
    "author_year_order": _chk_author_year_order,
    "author_title_order": _chk_author_title_order,
    "author_title_place_order": _chk_author_title_place_order,
    "numbered_ref": _chk_numbered_ref,
}


def _check_reference(style: str, entries: list, spec, source: str = "spec") -> Finding:
    """参考文献格式启发式检查（著录形态，非逐条语义）。数据驱动。

    source: 该结论的来源标注（FR4 要求「并标注该结论来源」）。
    """
    if not entries:
        return Finding("reference", "warn",
                       "No reference entries detected (missing References / Bibliography / Works Cited section)",
                       expected=style, actual="0 entries", source=source)
    tokens = spec.reference.checks if spec and spec.reference.checks else ()
    if not tokens:
        return Finding("reference", "info",
                       f"Custom/unregistered check items ({len(entries)} entries); please verify against your school template",
                       expected=style, actual=f"{len(entries)} entries", source=source)
    results = []
    for tok in tokens:
        fn = _REF_CHECKS.get(tok)
        if fn:
            results.append(fn(entries, style))
    if not results:
        return Finding("reference", "info",
                       f"[{style}] {len(entries)} entries detected; please verify manually.",
                       expected=style, actual=f"{len(entries)} entries", source=source)
    # 取最严重（severity 数值最大）的一项
    worst = max(results, key=lambda r: SEVERITY_ORDER[r[0]])
    return Finding("reference", worst[0],
                   worst[1], expected=style, actual=f"{len(entries)} entries", source=source)


def _apa_reference_content_hints(entries: list, italics: list) -> list:
    """APA 参考文献「内容级」稳健提示（仅 info，绝不 fail/warn，且不误伤官方样例）。

    只针对整份文献的宏观可疑信号给提示，避免对逐条语义做武断判定：
    1) 至少 2 条、且全部无斜体 → 期刊名 / 书名可能未斜体（APA 7th §6.22 / 10.1）。
       官方标注样例含大量斜体期刊名，不会触发；纯网页型条目（罕见）最多给一条 info。
    2) 任一作者段出现 ", and "（逗号后接 and）→ 应用 "&" 连接最后两位作者（APA 7th §9.7）。
       只检查作者段（首个 "(" 或首个 4 位年份之前），避免误伤书名/标题中的 and
       （如 "War and Peace" 出现在年份之后，不计入作者段）。

    说明：官方样例本身存在「无 DOI」「缺期号」的条目（其教学批注亦如此标注），
    故此处刻意不检查 DOI / 期号，否则会误报金标准论文。
    """
    out: list = []
    # 1) 斜体提示：要求至少 2 条且确实存在「像文献」的条目（含 4 位年份），降低噪声
    if len(entries) >= 2 and any(re.search(r"\b(19|20)\d{2}\b", e) for e in entries):
        if not any(italics):
            out.append(Finding(
                dimension="reference",
                severity="info",
                message="Reference italics",
                expected="Journal names and book/report titles should be italicized (APA 7th §6.22 / 10.1)",
                actual="No italicized text found in any reference entry; verify journal names and book titles are italicized",
                source="spec",
            ))
    # 2) & 提示
    and_entries = 0
    for e in entries:
        cut = len(e)
        m_paren = re.search(r"\(", e)
        m_year = re.search(r"\b(19|20)\d{2}\b", e)
        if m_paren:
            cut = min(cut, m_paren.start())
        if m_year:
            cut = min(cut, m_year.start())
        author_seg = e[:cut]
        if re.search(r",\s+and\s+", author_seg, re.IGNORECASE):
            and_entries += 1
    if and_entries:
        out.append(Finding(
            dimension="reference",
            severity="info",
            message="Reference author connector",
            expected="Use '&' before the last author (APA 7th §9.7), e.g. 'Smith, J., & Jones, K.'",
            actual=f"{and_entries} entr{'y' if and_entries == 1 else 'ies'} use 'and' between authors instead of '&'",
            source="spec",
        ))
    return out


# ---------------------------------------------------------------------------
# 主检查
# ---------------------------------------------------------------------------

def _is_sorted_alphabetically(entries: list) -> bool:
    """保守判断参考文献条目是否按首作者姓氏字母序排列（APA 7th 要求）。

    不解析语义，仅比较每条起始 token（去掉前缀序号/标点）的小写形式。
    任一相邻逆序即判为未排序。用于「仅提示、不动内容」。
    """

    def _key(e: str) -> str:
        s = re.sub(r"^[\s\d\.\(\)\-–—]+", "", e.strip())
        return s.lower()

    if len(entries) < 2:
        return True
    keys = [_key(e) for e in entries]
    return all(keys[i] <= keys[i + 1] for i in range(len(keys) - 1))


def run_check(docx_path: str, target: TargetProfile) -> tuple[DocumentProfile, list[Finding]]:
    prof = read_docx(docx_path)
    findings: list[Finding] = []

    # —— 页边距 ——
    for key, side, _label in MARGIN_DIMS:
        if key in target.page:
            exp = target.page[key]
            origin = target.resolution.get(key, "spec")
            sev, actual = _cmp_margin(prof.margins.get(side), exp)
            sev = _downgrade(sev, origin, is_page_or_heading=True,
                             dim="margins", spec_key=target.spec_key)
            findings.append(Finding(
                dimension="margins",
                severity=sev,
                message=f"{side} margin",
                expected=f'{exp}"',
                actual=actual,
                source=origin,
            ))

    # —— 栏数（单栏/双栏；结构性维度）——
    # 与页边距/字体等「规范典型值仅作兜底、spec 来源降级 info」不同：选中某个规范作为
    # **来源**（而非只是引用风格）就是对该规范布局的显式声明，IEEE 双栏更是格式的定义性
    # 特征。故本维度**不**走 _downgrade 的 spec→info 软参考策略，不匹配照常 warn。
    # 学校模板来源（resolution="school"）/ 用户显式设置（"user"）本就是硬判，一致。
    if target.columns is not None:
        exp_cols = int(target.columns)
        origin = target.resolution.get("columns", "spec")
        act_cols = prof.columns if prof.columns is not None else 1
        sev = "pass" if act_cols == exp_cols else "warn"
        findings.append(Finding(
            dimension="columns",
            severity=sev,
            message="Column layout",
            expected="%d column%s" % (exp_cols, "" if exp_cols == 1 else "s"),
            actual="%d column%s" % (act_cols, "" if act_cols == 1 else "s"),
            source=origin,
        ))

    # —— 纸张大小（Letter / A4；**通用维度，不归属任何引用规范**）——
    # 北美（美/加）用 Letter 8.5×11，欧洲/中国/澳洲用 A4 210×297。同一篇 APA 论文
    # 在纽约交是 Letter、在伦敦交是 A4 —— 所以规范给的只是「典型值参考」，spec 来源
    # 一律降级 info（只提示，绝不因纸型给客户判红）；学校模板 / 问卷 / AI 显式给出
    # 时才硬判（那是客户自己声明的意图）。
    if target.paper_size:
        exp_paper = str(target.paper_size)
        origin = target.resolution.get("paper_size", "spec")
        act_name = getattr(prof, "paper_size", None)
        act_desc = describe_paper_size(
            getattr(prof, "paper_width_in", None),
            getattr(prof, "paper_height_in", None))
        if act_name is None and act_desc == "unknown":
            sev, actual = "info", "Could not read the page size"
        elif act_name == exp_paper:
            sev, actual = "pass", act_desc
        else:
            sev, actual = "warn", act_desc
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="paper_size", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="paper_size", severity=sev,
            message="Paper size",
            expected=exp_paper,
            actual=actual,
            source=origin,
        ))

    # —— 字体 ——
    if "font_family" in target.page and target.page["font_family"]:
        exp_font = target.page["font_family"]
        origin = target.resolution.get("font_family", "spec")
        act_font = prof.font_family or "unknown"
        ok = act_font.lower() == exp_font.lower() or exp_font.lower() in act_font.lower()
        # 🔴 APA 7th §2.19 允许 6 款字体（每款有自己的合规字号），spec 里的
        # page.font_family = Times New Roman 只是「该风格典型首选」，**不是唯一正确答案**。
        # 客户用 Calibri / Arial / Georgia 同样合规 → 不该判「字体不符」，更不该在修正时
        # 把客户的字体换成 Times New Roman（换字体是越权，且会毁掉中文文档的 eastAsia 字体）。
        # 字号由 font_size_by_font 按「该字体自己的规范字号」校准 —— 这才是配对表的用处。
        # 只对**规范来源**放宽：学校模板 / 客户明确要求的字体必须照办，绝不放宽。
        if not ok and origin == "spec":
            try:
                _allow = [f.lower() for f in (getattr(resolve_spec(target),
                                                      "font_allowlist", ()) or ())]
            except Exception:
                _allow = []
            if _allow and act_font.lower() in _allow:
                ok = True
        sev = "pass" if ok else "warn"
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="font", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="font", severity=sev,
            message="Body font",
            expected=exp_font, actual=act_font,
            source=origin,
        ))

    # —— 字号 ——
    # 🔴 字体命中规范的允许清单时，**该字体有自己的合规字号**（APA：Calibri 11 / TNR 12 /
    # Georgia 11 …），spec 里的「典型值 12pt」不再是正确答案 → 本维度让位给
    # font_size_by_font（逐 run 按各自字体判），避免同一件事出两条互相矛盾的结论。
    _size_superseded = False
    if "font_size_pt" in target.page:
        _size_origin = target.resolution.get("font_size_pt", "spec")
        if _size_origin == "spec":
            try:
                _pairs_pre = getattr(resolve_spec(target), "font_allowlist_sizes", {}) or {}
            except Exception:
                _pairs_pre = {}
            if size_for_font(prof.font_family, _pairs_pre):
                _size_superseded = True
    if "font_size_pt" in target.page and not _size_superseded:
        exp = target.page["font_size_pt"]
        origin = target.resolution.get("font_size_pt", "spec")
        sev, actual = _cmp_num(prof.font_size_pt, exp, SIZE_TOL_FAIL, SIZE_TOL_WARN, "pt")
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="font_size", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="font_size", severity=sev,
            message="Body font size",
            expected=f"{exp}pt", actual=actual,
            source=origin,
        ))

    # —— 行距 ——
    if "line_spacing" in target.page:
        exp = target.page["line_spacing"]
        origin = target.resolution.get("line_spacing", "spec")
        sev, actual = _cmp_num(prof.line_spacing, exp, SPACING_TOL_FAIL, SPACING_TOL_WARN, "x")
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="line_spacing", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="line_spacing", severity=sev,
            message="Line spacing",
            expected=f"{exp}x", actual=actual,
            source=origin,
        ))

    # —— 正文对齐（段落级；spec 来源降级 info，学校/问卷来源硬判）——
    if target.page.get("body_alignment"):
        exp_al = str(target.page["body_alignment"])
        origin = target.resolution.get("body_alignment", "spec")
        act_al = prof.body_alignment
        if act_al is None:
            sev, actual = "warn", "Could not read (all paragraph alignment inherits from styles)"
        elif act_al == exp_al:
            sev, actual = "pass", ALIGN_LABEL.get(act_al, act_al)
        else:
            sev, actual = "warn", ALIGN_LABEL.get(act_al, act_al)
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="alignment", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="alignment", severity=sev, message="Body alignment",
            expected=ALIGN_LABEL.get(exp_al, exp_al), actual=actual, source=origin,
        ))

    # —— 正文首行缩进（段落级）——
    if target.page.get("first_line_indent_in") is not None:
        exp = float(target.page["first_line_indent_in"])
        origin = target.resolution.get("first_line_indent_in", "spec")
        sev, actual = _cmp_num(prof.first_line_indent_in, exp,
                               INDENT_TOL_FAIL, INDENT_TOL_WARN, "″")
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="first_line_indent", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="first_line_indent", severity=sev, message="Body first-line indent",
            expected=f'{exp}″', actual=actual, source=origin,
        ))

    # —— 页眉距页顶（MLA 9th 要求 ½″；页面维度，spec 来源降级 info，绝不 fail）——
    if target.page.get("header_distance_in") is not None:
        exp = float(target.page["header_distance_in"])
        origin = target.resolution.get("header_distance_in", "spec")
        act = prof.header_distance_in
        if act is None:
            sev, actual = "info", "Could not read (document has no section header distance set)"
        else:
            sev, actual = _cmp_num(act, exp, HEADER_TOL_FAIL, HEADER_TOL_WARN, "″")
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="header_distance", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="header_distance", severity=sev, message="Header distance from top",
            expected=f'{exp}″', actual=actual, source=origin,
        ))

    # —— 段后间距（段落级）——
    if target.page.get("space_after_pt") is not None:
        exp = float(target.page["space_after_pt"])
        origin = target.resolution.get("space_after_pt", "spec")
        sev, actual = _cmp_num(prof.space_after_pt, exp,
                               SPACE_TOL_FAIL, SPACE_TOL_WARN, "pt")
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="space_after", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="space_after", severity=sev, message="Space after paragraph",
            expected=f"{exp}pt", actual=actual, source=origin,
        ))

    # —— 标题层级 ——
    if target.heading_levels and target.heading_levels > 0:
        origin = target.resolution.get("heading_levels", "spec")
        _pseudo_lvls = sorted(getattr(prof, "heading_formats", None) or {}) \
            if getattr(prof, "heading_formats_are_pseudo", False) else []
        if prof.heading_levels_used == 0 and not _pseudo_lvls:
            msg = "No Heading styles detected (manual bold may be used instead of styles)"
            sev = _downgrade("warn", origin, is_page_or_heading=True,
                             dim="heading", spec_key=target.spec_key)
        elif prof.heading_levels_used == 0:
            # 无 Heading 样式，但靠「整段加粗 + 字号档」认出 N 级手排标题。
            # 层级属推断口径，必须如实告知，不能装作读到了标题样式。
            msg = ("No Heading styles; %d level%s recognised from manual bold headings "
                   "(levels are inferred from font size)"
                   % (len(_pseudo_lvls), "" if len(_pseudo_lvls) == 1 else "s"))
            sev = _downgrade("warn", origin, is_page_or_heading=True,
                             dim="heading", spec_key=target.spec_key)
        elif prof.heading_levels_used > target.heading_levels:
            msg = f"Detected {prof.heading_levels_used} levels (exceeds ≤{target.heading_levels})"
            sev = _downgrade("warn", origin, is_page_or_heading=True,
                             dim="heading", spec_key=target.spec_key)
        else:
            msg = f"Detected {prof.heading_levels_used} levels (styles: {', '.join(prof.heading_styles) or 'none'})"
            sev = _downgrade("pass", origin, is_page_or_heading=True,
                             dim="heading", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="heading", severity=sev,
            message="Heading levels",
            expected=f"≤{target.heading_levels} levels",
            actual=msg,
            source=origin,
        ))

    # —— 每级标题的具体格式（字体 / 字号 / 对齐 / 加粗 / 斜体 / 缩进量）——
    # v2.3.65 新增维度。背景：旧版只查「层级数 ≤ N」，至于「一级标题该是黑体三号
    # 居中、二级该是楷体四号左对齐」，学校模板读进来、问卷填进来、AI 填进来，
    # 检查器全都不看 —— 客户填了 8 个字段，报告只回一句「层级数对」。
    # 只对**目标显式给出**的层级与属性做比对（None = 不规定 → 不碰），不臆造。
    heading_targets = getattr(target, "heading_formats", None) or {}
    if heading_targets:
        hf_origin = target.resolution.get("heading_formats", target.source or "school")
        want_map = resolve_heading_formats(heading_targets, None, target.heading_levels)
        have_map = prof.heading_formats or {}
        for lvl in sorted(want_map):
            want = want_map[lvl]
            got = have_map.get(lvl)
            if got is None:
                continue        # 文档里没这一级标题 → 不凭空判缺失
            exp_bits: list = []
            act_bits: list = []
            unreadable: list = []

            def _num_diff(label, want_v, got_v, unit=""):
                if want_v is None:
                    return
                if got_v is None:
                    unreadable.append(label)
                    return
                if abs(float(got_v) - float(want_v)) <= 1e-6:
                    return
                exp_bits.append("%s %g%s" % (label, float(want_v), unit))
                act_bits.append("%s %g%s" % (label, float(got_v), unit))

            if want.alignment:
                if got.alignment is None:
                    unreadable.append("alignment")
                elif got.alignment != want.alignment:
                    exp_bits.append(ALIGN_LABEL.get(want.alignment, want.alignment))
                    act_bits.append(ALIGN_LABEL.get(got.alignment, got.alignment))
            if want.font_family:
                if got.font_family is None:
                    unreadable.append("font")
                elif str(got.font_family).strip().lower() != str(want.font_family).strip().lower():
                    exp_bits.append(str(want.font_family))
                    act_bits.append(str(got.font_family))
            _num_diff("size", want.size_pt, got.size_pt, "pt")
            indent_want = want.indent_in if want.indent_in is not None else (
                0.5 if want.indent else None)
            _num_diff("indent", indent_want, got.indent_in, "″")
            for label, want_b, got_b in (("bold", want.bold, got.bold),
                                         ("italic", want.italic, got.italic)):
                if want_b is None or got_b is None:
                    continue
                if bool(got_b) != bool(want_b):
                    exp_bits.append(label if want_b else "not " + label)
                    act_bits.append(label if got_b else "not " + label)

            def _actual_text() -> str:
                bits = list(act_bits)
                if unreadable:
                    bits.append("could not read: " + ", ".join(unreadable))
                return ", ".join(bits) if bits else "Matches"

            if unreadable and not exp_bits:
                sev = _downgrade("info", hf_origin, is_page_or_heading=True,
                                 dim="heading", spec_key=target.spec_key)
                findings.append(Finding(
                    dimension="heading", severity=sev,
                    message="Level %d heading format" % lvl,
                    expected=describe_heading(want),
                    actual="Could not read: " + ", ".join(unreadable),
                    source=hf_origin,
                ))
                continue
            sev = "pass" if not exp_bits else "fail"
            sev = _downgrade(sev, hf_origin, is_page_or_heading=True,
                             dim="heading", spec_key=target.spec_key)
            findings.append(Finding(
                dimension="heading", severity=sev,
                message="Level %d heading format" % lvl,
                expected=", ".join(exp_bits) if exp_bits else describe_heading(want),
                actual=_actual_text(),
                source=hf_origin,
            ))

    # —— 标题页格式（标题 / 作者单位的居中与加粗）——
    # 优先级：客户显式给出（学校模板实测 / 问卷 / AI）> 内置规范的 title_page。
    # 过去这条链只有 fixer 一半：spec 有 title_page 就**悄悄改**标题页，检查器从不报告、
    # 预览里也没有 —— 违反「检查→修正单一真相源」。现在先出结论，再允许施加。
    _tp_origin = None
    _tp_vals = {}
    for _k in ("title_page_title_centered", "title_page_title_bold",
               "title_page_author_centered"):
        _v = getattr(target, _k, None)
        if _v is not None:
            _tp_vals[_k] = _v
            _tp_origin = _tp_origin or target.resolution.get(_k, "school")
    _spec_tp = getattr(resolve_spec(target), "title_page", None)
    if not _tp_vals and _spec_tp is not None:
        # 规范兜底：spec 里有声明就照它判，来源记 spec（页面维度 → 提示语气）
        if _spec_tp.title_centered is not None:
            _tp_vals["title_page_title_centered"] = bool(_spec_tp.title_centered)
        if _spec_tp.title_bold is not None:
            _tp_vals["title_page_title_bold"] = bool(_spec_tp.title_bold)
        if _spec_tp.author_centered is not None:
            _tp_vals["title_page_author_centered"] = bool(_spec_tp.author_centered)
        _tp_origin = "spec"
    if _tp_vals:
        _tp_origin = _tp_origin or "spec"
        if not getattr(prof, "title_page_detected", False):
            # 没标题页不算错：很多学校要求封面另附或直接首页开正文
            findings.append(Finding(
                dimension="title_page", severity="info",
                message="Title page format",
                expected=_tp_expected_text(_tp_vals),
                actual="No title page detected (the first page looks like body text)",
                source=_tp_origin,
            ))
        else:
            _bad, _act = [], []
            _want_center = _tp_vals.get("title_page_title_centered")
            _cur_al = getattr(prof, "tp_title_alignment", None)
            if _want_center is not None:
                if _cur_al is None:
                    _act.append("could not read the title alignment")
                elif (_cur_al == "center") != bool(_want_center):
                    # 文案说「文档哪里不对」，不是说「目标是怎样」：
                    # 要求居中而它没居中 → "not centered"。
                    _bad.append("not centered" if _want_center else "centered")
                    _act.append(ALIGN_LABEL.get(_cur_al, _cur_al))
                else:
                    _act.append(ALIGN_LABEL.get(_cur_al, _cur_al))
            _want_bold = _tp_vals.get("title_page_title_bold")
            _cur_b = getattr(prof, "tp_title_bold", None)
            if _want_bold is not None:
                if _cur_b is None:
                    _act.append("could not read the title weight")
                elif bool(_cur_b) != bool(_want_bold):
                    _bad.append("not bold" if _want_bold else "bolded")
                    _act.append("bold" if _cur_b else "not bold")
                else:
                    _act.append("bold" if _cur_b else "not bold")
            _want_ac = _tp_vals.get("title_page_author_centered")
            _cur_ac = getattr(prof, "tp_author_alignment", None)
            if _want_ac is not None:
                if _cur_ac is None:
                    _act.append("could not read the author-line alignment")
                elif (_cur_ac == "center") != bool(_want_ac):
                    _bad.append("author lines not centered" if _want_ac
                                else "author lines centered")
                    _act.append("authors " + ALIGN_LABEL.get(_cur_ac, _cur_ac))
                else:
                    _act.append("authors " + ALIGN_LABEL.get(_cur_ac, _cur_ac))
            _sev = "fail" if _bad else "pass"
            _sev = _downgrade(_sev, _tp_origin, is_page_or_heading=True,
                              dim="title_page", spec_key=target.spec_key)
            findings.append(Finding(
                dimension="title_page", severity=_sev,
                message="Title page format",
                expected=_tp_expected_text(_tp_vals),
                actual=("Title page: " + ", ".join(_bad)) if _bad else
                       (", ".join(_act) if _act else "Matches"),
                source=_tp_origin,
            ))

    # —— 参考文献（始终硬判，引用归规范）——
    ref_style = target.reference_style or "Other"
    # 用 spec_key 取完整规范：学生/专业共用 style_key="APA"，但 tier 标记
    # (running_head/expect_abstract/expect_author_note/running_head_max_chars) 按
    # spec_key 区分。若只按 ref_style("APA") 取，会永远落到学生版，专业版差异项不生效。
    # 单一真相源：checker 与 fixer 共用 resolve_spec，杜绝「检查/修正用不同规范」的漂移
    spec = resolve_spec(target)
    ref_label = "Works Cited" if spec.reference.style_key == "MLA" else "References"
    ref_source = target.resolution.get("reference_style", "spec")
    ref_finding = _check_reference(ref_style, prof.reference_entries, spec, ref_source)
    # 顺序编号制（IEEE 等）且存在未编号条目 → 编号缺失本身是硬问题，须提示并补 [n]，
    # 与 fixer 的 reference_numbered 修正严格一致（check→fix 单一真相源）。
    # 覆盖 _check_reference 可能给出的 pass/info（例如已有 [3] 时整体被误判为「已合规」），
    # 否则 fixer 会因 fix_dims 不含 reference_numbered 而跳过编号，造成「该补却没补」。
    if target.reference_numbered and prof.reference_entries:
        _missing = [t for t in prof.reference_entries
                    if t.strip() and not _has_leading_number(t)]
        if _missing:
            ref_finding = Finding(
                dimension="reference",
                severity="fail" if len(_missing) == len(prof.reference_entries) else "warn",
                message="Reference sequential numbering",
                expected="Each reference entry should carry a sequential [n] code (e.g., [1], [2])",
                actual=f"{len(_missing)} of {len(prof.reference_entries)} entries not numbered",
                source=ref_source,
            )
    findings.append(ref_finding)

    # —— APA 参考文献内容级提示（仅信息，不误伤官方样例；见 _apa_reference_content_hints）——
    if spec.reference.style_key == "APA" and prof.reference_entries:
        findings.extend(_apa_reference_content_hints(
            prof.reference_entries, prof.reference_entries_italic))

    # —— 参考文献悬挂缩进（引用维度，硬判但容差略宽）——
    if target.reference_hanging_indent_in is not None:
        exp_h = float(target.reference_hanging_indent_in)
        act_h = prof.reference_hanging_indent_in
        if act_h is None:
            sev_h, actual_h = "info", "Could not read (no reference section detected or no explicit indent)"
        else:
            diff = abs(act_h - exp_h)
            if diff <= REF_HANG_TOL_WARN:
                sev_h, actual_h = "pass", f'{act_h}″'
            elif diff <= REF_HANG_TOL_FAIL:
                sev_h, actual_h = "warn", f'{act_h}″ (deviation {diff:.2f}″)'
            else:
                sev_h, actual_h = "fail", f'{act_h}″ (deviation {diff:.2f}″)'
        findings.append(Finding(
            dimension="reference_hanging_indent", severity=sev_h,
            message="Reference hanging indent",
            expected=f'{exp_h}″', actual=actual_h,
            source=target.resolution.get("reference_hanging_indent_in", "spec"),
        ))

    # —— 参考文献条目自身的排版（字体 / 字号 / 行距）——
    # v2.3.65 新增。学校模板里「条目五号宋体、单倍行距」是硬要求，且常与正文不同口径；
    # 这几项以前 target 根本承载不了（只有 style/悬挂/编号），模板读到了也无处放 → 无人检查。
    for _key, _dim, _label, _unit, _tol in (
            ("reference_font_family", "reference_font", "Reference entry font", "", None),
            ("reference_size_pt", "reference_size", "Reference entry font size", "pt",
             (SIZE_TOL_FAIL, SIZE_TOL_WARN)),
            ("reference_line_spacing", "reference_spacing", "Reference entry line spacing", "x",
             (SPACING_TOL_FAIL, SPACING_TOL_WARN))):
        _want = getattr(target, _key, None)
        if _want is None:
            continue
        _origin = target.resolution.get(_key, "school")
        if _key == "reference_font_family":
            _got = prof.reference_font_family
            if _got is None:
                # 区分「没有参考文献段」与「条目没显式设字体」——后者不能当缺陷
                _sev = "info"
                _act = ("Could not read (the entries inherit their font, no explicit font set)"
                        if prof.reference_entries else
                        "Could not read (no reference section detected)")
            elif str(_got).strip().lower() == str(_want).strip().lower():
                _sev, _act = "pass", str(_got)
            else:
                _sev, _act = "fail", str(_got)
            _exp = str(_want)
        else:
            _got = (prof.reference_size_pt if _key == "reference_size_pt"
                    else prof.reference_line_spacing)
            if _got is None:
                # 读不到就只提示，不判 warn/fail（条目常继承正文字号，读不到不等于不合规）
                _sev = "info"
                _act = ("Could not read (the entries inherit this from the body text)"
                        if prof.reference_entries else
                        "Could not read (no reference section detected)")
            else:
                _sev, _act = _cmp_num(_got, float(_want), _tol[0], _tol[1], _unit)
            _exp = "%g%s" % (float(_want), _unit)
        findings.append(Finding(
            dimension=_dim, severity=_sev, message=_label,
            expected=_exp, actual=_act, source=_origin,
        ))

    # —— 参考文献：另起一页 / 标题加粗 / 标题居中 ——
    # 目标显式给出（学校模板实测 / 问卷 / AI）时硬判并可修正；未给出时保持下面
    # 「仅按 APA/MLA 规范提示」的旧行为，不越界。
    if target.reference_new_page is not None:
        _want_np = bool(target.reference_new_page)
        _got_np = prof.reference_new_page
        _origin = target.resolution.get("reference_new_page", "school")
        if _got_np is None:
            _sev, _act = "info", "No reference section detected"
        elif bool(_got_np) == _want_np:
            _sev, _act = "pass", "Starts on a new page" if _want_np else "Does not force a new page"
        else:
            _sev, _act = "fail", ("Does not begin on a new page" if _want_np
                                  else "Begins on a new page (a page break is present)")
        findings.append(Finding(
            dimension="reference_new_page", severity=_sev,
            message="References on new page",
            expected="begin on a new page" if _want_np else "do not force a new page",
            actual=_act, source=_origin,
        ))

    if target.reference_heading_bold is not None or target.reference_heading_centered is not None:
        _origin = target.resolution.get(
            "reference_heading_bold",
            target.resolution.get("reference_heading_centered", "school"))
        if not prof.reference_entries:
            findings.append(Finding(
                dimension="reference_heading", severity="info",
                message="Reference heading format",
                expected=", ".join(filter(None, [
                    None if target.reference_heading_bold is None else
                    ("bold" if target.reference_heading_bold else "not bold"),
                    None if target.reference_heading_centered is None else
                    ("centered" if target.reference_heading_centered else "left-aligned")])),
                actual="No reference section detected", source=_origin,
            ))
        else:
            _bad = []
            if target.reference_heading_bold is not None \
                    and bool(prof.reference_heading_bold) != bool(target.reference_heading_bold):
                _bad.append("bolded" if prof.reference_heading_bold else "not bold")
            if target.reference_heading_centered is not None \
                    and bool(prof.reference_heading_centered) != bool(target.reference_heading_centered):
                _bad.append("centered" if prof.reference_heading_centered else "not centered")
            findings.append(Finding(
                dimension="reference_heading",
                severity="fail" if _bad else "pass",
                message="Reference heading format",
                expected=", ".join(filter(None, [
                    None if target.reference_heading_bold is None else
                    ("bold" if target.reference_heading_bold else "not bold"),
                    None if target.reference_heading_centered is None else
                    ("centered" if target.reference_heading_centered else "left-aligned")])),
                actual=("The heading is " + " and ".join(_bad)) if _bad else "Matches",
                source=_origin,
            ))

    # —— APA 专属保守检查（仅当 spec 显式声明时生效；其余风格取中性默认值自动跳过）——
    # 这些检查只做「提示」，不产生 fail，也不修改文档内容（符合「模板驱动 / 只提示」原则）。
    # 页码 / running head 属页面维度，来源为 spec → 按代码约定降级为 info（不强制）。
    apa_spec = spec  # 当前 spec 已按 target.reference_style 取出（非 APA 时各字段为中性默认值）
    # —— 页码位置（v2.3.66 重写）——
    # 三处旧错：
    #   ① 只判「含页码的段落是否右对齐」，**不分页眉/页脚** → 页脚右对齐也判符合 APA top-right；
    #   ② 客户手打「第 1 页」被判「没有页码」→ 一修就出现两个页码；
    #   ③ 对齐读不到（继承样式）时落进 else 报 pass → 假通过。
    # 现在按 PAGE_NUMBER_LAYOUT 把「位置」与「对齐」一起比，且区分手打/自动。
    _pn_pos = target.page_number_position or apa_spec.page_number_position
    _pn_res = origin_of_source(target.resolution.get("page_number_position"))
    _pn_explicit = _pn_res in ("user", "school", "ai_import")
    _pn_origin = _pn_res or "spec"
    if _pn_pos:
        _want_in, _want_align = PAGE_NUMBER_LAYOUT.get(_pn_pos, ("header", "right"))
        _pos_label = {
            "top-right": "top-right of the header",
            "bottom-center": "bottom-center of the footer",
            "bottom-right": "bottom-right of the footer",
        }.get(_pn_pos, _pn_pos)
        if not prof.page_number_present:
            findings.append(Finding(
                dimension="page_number",
                severity=_downgrade("warn", _pn_origin, True, "page_number", target.spec_key),
                message="Page number",
                expected=f"Page number at {_pos_label}",
                actual="No page number found in the header or footer",
                source=_pn_origin,
            ))
        else:
            _bad = []
            if prof.page_number_in and prof.page_number_in != _want_in:
                _bad.append("placed in the " + prof.page_number_in)
            if prof.page_number_alignment and prof.page_number_alignment != _want_align:
                _bad.append(prof.page_number_alignment + "-aligned")
            if _bad:
                findings.append(Finding(
                    dimension="page_number",
                    severity=_downgrade("warn", _pn_origin, True, "page_number", target.spec_key),
                    message="Page number position",
                    expected=_pos_label,
                    actual="Page number is " + " and ".join(_bad),
                    source=_pn_origin,
                ))
            else:
                findings.append(Finding(
                    dimension="page_number", severity="pass",
                    message="Page number",
                    expected=_pos_label,
                    actual="Present at " + _pos_label,
                    source=_pn_origin,
                ))
        # 手打页码单独给一条提示：修正时会**原地**把数字换成 PAGE 域（保留「第…页」前后缀），
        # 不是再加一条 —— 客户必须知道自己那串数字被动过。
        if prof.page_number_present and prof.page_number_is_static:
            findings.append(Finding(
                dimension="page_number", severity="info",
                message="Page number is typed, not automatic",
                expected="An automatic PAGE field (so Word renumbers every page by itself)",
                actual='Typed page number "%s" — the fix replaces only the digits with a PAGE '
                       "field and keeps the surrounding text" % prof.page_number_static_text,
                source="doc",
            ))

    # 目标已显式给出「另起一页」时，上面已经硬判并可修，这里不再重复提示
    if apa_spec.reference_new_page and target.reference_new_page is None:
        if prof.reference_new_page is None:
            findings.append(Finding(
                dimension="reference_new_page", severity="info",
                message=f"{ref_label} on new page",
                expected=f"{ref_label} begin on a new page",
                actual="No reference section detected",
                source="spec",
            ))
        elif not prof.reference_new_page:
            findings.append(Finding(
                dimension="reference_new_page", severity="info",
                message=f"{ref_label} on new page",
                expected=f"{ref_label} begin on a new page",
                actual=f"{ref_label} heading does not begin on a new page (insert a page break before {ref_label})",
                source="spec",
            ))
        else:
            findings.append(Finding(
                dimension="reference_new_page", severity="pass",
                message=f"{ref_label} on new page",
                expected=f"{ref_label} begin on a new page",
                actual="Starts on a new page",
                source="spec",
            ))

    # —— 字体白名单（APA 7th §2.19）：**字体 + 字号配对**，不是一串字体名 ——
    # 旧版只比字体名：客户用 Calibri 12pt（APA 要求 Calibri 11pt）会被判「合格」。
    # 修正策略（老板 2026-10-19 拍板）：**保留客户选的字体，只把字号校正到该字体的
    # 规范值** —— 替客户换字体是越权（6 款都合规，没有唯一正确答案），改字号才有。
    # 白名单之外的字体绝不猜字号（只提示），客户显式给了正文字号时也不抢（由 font_size 管）。
    allow_pairs = getattr(apa_spec, "font_allowlist_sizes", None) or {}
    font_size_by_font: dict = {}          # 字体 → 该字体应有的规范字号（只含要改的）
    if apa_spec.font_allowlist and prof.font_family:
        allowed = [f.lower() for f in apa_spec.font_allowlist]
        # 按「每种字体各自的众数字号」判，而不是只看全篇众数：
        # 正文里混用 Times New Roman（APA 要 12pt）与 Calibri（要 11pt）时，
        # 只看众数会把 Calibri 那几段彻底漏掉。
        _fam_sizes = dict(getattr(prof, "font_size_by_family", None) or {})
        if not _fam_sizes:
            _fam_sizes = {prof.font_family: prof.font_size_pt}
        _outside = [f for f in _fam_sizes if f.lower() not in allowed]
        if _outside:
            findings.append(Finding(
                dimension="font_allowlist", severity="info",
                message="Body font (APA allowlist)",
                expected="One of: " + ", ".join(
                    f"{k} {v:g}pt" for k, v in allow_pairs.items()) if allow_pairs
                else "One of: " + ", ".join(apa_spec.font_allowlist),
                actual=", ".join(_outside) + " (APA permits a range of readable fonts; "
                       "verify your school's requirement)",
                source="spec",
            ))
        # 「客户显式给了正文字号」时不抢：那种情况由 font_size 维度硬判硬改。
        # 注意必须看 resolution —— 选 APA 规范时 spec 自带 font_size_pt=12，
        # 那是「该风格典型值」，不属于客户显式要求，不能拿它挡掉字体配对的字号修正。
        _size_owned = origin_of_source(
            target.resolution.get("font_size_pt")) in ("user", "school", "ai_import")
        _cur_map: dict = {}
        if not _size_owned:
            for _fam, _cur in _fam_sizes.items():
                if _fam.lower() not in allowed:
                    continue
                _want = size_for_font(_fam, allow_pairs)
                if not _want or not _cur:
                    continue
                if abs(float(_cur) - float(_want)) > 0.01:
                    font_size_by_font[_fam] = float(_want)
                    _cur_map[_fam] = float(_cur)
        if font_size_by_font:
            findings.append(Finding(
                dimension="font_size_by_font", severity="warn",
                message="Body text size for "
                        + ("this font" if len(font_size_by_font) == 1 else "these fonts"),
                expected=", ".join("%s %g pt" % (f, s) for f, s in font_size_by_font.items()),
                actual=", ".join("%s %g pt" % (f, s) for f, s in _cur_map.items()),
                source="spec",
            ))
        elif not _outside:
            findings.append(Finding(
                dimension="font_allowlist", severity="pass",
                message="Body font (APA allowlist)",
                expected="One of: " + ", ".join(
                    f"{k} {v:g}pt" for k, v in allow_pairs.items()) if allow_pairs
                else "One of: " + ", ".join(apa_spec.font_allowlist),
                actual=", ".join("%s %g pt" % (f, s) for f, s in _fam_sizes.items()),
                source="spec",
            ))

    if apa_spec.running_head and getattr(apa_spec, "running_head_kind", "apa") == "apa":
        # 仅专业版 running_head=True 且 kind="apa" 才进此分支；学生版 running_head=False 完全跳过
        # （学生版不要求页眉短标题，绝不被误报「缺 running head」）。属页面维度，spec 来源 → info。
        rh = prof.header_text or ""
        if not rh:
            findings.append(Finding(
                dimension="running_head", severity="info",
                message="Running head",
                expected=f"Short title in the header (≤ {apa_spec.running_head_max_chars} characters) + page number",
                actual="No running-head text detected in the header (APA Professional requires a short title)",
                source="spec",
            ))
        elif apa_spec.running_head_max_chars and len(rh) > apa_spec.running_head_max_chars:
            findings.append(Finding(
                dimension="running_head", severity="info",
                message="Running head length",
                expected=f"≤ {apa_spec.running_head_max_chars} characters",
                actual=f"{len(rh)} characters: {rh[:60]}{'…' if len(rh) > 60 else ''}",
                source="spec",
            ))
        else:
            findings.append(Finding(
                dimension="running_head", severity="pass",
                message="Running head length",
                expected=f"≤ {apa_spec.running_head_max_chars} characters",
                actual=f"{len(rh)} characters",
                source="spec",
            ))

    if spec.running_head and getattr(spec, "running_head_kind", "apa") == "mla":
        # MLA running head：姓氏 + 空格 + 页码，每页右上（Purdue OWL general_format）。
        # 页码常以 PAGE 域形式存在（header_text 中不含数字），故同时看 header_text 与 page_number_present。
        # 属页面维度，spec 来源 → info；绝不改动内容。
        rh = prof.header_text or ""
        name_part = re.sub(r"\d+", "", rh).strip()
        has_name = bool(re.search(r"[A-Za-z]", name_part)) or (rh.strip() != "" and not rh.strip().isdigit())
        has_page = bool(prof.page_number_present) or bool(re.search(r"\d", rh))
        if not rh.strip() and not prof.page_number_present:
            findings.append(Finding(
                dimension="running_head", severity="info",
                message="Running head",
                expected="Your last name followed by the page number, in the upper-right corner of every page",
                actual="No header text detected (MLA requires a running head with your last name + page number)",
                source="spec",
            ))
        elif not has_name:
            findings.append(Finding(
                dimension="running_head", severity="info",
                message="Running head",
                expected="Your last name + page number (e.g., 'Smith 1')",
                actual=f"Header text detected: {rh[:60]!r}; MLA expects your last name followed by the page number",
                source="spec",
            ))
        elif not has_page:
            findings.append(Finding(
                dimension="running_head", severity="info",
                message="Running head",
                expected="Your last name + page number",
                actual="Running head has a name but no page number detected",
                source="spec",
            ))
        else:
            findings.append(Finding(
                dimension="running_head", severity="pass",
                message="Running head",
                expected="Your last name + page number",
                actual=rh[:60] or "(name + page number field)",
                source="spec",
            ))

    # —— 前置材料：按各自 tier 分流（学生/专业差异项，不混用）——
    # 专业版（expect_abstract / expect_author_note = True）才提示「应有」；
    # 学生版（均 False）不要求，仅在「多写了作者注」时温和提示非必需。
    if apa_spec.expect_abstract and not prof.has_abstract:
        findings.append(Finding(
            dimension="abstract", severity="info",
            message="Abstract & Keywords",
            expected="An Abstract with Keywords is required (APA Professional)",
            actual="No Abstract section detected",
            source="spec",
        ))
    if apa_spec.expect_author_note and not prof.has_author_note:
        findings.append(Finding(
            dimension="author_note", severity="info",
            message="Author Note",
            expected="An Author Note is expected (APA Professional)",
            actual="No Author Note section detected",
            source="spec",
        ))
    if (not apa_spec.expect_author_note) and apa_spec.reference.style_key == "APA" and prof.has_author_note:
        # 仅「APA 学生版」触发（running_head=False 且 style_key="APA"）：学生论文通常不要求作者注，
        # 写了也仅作温和提示。注意：MLA/IEEE/Other 的 expect_author_note 同为 False，但绝不应触发此提示
        # （它们不是 APA 学生版），故用 style_key=="APA" 收窄，避免误伤其他风格。
        findings.append(Finding(
            dimension="author_note", severity="info",
            message="Author Note (not required for student papers)",
            expected="Usually not required for APA Student papers",
            actual="An Author Note section was found; include it only if your instructor asks",
            source="spec",
        ))

    if apa_spec.reference_alpha_sort_hint:
        # 参考文献字母序仅做提示，不改动内容（用户明确：排序涉及改内容，只提示）。
        if prof.reference_entries and len(prof.reference_entries) > 1:
            if _is_sorted_alphabetically(prof.reference_entries):
                findings.append(Finding(
                    dimension="reference_sort", severity="pass",
                    message="Reference list order",
                    expected="Entries sorted alphabetically by first author's surname",
                    actual="Appears alphabetically ordered",
                    source="spec",
                ))
            else:
                findings.append(Finding(
                    dimension="reference_sort", severity="info",
                    message="Reference list order",
                    expected="Entries sorted alphabetically by first author's surname",
                    actual="Entries do not appear to be in alphabetical order (MLA/APA require alphabetical order)",
                    source="spec",
                ))

    # —— 参考文献标题格式（APA 7th：'References' 须粗体+居中于页顶；MLA 9th：'Works Cited' 居中、不加粗）——
    # 安全处理：官方样例用 Heading 1 样式承载，粗体/居中来自样式（python-docx 读不出样式解析值），
    # 故 is_style 为真时「信任样式」直接判 pass，绝不误伤金标准论文；仅当标题是「手动段落」
    # （非 Heading 样式）时才检查直接格式，缺失则给 info 提示（不 fail、不改内容）。
    if prof.reference_entries and spec.reference.style_key in ("APA", "MLA") \
            and target.reference_heading_bold is None \
            and target.reference_heading_centered is None:
        if prof.reference_heading_is_style:
            findings.append(Finding(
                dimension="reference_heading", severity="pass",
                message="Reference heading format",
                expected=f"{ref_label} uses a Heading style"
                          + (" (bold/centered supplied by the style)" if spec.reference_heading_bold_mode == "require"
                             else " (centered, not bold supplied by the style)"),
                actual="Uses a Heading style",
                source="spec",
            ))
        else:
            problems = []
            if spec.reference_heading_bold_mode == "forbid":
                if prof.reference_heading_bold:
                    problems.append("bolded")
            elif spec.reference_heading_bold and not prof.reference_heading_bold:
                problems.append("not bold")
            if not prof.reference_heading_centered:
                problems.append("not centered")
            if problems:
                expected = f"{ref_label} should be centered" + \
                           (" and not bold" if spec.reference_heading_bold_mode == "forbid"
                            else " and bold" if spec.reference_heading_bold else "") + \
                           " at the top of a new page"
                findings.append(Finding(
                    dimension="reference_heading", severity="info",
                    message="Reference heading format",
                    expected=expected,
                    actual=f"The '{ref_label}' heading is " + " and ".join(problems),
                    source="spec",
                ))

    # —— MLA 专属版式检查（仅信息/提示级，绝不改动内容；引文内容不查）——
    bq: list = []      # 待修块引用候选（供下方 FixAction 绑定用；非 MLA 时为 [];）
    if spec.reference.style_key == "MLA":
        # G3 标题格式：居中、不加粗、不斜体、不下划线（Purdue OWL general_format）
        if prof.title_text:
            t_problems = []
            if prof.title_alignment != "center":
                t_problems.append("not centered")
            if prof.title_bold:
                t_problems.append("bolded")
            if prof.title_italic:
                t_problems.append("italicized")
            if prof.title_underline:
                t_problems.append("underlined")
            if t_problems:
                findings.append(Finding(
                    dimension="title_format", severity="info",
                    message="Title format",
                    expected="Centered, not bold / italic / underlined (MLA)",
                    actual="The title is " + " and ".join(t_problems),
                    source="spec",
                ))
            else:
                findings.append(Finding(
                    dimension="title_format", severity="pass",
                    message="Title format",
                    expected="Centered, not bold/italic/underlined",
                    actual="Centered, unformatted",
                    source="spec",
                ))
        # G2 首页信息块：标题之前应有左上对齐的 姓名/教授/课程/日期（双行距）共约 4 行
        if prof.title_text and prof.first_page_left_lines < 3:
            findings.append(Finding(
                dimension="first_page_heading", severity="info",
                message="First-page heading block",
                expected="Upper left: your name, instructor, course, date (double-spaced)",
                actual="Could not clearly detect the 4-line MLA heading block before the title",
                source="spec",
            ))
        # G5 块引用：含引号且较长的段落应缩进 0.5″、不加引号（仅提示，避免误伤普通长段落）
        bq = [c for c in prof.block_quote_candidates
              if c["has_quote"] and (c["left_indent_in"] is None
                                     or abs((c["left_indent_in"] or 0) - 0.5) > 0.15)]
        if bq:
            findings.append(Finding(
                dimension="block_quote", severity="info",
                message="Block quotation",
                expected="Quotations longer than 4 lines should be block quotes (indent 0.5\" from left, no quotation marks)",
                actual=f"{len(bq)} long quotation paragraph(s) not formatted as a block quote",
                source="spec",
            ))

    # —— 把「需要修正」的 Finding 绑定 FixAction（check→fix 单一真相源）——
    # fixer 只消费带 fix 的 Finding，绝不再独立重算「要改什么」，从而确保
    # 「检查出来的问题」与「实际改的内容」必然一致。
    _fix_value = {
        "margins": target.page.get("margin_top_in"),
        "font": target.page.get("font_family"),
        "font_size": target.page.get("font_size_pt"),
        "line_spacing": target.page.get("line_spacing"),
        "alignment": target.page.get("body_alignment"),
        "first_line_indent": target.page.get("first_line_indent_in"),
        "space_after": target.page.get("space_after_pt"),
        "header_distance": target.page.get("header_distance_in"),
        "columns": target.columns,
        "paper_size": target.paper_size,
        "reference_hanging_indent": target.reference_hanging_indent_in,
    }
    # 参考文献细项只在「目标显式给出」时才允许修正：值来自规范软参考（spec）时保持
    # 「仅提示」，避免把提示类维度悄悄升级成会改文档的动作。
    for _k, _dim in (("reference_font_family", "reference_font"),
                     ("reference_size_pt", "reference_size"),
                     ("reference_line_spacing", "reference_spacing"),
                     ("reference_new_page", "reference_new_page"),
                     ("reference_heading", "reference_heading")):
        if _k == "reference_heading":
            _v = (target.reference_heading_bold
                  if target.reference_heading_bold is not None
                  else target.reference_heading_centered)
            if target.reference_heading_bold is None \
                    and target.reference_heading_centered is None:
                continue
        else:
            _v = getattr(target, _k, None)
        if _v is not None:
            _fix_value[_dim] = _v
    # 标题页 / 块引用：这两项**会实际改动文档**，必须由检查结论驱动。
    # 旧版这两处是 fixer 里不受 gating 的自动动作（客户看不到、报告里也没有）——
    # 违反「检查→修正单一真相源」。现在补成正常维度：检查出结论 → 绑 FixAction →
    # 修正器才动手 → 修正预览里也能看见。
    if _tp_vals:
        _fix_value["title_page"] = True
    if bq:
        _fix_value["block_quote"] = 0.5
    # 页码位置：**只有客户/学校/AI 显式要求时才动手**（来源 spec 时只提示）。
    # 理由与纸型一致 —— 国内学校普遍要「页脚居中」，把 APA 的 top-right 当硬要求去改，
    # 会把学生本来正确的页脚居中改成页眉右上。客户显式选了才改。
    if _pn_pos and _pn_explicit and prof.page_number_present:
        _fix_value["page_number"] = _pn_pos
    elif _pn_pos and _pn_explicit and not prof.page_number_present:
        _fix_value["page_number"] = _pn_pos      # 完全没有页码时，显式要求就补一个
    # 字体白名单的字号配对：值传**整张配对表**（不是单一字号）—— fixer 需要逐 run
    # 按各自字体查自己的规范字号，正文混用 TNR 与 Calibri 时才能各改各的。
    if font_size_by_font:
        _fix_value["font_size_by_font"] = dict(allow_pairs)

    for f in findings:
        if f.severity == "pass":
            continue
        if f.dimension in _fix_value:
            f.fix = FixAction(dimension=f.dimension, value=_fix_value[f.dimension])
        elif f.dimension == "reference" and target.reference_numbered:
            # IEEE 顺序编码制：references 维度发现偏差（缺 [n]）→ 绑定编号修正
            f.fix = FixAction(dimension="reference_numbered", value=True)
        # running_head / reference_new_page / font_allowlist（字体本身，不换字体）/
        # reference_heading / reference_sort / title_format / first_page_heading /
        # abstract / author_note / 内容提示类：均为「仅提示、改内容」项，fix=None → fixer 绝不碰。
        # （page_number 与 font_size_by_font 已在 _fix_value 里，可修。）

    # 标题层级编号/格式：由 target.heading_levels 驱动（而非 finding 严重度），
    # 只要目标要求标题层级 > 0 即绑定 heading 修正，与 fixer 既有行为一致。
    # v2.3.65：每级标题格式的 finding 也挂在 heading 维度上，这里一并绑定（不再 break 一条）。
    if (target.heading_levels and target.heading_levels > 0) or heading_targets:
        _lvl_value = target.heading_levels or (max(heading_targets) if heading_targets else 0)
        for f in findings:
            if f.dimension == "heading":
                f.fix = FixAction(dimension="heading", value=_lvl_value)

    return prof, findings


def summarize(findings: list[Finding]) -> dict:
    counts = {"pass": 0, "info": 0, "warn": 0, "fail": 0}
    for f in findings:
        counts[f.severity] = counts.get(f.severity, 0) + 1
    return counts
