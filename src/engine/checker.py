"""Rule Engine（确定性规则引擎）：把实测文档画像与 TargetProfile 比对。

原则：
- 全部为确定性判断，不调用任何大模型 / 网络。
- 每条 Finding 标注来源（spec / school / user / ai），供对比报告裁决。
- 页面排版来自「规范」时仅作参考（info），不强制 fail；
  页面排版来自学校/问卷/AI 时硬判。参考文献始终以规范硬判。
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Any, Optional

from .docx_reader import (DocumentProfile, read_docx, DocxReadError,
                          describe_paper_size, identify_paper_size,
                          find_abstract_section)
from .questionnaire import TargetProfile, PAGE_NUMBER_LAYOUT, origin_of_source
from .specs import (MARGIN_DIMS, get_spec, resolve_spec,
                    resolve_heading_formats, describe_heading, size_for_font,
                    spec_is_baseline)


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
    # v2.3.70：**结构化明细**——把「检查具体报了哪一级」带给修正器，
    # 让它只改被报错的对象，而不是整个维度重刷一遍（老板：修正一定按检查结果来，
    # 检查结果要详细、修正要能读到）。标题维度按「级」细化；其它维度留空=整维度。
    levels: tuple = ()


def _finding_heading_level(msg: str) -> Optional[int]:
    """从「Level N heading format」这类 finding 文案里取出**具体层级**。

    只认固定前缀（我们自己产出），取不到就返回 None —— 调用方据此判定
    「这条结论没有细到级」，退回整维度处理，绝不猜。
    """
    m = re.match(r"^Level\s+(\d+)\s", msg or "")
    if not m:
        return None
    try:
        lv = int(m.group(1))
    except (TypeError, ValueError):
        return None
    return lv if lv > 0 else None


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

# 章节编号形态的人类可读提示（进报告 expected 列；面向客户，一律英文）。
_NUMBER_PREFIX_HINT = {
    "roman": "I., II., III.",
    "upper_letter": "A., B., C.",
    "number": "1., 2., 3.",
    "number_paren": "1), 2), 3)",
    "lower_letter": "a), b), c)",
    "lower_roman": "i), ii), iii)",
}


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
    # 客户读的报告里不许出现内部键名：spec_key 缺失时 style 会是 "Other"，
    # 原样写进 expected 会让客户以为「应该写成 Other」（v2.3.71 客户视角实测）。
    _disp = "" if str(style or "").strip().lower() in ("", "other", "none") else style
    style = _disp
    if not entries:
        return Finding("reference", "warn",
                       "No reference entries detected (missing References / Bibliography / Works Cited section)",
                       expected=style, actual="0 entries", source=source)
    tokens = spec.reference.checks if spec and spec.reference.checks else ()
    if not tokens:
        # v2.3.71（客户视角）：这里原来写的是
        #   "Custom/unregistered check items (N entries); please verify against your school template"
        # 客户读不懂（"Custom/unregistered" 是我们的内部概念），而且 expected 直接打出
        # "Other" 这种内部键名，客户会误以为「应该写成 Other」。真正该说清的是两件事：
        #   · 没选引用风格 → 参考文献根本没被检查；选了 APA/MLA/IEEE 才会查；
        #   · 该风格没有内置条目检查 → 需要人工核对。
        if str(style or "").strip().lower() in ("", "other", "none"):
            return Finding("reference", "info",
                           "Reference style not specified — reference entries were NOT checked. "
                           "Choose APA, MLA or IEEE to have them checked",
                           expected="", actual=f"{len(entries)} entries found", source=source)
        return Finding("reference", "info",
                       f"Reference style {style} has no built-in entry checks; "
                       f"{len(entries)} entries were found — please verify them manually "
                       f"against your school template",
                       expected="", actual=f"{len(entries)} entries", source=source)
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

    # —— 论文大标题（v2.3.73）——
    # 有摘要的论文由 title_page 维度管；**没有摘要**的论文以前会把大标题当成章节
    # 一级标题，报告于是写「Level 1 heading format: Act Center 16pt」，
    # 修正就把论文大标题改成模板的 14pt 左对齐。现在单独说清楚：认出来了、
    # 它不参与层级比对 —— 客户才不会把「只查到 N 级」误读成漏检。
    if getattr(prof, "doc_title_index", None) is not None \
            and not getattr(prof, "title_page_detected", False):
        _dt_sz = getattr(prof, "doc_title_size_pt", None)
        findings.append(Finding(
            dimension="doc_title", severity="info",
            message="Document title",
            expected="",
            actual=("Found on page 1 (%spt, centered, bold) — it is the paper title, "
                    "so it was not counted as a heading level"
                    % (("%g" % _dt_sz) if _dt_sz else "?")),
            source="document",
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
                if want_b is None:
                    continue
                if got_b is None:
                    # 目标规定了、文档读不到（run 与样式链都没写）→ 如实记「读不到」，
                    # 别静默跳过。旧写法直接 continue，于是「三级标题该斜体」这类要求在
                    # 报告里彻底消失，修正器却照样加斜体 —— 检查看不见的项被改了（v2.3.70）。
                    unreadable.append(label)
                    continue
                if bool(got_b) != bool(want_b):
                    exp_bits.append(label if want_b else "not " + label)
                    act_bits.append(label if got_b else "not " + label)
            # 编号形态（IEEE 的 I. / A. / 1) 全在 number_prefix 里）：v2.3.70 前
            # 检查器**从不检查**它，修正器却一直在补编号 —— 报告说「标题格式合规」，
            # 标题却多出了 "A."。现在缺编号要说出来，修正才有据可依。
            # 只报「缺」不报「形态不同」：作者手打的 "1. Introduction" 若被判成
            # 「应为 I.」，修正器会在前面再插一个前缀（I. 1. Introduction），宁可不碰。
            if want.number_prefix:
                if got.number_prefix is None:
                    exp_bits.append("numbering (%s)"
                                    % _NUMBER_PREFIX_HINT.get(want.number_prefix,
                                                              want.number_prefix))
                    act_bits.append("no numbering")
            # 全大写：只报「该大写却没大写」；反向（作者全大写、目标不要求）
            # 不报也不改 —— 小写化会破坏专有名词。读不到（文本无字母等）也不硬判。
            if want.uppercase and got.uppercase is False:
                exp_bits.append("uppercase")
                act_bits.append("not uppercased")
            # 段前分页（v2.3.67）：三态。want 为 None=不规定；got 为 None=读不到（可能由
            # 样式承载）→ 不硬判，只记 unreadable；两者都非 None 且不一致才判 fail。
            if want.page_break_before is not None:
                if got.page_break_before is None:
                    unreadable.append("page break before")
                elif bool(got.page_break_before) != bool(want.page_break_before):
                    exp_bits.append("page break before" if want.page_break_before
                                    else "no page break before")
                    act_bits.append("page break before" if got.page_break_before
                                    else "no page break before")

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

        # —— 超出目标声明范围的更深层级（v2.3.73）——
        # 目标声明了完整多级（APA/IEEE 五级）而文档更深（第 6 级）：检查要**明说**
        # 「第 6 级会按第 5 级规则处理」—— 旧版检查一个字不提、修正器却静默刷了一遍。
        # ⚠️ 只在 len(want_map) >= 2 时报：目标只声明 1 级（学校模板常见）时，
        # 更深层级**没有任何依据**，fixer 根本不会动它（v2.3.72 明确禁止拿一级格式套二级），
        # 这时写「会按第 1 级处理」就是假话。那种情形由上面的「层级数」结论负责说明。
        _declared = max(want_map) if want_map else 0
        _extra = sorted(set(have_map) - set(want_map))
        if _extra and _declared >= 1 and len(want_map) >= 2:
            _deep_want = want_map[_declared]
            for _lv in _extra:
                _got_x = have_map[_lv]
                _cur_bits = []
                for _lbl, _v, _unit in (("size", getattr(_got_x, "size_pt", None), "pt"),
                                        ("align", getattr(_got_x, "alignment", None), "")):
                    if _v:
                        _cur_bits.append("%s %s%s" % (_lbl, _v, _unit))
                findings.append(Finding(
                    dimension="heading",
                    severity=_downgrade("warn", hf_origin, is_page_or_heading=True,
                                        dim="heading", spec_key=target.spec_key),
                    message="Level %d heading format" % _lv,
                    expected=("Falls back to level %d rules (%s)"
                              % (_declared, describe_heading(_deep_want))),
                    actual=(", ".join(_cur_bits) or "Could not read"),
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
    # 🔴 v2.3.73：规范的 title_page 只在**规范基线**模式下兜底。非规范基线（学校模板 /
    # 手动 / AI）时 spec_key 为空，resolve_spec 只是按 reference_style 兜底取到整套规范 ——
    # 客户只选了「参考文献按 APA」，不该因此被要求「标题页居中加粗」。
    _spec_tp = (getattr(resolve_spec(target), "title_page", None)
                if spec_is_baseline(target) else None)
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

    # —— 摘要页特化（v2.3.67）——
    # 摘要正文口径与正文相反：APA 摘要**不首行缩进**（顶格）、专业版摘要**另起一页**。
    # 以前「正文首行缩进 0.5″」会把摘要段一起刷成 0.5″（越权），「参考文献另起一页」也不管摘要。
    # 目标显式给出（spec/school/user/ai）才判；None = 不规定 → 不碰。
    _abs_fli = target.abstract_first_line_indent_in
    _abs_np = target.abstract_new_page
    # 目标没显式给摘要字段、但规范带了（如 APA 专业版 abstract_new_page=True）→ 用规范值
    if _abs_fli is None:
        _abs_fli = apa_spec.abstract_first_line_indent_in
    if _abs_np is None:
        _abs_np = apa_spec.abstract_new_page
    if _abs_fli is not None or _abs_np is not None:
        # 直接用 prof 已采样值（不重复读文件）；摘要段是否存在由采样值是否 None 判断
        _abs_fli_origin = target.resolution.get("abstract_first_line_indent_in") \
            if target.abstract_first_line_indent_in is not None else "spec"
        _abs_np_origin = target.resolution.get("abstract_new_page") \
            if target.abstract_new_page is not None else "spec"
        if _abs_fli is not None:
            _cur_fli = prof.abstract_first_line_indent_in
            if _cur_fli is None:
                # 无摘要段 → 不判缺失（很多论文摘要与正文同页/无独立摘要段）
                findings.append(Finding(
                    dimension="abstract_indent", severity="info",
                    message="Abstract first-line indent",
                    expected="Abstract text is flush left (no first-line indent)"
                    if abs(float(_abs_fli)) < 1e-3 else
                    f"Abstract text first-line indent = {_abs_fli:g}″",
                    actual="No Abstract section detected (nothing to change)",
                    source=_abs_fli_origin,
                ))
            elif abs(float(_cur_fli) - float(_abs_fli)) > 0.01:
                findings.append(Finding(
                    dimension="abstract_indent", severity="warn",
                    message="Abstract first-line indent",
                    expected=f"{_abs_fli:g}″",
                    actual=f"{_cur_fli:g}″",
                    source=_abs_fli_origin,
                ))
            else:
                findings.append(Finding(
                    dimension="abstract_indent", severity="pass",
                    message="Abstract first-line indent",
                    expected=f"{_abs_fli:g}″",
                    actual=f"{_cur_fli:g}″",
                    source=_abs_fli_origin,
                ))
        if _abs_np is not None:
            _cur_np = prof.abstract_new_page
            if _cur_np is None:
                findings.append(Finding(
                    dimension="abstract_new_page", severity="info",
                    message="Abstract on new page",
                    expected="Abstract begins on a new page"
                    if _abs_np else "Abstract does not start on a new page",
                    actual="No Abstract section detected",
                    source=_abs_np_origin,
                ))
            elif bool(_cur_np) != bool(_abs_np):
                findings.append(Finding(
                    dimension="abstract_new_page", severity="warn",
                    message="Abstract on new page",
                    expected="begins on a new page" if _abs_np else "continues on the same page",
                    actual="begins on a new page" if _cur_np else "continues on the same page",
                    source=_abs_np_origin,
                ))
            else:
                findings.append(Finding(
                    dimension="abstract_new_page", severity="pass",
                    message="Abstract on new page",
                    expected="begins on a new page" if _abs_np else "continues on the same page",
                    actual="Matches",
                    source=_abs_np_origin,
                ))

    # —— Running head 显式维度（v2.3.67）——
    # 客户/学校/AI 显式声明「要 running head」或「页码含作者姓氏」时硬判；
    # 否则回落到已有 spec 驱动的 info 提示（APA 专业版 / MLA）。文本内容（短标题/姓氏）
    # 涉及「改内容」且无法可靠自动生成 → 只提示客户手动补；页码位置由 page_number 维度可修。
    _rh_req = target.running_head_required
    _rh_name = target.page_number_include_author_name
    _rh_origin = target.resolution.get("running_head_required") or \
        target.resolution.get("page_number_include_author_name") or "spec"
    _rh_origin = origin_of_source(_rh_origin) or "spec"
    if _rh_req is not None or _rh_name is not None:
        _rh_text = prof.header_text or ""
        _rh_has_text = bool(_rh_text.strip())
        _rh_has_page = bool(prof.page_number_present) or bool(re.search(r"\d", _rh_text))
        # ① 是否要求 running head
        if _rh_req is True:
            if not _rh_has_text and not _rh_has_page:
                findings.append(Finding(
                    dimension="running_head", severity="warn",
                    message="Running head",
                    expected="A running head in the header (short title or last name + page number)",
                    actual="No running-head text or page number detected in the header",
                    source=_rh_origin,
                ))
            else:
                findings.append(Finding(
                    dimension="running_head", severity="pass",
                    message="Running head",
                    expected="A running head in the header",
                    actual=_rh_text[:60] or "(page number in header)",
                    source=_rh_origin,
                ))
        elif _rh_req is False:
            # 明确不要求 running head：不判缺失，仅当有多余页眉文本时温和提示
            if _rh_has_text:
                findings.append(Finding(
                    dimension="running_head", severity="info",
                    message="Running head (not required)",
                    expected="No running head required",
                    actual=f"Header text detected: {_rh_text[:60]!r}",
                    source=_rh_origin,
                ))
        # ② 页码是否含作者姓氏（MLA 的「Smith 1」形态）
        if _rh_name is True:
            _name_part = re.sub(r"\d+", "", _rh_text).strip()
            _has_name = bool(re.search(r"[A-Za-z]", _name_part))
            if not _has_name:
                findings.append(Finding(
                    dimension="running_head_name", severity="warn",
                    message="Running head author name",
                    expected="Your last name followed by the page number (e.g., 'Smith 1')",
                    actual="No author name detected in the header",
                    source=_rh_origin,
                ))
            else:
                findings.append(Finding(
                    dimension="running_head_name", severity="pass",
                    message="Running head author name",
                    expected="Last name + page number",
                    actual=_rh_text[:60],
                    source=_rh_origin,
                ))

    # —— 图表 caption（v2.3.67）——
    # 字号/行距可查可修（识别 Caption 段落）；表题/图题位置只提示（图片 caption 位置
    # 无法可靠判定）。目标显式给出才判；None = 不规定。
    _cap_size = target.caption_font_size_pt
    _cap_spacing = target.caption_line_spacing
    _cap_tbl_pos = target.caption_position_table
    _cap_fig_pos = target.caption_position_figure
    # 目标没显式给、但规范带了 → 用规范值（来源 spec → 只提示语气）
    if _cap_size is None:
        _cap_size = apa_spec.caption_font_size_pt
    if _cap_spacing is None:
        _cap_spacing = apa_spec.caption_line_spacing
    if _cap_tbl_pos is None:
        _cap_tbl_pos = apa_spec.caption_position_table
    if _cap_fig_pos is None:
        _cap_fig_pos = apa_spec.caption_position_figure
    _has_captions = bool(getattr(prof, "caption_count", 0))
    # ① 字号（可修）
    if _cap_size is not None:
        _cap_size_origin = target.resolution.get("caption_font_size_pt", "spec")
        _cur_cap_size = prof.caption_font_size_pt
        if not _has_captions or _cur_cap_size is None:
            findings.append(Finding(
                dimension="caption_size", severity="info",
                message="Figure/table caption font size",
                expected=f"{_cap_size:g} pt",
                actual="No captions detected" if not _has_captions else "Could not read caption size",
                source=_cap_size_origin,
            ))
        elif abs(float(_cur_cap_size) - float(_cap_size)) > 0.01:
            findings.append(Finding(
                dimension="caption_size", severity="warn",
                message="Figure/table caption font size",
                expected=f"{_cap_size:g} pt",
                actual=f"{_cur_cap_size:g} pt",
                source=_cap_size_origin,
            ))
        else:
            findings.append(Finding(
                dimension="caption_size", severity="pass",
                message="Figure/table caption font size",
                expected=f"{_cap_size:g} pt",
                actual=f"{_cur_cap_size:g} pt",
                source=_cap_size_origin,
            ))
    # ② 行距（可修）
    if _cap_spacing is not None:
        _cap_spacing_origin = target.resolution.get("caption_line_spacing", "spec")
        _cur_cap_sp = prof.caption_line_spacing
        if not _has_captions or _cur_cap_sp is None:
            findings.append(Finding(
                dimension="caption_spacing", severity="info",
                message="Figure/table caption line spacing",
                expected=f"{_cap_spacing:g}x",
                actual="No captions detected" if not _has_captions else "Could not read caption spacing",
                source=_cap_spacing_origin,
            ))
        elif abs(float(_cur_cap_sp) - float(_cap_spacing)) > 0.01:
            findings.append(Finding(
                dimension="caption_spacing", severity="warn",
                message="Figure/table caption line spacing",
                expected=f"{_cap_spacing:g}x",
                actual=f"{_cur_cap_sp:g}x",
                source=_cap_spacing_origin,
            ))
        else:
            findings.append(Finding(
                dimension="caption_spacing", severity="pass",
                message="Figure/table caption line spacing",
                expected=f"{_cap_spacing:g}x",
                actual=f"{_cur_cap_sp:g}x",
                source=_cap_spacing_origin,
            ))
    # ③ 表题位置（只提示）
    if _cap_tbl_pos is not None:
        _cap_tbl_origin = target.resolution.get("caption_position_table", "spec")
        _tbl_positions = getattr(prof, "caption_table_positions", None) or []
        if not _tbl_positions:
            findings.append(Finding(
                dimension="caption_position_table", severity="info",
                message="Table caption position",
                expected="table caption " + ("above the table" if _cap_tbl_pos == "top" else "below the table"),
                actual="No table caption position could be determined",
                source=_cap_tbl_origin,
            ))
        else:
            _mode_pos = Counter(_tbl_positions).most_common(1)[0][0]
            _mode_label = "above" if _mode_pos == "top" else "below"
            _want_label = "above" if _cap_tbl_pos == "top" else "below"
            if _mode_pos != _cap_tbl_pos:
                findings.append(Finding(
                    dimension="caption_position_table", severity="info",
                    message="Table caption position",
                    expected="table caption " + _want_label + " the table",
                    actual="table caption " + _mode_label + " the table",
                    source=_cap_tbl_origin,
                ))
            else:
                findings.append(Finding(
                    dimension="caption_position_table", severity="pass",
                    message="Table caption position",
                    expected="table caption " + _want_label + " the table",
                    actual="Matches",
                    source=_cap_tbl_origin,
                ))
    # ④ 图题位置（只提示）：图片 caption 位置无法可靠判定，仅当有 caption 时给说明性提示
    if _cap_fig_pos is not None:
        _cap_fig_origin = target.resolution.get("caption_position_figure", "spec")
        _fig_want = "above the figure" if _cap_fig_pos == "top" else "below the figure"
        findings.append(Finding(
            dimension="caption_position_figure", severity="info",
            message="Figure caption position",
            expected="figure caption " + _fig_want,
            actual="Figure caption position is not auto-checked (figures are embedded in paragraphs); "
                   "verify it manually",
            source=_cap_fig_origin,
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
    # 摘要页特化：**只有客户/学校/AI 显式要求时才动手**（来源 spec 时只提示）。
    # 理由与页码位置一致 —— 摘要是否缩进/另起一页，不同学校要求不同，APA 的
    # 「顶格 + 另起一页」只是典型值，不能替客户拍板去改摘要页。
    if target.abstract_first_line_indent_in is not None \
            and origin_of_source(target.resolution.get("abstract_first_line_indent_in")) \
            in ("user", "school", "ai_import"):
        _fix_value["abstract_indent"] = float(target.abstract_first_line_indent_in)
    if target.abstract_new_page is not None \
            and origin_of_source(target.resolution.get("abstract_new_page")) \
            in ("user", "school", "ai_import"):
        _fix_value["abstract_new_page"] = bool(target.abstract_new_page)
    # 图表 caption 字号/行距可修：**只有客户/学校/AI 显式要求时才动手**（spec 只提示）。
    # caption 位置只提示（不绑 fix）。
    if target.caption_font_size_pt is not None \
            and origin_of_source(target.resolution.get("caption_font_size_pt")) \
            in ("user", "school", "ai_import"):
        _fix_value["caption_size"] = float(target.caption_font_size_pt)
    if target.caption_line_spacing is not None \
            and origin_of_source(target.resolution.get("caption_line_spacing")) \
            in ("user", "school", "ai_import"):
        _fix_value["caption_spacing"] = float(target.caption_line_spacing)

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
    # v2.3.70：把**具体层级**写进 FixAction.levels —— 「Level 3 heading format」这类
    # 检查结果本来就精确到级，修正器据此只改被报错的那级（否则全五级重刷＝不按结果修）。
    if (target.heading_levels and target.heading_levels > 0) or heading_targets:
        _lvl_value = target.heading_levels or (max(heading_targets) if heading_targets else 0)
        for f in findings:
            if f.dimension == "heading":
                _lv = _finding_heading_level(f.message)
                f.fix = FixAction(dimension="heading", value=_lvl_value,
                                  levels=(_lv,) if _lv else ())

    return prof, findings


def summarize(findings: list[Finding]) -> dict:
    counts = {"pass": 0, "info": 0, "warn": 0, "fail": 0}
    for f in findings:
        counts[f.severity] = counts.get(f.severity, 0) + 1
    return counts
