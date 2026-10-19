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

from .docx_reader import DocumentProfile, read_docx, DocxReadError
from .questionnaire import TargetProfile
from .specs import MARGIN_DIMS, get_spec, resolve_spec


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

    # —— 字体 ——
    if "font_family" in target.page and target.page["font_family"]:
        exp_font = target.page["font_family"]
        origin = target.resolution.get("font_family", "spec")
        act_font = prof.font_family or "unknown"
        ok = act_font.lower() == exp_font.lower() or exp_font.lower() in act_font.lower()
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
    if "font_size_pt" in target.page:
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
        if prof.heading_levels_used == 0:
            msg = "No Heading styles detected (manual bold may be used instead of styles)"
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

    # —— APA 专属保守检查（仅当 spec 显式声明时生效；其余风格取中性默认值自动跳过）——
    # 这些检查只做「提示」，不产生 fail，也不修改文档内容（符合「模板驱动 / 只提示」原则）。
    # 页码 / running head 属页面维度，来源为 spec → 按代码约定降级为 info（不强制）。
    apa_spec = spec  # 当前 spec 已按 target.reference_style 取出（非 APA 时各字段为中性默认值）
    if apa_spec.page_number_position:
        if not prof.page_number_present:
            findings.append(Finding(
                dimension="page_number", severity="info",
                message="Page number",
                expected=apa_spec.page_number_position,
                actual="No PAGE field detected in header/footer (MLA/APA require a page number on every page)",
                source="spec",
            ))
        elif apa_spec.page_number_position == "top-right" and prof.page_number_right_aligned is False:
            findings.append(Finding(
                dimension="page_number", severity="info",
                message="Page number position",
                expected="top-right (right-aligned in the header)",
                actual="Page number present but not right-aligned in the header",
                source="spec",
            ))
        else:
            findings.append(Finding(
                dimension="page_number", severity="pass",
                message="Page number",
                expected=apa_spec.page_number_position,
                actual="Present" + (" (right-aligned)" if prof.page_number_right_aligned else ""),
                source="spec",
            ))

    if apa_spec.reference_new_page:
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

    if apa_spec.font_allowlist:
        if prof.font_family:
            allowed = [f.lower() for f in apa_spec.font_allowlist]
            if prof.font_family.lower() not in allowed:
                findings.append(Finding(
                    dimension="font_allowlist", severity="info",
                    message="Body font (APA allowlist)",
                    expected="One of: " + ", ".join(apa_spec.font_allowlist),
                    actual=f"{prof.font_family} (APA permits a range of readable fonts; verify your school's requirement)",
                    source="spec",
                ))
            else:
                findings.append(Finding(
                    dimension="font_allowlist", severity="pass",
                    message="Body font (APA allowlist)",
                    expected="One of: " + ", ".join(apa_spec.font_allowlist),
                    actual=prof.font_family,
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
    if prof.reference_entries and spec.reference.style_key in ("APA", "MLA"):
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
        "reference_hanging_indent": target.reference_hanging_indent_in,
    }
    for f in findings:
        if f.severity == "pass":
            continue
        if f.dimension in _fix_value:
            f.fix = FixAction(dimension=f.dimension, value=_fix_value[f.dimension])
        elif f.dimension == "reference" and target.reference_numbered:
            # IEEE 顺序编码制：references 维度发现偏差（缺 [n]）→ 绑定编号修正
            f.fix = FixAction(dimension="reference_numbered", value=True)
        # header_distance / running_head / page_number / reference_new_page /
        # font_allowlist / reference_heading / reference_sort / title_format /
        # first_page_heading / block_quote / abstract / author_note / 内容提示类：
        # 均为「仅提示、改内容」项，fix=None → fixer 绝不碰。

    # 标题层级编号/格式：由 target.heading_levels 驱动（而非 finding 严重度），
    # 只要目标要求标题层级 > 0 即绑定 heading 修正，与 fixer 既有行为一致。
    if target.heading_levels and target.heading_levels > 0:
        for f in findings:
            if f.dimension == "heading":
                f.fix = FixAction(dimension="heading", value=target.heading_levels)
                break

    return prof, findings


def summarize(findings: list[Finding]) -> dict:
    counts = {"pass": 0, "info": 0, "warn": 0, "fail": 0}
    for f in findings:
        counts[f.severity] = counts.get(f.severity, 0) + 1
    return counts
