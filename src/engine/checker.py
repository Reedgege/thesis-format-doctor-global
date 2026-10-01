"""Rule Engine（确定性规则引擎）：把实测文档画像与 TargetProfile 比对。

原则：
- 全部为确定性判断，不调用任何大模型 / 网络。
- 每条 Finding 标注来源（spec / school / user / ai），供对比报告裁决。
- 页面排版来自「规范」时仅作参考（info），不强制 fail；
  页面排版来自学校/问卷/AI 时硬判。参考文献始终以规范硬判。
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional

from .docx_reader import DocumentProfile, read_docx, DocxReadError
from .questionnaire import TargetProfile
from .specs import MARGIN_DIMS, get_spec

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
SPACE_TOL_FAIL = 1.5        # 磅（段后间距）
SPACE_TOL_WARN = 0.8
REF_HANG_TOL_FAIL = 0.15    # 英寸（参考文献悬挂缩进；引用维度，容差略宽）
REF_HANG_TOL_WARN = 0.03

SEVERITY_ORDER = {"pass": 0, "info": 1, "warn": 2, "fail": 3}

# Alignment labels (also reused by report.py to avoid duplicate hard-coding)
ALIGN_LABEL = {"left": "Left", "justify": "Justified",
               "center": "Center", "right": "Right"}


@dataclass
class Finding:
    dimension: str
    severity: str            # pass / warn / fail / info
    message: str
    expected: str = ""
    actual: str = ""
    source: str = ""         # spec / school / user / ai


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


def _downgrade(sev: str, origin: str, is_page_or_heading: bool) -> str:
    """规范来源的页面/标题维度仅作参考（info），不强制 fail/warn。"""
    if is_page_or_heading and origin == "spec" and sev in ("warn", "fail"):
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


# ---------------------------------------------------------------------------
# 主检查
# ---------------------------------------------------------------------------

def run_check(docx_path: str, target: TargetProfile) -> tuple[DocumentProfile, list[Finding]]:
    prof = read_docx(docx_path)
    findings: list[Finding] = []

    # —— 页边距 ——
    for key, side, _label in MARGIN_DIMS:
        if key in target.page:
            exp = target.page[key]
            origin = target.resolution.get(key, "spec")
            sev, actual = _cmp_margin(prof.margins.get(side), exp)
            sev = _downgrade(sev, origin, is_page_or_heading=True)
            findings.append(Finding(
                dimension="margins",
                severity=sev,
                message=f"{side} margin",
                expected=f'{exp}"',
                actual=actual,
                source=origin,
            ))

    # —— 字体 ——
    if "font_family" in target.page and target.page["font_family"]:
        exp_font = target.page["font_family"]
        origin = target.resolution.get("font_family", "spec")
        act_font = prof.font_family or "unknown"
        ok = act_font.lower() == exp_font.lower() or exp_font.lower() in act_font.lower()
        sev = "pass" if ok else "warn"
        sev = _downgrade(sev, origin, is_page_or_heading=True)
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
        sev = _downgrade(sev, origin, is_page_or_heading=True)
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
        sev = _downgrade(sev, origin, is_page_or_heading=True)
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
        sev = _downgrade(sev, origin, is_page_or_heading=True)
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
        sev = _downgrade(sev, origin, is_page_or_heading=True)
        findings.append(Finding(
            dimension="first_line_indent", severity=sev, message="Body first-line indent",
            expected=f'{exp}″', actual=actual, source=origin,
        ))

    # —— 段后间距（段落级）——
    if target.page.get("space_after_pt") is not None:
        exp = float(target.page["space_after_pt"])
        origin = target.resolution.get("space_after_pt", "spec")
        sev, actual = _cmp_num(prof.space_after_pt, exp,
                               SPACE_TOL_FAIL, SPACE_TOL_WARN, "pt")
        sev = _downgrade(sev, origin, is_page_or_heading=True)
        findings.append(Finding(
            dimension="space_after", severity=sev, message="Space after paragraph",
            expected=f"{exp}pt", actual=actual, source=origin,
        ))

    # —— 标题层级 ——
    if target.heading_levels and target.heading_levels > 0:
        origin = target.resolution.get("heading_levels", "spec")
        if prof.heading_levels_used == 0:
            msg = "No Heading styles detected (manual bold may be used instead of styles)"
            sev = _downgrade("warn", origin, is_page_or_heading=True)
        elif prof.heading_levels_used > target.heading_levels:
            msg = f"Detected {prof.heading_levels_used} levels (exceeds ≤{target.heading_levels})"
            sev = _downgrade("warn", origin, is_page_or_heading=True)
        else:
            msg = f"Detected {prof.heading_levels_used} levels (styles: {', '.join(prof.heading_styles) or 'none'})"
            sev = _downgrade("pass", origin, is_page_or_heading=True)
        findings.append(Finding(
            dimension="heading", severity=sev,
            message="Heading levels",
            expected=f"≤{target.heading_levels} levels",
            actual=msg,
            source=origin,
        ))

    # —— 参考文献（始终硬判，引用归规范）——
    ref_style = target.reference_style or "Other"
    spec = get_spec(ref_style)
    ref_source = target.resolution.get("reference_style", "spec")
    findings.append(_check_reference(ref_style, prof.reference_entries, spec, ref_source))

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

    return prof, findings


def summarize(findings: list[Finding]) -> dict:
    counts = {"pass": 0, "info": 0, "warn": 0, "fail": 0}
    for f in findings:
        counts[f.severity] = counts.get(f.severity, 0) + 1
    return counts
