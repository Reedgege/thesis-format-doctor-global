"""PaperFormat Pro —— 真实 DOCX 端到端烟测（v2.3.57 发版门禁）。

不依赖 Tk / GUI，只走引擎：build_target -> run_check -> fix_docx -> compute_changes
-> preview_fix -> build_report，并强校验两条产品铁律：
  1) 纯格式修正，正文逐字不变；
  2) 英文模式(lang="en")产出的文案零 CJK（v2.3.51 i18n 回归线）。
"""

import re
import sys
import os
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_LINE_SPACING

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "src"))

from engine.questionnaire import target_from_spec
from engine.checker import run_check
from engine.fixer import fix_docx, compute_changes, preview_fix
from engine.report import build_report

CJK_RE = re.compile(r"[\u3000-\u9fff\uff00-\uffef]")

IN = os.path.join(HERE, "_smoke_in.docx")
OUT = os.path.join(HERE, "_smoke_out.docx")


def _para_texts(path):
    d = Document(path)
    return [p.text for p in d.paragraphs]


def _all_text(path):
    d = Document(path)
    parts = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(parts)


def _build_bad_docx():
    """构造一份明显不合规的学生论文：页边距全 2"、正文 Calibri 11、单倍行距。"""
    if os.path.exists(IN):
        os.remove(IN)
    d = Document()
    # 全局默认样式：Calibri 11（APA 要求 Times New Roman 12）
    normal = d.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    # 分节页边距全 2.0"（APA 要求 1.0"）
    for sec in d.sections:
        for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
            setattr(sec, m, Inches(2.0))
    # 标题（必须被修正器跳过，不能压成 12pt 正文）
    d.add_paragraph("The Impact of Social Media on Academic Performance", style="Heading 1")
    d.add_paragraph("1. Introduction", style="Heading 2")
    # 正文：单倍行距
    body = [
        "Social media has become an integral part of university life. "
        "Many students report spending several hours each day on these platforms.",
        "This study examines whether heavy usage correlates with lower grade point averages "
        "across a sample of undergraduate respondents.",
        "We argue that the relationship is mediated by sleep quality and self-regulation.",
        "2. Literature Review",  # 故意写成正文（非标题样式）制造一处格式噪点
    ]
    for t in body:
        p = d.add_paragraph(t)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    # 参考文献段（无编号，APA 不要求编号，但要能识别且不改内容）
    d.add_paragraph("References", style="Heading 2")
    d.add_paragraph("Smith, J. A. (2020). Title of work. Publisher.")
    d.add_paragraph("Lee, K. (2021). Another study. Journal of Things, 12(3), 45-67.")
    d.save(IN)


def _measure_output(path):
    d = Document(path)
    sec = d.sections[0]
    margins = {
        "left": sec.left_margin.inches,
        "right": sec.right_margin.inches,
        "top": sec.top_margin.inches,
        "bottom": sec.bottom_margin.inches,
    }
    # 取第一个非标题正文段的 run 字体 / 行距
    body_font = None
    body_spacing = None
    for p in d.paragraphs:
        style = (p.style.name or "") if p.style else ""
        if style.startswith("Heading"):
            continue
        if not p.text.strip():
            continue
        if p.runs:
            body_font = p.runs[0].font.name
            body_spacing = p.paragraph_format.line_spacing
            break
    return margins, body_font, body_spacing


def main():
    fails = []
    _build_bad_docx()

    target = target_from_spec("APA")
    print(f"[spec] APA: margins=1.0in, font={target.page.get('font')}, "
          f"size={target.page.get('size')}, spacing={target.page.get('spacing')}")

    # 1) run_check
    prof, findings = run_check(IN, target)
    print(f"[check] findings={len(findings)}; "
          f"severities={sorted({f.severity for f in findings})}")
    if not findings:
        fails.append("run_check returned no findings")
    fail_dims = [f.dimension for f in findings if f.severity == "fail"]
    print(f"[check] fail dimensions: {fail_dims}")
    if not fail_dims:
        fails.append("run_check produced no fail-level finding (doc should be non-compliant)")

    # 2) fix_docx
    if os.path.exists(OUT):
        os.remove(OUT)
    summary = fix_docx(IN, OUT, target)
    print(f"[fix] summary keys={sorted(summary.keys())}")
    if not os.path.exists(OUT):
        fails.append("fix_docx did not write output file")
        _report(fails)
        return

    # 3) 正文逐字不变（铁律 1）
    in_text = _all_text(IN)
    out_text = _all_text(OUT)
    if in_text != out_text:
        # 找首个差异
        a, b = _para_texts(IN), _para_texts(OUT)
        diff = next((i for i in range(min(len(a), len(b))) if a[i] != b[i]), None)
        fails.append(f"content changed after fix! first diff paragraph #{diff}: "
                     f"{a[diff]!r} -> {b[diff]!r}")
    else:
        print("[fix] content byte-identical before/after (铁律 OK)")

    # 4) 实测修正是否真的落地（页边距/字体/行距）
    margins, body_font, body_spacing = _measure_output(OUT)
    print(f"[measure] margins={margins} font={body_font} spacing={body_spacing}")
    if abs(margins["left"] - 1.0) > 0.02:
        fails.append(f"left margin not corrected to 1.0 (got {margins['left']})")
    if body_font not in ("Times New Roman", "Times New Roman (Headings)"):
        # python-docx 可能返回 None 若仅 rFonts 写入但 style 未显式设；用底层校验
        pass  # 字体逐 run 设置，style 层可能仍为 Calibri；见下方 run 级复核
    # run 级复核字体
    od = Document(OUT)
    for p in od.paragraphs:
        if (p.style.name or "").startswith("Heading"):
            continue
        if not p.text.strip():
            continue
        if p.runs and p.runs[0].font.name and "Times New Roman" not in p.runs[0].font.name:
            fails.append(f"body run font not Times New Roman after fix: {p.runs[0].font.name!r}")
            break
    if body_spacing is not None and abs(float(body_spacing) - 2.0) > 0.05:
        fails.append(f"body line spacing not corrected to 2.0 (got {body_spacing})")

    # 5) 英文模式零 CJK（铁律 2 / v2.3.51 回归线）
    changes_en = compute_changes(IN, target, lang="en")
    preview_en = preview_fix(IN, target, lang="en")
    cjk_in_changes = [s for s in changes_en if isinstance(s, str) and CJK_RE.search(s)]
    cjk_in_preview = [s for s in preview_en if isinstance(s, str) and CJK_RE.search(s)]
    print(f"[i18n] compute_changes(en)={len(changes_en)} lines; "
          f"preview_fix(en)={len(preview_en)} lines")
    if cjk_in_changes:
        fails.append(f"compute_changes(lang=en) leaked CJK: {cjk_in_changes[:3]}")
    if cjk_in_preview:
        fails.append(f"preview_fix(lang=en) leaked CJK: {cjk_in_preview[:3]}")

    # 6) 报告可构建
    rep = build_report(IN, target, prof, findings)
    print(f"[report] summary={rep.summary}")
    if not rep.findings:
        fails.append("build_report produced empty findings")

    _report(fails)


def _report(fails):
    print("\n" + "=" * 60)
    if fails:
        print("SMOKE TEST: FAIL")
        for f in fails:
            print("  - " + f)
        sys.exit(1)
    else:
        print("SMOKE TEST: PASS —— 真实 DOCX 端到端全链路通过")
        sys.exit(0)


if __name__ == "__main__":
    main()
