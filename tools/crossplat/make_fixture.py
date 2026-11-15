# -*- coding: utf-8 -*-
"""生成一份「结构完整、格式全错」的 APA 学生版论文，作为跨平台功能验收素材。

为什么要自己造
--------------
原来的 bad_apa.docx 有个硬伤：正文里有 2 条参考文献条目，却**没有 References
标题**。于是检查器报 "No reference entries detected"，参考文献重排那条路径压根
没被走到；复检剩下的 warn 也是素材自身的毛病，不是产品改不动。用这种素材做验收，
等于自己给自己发合格证。

这份素材的标准
--------------
**结构必须自洽**（客户真实投稿就该长这样），**格式必须全错**（否则没什么可修）：

    结构齐：题名页 / Abstract + Keywords / 三级标题 / 正文多段 / References 标题 + 6 条条目
    格式错：四边 1.25" 页边距、Arial 10pt、单倍行距、无首行缩进、
            页码在页脚居中、题名页与正文挤在同一页、参考文献无悬挂缩进

输出：tools/crossplat/fixtures/paper_apa_student.docx

内容全部是占位学术文本，不涉及任何真实作者或真实研究，可公开。

用法
----
    python tools/crossplat/make_fixture.py
"""

from __future__ import annotations

import os

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "fixtures", "paper_apa_student.docx")

TITLE = "The Impact of AI Assistants on Knowledge Work"

ABSTRACT = (
    "This study examines how AI writing assistants change the way knowledge workers "
    "plan, draft, and revise documents. Drawing on a survey of 214 professionals and "
    "a follow-up interview series, it finds that assistants shorten first-draft time "
    "while shifting effort toward review and verification. Participants reported "
    "higher output volume but mixed confidence in factual accuracy. The paper argues "
    "that the value of an assistant lies less in generation than in how well it "
    "supports judgment, and it closes with implications for tool design and for the "
    "training of new professionals."
)

KEYWORDS = "AI assistants, knowledge work, writing process, verification, productivity"

# (层级, 标题) —— Level 1 与 Level 2 都出现，验标题层级处理
HEADINGS = [
    (1, "Introduction"),
    (2, "Background"),
    (2, "Research Questions"),
    (1, "Method"),
    (2, "Participants"),
    (2, "Procedure"),
    (1, "Results"),
    (2, "Drafting Speed"),
    (2, "Accuracy Concerns"),
    (1, "Discussion"),
    (2, "Limitations"),
    (1, "Conclusion"),
]

BODY = (
    "Knowledge work has always involved more than producing text, yet most measures "
    "of productivity still count words and pages. This paragraph is representative of "
    "the body text used to exercise paragraph-level formatting: it is long enough to "
    "wrap across several lines, which is what makes first-line indent, line spacing, "
    "and font size observable in the output document."
)

REFERENCES = [
    "Brown, T., Mann, B., Ryder, N., Subbiah, M., Kaplan, J., Dhariwal, P., "
    "Neelakantan, A., Shyam, P., Sastry, G., Askell, A., Agarwal, S., Herbert-Voss, "
    "A., Krueger, G., Henighan, T., Child, R., Ramesh, A., Ziegler, D., Wu, J., "
    "Winter, C., ... Amodei, D. (2020). Language models are few-shot learners. "
    "Advances in Neural Information Processing Systems, 33, 1877-1901.",

    "Davenport, T. H., & Kirby, J. (2016). Just how smart are smart machines? "
    "MIT Sloan Management Review, 57(3), 21-25.",

    "Gawande, A. (2009). The checklist manifesto: How to get things right. "
    "Metropolitan Books.",

    "Kalliamvakou, E., Gousios, G., Blincoe, K., Singer, L., German, D. M., & "
    "Damian, D. (2016). An in-depth study of the promises and perils of mining "
    "GitHub. Empirical Software Engineering, 21(5), 2035-2071.",

    "Oppenlaender, J. (2022). The creativity of text-to-image generation. "
    "Proceedings of the 25th International Academic Mindtrek Conference, 192-202.",

    "Suchman, L. A. (2007). Human-machine reconfigurations: Plans and situated "
    "actions (2nd ed.). Cambridge University Press.",
]


def _set_font(run, name="Arial", size=10):
    """故意用 Arial 10pt —— APA 要求 Times New Roman 12pt，这里要的是"错"。"""
    run.font.name = name
    run.font.size = Pt(size)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), name)


def _plain(doc, text, bold=False, align=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    _set_font(run)
    return p


def build() -> str:
    doc = Document()

    # ---- 故意不合规的页面设置：四边 1.25"，APA 要 1" ----
    for sec in doc.sections:
        sec.top_margin = Inches(1.25)
        sec.bottom_margin = Inches(1.25)
        sec.left_margin = Inches(1.25)
        sec.right_margin = Inches(1.25)

    # ---- 默认样式也一起调坏：单倍行距、Arial 10 ----
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10)
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Inches(0)

    # ---- 题名页信息：故意和正文挤在同一页（APA 学生版题名页要独立成页） ----
    _plain(doc, TITLE, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    _plain(doc, "Jordan A. Placeholder")
    _plain(doc, "Department of Information Systems, Example University")
    _plain(doc, "BUS 610: Strategic Analysis")
    _plain(doc, "Dr. R. Placeholder")
    _plain(doc, "March 3, 2026")

    # ---- 摘要 ----
    doc.add_heading("Abstract", level=1)
    _plain(doc, ABSTRACT)
    _plain(doc, "Keywords: " + KEYWORDS)

    # ---- 正文 + 多级标题 ----
    for level, text in HEADINGS:
        doc.add_heading(text, level=level)
        _plain(doc, BODY)
        if level == 1:
            _plain(doc, BODY)

    # ---- 参考文献：有标题、有条目（这是上一版素材缺的东西） ----
    doc.add_heading("References", level=1)
    for entry in REFERENCES:
        p = _plain(doc, entry)
        # 故意不设悬挂缩进 —— APA 要求 0.5" hanging，这里要的就是"错"
        p.paragraph_format.left_indent = Inches(0)
        p.paragraph_format.first_line_indent = Inches(0)

    # ---- 页码：故意放页脚居中（APA 学生版要页眉右上） ----
    footer = doc.sections[0].footer
    fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run()
    _set_font(run)
    fld = run._element.makeelement(qn("w:fldSimple"), {qn("w:instr"): "PAGE"})
    run._element.addnext(fld)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print("written: %s (%d bytes)" % (OUT, os.path.getsize(OUT)))
    return OUT


if __name__ == "__main__":
    build()
