"""MLA 9th 版式层回归测试（Purdue OWL general_format + 样例论文，均已逐字核对）。

覆盖 G1 页眉(姓氏+页码) / G2 首页信息块 / G3 标题(居中不加粗) /
G4 Works Cited(新页+居中标题+字母序) / G5 块引用(长引文应缩进0.5″)。
所有新增检查均为 info/pass 级，绝不 fail、绝不改内容。
"""

from __future__ import annotations

import pytest
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from src.engine.docx_reader import read_docx
from src.engine.checker import run_check
from src.engine.questionnaire import target_from_spec


def _add_page_field(paragraph):
    """在段落末尾插入一个 PAGE 域（用于 MLA 页眉页码）。"""
    run = paragraph.add_run()
    f1 = OxmlElement("w:fldChar")
    f1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    f2 = OxmlElement("w:fldChar")
    f2.set(qn("w:fldCharType"), "end")
    run._r.append(f1)
    run._r.append(instr)
    run._r.append(f2)


def _set_body(p):
    p.paragraph_format.line_spacing = 2.0
    p.paragraph_format.first_line_indent = Inches(0.5)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return p


def _make_mla_doc(path, *,
                  with_rh=True,
                  title_bold=False,
                  with_info_block=True,
                  block_quote_wrong=False,
                  works_cited_new_page=True,
                  works_cited_bold=False,
                  alphabetical=True,
                  header_distance_in=None):
    doc = Document()
    # 默认 Normal 字体 / 字号（python-docx 默认页边距即 1″）
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 2.0

    # 页眉 running head：姓氏 + 页码，右上
    sec = doc.sections[0]
    if header_distance_in is not None:
        sec.header_distance = Inches(header_distance_in)
    hp = sec.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if with_rh:
        hp.add_run("Doe ")
        _add_page_field(hp)

    # 首页信息块：姓名/教授/课程/日期（左上，双倍）
    if with_info_block:
        for line in ("Jane P. Doe", "Dr. Samira Hale",
                     "English 624", "12 February 2025"):
            p = doc.add_paragraph(line)
            p.paragraph_format.line_spacing = 2.0
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # 标题（居中，不加粗/不斜体/不下划线）
    title = doc.add_paragraph("Toward a Recovery of Nineteenth Century Farming Handbooks")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if title_bold:
        title.runs[0].font.bold = True

    # 正文
    for txt in (
        "While researching nineteenth century farming practices, I came across agricultural "
        "journals, newspapers, pamphlets, and brochures that contained a unique approach.",
        "Authors often placed the farming literature they were studying into a historical "
        "context by discussing the important events in agriculture of the year in question.",
    ):
        _set_body(doc.add_paragraph(txt))

    if block_quote_wrong:
        # 长引文但首行缩进、未做块引用缩进 → 应触发提示
        long_q = ("“Substantial increase in a farmer’s ability to earn income” (Danhof 5). "
                  "This improvement allowed the relations between the rural and urban populations "
                  "to strengthen, resulting in an increase in trade, and the urban population "
                  "grew rapidly after 1820, changing the structure of American life in ways that "
                  "were previously unimaginable to the farmers who tilled the soil each spring.")
        _set_body(doc.add_paragraph(long_q))

    # Works Cited
    wc = doc.add_paragraph("Works Cited")
    wc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if works_cited_bold:
        wc.runs[0].font.bold = True
    if works_cited_new_page:
        wc.paragraph_format.page_break_before = True

    entries = [
        "Allen, R.L. The American Farm Book; or Compend of American Agriculture. Saxton, 1849.",
        "Baker, Gladys L. Century of Service: The First 100 Years of the United States "
        "Department of Agriculture. Federal Government, 1996.",
    ] if alphabetical else [
        "Baker, Gladys L. Century of Service: The First 100 Years of the United States "
        "Department of Agriculture. Federal Government, 1996.",
        "Allen, R.L. The American Farm Book; or Compend of American Agriculture. Saxton, 1849.",
    ]
    for e in entries:
        ep = doc.add_paragraph(e)
        ep.paragraph_format.left_indent = Inches(0.5)
        ep.paragraph_format.first_line_indent = Inches(-0.5)
        ep.paragraph_format.line_spacing = 2.0

    doc.save(str(path))


def _find(findings, dim):
    return [f for f in findings if f.dimension == dim]


def _sev(findings, dim):
    fs = _find(findings, dim)
    return fs[0].severity if fs else None


# —— G1 页眉 ——
def test_mla_running_head_present_pass(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx)
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    assert _sev(findings, "running_head") == "pass", \
        f"running_head should pass, got {[(f.severity,f.actual) for f in _find(findings,'running_head')]}"


def test_mla_running_head_missing_info(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx, with_rh=False)
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    assert _sev(findings, "running_head") == "info"
    assert "No header text" in _find(findings, "running_head")[0].actual


# —— G3 标题 ——
def test_mla_title_format_pass(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx)
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    assert _sev(findings, "title_format") == "pass"


def test_mla_title_bold_info(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx, title_bold=True)
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    assert _sev(findings, "title_format") == "info"
    assert "bolded" in _find(findings, "title_format")[0].actual


# —— G4 Works Cited ——
def test_mla_works_cited_heading_pass(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx)
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    # 居中、不加粗 → 不应出现 reference_heading 的 info/警告
    rh = _find(findings, "reference_heading")
    assert not rh or rh[0].severity == "pass", \
        f"reference_heading unexpected: {[(f.severity,f.actual) for f in rh]}"
    assert _sev(findings, "reference_new_page") == "pass"
    assert _sev(findings, "reference_sort") == "pass"


def test_mla_works_cited_not_new_page_info(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx, works_cited_new_page=False)
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    assert _sev(findings, "reference_new_page") == "info"


def test_mla_works_cited_bold_info(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx, works_cited_bold=True)
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    # MLA 的 Works Cited 不应加粗；加粗应提示
    assert _sev(findings, "reference_heading") == "info"


def test_mla_works_cited_out_of_order_info(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx, alphabetical=False)
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    assert _sev(findings, "reference_sort") == "info"


# —— G2 首页信息块 ——
def test_mla_first_page_block_present_no_finding(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx)
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    assert _sev(findings, "first_page_heading") is None


def test_mla_first_page_block_missing_info(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx, with_info_block=False)
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    assert _sev(findings, "first_page_heading") == "info"


# —— G5 块引用 ——
def test_mla_block_quote_wrong_info(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx, block_quote_wrong=True)
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    assert _sev(findings, "block_quote") == "info"


def test_mla_block_quote_absent_no_false_positive(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx)  # 无长引文
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    assert _sev(findings, "block_quote") is None


# —— 页眉距页顶 ½″（MLA 9th：header one-half inch from the top）——
def test_mla_header_distance_correct_pass(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx, header_distance_in=0.5)
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    assert _sev(findings, "header_distance") == "pass", \
        f"header_distance should pass at 0.5\", got {[(f.severity, f.actual) for f in _find(findings, 'header_distance')]}"


def test_mla_header_distance_off_info(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx, header_distance_in=1.0)  # 偏离 ½″
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    # spec 来源的页面维度仅作提示（info），绝不 warn/fail
    assert _sev(findings, "header_distance") == "info", \
        f"header_distance off should be info, got {[(f.severity, f.actual) for f in _find(findings, 'header_distance')]}"
    assert "1.0" in _find(findings, "header_distance")[0].actual


# —— 金标准等价：完整正确的 MLA 文档零误报（无 warn/fail）——
def test_mla_full_sample_zero_false_positive(tmp_path):
    docx = tmp_path / "m.docx"
    _make_mla_doc(docx)
    prof, findings = run_check(str(docx), target_from_spec("MLA"))
    bad = [f for f in findings if f.severity in ("warn", "fail")]
    assert not bad, f"MLA sample produced warn/fail: {[(f.dimension,f.severity,f.actual) for f in bad]}"
