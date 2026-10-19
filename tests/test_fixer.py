"""一键修正引擎断言式测试（pytest）。

守护铁律：一键修正只改格式、绝不改正文内容。
- 正文段落文本修正前后逐字一致；
- 页边距 / 字体 / 字号 / 行距按 TargetProfile 生效；
- 参考文献条目数不变、核心 token（作者/年份/题名）保留；
- 绝不覆盖原件（同路径必须报错）。
"""

from __future__ import annotations

import os
import sys

import pytest

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine.questionnaire import TargetProfile, target_from_spec
from src.engine import checker
from src.engine.fixer import fix_docx, preview_fix
from src.engine.docx_reader import reference_paragraph_indices


def _make_sample(path):
    doc = Document()
    for t in ("Body paragraph one.", "Body paragraph two with more words.", "Body paragraph three."):
        p = doc.add_paragraph()
        r = p.add_run(t)
        r.font.name = "Calibri"
        r.font.size = Pt(10)
    doc.add_paragraph("References")
    doc.add_paragraph("Smith, J. (2020). Title of work. Publisher.")
    doc.add_paragraph("Lee, K. (2019). Another title. Journal.")
    for sec in doc.sections:
        sec.top_margin = Inches(2.0)
        sec.bottom_margin = Inches(2.0)
        sec.left_margin = Inches(2.0)
        sec.right_margin = Inches(2.0)
    doc.save(path)


def _body_texts(path):
    doc = Document(path)
    ref_idx = set(reference_paragraph_indices(doc))
    return [doc.paragraphs[i].text
            for i in range(len(doc.paragraphs)) if i not in ref_idx]


def _ref_texts(path):
    doc = Document(path)
    return [doc.paragraphs[i].text for i in reference_paragraph_indices(doc)]


def test_fix_preserves_body_text(tmp_path):
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    _make_sample(src)
    before = _body_texts(src)
    target = TargetProfile(
        page={"margin_top_in": 1.25, "margin_bottom_in": 1.25,
              "margin_left_in": 1.0, "margin_right_in": 1.0,
              "font_family": "Arial", "font_size_pt": 11, "line_spacing": 1.5},
        reference_style="APA", reference_hanging_indent_in=0.5, reference_numbered=False,
    )
    fix_docx(str(src), str(out), target)
    assert _body_texts(out) == before


def test_fix_applies_margins(tmp_path):
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    _make_sample(src)
    target = TargetProfile(page={"margin_top_in": 1.25, "margin_bottom_in": 1.25,
                                 "margin_left_in": 1.0, "margin_right_in": 1.0})
    fix_docx(str(src), str(out), target)
    doc = Document(str(out))
    for sec in doc.sections:
        assert sec.top_margin.inches == pytest.approx(1.25)
        assert sec.bottom_margin.inches == pytest.approx(1.25)
        assert sec.left_margin.inches == pytest.approx(1.0)
        assert sec.right_margin.inches == pytest.approx(1.0)


def test_fix_applies_font(tmp_path):
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    _make_sample(src)
    target = TargetProfile(page={"font_family": "Arial", "font_size_pt": 11})
    fix_docx(str(src), str(out), target)
    doc = Document(str(out))
    for p in doc.paragraphs:
        for r in p.runs:
            if r.text.strip():
                assert r.font.name == "Arial"
                assert r.font.size == Pt(11)


def test_fix_applies_line_spacing(tmp_path):
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    _make_sample(src)
    target = TargetProfile(page={"line_spacing": 1.5})
    fix_docx(str(src), str(out), target)
    doc = Document(str(out))
    for p in doc.paragraphs:
        assert p.paragraph_format.line_spacing == pytest.approx(1.5)
        # 必须是倍数制（命名 1.0/1.5/2.0 或 MULTIPLE），不能是固定值/最小值
        assert p.paragraph_format.line_spacing_rule in (
            WD_LINE_SPACING.MULTIPLE, WD_LINE_SPACING.SINGLE,
            WD_LINE_SPACING.ONE_POINT_FIVE, WD_LINE_SPACING.DOUBLE,
        )


def test_fix_reference_count_unchanged(tmp_path):
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    _make_sample(src)
    before = reference_paragraph_indices(Document(str(src)))
    target = TargetProfile(reference_style="IEEE", reference_hanging_indent_in=0.0,
                           reference_numbered=True)
    fix_docx(str(src), str(out), target)
    after = reference_paragraph_indices(Document(str(out)))
    assert len(before) == len(after) == 2


def test_fix_reference_tokens_preserved_numbered(tmp_path):
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    _make_sample(src)
    target = TargetProfile(reference_style="IEEE", reference_hanging_indent_in=0.0,
                           reference_numbered=True)
    fix_docx(str(src), str(out), target)
    joined = "\n".join(_ref_texts(str(out)))
    for tok in ("Smith", "2020", "Title of work", "Lee", "2019", "Another title"):
        assert tok in joined
    refs = _ref_texts(str(out))
    assert refs[0].startswith("[1]")
    assert refs[1].startswith("[2]")


def test_fix_reference_tokens_preserved_author_year(tmp_path):
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    _make_sample(src)
    target = TargetProfile(reference_style="APA", reference_hanging_indent_in=0.5,
                           reference_numbered=False)
    fix_docx(str(src), str(out), target)
    joined = "\n".join(_ref_texts(str(out)))
    for tok in ("Smith", "2020", "Title of work"):
        assert tok in joined
    # 非编号制不应加 [n] 前缀
    assert not _ref_texts(str(out))[0].lstrip().startswith("[")


def test_fix_writes_new_file_not_overwrite(tmp_path):
    src = tmp_path / "in.docx"
    _make_sample(src)
    target = TargetProfile(page={"font_family": "Arial", "font_size_pt": 11})
    with pytest.raises(ValueError):
        fix_docx(str(src), str(src), target)  # 同路径必须报错
    out = tmp_path / "out.docx"
    before = _body_texts(str(src))
    fix_docx(str(src), str(out), target)
    assert _body_texts(str(src)) == before  # 原件未被改
    assert out.exists()


def _make_apa_titlepage(path):
    """构造一篇带标题页的 APA 论文：标题 / 作者 / 单位 / Abstract / Keywords / 正文。"""
    doc = Document()
    doc.add_paragraph().add_run(
        "The relationship between sleep quality and cognitive performance in young adults")
    doc.add_paragraph().add_run("Avery N. Campbell")
    doc.add_paragraph().add_run("Department of Psychology, Riverside State University")
    doc.add_paragraph().add_run("Abstract. Emerging evidence suggests a close link.")
    doc.add_paragraph().add_run("Keywords: sleep, quality, cognition")
    for txt in ("Introduction", "Body paragraph one.", "Body paragraph two."):
        doc.add_paragraph().add_run(txt)
    doc.save(path)


def test_fix_apa_title_page_centers_and_bolds_title(tmp_path):
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    _make_apa_titlepage(src)
    target = TargetProfile(spec_key="APA",
                           page={"font_family": "Times New Roman",
                                 "font_size_pt": 12, "line_spacing": 2.0})
    fix_docx(str(src), str(out), target)
    doc = Document(str(out))
    # 标题（p0）：居中 + 加粗（此前漏掉，用户实测标题仍是左对齐未加粗）
    assert doc.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert doc.paragraphs[0].runs[0].bold is True
    # 作者（p1）/ 单位（p2）：居中（不加粗）
    assert doc.paragraphs[1].alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert doc.paragraphs[2].alignment == WD_ALIGN_PARAGRAPH.CENTER
    # Abstract（p3）：标题页信号之后，绝不被当成标题居中（正文对齐由 body 规则决定）
    assert doc.paragraphs[3].alignment != WD_ALIGN_PARAGRAPH.CENTER


def test_fix_no_title_page_signal_leaves_first_para_alone(tmp_path):
    """无标题页正面信号（无 Abstract/Keywords）时，不得盲猜首段为标题去居中/加粗。"""
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    doc = Document()
    doc.add_paragraph().add_run("Introduction")          # 首段是章节标题样式文本
    doc.add_paragraph().add_run("Body text here.")
    doc.add_paragraph("References")
    doc.add_paragraph("Smith, J. (2020). Title. Publisher.")
    doc.save(src)
    target = TargetProfile(spec_key="APA",
                           page={"font_family": "Times New Roman", "font_size_pt": 12})
    fix_docx(str(src), str(out), target)
    out_doc = Document(str(out))
    # 首段既非已识别标题、又无标题页信号 → 对齐保持原样，不被居中
    assert out_doc.paragraphs[0].alignment in (None, WD_ALIGN_PARAGRAPH.LEFT)


def test_fix_apa_applies_left_align_and_halfinch_indent(tmp_path):
    """APA 7th 硬性正文规则：左对齐 + 首行缩进 0.5″（§2.22 / §2.24）。

    此前规范 PageLayout 未设 body_alignment / first_line_indent_in，导致 checker 与
    fixer 对这两项「零覆盖」。补入后：正文应改为 left + 0.5″，标题（APA L1 居中加粗）不受影响。
    """
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    doc = Document()
    doc.add_heading("Method", level=1)          # 应被格式化为 APA L1（居中加粗）
    for t in ("Body one.", "Body two.", "Body three."):
        p = doc.add_paragraph(t)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Inches(0)
    doc.add_paragraph("References")
    doc.save(src)

    target = target_from_spec("APA-Student")
    assert target.page.get("body_alignment") == "left"
    assert target.page.get("first_line_indent_in") == 0.5

    fix_docx(str(src), str(out), target)
    o = Document(str(out))
    head = o.paragraphs[0]
    assert head.alignment == WD_ALIGN_PARAGRAPH.CENTER   # L1 不被正文对齐覆盖
    assert head.runs[0].bold is True
    for p in o.paragraphs[1:4]:
        assert p.alignment == WD_ALIGN_PARAGRAPH.LEFT
        fl = p.paragraph_format.first_line_indent
        assert fl is not None and abs(fl.inches - 0.5) < 1e-6


def test_fix_apa_heading_levels_land_l1_to_l5(tmp_path):
    """APA 7th 五级标题须按 §2.26–2.27 落盘为直接格式：L1 居中加粗 / L2 左对齐加粗 /
    L3 左对齐加粗斜体 / L4 缩进加粗接排(句点) / L5 缩进加粗斜体接排(句点)。

    修正引擎把标题格式写成「直接格式」而非 Heading 样式，故 python-docx 可直接读出，
    这里即头less 抽检 L1–L5 是否真的落地正确（商业级验收点）。"""
    src = tmp_path / "h.docx"
    out = tmp_path / "h_out.docx"
    doc = Document()
    specs = [(1, "First Level"), (2, "Second Level"), (3, "Third Level"),
             (4, "Fourth Level"), (5, "Fifth Level")]
    for lvl, txt in specs:
        p = doc.add_paragraph()
        p.style = doc.styles[f"Heading {lvl}"]
        p.add_run(txt)
    doc.save(src)

    fix_docx(str(src), str(out), target_from_spec("APA-Student"))
    o = Document(str(out))
    heading_texts = {t for _lvl, t in specs}
    by = {}
    for p in o.paragraphs:
        t = p.text.strip().rstrip(".")   # 接排标题被补句点，匹配时剥掉
        if t in heading_texts:
            by[t] = p
    assert len(by) == 5, f"expected 5 headings, got {list(by)}"

    def runs_bold_italic(p):
        return (any(r.font.bold is True for r in p.runs),
                any(r.font.italic is True for r in p.runs))

    # L1
    p = by["First Level"]
    assert p.alignment == WD_ALIGN_PARAGRAPH.CENTER
    b, i = runs_bold_italic(p)
    assert b and not i and not p.text.strip().endswith(".")
    # L2
    p = by["Second Level"]
    assert p.alignment == WD_ALIGN_PARAGRAPH.LEFT
    b, i = runs_bold_italic(p)
    assert b and not i and not p.text.strip().endswith(".")
    # L3
    p = by["Third Level"]
    assert p.alignment == WD_ALIGN_PARAGRAPH.LEFT
    b, i = runs_bold_italic(p)
    assert b and i and not p.text.strip().endswith(".")
    # L4 (run-in, indent 0.5", trailing period)
    p = by["Fourth Level"]
    assert p.alignment == WD_ALIGN_PARAGRAPH.LEFT
    b, i = runs_bold_italic(p)
    assert b and not i and p.text.strip().endswith(".")
    fl = p.paragraph_format.first_line_indent
    assert fl is not None and abs(fl.inches - 0.5) < 1e-6
    # L5 (run-in, indent 0.5", italic, trailing period)
    p = by["Fifth Level"]
    assert p.alignment == WD_ALIGN_PARAGRAPH.LEFT
    b, i = runs_bold_italic(p)
    assert b and i and p.text.strip().endswith(".")
    fl = p.paragraph_format.first_line_indent
    assert fl is not None and abs(fl.inches - 0.5) < 1e-6


def _make_ieee_titlepage(path):
    """构造一篇带标题页的 IEEE 论文：标题 / 作者 / 单位 / Abstract / Keywords / 正文。"""
    doc = Document()
    doc.add_paragraph().add_run(
        "A Neural Network Approach to Low-Power Signal Classification")
    doc.add_paragraph().add_run("Jordan M. Rivera")
    doc.add_paragraph().add_run("Department of Electrical Engineering, Lakeside University")
    doc.add_paragraph().add_run("Abstract. We propose a low-power classifier.")
    doc.add_paragraph().add_run("Keywords: neural network, signal, low-power")
    for txt in ("Introduction", "Body paragraph one.", "Body paragraph two."):
        doc.add_paragraph().add_run(txt)
    doc.save(path)


def test_fix_ieee_title_page_centers_and_bolds_title(tmp_path):
    """IEEE 标题页：标题居中加粗、作者/单位居中（spec.title_page 显式声明，不再靠 L1 启发）。"""
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    _make_ieee_titlepage(src)
    target = target_from_spec("IEEE")
    fix_docx(str(src), str(out), target)
    doc = Document(str(out))
    # 标题（p0）：居中 + 加粗
    assert doc.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert doc.paragraphs[0].runs[0].bold is True
    # 作者（p1）/ 单位（p2）：居中
    assert doc.paragraphs[1].alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert doc.paragraphs[2].alignment == WD_ALIGN_PARAGRAPH.CENTER
    # Abstract（p3）：标题页信号之后，绝不被当成标题居中（正文对齐由 body 规则决定）
    assert doc.paragraphs[3].alignment != WD_ALIGN_PARAGRAPH.CENTER


def _make_ieee_headings_doc(path):
    """构造一篇用 Heading 样式标好层级的 IEEE 文稿（无标题页信号）。"""
    doc = Document()
    doc.add_paragraph("Introduction", style="Heading 1")
    doc.add_paragraph("Background", style="Heading 2")
    doc.add_paragraph("Prior work", style="Heading 3")
    doc.add_paragraph("Methods", style="Heading 1")
    doc.add_paragraph("Sampling", style="Heading 2")
    for t in ("Body paragraph one.", "Body paragraph two."):
        doc.add_paragraph().add_run(t)
    doc.save(path)


def test_fix_ieee_headings_numbered_and_bolded(tmp_path):
    """IEEE 章节标题：罗马/字母/数字编号 + L1 全大写 + 加粗，且遇更浅层重置更深层编号。

    对齐 IEEE_Paper_Format_Rules V1.0 §3：仅 L1 全大写（"I. INTRODUCTION"），
    L2–L4 保留作者原大小写（"A. Background"），L3 为 "1)" 右括号编号。
    """
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    _make_ieee_headings_doc(src)
    target = target_from_spec("IEEE")
    fix_docx(str(src), str(out), target)
    doc = Document(str(out))
    # H1: I. INTRODUCTION — 居中 + 加粗 + 全大写
    assert doc.paragraphs[0].text == "I. INTRODUCTION"
    assert doc.paragraphs[0].runs[0].bold is True
    assert doc.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER
    # H2: A. Background — 左对齐 + 加粗 + 保留原大小写（不全大写）
    assert doc.paragraphs[1].text == "A. Background"
    assert doc.paragraphs[1].alignment == WD_ALIGN_PARAGRAPH.LEFT
    # H3: 1) Prior work（数字 + 右括号编号）
    assert doc.paragraphs[2].text == "1) Prior work"
    # 第二个 H1 → II. METHODS（罗马递增）
    assert doc.paragraphs[3].text == "II. METHODS"
    # 其下 H2 重置为 A. Sampling（不是 B.）
    assert doc.paragraphs[4].text == "A. Sampling"


def test_fix_ieee_headings_idempotent_on_rerun(tmp_path):
    """重跑不应叠加编号前缀（I. I. ...）；用独立输出文件避免原地改写。"""
    src = tmp_path / "in.docx"
    out1 = tmp_path / "out1.docx"
    out2 = tmp_path / "out2.docx"
    _make_ieee_headings_doc(src)
    target = target_from_spec("IEEE")
    fix_docx(str(src), str(out1), target)
    fix_docx(str(out1), str(out2), target)          # 第二次跑在 out1 产物上
    doc = Document(str(out2))
    assert doc.paragraphs[0].text == "I. INTRODUCTION"
    assert doc.paragraphs[3].text == "II. METHODS"


def _make_ieee_apa_refs_doc(path):
    """构造一篇含 APA 作者-年参考文献的文稿（IEEE 应 best-effort 改写为编号引文）。"""
    doc = Document()
    doc.add_paragraph("References", style="Heading 1")
    doc.add_paragraph().add_run(
        "Anderson, K., & Lee, S. (2019). Sleep quality and memory consolidation. "
        "Journal of Sleep Research, 12(3), 45-67.")
    doc.add_paragraph().add_run("Smith, J. A. (2020). Title of Work. Publisher.")
    doc.save(path)


def test_fix_ieee_references_rewritten_to_ieee_style(tmp_path):
    """IEEE 参考文献：APA 作者-年条目 best-effort 改写为 IEEE 编号引文（作者首字母+引号+vol/no/pp+年）。"""
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    _make_ieee_apa_refs_doc(src)
    target = target_from_spec("IEEE")
    fix_docx(str(src), str(out), target)
    doc = Document(str(out))
    # 期刊条目 → IEEE 编号引文 + [1] 前缀
    assert doc.paragraphs[1].text == (
        '[1] K. Anderson and S. Lee, "Sleep quality and memory consolidation," '
        "Journal of Sleep Research, vol. 12, no. 3, pp. 45-67, 2019.")
    # 书条目 → IEEE 书引文 + [2] 前缀
    assert doc.paragraphs[2].text == (
        "[2] J. A. Smith, Title of Work. Publisher, 2020.")


def test_fix_ieee_references_rewrite_idempotent(tmp_path):
    """IEEE 参考文献重写幂等：重跑不产生双重编号 / 不二次改写；用独立输出文件。"""
    src = tmp_path / "in.docx"
    out1 = tmp_path / "out1.docx"
    out2 = tmp_path / "out2.docx"
    _make_ieee_apa_refs_doc(src)
    target = target_from_spec("IEEE")
    fix_docx(str(src), str(out1), target)
    fix_docx(str(out1), str(out2), target)          # 第二次跑在 out1 产物上
    doc = Document(str(out2))
    assert doc.paragraphs[1].text == (
        '[1] K. Anderson and S. Lee, "Sleep quality and memory consolidation," '
        "Journal of Sleep Research, vol. 12, no. 3, pp. 45-67, 2019.")
    # 不应出现 [1][1] 或双重 vol.
    assert "[1][1]" not in doc.paragraphs[1].text
    assert doc.paragraphs[1].text.count("vol.") == 1



# ---------------------------------------------------------------------------
# 真实样例回归（2026-10-06）：用户样例 sample-paper-apa_fixed21.docx 暴露的三类真 bug
# ---------------------------------------------------------------------------


def _make_real_ieee_doc(path):
    """复刻用户真实样例的形态：标题页 + 12pt 整段加粗的伪标题 + APA 参考文献。

    关键点：章节标题**不是 Heading 样式**，而是 12pt、整段加粗的 Normal 段（真实学生
    论文最常见形态）。旧实现要求 ≥13pt 才认伪标题，于是这些标题既不被编号也不被保护。
    """
    doc = Document()
    doc.add_paragraph("The relationship between sleep quality and cognitive "
                      "performance in young adults")
    doc.add_paragraph("Avery N. Campbell")
    doc.add_paragraph("Department of Psychology, Riverside State University")
    doc.add_paragraph("Abstract. Emerging evidence suggests that sleep quality is "
                      "closely linked to cognitive performance.")
    doc.add_paragraph("Keywords: sleep quality, cognitive performance, young adults, "
                      "systematic review")
    for text in ("Introduction", "Literature Review", "Method", "Results",
                 "Discussion", "Conclusion", "References"):
        p = doc.add_paragraph(text)
        for r in p.runs:                     # 12pt 整段加粗、无 Heading 样式
            r.bold = True
            r.font.size = Pt(12)
    doc.add_paragraph("1. Anderson, K., & Lee, S. (2019). Sleep duration and memory in "
                      "college students. Journal of Sleep Research, 12(3), 45-67.")
    doc.add_paragraph("2. Becker, T., Muller, R., & Schmidt, F. (2021). Executive "
                      "function. Journal of Educational Psychology, 113(4), 612-630.")
    doc.save(str(path))
    return str(path)


def test_fix_ieee_realistic_bold_normal_headings_numbered(tmp_path):
    """真实样例形态（12pt 加粗 Normal 伪标题）也必须被识别 + 编号（I./II./…）。

    守护 2026-10-06 修复：旧阈值 13pt 让 12pt 加粗标题整体漏判 →
    headings_formatted=0，标题既不编号又被正文格式污染。
    """
    src = _make_real_ieee_doc(tmp_path / "real.docx")
    out = tmp_path / "real_out.docx"
    result = fix_docx(src, str(out), target_from_spec("IEEE"))
    assert result["headings_formatted"] >= 7, "12pt 加粗伪标题应被识别并格式化"
    texts = [p.text for p in Document(str(out)).paragraphs]
    assert "I. INTRODUCTION" in texts
    assert "II. LITERATURE REVIEW" in texts
    # References 是后置部分：按 IEEE 惯例**不编号**，只套格式
    assert "REFERENCES" in texts
    assert not any(t.startswith("VII. REFERENCES") for t in texts)


def test_fix_ieee_realistic_doc_is_idempotent(tmp_path):
    """整篇跑两遍，第二遍必须是「什么都没改」。

    守护 2026-10-06 修复的两个隐形打架：
    ① 标题保护用 is_heading(13pt)、标题格式化用 _assign_heading_levels(10pt)，
       两套判定不一致 → 正文先刷缩进、标题再改回来，每遍恒报「已改」；
    ② 标题页段落被正文档对齐(justify)+首行缩进刷一遍，标题页流程再改回来，同样恒报。
    两者都会让用户「跑第二遍才干净」，且汇总批注满嘴假话。
    """
    src = _make_real_ieee_doc(tmp_path / "real.docx")
    out1 = tmp_path / "o1.docx"
    out2 = tmp_path / "o2.docx"
    fix_docx(src, str(out1), target_from_spec("IEEE"))
    result2 = fix_docx(str(out1), str(out2), target_from_spec("IEEE"))
    assert result2["headings_formatted"] == 0, "第二遍不应再改标题"
    assert result2["title_page_formatted"] == 0, "第二遍不应再改标题页"
    assert result2["reference_numbering_added"] == 0, "第二遍不应重复编号"
    assert result2.get("applied") in ({}, None), "第二遍不应有任何页面级改动"
    # 文本也必须逐段一致
    a = [p.text for p in Document(str(out1)).paragraphs]
    b = [p.text for p in Document(str(out2)).paragraphs]
    assert a == b


def test_ieee_reference_rewrite_requires_explicit_opt_in(tmp_path):
    """参考文献内容改写必须显式开启：裸 TargetProfile(reference_style='IEEE') 不得改字。

    守护产品铁律「只改格式、不改内容」：APA→IEEE 引文改写属内容层改动，
    不能只凭 reference_style 字符串就触发（否则任何单参调用/单测都被改字）。
    """
    src = tmp_path / "ref.docx"
    _make_ieee_apa_refs_doc(src)
    out = tmp_path / "ref_out.docx"
    before = Document(src).paragraphs[1].text
    fix_docx(src, str(out),
             TargetProfile(reference_style="IEEE", reference_hanging_indent_in=0.0,
                           reference_numbered=True))
    after = Document(str(out)).paragraphs[1].text
    assert after.startswith("[1] "), "仍应补编号前缀"
    assert after[len("[1] "):] == before, "未显式开启时，条目文字必须逐字不动"
    # 而真实产品路径（target_from_spec）会开启改写
    assert target_from_spec("IEEE").reference_rewrite is True
    assert target_from_spec("APA").reference_rewrite is False


def test_ieee_reference_rewrite_handles_multi_author(tmp_path):
    """一个 & 段里有多位作者（'Becker, T., Muller, R., & Schmidt, F.'）也要正确重写。"""
    src = tmp_path / "multi.docx"
    doc = Document()
    doc.add_paragraph("Body paragraph for the thesis.")
    doc.add_paragraph("References")
    doc.add_paragraph("Becker, T., Muller, R., & Schmidt, F. (2021). Executive "
                      "function. Journal of Educational Psychology, 113(4), 612-630.")
    doc.save(src)
    out = tmp_path / "multi_out.docx"
    fix_docx(src, str(out), target_from_spec("IEEE"))
    text = Document(str(out)).paragraphs[2].text
    assert text.startswith("[1] T. Becker, R. Muller, and F. Schmidt,"), text
    assert "vol. 113, no. 4, pp. 612-630, 2021." in text
    # 绝不能出现把编号吞进姓氏的垃圾
    assert "1. Becker" not in text and "K. 1." not in text


def test_spec_key_none_reference_style_ieee_resolves_same_spec(tmp_path):
    """钉死「检查/修正规范还不一样」漂移 bug：spec_key=None 但 reference_style='IEEE' 时，
    checker 与 fixer 必须解析到同一个 IEEE 规范（修复前 checker 用 ref_style、fixer 用
    'Other'，编号制漏判 → 预览说要改、修了却没改 / 反之）。

    单一真相源：两侧都走 resolve_spec(target) = get_spec(spec_key or reference_style or "Other")。
    """
    from src.engine.checker import resolve_spec
    from src.engine.specs import get_spec
    from docx import Document

    target = TargetProfile(spec_key=None, reference_style="IEEE",
                           reference_hanging_indent_in=0.0, reference_numbered=True)
    # checker/fixer 都把 reference_style 当回退键 → 解析到同一份 IEEE 规范
    assert resolve_spec(target).reference.style_key == "IEEE"
    # 构造一篇参考文献未编号的文稿
    src = tmp_path / "r.docx"
    doc = Document()
    doc.add_paragraph("Body text for the thesis.")
    doc.add_paragraph("References")
    doc.add_paragraph("Smith, J. (2020). Title of work. Publisher.")
    doc.save(str(src))
    # checker 据此产出的 references finding 必须绑定编号修正动作，
    # 否则 fixer 不会补 [n]，预览与落盘就脱节了
    _prof, findings = checker.run_check(str(src), target)
    numbered = [f for f in findings if getattr(f, "fix", None)
                and f.fix.dimension == "reference_numbered"]
    assert numbered, "IEEE 编号制须绑定 reference_numbered 修正，否则 fixer 不会补 [n]"


def test_ieee_title_page_never_numbered_as_section(tmp_path):
    """标题页标题（被加粗后）绝不能被当成章节标题编号。

    守护 2026-10-06 修复：标题页流程会把大标题加粗，而加粗 + 短文本恰好满足伪标题
    条件 → 第二遍把论文大标题编成 "I."，并把后续章节编号整体推一位，产出
    "II. I. INTRODUCTION" 这种脏结果（用户肉眼看不出编号错位，只会以为软件改坏了）。
    构造时标题刻意取 80 字以内（伪标题长度上限内），确保真的能触发该误判。
    """
    doc = Document()
    doc.add_paragraph("Sleep quality and cognitive performance in young adults")
    doc.add_paragraph("Avery N. Campbell")
    doc.add_paragraph("Department of Psychology, Riverside State University")
    doc.add_paragraph("Abstract. Sleep quality is closely linked to cognition.")
    doc.add_paragraph("Keywords: sleep quality, cognitive performance, adults")
    for text in ("Introduction", "Method", "References"):
        p = doc.add_paragraph(text)
        for r in p.runs:
            r.bold = True
            r.font.size = Pt(12)
    src = tmp_path / "tp.docx"
    doc.save(str(src))
    out1, out2 = tmp_path / "o1.docx", tmp_path / "o2.docx"
    fix_docx(str(src), str(out1), target_from_spec("IEEE"))
    fix_docx(str(out1), str(out2), target_from_spec("IEEE"))
    texts = [p.text for p in Document(str(out2)).paragraphs]
    assert texts[0] == "Sleep quality and cognitive performance in young adults", \
        f"标题页标题被当成章节标题改了：{texts[0]!r}"
    assert "I. INTRODUCTION" in texts
    assert "II. METHOD" in texts
    assert not any("I. I." in t or "II. I." in t for t in texts), "出现编号错位"
