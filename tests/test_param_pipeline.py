"""参数全链路回归：填 -> 查 -> 修 -> 复测（v2.3.65）。

本文件守护的是「界面上有的字段，真的能跑完全程」这件事。起因是老板一句话：
> 「你再检查一下其他参数，要具体可执行」

逐项审计（_smoke/audit_params_e2e.py）过程中查出三类真问题，都在这里钉死：

1. **无 Heading 样式的文档（学生论文最常见形态）标题维度是死的**：
   docx_reader 只采 Heading 样式 → 客户填了「一级 16pt 居中」后，检查器只回一句
   「No Heading styles detected」、修正预览一行不出，而修正器其实会照改（认伪标题）。
   检查 / 预览 / 修正三方脱节。
2. **学校模板给了每级标题格式但没给层级数 → 修正链整体空转**。
3. **学校模板里的 "References" 标题被当成二级标题**，凭空造出一条「二级标题
   12pt 居中加粗」的学校要求。
"""

from __future__ import annotations

import os
import sys

import pytest

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine import checker
from src.engine.docx_reader import read_docx
from src.engine.fixer import fix_docx, _compute_changes
from src.engine.questionnaire import (
    target_from_dict, target_from_template_docx, TargetProfile,
)


# ---------------------------------------------------------------------------
# 夹具：真实学生论文形态（无 Heading 样式，章节标题是整段加粗的 Normal 段）
# ---------------------------------------------------------------------------

def _student_doc(path, *, body_font="Arial", body_size=9.0, line_spacing=1.0,
                 alignment=WD_ALIGN_PARAGRAPH.RIGHT, indent_in=0.0,
                 space_after_pt=12.0,
                 margins=(0.6, 0.6, 0.6, 0.6),
                 h1_size=14.0, h1_align=WD_ALIGN_PARAGRAPH.LEFT,
                 ref_head_bold=False, ref_head_center=False,
                 ref_font="Arial", ref_size=9.0, ref_spacing=1.5,
                 ref_hang_in=0.0, ref_new_page=False):
    doc = Document()
    for s in doc.sections:
        s.top_margin, s.bottom_margin, s.left_margin, s.right_margin = (
            Inches(margins[0]), Inches(margins[1]), Inches(margins[2]), Inches(margins[3]))

    def _body(text):
        p = doc.add_paragraph()
        r = p.add_run(text)
        r.font.name = body_font
        r.font.size = Pt(body_size)
        pf = p.paragraph_format
        pf.line_spacing = line_spacing
        pf.alignment = alignment
        pf.first_line_indent = Inches(indent_in)
        pf.space_after = Pt(space_after_pt)
        return p

    for t in ("Body paragraph one with enough words to look like real prose.",
              "Body paragraph two, still ordinary running text of the thesis.",
              "Body paragraph three, closing the introductory section."):
        _body(t)

    # 章节标题：普通段落 + 整段加粗（无 Heading 样式）
    h1 = doc.add_paragraph()
    h1.style = doc.styles["Normal"]
    hr = h1.add_run("Chapter One")
    hr.font.name = body_font
    hr.font.size = Pt(h1_size)
    hr.bold = True
    h1.paragraph_format.alignment = h1_align
    h1.paragraph_format.first_line_indent = Inches(0)

    # 参考文献：标题 + 条目
    rh = doc.add_paragraph()
    rh.paragraph_format.page_break_before = bool(ref_new_page)
    rhr = rh.add_run("References")
    rhr.font.name = body_font
    rhr.font.size = Pt(ref_size)
    rhr.bold = ref_head_bold
    rh.paragraph_format.alignment = (WD_ALIGN_PARAGRAPH.CENTER if ref_head_center
                                     else WD_ALIGN_PARAGRAPH.LEFT)
    rh.paragraph_format.first_line_indent = Inches(0)

    e = doc.add_paragraph()
    er = e.add_run("Smith, J. A. (2020). Title of the work. Publisher.")
    er.font.name = ref_font
    er.font.size = Pt(ref_size)
    e.paragraph_format.line_spacing = ref_spacing
    e.paragraph_format.left_indent = Inches(ref_hang_in)
    e.paragraph_format.first_line_indent = Inches(-ref_hang_in)
    doc.save(str(path))
    return str(path)


def _school_template(path):
    """一份「像学校模板」的 docx：正文宋体小四 1.5 倍两端对齐；
    一级标题 18pt 加粗居中（手排，无样式）；参考文献标题加粗居中 + 另起一页，
    条目 10.5pt 单倍、悬挂 0.25″。"""
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(1.25)
        s.bottom_margin = Inches(1.25)
        s.left_margin = Inches(1.5)
        s.right_margin = Inches(1.0)
    p = doc.add_paragraph()
    r = p.add_run("School template body sample text.")
    r.font.name = "SimSun"
    r.font.size = Pt(12)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Inches(0.35)

    h = doc.add_paragraph()
    hr = h.add_run("Chapter One")
    hr.font.name = "SimSun"
    hr.font.size = Pt(18)
    hr.bold = True
    h.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.paragraph_format.first_line_indent = Inches(0)

    rh = doc.add_paragraph()
    rh.paragraph_format.page_break_before = True
    rhr = rh.add_run("References")
    rhr.font.name = "SimSun"
    rhr.font.size = Pt(12)
    rhr.bold = True
    rh.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rh.paragraph_format.first_line_indent = Inches(0)

    e = doc.add_paragraph()
    er = e.add_run("Smith, J. A. (2020). Title of the work. Publisher.")
    er.font.name = "SimSun"
    er.font.size = Pt(10.5)
    e.paragraph_format.line_spacing = 1.0
    e.paragraph_format.left_indent = Inches(0.25)
    e.paragraph_format.first_line_indent = Inches(-0.25)
    doc.save(str(path))
    return str(path)


def _para_by_text(path, needle):
    for p in Document(str(path)).paragraphs:
        if needle in (p.text or ""):
            return p
    return None


# ---------------------------------------------------------------------------
# ① 无 Heading 样式：检查 / 预览 / 修正必须三方一致
# ---------------------------------------------------------------------------

def test_pseudo_heading_is_sampled_by_reader(tmp_path):
    """无 Heading 样式的文档也要能采到每级标题格式（否则检查无从比对）。"""
    src = _student_doc(tmp_path / "s.docx")
    prof = read_docx(src)
    assert prof.heading_levels_used == 0, "该文档确实没有 Heading 样式"
    assert prof.heading_formats_are_pseudo is True
    assert 1 in prof.heading_formats
    assert prof.heading_formats[1].size_pt == pytest.approx(14.0)
    assert prof.heading_formats[1].bold is True


def test_pseudo_heading_format_is_checked(tmp_path):
    """填了每级标题格式，检查报告必须给出具体差异（不能只说「没检测到标题样式」）。"""
    src = _student_doc(tmp_path / "s.docx")
    target = target_from_dict({
        "heading_levels": 1,
        "heading_formats": {1: {"alignment": "center", "bold": True,
                                "size_pt": 16.0, "font_family": "Times New Roman"}},
    })
    _, findings = checker.run_check(src, target)
    lvl_f = [f for f in findings
             if f.dimension == "heading" and f.message == "Level 1 heading format"]
    assert lvl_f, "每级标题格式必须单独出一条结论"
    assert lvl_f[0].severity == "fail"
    assert "16pt" in lvl_f[0].expected
    assert lvl_f[0].fix is not None, "不达标项必须绑定修正动作"


def test_pseudo_heading_change_is_previewed(tmp_path):
    """修正预览必须列出标题改动（否则客户点「修正」前根本不知道会改标题）。"""
    src = _student_doc(tmp_path / "s.docx")
    target = target_from_dict({
        "heading_levels": 1,
        "heading_formats": {1: {"alignment": "center", "size_pt": 16.0}},
    })
    prof, findings = checker.run_check(src, target)
    lines = _compute_changes(prof, target, lang="en", findings=findings)
    assert any("Level 1 heading" in ln for ln in lines), lines
    assert any("16pt" in ln for ln in lines), lines
    # 中文界面同样要出（且不能夹带英文标签以外的中文泄漏到英文预览）
    zh = _compute_changes(prof, target, lang="zh", findings=findings)
    assert any("第 1 级标题" in ln for ln in zh), zh


def test_pseudo_heading_is_actually_fixed(tmp_path):
    """预览说会改，就必须真的改到文档里。"""
    src = _student_doc(tmp_path / "s.docx")
    out = tmp_path / "s_out.docx"
    target = target_from_dict({
        "heading_levels": 1,
        "heading_formats": {1: {"alignment": "center", "bold": True,
                                "size_pt": 16.0, "font_family": "Times New Roman"}},
    })
    _, findings = checker.run_check(src, target)
    fix_docx(src, str(out), target, findings=findings)

    h1 = _para_by_text(out, "Chapter One")
    assert h1.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
    run = [r for r in h1.runs if (r.text or "").strip()][0]
    assert run.font.size.pt == pytest.approx(16.0)
    assert run.bold is True
    assert run.font.name == "Times New Roman"


# ---------------------------------------------------------------------------
# ② 学校模板：层级数反推 + 标题真的被改 + References 不被当二级标题
# ---------------------------------------------------------------------------

def test_template_profile_infers_heading_levels(tmp_path):
    """模板只有每级标题格式、没有 Heading 样式时，层级数必须由样本反推。"""
    tpl = _school_template(tmp_path / "tpl.docx")
    tgt = target_from_template_docx(read_docx(tpl))
    assert tgt.heading_formats, "模板里的手排标题必须被采到"
    assert tgt.heading_levels == max(tgt.heading_formats)


def test_template_reference_heading_not_taken_as_level2(tmp_path):
    """模板里的 "References"（12pt 居中加粗）不得被当成二级标题。

    否则客户选完学校模板，系统会凭空认为「学校要求二级标题 12pt 居中加粗」。
    """
    tpl = _school_template(tmp_path / "tpl.docx")
    tgt = target_from_template_docx(read_docx(tpl))
    assert 2 not in tgt.heading_formats, tgt.heading_formats
    assert tgt.heading_levels == 1


def test_template_end_to_end_all_params(tmp_path):
    """学校模板 -> 检查 -> 修正 -> 复测：模板里每一条硬要求都必须真的落到学生文档。"""
    tpl = _school_template(tmp_path / "tpl.docx")
    tgt = target_from_template_docx(read_docx(tpl))
    src = _student_doc(tmp_path / "s.docx")
    out = tmp_path / "s_out.docx"

    _, findings = checker.run_check(src, tgt)
    fix_docx(src, str(out), tgt, findings=findings)
    after = read_docx(out)

    # 页面
    assert after.margins["top"] == pytest.approx(1.25)
    assert after.margins["left"] == pytest.approx(1.5)
    assert after.font_family == "SimSun"
    assert after.font_size_pt == pytest.approx(12.0)
    assert after.line_spacing == pytest.approx(1.5)
    assert after.body_alignment == "justify"
    assert after.first_line_indent_in == pytest.approx(0.35)
    assert after.space_after_pt == pytest.approx(0.0)
    # 参考文献
    assert after.reference_hanging_indent_in == pytest.approx(0.25)
    assert after.reference_font_family == "SimSun"
    assert after.reference_size_pt == pytest.approx(10.5)
    assert after.reference_line_spacing == pytest.approx(1.0)
    assert after.reference_new_page is True
    assert after.reference_heading_bold is True
    assert after.reference_heading_centered is True
    # 标题
    h1 = _para_by_text(out, "Chapter One")
    assert h1.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert [r for r in h1.runs if (r.text or "").strip()][0].font.size.pt == pytest.approx(18.0)


def test_template_conflict_note_when_hanging_indent_differs(tmp_path):
    """模板实测与所选规范的悬挂缩进冲突时，必须进 notes 说明「以模板为准」。"""
    tpl = _school_template(tmp_path / "tpl.docx")     # 模板 0.25″
    tgt = target_from_template_docx(read_docx(tpl), reference_style="APA")  # APA 0.5″
    assert tgt.reference_hanging_indent_in == pytest.approx(0.25)
    joined = " ".join(tgt.notes)
    assert "0.25" in joined and "APA" in joined


# ---------------------------------------------------------------------------
# ③ 参考文献细项：可填 / 可查 / 可修 / 可复测
# ---------------------------------------------------------------------------

REF_CASES = [
    ("reference_font_family", "SimSun", lambda p: p.reference_font_family),
    ("reference_size_pt", 10.5, lambda p: p.reference_size_pt),
    ("reference_line_spacing", 1.0, lambda p: p.reference_line_spacing),
    ("reference_new_page", True, lambda p: p.reference_new_page),
    ("reference_heading_bold", True, lambda p: p.reference_heading_bold),
    ("reference_heading_centered", True, lambda p: p.reference_heading_centered),
]


@pytest.mark.parametrize("key,want,getter", REF_CASES)
def test_reference_detail_param_roundtrip(tmp_path, key, want, getter):
    """逐个参考文献细项：填了 → 能查到差异 → 能修 → 复测达标。"""
    src = _student_doc(tmp_path / "s.docx", ref_font="Arial", ref_size=9.0,
                       ref_spacing=1.5, ref_new_page=False,
                       ref_head_bold=False, ref_head_center=False)
    out = tmp_path / "s_out.docx"
    target = target_from_dict({key: want})

    _, findings = checker.run_check(src, target)
    target_dims = {f.dimension for f in findings if f.fix is not None}
    assert target_dims, f"{key} 填了却查不出可修项"

    fix_docx(src, str(out), target, findings=findings)
    assert getter(read_docx(out)) == want


# ---------------------------------------------------------------------------
# ④ 正文对齐 / 段后间距（审计里验过，固化成回归）
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("key,want,getter", [
    ("body_alignment", "left", lambda p: p.body_alignment),
    ("space_after_pt", 0.0, lambda p: p.space_after_pt),
    ("first_line_indent_in", 0.5, lambda p: p.first_line_indent_in),
])
def test_body_layout_param_roundtrip(tmp_path, key, want, getter):
    src = _student_doc(tmp_path / "s.docx")
    out = tmp_path / "s_out.docx"
    target = target_from_dict({key: want})
    _, findings = checker.run_check(src, target)
    assert any(f.fix is not None for f in findings), f"{key} 填了却查不出可修项"
    fix_docx(src, str(out), target, findings=findings)
    assert getter(read_docx(out)) == want
