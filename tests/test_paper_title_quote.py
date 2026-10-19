"""纸张大小（Letter / A4）+ 标题页 + 块引用 三条链路回归（v2.3.65）。

三条链路的共同点：都是**通用格式维度，不归属任何引用规范**，且都曾经「只做半截」。

1. **纸张大小**：北美用 Letter（8.5×11 in），欧洲/中国/澳洲用 A4（210×297 mm）。
   同一篇 APA 论文在纽约交是 Letter、在伦敦交是 A4 —— 所以「规范给的只是典型值参考」，
   真正说了算的是学校模板/客户。以前引擎**完全不认纸型**：客户把 A4 稿交给要求 Letter
   的学校，工具一声不吭。
2. **标题页**：修正器里有一条**不受检查驱动**的流程会悄悄改标题页（客户看不到、报告里
   没有），且客户无处填「标题要不要居中/加粗」。
3. **块引用**：检查按「长引文」判、修正只认 Quote 样式 → 客户看到问题，点完修正一处没动。
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
from src.engine.docx_reader import read_docx, identify_paper_size
from src.engine.fixer import fix_docx, _compute_changes
from src.engine.questionnaire import (
    target_from_dict, target_from_spec, target_from_template_docx,
)


# ---------------------------------------------------------------------------
# 夹具
# ---------------------------------------------------------------------------

def _doc(path, paper=("Letter", 8.5, 11.0), landscape=False,
         title_center=False, title_bold=False, author_center=False,
         with_title_page=True):
    doc = Document()
    s = doc.sections[0]
    w, h = paper[1], paper[2]
    if landscape:
        w, h = h, w
    s.page_width, s.page_height = Inches(w), Inches(h)
    t = doc.add_paragraph()
    t.add_run("The relationship between sleep quality and cognitive performance")
    if title_center:
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if title_bold:
        for r in t.runs:
            r.bold = True
    a1 = doc.add_paragraph()
    a1.add_run("Avery N. Campbell")
    a2 = doc.add_paragraph()
    a2.add_run("Department of Psychology, Riverside State University")
    if author_center:
        a1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        a2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if with_title_page:
        doc.add_paragraph().add_run("Abstract. Emerging evidence suggests a close link.")
        doc.add_paragraph().add_run("Keywords: sleep, quality, cognition")
    doc.add_paragraph().add_run("Introduction")
    doc.add_paragraph().add_run("Body paragraph one, ordinary running text of the paper.")
    doc.save(str(path))
    return str(path)


def _doc_with_quotation(path):
    doc = Document()
    doc.add_paragraph().add_run("Body paragraph one, ordinary text of the paper.")
    q = doc.add_paragraph()
    q.add_run('\u201cThis is a very long quotation that clearly runs well beyond four lines of '
              'text and therefore ought to be set as a block quotation with a half inch indent '
              'according to the style guide that the student follows, because it exceeds the '
              'four line threshold used by most guides when deciding whether a quotation has to '
              'be set off from the surrounding body text as a separate block.\u201d')
    doc.add_paragraph().add_run("Body paragraph after the quotation.")
    doc.save(str(path))
    return str(path)


# ---------------------------------------------------------------------------
# ① 纸张大小
# ---------------------------------------------------------------------------

def test_identify_paper_size_core_sizes():
    assert identify_paper_size(8.5, 11.0) == "Letter"
    assert identify_paper_size(8.27, 11.69) == "A4"
    assert identify_paper_size(8.28, 11.69) == "A4"        # Word 常见舍入
    assert identify_paper_size(11.0, 8.5) == "Letter"      # 横向归一
    assert identify_paper_size(7.5, 10.0) is None          # 非标准尺寸不臆造


def test_paper_size_read_from_docx(tmp_path):
    p = _doc(tmp_path / "a4.docx", paper=("A4", 8.27, 11.69))
    prof = read_docx(p)
    assert prof.paper_size == "A4"
    assert prof.paper_width_in == pytest.approx(8.27)
    assert prof.paper_height_in == pytest.approx(11.69)


def test_paper_size_mismatch_is_reported_and_previewed(tmp_path):
    """Letter 稿 vs 要求 A4：必须报告 + 进预览（不能悄悄不管）。"""
    p = _doc(tmp_path / "letter.docx", paper=("Letter", 8.5, 11.0))
    target = target_from_dict({"paper_size": "A4"})
    prof, findings = checker.run_check(p, target)
    f = [x for x in findings if x.dimension == "paper_size"]
    assert f and f[0].severity == "warn"
    assert f[0].expected == "A4" and "Letter" in f[0].actual
    assert f[0].fix is not None
    lines = _compute_changes(prof, target, lang="en", findings=findings)
    assert any("Paper size" in ln and "A4" in ln for ln in lines), lines


def test_paper_size_is_actually_converted(tmp_path):
    p = _doc(tmp_path / "letter.docx", paper=("Letter", 8.5, 11.0))
    out = tmp_path / "a4.docx"
    target = target_from_dict({"paper_size": "A4"})
    _, findings = checker.run_check(p, target)
    fix_docx(p, str(out), target, findings=findings)
    prof = read_docx(str(out))
    assert prof.paper_size == "A4"
    assert prof.paper_width_in == pytest.approx(8.27, abs=0.02)
    assert prof.paper_height_in == pytest.approx(11.69, abs=0.02)


def test_paper_size_conversion_keeps_landscape(tmp_path):
    """横向 A4 改成 Letter 后必须仍是横向，不能被打回纵向（否则版面全乱）。"""
    p = _doc(tmp_path / "ls.docx", paper=("A4", 8.27, 11.69), landscape=True)
    prof = read_docx(p)
    assert prof.paper_size == "A4"
    out = tmp_path / "ls_letter.docx"
    target = target_from_dict({"paper_size": "Letter"})
    _, findings = checker.run_check(p, target)
    fix_docx(p, str(out), target, findings=findings)
    after = read_docx(str(out))
    assert after.paper_size == "Letter"
    assert after.paper_width_in > after.paper_height_in, "必须保持横向"


def test_paper_size_from_spec_is_only_a_hint(tmp_path):
    """选 APA（规范典型值 Letter）不该因为客户是 A4 稿就判红 —— 欧洲学生也用 APA。"""
    p = _doc(tmp_path / "a4.docx", paper=("A4", 8.27, 11.69))
    _, findings = checker.run_check(p, target_from_spec("APA"))
    f = [x for x in findings if x.dimension == "paper_size"]
    assert f and f[0].severity == "info", f[0].severity if f else "missing"


def test_paper_size_from_school_template(tmp_path):
    """学校模板的纸型必须带进目标画像（通用项，模板说了算）。"""
    tpl = _doc(tmp_path / "tpl.docx", paper=("A4", 8.27, 11.69))
    tgt = target_from_template_docx(read_docx(tpl))
    assert tgt.paper_size == "A4"
    assert tgt.resolution.get("paper_size") == "school"


def test_paper_size_nonstandard_goes_to_notes(tmp_path):
    """模板用非标纸（如 8.5×13）时不硬塞成 Letter/A4，而是写进提示。"""
    tpl = _doc(tmp_path / "tpl.docx", paper=("X", 8.5, 13.0))
    tgt = target_from_template_docx(read_docx(tpl))
    assert tgt.paper_size is None
    assert any("not a standard" in n for n in tgt.notes), tgt.notes


def test_ai_template_carries_paper_size():
    """AI 模板必须能填纸型（且明确写清「与引用规范无关」）。"""
    assert target_from_dict({"paper_size": "A4"}).paper_size == "A4"
    assert target_from_dict({"paper_size": "a4"}).paper_size == "A4"      # 大小写归一
    # 白名单外的值（B5 不在支持列表）必须被丢弃；只剩它 → 整份画像为空 → 明确报错，
    # 绝不静默当成「已设置」把不认识的纸型塞进检查。
    with pytest.raises(ValueError):
        target_from_dict({"paper_size": "B5"})


# ---------------------------------------------------------------------------
# ② 标题页
# ---------------------------------------------------------------------------

def test_title_page_reported_and_previewed(tmp_path):
    """客户填了标题页三项 → 要报告、要进预览、要能修。"""
    p = _doc(tmp_path / "tp.docx")
    target = target_from_dict({"title_page_title_centered": True,
                               "title_page_title_bold": True,
                               "title_page_author_centered": True})
    prof, findings = checker.run_check(p, target)
    f = [x for x in findings if x.dimension == "title_page"]
    assert f and f[0].severity == "fail"
    assert "not centered" in f[0].actual and "not bold" in f[0].actual
    assert f[0].fix is not None, "标题页不达标必须绑定修正动作"
    lines = _compute_changes(prof, target, lang="en", findings=findings)
    assert any("Title page" in ln for ln in lines), lines


def test_title_page_is_actually_fixed(tmp_path):
    p = _doc(tmp_path / "tp.docx")
    out = tmp_path / "tp_out.docx"
    target = target_from_dict({"title_page_title_centered": True,
                               "title_page_title_bold": True,
                               "title_page_author_centered": True})
    _, findings = checker.run_check(p, target)
    fix_docx(p, str(out), target, findings=findings)
    doc = Document(str(out))
    assert doc.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert doc.paragraphs[0].runs[0].bold is True
    assert doc.paragraphs[1].alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert doc.paragraphs[2].alignment == WD_ALIGN_PARAGRAPH.CENTER
    # 修完复检必须 pass
    _, after = checker.run_check(str(out), target)
    f = [x for x in after if x.dimension == "title_page"]
    assert f and f[0].severity == "pass", f[0].severity if f else "missing"


def test_title_page_not_changed_without_check_verdict(tmp_path):
    """没有 title_page 结论时，修正器绝不能碰标题页（旧版曾无条件改）。"""
    p = _doc(tmp_path / "tp.docx")
    out = tmp_path / "out.docx"
    # 只给一个与标题页无关的目标 → checker 不出 title_page 结论
    target = TargetProfileNoTp = target_from_dict({"font_size_pt": 11.0})
    fix_docx(p, str(out), target)
    doc = Document(str(out))
    assert doc.paragraphs[0].alignment is None, "标题页不该被改动"
    assert doc.paragraphs[0].runs[0].bold in (None, False)


def test_title_page_spec_fallback_still_works(tmp_path):
    """客户没填时按内置规范兜底（APA 要求标题居中加粗），保持旧能力不丢。"""
    p = _doc(tmp_path / "tp.docx")
    out = tmp_path / "out.docx"
    fix_docx(p, str(out), target_from_spec("APA"))
    doc = Document(str(out))
    assert doc.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert doc.paragraphs[0].runs[0].bold is True


def test_title_page_template_values_are_read(tmp_path):
    """学校模板的标题页格式必须能读出来（居中/加粗/作者行居中）。"""
    tpl = _doc(tmp_path / "tpl.docx", title_center=True, title_bold=True,
               author_center=True)
    tgt = target_from_template_docx(read_docx(tpl))
    assert tgt.title_page_title_centered is True
    assert tgt.title_page_title_bold is True
    assert tgt.title_page_author_centered is True
    assert tgt.resolution.get("title_page_title_centered") == "school"


# ---------------------------------------------------------------------------
# ③ 块引用（检查与修正同源）
# ---------------------------------------------------------------------------

def test_block_quote_reported_previewed_and_fixed(tmp_path):
    """长引文：报告 → 预览 → 真缩进 0.5″ → 复检不再报。"""
    p = _doc_with_quotation(tmp_path / "q.docx")
    out = tmp_path / "q_out.docx"
    target = target_from_spec("MLA")
    prof, findings = checker.run_check(p, target)
    f = [x for x in findings if x.dimension == "block_quote"]
    assert f, "长引文应被识别"
    assert f[0].fix is not None, "块引用必须绑定修正动作（旧版会改但不在检查结论里）"
    lines = _compute_changes(prof, target, lang="en", findings=findings)
    assert any("Block quotation" in ln for ln in lines), lines

    fix_docx(p, str(out), target, findings=findings)
    q = [x for x in Document(str(out)).paragraphs if "long quotation" in x.text][0]
    assert q.paragraph_format.left_indent.inches == pytest.approx(0.5)

    _, after = checker.run_check(str(out), target)
    assert not [x for x in after if x.dimension == "block_quote"], "修完不该再报"


def test_block_quote_no_false_preview_without_candidates(tmp_path):
    """没有长引文时，预览里不得凭空出现块引用行。"""
    p = _doc(tmp_path / "plain.docx")
    target = target_from_spec("MLA")
    prof, findings = checker.run_check(p, target)
    lines = _compute_changes(prof, target, lang="en", findings=findings)
    assert not any("Block quotation" in ln for ln in lines), lines
