# -*- coding: utf-8 -*-
"""标题页 / 摘要 / 图表 caption 三处的**段前段后**全链路回归（v2.3.76）。

背景（老板 2026-10-20）：「标题页 / 摘要 / 图表 caption 目前只做了『行距』，没有『段前段后』。
要不要一起补？」—— 答：要。

补之前这三处只有「行距」一个 spacing 维度，段前段后既读不到、查不出、也改不了：
学校模板写着「摘要上方 12 磅、图题下方 6 磅」，客户填进问卷也不生效。
本文件按「模板能读 → 检查报得出 → 预览提得前 → 修正改得对 → 二次幂等 → 没要求绝不动」
六环逐环钉死，并把「标题页段前段后判成 pass」这个顺序坑（_sev 在比对前就定了）钉住。
"""
import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt

from src.engine.checker import run_check
from src.engine.docx_reader import read_docx
from src.engine.fixer import compute_changes, fix_docx
from src.engine.questionnaire import TargetProfile, target_from_dict, target_from_template_docx


# ---------------------------------------------------------------------------
# 造数据工具
# ---------------------------------------------------------------------------

def _set_lines(para, before=None, after=None):
    """给段落写 w:beforeLines / w:afterLines（1/100 行）——中文学校模板常态。"""
    pPr = para._p.get_or_add_pPr()
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        sp = pPr.makeelement(qn("w:spacing"), {})
        pPr.append(sp)
    if before is not None:
        sp.set(qn("w:beforeLines"), str(int(round(before * 100))))
    if after is not None:
        sp.set(qn("w:afterLines"), str(int(round(after * 100))))


def _set_pt(para, before=None, after=None):
    if before is not None:
        para.paragraph_format.space_before = Pt(before)
    if after is not None:
        para.paragraph_format.space_after = Pt(after)


def _title_page_paper(path, *, before_lines=None, before_pt=0.0, after_pt=0.0,
                      title_size=16.0):
    """标题页型论文：居中加粗大标题 + 作者行 + Abstract（Heading 1）+ 摘要正文。"""
    doc = Document()
    n = doc.styles["Normal"]
    n.font.name = "Times New Roman"
    n.font.size = Pt(12)
    n.paragraph_format.line_spacing = 2.0

    title = doc.add_paragraph()
    run = title.add_run("The Impact of Remote Work on Team Performance")
    run.font.size = Pt(title_size)
    run.font.bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if before_lines is not None:
        _set_lines(title, before=before_lines)
    else:
        _set_pt(title, before=before_pt, after=after_pt)

    author = doc.add_paragraph("Jane Doe")
    author.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("Abstract", style="Heading 1")
    para = doc.add_paragraph("This study examines how remote work shapes team performance. " * 3)
    doc.add_paragraph("Body paragraph after the abstract, long enough to be sampled. " * 3)
    doc.save(str(path))
    return str(path)


def _abstract_paper(path, *, before_pt=0.0, after_pt=0.0, before_lines=None):
    """带摘要的论文（无标题页争议）：摘要正文段的段前段后可按磅或「行」写。"""
    doc = Document()
    n = doc.styles["Normal"]
    n.font.name = "Times New Roman"
    n.font.size = Pt(12)
    doc.add_paragraph("Introduction", style="Heading 1")
    doc.add_paragraph("Body text before the abstract, long enough to be sampled. " * 2)
    doc.add_paragraph("Abstract", style="Heading 1")
    abs_p = doc.add_paragraph("This study examines how remote work shapes team performance. " * 3)
    if before_lines is not None:
        _set_lines(abs_p, before=before_lines, after=before_lines)
    else:
        _set_pt(abs_p, before=before_pt, after=after_pt)
    doc.add_paragraph("Another abstract paragraph so the mode is stable. " * 3)
    doc.add_paragraph("Conclusion", style="Heading 1")
    doc.add_paragraph("Body text after the abstract, long enough to be sampled. " * 2)
    doc.save(str(path))
    return str(path)


def _caption_paper(path, *, before_pt=0.0, after_pt=0.0):
    """带 Caption 段（Word 内置 Caption 样式）的论文。"""
    doc = Document()
    n = doc.styles["Normal"]
    n.font.name = "Times New Roman"
    n.font.size = Pt(12)
    doc.add_paragraph("Results", style="Heading 1")
    doc.add_paragraph("Body text introducing the table, long enough for sampling. " * 2)
    cap = doc.add_paragraph("Table 1 Results of the regression", style="Caption")
    _set_pt(cap, before=before_pt, after=after_pt)
    t = doc.add_table(rows=2, cols=2)
    t.cell(0, 0).text = "Item"
    t.cell(0, 1).text = "Value"
    doc.add_paragraph("Body text after the table, long enough for sampling. " * 2)
    doc.save(str(path))
    return str(path)


def _mk_target(**extra):
    data = {"font_size_pt": 12}
    data.update(extra)
    return target_from_dict(data, "user")


def _para_with(doc, needle):
    return [p for p in doc.paragraphs if needle in (p.text or "")][0]


# ---------------------------------------------------------------------------
# 一、模型 / 键表
# ---------------------------------------------------------------------------

def test_special_spacing_fields_default_to_none():
    t = TargetProfile()
    for k in ("title_page_space_before_pt", "title_page_space_after_pt",
              "abstract_space_before_pt", "abstract_space_after_pt",
              "caption_space_before_pt", "caption_space_after_pt"):
        assert getattr(t, k) is None


def test_keys_are_registered_in_source_tables():
    """键表不登记 → target_from_dict / overlay 会静默丢掉客户填的值。"""
    for k in ("title_page_space_before_pt", "title_page_space_after_pt"):
        assert k in TargetProfile.TITLE_PAGE_KEYS
    for k in ("abstract_space_before_pt", "abstract_space_after_pt"):
        assert k in TargetProfile.ABSTRACT_KEYS
    for k in ("caption_space_before_pt", "caption_space_after_pt"):
        assert k in TargetProfile.CAPTION_KEYS


def test_target_from_dict_carries_six_fields():
    t = _mk_target(title_page_space_before_pt=0, title_page_space_after_pt=12,
                   abstract_space_before_pt=12, abstract_space_after_pt=6,
                   caption_space_before_pt=6, caption_space_after_pt=6)
    assert t.title_page_space_before_pt == 0.0
    assert t.title_page_space_after_pt == 12.0
    assert t.abstract_space_before_pt == 12.0
    assert t.abstract_space_after_pt == 6.0
    assert t.caption_space_before_pt == 6.0
    assert t.caption_space_after_pt == 6.0
    for k in ("abstract_space_before_pt", "caption_space_after_pt"):
        assert t.resolution.get(k) == "user"


def test_zero_is_not_treated_as_unset():
    """0 磅 = 明确要求「无间距」，必须与「不规定（None）」分开。"""
    t = _mk_target(abstract_space_before_pt=0)
    assert t.abstract_space_before_pt == 0.0
    assert t.abstract_space_after_pt is None


# ---------------------------------------------------------------------------
# 二、读取层
# ---------------------------------------------------------------------------

def test_reads_title_page_title_line_units(tmp_path):
    """标题页大标题「段前 1 行」@16pt → 1.2 × 16 = 19.2 磅。"""
    p = _title_page_paper(tmp_path / "tp.docx", before_lines=1.0)
    prof = read_docx(p)
    assert prof.title_page_detected is True
    assert prof.tp_title_space_before_pt == pytest.approx(19.2)


def test_reads_title_page_title_points(tmp_path):
    p = _title_page_paper(tmp_path / "tp2.docx", before_pt=12.0, after_pt=6.0)
    prof = read_docx(p)
    assert prof.tp_title_space_before_pt == pytest.approx(12.0)
    assert prof.tp_title_space_after_pt == pytest.approx(6.0)


def test_reads_abstract_body_spacing(tmp_path):
    p = _abstract_paper(tmp_path / "abs.docx", before_pt=12.0, after_pt=6.0)
    prof = read_docx(p)
    assert prof.abstract_space_before_pt == pytest.approx(12.0)
    assert prof.abstract_space_after_pt == pytest.approx(6.0)


def test_reads_abstract_line_units(tmp_path):
    """摘要正文「段前 0.5 行」@12pt → 0.6 × 12 = 7.2 磅。"""
    p = _abstract_paper(tmp_path / "abs2.docx", before_lines=0.5)
    prof = read_docx(p)
    assert prof.abstract_space_before_pt == pytest.approx(7.2)


def test_reads_caption_spacing(tmp_path):
    p = _caption_paper(tmp_path / "cap.docx", before_pt=6.0, after_pt=6.0)
    prof = read_docx(p)
    assert prof.caption_count == 1
    assert prof.caption_space_before_pt == pytest.approx(6.0)
    assert prof.caption_space_after_pt == pytest.approx(6.0)


def test_no_abstract_reads_none(tmp_path):
    """文档里没有摘要 → 画像字段留 None（不是 0），检查据此出 info 而非 fail。"""
    p = _caption_paper(tmp_path / "noabs.docx")
    prof = read_docx(p)
    assert prof.abstract_space_before_pt is None
    assert prof.abstract_space_after_pt is None


# ---------------------------------------------------------------------------
# 三、检查层：报得出 / 达成时判 pass / 没规定就不报
# ---------------------------------------------------------------------------

def test_checker_reports_title_page_spacing_mismatch(tmp_path):
    """模板要求段前 0，论文写着「段前 1 行」(19.2pt) → 必须非 pass 且文案点出段前。"""
    p = _title_page_paper(tmp_path / "tp.docx", before_lines=1.0)
    t = _mk_target(title_page_space_before_pt=0, title_page_space_after_pt=0)
    _, findings = run_check(p, t)
    f = [x for x in findings if x.dimension == "title_page"][0]
    assert f.severity != "pass"
    assert "title space before" in f.actual
    assert f.fix is not None and f.fix.dimension == "title_page"


def test_checker_passes_title_page_spacing_when_matching(tmp_path):
    p = _title_page_paper(tmp_path / "tp.docx", before_pt=0.0, after_pt=0.0)
    t = _mk_target(title_page_space_before_pt=0, title_page_space_after_pt=0)
    _, findings = run_check(p, t)
    f = [x for x in findings if x.dimension == "title_page"][0]
    assert f.severity == "pass"
    assert "title space before 0 pt" in f.actual


def test_checker_reports_abstract_spacing(tmp_path):
    p = _abstract_paper(tmp_path / "abs.docx", before_pt=0.0, after_pt=0.0)
    t = _mk_target(abstract_space_before_pt=12, abstract_space_after_pt=6)
    _, findings = run_check(p, t)
    dims = {x.dimension: x for x in findings}
    assert dims["abstract_space_before"].severity != "pass"
    assert dims["abstract_space_after"].severity != "pass"
    assert dims["abstract_space_before"].fix is not None


def test_checker_passes_abstract_spacing_when_matching(tmp_path):
    p = _abstract_paper(tmp_path / "abs.docx", before_pt=12.0, after_pt=6.0)
    t = _mk_target(abstract_space_before_pt=12, abstract_space_after_pt=6)
    _, findings = run_check(p, t)
    dims = {x.dimension: x for x in findings}
    assert dims["abstract_space_before"].severity == "pass"
    assert dims["abstract_space_after"].severity == "pass"


def test_checker_reports_caption_spacing(tmp_path):
    p = _caption_paper(tmp_path / "cap.docx", before_pt=0.0, after_pt=0.0)
    t = _mk_target(caption_space_before_pt=6, caption_space_after_pt=6)
    _, findings = run_check(p, t)
    dims = {x.dimension: x for x in findings}
    assert dims["caption_space_before"].severity != "pass"
    assert dims["caption_space_after"].severity != "pass"
    assert dims["caption_space_before"].fix is not None


def test_caption_spacing_is_info_when_no_captions(tmp_path):
    """没有图表标题 → 只能说「没检测到 caption」，不能判客户错。"""
    p = _abstract_paper(tmp_path / "nocap.docx")
    t = _mk_target(caption_space_before_pt=6)
    _, findings = run_check(p, t)
    f = [x for x in findings if x.dimension == "caption_space_before"][0]
    assert f.severity == "info"
    assert "No captions detected" in f.actual


def test_checker_stays_silent_when_target_unspecified(tmp_path):
    """目标一个字没说 → 这三处**一条 finding 都不许有**（不做无根据的猜测）。"""
    p = _caption_paper(tmp_path / "cap.docx", before_pt=3.0, after_pt=3.0)
    t = _mk_target()
    _, findings = run_check(p, t)
    dims = {x.dimension for x in findings}
    for d in ("abstract_space_before", "abstract_space_after",
              "caption_space_before", "caption_space_after"):
        assert d not in dims


# ---------------------------------------------------------------------------
# 四、修正层：改得对 / 清「行」单位 / 幂等 / 没要求绝不动
# ---------------------------------------------------------------------------

def test_fix_applies_abstract_spacing(tmp_path):
    src = _abstract_paper(tmp_path / "abs.docx", before_pt=0.0, after_pt=0.0)
    out = str(tmp_path / "abs_fixed.docx")
    t = _mk_target(abstract_space_before_pt=12, abstract_space_after_pt=6)
    info = fix_docx(src, out, t)
    assert info["applied"].get("abstract_space_before_pt") == 12.0
    para = _para_with(Document(out), "This study examines")
    assert para.paragraph_format.space_before == Pt(12)
    assert para.paragraph_format.space_after == Pt(6)


def test_fix_applies_caption_spacing(tmp_path):
    src = _caption_paper(tmp_path / "cap.docx", before_pt=0.0, after_pt=0.0)
    out = str(tmp_path / "cap_fixed.docx")
    t = _mk_target(caption_space_before_pt=6, caption_space_after_pt=12)
    info = fix_docx(src, out, t)
    assert info["applied"].get("caption_space_after_pt") == 12.0
    cap = _para_with(Document(out), "Table 1 Results")
    assert cap.paragraph_format.space_before == Pt(6)
    assert cap.paragraph_format.space_after == Pt(12)


def test_fix_applies_title_page_spacing(tmp_path):
    src = _title_page_paper(tmp_path / "tp.docx", before_lines=1.0)
    out = str(tmp_path / "tp_fixed.docx")
    t = _mk_target(title_page_title_centered=True, title_page_title_bold=True,
                   title_page_space_before_pt=0, title_page_space_after_pt=0)
    fix_docx(src, out, t)
    title = _para_with(Document(out), "The Impact of Remote Work")
    assert float(title.paragraph_format.space_before.pt) == 0.0
    assert float(title.paragraph_format.space_after.pt) == 0.0


def test_fix_clears_line_units_so_points_actually_win(tmp_path):
    """Word 里 w:beforeLines 优先于 w:before —— 不删掉它，写了磅值也等于没改。"""
    src = _abstract_paper(tmp_path / "abs.docx", before_lines=1.0)
    out = str(tmp_path / "abs_fixed.docx")
    t = _mk_target(abstract_space_before_pt=12)
    fix_docx(src, out, t)
    para = _para_with(Document(out), "This study examines")
    pPr = para._p.pPr
    sp = pPr.find(qn("w:spacing")) if pPr is not None else None
    assert sp is None or sp.get(qn("w:beforeLines")) is None
    assert read_docx(out).abstract_space_before_pt == pytest.approx(12.0)


def test_fix_is_idempotent(tmp_path):
    src = _abstract_paper(tmp_path / "abs.docx", before_pt=0.0, after_pt=0.0)
    out1 = str(tmp_path / "abs_fixed1.docx")
    out2 = str(tmp_path / "abs_fixed2.docx")
    t = _mk_target(abstract_space_before_pt=12, abstract_space_after_pt=6)
    fix_docx(src, out1, t)
    info2 = fix_docx(out1, out2, t)
    assert info2["applied"].get("abstract_space_before_pt") is None
    assert read_docx(out2).abstract_space_before_pt == pytest.approx(12.0)


def test_fix_does_not_touch_specials_when_not_requested(tmp_path):
    """目标没说段前段后 → 这三处一个字节都不许动（模板里带着 3 磅也得留着）。"""
    src = _caption_paper(tmp_path / "cap.docx", before_pt=3.0, after_pt=3.0)
    out = str(tmp_path / "cap_fixed.docx")
    t = _mk_target(caption_font_size_pt=10)
    fix_docx(src, out, t)
    cap = _para_with(Document(out), "Table 1 Results")
    assert cap.paragraph_format.space_before == Pt(3)
    assert cap.paragraph_format.space_after == Pt(3)


def test_fix_does_not_touch_abstract_when_only_caption_requested(tmp_path):
    """维度各自门控：只要求 caption 时，摘要段前段后必须原样保留。"""
    src = _abstract_paper(tmp_path / "abs.docx", before_pt=9.0, after_pt=9.0)
    out = str(tmp_path / "abs_fixed.docx")
    t = _mk_target(caption_space_before_pt=6)
    fix_docx(src, out, t)
    para = _para_with(Document(out), "This study examines")
    assert para.paragraph_format.space_before == Pt(9)


# ---------------------------------------------------------------------------
# 五、预览层：改之前必须看得见
# ---------------------------------------------------------------------------

def test_preview_lists_abstract_spacing(tmp_path):
    p = _abstract_paper(tmp_path / "abs.docx", before_pt=0.0, after_pt=0.0)
    t = _mk_target(abstract_space_before_pt=12)
    lines = compute_changes(p, t, "en")
    assert any("Abstract space before" in ln and "12" in ln for ln in lines)


def test_preview_lists_caption_spacing(tmp_path):
    p = _caption_paper(tmp_path / "cap.docx", before_pt=0.0, after_pt=0.0)
    t = _mk_target(caption_space_after_pt=6)
    lines = compute_changes(p, t, "en")
    assert any("Caption space after" in ln and "6" in ln for ln in lines)


def test_preview_lists_title_page_spacing(tmp_path):
    p = _title_page_paper(tmp_path / "tp.docx", before_pt=0.0, after_pt=0.0)
    t = _mk_target(title_page_space_before_pt=12)
    lines = compute_changes(p, t, "en")
    assert any("Title page" in ln and "title space before" in ln and "12" in ln
               for ln in lines)


def test_preview_empty_on_already_compliant_document(tmp_path):
    """已经合规 → 预览为空（客户不会被拉去点一次「什么都没改」的修正）。"""
    p = _abstract_paper(tmp_path / "abs.docx", before_pt=12.0, after_pt=6.0)
    t = _mk_target(abstract_space_before_pt=12, abstract_space_after_pt=6)
    lines = compute_changes(p, t, "en")
    assert not any("Abstract space" in ln for ln in lines)


# ---------------------------------------------------------------------------
# 六、模板通道：模板能读 → 问卷/目标带得走
# ---------------------------------------------------------------------------

def test_template_carries_title_page_and_abstract_spacing(tmp_path):
    p = _title_page_paper(tmp_path / "tp.docx", before_pt=12.0, after_pt=6.0)
    t = target_from_template_docx(read_docx(p))
    assert t.title_page_space_before_pt == pytest.approx(12.0)
    assert t.title_page_space_after_pt == pytest.approx(6.0)
    assert t.resolution.get("title_page_space_before_pt") == "school"


def test_template_carries_caption_spacing(tmp_path):
    p = _caption_paper(tmp_path / "cap.docx", before_pt=6.0, after_pt=6.0)
    t = target_from_template_docx(read_docx(p))
    assert t.caption_space_before_pt == pytest.approx(6.0)
    assert t.caption_space_after_pt == pytest.approx(6.0)
    assert t.resolution.get("caption_space_before_pt") == "school"


def test_template_carries_abstract_spacing(tmp_path):
    p = _title_page_paper(tmp_path / "tp.docx")
    # 摘要正文段写 12 磅段前
    doc = Document(p)
    _set_pt(_para_with(doc, "This study examines"), before=12.0, after=6.0)
    p2 = str(tmp_path / "tp_abs.docx")
    doc.save(p2)
    t = target_from_template_docx(read_docx(p2))
    assert t.abstract_space_before_pt == pytest.approx(12.0)
    assert t.abstract_space_after_pt == pytest.approx(6.0)
