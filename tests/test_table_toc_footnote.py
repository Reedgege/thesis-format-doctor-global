# -*- coding: utf-8 -*-
"""表格内段落 / 目录（TOC）/ 脚注尾注提示 —— v2.3.77 全链路回归。

老板 2026-10-20 定的范围（脚注明确**不改**）：
  ① 表格内段落段前段后 → 可查可修（表格与正文是两套口径）；
  ② 目录 TOC 第一档 → 目录**标题段**可查可修；**条目只核对不修改**
     （条目是域段落，Word「更新目录」会用 TOC 1..9 样式整体重排，改直接格式等于白改）；
  ③ 脚注 / 尾注 → **不修改**，只给客户三层提示（已知限制 / 一条 info / 批注 Scope 一句），
     并且正文改正文字号时不许把脚注引用标记一起改。

本文件逐环节钉死：模板能读 → 检查报得出 → 预览提得前 → 修正改得对 → 二次幂等 →
没有要求时一个字都不动。
"""
import pytest
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from src.engine.checker import run_check
from src.engine.docx_reader import (count_note_entries, detect_toc, read_docx)
from src.engine.fixer import compute_changes, fix_docx
from src.engine.questionnaire import target_from_dict, target_from_template_docx
from src.engine.report import KNOWN_LIMITATIONS


# ---------------------------------------------------------------------------
# 造数据工具
# ---------------------------------------------------------------------------

def _set_pt(para, before=None, after=None):
    if before is not None:
        para.paragraph_format.space_before = Pt(before)
    if after is not None:
        para.paragraph_format.space_after = Pt(after)


def _set_lines(para, before=None, after=None):
    """写 w:beforeLines / w:afterLines（1/100 行），中文模板常见写法。"""
    pPr = para._p.get_or_add_pPr()
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        sp = pPr.makeelement(qn("w:spacing"), {})
        pPr.append(sp)
    if before is not None:
        sp.set(qn("w:beforeLines"), str(int(round(before * 100))))
    if after is not None:
        sp.set(qn("w:afterLines"), str(int(round(after * 100))))


def _read_spacing(para):
    """(段前, 段后) 磅值；读不到返回 (None, None)。"""
    pPr = para._p.pPr
    if pPr is None:
        return None, None
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        return None, None

    def _one(pt_attr, ln_attr):
        raw_ln = sp.get(qn(ln_attr))
        if raw_ln is not None:
            return ("lines", raw_ln)
        raw = sp.get(qn(pt_attr))
        return ("pt", raw) if raw is not None else (None, None)

    return _one("w:before", "w:beforeLines"), _one("w:after", "w:afterLines")


def _table_paper(path, *, cell_before=6.0, cell_after=6.0, body_before=12.0):
    """一篇带三线表的论文：表格内 6 磅、正文 12 磅（两套口径，必须分开判）。"""
    doc = Document()
    n = doc.styles["Normal"]
    n.font.name = "Times New Roman"
    n.font.size = Pt(12)
    doc.add_paragraph("Introduction", style="Heading 1")
    p1 = doc.add_paragraph("Body text introducing the table, long enough to be sampled. " * 2)
    _set_pt(p1, before=body_before, after=0)
    t = doc.add_table(rows=2, cols=2)
    for r in t.rows:
        for c in r.cells:
            cp = c.paragraphs[0]
            cp.text = "Cell text"
            _set_pt(cp, before=cell_before, after=cell_after)
    p2 = doc.add_paragraph("Body text after the table, long enough to be sampled. " * 2)
    _set_pt(p2, before=body_before, after=0)
    doc.save(str(path))
    return str(path)


def _toc_entry(doc, text, anchor):
    """造一条「像 Word 生成的」目录条目：带指向 _Toc 书签的超链接（真实信号）。"""
    p = doc.add_paragraph()
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("w:anchor"), anchor)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    hl.append(r)
    p._p.append(hl)
    return p


def _toc_paper(path, *, head_before=12.0, head_after=12.0,
               entry_before=6.0, entry_after=6.0):
    """带目录的论文：目录标题段 + 3 条目录条目。"""
    doc = Document()
    n = doc.styles["Normal"]
    n.font.name = "Times New Roman"
    n.font.size = Pt(12)
    head = doc.add_paragraph("Contents", style="TOC Heading")
    _set_pt(head, before=head_before, after=head_after)
    for i, t in enumerate(("1 Introduction", "2 Methods", "3 Results"), start=1):
        p = _toc_entry(doc, t, "_Toc%09d" % i)
        _set_pt(p, before=entry_before, after=entry_after)
    doc.add_paragraph("Introduction", style="Heading 1")
    doc.add_paragraph("Body paragraph, long enough to be sampled as body text. " * 3)
    doc.save(str(path))
    return str(path)


_FOOT_CT = ("application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.footnotes+xml")
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _add_footnotes(doc, notes):
    """给文档装上真实的 footnotes.xml 部件（含 3 条 id<=0 的特殊条目，用来验证跳过）。"""
    specials = (
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:t>sep</w:t></w:r></w:p></w:footnote>'
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p/></w:footnote>'
        '<w:footnote w:type="continuationNotice" w:id="1"><w:p/></w:footnote>'
    )
    body = "".join(
        '<w:footnote w:id="%d"><w:p><w:r><w:t>%s</w:t></w:r></w:p></w:footnote>'
        % (i, text) for i, text in enumerate(notes, start=2)
    )
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:footnotes xmlns:w="%s">%s%s</w:footnotes>' % (_W_NS, specials, body))
    part = Part(PackURI("/word/footnotes.xml"), _FOOT_CT, xml.encode("utf-8"),
                doc.part.package)
    doc.part.relate_to(part, RT.FOOTNOTES)
    # 不需要手工注册 [Content_Types].xml 的 Override：python-docx 保存时会按
    # package.iter_parts() 重建整个 Content_Types（pkgwriter._ContentTypesItem.from_parts），
    # 裸 Part 的 content_type 会被自动写成 Override。1.2.0 的 Package 也没有
    # content_types 属性（那是旧版 API），强行调用会 AttributeError。


def _footnote_paper(path, notes=("First note.", "Second note.", "Third note.")):
    """正文里带脚注引用标记 + footnotes.xml 的论文。"""
    doc = Document()
    n = doc.styles["Normal"]
    n.font.name = "Times New Roman"
    n.font.size = Pt(12)
    doc.add_paragraph("Introduction", style="Heading 1")
    p = doc.add_paragraph("Body text with a footnote reference inside it. " * 2)
    ref = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rstyle = OxmlElement("w:rStyle")
    rstyle.set(qn("w:val"), "FootnoteReference")
    rpr.append(rstyle)
    ref.append(rpr)
    mark = OxmlElement("w:footnoteReference")
    mark.set(qn("w:id"), "2")
    ref.append(mark)
    p._p.append(ref)
    _add_footnotes(doc, list(notes))
    doc.save(str(path))
    return str(path)


def _mk_target(**extra):
    data = {"font_size_pt": 12}
    data.update(extra)
    return target_from_dict(data, "user")


def _cell_paras(path):
    return [c.paragraphs[0] for t in Document(path).tables for r in t.rows for c in r.cells]


# ---------------------------------------------------------------------------
# 一、表格内段落：读取 / 不污染正文 / 检查 / 修正 / 预览 / 模板
# ---------------------------------------------------------------------------

def test_reads_table_cell_spacing(tmp_path):
    p = _table_paper(tmp_path / "t.docx", cell_before=6.0, cell_after=3.0)
    prof = read_docx(p)
    assert prof.table_space_before_pt == pytest.approx(6.0)
    assert prof.table_space_after_pt == pytest.approx(3.0)


def test_table_spacing_does_not_pollute_body_mode(tmp_path):
    """表格 6 磅、正文 12 磅 —— 两个众数必须各归各的（混起来两边都判错）。"""
    p = _table_paper(tmp_path / "t.docx", cell_before=6.0, body_before=12.0)
    prof = read_docx(p)
    assert prof.space_before_pt == pytest.approx(12.0)
    assert prof.table_space_before_pt == pytest.approx(6.0)


def test_checker_reports_table_spacing(tmp_path):
    p = _table_paper(tmp_path / "t.docx", cell_before=6.0)
    t = _mk_target(table_space_before_pt=0)
    _, findings = run_check(p, t)
    f = [x for x in findings if x.dimension == "table_space_before"][0]
    assert f.severity != "pass"
    assert f.fix is not None and f.fix.dimension == "table_space_before"


def test_checker_gives_info_when_no_table(tmp_path):
    p = tmp_path / "n.docx"
    doc = Document()
    doc.add_paragraph("Body paragraph without any table, long enough. " * 3)
    doc.save(str(p))
    t = _mk_target(table_space_before_pt=0)
    _, findings = run_check(str(p), t)
    f = [x for x in findings if x.dimension == "table_space_before"][0]
    assert f.severity == "info"
    assert "No table text detected" in f.actual


def test_fix_applies_table_spacing_without_touching_body(tmp_path):
    src = _table_paper(tmp_path / "t.docx", cell_before=6.0, body_before=12.0)
    out = str(tmp_path / "t_fixed.docx")
    t = _mk_target(table_space_before_pt=0, table_space_after_pt=0)
    info = fix_docx(src, out, t)
    assert info["applied"].get("table_space_before_pt") == 0.0
    for cp in _cell_paras(out):
        assert cp.paragraph_format.space_before == Pt(0)
        assert cp.paragraph_format.space_after == Pt(0)
    body = [p for p in Document(out).paragraphs if "Body text introducing" in p.text][0]
    assert body.paragraph_format.space_before == Pt(12), "正文段前不该被表格口径改"


def test_fix_table_spacing_clears_line_units(tmp_path):
    """表格里写「段前 0.5 行」也要能改成磅值（必须清掉 w:*Lines，否则改了等于没改）。"""
    src = tmp_path / "t.docx"
    doc = Document()
    doc.add_paragraph("Body text, long enough to be sampled as body text. " * 3)
    tb = doc.add_table(rows=1, cols=1)
    cp = tb.cell(0, 0).paragraphs[0]
    cp.text = "Cell"
    _set_lines(cp, before=0.5, after=0.5)
    doc.save(str(src))
    out = str(tmp_path / "t_fixed.docx")
    fix_docx(str(src), out, _mk_target(table_space_before_pt=0))
    cp2 = _cell_paras(out)[0]
    kind, _raw = _read_spacing(cp2)[0]
    assert kind == "pt", "w:beforeLines 必须被清掉，否则 Word 仍按「行」渲染"


def test_table_fix_is_idempotent(tmp_path):
    src = _table_paper(tmp_path / "t.docx", cell_before=6.0)
    out1, out2 = str(tmp_path / "a.docx"), str(tmp_path / "b.docx")
    t = _mk_target(table_space_before_pt=0)
    fix_docx(src, out1, t)
    info2 = fix_docx(out1, out2, t)
    assert info2["applied"].get("table_space_before_pt") is None


def test_table_untouched_when_not_requested(tmp_path):
    src = _table_paper(tmp_path / "t.docx", cell_before=6.0)
    out = str(tmp_path / "t_fixed.docx")
    fix_docx(src, out, _mk_target(font_size_pt=12))
    for cp in _cell_paras(out):
        assert cp.paragraph_format.space_before == Pt(6)


def test_preview_lists_table_spacing(tmp_path):
    p = _table_paper(tmp_path / "t.docx", cell_before=6.0)
    lines = compute_changes(p, _mk_target(table_space_before_pt=0), "en")
    assert any("Table cell space before" in ln and "6" in ln for ln in lines)


def test_template_carries_table_spacing(tmp_path):
    p = _table_paper(tmp_path / "t.docx", cell_before=6.0, cell_after=0.0)
    t = target_from_template_docx(read_docx(p))
    assert t.table_space_before_pt == pytest.approx(6.0)
    assert t.table_space_after_pt == pytest.approx(0.0)
    assert t.resolution.get("table_space_before_pt") == "school"


# ---------------------------------------------------------------------------
# 二、目录：识别 / 标题可修 / 条目只提示 / 不被正文流程改
# ---------------------------------------------------------------------------

def test_detects_toc_heading_and_entries(tmp_path):
    p = _toc_paper(tmp_path / "toc.docx")
    info = detect_toc(Document(p))
    assert info is not None
    assert info["heading"] is not None
    assert len(info["entries"]) == 3
    assert Document(p).paragraphs[info["heading"]].text.strip() == "Contents"


def test_reads_toc_spacing(tmp_path):
    p = _toc_paper(tmp_path / "toc.docx", head_before=12.0, head_after=6.0,
                   entry_before=3.0, entry_after=0.0)
    prof = read_docx(p)
    assert prof.toc_detected is True
    assert prof.toc_entry_count == 3
    assert prof.toc_heading_space_before_pt == pytest.approx(12.0)
    assert prof.toc_heading_space_after_pt == pytest.approx(6.0)
    assert prof.toc_entry_space_before_pt == pytest.approx(3.0)


def test_toc_heading_is_not_sampled_as_body(tmp_path):
    """目录标题段与条目不能混进正文众数（否则短文档里正文口径会被带偏）。"""
    p = _toc_paper(tmp_path / "toc.docx", entry_before=99.0, head_before=99.0)
    prof = read_docx(p)
    assert prof.space_before_pt == pytest.approx(0.0)


def test_checker_reports_toc_heading_spacing(tmp_path):
    p = _toc_paper(tmp_path / "toc.docx", head_before=12.0)
    t = _mk_target(toc_heading_space_before_pt=0)
    _, findings = run_check(p, t)
    f = [x for x in findings if x.dimension == "toc_heading_space_before"][0]
    assert f.severity != "pass"
    assert f.fix is not None


def test_toc_entry_spacing_is_info_only_never_fixed(tmp_path):
    """目标给了条目间距 → 只核对、不修改（fix 必须为 None）。"""
    p = _toc_paper(tmp_path / "toc.docx", entry_before=6.0)
    t = _mk_target(toc_entry_space_before_pt=0)
    _, findings = run_check(p, t)
    f = [x for x in findings if x.dimension == "toc_entry_spacing"][0]
    assert f.severity == "info"
    assert f.fix is None
    assert "TOC 1-9 styles" in f.actual


def test_fix_applies_toc_heading_spacing_only(tmp_path):
    src = _toc_paper(tmp_path / "toc.docx", head_before=12.0, entry_before=6.0)
    out = str(tmp_path / "toc_fixed.docx")
    t = _mk_target(toc_heading_space_before_pt=0)
    info = fix_docx(src, out, t)
    assert info["applied"].get("toc_heading_space_before_pt") == 0.0
    doc = Document(out)
    head = [p for p in doc.paragraphs if p.text.strip() == "Contents"][0]
    assert head.paragraph_format.space_before == Pt(0)
    # 条目一律不动
    info2 = detect_toc(doc)
    for i in info2["entries"]:
        assert doc.paragraphs[i].paragraph_format.space_before == Pt(6)


def test_fix_never_touches_toc_entries_even_if_requested(tmp_path):
    """客户填了条目间距也不改文档 —— 只提示。条目必须原样。"""
    src = _toc_paper(tmp_path / "toc.docx", entry_before=6.0)
    out = str(tmp_path / "toc_fixed.docx")
    t = _mk_target(toc_entry_space_before_pt=0, toc_entry_space_after_pt=0)
    info = fix_docx(src, out, t)
    assert "toc_entry_space_before_pt" not in info["applied"]
    doc = Document(out)
    for i in detect_toc(doc)["entries"]:
        assert doc.paragraphs[i].paragraph_format.space_before == Pt(6)


def test_toc_heading_excluded_from_body_spacing(tmp_path):
    """目录标题段的段前段后由 toc_heading 维度独占：正文段前改动不能刷它。"""
    src = _toc_paper(tmp_path / "toc.docx", head_before=12.0)
    out = str(tmp_path / "toc_fixed.docx")
    t = _mk_target(space_before_pt=36)      # 只改正文段前
    fix_docx(src, out, t)
    head = [p for p in Document(out).paragraphs if p.text.strip() == "Contents"][0]
    assert head.paragraph_format.space_before == Pt(12), "目录标题段不该吃正文段前"


def test_checker_gives_info_when_no_toc(tmp_path):
    p = _table_paper(tmp_path / "t.docx")
    t = _mk_target(toc_heading_space_before_pt=0)
    _, findings = run_check(p, t)
    f = [x for x in findings if x.dimension == "toc_heading_space_before"][0]
    assert f.severity == "info"
    assert "No table of contents detected" in f.actual


def test_preview_lists_toc_heading_spacing(tmp_path):
    p = _toc_paper(tmp_path / "toc.docx", head_before=12.0)
    lines = compute_changes(p, _mk_target(toc_heading_space_before_pt=0), "en")
    assert any("Contents heading space before" in ln and "12" in ln for ln in lines)


def test_template_carries_toc_heading_spacing(tmp_path):
    p = _toc_paper(tmp_path / "toc.docx", head_before=12.0, head_after=6.0)
    t = target_from_template_docx(read_docx(p))
    assert t.toc_heading_space_before_pt == pytest.approx(12.0)
    assert t.toc_heading_space_after_pt == pytest.approx(6.0)


# ---------------------------------------------------------------------------
# 三、脚注 / 尾注：只数不改 + 三层提示
# ---------------------------------------------------------------------------

def test_counts_footnotes_and_skips_special_entries(tmp_path):
    """3 条真脚注 + 3 条 id<=0 的特殊条目 → 只数 3。"""
    p = _footnote_paper(tmp_path / "f.docx")
    assert count_note_entries(Document(p), "w:footnote") == 3
    assert read_docx(p).footnote_count == 3


def test_no_notes_reads_zero(tmp_path):
    p = _table_paper(tmp_path / "t.docx")
    prof = read_docx(p)
    assert prof.footnote_count == 0 and prof.endnote_count == 0


def test_checker_hints_footnotes_are_not_checked(tmp_path):
    p = _footnote_paper(tmp_path / "f.docx")
    _, findings = run_check(p, _mk_target())
    f = [x for x in findings if x.dimension == "footnote"][0]
    assert f.severity == "info"
    assert f.fix is None, "脚注不修改：这条永远不许绑修正动作"
    assert "3 footnotes" in f.actual
    assert "left untouched" in f.actual


def test_no_footnote_hint_when_document_has_none(tmp_path):
    p = _table_paper(tmp_path / "t.docx")
    _, findings = run_check(p, _mk_target())
    assert not [x for x in findings if x.dimension == "footnote"]


def test_footnote_mark_run_is_not_resized(tmp_path):
    """正文改正文字号时，脚注引用标记（上标数字）必须保持原样。"""
    src = _footnote_paper(tmp_path / "f.docx")
    out = str(tmp_path / "f_fixed.docx")
    fix_docx(src, out, _mk_target(font_size_pt=11))
    doc = Document(out)
    para = [p for p in doc.paragraphs if "footnote reference" in p.text][0]
    marks = [r for r in para.runs
             if r._r.find(qn("w:footnoteReference")) is not None]
    assert marks, "测试前提：段落里必须有脚注引用标记"
    for r in marks:
        assert r.font.size is None, "引用标记的字号不许被正文流程改"


def test_known_limitations_declares_footnotes_and_toc():
    txt = " ".join(KNOWN_LIMITATIONS)
    assert "Footnotes and endnotes are NOT checked" in txt
    assert "TOC 1-9 styles" in txt
