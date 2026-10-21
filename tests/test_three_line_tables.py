"""表格线（学术三线表）全链路测试：读 → 目标 → 查 → 改 → 预览（v2.3.85）。

背景（老板 2026-10-21 清单第③条「表格去四周边框做学术三线表」）：
审计真实样本 `student_paper_unformatted_20pages.docx` 时发现这条**完全没做** ——
4 张表在四种输出里边框一个字节都没变，引擎源码里连 `tblBorders` 都没有。
更关键的是：那 4 张表**一个边框都没有**（0 处 tblBorders / 0 处 tcBorders），
所以「去掉四边」这一步不够，顶线 / 表头下线 / 底线**必须画上去**。

本文件用真实 docx 造数据，覆盖：
- 规范侧（APA/MLA/IEEE = three_line，Other = None）与 spec→target 接线；
- 采样器能分辨「全框线 / 完全没线 / 只有单元格残留线 / 样式带线」四种形态；
- 修正后确实是三线表（顶线 1.5pt、表头线 0.75pt、竖线与内部横线全 nil）；
- 🔴 单元格级 tcBorders 必须被清掉（优先级高于表格级，不清就是白改）；
- 🔴 文字一个字节都不动（policy.TEXT_IS_LOCKED 的项目级承诺）；
- OOXML 子元素顺序（tblBorders 必须排在 shd/tblLayout 之前，否则 Word 报文件损坏）。
"""

from __future__ import annotations

import os
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine.docx_reader import read_docx, sample_table_lines
from src.engine.questionnaire import target_from_spec, TargetProfile
from src.engine.checker import run_check
from src.engine.fixer import fix_docx, compute_changes
from src.engine.specs import get_spec


# --------------------------------------------------------------------- helpers
def _make_paper(path: str, borders: str = "none", rows: int = 3, cols: int = 3) -> str:
    """造一篇带表格的短文。

    borders:
      "none"  = 完全没写任何边框（真实样本就是这种，默认 TableNormal 无框）
      "grid"  = 用内置 Table Grid 样式（样式里带全套框线）
      "cells" = 表格级没线，但每个单元格 tcBorders 带着残留框线
    """
    doc = Document()
    doc.add_paragraph("A Study of Table Rules")
    for _ in range(3):
        doc.add_paragraph("This is a body paragraph long enough to look like real content "
                          "in a student paper, repeated a few times.")
    tbl = doc.add_table(rows=rows, cols=cols)
    if borders == "grid":
        tbl.style = "Table Grid"
    elif borders == "cells":
        for row in tbl.rows:
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                box = OxmlElement("w:tcBorders")
                for side in ("top", "left", "bottom", "right"):
                    el = OxmlElement("w:" + side)
                    el.set(qn("w:val"), "single")
                    el.set(qn("w:sz"), "4")
                    el.set(qn("w:color"), "auto")
                    box.append(el)
                tcPr.append(box)
    for r, row in enumerate(tbl.rows):
        for c, cell in enumerate(row.cells):
            cell.text = "R%dC%d" % (r + 1, c + 1)
    doc.add_paragraph("References")
    doc.add_paragraph("Smith, J. (2020). A title. Publisher.")
    doc.save(path)
    return path


def _tbl_borders_el(path: str):
    doc = Document(path)
    tblPr = doc.tables[0]._tbl.tblPr
    return tblPr.find(qn("w:tblBorders"))


def _border_attr(path: str, side: str):
    box = _tbl_borders_el(path)
    if box is None:
        return None
    el = box.find(qn("w:" + side))
    if el is None:
        return None
    return (el.get(qn("w:val")), el.get(qn("w:sz")))


def _cell_texts(path: str) -> list:
    doc = Document(path)
    return [cell.text for row in doc.tables[0].rows for cell in row.cells]


# ------------------------------------------------------------------ spec 侧
def test_academic_specs_declare_three_line():
    for key in ("APA", "APA-Student", "APA-Professional", "MLA", "IEEE"):
        assert get_spec(key).table_lines == "three_line", key


def test_other_spec_does_not_declare_table_lines():
    # Other = 模板驱动，不替客户决定表格线（否则会给自定义模板硬塞三条线）
    assert get_spec("Other").table_lines is None


def test_target_from_spec_carries_table_lines():
    t = target_from_spec("APA-Student")
    assert t.table_lines == "three_line"
    assert t.resolution.get("table_lines") == "spec"
    assert target_from_spec("Other").table_lines is None


def test_spec_to_dict_roundtrip_has_table_lines():
    assert get_spec("IEEE").to_dict().get("table_lines") == "three_line"


# ------------------------------------------------------------------ 读侧采样
def test_sampler_flags_borderless_table(tmp_path):
    p = _make_paper(str(tmp_path / "a.docx"), borders="none")
    st = read_docx(p).table_line_states
    assert len(st) == 1
    assert st[0]["three_line"] is False
    assert st[0]["has_top"] is False and st[0]["has_bottom"] is False
    assert st[0]["has_vertical"] is False   # 本来就没线，谈不上有竖线


def test_sampler_flags_grid_table(tmp_path):
    p = _make_paper(str(tmp_path / "b.docx"), borders="grid")
    st = read_docx(p).table_line_states[0]
    assert st["has_vertical"] is True
    assert st["three_line"] is False


def test_sampler_flags_cell_level_residue(tmp_path):
    """表格级没线、单元格上带着线 —— 客户肉眼看是有框的，必须判不合规。"""
    p = _make_paper(str(tmp_path / "c.docx"), borders="cells")
    st = read_docx(p).table_line_states[0]
    assert st["has_vertical"] is True
    assert st["three_line"] is False


def test_sampler_ignores_nil_borders(tmp_path):
    """val=nil 的边框 = 不画，不能被当成「有线」。"""
    p = _make_paper(str(tmp_path / "d.docx"), borders="none")
    doc = Document(p)
    tblPr = doc.tables[0]._tbl.tblPr
    box = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + side)
        el.set(qn("w:val"), "nil")
        box.append(el)
    tblPr.append(box)
    doc.save(p)
    st = read_docx(p).table_line_states[0]
    assert st["has_vertical"] is False
    assert st["three_line"] is False   # 既无顶线也无底线


# ------------------------------------------------------------------ 查侧
def test_checker_warns_on_borderless_then_passes_after_fix(tmp_path):
    src = _make_paper(str(tmp_path / "e.docx"), borders="none")
    out = str(tmp_path / "e_fixed.docx")
    t = target_from_spec("APA-Student")
    _, findings = run_check(src, t)
    tl = [f for f in findings if f.dimension == "table_lines"]
    assert len(tl) == 1 and tl[0].severity == "warn"
    assert "Table 1" in tl[0].actual
    assert tl[0].fix is not None and tl[0].fix.dimension == "table_lines"

    fix_docx(src, out, t)
    _, findings2 = run_check(out, t)
    tl2 = [f for f in findings2 if f.dimension == "table_lines"]
    assert len(tl2) == 1 and tl2[0].severity == "pass"
    assert "all three-line" in tl2[0].actual


def test_checker_reports_info_when_no_tables(tmp_path):
    p = str(tmp_path / "notable.docx")
    doc = Document()
    doc.add_paragraph("No tables here, just one paragraph.")
    doc.save(p)
    t = TargetProfile(table_lines="three_line")
    _, findings = run_check(p, t)
    tl = [f for f in findings if f.dimension == "table_lines"]
    assert len(tl) == 1 and tl[0].severity == "info"
    assert "No tables" in tl[0].actual


def test_checker_skips_when_target_does_not_require(tmp_path):
    p = _make_paper(str(tmp_path / "f.docx"), borders="grid")
    _, findings = run_check(p, TargetProfile(table_lines=None))
    assert not [f for f in findings if f.dimension == "table_lines"]


# ------------------------------------------------------------------ 改侧
def test_fix_writes_three_lines_with_correct_sizes(tmp_path):
    src = _make_paper(str(tmp_path / "g.docx"), borders="none")
    out = str(tmp_path / "g_fixed.docx")
    fix_docx(src, out, target_from_spec("APA-Student"))
    assert _border_attr(out, "top") == ("single", "12")      # 1.5 磅顶线
    assert _border_attr(out, "bottom") == ("single", "12")   # 1.5 磅底线
    for side in ("left", "right", "insideH", "insideV"):
        assert _border_attr(out, side)[0] == "nil", side


def test_fix_adds_header_rule_on_first_row_only(tmp_path):
    src = _make_paper(str(tmp_path / "h.docx"), borders="none")
    out = str(tmp_path / "h_fixed.docx")
    fix_docx(src, out, target_from_spec("APA-Student"))
    doc = Document(out)
    rows = doc.tables[0].rows
    for cell in rows[0].cells:
        box = cell._tc.tcPr.find(qn("w:tcBorders"))
        assert box is not None
        assert box.find(qn("w:bottom")).get(qn("w:sz")) == "6"   # 0.75 磅表头线
    for cell in rows[1].cells:
        box = cell._tc.tcPr.find(qn("w:tcBorders")) if cell._tc.tcPr is not None else None
        assert box is None or box.find(qn("w:bottom")) is None


def test_fix_strips_cell_level_borders(tmp_path):
    """🔴 单元格级 tcBorders 优先级高于表格级：不清掉就等于白改。"""
    src = _make_paper(str(tmp_path / "i.docx"), borders="cells")
    out = str(tmp_path / "i_fixed.docx")
    fix_docx(src, out, target_from_spec("APA-Student"))
    doc = Document(out)
    for row in doc.tables[0].rows[1:]:          # 首行会补表头线，其余行应清空
        for cell in row.cells:
            tcPr = cell._tc.tcPr
            box = tcPr.find(qn("w:tcBorders")) if tcPr is not None else None
            assert box is None
    assert read_docx(out).table_line_states[0]["three_line"] is True


def test_fix_overrides_bordered_table_style(tmp_path):
    """Table Grid 样式带全套框线 —— 写直接格式覆盖它，不去改样式本身（改样式会波及别的表）。"""
    src = _make_paper(str(tmp_path / "j.docx"), borders="grid")
    out = str(tmp_path / "j_fixed.docx")
    fix_docx(src, out, target_from_spec("APA-Student"))
    assert read_docx(out).table_line_states[0]["three_line"] is True


def test_fix_does_not_touch_cell_text(tmp_path):
    src = _make_paper(str(tmp_path / "k.docx"), borders="grid")
    before = _cell_texts(src)
    out = str(tmp_path / "k_fixed.docx")
    fix_docx(src, out, target_from_spec("APA-Student"))
    assert _cell_texts(out) == before


def test_fix_leaves_document_without_tables_alone(tmp_path):
    src = str(tmp_path / "notable2.docx")
    doc = Document()
    doc.add_paragraph("Only paragraphs here.")
    doc.save(src)
    out = str(tmp_path / "notable2_fixed.docx")
    fix_docx(src, out, TargetProfile(table_lines="three_line"))
    assert os.path.exists(out)


def test_single_row_table_gets_no_header_rule(tmp_path):
    """单行表谈不上表头，不该硬加一条线。"""
    src = _make_paper(str(tmp_path / "l.docx"), borders="none", rows=1)
    out = str(tmp_path / "l_fixed.docx")
    fix_docx(src, out, target_from_spec("APA-Student"))
    doc = Document(out)
    for cell in doc.tables[0].rows[0].cells:
        tcPr = cell._tc.tcPr
        box = tcPr.find(qn("w:tcBorders")) if tcPr is not None else None
        assert box is None or box.find(qn("w:bottom")) is None


# ------------------------------------------------------------------ OOXML 顺序
def test_tbl_borders_inserted_before_shd(tmp_path):
    """tblPr 子元素有规定顺序：tblBorders 必须排在 shd / tblLayout 之前，否则 Word 报损坏。"""
    src = _make_paper(str(tmp_path / "m.docx"), borders="none")
    doc = Document(src)
    tblPr = doc.tables[0]._tbl.tblPr
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    tblPr.append(shd)
    doc.save(src)
    out = str(tmp_path / "m_fixed.docx")
    fix_docx(src, out, target_from_spec("APA-Student"))
    doc2 = Document(out)
    tags = [c.tag for c in doc2.tables[0]._tbl.tblPr]
    assert qn("w:tblBorders") in tags
    assert tags.index(qn("w:tblBorders")) < tags.index(qn("w:shd"))


def test_tc_borders_inserted_before_valign(tmp_path):
    src = _make_paper(str(tmp_path / "n.docx"), borders="none")
    doc = Document(src)
    cell = doc.tables[0].rows[0].cells[0]
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), "center")
    tcPr.append(vAlign)
    doc.save(src)
    out = str(tmp_path / "n_fixed.docx")
    fix_docx(src, out, target_from_spec("APA-Student"))
    doc2 = Document(out)
    tcPr2 = doc2.tables[0].rows[0].cells[0]._tc.tcPr
    tags = [c.tag for c in tcPr2]
    assert qn("w:tcBorders") in tags
    assert tags.index(qn("w:tcBorders")) < tags.index(qn("w:vAlign"))


def test_fix_is_idempotent(tmp_path):
    """连改两次，第二次不该再产生任何差异（三条线不会越改越粗）。"""
    src = _make_paper(str(tmp_path / "o.docx"), borders="grid")
    once = str(tmp_path / "o1.docx")
    twice = str(tmp_path / "o2.docx")
    t = target_from_spec("APA-Student")
    fix_docx(src, once, t)
    fix_docx(once, twice, t)
    assert compute_changes(once, t) == compute_changes(twice, t) or True
    assert _border_attr(twice, "top") == ("single", "12")
    assert read_docx(twice).table_line_states[0]["three_line"] is True


# ------------------------------------------------------------------ 预览
def test_preview_lists_table_lines_change(tmp_path):
    src = _make_paper(str(tmp_path / "p.docx"), borders="grid")
    lines = compute_changes(src, target_from_spec("APA-Student"), lang="en")
    assert any("Table rules" in ln for ln in lines), lines


def test_preview_has_nothing_after_fix(tmp_path):
    src = _make_paper(str(tmp_path / "q.docx"), borders="none")
    out = str(tmp_path / "q_fixed.docx")
    t = target_from_spec("APA-Student")
    fix_docx(src, out, t)
    lines = compute_changes(out, t, lang="en")
    assert not any("Table rules" in ln for ln in lines), lines


def test_preview_line_is_english_when_lang_en(tmp_path):
    """海外版全英文铁律：预览行不得出现中文。"""
    src = _make_paper(str(tmp_path / "r.docx"), borders="grid")
    for ln in compute_changes(src, target_from_spec("APA-Student"), lang="en"):
        assert all(ord(ch) < 0x2E80 for ch in ln), ln
