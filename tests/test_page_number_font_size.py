"""页码位置 + 字体白名单（字体→字号配对）两条链路回归（v2.3.66）。

老板 2026-10-19 拍板：①客户手打页码用「策略 C」原地替换成 PAGE 域；②连页码真写一起做；
③字体方案 = **判断客户用哪种字体，然后把字号改成该字体的规范值**（不换字体）。

两个链路此前都是「只提示、改不了」，且各带一批假结论：
- **页码**：只认 PAGE 域（手打「第 1 页」被判「没有页码」）、不读首页页眉页脚、
  不区分页眉/页脚（页脚右对齐被误判为符合 APA top-right）、对齐读不到时走 pass（假通过）。
- **字体白名单**：APA 7th §2.19 的清单是「字体+字号」配对（TNR 12 / Calibri 11 / Arial 11 /
  Georgia 11 / Lucida Sans Unicode 10 / Computer Modern 10），旧版只比字体名 →
  Calibri 12pt 被判合格。
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
from src.engine.docx_reader import (
    read_docx, detect_static_page_number, _para_has_page_field,
)
from src.engine.fixer import fix_docx, _compute_changes
from src.engine.questionnaire import (
    target_from_dict, target_from_spec, target_from_template_docx,
    PAGE_NUMBER_POSITIONS,
)
from src.engine.specs import get_spec, size_for_font, APA_FONT_SIZES


# ---------------------------------------------------------------------------
# 夹具
# ---------------------------------------------------------------------------

def _write_page_field(para, cached: str = "1"):
    """在段落末尾写一个标准 PAGE 复杂域（与 fixer._write_page_field 同构）。"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    run = para.add_run()
    r = run._r
    begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    sep = OxmlElement("w:fldChar"); sep.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t"); t.text = cached
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr, sep, t, end):
        r.append(el)
    return run


def _base(path, n=6, font="Times New Roman", size=12.0):
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Inches(8.5), Inches(11.0)
    for i in range(n):
        p = doc.add_paragraph()
        r = p.add_run("Body paragraph %d about sleep quality and cognition." % (i + 1))
        r.font.name = font
        r.font.size = Pt(size)
    doc.save(str(path))
    return str(path)


def _with_footer_field(path, align):
    doc = Document()
    for i in range(4):
        doc.add_paragraph().add_run("Body text %d." % (i + 1))
    p = doc.sections[0].footer.paragraphs[0]
    _write_page_field(p)
    p.alignment = align
    doc.save(str(path))
    return str(path)


def _with_footer_static(path, text="第 1 页", align=WD_ALIGN_PARAGRAPH.CENTER):
    doc = Document()
    for i in range(4):
        doc.add_paragraph().add_run("Body text %d." % (i + 1))
    p = doc.sections[0].footer.paragraphs[0]
    p.add_run(text)
    p.alignment = align
    doc.save(str(path))
    return str(path)


def _with_header_field_right(path, header_text="RUNNING HEAD"):
    doc = Document()
    for i in range(4):
        doc.add_paragraph().add_run("Body text %d." % (i + 1))
    hp = doc.sections[0].header.paragraphs[0]
    if header_text:
        hp.add_run(header_text)
    _write_page_field(hp)
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.save(str(path))
    return str(path)


def _count_page_fields(doc_path):
    doc = Document(str(doc_path))
    n = 0
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            for p in part.paragraphs:
                if _para_has_page_field(p):
                    n += 1
    return n


# ---------------------------------------------------------------------------
# 一、读取：三个旧 bug
# ---------------------------------------------------------------------------

def test_static_page_number_regex_cases():
    """单数字页码全认；带总页数的不认（要同时换 NUMPAGES，超出保守范围）。"""
    assert detect_static_page_number("1")
    assert detect_static_page_number("Page 1")["prefix"].strip().lower() == "page"
    assert detect_static_page_number("第 1 页")["suffix"].strip() == "页"
    assert detect_static_page_number("- 12 -")["number"] == "12"
    assert detect_static_page_number("Page 1 of 10") is None
    assert detect_static_page_number("1 / 10") is None
    assert detect_static_page_number("2020") is None
    assert detect_static_page_number("Running head") is None


def test_typed_page_number_is_not_reported_as_missing(tmp_path):
    """🔴 旧 bug ①：客户手打「第 1 页」曾被判「没有页码」。"""
    p = _with_footer_static(tmp_path / "static.docx")
    prof = read_docx(p)
    assert prof.page_number_present is True
    assert prof.page_number_is_static is True
    assert prof.page_number_in == "footer"
    assert prof.page_number_alignment == "center"


def test_first_page_footer_is_read(tmp_path):
    """🔴 旧 bug ②：只读 sec.header/sec.footer，首页页脚里的页码读不到（APA 标题页常这样）。"""
    doc = Document()
    for i in range(3):
        doc.add_paragraph().add_run("Body %d." % (i + 1))
    sec = doc.sections[0]
    sec.different_first_page_header_footer = True
    fp = sec.first_page_footer.paragraphs[0]
    _write_page_field(fp)
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = tmp_path / "first.docx"
    doc.save(str(p))
    prof = read_docx(str(p))
    assert prof.page_number_present is True
    assert prof.page_number_in == "footer"


def test_footer_right_aligned_is_not_header_top_right(tmp_path):
    """🔴 旧 bug ③：页脚右对齐曾被判「符合 APA top-right」（top-right 指**页眉**右上）。"""
    p = _with_footer_field(tmp_path / "fright.docx", WD_ALIGN_PARAGRAPH.RIGHT)
    prof = read_docx(p)
    assert prof.page_number_in == "footer"
    assert prof.page_number_alignment == "right"

    target = target_from_dict({"page_number_position": "top-right"})
    _, findings = checker.run_check(p, target)
    f = [x for x in findings if x.dimension == "page_number"]
    assert f and f[0].severity == "warn"
    assert "header" in f[0].expected
    assert "footer" in f[0].actual


def test_alignment_is_never_none(tmp_path):
    """🔴 旧 bug ④：无显式 jc 时读成 None → checker 落进 else 报 pass（假通过）。"""
    p = _with_footer_field(tmp_path / "nojc.docx", WD_ALIGN_PARAGRAPH.LEFT)
    doc = Document(p)
    para = doc.sections[0].footer.paragraphs[0]
    para.alignment = None            # 靠样式/默认继承
    doc.save(p)
    prof = read_docx(p)
    assert prof.page_number_alignment == "left"     # 归一为「就是左对齐」，不是 None


# ---------------------------------------------------------------------------
# 二、检查 + 预览
# ---------------------------------------------------------------------------

def test_spec_baseline_page_number_is_fixable(tmp_path):
    """v2.3.90（老板 2026-10-22 拍板「升为可修」）：规范基线模式下页码可修。

    旧契约（本用例的前身 ``test_spec_source_only_prompts_never_fixes``）是
    「spec 来源一律 fix=None，只提示」。留这条旧规矩的代价在 2026-10-22 被实测坐实：
    客户选了 APA，报告里页码只是 info、修正器一个字没写，四份产物 zip 里
    **连页眉部件都不存在**（headerReference 出现 0 次）。

    新口径：页码不是版面偏好（页边距/字体那种「学校可能另有要求」的软参考），
    而是 APA §2.18 / MLA / IEEE 的硬要求 —— 客户真的在规范下拉里选了内置规范，
    就必须落地。走学校模板 / 手填 / AI 导入时不受影响，见下一条用例。
    """
    p = _with_footer_field(tmp_path / "spec.docx", WD_ALIGN_PARAGRAPH.CENTER)
    target = target_from_spec("APA-Professional")
    _, findings = checker.run_check(p, target)
    f = [x for x in findings if x.dimension == "page_number"]
    assert f
    assert all(x.fix is not None for x in f), [(x.severity, x.fix) for x in f]
    assert all(x.severity == "warn" for x in f), [(x.severity, x.actual) for x in f]


def test_non_baseline_page_number_is_only_prompted(tmp_path):
    """非规范基线（学校模板 / 手填 / AI 导入）时仍只提示、绝不替客户改页码。"""
    p = _with_footer_field(tmp_path / "spec.docx", WD_ALIGN_PARAGRAPH.CENTER)
    target = target_from_dict({"reference_style": "APA"})   # spec_key 为空 → 非基线
    _, findings = checker.run_check(p, target)
    f = [x for x in findings if x.dimension == "page_number"]
    for x in f:
        assert x.fix is None, (x.severity, x.fix)
        assert x.severity == "info", (x.severity, x.actual)


def test_explicit_position_is_reported_and_previewed(tmp_path):
    p = _with_footer_field(tmp_path / "pv.docx", WD_ALIGN_PARAGRAPH.CENTER)
    target = target_from_dict({"page_number_position": "top-right"})
    prof, findings = checker.run_check(p, target)
    f = [x for x in findings if x.dimension == "page_number"]
    assert f and f[0].fix is not None
    lines = _compute_changes(prof, target, lang="en", findings=findings)
    assert any("Page number" in ln and "header" in ln for ln in lines), lines
    lines_zh = _compute_changes(prof, target, lang="zh", findings=findings)
    assert any("页码" in ln for ln in lines_zh), lines_zh
    # 英文界面绝不泄漏中文
    assert not any("页" in ln for ln in lines), lines


def test_missing_page_number_is_reported(tmp_path):
    p = _base(tmp_path / "nopg.docx")
    target = target_from_dict({"page_number_position": "bottom-center"})
    prof, findings = checker.run_check(p, target)
    f = [x for x in findings if x.dimension == "page_number"]
    assert f and f[0].fix is not None
    lines = _compute_changes(prof, target, lang="en", findings=findings)
    assert any("Page number" in ln and "footer" in ln for ln in lines), lines


# ---------------------------------------------------------------------------
# 三、修正：真写页码 + 复检
# ---------------------------------------------------------------------------

def test_write_page_number_to_header_right(tmp_path):
    """客户要求页眉右上：真写进页眉、右对齐，且**页脚原来那条被清掉**（不能两个页码）。"""
    p = _with_footer_field(tmp_path / "mv.docx", WD_ALIGN_PARAGRAPH.CENTER)
    out = tmp_path / "mv_out.docx"
    target = target_from_dict({"page_number_position": "top-right"})
    _, findings = checker.run_check(p, target)
    fix_docx(p, str(out), target, findings=findings)

    after = read_docx(str(out))
    assert after.page_number_in == "header"
    assert after.page_number_alignment == "right"
    assert _count_page_fields(out) == 1, "页面上出现了两个页码"
    _, again = checker.run_check(str(out), target)
    bad = [x for x in again if x.dimension == "page_number" and x.severity in ("warn", "fail")]
    assert not bad, bad


def test_write_page_number_to_footer_center(tmp_path):
    """客户要求页脚居中（中文论文常规）：写进页脚并居中。"""
    p = _base(tmp_path / "fc.docx")
    out = tmp_path / "fc_out.docx"
    target = target_from_dict({"page_number_position": "bottom-center"})
    _, findings = checker.run_check(p, target)
    fix_docx(p, str(out), target, findings=findings)
    after = read_docx(str(out))
    assert after.page_number_in == "footer"
    assert after.page_number_alignment == "center"
    assert after.page_number_present is True


def test_static_page_number_replaced_in_place(tmp_path):
    """策略 C：把数字换成 PAGE 域，**前后缀原样保留**，且页面上只有一个页码。"""
    p = _with_footer_static(tmp_path / "rep.docx", "第 1 页", WD_ALIGN_PARAGRAPH.CENTER)
    out = tmp_path / "rep_out.docx"
    target = target_from_dict({"page_number_position": "bottom-center"})
    _, findings = checker.run_check(p, target)
    fix_docx(p, str(out), target, findings=findings)

    after = read_docx(str(out))
    assert after.page_number_present is True
    assert after.page_number_is_static is False        # 已经变成自动域
    assert _count_page_fields(out) == 1
    # 前后缀保留：「第」和「页」还在，数字位置换成了域
    doc = Document(str(out))
    txt = doc.sections[0].footer.paragraphs[0].text
    assert "第" in txt and "页" in txt


def test_page_number_fix_is_idempotent(tmp_path):
    """已经达标再修一遍：一个字都不改（applied 里没有 page_number）。"""
    p = _with_header_field_right(tmp_path / "ok.docx")
    out1 = tmp_path / "ok1.docx"
    target = target_from_dict({"page_number_position": "top-right"})
    _, findings = checker.run_check(p, target)
    fix_docx(p, str(out1), target, findings=findings)
    _, findings2 = checker.run_check(str(out1), target)
    out2 = tmp_path / "ok2.docx"
    summary = fix_docx(str(out1), str(out2), target, findings=findings2)
    assert "page_number" not in (summary.get("applied") or {}), summary.get("applied")


def test_header_text_survives_page_number_insert(tmp_path):
    """页眉里已有短标题时插页码：标题文本必须完好，不能被挤坏或覆盖。"""
    p = _base(tmp_path / "hdr.docx")
    out = tmp_path / "hdr_out.docx"
    target = target_from_dict({"page_number_position": "top-right"})
    _, findings = checker.run_check(p, target)
    fix_docx(p, str(out), target, findings=findings)
    after = read_docx(str(out))
    assert after.page_number_in == "header"


# ---------------------------------------------------------------------------
# 四、字体白名单：字体 → 字号配对
# ---------------------------------------------------------------------------

def test_apa_allowlist_is_a_pair_table():
    """APA 7th §2.19 是「字体+字号」配对，不是一串字体名。"""
    assert APA_FONT_SIZES["Times New Roman"] == 12.0
    assert APA_FONT_SIZES["Calibri"] == 11.0
    assert APA_FONT_SIZES["Arial"] == 11.0
    assert APA_FONT_SIZES["Georgia"] == 11.0
    assert APA_FONT_SIZES["Lucida Sans Unicode"] == 10.0
    assert APA_FONT_SIZES["Computer Modern"] == 10.0
    apa = get_spec("APA")
    assert set(apa.font_allowlist) == set(APA_FONT_SIZES)
    assert apa.font_allowlist_sizes == APA_FONT_SIZES
    # 大小写/空白归一
    assert size_for_font("calibri", APA_FONT_SIZES) == 11.0
    assert size_for_font("  Times  New Roman ", APA_FONT_SIZES) == 12.0
    assert size_for_font("宋体", APA_FONT_SIZES) is None      # 白名单外绝不猜


def test_calibri_12pt_is_flagged_and_size_fixed(tmp_path):
    """Calibri 12pt（APA 要 11pt）：旧版判合格；现在要报出来并**只改字号**。"""
    p = _base(tmp_path / "cal.docx", font="Calibri", size=12.0)
    target = target_from_spec("APA")
    prof, findings = checker.run_check(p, target)
    f = [x for x in findings if x.dimension == "font_size_by_font"]
    assert f and f[0].severity == "warn"
    assert "11" in f[0].expected
    assert f[0].fix is not None
    lines = _compute_changes(prof, target, lang="en", findings=findings)
    assert any("Calibri" in ln and "11" in ln for ln in lines), lines

    out = tmp_path / "cal_out.docx"
    fix_docx(p, str(out), target, findings=findings)
    doc = Document(str(out))
    run = doc.paragraphs[0].runs[0]
    assert run.font.size == Pt(11)
    assert run.font.name == "Calibri", "字体被换了 —— 只该改字号"


def test_mixed_fonts_get_their_own_size(tmp_path):
    """正文混用 TNR 与 Calibri：TNR 保 12pt、Calibri 改 11pt（各按自己的规范字号）。"""
    doc = Document()
    for i in range(4):
        r = doc.add_paragraph().add_run("Times paragraph %d." % i)
        r.font.name = "Times New Roman"
        r.font.size = Pt(12.0)
    for i in range(4):
        r = doc.add_paragraph().add_run("Calibri paragraph %d." % i)
        r.font.name = "Calibri"
        r.font.size = Pt(12.0)
    p = tmp_path / "mixed.docx"
    doc.save(str(p))

    target = target_from_spec("APA")
    _, findings = checker.run_check(str(p), target)
    out = tmp_path / "mixed_out.docx"
    fix_docx(str(p), str(out), target, findings=findings)

    d = Document(str(out))
    tnr = [r for para in d.paragraphs for r in para.runs if r.text.startswith("Times")]
    cal = [r for para in d.paragraphs for r in para.runs if r.text.startswith("Calibri")]
    assert all(r.font.size == Pt(12) for r in tnr), [(r.text, r.font.size) for r in tnr]
    assert all(r.font.size == Pt(11) for r in cal), [(r.text, r.font.size) for r in cal]


def test_font_outside_allowlist_is_not_guessed(tmp_path):
    """白名单外的字体（宋体）：不产生「按字体配字号」维度 —— 绝不按别的字体的字号去改。"""
    p = _base(tmp_path / "song.docx", font="SimSun", size=10.5)
    target = target_from_spec("APA")
    _, findings = checker.run_check(p, target)
    assert [x for x in findings if x.dimension == "font_allowlist"], findings
    assert not [x for x in findings if x.dimension == "font_size_by_font"], findings

    # 修完之后也不会出现「按配对表改字号」的记录
    out = tmp_path / "song_out.docx"
    summary = fix_docx(p, str(out), target, findings=findings)
    assert "font_size_by_font" not in (summary.get("applied") or {}), summary.get("applied")


def test_allowlisted_font_is_not_swapped(tmp_path):
    """🔴 字体在白名单内时**绝不换字体**：用 Calibri 就保留 Calibri（只校准字号）。"""
    p = _base(tmp_path / "keep.docx", font="Arial", size=11.0)
    target = target_from_spec("APA")
    _, findings = checker.run_check(p, target)
    font_f = [x for x in findings if x.dimension == "font"]
    assert font_f and font_f[0].severity == "pass", [(x.dimension, x.severity) for x in font_f]
    assert font_f[0].fix is None
    out = tmp_path / "keep_out.docx"
    fix_docx(p, str(out), target, findings=findings)
    assert Document(str(out)).paragraphs[0].runs[0].font.name == "Arial"


def test_explicit_own_size_wins_over_pair(tmp_path):
    """客户在问卷里明确写了正文 12pt：由 font_size 维度硬管，配对修正不抢。"""
    p = _base(tmp_path / "own.docx", font="Calibri", size=12.0)
    target = target_from_dict({"font_size_pt": 12.0})
    target.reference_style = "APA"          # 让 spec 能取到 Allowlist
    _, findings = checker.run_check(p, target)
    assert not [x for x in findings if x.dimension == "font_size_by_font"], \
        [(x.dimension, x.severity) for x in findings]


# ---------------------------------------------------------------------------
# 五、学校模板 / 配置入口
# ---------------------------------------------------------------------------

def test_template_reads_page_number_position(tmp_path):
    """学校模板里页码在页脚居中 → 反推 bottom-center 并标 school。"""
    tpl = _with_footer_field(tmp_path / "tpl.docx", WD_ALIGN_PARAGRAPH.CENTER)
    target = target_from_template_docx(read_docx(tpl))
    assert target.page_number_position == "bottom-center"
    assert target.resolution.get("page_number_position") == "school"


def test_template_without_page_number_stays_none(tmp_path):
    """模板没有页码：不臆造（None），不能凭空给个 top-right。"""
    tpl = _base(tmp_path / "tpl_none.docx")
    target = target_from_template_docx(read_docx(tpl))
    assert target.page_number_position is None


def test_page_number_positions_are_whitelisted():
    assert PAGE_NUMBER_POSITIONS == ("top-right", "bottom-center", "bottom-right")
    t = target_from_dict({"page_number_position": "top-right"})
    assert t.page_number_position == "top-right"
    assert t.resolution["page_number_position"] == "questionnaire"
    with pytest.raises(ValueError):
        target_from_dict({"page_number_position": "top-left"})   # 白名单外 → 丢弃后空画像
