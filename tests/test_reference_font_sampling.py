# -*- coding: utf-8 -*-
"""参考文献**条目字体**采样回归（v2.3.115）。

老板 2026-10-30 实测：把学校模板参考文献的编号括号 `[1]` 手动删成 `1` 之后，
修正稿的参考文献字体被刷成了「宋体」。根因：Word 给新敲的编号 run 自动贴上
``w:rFonts w:hint="eastAsia" w:eastAsia="宋体"``，而条目正文 run 一个字体都没写
（继承 Normal 的 Times New Roman）→ 整块唯一的「显式字体」就是那个宋体 →
众数一取就把参考文献字体判成宋体。

本文件锁住四条口径：
1. 编号 run 的 eastAsia 不得被当成条目字体证据；
2. 条目 run 真写了拉丁字体（ascii/hAnsi）时照常读到；
3. 条目文本含东亚字符时，eastAsia 是真证据（中文文献条目）；
4. 同一份模板「带括号 / 不带括号」两种形态，读出的条目字体必须一致（回归本体）。
"""
from __future__ import annotations

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

from src.engine.docx_reader import read_docx


def _set_rfonts(run, **kw):
    """给 run 写 w:rFonts（模拟 Word 自动贴上的 eastAsia / hint）。"""
    rPr = run._r.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = rPr.makeelement(qn("w:rFonts"), {})
        rPr.append(rf)
    for k, v in kw.items():
        rf.set(qn("w:" + k), v)


def _make_template(path, marker_run_eastasia=True, entry_text=None):
    """造一份最小模板：正文 + References 标题 + 一条参考文献条目。

    marker_run_eastasia=True 时，把编号单独做成一个 run 并贴上
    ``hint=eastAsia / eastAsia=宋体``（这正是 Word 在中文输入法下的行为）。
    """
    doc = Document()
    doc.add_paragraph("Body text sample.")
    doc.add_paragraph("References")
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(36)
    p.paragraph_format.first_line_indent = Pt(-36)
    if marker_run_eastasia:
        r0 = p.add_run("1")
        _set_rfonts(r0, hint="eastAsia", eastAsia="宋体")
        p.add_run(entry_text or " Smith, J. (2020). An English title. Publisher.")
    else:
        p.add_run(entry_text or "1 Smith, J. (2020). An English title. Publisher.")
    doc.save(path)
    return str(path)


def test_marker_run_eastasia_is_not_font_evidence(tmp_path):
    """编号 run 只有 eastAsia=宋体、条目正文没写字体 → 读不到条目字体（None）。

    绝不能读成「宋体」—— 否则修正器会把英文条目全刷成宋体。
    """
    p = _make_template(tmp_path / "t.docx", marker_run_eastasia=True)
    prof = read_docx(p)
    assert prof.reference_entries, "夹具应被识别出参考文献条目"
    assert prof.reference_font_family is None, (
        "编号 run 的 eastAsia 不得当成条目字体证据，实测读到 %r"
        % (prof.reference_font_family,))


def test_explicit_latin_font_still_read(tmp_path):
    """条目 run 真写了拉丁字体 → 必须照常读到（不能因为加固而漏读）。"""
    doc = Document()
    doc.add_paragraph("Body text.")
    doc.add_paragraph("References")
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(36)
    p.paragraph_format.first_line_indent = Pt(-36)
    r = p.add_run("Smith, J. (2020). An English title. Publisher.")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    src = str(tmp_path / "explicit.docx")
    doc.save(src)

    assert read_docx(src).reference_font_family == "Times New Roman"


def test_cjk_entry_eastasia_is_real_evidence(tmp_path):
    """条目文本含东亚字符（中文题名）→ eastAsia 是真证据，必须读到宋体。"""
    doc = Document()
    doc.add_paragraph("Body text.")
    doc.add_paragraph("References")
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(36)
    p.paragraph_format.first_line_indent = Pt(-36)
    r = p.add_run("Zhang, S. (2020). 供应链韧性的实证研究. 管理科学, 12(3), 45-52.")
    _set_rfonts(r, hint="eastAsia", eastAsia="宋体")
    r.font.size = Pt(10.5)
    src = str(tmp_path / "cjk.docx")
    doc.save(src)

    assert read_docx(src).reference_font_family == "宋体", (
        "中文条目只设 eastAsia 时必须读到宋体"
    )


def test_bracket_removal_does_not_change_reference_font(tmp_path):
    """回归本体：同一份模板「带 [1] / 不带括号」两种编号形态，条目字体必须一致。

    老板就是删了括号才出问题，所以这一条直接对照两种形态。
    """
    with_brackets = _make_template(
        tmp_path / "with.docx", marker_run_eastasia=False,
        entry_text="[1] Smith, J. (2020). An English title. Publisher.")
    # 删括号后：Word 把 "1" 拆成独立 run 并贴上 eastAsia=宋体
    without_brackets = _make_template(
        tmp_path / "without.docx", marker_run_eastasia=True)

    f1 = read_docx(with_brackets).reference_font_family
    f2 = read_docx(without_brackets).reference_font_family
    assert f1 == f2, (
        "带括号与不带括号读出的条目字体必须一致，实测 %r vs %r" % (f1, f2))
    assert f1 is None, "模板条目没写字体时不应臆造字体，实测 %r" % (f1,)
