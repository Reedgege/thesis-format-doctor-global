# -*- coding: utf-8 -*-
"""锁住两个「报告说改好了、文档里其实没改/改坏了」的缺陷。

都是 2026-11-15 站客户角度验收时抓到的，共同点：**回执与文档不一致**。
这类问题的杀伤力比"直接报错"大得多 —— 客户看到"字体已改为 TNR 12"，
打开文档还是 Arial，从此不再信任任何结论。

缺陷 1：页脚里的简单域（fldSimple）页码清不掉
--------------------------------------------
Word 的「插入 → 页码」生成的是 ``w:fldSimple`` 简单域。改页码位置时，
``_strip_page_numbers_in_part`` 只遍历 ``p.runs``，而 fldSimple **不是 w:r**，
删不到它。后果：页眉加了页码、页脚原来那条还在 → 页面上**两个页码**，
而 ``removed_elsewhere`` 还报"已清理 1 处"。

缺陷 2：改动清单按「目标值」而非「真改了什么」渲染
--------------------------------------------------
``_summary_lines`` 里 margins / Body font / Line spacing 三条读的是 ``page``
（规范要求什么），其余读 ``applied``（实际改了什么）。后果：只要客户选了规范，
就无条件打印「Body font → Times New Roman 12pt」，哪怕一个字都没动。
"""

from __future__ import annotations

import docx
import pytest

from src import main as cli
from src.engine.fixer import _strip_page_numbers_in_part


# ---------------------------------------------------------------------------
# 缺陷 1：fldSimple 简单域
# ---------------------------------------------------------------------------
def _para_with_simple_page_field(doc):
    """造一个「页脚里 Word 那种简单域页码」。"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    p = doc.sections[0].footer.paragraphs[0]
    run = p.add_run()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), " PAGE  \\* MERGEFORMAT ")
    t = OxmlElement("w:t")
    t.text = "1"
    fld.append(t)
    run._r.addnext(fld)
    return p


def test_simple_field_page_number_is_removed():
    """fldSimple 形态的页码必须能被清掉（否则页面上出现两个页码）。"""
    from src.engine.docx_reader import _para_has_page_field

    doc = docx.Document()
    doc.add_paragraph("Body")
    p = _para_with_simple_page_field(doc)
    assert _para_has_page_field(p), "前置条件：这段确实带 PAGE 域"

    n = _strip_page_numbers_in_part(doc.sections[0].footer)
    assert n == 1, "应该报清理了 1 处"
    assert not _para_has_page_field(p), "清理之后不该再有 PAGE 域"
    assert p._p.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldSimple") is None


def test_complex_field_page_number_still_removed():
    """复杂域（fldChar/instrText）这条路不能被顺手改坏。"""
    from src.engine.docx_reader import _para_has_page_field
    from src.engine.fixer import _write_page_field

    doc = docx.Document()
    doc.add_paragraph("Body")
    p = doc.sections[0].footer.paragraphs[0]
    _write_page_field(p)
    assert _para_has_page_field(p)

    assert _strip_page_numbers_in_part(doc.sections[0].footer) == 1
    assert not _para_has_page_field(p)


def test_no_page_field_returns_zero():
    """没有页码的页脚不该被误报「清理了 N 处」。"""
    doc = docx.Document()
    doc.add_paragraph("Body")
    footer = doc.sections[0].footer
    footer.paragraphs[0].add_run("Confidential")
    assert _strip_page_numbers_in_part(footer) == 0
    assert footer.paragraphs[0].text == "Confidential", "有实质文字时一个字都不能动"


# ---------------------------------------------------------------------------
# 缺陷 2：改动清单不许谎报
# ---------------------------------------------------------------------------
def _summary(applied: dict) -> dict:
    return {
        "page": {
            "margin_top_in": 1.0, "margin_bottom_in": 1.0,
            "margin_left_in": 1.0, "margin_right_in": 1.0,
            "font_family": "Times New Roman", "font_size_pt": 12.0,
            "line_spacing": 2.0,
        },
        "applied": applied,
    }


def test_changes_do_not_claim_font_when_untouched():
    """字体一个字没改时，绝不能出现 "Body font → …"。"""
    lines = cli._summary_lines(_summary({"margins": 1.0, "line_spacing": 2.0}))
    joined = "\n".join(lines)
    assert "Body font" not in joined, "谎报字体已改：\n" + joined
    assert "Line spacing" in joined, "行距真改了，应该写上"
    assert "Margins" in joined, "页边距真改了，应该写上"


def test_changes_do_not_claim_line_spacing_when_untouched():
    lines = cli._summary_lines(_summary({"font_family": "Times New Roman",
                                        "font_size_pt": 12.0}))
    joined = "\n".join(lines)
    assert "Line spacing" not in joined, "谎报行距已改：\n" + joined
    assert "Body font" in joined, "字体真改了，应该写上"


def test_changes_are_empty_when_nothing_applied():
    """什么都没改时，清单必须是空的（而不是把规范目标抄一遍）。"""
    assert cli._summary_lines(_summary({})) == []
