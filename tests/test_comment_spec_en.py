# -*- coding: utf-8 -*-
"""英文模板批注解析器（comment_spec_en）的单元测试。

覆盖：示例模板 15 条批注、标题层级变体、单位换算、栏数、三线表、表后间隙、
参考文献流派、图表题位置、否定语义、以及「编辑性批注不得被误读」的门控。
"""

from __future__ import annotations

import pytest

from src.engine.comment_spec_en import parse_comments


def _parse(*texts):
    return parse_comments([{"id": i + 1, "text": t} for i, t in enumerate(texts)])


# ---------------------------------------------------------------------------
# 整体：示例模板的 15 条英文批注
# ---------------------------------------------------------------------------
def test_sample_template_comments_end_to_end():
    from tools.make_sample_template_en import COMMENTS

    spec = _parse(*COMMENTS)

    assert spec["count"] == len(COMMENTS)
    assert spec["unparsed"] == [], spec["unparsed"]

    page = spec["page"]
    assert page["font_family"] == "Times New Roman"
    assert page["font_size_pt"] == 12.0
    assert page["line_spacing"] == 2.0
    assert page["body_alignment"] == "justify"
    assert page["first_line_indent_in"] == 0.5
    assert [page["margin_%s_in" % s] for s in ("top", "bottom", "left", "right")] == [1.0] * 4

    assert spec["structural"]["paper_size"] == "Letter"
    assert spec["structural"]["table_lines"] == "three_line"
    assert spec["structural"]["columns"] == 1

    h = spec["heading_formats"]
    assert h[1]["size_pt"] == 18.0 and h[1]["alignment"] == "center"
    assert h[1]["space_before_pt"] == 12.0 and h[1]["space_after_pt"] == 6.0
    assert h[2]["size_pt"] == 14.0
    assert h[3]["italic"] is True

    assert spec["title_format"]["size_pt"] == 24.0
    assert spec["title_format"]["line_spacing"] == 1.0
    assert spec["title_page"]["title_page_title_centered"] is True
    assert spec["title_page"]["title_page_title_bold"] is True

    assert spec["abstract"]["abstract_heading_alignment"] == "center"
    assert spec["abstract"]["abstract_heading_bold"] is True
    assert spec["abstract"]["abstract_first_line_indent_in"] == 0.0
    assert spec["abstract"]["keywords_label_italic"] is True

    assert spec["table"]["table_gap_pt"] == 24.0
    assert spec["caption"]["caption_position_table"] == "top"
    assert spec["caption"]["caption_position_figure"] == "bottom"

    assert spec["reference"]["reference_style"] == "apa"
    assert spec["reference"]["reference_hanging_indent_in"] == 0.5
    assert spec["reference"]["reference_line_spacing"] == 2.0


# ---------------------------------------------------------------------------
# 标题层级：多种写法都要认
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("text,level", [
    ("Heading 1: Times New Roman 18 pt, bold", 1),
    ("H2: Arial 13 pt, bold, left-aligned", 2),
    ("Level 3 headings: 11 pt, italic", 3),
    ("First-level heading: 16 pt, bold, centered", 1),
    ("second level heading: 12 pt", 2),
])
def test_heading_level_variants(text, level):
    spec = _parse(text)
    assert level in spec["heading_formats"]
    assert spec["unparsed"] == []


# ---------------------------------------------------------------------------
# 单位换算 / 页面结构
# ---------------------------------------------------------------------------
def test_margins_in_cm_converted_to_inches():
    spec = _parse("Margins: 2.54 cm")
    page = spec["page"]
    for side in ("top", "bottom", "left", "right"):
        assert page["margin_%s_in" % side] == pytest.approx(1.0, abs=0.01)


def test_margins_per_side():
    spec = _parse("Margins: top 1.5 inch, bottom 1 inch, left 1.25 in, right 1 in.")
    page = spec["page"]
    assert page["margin_top_in"] == 1.5
    assert page["margin_left_in"] == 1.25
    assert page["margin_bottom_in"] == 1.0


def test_paper_size_aliases():
    assert _parse("Paper size: A4")["structural"]["paper_size"] == "A4"
    assert _parse("Paper size: Letter, 8.5 x 11 inches.")["structural"]["paper_size"] == "Letter"


def test_paper_size_from_dimensions_only():
    """只写尺寸没写纸型名时，要复用 identify_paper_size 归一（口径与实测一致）。"""
    assert _parse("Paper size: 210 x 297 mm")["structural"]["paper_size"] == "A4"
    assert _parse("Paper size: 8.5 x 11 in")["structural"]["paper_size"] == "Letter"


@pytest.mark.parametrize("text,cols", [
    ("Layout: single column, one column per page.", 1),
    ("Layout: two columns", 2),
    ("Double-column layout required", 2),
])
def test_columns(text, cols):
    assert _parse(text)["structural"]["columns"] == cols


# ---------------------------------------------------------------------------
# 三线表 / 表后间隙
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("text", [
    "Tables: use three-line tables with no vertical lines.",
    "Tables: no vertical lines, horizontal rules only.",
    "Tables: formal table style.",
])
def test_three_line_table(text):
    assert _parse(text)["structural"]["table_lines"] == "three_line"


def test_table_gap_in_pt():
    assert _parse("Tables: leave 12 pt after each table.")["table"]["table_gap_pt"] == 12.0


def test_table_gap_blank_line_uses_body_metrics():
    spec = _parse("Body text: 12 pt, double-spaced.",
                  "Leave one blank line after each table before the following text.")
    # 一行 ≈ 12pt × 2.0 = 24pt
    assert spec["table"]["table_gap_pt"] == pytest.approx(24.0)


def test_table_gap_not_confused_with_three_line():
    spec = _parse("Tables: three-line tables, no vertical lines.")
    assert "table_gap_pt" not in spec["table"]


# ---------------------------------------------------------------------------
# 参考文献
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("text,style", [
    ("References: APA 7th edition.", "apa"),
    ("References: MLA 9th edition, double-spaced.", "mla"),
    ("References: IEEE style, numbered [1].", "ieee"),
])
def test_reference_style(text, style):
    assert _parse(text)["reference"]["reference_style"] == style


def test_reference_details():
    spec = _parse("References: APA 7th edition, hanging indent 0.5 inch, "
                  "Times New Roman 10 pt, double-spaced, numbered.")
    ref = spec["reference"]
    assert ref["reference_hanging_indent_in"] == 0.5
    assert ref["reference_size_pt"] == 10.0
    assert ref["reference_font_family"] == "Times New Roman"
    assert ref["reference_line_spacing"] == 2.0
    assert ref["reference_numbered"] is True


def test_reference_not_numbered():
    assert _parse("References: not numbered.")["reference"]["reference_numbered"] is False


# ---------------------------------------------------------------------------
# 图表题位置 / 否定语义
# ---------------------------------------------------------------------------
def test_caption_positions():
    spec = _parse("Table captions: 12 pt, placed above the table.",
                  "Figure captions: 12 pt, placed below the figure.")
    assert spec["caption"]["caption_position_table"] == "top"
    assert spec["caption"]["caption_position_figure"] == "bottom"


def test_negative_indent_and_bold():
    spec = _parse("Abstract body: 12 pt, double-spaced, no first-line indent.")
    assert spec["abstract"]["abstract_first_line_indent_in"] == 0.0

    spec2 = _parse("Heading 2: 14 pt, not bold, left-aligned.")
    assert spec2["heading_formats"][2]["bold"] is False


# ---------------------------------------------------------------------------
# 门控：编辑性批注绝不能当格式规范
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("text", [
    "Please fix this later.",
    "TODO: check with supervisor.",
    "This paragraph was moved from chapter 2.",
])
def test_editorial_comment_is_ignored(text):
    spec = _parse(text)
    assert spec["unparsed"] and spec["unparsed"][0]["text"] == text
    assert spec["page"] == {}
    assert spec["heading_formats"] == {}


def test_empty_comments_list():
    spec = parse_comments([])
    assert spec["count"] == 0
    assert spec["unparsed"] == []
    assert spec["heading_formats"] == {}
