"""Tests for the redesigned format-source routing (Tasks #26 / #27).

- docx_reader extracts per-level heading formats.
- target_from_template_docx maps header_distance_in + heading_formats and
  injects the chosen reference style's switches (no-spec-baseline mode).
- build_target supports a no-spec-baseline mode (school / custom source):
  spec_key stays None, page comes from the template, reference is driven by
  the reference_style argument (school sub-dropdown) or questionnaire/ai data.
- Reference style left empty -> no reference checking (reference_style None).
"""

from __future__ import annotations

import os
import sys

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine.docx_reader import read_docx
from src.engine.questionnaire import (
    target_from_template_docx, build_target, TargetProfile, apply_page_overrides,
)
from src.engine.specs import resolve_spec


def _make_heading_doc(path):
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)
        s.header_distance = Inches(0.5)
    # L1: centered, bold, 16pt, numbered "1."
    h1 = doc.add_paragraph("1. Introduction", style="Heading 1")
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in h1.runs:
        r.font.bold = True
        r.font.size = Pt(16)
    # L2: left, bold (not italic), numbered "A."
    h2 = doc.add_paragraph("A. Method", style="Heading 2")
    h2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for r in h2.runs:
        r.font.bold = True
        r.font.size = Pt(14)
    # L3: left, bold italic, numbered "I."
    h3 = doc.add_paragraph("I. Results", style="Heading 3")
    h3.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for r in h3.runs:
        r.font.bold = True
        r.font.italic = True
        r.font.size = Pt(12)
    # L4: left, bold, all-caps (no numbering)
    h4 = doc.add_paragraph("APPENDIX", style="Heading 4")
    h4.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for r in h4.runs:
        r.font.bold = True
    # body
    b = doc.add_paragraph("Body paragraph text here to inspect margins and fonts.")
    b.runs[0].font.name = "Arial"
    b.runs[0].font.size = Pt(11)
    doc.save(path)
    return path


def test_heading_formats_extracted(tmp_path):
    p = _make_heading_doc(str(tmp_path / "h.docx"))
    prof = read_docx(p)
    assert prof.heading_levels_used == 4
    assert set(prof.heading_formats.keys()) == {1, 2, 3, 4}

    f1 = prof.heading_formats[1].to_dict()
    assert f1["alignment"] == "center"
    assert f1["bold"] is True
    assert isinstance(f1["size_pt"], float) and f1["size_pt"] >= 12
    assert f1["number_prefix"] == "number"

    f2 = prof.heading_formats[2].to_dict()
    assert f2["alignment"] == "left"
    assert f2["bold"] is True
    assert f2["number_prefix"] == "upper_letter"

    f3 = prof.heading_formats[3].to_dict()
    assert f3["italic"] is True
    assert f3["number_prefix"] == "roman"

    f4 = prof.heading_formats[4].to_dict()
    assert f4["uppercase"] is True
    assert f4["number_prefix"] is None

    # header_distance_in read from the section
    assert prof.header_distance_in == 0.5


def test_template_docx_maps_header_and_headings(tmp_path):
    p = _make_heading_doc(str(tmp_path / "t.docx"))
    prof = read_docx(p)
    t = target_from_template_docx(prof, reference_style=None)
    # 无规范基线：spec_key 必须为 None
    assert t.spec_key is None
    # 页面值来自学校模板（origin=school）
    assert t.page.get("header_distance_in") == 0.5
    assert t.page.get("margin_top_in") == 1.0
    assert t.resolution.get("header_distance_in") == "school"
    # 每级标题格式被带入
    assert t.heading_formats and 1 in t.heading_formats
    assert t.heading_formats[1]["alignment"] == "center"
    # 该合成模板**没有参考文献批注**、也未选流派 -> 才真的不查
    # （「留空」本身 = 按学校模板，不是「不检查参考文献」）
    assert t.reference_style is None
    assert t.reference_hanging_indent_in is None


def test_template_docx_injects_reference_style(tmp_path):
    p = _make_heading_doc(str(tmp_path / "t2.docx"))
    prof = read_docx(p)
    t = target_from_template_docx(prof, reference_style="APA")
    assert t.spec_key is None
    assert t.reference_style == "APA"
    # 悬挂缩进 / 编码方式自动带入该风格规范值（APA=0.5 / 著者-年制）
    assert t.reference_hanging_indent_in == 0.5
    assert t.reference_numbered is False
    assert t.resolution.get("reference_style") == "school"


def test_build_target_no_spec_baseline_school(tmp_path):
    p = _make_heading_doc(str(tmp_path / "b.docx"))
    # 学校模板来源：spec_key=None，reference_style 由子下拉驱动
    t = build_target(spec_key=None, template=p, reference_style="MLA")
    assert t.spec_key is None
    assert "school" in t.source
    # 页面来自模板，不被任何规范值覆盖
    assert t.page.get("margin_top_in") == 1.0
    assert t.resolution.get("margin_top_in") == "school"
    # 引用风格生效
    assert t.reference_style == "MLA"
    # 不应带规范基线页面值（如 APA 的 first_line_indent 0.5 劫持）
    # 学校模板正文无首行缩进 -> 实测为 0.0，绝不可能是规范的 0.5
    assert t.page.get("first_line_indent_in") == 0.0


def test_build_target_spec_baseline_preserved():
    # 规范模式行为不变：基线为规范，页面软参考
    t = build_target("APA-Student")
    assert t.spec_key == "APA-Student"
    assert t.reference_style == "APA"
    # 规范页面值带 resolution=spec（checker 会降级为 info）
    assert t.resolution.get("margin_top_in") == "spec"
    assert t.page.get("first_line_indent_in") == 0.5


def test_build_target_no_reference_means_no_check(tmp_path):
    p = _make_heading_doc(str(tmp_path / "n.docx"))
    t = build_target(spec_key=None, template=p, reference_style=None)
    assert t.reference_style is None


def test_apply_page_overrides_only_changed_upgrade():
    """复核弹窗写回：仅改动过的字段升级为 user，未改动保持原 resolution。

    关键不变量：规范模式里弹窗预填了 spec 页面值，用户「什么都不改直接 Run」时，
    绝不能把 spec 软参考（info 降级）误升级成 user 硬规则（warn/fail）。
    """
    t = TargetProfile(
        source="spec", spec_key="APA-Student",
        page={"margin_top_in": 1.0, "font_family": "Times New Roman"},
        resolution={"margin_top_in": "spec", "font_family": "spec"},
    )
    # 模拟弹窗回传：margin_top 与规范一致（未改）、font_family 改了、font_size 新填、空值跳过
    apply_page_overrides(t, {
        "margin_top_in": "1.0",
        "font_family": "Arial",
        "font_size_pt": "12.0",
        "header_distance_in": "",
    })
    # 未改动 -> 维持 spec 软参考
    assert t.page["margin_top_in"] == 1.0
    assert t.resolution["margin_top_in"] == "spec"
    # 改动 -> 升级为 user
    assert t.page["font_family"] == "Arial"
    assert t.resolution["font_family"] == "user"
    # 新填 -> 升级为 user（数值化）
    assert t.page["font_size_pt"] == 12.0
    assert t.resolution["font_size_pt"] == "user"
    # 空值 -> 不写回
    assert "header_distance_in" not in t.page


def test_apply_page_overrides_rejects_garbage():
    """非数值字段填了非法 float 应被忽略，不污染 target。"""
    t = TargetProfile(source="school",
                      page={"margin_top_in": 1.0},
                      resolution={"margin_top_in": "school"})
    apply_page_overrides(t, {"margin_top_in": "abc"})
    assert t.page["margin_top_in"] == 1.0  # 非法值被跳过
    assert t.resolution["margin_top_in"] == "school"
    # checker 经 resolve_spec 落到 Other，无引用硬判
    spec = resolve_spec(t)
    assert spec.key == "Other"


def test_empty_spec_key_means_no_spec_baseline():
    """spec_key 为空串时必须等同于「无规范」，不能当成一个空规范。

    v2.3.65 实测到的真问题：UI 的 spec_key 是 StringVar，「无规范」存的是 ``""``；
    旧判定写的是 ``spec_key is not None``，于是空串被当成合法规范键 → 客户选了
    「学校模板 / 自定义」，背后仍套着一层规范基线，凭空注入
    reference_style / reference_numbered / reference_hanging_indent_in 三项硬规则，
    跟「留空即不检查」直接矛盾。三种"无规范"写法必须完全等价。
    """
    for sk in (None, "", "Other"):
        t = build_target(sk, questionnaire={"margin_left_in": "1.5"})
        assert t.source == "user", "spec_key=%r 不该带规范基线（source=%r）" % (sk, t.source)
        assert set(t.resolution) == {"margin_left_in"}, \
            "spec_key=%r 凭空注入了别的维度：%r" % (sk, sorted(t.resolution))
        assert t.reference_style is None


def test_no_spec_baseline_does_not_inject_reference_rules():
    """无规范基线 + 问卷没写参考文献 → 参考文献三项都得是 None（不检查）。"""
    t = build_target("", questionnaire={"font_family": "Arial"})
    assert t.reference_style is None
    assert t.reference_numbered is None
    assert t.reference_hanging_indent_in is None
