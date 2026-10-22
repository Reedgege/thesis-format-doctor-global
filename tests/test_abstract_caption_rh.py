"""摘要页特化 + running head + heading 分页 + 图表 caption 四条链路回归（v2.3.67）。

老板 2026-10-19 从海外学术市场（APA/MLA/Chicago/IEEE/ACM/Harvard）角度指出模板 5 个
高频致命缺失项，本轮落地其中 4 项（脚注/尾注因 python-docx 硬限制只提示不做）：
①摘要页特化（abstract_first_line_indent_in + abstract_new_page）——全链路；
②Running Head（running_head_required + page_number_include_author_name）——可查可提示；
③heading page_break_before——全链路；
④图表 caption（字号/行距可修 + 位置仅提示）——部分。
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
    read_docx, find_abstract_section, is_caption_paragraph,
)
from src.engine.fixer import fix_docx, _compute_changes
from src.engine.questionnaire import (
    target_from_dict, target_from_spec, TargetProfile,
)
from src.engine.specs import get_spec


# ---------------------------------------------------------------------------
# 夹具
# ---------------------------------------------------------------------------

def _doc_with_abstract(path, indent=0.5, new_page=False):
    """构造一个含摘要（Heading 1 "Abstract" + 摘要正文）的文档。"""
    doc = Document()
    h = doc.add_paragraph("Abstract")
    h.style = doc.styles["Heading 1"]
    if new_page:
        h.paragraph_format.page_break_before = True
    a = doc.add_paragraph("This is the abstract body text that summarizes the study.")
    a.paragraph_format.first_line_indent = Inches(indent)
    # 正文若干段
    for i in range(4):
        p = doc.add_paragraph("Body paragraph %d about the research methods." % (i + 1))
        p.paragraph_format.first_line_indent = Inches(0.5)
    doc.save(str(path))
    return str(path)


def _doc_with_captions(path):
    """构造一个含表格 caption（上方）+ 图 caption（Caption 样式）的文档。"""
    doc = Document()
    # 表题在上
    tc = doc.add_paragraph("Table 1. Sample table caption")
    tc.paragraph_format.first_line_indent = Inches(0)
    for r in tc.runs:
        r.font.size = Pt(12)
    doc.add_table(rows=1, cols=1)
    # 图题（Caption 样式）在下
    fc = doc.add_paragraph("Figure 1. Sample figure caption")
    fc.style = doc.styles["Caption"]
    for r in fc.runs:
        r.font.size = Pt(12)
    for i in range(3):
        doc.add_paragraph("Body text %d." % (i + 1))
    doc.save(str(path))
    return str(path)


# ---------------------------------------------------------------------------
# ① 摘要页特化
# ---------------------------------------------------------------------------

def test_read_abstract_indent_and_new_page(tmp_path):
    p = _doc_with_abstract(tmp_path / "a.docx", indent=0.5, new_page=True)
    prof = read_docx(p)
    assert prof.has_abstract is True
    assert prof.abstract_first_line_indent_in == pytest.approx(0.5, abs=0.01)
    assert prof.abstract_new_page is True


def test_find_abstract_section_returns_head_and_body(tmp_path):
    p = _doc_with_abstract(tmp_path / "a.docx")
    doc = Document(p)
    head, body = find_abstract_section(list(doc.paragraphs))
    assert head is not None
    assert len(body) >= 1
    assert "abstract" in doc.paragraphs[head].text.lower()


def test_check_abstract_indent_warn(tmp_path):
    p = _doc_with_abstract(tmp_path / "a.docx", indent=0.5)
    t = target_from_dict({"abstract_first_line_indent_in": 0.0}, "user")
    prof, findings = checker.run_check(p, t)
    dims = {f.dimension: f for f in findings if f.dimension == "abstract_indent"}
    assert dims, "应有 abstract_indent 维度"
    assert dims["abstract_indent"].severity in ("warn", "fail")


def test_check_abstract_new_page(tmp_path):
    p = _doc_with_abstract(tmp_path / "a.docx", new_page=False)
    t = target_from_dict({"abstract_new_page": True}, "user")
    prof, findings = checker.run_check(p, t)
    dims = {f.dimension: f for f in findings if f.dimension == "abstract_new_page"}
    assert dims, "应有 abstract_new_page 维度"
    assert dims["abstract_new_page"].severity in ("warn", "fail")


def test_fix_abstract_indent_to_zero(tmp_path):
    p = _doc_with_abstract(tmp_path / "a.docx", indent=0.5)
    out = tmp_path / "out.docx"
    t = target_from_dict({"abstract_first_line_indent_in": 0.0}, "user")
    summary = fix_docx(p, str(out), t)
    # 摘要正文段应被设为 0 缩进
    prof2 = read_docx(str(out))
    assert prof2.abstract_first_line_indent_in == pytest.approx(0.0, abs=0.01)


def test_fix_abstract_new_page(tmp_path):
    p = _doc_with_abstract(tmp_path / "a.docx", new_page=False)
    out = tmp_path / "out.docx"
    t = target_from_dict({"abstract_new_page": True}, "user")
    summary = fix_docx(p, str(out), t)
    prof2 = read_docx(str(out))
    assert prof2.abstract_new_page is True


def test_abstract_new_page_is_fixable_on_spec_baseline(tmp_path):
    """规范基线（客户真在规范下拉里选了内置规范）时，摘要另起一页可修。

    v2.3.91 契约升级（与 page_number / running_head 同口径）：APA 7 §2.9 要求摘要自成
    一页，客户既然明确选了 APA，这就是硬要求，不该停在「提示」上。旧断言
    「spec 来源一律只提示」在 v2.3.90 已被页码打破，这里是同一条链的补全。
    """
    p = _doc_with_abstract(tmp_path / "a.docx", new_page=False)
    t = target_from_spec("APA-Professional")
    prof, findings = checker.run_check(p, t)
    abs_findings = [f for f in findings if f.dimension == "abstract_new_page"]
    assert abs_findings
    assert any(getattr(f, "fix", None) is not None for f in abs_findings), \
        "规范基线下的摘要另起一页应当可修"
    # 真修一遍 → 复检必须收敛
    out = tmp_path / "out.docx"
    fix_docx(p, str(out), t)
    assert read_docx(str(out)).abstract_new_page is True


def test_abstract_new_page_stays_prompt_only_when_not_baseline(tmp_path):
    """非基线（学校模板 / 手填 / AI 导入）即便引用了 APA，也只提示不动手。

    模板说了算：客户的学校可能明确要求「摘要与正文接排」，工具不许替它拍板。
    """
    p = _doc_with_abstract(tmp_path / "a.docx", new_page=False)
    t = target_from_dict({"reference_style": "APA"}, "school")
    prof, findings = checker.run_check(p, t)
    abs_findings = [f for f in findings if f.dimension == "abstract_new_page"]
    assert abs_findings
    assert all(getattr(f, "fix", None) is None for f in abs_findings), \
        "非规范基线的摘要另起一页只应提示"


# ---------------------------------------------------------------------------
# ② Running head
# ---------------------------------------------------------------------------

def test_target_dict_running_head(tmp_path):
    t = target_from_dict({"running_head_required": True,
                          "page_number_include_author_name": True}, "user")
    assert t.running_head_required is True
    assert t.page_number_include_author_name is True


def test_check_running_head_required_warn(tmp_path):
    """无页眉文本时，显式 running_head_required=True 应 warn。"""
    doc = Document()
    for i in range(4):
        doc.add_paragraph("Body %d." % (i + 1))
    p = tmp_path / "rh.docx"
    doc.save(str(p))
    t = target_from_dict({"running_head_required": True}, "user")
    prof, findings = checker.run_check(str(p), t)
    rh = [f for f in findings if f.dimension == "running_head"]
    assert rh
    assert any(f.severity in ("warn", "fail") for f in rh)


def test_check_page_number_include_author_name(tmp_path):
    """页眉无姓氏时，显式 page_number_include_author_name=True 应 warn。"""
    doc = Document()
    for i in range(4):
        doc.add_paragraph("Body %d." % (i + 1))
    # 页眉只有一个页码域，无姓氏文本
    doc.sections[0].header.paragraphs[0].text = "1"
    p = tmp_path / "rh2.docx"
    doc.save(str(p))
    t = target_from_dict({"page_number_include_author_name": True}, "user")
    prof, findings = checker.run_check(str(p), t)
    rh = [f for f in findings if f.dimension == "running_head_name"]
    assert rh
    assert any(f.severity in ("warn", "fail") for f in rh)


# ---------------------------------------------------------------------------
# ③ heading page_break_before
# ---------------------------------------------------------------------------

def test_heading_level_format_parses_page_break_before():
    from src.engine.specs import heading_level_format_from_dict
    f = heading_level_format_from_dict({"alignment": "left", "page_break_before": True})
    assert f.page_break_before is True
    f2 = heading_level_format_from_dict({"page_break_before": False})
    assert f2.page_break_before is False


def test_target_dict_heading_page_break(tmp_path):
    t = target_from_dict({"heading_formats": {"1": {"page_break_before": True}}}, "user")
    assert t.heading_formats[1].get("page_break_before") is True


def test_fix_heading_page_break_before(tmp_path):
    """Heading 1 标题原本不分页，目标要求 page_break_before=True 后应分页。"""
    doc = Document()
    h = doc.add_paragraph("Introduction")
    h.style = doc.styles["Heading 1"]
    for i in range(3):
        doc.add_paragraph("Body %d." % (i + 1))
    p = tmp_path / "h.docx"
    doc.save(str(p))
    out = tmp_path / "h_out.docx"
    t = target_from_dict({"heading_formats": {"1": {"page_break_before": True}}}, "user")
    fix_docx(str(p), str(out), t)
    # 重新读，检查 Heading 1 段前分页
    from docx import Document as _D
    d2 = _D(str(out))
    h2 = None
    for para in d2.paragraphs:
        if para.text == "Introduction":
            h2 = para
            break
    assert h2 is not None
    assert h2.paragraph_format.page_break_before is True


# ---------------------------------------------------------------------------
# ④ 图表 caption
# ---------------------------------------------------------------------------

def test_read_captions(tmp_path):
    p = _doc_with_captions(tmp_path / "c.docx")
    prof = read_docx(p)
    assert prof.caption_count >= 2
    assert prof.caption_font_size_pt is not None


def test_is_caption_paragraph_detects_prefix():
    doc = Document()
    p1 = doc.add_paragraph("Table 1. Results")
    p2 = doc.add_paragraph("This is not a caption.")
    p3 = doc.add_paragraph("Figure 2. Overview")
    assert is_caption_paragraph(p1) is True
    assert is_caption_paragraph(p2) is False
    assert is_caption_paragraph(p3) is True


def test_fix_caption_size(tmp_path):
    p = _doc_with_captions(tmp_path / "c.docx")
    out = tmp_path / "c_out.docx"
    t = target_from_dict({"caption_font_size_pt": 10.0}, "user")
    fix_docx(str(p), str(out), t)
    prof2 = read_docx(str(out))
    # caption 字号应被改为 10pt（众数）
    assert prof2.caption_font_size_pt == pytest.approx(10.0, abs=0.5)


def test_check_caption_position_table_hint(tmp_path):
    """表题位置只提示（info），不硬判、不改。"""
    p = _doc_with_captions(tmp_path / "c.docx")
    t = target_from_dict({"caption_position_table": "top"}, "user")
    prof, findings = checker.run_check(str(p), t)
    pos = [f for f in findings if f.dimension == "caption_position_table"]
    assert pos
    # 位置维度不绑 fix
    assert all(getattr(f, "fix", None) is None for f in pos)


# ---------------------------------------------------------------------------
# 预览 & 幂等
# ---------------------------------------------------------------------------

def test_preview_lists_abstract_and_caption_changes(tmp_path):
    p = _doc_with_abstract(tmp_path / "a.docx", indent=0.5)
    t = target_from_dict({"abstract_first_line_indent_in": 0.0,
                          "abstract_new_page": True}, "user")
    prof, findings = checker.run_check(p, t)
    lines = _compute_changes(prof, t, "en", findings=findings)
    joined = " ".join(lines)
    assert "abstract" in joined.lower() or "Abstract" in joined


def test_fix_is_idempotent_for_caption(tmp_path):
    p = _doc_with_captions(tmp_path / "c.docx")
    out1 = tmp_path / "c1.docx"
    out2 = tmp_path / "c2.docx"
    t = target_from_dict({"caption_font_size_pt": 10.0}, "user")
    fix_docx(str(p), str(out1), t)
    fix_docx(str(out1), str(out2), t)
    prof1 = read_docx(str(out1))
    prof2 = read_docx(str(out2))
    assert prof1.caption_font_size_pt == pytest.approx(prof2.caption_font_size_pt, abs=0.01)


# ---------------------------------------------------------------------------
# 手动输入 == AI 导入：同一套 schema / 同一份 target_from_dict
# ---------------------------------------------------------------------------

def _manual_form_save_shape():
    """逼真复刻 app.py ``_open_questionnaire`` 的 ``_save`` 实际产出：

    - float/int 字段经 Entry 拿到的是**字符串**（_save 不转 float，留给 _coerce 转）；
    - tri 字段产出 True / False（「—」被跳过，不进数据）；
    - choice 字段产出 schema options 里的原值字符串（"top"/"bottom"）。
    这张字典就是手填弹窗「保存」后 ``self.questionnaire_data`` 的真实形态。
    """
    return {
        # 既有字段（保证解析非空）
        "margin_top_in": "1.0",
        "font_size_pt": "12.0",
        # —— 本轮新增 10 个字段，全部以手填弹窗的真实输出类型给出 ——
        "abstract_first_line_indent_in": "0.0",       # float → 字符串
        "abstract_new_page": True,                    # tri → bool
        "running_head_required": True,                # tri → bool
        "page_number_include_author_name": False,      # tri → bool
        "caption_font_size_pt": "9.0",                # float → 字符串
        "caption_line_spacing": "1.0",                # float → 字符串
        "caption_position_table": "top",              # choice → 选项原值
        "caption_position_figure": "bottom",          # choice → 选项原值
    }


def test_manual_form_output_maps_into_targetprofile():
    """手填弹窗保存的字典能正确落进 TargetProfile（resolution=user）。"""
    t = target_from_dict(_manual_form_save_shape(), "user")
    # 类型被 _coerce 正确还原
    assert t.abstract_first_line_indent_in == 0.0
    assert t.abstract_new_page is True
    assert t.running_head_required is True
    assert t.page_number_include_author_name is False
    assert t.caption_font_size_pt == 9.0
    assert t.caption_line_spacing == 1.0
    assert t.caption_position_table == "top"
    assert t.caption_position_figure == "bottom"
    # 来源记成 user（手填），与 AI 导入的 ai_import 区分但走同一函数
    assert t.resolution.get("abstract_new_page") == "user"
    assert t.resolution.get("caption_position_table") == "user"


def test_manual_and_ai_import_converge_to_same_target():
    """老板原话：手动输入跟 AI 导入是同一套路径。

    同一份字段值，分别以 source="user"（手填）与 source="ai_import"（AI 产出文件再导入）
    喂给同一个 ``target_from_dict``，得到的 TargetProfile 关键字段必须完全一致——
    证明两者不是两套逻辑，只差一个来源标签。
    """
    payload = _manual_form_save_shape()
    t_user = target_from_dict(payload, "user")
    t_ai = target_from_dict(payload, "ai_import")
    for key in ("abstract_first_line_indent_in", "abstract_new_page",
                "running_head_required", "page_number_include_author_name",
                "caption_font_size_pt", "caption_line_spacing",
                "caption_position_table", "caption_position_figure"):
        assert getattr(t_user, key) == getattr(t_ai, key), key
    # 只有来源标签不同，字段值完全相同
    assert t_user.resolution.get("caption_font_size_pt") == "user"
    assert t_ai.resolution.get("caption_font_size_pt") == "ai_import"
