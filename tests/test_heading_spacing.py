# -*- coding: utf-8 -*-
"""标题级 **段前 / 段后 / 行距** 的完整链路回归（v2.3.75）。

背景（老板 2026-10-20 追问）：「咱们整个调整格式之中，段前段后有没有考虑？比如说
标题的段前段后……这是不是也是非常核心的东西？你读读那些规范，那些规范里面有没有要求？」

事实核查结论：此前整条链路只有**正文段后**一项（PageLayout.space_after_pt），
标题级连字段都没有 —— 学校模板写着「章标题段前 1 行、段后 0.5 行」，
读不到、查不出、也改不了。规范侧确有硬要求：
  · APA 7th §2.21 Do not add blank lines before or after headings → 段前段后都为 0
  · APA 7th §2.21 Double-space the entire paper（含标题）→ 标题行距 2.0
  · IEEE 官方未规定具体磅值（IEEEtran 模板各自为政）→ 一律 None，绝不臆造

本文件逐环节钉死：模板能读 → 检查报得出 → 修正改得对 → 预览提得前 → 二次幂等。
"""
import os

import pytest
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

from src.engine.docx_reader import read_docx, _spacing_pair_of_para
from src.engine.checker import run_check
from src.engine.fixer import fix_docx, compute_changes
from src.engine.specs import (HeadingLevelFormat, describe_heading,
                              heading_level_format_from_dict, get_spec)
from src.engine.questionnaire import (HEADING_LEVEL_KEYS, TargetProfile,
                                      build_target, target_from_dict,
                                      target_from_template_docx)


# ---------------------------------------------------------------------------
# 造数据工具
# ---------------------------------------------------------------------------

def _set_lines(para, before=None, after=None):
    """给段落写 w:beforeLines / w:afterLines（单位 1/100 行）——中文学校模板常态。"""
    pPr = para._p.get_or_add_pPr()
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        sp = pPr.makeelement(qn("w:spacing"), {})
        pPr.append(sp)
    if before is not None:
        sp.set(qn("w:beforeLines"), str(int(round(before * 100))))
    if after is not None:
        sp.set(qn("w:afterLines"), str(int(round(after * 100))))


def _style_spacing(doc, style_name, before_pt=None, after_pt=None):
    """把段前/段后写到**样式**上（模板制作者最常干的事）。"""
    st = doc.styles[style_name]
    if before_pt is not None:
        st.paragraph_format.space_before = Pt(before_pt)
    if after_pt is not None:
        st.paragraph_format.space_after = Pt(after_pt)


def _template(path, *, before_lines=1.0, after_lines=0.5, size_h1=14.0):
    """学校模板：L1 用「行」单位写段前段后；L2 用磅值。"""
    doc = Document()
    n = doc.styles["Normal"]
    n.font.name = "Times New Roman"
    n.font.size = Pt(12)
    h = doc.styles["Heading 1"]
    h.font.name = "Times New Roman"
    h.font.size = Pt(size_h1)
    h1 = doc.add_paragraph("Chapter One Introduction", style="Heading 1")
    _set_lines(h1, before=before_lines, after=after_lines)
    doc.add_paragraph("Body text long enough to be sampled as body text. " * 3)
    h2 = doc.add_paragraph("Section One Background", style="Heading 2")
    h2.paragraph_format.space_before = Pt(6)
    h2.paragraph_format.space_after = Pt(6)
    doc.add_paragraph("More body text so the body sample is stable. " * 3)
    doc.save(str(path))
    return str(path)


def _paper(path, *, h1_before=0.0, h1_after=0.0, h2_before=0.0, h2_after=0.0):
    """学生论文：章节标题段前段后全 0（最常见形态）。"""
    doc = Document()
    n = doc.styles["Normal"]
    n.font.name = "Times New Roman"
    n.font.size = Pt(12)
    n.paragraph_format.line_spacing = 2.0
    h1 = doc.add_paragraph("Introduction", style="Heading 1")
    h1.paragraph_format.space_before = Pt(h1_before)
    h1.paragraph_format.space_after = Pt(h1_after)
    doc.add_paragraph("The body paragraph that must keep its own spacing rules. " * 4)
    h2 = doc.add_paragraph("Background", style="Heading 2")
    h2.paragraph_format.space_before = Pt(h2_before)
    h2.paragraph_format.space_after = Pt(h2_after)
    doc.add_paragraph("Another body paragraph, long enough for sampling. " * 4)
    doc.save(str(path))
    return str(path)


def _hf(before=None, after=None, line=None):
    d: dict = {}
    if before is not None:
        d["space_before_pt"] = before
    if after is not None:
        d["space_after_pt"] = after
    if line is not None:
        d["line_spacing"] = line
    return d


def _target(hf_map, **page):
    return target_from_dict({"heading_formats": hf_map, **page}, "user")


def _para_by_text(path, text, *, contains=False):
    paras = Document(path).paragraphs
    if contains:
        return [p for p in paras if text in p.text][0]
    return [p for p in paras if p.text.strip() == text][0]


def _h1_para(path):
    """取一级标题段（IEEE 会把它编号 + 全大写，按文字匹配不可靠）。"""
    return [p for p in Document(path).paragraphs
            if (p.style.name or "") == "Heading 1"][0]


# ---------------------------------------------------------------------------
# 一、模型层：三态字段 + 序列化 + describe
# ---------------------------------------------------------------------------

def test_heading_model_has_tri_state_spacing_fields():
    f = HeadingLevelFormat()
    assert f.space_before_pt is None and f.space_after_pt is None and f.line_spacing is None
    assert HeadingLevelFormat(space_before_pt=0.0).space_before_pt == 0.0


def test_dict_round_trip_keeps_spacing():
    raw = {"space_before_pt": "12", "space_after_pt": 6, "line_spacing": "1.5"}
    f = heading_level_format_from_dict(raw)
    assert (f.space_before_pt, f.space_after_pt, f.line_spacing) == (12.0, 6.0, 1.5)


def test_describe_heading_mentions_apa_no_extra_space():
    txt = describe_heading(get_spec("APA-Student").heading_formats[1])
    assert "no extra space above or below" in txt
    assert "2x line spacing" in txt


def test_ieee_leaves_spacing_unspecified():
    """IEEE 官方没规定标题段前段后磅值 → 必须留 None，不许臆造。"""
    for lv in (1, 2, 3, 4, 5):
        f = get_spec("IEEE").heading_formats[lv]
        assert f.space_before_pt is None
        assert f.space_after_pt is None
        assert f.line_spacing is None


def test_apa_requires_no_extra_space_on_every_level():
    for lv in (1, 2, 3, 4, 5):
        f = get_spec("APA-Student").heading_formats[lv]
        assert f.space_before_pt == 0.0
        assert f.space_after_pt == 0.0
        assert f.line_spacing == 2.0


# ---------------------------------------------------------------------------
# 二、读取层：磅值 / 「行」单位 / 样式链 三级回退
# ---------------------------------------------------------------------------

def test_reads_point_values_directly(tmp_path):
    p = _paper(tmp_path / "a.docx", h1_before=6.0, h1_after=3.0)
    prof = read_docx(p)
    s = prof.heading_formats[1]
    assert s.space_before_pt == pytest.approx(6.0)
    assert s.space_after_pt == pytest.approx(3.0)


def test_converts_line_units_to_points(tmp_path):
    """w:beforeLines=100（1 行）@14pt → 1.2 × 14 = 16.8 磅；段后 0.5 行 → 8.4 磅。"""
    tpl = _template(tmp_path / "tpl.docx", before_lines=1.0, after_lines=0.5, size_h1=14.0)
    prof = read_docx(tpl)
    s = prof.heading_formats[1]
    assert s.space_before_pt == pytest.approx(16.8)
    assert s.space_after_pt == pytest.approx(8.4)


def test_reads_spacing_carried_by_style_chain(tmp_path):
    """段前段后挂在 Heading 1 样式上（模板常态），段落本身没有直接格式。"""
    p = tmp_path / "s.docx"
    doc = Document()
    _style_spacing(doc, "Heading 1", before_pt=18, after_pt=9)
    doc.add_paragraph("Introduction", style="Heading 1")
    doc.add_paragraph("Body text for sampling, long enough. " * 3)
    doc.save(str(p))
    s = read_docx(str(p)).heading_formats[1]
    assert s.space_before_pt == pytest.approx(18.0)
    assert s.space_after_pt == pytest.approx(9.0)


def test_unset_spacing_reads_as_zero_not_none(tmp_path):
    """整条链都没写 → Word 渲染就是 0 磅。返回 None 会被判「读不到」→ 检查有洞却照改。"""
    p = tmp_path / "z.docx"
    doc = Document()
    para = doc.add_paragraph("A plain paragraph whose style carries no spacing at all.")
    doc.save(str(p))
    b, a = _spacing_pair_of_para(para)
    assert b == 0.0 and a == 0.0


def test_template_target_carries_spacing(tmp_path):
    tpl = _template(tmp_path / "tpl.docx")
    t = target_from_template_docx(read_docx(tpl))
    assert t.heading_formats[1]["space_before_pt"] == pytest.approx(16.8)
    assert t.heading_formats[1]["space_after_pt"] == pytest.approx(8.4)
    assert t.heading_formats[2]["space_before_pt"] == pytest.approx(6.0)


# ---------------------------------------------------------------------------
# 三、检查层：报得出，且只在「目标规定了」时报
# ---------------------------------------------------------------------------

def test_checker_reports_spacing_mismatch_per_level(tmp_path):
    tpl = _template(tmp_path / "tpl.docx")
    paper = _paper(tmp_path / "paper.docx")
    t = build_target(spec_key=None, template=tpl)
    _, findings = run_check(paper, t)
    lvl1 = [f for f in findings
            if f.dimension == "heading" and f.message == "Level 1 heading format"]
    assert lvl1, "一级标题的段前段后差异必须报出来"
    f = lvl1[0]
    assert f.severity != "pass"
    assert "space before 16.8 pt" in f.expected
    assert "space after 8.4 pt" in f.expected
    assert "space before 0 pt" in f.actual
    # 精确到级：修正器据此只改报错的那一级
    assert f.fix is not None and f.fix.levels == (1,)


def test_checker_silent_when_target_says_nothing(tmp_path):
    """目标没规定段前段后（None）→ 一个字都不许报，更不许改。"""
    paper = _paper(tmp_path / "paper.docx", h1_before=24.0, h1_after=12.0)
    t = _target({"1": {"size_pt": 14}})
    _, findings = run_check(paper, t)
    hits = [f for f in findings
            if f.dimension == "heading" and "space" in (f.expected or "").lower()]
    assert not hits


def test_checker_passes_when_spacing_already_matches(tmp_path):
    paper = _paper(tmp_path / "paper.docx", h1_before=6.0, h1_after=6.0)
    t = _target({"1": _hf(before=6.0, after=6.0)})
    _, findings = run_check(paper, t)
    f = [x for x in findings if x.message == "Level 1 heading format"][0]
    assert f.severity == "pass", (f.expected, f.actual)


# ---------------------------------------------------------------------------
# 四、修正层：改得对 + 幂等 + 「行」单位必须清掉
# ---------------------------------------------------------------------------

def test_fixer_applies_spacing(tmp_path):
    paper = _paper(tmp_path / "paper.docx")
    out = str(tmp_path / "fixed.docx")
    t = _target({"1": _hf(before=12.0, after=6.0)})
    _, findings = run_check(paper, t)
    fix_docx(paper, out, t, findings=findings)
    h = _para_by_text(out, "Introduction").paragraph_format
    assert h.space_before == Pt(12)
    assert h.space_after == Pt(6)
    # 正文段落的段后不受标题规则牵连
    body = [p for p in Document(out).paragraphs if p.text.startswith("The body paragraph")][0]
    assert body.paragraph_format.space_before is None


def test_fixer_drops_line_units_so_the_value_actually_sticks(tmp_path):
    """🔴 写磅值前必须删 w:beforeLines：OOXML 里「行」单位优先，留着等于没改。"""
    p = tmp_path / "lines.docx"
    doc = Document()
    para = doc.add_paragraph("Introduction", style="Heading 1")
    _set_lines(para, before=2.0, after=1.0)
    doc.add_paragraph("Body text for sampling, long enough. " * 3)
    doc.save(str(p))

    out = str(tmp_path / "fixed.docx")
    t = _target({"1": _hf(before=6.0, after=6.0)})
    _, findings = run_check(str(p), t)
    fix_docx(str(p), out, t, findings=findings)

    sp = _para_by_text(out, "Introduction")._p.pPr.find(qn("w:spacing"))
    assert sp.get(qn("w:beforeLines")) is None, "w:beforeLines 没删掉，Word 会忽略磅值"
    assert sp.get(qn("w:afterLines")) is None
    s = read_docx(out).heading_formats[1]
    assert s.space_before_pt == pytest.approx(6.0)
    assert s.space_after_pt == pytest.approx(6.0)


def test_fixer_applies_heading_line_spacing(tmp_path):
    paper = _paper(tmp_path / "paper.docx")
    out = str(tmp_path / "fixed.docx")
    t = _target({"1": _hf(line=1.5)})
    _, findings = run_check(paper, t)
    fix_docx(paper, out, t, findings=findings)
    h = _para_by_text(out, "Introduction").paragraph_format
    assert h.line_spacing == pytest.approx(1.5)


def test_second_pass_is_a_no_op(tmp_path):
    tpl = _template(tmp_path / "tpl.docx")
    paper = _paper(tmp_path / "paper.docx")
    t = build_target(spec_key=None, template=tpl)
    out1 = str(tmp_path / "f1.docx")
    out2 = str(tmp_path / "f2.docx")
    _, findings = run_check(paper, t)
    fix_docx(paper, out1, t, findings=findings)
    _, findings2 = run_check(out1, t)
    fix_docx(out1, out2, t, findings=findings2)
    # 二次修正：段前段后相关的结论必须全部 pass，预览必须为空
    bad = [f for f in findings2
           if f.dimension == "heading" and f.severity != "pass"
           and "space" in (f.expected or "").lower()]
    assert not bad, [(f.message, f.expected, f.actual) for f in bad]
    assert compute_changes(out1, t, lang="en") == []


# ---------------------------------------------------------------------------
# 五、预览层：预览说的就是修正真会做的
# ---------------------------------------------------------------------------

def test_preview_lists_spacing_changes(tmp_path):
    paper = _paper(tmp_path / "paper.docx")
    t = _target({"1": _hf(before=12.0, after=6.0)})
    lines = compute_changes(paper, t, lang="zh")
    joined = " | ".join(lines)
    assert "段前" in joined and "段后" in joined
    assert "0pt → 12pt" in joined
    assert "0pt → 6pt" in joined


def test_preview_is_empty_when_nothing_differs(tmp_path):
    paper = _paper(tmp_path / "paper.docx", h1_before=6.0, h1_after=6.0)
    t = _target({"1": _hf(before=6.0, after=6.0)})
    assert compute_changes(paper, t, lang="en") == []


# ---------------------------------------------------------------------------
# 六、规范基线：APA 硬性「无额外段前段后 + 双倍行距」
# ---------------------------------------------------------------------------

def test_apa_baseline_strips_heading_spacing(tmp_path):
    paper = _paper(tmp_path / "paper.docx", h1_before=24.0, h1_after=12.0)
    t = build_target(spec_key="APA-Student")
    _, findings = run_check(paper, t)
    lvl1 = [f for f in findings if f.message == "Level 1 heading format"]
    assert lvl1 and "space before 0 pt" in lvl1[0].expected

    out = str(tmp_path / "fixed.docx")
    fix_docx(paper, out, t, findings=findings)
    h = _para_by_text(out, "Introduction").paragraph_format
    assert h.space_before == Pt(0)
    assert h.space_after == Pt(0)
    assert h.line_spacing == pytest.approx(2.0)


def test_ieee_baseline_never_touches_heading_spacing(tmp_path):
    """IEEE 没规定 → 客户自己的段前段后原样保留。"""
    paper = _paper(tmp_path / "paper.docx", h1_before=24.0, h1_after=12.0)
    t = build_target(spec_key="IEEE")
    out = str(tmp_path / "fixed.docx")
    _, findings = run_check(paper, t)
    fix_docx(paper, out, t, findings=findings)
    # IEEE 会给一级标题补罗马数字编号并全大写（"I. INTRODUCTION"）→ 按样式取段
    h = _h1_para(out).paragraph_format
    assert h.space_before == Pt(24)
    assert h.space_after == Pt(12)


# ---------------------------------------------------------------------------
# 七、问卷 / AI 通道：字段进得去、认得下
# ---------------------------------------------------------------------------

def test_questionnaire_schema_knows_the_new_fields():
    for k in ("space_before_pt", "space_after_pt", "line_spacing"):
        assert k in HEADING_LEVEL_KEYS


def test_questionnaire_keeps_spacing_through_target_from_dict():
    t = target_from_dict({"heading_formats": {"1": _hf(before=12, after=6, line=1.5)}}, "user")
    f = t.heading_formats[1]
    assert f["space_before_pt"] == 12.0
    assert f["space_after_pt"] == 6.0
    assert f["line_spacing"] == 1.5


def test_ai_template_documents_the_new_fields():
    """AI 填表模板必须与代码同步，否则 AI 永远填不出段前段后。"""
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    p = os.path.join(root, "src", "data", "ai_questionnaire_template.yaml")
    text = open(p, "r", encoding="utf-8").read()
    for k in ("space_before_pt", "space_after_pt", "line_spacing"):
        assert k in text, "AI 模板缺字段 %s" % k


# ---------------------------------------------------------------------------
# 八、参考文献条目之间的段前 / 段后（v2.3.75）
# ---------------------------------------------------------------------------

_ENTRY1 = "Anderson, K., & Lee, S. (2019). Sleep quality and memory consolidation. " \
          "Journal of Sleep Research, 12(3), 45-67."
_ENTRY2 = "Smith, J. A. (2020). Title of Work. Publisher."


def _refs(path, *, entry_before=0.0, entry_after=0.0, lines=None):
    """带参考文献段的论文；条目段前段后可控（可用「行」单位写）。"""
    doc = Document()
    n = doc.styles["Normal"]
    n.font.name = "Times New Roman"
    n.font.size = Pt(12)
    doc.add_paragraph("Body paragraph for the thesis. " * 3)
    doc.add_paragraph("References")
    for txt in (_ENTRY1, _ENTRY2):
        p = doc.add_paragraph(txt)
        if lines is not None:
            _set_lines(p, before=lines[0], after=lines[1])
        else:
            p.paragraph_format.space_before = Pt(entry_before)
            p.paragraph_format.space_after = Pt(entry_after)
    doc.save(str(path))
    return str(path)


def test_template_reads_reference_entry_spacing(tmp_path):
    tpl = _refs(tmp_path / "tpl.docx", entry_before=6.0, entry_after=6.0)
    prof = read_docx(tpl)
    assert prof.reference_space_before_pt == pytest.approx(6.0)
    assert prof.reference_space_after_pt == pytest.approx(6.0)


def test_reference_entry_spacing_converts_line_units(tmp_path):
    tpl = _refs(tmp_path / "tpl.docx", lines=(1.0, 0.5))
    prof = read_docx(tpl)
    # 1 行 @12pt = 1.2 × 12 = 14.4；0.5 行 = 7.2
    assert prof.reference_space_before_pt == pytest.approx(14.4)
    assert prof.reference_space_after_pt == pytest.approx(7.2)


def test_template_target_carries_reference_entry_spacing(tmp_path):
    tpl = _refs(tmp_path / "tpl.docx", entry_before=6.0, entry_after=3.0)
    t = target_from_template_docx(read_docx(tpl))
    assert t.reference_space_before_pt == pytest.approx(6.0)
    assert t.reference_space_after_pt == pytest.approx(3.0)
    assert t.resolution.get("reference_space_before_pt") == "school"


def test_checker_reports_reference_entry_spacing(tmp_path):
    tpl = _refs(tmp_path / "tpl.docx", entry_before=6.0, entry_after=6.0)
    paper = _refs(tmp_path / "paper.docx")
    t = build_target(spec_key=None, template=tpl)
    _, findings = run_check(paper, t)
    sb = [f for f in findings if f.dimension == "reference_space_before"]
    sa = [f for f in findings if f.dimension == "reference_space_after"]
    assert sb and sb[0].severity == "fail" and sb[0].expected == "6pt"
    assert sa and sa[0].severity == "fail"
    assert sb[0].fix is not None and sb[0].fix.dimension == "reference_space_before"


def test_fixer_applies_reference_entry_spacing(tmp_path):
    paper = _refs(tmp_path / "paper.docx")
    out = str(tmp_path / "fixed.docx")
    t = _target({}, font_size_pt=12)   # 空画像会被拒 → 给一个占位字段
    t.reference_space_before_pt = 6.0
    t.reference_space_after_pt = 3.0
    t.resolution["reference_space_before_pt"] = "school"
    t.resolution["reference_space_after_pt"] = "school"
    _, findings = run_check(paper, t)
    fix_docx(paper, out, t, findings=findings)
    doc = Document(out)
    for txt in (_ENTRY1, _ENTRY2):
        pf = [p for p in doc.paragraphs if p.text.strip() == txt][0].paragraph_format
        assert pf.space_before == Pt(6)
        assert pf.space_after == Pt(3)
    # 幂等
    _, f2 = run_check(out, t)
    bad = [f for f in f2 if f.dimension.startswith("reference_space")
           and f.severity != "pass"]
    assert not bad, [(f.message, f.expected, f.actual) for f in bad]


def test_reference_entry_spacing_preview(tmp_path):
    paper = _refs(tmp_path / "paper.docx")
    t = _target({}, font_size_pt=12)
    t.reference_space_before_pt = 6.0
    t.resolution["reference_space_before_pt"] = "school"
    lines = compute_changes(paper, t, lang="zh")
    assert any("条目段前" in ln and "6.0pt" in ln for ln in lines), lines


def test_reference_entry_spacing_silent_when_unspecified(tmp_path):
    paper = _refs(tmp_path / "paper.docx", entry_before=18.0, entry_after=18.0)
    t = _target({}, font_size_pt=12)
    _, findings = run_check(paper, t)
    assert not [f for f in findings if f.dimension.startswith("reference_space")]
