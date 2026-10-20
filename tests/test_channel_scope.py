"""通道边界护栏（v2.3.73）：参考文献选了规范，绝不许外溢到标题 / 页面 / 标题页。

老板 2026-10-20 定的口径（逐字）：
  「如果用户是选择的现成的规范，那就全部按规范走。……你想想人家为什么要上传学校模板？
   为什么要手动？为什么要用 AI 导入？那不就是现行的模板满足不了他的需求嘛……
   那这个地方你就是只改它的参考文献呀，那其他地方你就按人家用户上传的地方改呀。」

由此得出两条铁律：
1. **规范基线**（客户在规范下拉里真选了 APA/IEEE/MLA）→ 页面 / 标题 / 标题页 / 参考文献
   全部按规范；
2. **非规范基线**（学校模板 / 手动输入 / AI 导入）→ 以客户提供的为准；参考文献那栏若选了
   某个引用规范，**只**按该规范核参考文献，页面与标题一个字都不许跟着动。

实测事故（v2.3.72 及以前，本文件钉死）：手动输入只填页面 + 参考文献选 IEEE + 标题层级填 2，
点修正后论文大标题变成 ``I. THE IMPACT OF AI ON ACADEMIC WRITING``、章节标题变成
``A. Introduction / B. Method / C. Results``，而预览里一个字都没提（背着客户改文字）。

造数据一律用真实形态：16pt 居中加粗的论文大标题（**没有 Abstract**，学生论文常态）、
14pt 手排章节标题（Normal 样式 + 整段加粗）、带参考文献段。
"""
from __future__ import annotations

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

from src.engine.checker import run_check
from src.engine.docx_reader import doc_title_index, read_docx
from src.engine.fixer import _compute_changes, fix_docx
from src.engine.questionnaire import (QUESTIONNAIRE_SCHEMA, build_target,
                                      target_from_dict)
from src.engine.specs import (REFERENCE_STYLE_CUSTOM, REFERENCE_STYLE_KEYS,
                              get_spec, spec_is_baseline)

TITLE = "The Impact of AI on Academic Writing"
CH = ("Introduction", "Method", "Results")
ENTRY = "Smith, J. A. (2020). Title of work. Publisher."


def _paper(tmp_path, name="paper.docx", *, title_size=16.0, ch_size=14.0):
    """学生真实形态的论文（无摘要、大标题居中、章节标题手排左对齐）。"""
    p = tmp_path / name
    doc = Document()
    n = doc.styles["Normal"]
    n.font.name = "Times New Roman"
    n.font.size = Pt(12)
    n.paragraph_format.line_spacing = 2.0
    s = doc.sections[0]
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(s, m, Inches(1.25))          # 与 APA 的 1.0″ 不同：一旦外溢就会露馅

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(TITLE)
    r.bold = True
    r.font.size = Pt(title_size)

    for name_ in CH:
        par = doc.add_paragraph()
        rr = par.add_run(name_)
        rr.bold = True
        rr.font.size = Pt(ch_size)
        doc.add_paragraph("Body text of the section. " * 6)

    h = doc.add_paragraph()
    hr = h.add_run("References")
    hr.bold = True
    hr.font.size = Pt(ch_size)
    doc.add_paragraph(ENTRY)
    doc.save(p)
    return str(p)


def _heading_texts(path):
    return [p.text.strip() for p in Document(path).paragraphs if p.text.strip()]


def _manual(page_only_ref, *, extra=None):
    """手动输入通道：只填页面值 + 参考文献那一栏。"""
    q = {"font_family": "Times New Roman", "font_size_pt": 12, "line_spacing": 2.0,
         "margin_top_in": 1.25, "margin_bottom_in": 1.25,
         "margin_left_in": 1.25, "margin_right_in": 1.25,
         "reference_style": page_only_ref}
    q.update(extra or {})
    return build_target(spec_key=None, questionnaire=q)


# ---------------------------------------------------------------------------
# 一、spec_is_baseline：只有「真在规范下拉里选了规范」才算基线
# ---------------------------------------------------------------------------

def test_spec_is_baseline_only_for_real_spec_selection():
    assert spec_is_baseline(build_target(spec_key="APA-Student"))
    assert spec_is_baseline(build_target(spec_key="IEEE"))
    # 「Other」= 不选规范（学校模板/自定义），不是基线
    assert not spec_is_baseline(build_target(spec_key="Other"))
    assert not spec_is_baseline(build_target(spec_key=None))


def test_manual_reference_style_never_becomes_spec_baseline():
    """参考文献栏选了 IEEE，也不许把画像升级成规范基线。"""
    t = _manual("IEEE")
    assert t.reference_style == "IEEE"      # 参考文献仍按 IEEE 核
    assert t.spec_key is None               # 但页面/标题不许跟着 IEEE 走
    assert not spec_is_baseline(t)


# ---------------------------------------------------------------------------
# 二、外溢事故本尊：手动输入选 IEEE + 标题层级 → 标题必须一个字都不动
# ---------------------------------------------------------------------------

def test_manual_ieee_reference_does_not_renumber_headings(tmp_path):
    paper = _paper(tmp_path)
    before = _heading_texts(paper)
    t = _manual("IEEE", extra={"heading_levels": "2"})

    prof, findings = run_check(paper, t)
    out = str(tmp_path / "fixed.docx")
    fix_docx(paper, out, t, findings=findings)

    after = _heading_texts(out)
    # 核心断言：论文大标题与章节标题原样保留（不许出现 "I. " / "A. " / 全大写重写）
    assert TITLE in after
    for name in CH:
        assert name in after, "章节标题被改写成 %r" % after
    assert after[:1] == before[:1]

    # 预览与实际一致：既然实际什么都没改，预览也不能空口说要改（反之亦然）
    changes = _compute_changes(prof, t, "en", findings=findings)
    heading_lines = [c for c in changes if "evel" in c and "heading" in c.lower()]
    assert not heading_lines


def test_manual_ieee_reference_still_numbers_reference_entries(tmp_path):
    """参考文献这一块仍要按 IEEE 走 —— 只改参考文献，不是什么都不做。"""
    paper = _paper(tmp_path)
    t = _manual("IEEE")
    prof, findings = run_check(paper, t)
    out = str(tmp_path / "fixed.docx")
    fix_docx(paper, out, t, findings=findings)
    body = "\n".join(_heading_texts(out))
    assert "[1] Smith" in body, "参考文献条目没按 IEEE 补编号：%r" % body


def test_manual_custom_reference_applies_no_citation_style(tmp_path):
    """「按我校要求（自定义）」= 不套任何引用体例，条目一个字不动。"""
    paper = _paper(tmp_path)
    t = _manual(REFERENCE_STYLE_CUSTOM)
    assert t.reference_style is None
    prof, findings = run_check(paper, t)
    out = str(tmp_path / "fixed.docx")
    fix_docx(paper, out, t, findings=findings)
    body = "\n".join(_heading_texts(out))
    assert ENTRY in body
    assert "[1]" not in body


# ---------------------------------------------------------------------------
# 三、规范基线：客户真选了 APA/IEEE，全篇仍按规范走（这条能力不许被削弱）
# ---------------------------------------------------------------------------

def test_spec_baseline_still_formats_headings_per_spec(tmp_path):
    paper = _paper(tmp_path)
    t = build_target(spec_key="APA-Student")
    prof, findings = run_check(paper, t)
    out = str(tmp_path / "fixed.docx")
    fix_docx(paper, out, t, findings=findings)
    # APA L1 = 居中加粗：论文大标题被识别为文档标题后，章节标题（L1）应被居中
    doc = Document(out)
    ch = [p for p in doc.paragraphs if p.text.strip() == "Introduction"][0]
    assert ch.alignment == WD_ALIGN_PARAGRAPH.CENTER


def test_spec_baseline_ieee_preview_lists_the_numbering_it_will_add(tmp_path):
    """编号/全大写是改文字 —— 规范基线模式下预览必须先写出来（预览 = 实际）。"""
    paper = _paper(tmp_path)
    t = build_target(spec_key="IEEE")
    prof, findings = run_check(paper, t)
    changes = _compute_changes(prof, t, "en", findings=findings)
    joined = " | ".join(changes)
    assert "numbering" in joined.lower(), "预览没提编号：%s" % joined
    out = str(tmp_path / "fixed.docx")
    fix_docx(paper, out, t, findings=findings)
    texts = _heading_texts(out)
    assert any(t_.startswith(("I.", "1.")) or t_.isupper() for t_ in texts), texts


# ---------------------------------------------------------------------------
# 四、论文大标题：不再被当成一级章节标题（检查结果不许张冠李戴）
# ---------------------------------------------------------------------------

def test_doc_title_is_recognised_without_abstract(tmp_path):
    paper = _paper(tmp_path)
    doc = Document(paper)
    assert doc_title_index(doc) == 0
    prof = read_docx(paper)
    assert prof.doc_title_index == 0
    assert prof.doc_title_size_pt == pytest.approx(16.0)


def test_doc_title_is_not_compared_as_heading_level(tmp_path):
    """模板 L1 = 14pt 左对齐；论文大标题是 16pt 居中 —— 以前会被判「L1 不符」并改小。"""
    tpl = Document()
    n = tpl.styles["Normal"]
    n.font.name = "Times New Roman"
    n.font.size = Pt(12)
    h1 = tpl.styles["Heading 1"]
    h1.font.name = "Times New Roman"
    h1.font.size = Pt(14)
    h1.font.bold = True
    ph = tpl.add_paragraph("Introduction", style="Heading 1")
    ph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    tpl_path = str(tmp_path / "tpl.docx")
    tpl.save(tpl_path)

    paper = _paper(tmp_path)
    t = build_target(spec_key=None, template=tpl_path)
    prof, findings = run_check(paper, t)

    # 检查结论里不许出现「L1 应当是 14pt 左对齐而实际 16pt 居中」这种张冠李戴
    # ⚠️ v2.3.75 起只筛「字号 / 对齐」两项：新加的段前段后 / 行距维度会合法地报出
    # 「模板章标题段前 24pt、论文是 0pt」——那不是张冠李戴，正是模板驱动该说的话，
    # 用 `assert not bad` 一刀切会把这个正确结论误判成本用例失败。
    bad = [f for f in findings
           if f.dimension == "heading" and f.severity != "pass"
           and "heading format" in f.message.lower()
           and ("16 pt" in (f.actual or "") or "center" in (f.actual or "").lower())]
    assert not bad, [(f.message, f.expected, f.actual) for f in bad]

    # 并且报告里必须有一条「认出论文大标题、没参与层级比对」的说明（客户才知道没漏检）
    dt = [f for f in findings if f.dimension == "doc_title"]
    assert dt and dt[0].severity == "info"
    assert "16pt" in dt[0].actual

    out = str(tmp_path / "fixed.docx")
    fix_docx(paper, out, t, findings=findings)
    doc = Document(out)
    title = [p for p in doc.paragraphs if p.text.strip() == TITLE][0]
    assert title.runs[0].font.size == Pt(16), "论文大标题被改小了"


# ---------------------------------------------------------------------------
# 五、问卷 schema / i18n：参考文献不许被定死成三选一
# ---------------------------------------------------------------------------

def test_reference_style_offers_custom_and_defaults_to_it():
    field = [f for f in QUESTIONNAIRE_SCHEMA if f["key"] == "reference_style"][0]
    assert field["default"] == REFERENCE_STYLE_CUSTOM
    assert REFERENCE_STYLE_CUSTOM in field["options"]
    for k in REFERENCE_STYLE_KEYS:
        assert k in field["options"]


def test_reference_style_option_labels_are_bilingual():
    from src.ui import i18n
    zh = i18n.questionnaire_option("zh", "reference_style", REFERENCE_STYLE_CUSTOM)
    en = i18n.questionnaire_option("en", "reference_style", REFERENCE_STYLE_CUSTOM)
    assert "自定义" in zh
    assert "custom" in en.lower()
    # 未登记的选项原样显示，旧数据不受影响
    assert i18n.questionnaire_option("en", "reference_style", "IEEE") == "IEEE"


def test_custom_reference_style_parses_to_none():
    t = target_from_dict({"reference_style": REFERENCE_STYLE_CUSTOM,
                          "font_size_pt": 12}, "user")
    assert t.reference_style is None
    assert t.spec_key is None
    # 选了真规范：参考文献仍取规范值（悬挂缩进/是否编号），但 spec_key 仍不许升级
    t2 = target_from_dict({"reference_style": "APA"}, "user")
    assert t2.reference_style == "APA"
    assert t2.reference_hanging_indent_in == get_spec("APA").reference.hanging_indent_in
    assert t2.spec_key is None


# ---------------------------------------------------------------------------
# 六、深层级：只有「声明了完整多级」才允许说「按最深一级处理」
# ---------------------------------------------------------------------------

def _doc_two_levels(tmp_path):
    """一二级标题都用样式（真实学校模板常见形态）。"""
    p = tmp_path / "two.docx"
    doc = Document()
    doc.add_paragraph("Chapter One", style="Heading 1")
    doc.add_paragraph("Section One", style="Heading 2")
    doc.add_paragraph("Body text.")
    doc.save(p)
    return str(p)


def test_deeper_level_fallback_is_only_announced_for_multi_level_targets(tmp_path):
    """学校模板只声明 1 级标题格式 + 文档有 2 级 → 检查**不许**说「二级按一级处理」。

    因为 fixer 根本不会拿一级格式套二级（v2.3.72 明令禁止，实测整篇层级崩掉），
    检查若这么写就是假话。这种情形由「层级数超限」那条结论负责说明。
    """
    tpl = Document()
    tpl.styles["Heading 1"].font.size = Pt(14)
    tpl.styles["Heading 1"].font.bold = True
    tpl_path = str(tmp_path / "tpl1.docx")
    tpl.save(tpl_path)

    paper = _doc_two_levels(tmp_path)
    t = build_target(spec_key=None, template=tpl_path)   # 模板只提供 1 级
    _, findings = run_check(paper, t)
    fallback = [f for f in findings if "Falls back" in (f.expected or "")]
    assert not fallback, [(f.message, f.expected) for f in fallback]


def test_deeper_level_fallback_is_announced_for_spec_baseline(tmp_path):
    """规范基线 APA 声明 5 级、文档有第 6 级 → 检查要明说「按第 5 级处理」。"""
    p = tmp_path / "six.docx"
    doc = Document()
    doc.add_paragraph("Deep Heading", style="Heading 6")
    doc.add_paragraph("Body one.")
    doc.add_paragraph("Body two.")
    doc.save(p)
    t = build_target(spec_key="APA-Student")
    _, findings = run_check(str(p), t)
    fallback = [f for f in findings if "Falls back to level 5" in (f.expected or "")]
    assert fallback, [f.message for f in findings]
