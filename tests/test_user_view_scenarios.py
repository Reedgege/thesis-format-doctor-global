"""客户视角护栏（v2.3.71）：四个真实场景里被客户骂过的行为，逐条钉死。

来源不是「代码看起来对不对」，而是把客户真实动作走了一遍（上传学校模板 → 检查论文 →
点修正 → 打开改完的文档）之后发现的四类事故：

1. **行距从来没被检查过**：行距写在 Normal 样式上是 Word 的常态，旧采样只读段落直接格式
   → 行距恒为 None → 单倍行距的论文交上来报告全绿（假通过）。行距是论文格式头号要求。
2. **干净文档点「修正」被乱改**：预览一行都没列，文档里 8 个标题却被重排、还插进 8 条批注。
3. **手排章节标题改了等于没改**：学校模板模式下论文大标题被当成伪标题占掉 L1，学生手排的
   章节标题被推成 L2 → 模板只声明 1 级 → 检查不报、修正跳过（报告却写"已修正"）。
4. **报告里是内部话术**："Custom/unregistered check items"、Expected: Other —— 客户看不懂。

造数据一律用真实形态：手排标题（Normal 样式 + 整段加粗）、行距挂在 Normal 样式上、
混合文档（部分标题用了样式、部分手打）—— 这正是学生论文最常见的样子。
"""
from __future__ import annotations

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

from src.engine.checker import run_check
from src.engine.docx_reader import read_docx
from src.engine.fixer import _compute_changes, _detect_title_page, fix_docx
from src.engine.questionnaire import TargetProfile, target_from_template_docx
from src.engine.report import build_report
from src.engine.specs import get_spec

TITLE = "The Impact of Artificial Intelligence on Academic Writing"
SECTIONS = ("Introduction", "Method", "Results", "Conclusion")


def _doc_with_style_spacing(tmp_path, name="paper.docx", *, body_font="Calibri",
                            body_size=10.5, spacing=1.0, manual_headings=True,
                            heading1=None, add_abstract=True):
    """造一份「学生真实形态」的论文。

    - 行距写在 **Normal 样式** 上（Word 里设行距的常态，段落无直接格式）；
    - 标题手排（Normal + 整段加粗 + 14pt）或真 Heading 样式（heading1 给定时）；
    - 首段是 16pt 居中的论文大标题（模板模式下曾被误当章节标题）。
    """
    p = tmp_path / name
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = body_font
    normal.font.size = Pt(body_size)
    normal.paragraph_format.line_spacing = spacing
    sec = doc.sections[0]                       # 页边距也要设成模板值，否则「干净文档」
    sec.left_margin = Inches(1.0)               # 光有对齐/字体对，页边距仍是 python-docx
    sec.right_margin = Inches(1.0)              # 默认的 1.25″ → 预览里会冒出页边距改动
    sec.top_margin = Inches(1.0)
    sec.bottom_margin = Inches(1.0)
    if heading1:
        for st, size in heading1.items():
            s = doc.styles[st]
            s.font.name = body_font
            s.font.size = Pt(size)
            s.font.bold = True

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(TITLE)
    r.bold = True
    r.font.size = Pt(16)

    if add_abstract:
        doc.add_paragraph().add_run("Abstract").bold = True
        doc.add_paragraph("This study examines how AI writing tools reshape "
                          "academic authorship and citation practices.")

    for head in SECTIONS:
        if manual_headings:
            ph = doc.add_paragraph()          # 手排标题：Normal 样式 + 加粗 + 14pt
            rr = ph.add_run(head)
            rr.bold = True
            rr.font.size = Pt(14)
        else:
            doc.add_heading(head, level=1)
        doc.add_paragraph("Body text for %s. It has a couple of sentences." % head)

    # 参考文献段（真实论文都有；没选引用风格时报告要说清「参考文献没被检查」）
    doc.add_paragraph().add_run("References").bold = True
    for e in ("Smith, J. A. (2021). The automation of authorship. "
              "Journal of Writing Research, 12(3), 45-67.",
              "Lee, K., & Patel, R. (2020). Citation practices in the age of machines. "
              "Educational Technology Review, 8(2), 112-130."):
        doc.add_paragraph(e)
    doc.save(p)
    return p


def _template_target(*, font="Times New Roman", size=12.0, spacing=2.0,
                     h1_size=14.0, ref_style=None):
    """模拟「客户上传学校模板」得到的目标（模板模式 spec_key=None）。"""
    return TargetProfile(
        spec_key=None,
        page={"font_family": font, "font_size_pt": size, "line_spacing": spacing,
              "margin_top_in": 1.0, "margin_bottom_in": 1.0,
              "margin_left_in": 1.0, "margin_right_in": 1.0},
        heading_levels=1,
        heading_formats={1: {"font_family": font, "size_pt": h1_size, "bold": True}},
        reference_style=ref_style,
        resolution={"font_family": "school", "line_spacing": "school",
                    "font_size_pt": "school", "heading_formats": "school"},
    )


# ---------------------------------------------------------------- 1. 行距必须查得出来
def test_line_spacing_inherited_from_normal_style_is_read(tmp_path):
    """行距挂在 Normal 样式上（Word 常态）也必须读得到，否则行距维度整个失效。"""
    p = _doc_with_style_spacing(tmp_path, spacing=2.0)
    prof = read_docx(str(p))
    assert prof.line_spacing == 2.0, "行距由 Normal 样式承载时读不到 → 检查维度失效"


def test_line_spacing_mismatch_is_reported_and_fixed(tmp_path):
    """客户交单倍行距论文、模板要求双倍 → 必须报出来，且修完真的变双倍。"""
    src = _doc_with_style_spacing(tmp_path, spacing=1.0)
    target = _template_target(spacing=2.0)
    _, findings = run_check(str(src), target)
    ls = [f for f in findings if f.dimension == "line_spacing" and f.severity != "pass"]
    assert ls, [(f.dimension, f.severity) for f in findings]

    out = tmp_path / "out.docx"
    fix_docx(str(src), str(out), target, findings=findings)
    doc = Document(str(out))
    body = [p for p in doc.paragraphs if p.text.startswith("Body text")][0]
    assert body.paragraph_format.line_spacing == 2.0, "修完正文行距必须真的是双倍"


def test_fixed_line_spacing_passes_recheck(tmp_path):
    """修完复检行距必须变绿（否则客户会觉得「改了也没用」）。"""
    src = _doc_with_style_spacing(tmp_path, spacing=1.0)
    target = _template_target(spacing=2.0)
    _, findings = run_check(str(src), target)
    out = tmp_path / "out2.docx"
    fix_docx(str(src), str(out), target, findings=findings)
    _, after = run_check(str(out), target)
    assert not [f for f in after if f.dimension == "line_spacing" and f.severity != "pass"]


# ---------------------------------------------------------------- 2. 干净文档不许被乱改
def test_clean_paper_preview_and_fix_do_nothing(tmp_path):
    """客户拿一份已经合规的论文点「修正」：预览为空、文档一个字都不许动。

    旧行为（v2.3.70 及以前）：预览 0 条，但 fixer 仍把 8 个标题重排 + 插 8 条批注 ——
    「预览说没改动、实际改了文档」，是最伤信任的一类 bug。
    """
    src = _doc_with_style_spacing(tmp_path, name="clean.docx", body_font="Times New Roman",
                                  body_size=12.0, spacing=2.0, manual_headings=False,
                                  heading1={"Heading 1": 14.0})
    target = _template_target(font="Times New Roman", size=12.0, spacing=2.0, h1_size=14.0)
    prof, findings = run_check(str(src), target)

    changes = _compute_changes(prof, target, "en", findings=findings)
    assert changes == [], "文档已合规，预览必须为空：%r" % (changes,)

    out = tmp_path / "clean_out.docx"
    summary = fix_docx(str(src), str(out), target, findings=findings)
    assert summary.get("headings_formatted", 0) == 0, "没报标题问题就不许重排标题"
    assert summary.get("comments_added", 0) == 0, "没报标题问题就不许往文档里插批注"


def test_dirty_paper_still_gets_fixed(tmp_path):
    """对照：真有问题的论文，标题照样要被修（别把门控做成一刀切不干活）。"""
    src = _doc_with_style_spacing(tmp_path, name="dirty.docx", body_font="Calibri",
                                  body_size=10.5, spacing=1.0, manual_headings=True,
                                  heading1={"Heading 1": 13.0})
    target = _template_target(font="Times New Roman", size=12.0, spacing=2.0, h1_size=14.0)
    prof, findings = run_check(str(src), target)
    assert _compute_changes(prof, target, "en", findings=findings), "有问题的论文预览不能为空"
    out = tmp_path / "dirty_out.docx"
    fix_docx(str(src), str(out), target, findings=findings)
    doc = Document(str(out))
    intro = [p for p in doc.paragraphs if p.text.strip() == "Introduction"][0]
    assert intro.runs[0].font.name == "Times New Roman", "手排章节标题必须被改成模板字体"
    assert intro.runs[0].font.size.pt == 14.0


# ---------------------------------------------------------------- 3. 手排章节标题必须被认出来
def test_manual_section_headings_are_recognised_not_skipped(tmp_path):
    """模板模式（spec_key=None）下，手排章节标题不能被论文大标题挤到二级而逃过修正。"""
    src = _doc_with_style_spacing(tmp_path, manual_headings=True)
    target = _template_target()
    doc = Document(str(src))
    tp = _detect_title_page(doc)
    assert tp is not None and tp["title"] == 0, "论文大标题必须被认成标题页标题（否则挤掉 L1 档）"

    _, findings = run_check(str(src), target)
    hd = [f for f in findings if f.dimension == "heading" and f.severity != "pass"]
    assert hd, "手排章节标题与模板不符，必须报出来：%r" % [
        (f.message, f.expected, f.actual) for f in findings if f.dimension == "heading"]


def test_template_mode_uses_template_headings(tmp_path):
    """端到端：模板只声明 1 级 → 6 个手排章节标题全部按模板 L1 修（不是只修 References）。"""
    src = _doc_with_style_spacing(tmp_path, manual_headings=True)
    target = _template_target(font="Times New Roman", h1_size=14.0)
    _, findings = run_check(str(src), target)
    out = tmp_path / "tpl_out.docx"
    fix_docx(str(src), str(out), target, findings=findings)
    doc = Document(str(out))
    for head in SECTIONS:
        p = [x for x in doc.paragraphs if x.text.strip() == head][0]
        assert p.runs[0].font.name == "Times New Roman", "%s 没被修" % head


# ---------------------------------------------------------------- 4. 报告必须说人话
def test_report_has_no_internal_jargon(tmp_path):
    """没选引用风格时，报告不许出现 "Custom/unregistered" 或 Expected: Other。"""
    src = _doc_with_style_spacing(tmp_path)
    target = _template_target(ref_style=None)
    prof, findings = run_check(str(src), target)
    rep = build_report(str(src), target, prof, findings)
    md = rep.markdown
    assert "Custom/unregistered" not in md
    assert "Expected: Other" not in md
    assert "Reference style not specified" in md, md
    # 所有 finding 的 expected 都不许是内部键名
    for f in findings:
        assert (f.expected or "").strip() != "Other", (f.dimension, f.message)
