# -*- coding: utf-8 -*-
"""学校模板驱动：题目字号/字体 + 摘要正文字号/加粗 的回归测试。

铁律：学校模板批注 + 模板自带样式是唯一事实源，严禁套内置规范硬编码值。
本测试覆盖：
- 模板批注「Title: 24 pt」→ 画像/检查/修正（题目变 24pt TNR，不套规范默认）。
- 模板批注「Abstract body: 12 pt」（未提 bold）→ 摘要正文 12pt 不加粗
  （不套 APA 规范的 9pt 加粗，即复现 2.3.113 的 bug 并验证已修）。
- 参考文献风格由 build_target 的 reference_style 参数单独驱动（IEEE），
  与页面/摘要版式解耦（学校模板写 APA 也不影响 IEEE 选择）。
"""
from __future__ import annotations

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from src.engine import checker, fixer
from src.engine.docx_reader import read_docx, doc_title_index
from src.engine.questionnaire import build_target

from tools.make_sample_template_en import make as make_sample_template


def _make_student_paper(path):
    """构造一篇「学校模板值全错」的学生稿：
    - 题目 16pt Calibri（应为 24pt TNR）
    - 摘要正文 9pt 加粗（应为 12pt 不加粗）
    - 参考文献用 IEEE 编号格式。
    """
    doc = Document()
    # 题目（标题页第一段，不套标题样式）
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run("Digital Transformation and Supply Chain Resilience")
    tr.font.name = "Calibri"
    tr.font.size = Pt(16)
    tr.bold = True
    # 作者行
    ap = doc.add_paragraph()
    ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ar = ap.add_run("John Doe")
    ar.font.name = "Calibri"
    ar.font.size = Pt(11)
    # 摘要标题（样式级标题，命中 checker 主检测路径）
    ah = doc.add_paragraph("Abstract", style="Heading 2")
    # 摘要正文（错：9pt 加粗）
    ab = doc.add_paragraph()
    abr = ab.add_run(
        "This paper studies how digital technologies reshape supply chain resilience "
        "under external shocks.")
    abr.font.name = "Times New Roman"
    abr.font.size = Pt(9)
    abr.bold = True
    # 关键词
    kw = doc.add_paragraph()
    kwr = kw.add_run("Keywords: digital transformation, supply chain, resilience")
    kwr.font.name = "Times New Roman"
    kwr.font.size = Pt(12)
    # 正文
    doc.add_paragraph("Chapter One", style="Heading 1")
    bp = doc.add_paragraph()
    bpr = bp.add_run("Body text with some content to inspect formatting.")
    bpr.font.name = "Times New Roman"
    bpr.font.size = Pt(12)
    # 参考文献（IEEE）
    doc.add_paragraph("References", style="Heading 1")
    rp = doc.add_paragraph(
        '[1] J. A. Smith, "Title of the work," J. Name, vol. 1, no. 2, pp. 3-4, 2020.')
    rp.paragraph_format.left_indent = Inches(0.5)
    rp.paragraph_format.first_line_indent = Inches(-0.5)
    doc.save(path)
    return path


def _title_runs(path):
    doc = Document(path)
    p0 = doc.paragraphs[0]
    return [{"sz": (r.font.size.pt if r.font.size else None),
             "bold": r.bold,
             "font": r.font.name} for r in p0.runs]


def _abstract_body_run(path):
    """返回摘要正文第一个段落的首 run 属性。"""
    doc = Document(path)
    paras = doc.paragraphs
    head_idx = next((i for i, p in enumerate(paras)
                    if p.text.strip().lower() == "abstract"
                    and p.style.name.startswith("Heading")), None)
    assert head_idx is not None, "测试夹具未命中摘要标题"
    for j in range(head_idx + 1, len(paras)):
        if paras[j].text.strip():
            r = paras[j].runs[0]
            return {"sz": (r.font.size.pt if r.font.size else None),
                    "bold": r.bold,
                    "font": r.font.name}
    raise AssertionError("未找到摘要正文段")


def test_template_drives_title_and_abstract(tmp_path):
    tpl, _ = make_sample_template(str(tmp_path / "tpl.docx"))
    paper = _make_student_paper(str(tmp_path / "paper.docx"))

    # 用户：学校模板 + 参考文献自己选 IEEE（只作用于参考文献）
    target = build_target(template=tpl, reference_style="IEEE")

    # 画像必须来自模板，不套规范硬编码
    assert target.title_page_title_size_pt == 24.0
    assert target.title_page_title_font == "Times New Roman"
    assert target.abstract_font_size_pt == 12.0
    assert target.abstract_bold is False
    assert target.reference_style == "IEEE"

    # 修正前：题目字号错、摘要正文被加粗 → 应有 title_page / abstract 维度告警
    prof, findings = checker.run_check(paper, target)
    dims = {f.dimension for f in findings if f.dimension in ("title_page", "abstract")}
    assert "title_page" in dims, "应检出题目页格式不达标"
    assert "abstract" in dims, "应检出摘要格式不达标（含误加粗）"

    # 修正前实测：题目 16pt / 摘要正文 9pt 加粗
    tr0 = _title_runs(paper)[0]
    assert tr0["sz"] == 16.0
    ab0 = _abstract_body_run(paper)
    assert ab0["sz"] == 9.0 and ab0["bold"] is True

    # 修正
    out = str(tmp_path / "paper_fixed.docx")
    fixer.fix_docx(paper, out, target)

    # 修正后实测：题目 24pt TNR 加粗；摘要正文 12pt 不加粗
    tr1 = _title_runs(out)[0]
    assert tr1["sz"] == 24.0, f"题目应修正为 24pt，实测 {tr1['sz']}"
    assert tr1["font"] == "Times New Roman", f"题目字体应修正为 TNR，实测 {tr1['font']}"
    assert tr1["bold"] is True
    ab1 = _abstract_body_run(out)
    assert ab1["sz"] == 12.0, f"摘要正文应修正为 12pt，实测 {ab1['sz']}"
    assert ab1["bold"] is False, "摘要正文不应被加粗"

    # 修正后复检：title_page / abstract 维度应全过
    prof2, findings2 = checker.run_check(out, target)
    bad = [f for f in findings2 if f.dimension in ("title_page", "abstract")
           and f.severity in ("fail", "warn")]
    assert not bad, f"修正后不应残留 title_page/abstract 告警：{bad}"


def test_ieee_reference_untouched_by_page_fix(tmp_path):
    """页面/摘要修正不得误伤用户自选的 IEEE 参考文献条目。"""
    tpl, _ = make_sample_template(str(tmp_path / "tpl.docx"))
    paper = _make_student_paper(str(tmp_path / "paper.docx"))
    target = build_target(template=tpl, reference_style="IEEE")
    out = str(tmp_path / "paper_fixed.docx")
    fixer.fix_docx(paper, out, target)

    doc = Document(out)
    ref_para = next(p for p in doc.paragraphs
                    if p.text.strip().startswith("[1]"))
    # IEEE 编号条目首字符保持 [1]
    assert ref_para.text.strip().startswith("[1]"), "IEEE 参考文献条目不应被改写"
    # 悬挂缩进应保持
    assert ref_para.paragraph_format.left_indent is not None


def _make_shrunk_title_paper(path):
    """构造「题目被错改小到与正文同号」的论文，复现 2.3.113 的题目字号读取盲区：
    - 题目 12pt（run）居中加粗，与正文 12pt（run）相同 → doc_title_index 因
      「题目不是全篇最大字号（须严格大于之后每段）」判不出 → 题目字号读不到。
    - 标题用样式字号（run 无显式字号），不参与 run 字号比较，避免干扰触发条件。
    """
    doc = Document()
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run("Paper Title Shrunk To Body Size")
    tr.font.name = "Times New Roman"
    tr.font.size = Pt(12)
    tr.bold = True
    # 作者行
    ap = doc.add_paragraph()
    ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ap.add_run("John Doe").font.size = Pt(11)
    # 摘要标题（样式级标题，命中 detect_title_page 主检测路径）
    doc.add_paragraph("Abstract", style="Heading 2")
    # 摘要正文 12pt（与题目 run 字号相同，制造盲区）
    ab = doc.add_paragraph()
    ab.add_run("This paper studies a topic of interest.").font.size = Pt(12)
    doc.add_paragraph("Keywords: a, b")
    # 正文 + 样式级标题（run 无显式字号，不参与 run 比较）
    doc.add_paragraph("Chapter One", style="Heading 1")
    bp = doc.add_paragraph()
    bp.add_run("Body text at twelve point, same as the title run size.").font.size = Pt(12)
    doc.add_paragraph("References", style="Heading 1")
    rp = doc.add_paragraph(
        '[1] J. A. Smith, "Title of the work," J. Name, vol. 1, no. 2, pp. 3-4, 2020.')
    rp.paragraph_format.left_indent = Inches(0.5)
    rp.paragraph_format.first_line_indent = Inches(-0.5)
    doc.save(path)
    return path




def test_title_size_readable_when_shrunk_to_body_size(tmp_path):
    """加固回归：题目字号读取盲区。

    doc_title_index 因「题目非全篇最大字号」判不出（题目 12pt == 正文 12pt）时，
    题目字号必须仍可经 detect_title_page 兜底读到，否则 title_page 维度会误报 pass、
    客户看不到题目字号错。prof.doc_title_index 保持 None（章节标题层级排除不动）。
    """
    import pytest
    from docx import Document as _D
    paper = _make_shrunk_title_paper(str(tmp_path / "shrunk.docx"))
    # 主路径 doc_title_index 因「题目非严格大于之后每段」判不出
    assert doc_title_index(_D(paper)) is None
    prof = read_docx(paper)
    # doc_title_index 不动（章节标题层级排除零影响）
    assert prof.doc_title_index is None
    # 加固：兜底仍能读到题目字号
    assert prof.doc_title_size_pt == pytest.approx(12.0)


def test_title_size_flagged_when_shrunk_below_max(tmp_path):
    """加固回归（端到端）：题目被错改小到与正文同号时，检查器仍能标红题目字号，
    不再误报 pass。模板题目 24pt，实测 12pt → title_page 维度应 fail/warn 且说明含
    「12 pt (expected 24 pt)」。"""
    tpl, _ = make_sample_template(str(tmp_path / "tpl.docx"))
    paper = _make_shrunk_title_paper(str(tmp_path / "shrunk.docx"))
    target = build_target(template=tpl, reference_style="IEEE")
    prof, findings = checker.run_check(paper, target)
    f = next((x for x in findings if x.dimension == "title_page"), None)
    assert f is not None, "应有 title_page 维度"
    assert f.severity in ("fail", "warn"), "题目字号错应被标红，而非误报 pass"
    assert "12 pt" in (f.actual or "") and "24 pt" in (f.actual or "")


# v2.3.114：三路齐平回归。模板 / 手填问卷 / AI 导入 三个通道必须都能把
# 「题目字号/字体 + 摘要正文字号/加粗/对齐」这 5 个字段带进画像，并经同一套
# 检查/修正路径生效。模板通道已由 test_template_drives_title_and_abstract 覆盖，
# 下面两条分别覆盖手填问卷、AI 导入两个通道。
_MANUAL_Q = {
    "title_page_title_size_pt": 24.0,
    "title_page_title_font": "Times New Roman",
    "abstract_font_size_pt": 12.0,
    "abstract_bold": False,          # 三态：False = 必须不加粗
    "abstract_alignment": "left",
    "reference_style": "IEEE",       # 只作用于参考文献，与页面/摘要版式解耦
}


def _ai_yaml_path(tmp_path):
    """写一份 AI 填表 YAML（扁平键值，与手填问卷同一组 schema key）。"""
    import yaml
    p = tmp_path / "ai_filled.yaml"
    p.write_text(
        "title_page_title_size_pt: 24.0\n"
        "title_page_title_font: Times New Roman\n"
        "abstract_font_size_pt: 12.0\n"
        "abstract_bold: false\n"
        "abstract_alignment: left\n"
        "reference_style: IEEE\n",
        encoding="utf-8",
    )
    # 落盘后回读一次，确认 YAML 解析路径与手写一致
    with open(p, "r", encoding="utf-8") as fh:
        assert yaml.safe_load(fh.read())["abstract_bold"] is False
    return str(p)


def test_manual_questionnaire_drives_title_and_abstract(tmp_path):
    """手填问卷通道：5 字段经 target_from_dict 进画像，检查/修正与模板通道一致。"""
    paper = _make_student_paper(str(tmp_path / "paper.docx"))
    target = build_target(questionnaire=_MANUAL_Q)

    # 画像必须来自手填值，不套规范硬编码
    assert target.title_page_title_size_pt == 24.0
    assert target.title_page_title_font == "Times New Roman"
    assert target.abstract_font_size_pt == 12.0
    assert target.abstract_bold is False
    assert target.abstract_alignment == "left"
    assert target.reference_style == "IEEE"

    # 修正前：题目 16pt Calibri / 摘要正文 9pt 加粗 → 应被标红
    prof, findings = checker.run_check(paper, target)
    dims = {f.dimension for f in findings if f.dimension in ("title_page", "abstract")}
    assert "title_page" in dims, "手填问卷应检出题目页格式不达标"
    assert "abstract" in dims, "手填问卷应检出摘要格式不达标（含误加粗）"

    # 修正
    out = str(tmp_path / "paper_fixed.docx")
    fixer.fix_docx(paper, out, target)

    # 修正后实测：题目 24pt TNR 加粗；摘要正文 12pt 不加粗
    tr1 = _title_runs(out)[0]
    assert tr1["sz"] == 24.0, f"题目应修正为 24pt，实测 {tr1['sz']}"
    assert tr1["font"] == "Times New Roman", f"题目字体应修正为 TNR，实测 {tr1['font']}"
    assert tr1["bold"] is True
    ab1 = _abstract_body_run(out)
    assert ab1["sz"] == 12.0, f"摘要正文应修正为 12pt，实测 {ab1['sz']}"
    assert ab1["bold"] is False, "摘要正文不应被加粗"

    # 修正后复检：title_page / abstract 维度应全过
    prof2, findings2 = checker.run_check(out, target)
    bad = [f for f in findings2 if f.dimension in ("title_page", "abstract")
           and f.severity in ("fail", "warn")]
    assert not bad, f"修正后不应残留 title_page/abstract 告警：{bad}"


def test_ai_import_drives_title_and_abstract(tmp_path):
    """AI 导入通道：5 字段经 target_from_ai（YAML → target_from_dict）进画像，
    检查/修正与模板通道一致。"""
    paper = _make_student_paper(str(tmp_path / "paper.docx"))
    ai_path = _ai_yaml_path(tmp_path)
    target = build_target(ai=ai_path)

    # 画像必须来自 AI 填表值，不套规范硬编码
    assert target.title_page_title_size_pt == 24.0
    assert target.title_page_title_font == "Times New Roman"
    assert target.abstract_font_size_pt == 12.0
    assert target.abstract_bold is False
    assert target.abstract_alignment == "left"
    assert target.reference_style == "IEEE"

    # 修正前：题目 16pt Calibri / 摘要正文 9pt 加粗 → 应被标红
    prof, findings = checker.run_check(paper, target)
    dims = {f.dimension for f in findings if f.dimension in ("title_page", "abstract")}
    assert "title_page" in dims, "AI 导入应检出题目页格式不达标"
    assert "abstract" in dims, "AI 导入应检出摘要格式不达标（含误加粗）"

    # 修正
    out = str(tmp_path / "paper_fixed.docx")
    fixer.fix_docx(paper, out, target)

    # 修正后实测：题目 24pt TNR 加粗；摘要正文 12pt 不加粗
    tr1 = _title_runs(out)[0]
    assert tr1["sz"] == 24.0, f"题目应修正为 24pt，实测 {tr1['sz']}"
    assert tr1["font"] == "Times New Roman", f"题目字体应修正为 TNR，实测 {tr1['font']}"
    assert tr1["bold"] is True
    ab1 = _abstract_body_run(out)
    assert ab1["sz"] == 12.0, f"摘要正文应修正为 12pt，实测 {ab1['sz']}"
    assert ab1["bold"] is False, "摘要正文不应被加粗"

    # 修正后复检：title_page / abstract 维度应全过
    prof2, findings2 = checker.run_check(out, target)
    bad = [f for f in findings2 if f.dimension in ("title_page", "abstract")
           and f.severity in ("fail", "warn")]
    assert not bad, f"修正后不应残留 title_page/abstract 告警：{bad}"
