# -*- coding: utf-8 -*-
"""首页结构三件套（v2.3.91）：题名页独立成页 / 标题垂直位置 + 零额外段距 / 正文另起一页。

老板 2026-10-22 拿 v2.3.90 的产物逐页核对后拍板「三件一起做，清掉」：
  ① 题目贴在顶边距第一行（APA 要求标题在页边距下 3–4 行）；
  ② 题名页与摘要挤在同一页（题名页 / 摘要 / 正文应各自成页）；
  ③ 原文残留的 12pt / 18pt 段距没清（APA 题名页整页双倍行距、不留额外段距）。

三条共同契约（本文件逐条钉住）：
  * **纯版式** —— 只写 ``w:pageBreakBefore`` 与 ``w:spacing``，段落数与文字逐字不变；
  * **规范基线才动手** —— 学校模板 / 手填 / AI 导入一律只提示（模板说了算）；
  * **修完必须收敛** —— 复检不许再报同一条，二遍修正不许再改一个字节。
"""

from __future__ import annotations

import os
import re
import sys
import zipfile

import pytest

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine import checker
from src.engine.docx_reader import read_docx
from src.engine.fixer import fix_docx, preview_fix
from src.engine.questionnaire import target_from_dict, target_from_spec

NEW_DIMS = ("title_page_new_page", "title_page_layout", "text_new_page")

TITLE = "Digital Transformation and Supply Chain Resilience"


# ---------------------------------------------------------------------------
# 夹具：贴近真实样本（标题 / 作者单位块 / 摘要 / 正文 / 参考文献）
# ---------------------------------------------------------------------------

def _make_paper(path, *, title_break=False, text_break=False, extra_spacing=True,
                title_space_before_pt=None, title_space_after_pt=None,
                blank_after_title=False, title=TITLE,
                course=None, instructor=None, due_date=None):
    """造一篇题名页结构完整的论文。

    ``title_break`` / ``text_break`` 控制客户**原本**有没有那两个分页；
    ``extra_spacing`` 控制原文是否残留 12pt/18pt 段距（老板实测的那份就有）；
    ``title_space_before_pt`` 直接指定标题段前（None = 不写，等同 0）；
    ``blank_after_title`` 用一个真空段落实现 APA 要求的那一个空行。
    """
    doc = Document()
    t = doc.add_paragraph()
    r = t.add_run(title)
    r.bold = True
    r.font.size = Pt(16)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if title_space_before_pt is not None:
        t.paragraph_format.space_before = Pt(title_space_before_pt)
    if title_space_after_pt is not None:
        t.paragraph_format.space_after = Pt(title_space_after_pt)
    elif extra_spacing:
        t.paragraph_format.space_after = Pt(18)

    if blank_after_title:
        doc.add_paragraph("")

    a1 = doc.add_paragraph("By: Alex Morgan")
    a1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    a2 = doc.add_paragraph("School of Business and Management")
    a2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if extra_spacing:
        a1.paragraph_format.space_after = Pt(12)
        a2.paragraph_format.space_after = Pt(12)

    # APA 学生版题名页的可选补充行（课程 / 教师 / 日期）：给不给由调用方决定
    for _line in (course, instructor, due_date):
        if _line:
            _p = doc.add_paragraph(_line)
            _p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # APA 学生版题名页的可选补充行（课程 / 教师 / 日期）：给不给由调用方决定
    for _line in (course, instructor, due_date):
        if _line:
            _p = doc.add_paragraph(_line)
            _p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    h = doc.add_paragraph("Abstract")
    h.style = doc.styles["Heading 1"]
    if title_break:
        h.paragraph_format.page_break_before = True
    doc.add_paragraph("This paper studies how digital transformation affects resilience.")
    doc.add_paragraph("Keywords: digital transformation, supply chain, resilience")

    b = doc.add_paragraph("1. Introduction")
    b.style = doc.styles["Heading 1"]
    if text_break:
        b.paragraph_format.page_break_before = True
    for i in range(4):
        doc.add_paragraph("Body paragraph %d about the research methods." % (i + 1))

    ref = doc.add_paragraph("References")
    ref.style = doc.styles["Heading 1"]
    for i in range(2):
        doc.add_paragraph("Ivanov, D. (2020). Predicting the impacts. Entry %d." % (i + 1))
    doc.save(str(path))
    return str(path)


def _dim(findings, dim):
    return [f for f in findings if f.dimension == dim]


def _sev(findings, dim):
    got = _dim(findings, dim)
    return got[0].severity if got else None


def _texts(path):
    return [p.text for p in Document(str(path)).paragraphs]


# ---------------------------------------------------------------------------
# ① 题名页独立成页
# ---------------------------------------------------------------------------

def test_title_page_new_page_warn_and_fixable_on_baseline(tmp_path):
    """规范基线 + 题名页与摘要挤在一页 → warn 且可修；修完复检必须收敛。"""
    p = _make_paper(tmp_path / "a.docx", title_break=False)
    t = target_from_spec("APA-Student")
    prof, findings = checker.run_check(p, t)
    assert prof.tp_next_break_before is False
    assert _sev(findings, "title_page_new_page") == "warn"
    assert any(getattr(f, "fix", None) for f in _dim(findings, "title_page_new_page"))

    out = tmp_path / "out.docx"
    summary = fix_docx(p, str(out), t)
    assert summary["applied"]["title_page_new_page"]["changed"] is True

    prof2, findings2 = checker.run_check(str(out), t)
    assert prof2.tp_next_break_before is True
    assert _sev(findings2, "title_page_new_page") == "pass"


def test_title_page_new_page_stays_info_when_not_baseline(tmp_path):
    """学校模板模式（非规范基线）即便引用了 APA，也只提示、不绑定修正。"""
    p = _make_paper(tmp_path / "a.docx", title_break=False)
    t = target_from_dict({"reference_style": "APA"}, "school")
    prof, findings = checker.run_check(p, t)
    assert _sev(findings, "title_page_new_page") == "info"
    assert all(getattr(f, "fix", None) is None for f in _dim(findings, "title_page_new_page"))


def test_title_page_already_on_own_page_is_pass(tmp_path):
    """客户本来就有分页 → pass，且一个字节都不动（绝不重写客户自己的分页符）。"""
    p = _make_paper(tmp_path / "a.docx", title_break=True)
    t = target_from_spec("APA-Student")
    prof, findings = checker.run_check(p, t)
    assert prof.tp_next_break_before is True
    assert _sev(findings, "title_page_new_page") == "pass"

    out = tmp_path / "out.docx"
    summary = fix_docx(p, str(out), t)
    assert "title_page_new_page" not in summary["applied"]


# ---------------------------------------------------------------------------
# ② 标题垂直位置 + 零额外段距
# ---------------------------------------------------------------------------

def test_title_sits_three_lines_below_top_margin_after_fix(tmp_path):
    """APA：标题应位于页边距下 3 行 = 3 × 12pt × 2.0 = 72pt。"""
    p = _make_paper(tmp_path / "a.docx")
    t = target_from_spec("APA-Student")
    out = tmp_path / "out.docx"
    summary = fix_docx(p, str(out), t)

    applied = summary["applied"]["title_page_layout"]
    assert applied["lines"] == pytest.approx(3.0)
    assert applied["want_pt"] == pytest.approx(72.0, abs=0.1)

    para = Document(str(out)).paragraphs[0]
    assert float(para.paragraph_format.space_before.pt) == pytest.approx(72.0, abs=0.2)
    # 复检收敛（读回来就是 72pt，且 checker 认为合规）
    prof2, findings2 = checker.run_check(str(out), t)
    assert prof2.tp_title_space_before_pt == pytest.approx(72.0, abs=0.5)
    assert _sev(findings2, "title_page_layout") == "pass"


def test_title_three_to_four_lines_is_a_range(tmp_path):
    """APA 原文是 "three to four lines" —— **3 行与 4 行都必须判 pass**。

    v2.3.91 按精确 3 行（72pt）判、容差 1pt，一份合规的 4 行标题（96pt）会被误报成错。
    老板 2026-10-23「规范里面没有要求的我们不要加，仔细看规范」—— 逐字读原文才发现。
    """
    t = target_from_spec("APA-Student")
    for pt in (72.0, 84.0, 96.0):
        p = _make_paper(tmp_path / ("a%d.docx" % int(pt)),
                         title_space_before_pt=pt, title_space_after_pt=24.0,
                         extra_spacing=False)
        prof, findings = checker.run_check(p, t)
        assert prof.tp_title_space_before_pt == pytest.approx(pt, abs=0.5)
        assert _sev(findings, "title_page_layout") == "pass", "%.0fpt 应判合规" % pt
        # 合规就不许绑修正动作（绝不白改客户文档）
        assert all(getattr(f, "fix", None) is None for f in _dim(findings, "title_page_layout"))


def test_title_far_below_or_above_range_is_flagged_and_fixed(tmp_path):
    """区间之外（贴顶 = 0pt / 5 行以上）才判错，修到 3 行。"""
    t = target_from_spec("APA-Student")
    for pt in (0.0, 200.0):
        p = _make_paper(tmp_path / ("b%d.docx" % int(pt)), title_space_before_pt=pt)
        prof, findings = checker.run_check(p, t)
        assert _sev(findings, "title_page_layout") == "warn"
        out = tmp_path / ("o%d.docx" % int(pt))
        fix_docx(p, str(out), t)
        prof2, findings2 = checker.run_check(str(out), t)
        assert prof2.tp_title_space_before_pt == pytest.approx(72.0, abs=0.5)
        assert _sev(findings2, "title_page_layout") == "pass"


def test_one_blank_line_after_title_is_kept_not_cleared(tmp_path):
    """APA：标题与作者之间要留**一个双倍行距空行**，不是清零。

    v2.3.91 把标题段后清成 0，等于删掉了这条要求（反向违规）。v2.3.92 改为「约一行」：
    原文的 18pt 在允许带内 → 一个字节都不动；作者等其余元素才清 0。
    """
    t = target_from_spec("APA-Student")
    p = _make_paper(tmp_path / "a.docx", extra_spacing=True)
    prof, _ = checker.run_check(p, t)
    assert prof.tp_title_space_after_pt == pytest.approx(18.0, abs=0.5)
    assert prof.tp_author_space_max_pt == pytest.approx(12.0, abs=0.5)

    out = tmp_path / "out.docx"
    fix_docx(p, str(out), t)
    prof2, findings2 = checker.run_check(str(out), t)
    # 空行必须还在
    assert prof2.tp_title_space_after_pt == pytest.approx(18.0, abs=0.5)
    # 其余元素之间清 0
    assert (prof2.tp_author_space_max_pt or 0.0) == pytest.approx(0.0, abs=0.1)
    assert _sev(findings2, "title_page_layout") == "pass"


def test_missing_blank_line_after_title_is_supplied(tmp_path):
    """标题后既没有空段落、段后也是 0 → 补齐到约一行（24pt = 1 行 × 12pt × 2.0）。"""
    t = target_from_spec("APA-Student")
    p = _make_paper(tmp_path / "a.docx", extra_spacing=False)
    prof, findings = checker.run_check(p, t)
    assert _sev(findings, "title_page_layout") == "warn"
    out = tmp_path / "out.docx"
    fix_docx(p, str(out), t)
    prof2, findings2 = checker.run_check(str(out), t)
    assert prof2.tp_title_space_after_pt == pytest.approx(24.0, abs=1.0)
    assert _sev(findings2, "title_page_layout") == "pass"


def test_blank_paragraph_between_title_and_author_is_accepted(tmp_path):
    """空行也可以用一个真空段落实现 —— 两种写法都得认，否则合规文档被误报。"""
    t = target_from_spec("APA-Student")
    p = _make_paper(tmp_path / "a.docx", extra_spacing=False,
                    blank_after_title=True, title_space_before_pt=72.0)
    prof, findings = checker.run_check(p, t)
    assert prof.tp_author_gap_paragraphs == 1
    assert _sev(findings, "title_page_layout") == "pass"

    out = tmp_path / "out.docx"
    summary = fix_docx(p, str(out), t)
    # 已有真空段 → 不许再补段后间距（否则一个空行变两个）
    assert "after_pt" not in (summary["applied"].get("title_page_layout") or {})


def test_title_page_layout_stays_info_when_not_baseline(tmp_path):
    p = _make_paper(tmp_path / "a.docx", title_space_before_pt=0.0)
    t = target_from_dict({"reference_style": "APA"}, "school")
    prof, findings = checker.run_check(p, t)
    assert _sev(findings, "title_page_layout") == "info"
    assert all(getattr(f, "fix", None) is None for f in _dim(findings, "title_page_layout"))


# ---------------------------------------------------------------------------
# ③ 摘要之后正文另起一页
# ---------------------------------------------------------------------------

def test_text_new_page_fixable_on_baseline(tmp_path):
    p = _make_paper(tmp_path / "a.docx", text_break=False)
    t = target_from_spec("APA-Student")
    prof, findings = checker.run_check(p, t)
    assert prof.text_break_before is False
    assert _sev(findings, "text_new_page") == "warn"

    out = tmp_path / "out.docx"
    summary = fix_docx(p, str(out), t)
    assert summary["applied"]["text_new_page"]["changed"] is True
    prof2, findings2 = checker.run_check(str(out), t)
    assert prof2.text_break_before is True
    assert _sev(findings2, "text_new_page") == "pass"


def test_text_new_page_stays_info_when_not_baseline(tmp_path):
    p = _make_paper(tmp_path / "a.docx", text_break=False)
    t = target_from_dict({"reference_style": "APA"}, "school")
    prof, findings = checker.run_check(p, t)
    assert _sev(findings, "text_new_page") == "info"
    assert all(getattr(f, "fix", None) is None for f in _dim(findings, "text_new_page"))


def test_text_new_page_absent_when_abstract_not_located(tmp_path):
    """摘要标题不是样式标题时定位不到 → 只出 info「无法判定」，绝不猜着插分页。

    猜错的后果是在错误位置插分页 —— 那是真的把客户文档版式改坏，
    比「不修」严重得多。真实样本 student_paper_unformatted_20pages.docx 就是这种。
    """
    doc = Document()
    doc.add_paragraph(TITLE)
    doc.add_paragraph("Abstract")
    doc.add_paragraph("Body of the abstract, but the heading carries no heading style.")
    for i in range(3):
        doc.add_paragraph("Body %d." % (i + 1))
    p = tmp_path / "a.docx"
    doc.save(str(p))

    t = target_from_spec("APA-Student")
    prof, findings = checker.run_check(str(p), t)
    assert prof.text_break_before is None
    assert _sev(findings, "text_new_page") == "info"
    assert all(getattr(f, "fix", None) is None for f in _dim(findings, "text_new_page"))


# ---------------------------------------------------------------------------
# 铁律：纯版式 —— 不新增段落、不改一个字
# ---------------------------------------------------------------------------

def test_structure_fixes_add_no_paragraph_and_change_no_text(tmp_path):
    p = _make_paper(tmp_path / "a.docx")
    before = _texts(p)
    t = target_from_spec("APA-Student")
    out = tmp_path / "out.docx"
    fix_docx(p, str(out), t)
    after = _texts(out)
    assert len(after) == len(before), "首页结构修正不许新增/删除任何段落"
    assert after == before, "首页结构修正不许改动正文文字（文字 100% 锁死）"


def test_page_break_written_as_page_break_before_not_new_paragraph(tmp_path):
    """分页必须写成段前分页属性，**不插空段、不插 w:br**。"""
    p = _make_paper(tmp_path / "a.docx")
    t = target_from_spec("APA-Student")
    out = tmp_path / "out.docx"
    fix_docx(p, str(out), t)

    doc = Document(str(out))
    assert doc.paragraphs[3].text.strip() == "Abstract"
    assert doc.paragraphs[3].paragraph_format.page_break_before is True
    # 整份文档不得有我们新加的硬分页符 run（客户原有的一律保留）
    n_br = sum(1 for br in doc.element.body.iter(qn("w:br"))
               if br.get(qn("w:type")) in (None, "page"))
    assert n_br == 0, "首页结构只许写 pageBreakBefore，不许插 w:br"


def test_customer_existing_page_break_is_preserved(tmp_path):
    """客户自己插的分页符绝不删：只需保证「有分页」这个结果成立。"""
    p = _make_paper(tmp_path / "a.docx", title_break=True, text_break=True)
    before = [p_.text for p_ in Document(str(p)).paragraphs]
    t = target_from_spec("APA-Student")
    out = tmp_path / "out.docx"
    fix_docx(p, str(out), t)
    doc = Document(str(out))
    assert doc.paragraphs[3].paragraph_format.page_break_before is True
    assert [x.text for x in doc.paragraphs] == before


# ---------------------------------------------------------------------------
# 收敛 & 幂等
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("spec_key", ["APA-Student", "APA-Professional"])
def test_three_items_all_pass_after_one_fix(tmp_path, spec_key):
    """一次修正后，首页三项必须全部 pass（不许留 warn，也不许再绑修正）。"""
    p = _make_paper(tmp_path / "a.docx")
    t = target_from_spec(spec_key)
    out = tmp_path / "out.docx"
    fix_docx(p, str(out), t)
    prof2, findings2 = checker.run_check(str(out), t)
    for dim in NEW_DIMS:
        assert _sev(findings2, dim) == "pass", "%s 未收敛：%s" % (dim, _sev(findings2, dim))
        assert all(getattr(f, "fix", None) is None for f in _dim(findings2, dim)), \
            "%s 已合规却仍绑了修正动作" % dim


_COMMENT_RUN_RE = re.compile(
    r"<w:r\b(?:(?!</w:r>).)*?<w:commentReference[^>]*/>(?:(?!</w:r>).)*?</w:r>", re.S)
_COMMENT_MARK_RE = re.compile(r"<w:comment(?:RangeStart|RangeEnd)[^>]*/>")


def _norm(b: bytes) -> str:
    s = b.decode("utf-8", "ignore")
    return _COMMENT_MARK_RE.sub("", _COMMENT_RUN_RE.sub("", s))


def test_structure_fix_is_idempotent(tmp_path):
    """二遍修正不许再改一个字节（批注锚点与 comments.xml 除外）。"""
    p = _make_paper(tmp_path / "a.docx")
    t = target_from_spec("APA-Student")
    o1 = tmp_path / "o1.docx"
    o2 = tmp_path / "o2.docx"
    fix_docx(p, str(o1), t)
    fix_docx(str(o1), str(o2), t)

    def parts(path):
        with zipfile.ZipFile(str(path)) as z:
            return {n: _norm(z.read(n)) for n in z.namelist()
                    if n != "word/comments.xml"}

    a, b = parts(o1), parts(o2)
    assert a == b, "差异部件: %s" % sorted(k for k in set(a) | set(b) if a.get(k) != b.get(k))


# ---------------------------------------------------------------------------
# 客户知情权：预览与批注必须提前说清「页数会变、标题会下移」
# ---------------------------------------------------------------------------

def test_preview_discloses_three_structural_changes(tmp_path):
    p = _make_paper(tmp_path / "a.docx")
    t = target_from_spec("APA-Student")
    zh = preview_fix(p, t, lang="zh")
    en = preview_fix(p, t, lang="en")
    joined_zh = "\n".join(zh)
    joined_en = "\n".join(en)
    assert "题名页独立成页" in joined_zh
    assert "标题下移 3 行（72.0 pt）" in joined_zh
    assert "其余题名页元素额外段距清 0" in joined_zh
    assert "正文另起一页" in joined_zh
    # 海外版英文界面绝不泄漏中文
    assert "Title page on its own page" in joined_en
    assert "3 lines (72.0 pt)" in joined_en
    assert not re.search(r"[\u4e00-\u9fff]", joined_en), "英文预览里出现中文"


def test_summary_comment_discloses_page_break(tmp_path):
    """结尾批注必须说清「题名页现在独立成页，且没有增删段落或字符」。"""
    p = _make_paper(tmp_path / "a.docx")
    t = target_from_spec("APA-Student")
    out = tmp_path / "out.docx"
    fix_docx(p, str(out), t)
    with zipfile.ZipFile(str(out)) as z:
        xml = z.read("word/comments.xml").decode("utf-8", "ignore")
    assert "title page now occupies its own page" in xml.lower()
    assert "no paragraph or character was added" in xml.lower()
    assert "below the top margin" in xml.lower()


# ---------------------------------------------------------------------------
# ④ 题名页内容完整性：只提示，永不代写（老板 2026-10-23「可以做提示」）
# ---------------------------------------------------------------------------

def test_student_title_page_missing_content_is_only_prompted(tmp_path):
    """APA 学生版还要求课程号+课名、教师名、提交日期 —— 缺了只出 info，绝不代写。

    这三项是**内容**不是版式，工具一个字都不会加（文字 100% 锁死），
    只能告诉客户还差什么。severity 恒为 info、fix 恒为 None。
    """
    p = _make_paper(tmp_path / "a.docx")
    t = target_from_spec("APA-Student")
    prof, findings = checker.run_check(p, t)
    got = _dim(findings, "title_page_content")
    assert got, "应出题名页内容提示"
    assert got[0].severity == "info"
    assert all(getattr(f, "fix", None) is None for f in got)
    assert "course number and name" in got[0].actual
    assert "instructor name" in got[0].actual
    assert "assignment due date" in got[0].actual

    out = tmp_path / "out.docx"
    summary = fix_docx(p, str(out), t)
    assert "title_page_content" not in summary["applied"]
    # 修正前后文字必须一模一样（绝不代写）
    assert _texts(out) == _texts(p)


def test_student_title_page_complete_content_passes(tmp_path):
    p = _make_paper(tmp_path / "a.docx", course="PSY 201: Introduction to Psychology",
                    instructor="Dr. Rowan J. Estes", due_date="October 18, 2020")
    t = target_from_spec("APA-Student")
    prof, findings = checker.run_check(p, t)
    got = _dim(findings, "title_page_content")
    assert got and got[0].severity == "pass"


def test_title_page_content_not_checked_for_professional(tmp_path):
    """APA 专业版题名页要的是作者注与 running head，**没有**课程/教师/日期这三项。

    规范里没要求的一律不查 —— 否则专业版客户会收到一条根本没依据的提示。
    """
    p = _make_paper(tmp_path / "a.docx")
    t = target_from_spec("APA-Professional")
    prof, findings = checker.run_check(p, t)
    assert not _dim(findings, "title_page_content")


def test_title_page_content_not_checked_when_not_baseline(tmp_path):
    """学校模板模式：模板可能压根不要课程/教师/日期，不许拿 APA 去指点客户。"""
    p = _make_paper(tmp_path / "a.docx")
    t = target_from_dict({"reference_style": "APA"}, "school")
    prof, findings = checker.run_check(p, t)
    assert not _dim(findings, "title_page_content")


def test_course_code_inside_paper_title_does_not_suppress_prompt(tmp_path):
    """论文大标题里出现 "COVID-19" 这类「字母+数字」，不能把它当成课程号。

    所以内容提示只在**标题以外**那块文本里找线索（reader 的 tp_author_block_text）。
    """
    p = _make_paper(tmp_path / "a.docx", title="COVID-19 and Supply Chain Resilience")
    t = target_from_spec("APA-Student")
    prof, findings = checker.run_check(p, t)
    got = _dim(findings, "title_page_content")
    assert got and got[0].severity == "info"
    assert "course number and name" in got[0].actual
