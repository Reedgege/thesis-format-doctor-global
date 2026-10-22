"""首页按规范分流（v2.3.90）：APA 学生版 / APA 专业版 / MLA 的页眉与页码。

背景（老板 2026-10-22 清单第 ② 条「首页按 APA-Stu / APA-Pro / MLA 分别处理」）：

v2.3.89 之前，**两份产物一模一样**。硬证据（`_smoke/probe_firstpage.py` 实测老板
自己出的四份产物）：
    apa-stu  document.xml = 95272 字节
    apa-p    document.xml = 95272 字节   ← 逐字节 100% 相同，text 节点 207 vs 207 全等
    四份产物 zip 里 **连页眉部件都不存在**，headerReference 出现 0 次
即：选 APA 学生版还是专业版，产出毫无区别；APA/MLA 要求的页码一个都没有。

根因两条：
  ① `page_number` 属于「页面维度」，来自规范时被 `_downgrade` 一律降为 info →
     checker 只提示、fixer 不动手 → 永远不会有页码；
  ② `running_head` 恒为 info 且 fix=None（注释写着「仅提示、改内容」）→
     而 APA-Pro 与 APA-Stu 在**排版层**唯一的真差异恰恰就是 running head。

老板拍板：① 页码升为可修；② running head 做（从客户自己的标题/姓名行派生）。

本文件钉死的契约：
- APA-Pro 页眉 = 全大写短标题（≤50 字符，词边界截断）+ 页码，**同段**左对齐 + 右制表位；
- APA-Stu 只有页码、**没有** running head（学生版不要求，绝不许误加）；
- MLA 页眉 = 姓氏 + 页码，右上；
- 两份产物必须**真的不同**（第 ② 条的验收标准）；
- 正文文字一行都不许变（文字 100% 锁死铁律）；
- 客户自己写过的页眉**绝不覆盖**；
- 非规范基线（学校模板 / 手填 / AI 导入）一个字节都不写；
- 修完复检，page_number / running_head 不许再有 warn（检查与修正必须同源）；
- 格式层幂等（第二遍不再改任何格式，只清掉已无话可说的旧批注）。
"""

from __future__ import annotations

import os
import re
import sys
import zipfile

from docx import Document
from docx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine.checker import run_check
from src.engine.docx_reader import read_docx
from src.engine.fixer import (
    _mla_surname_from_doc,
    _short_title_from_doc,
    compute_changes,
    fix_docx,
    preview_fix,
)
from src.engine.questionnaire import target_from_spec

TITLE = "Digital Transformation and Supply Chain Resilience"
BODY_1 = (
    "Digital technology adoption has reshaped how manufacturing firms coordinate "
    "with their upstream and downstream partners across regions."
)
BODY_2 = (
    "Existing literature on supply chain risk management can be broadly divided "
    "into two streams, neither of which fully explains the mechanism at work here."
)
_COMMENT_RUN_RE = re.compile(
    r"<w:r\b(?:(?!</w:r>).)*?<w:commentReference[^>]*/>(?:(?!</w:r>).)*?</w:r>", re.S)
_COMMENT_MARK_RE = re.compile(r"<w:comment(?:RangeStart|RangeEnd)[^>]*/>")


def _strip_comment_marks(text: str) -> str:
    """剥掉批注锚点：先删承载 commentReference 的整个 run，再删范围标记。

    只删标记会留下一堆空的 ``<w:r>`` 壳子 —— 首修文件里有、二修文件里没有，
    会被误判成「格式不幂等」。
    """
    return _COMMENT_MARK_RE.sub("", _COMMENT_RUN_RE.sub("", text))


# --------------------------------------------------------------------- helpers
def _make_paper(path: str, *, header_text: str = "") -> str:
    """造一篇首页结构贴近真实样本的短文（标题 + 作者/单位块 + 摘要 + 正文 + 参考文献）。"""
    doc = Document()
    p = doc.add_paragraph()
    r = p.add_run(TITLE)
    r.bold = True
    doc.add_paragraph("By: Alex Morgan")
    doc.add_paragraph("School of Business and Management")
    doc.add_paragraph("Abstract")
    doc.add_paragraph(BODY_1)
    doc.add_paragraph(BODY_2)
    doc.add_paragraph("References")
    doc.add_paragraph("Ivanov, D. (2020). Predicting the impacts of epidemic outbreaks.")
    if header_text:
        hdr = doc.sections[0].header
        hdr.is_linked_to_previous = False
        hdr.paragraphs[0].text = header_text
    doc.save(path)
    return path


def _text_lines(path: str) -> list:
    """按软换行拆成「行」做字符级比对（段数可能变，行数不该变）。"""
    doc = Document(path)
    out: list = []
    for p in doc.paragraphs:
        buf: list = []
        for child in p._p:
            if child.tag == qn("w:pPr"):
                continue
            if child.tag == qn("w:r"):
                for el in child:
                    if el.tag == qn("w:rPr"):
                        continue
                    if el.tag in (qn("w:br"), qn("w:cr")):
                        out.append("".join(buf))
                        buf = []
                    elif el.tag == qn("w:t"):
                        buf.append(el.text or "")
        out.append("".join(buf))
    return out


def _header_parts(path: str) -> dict:
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n).decode("utf8") for n in z.namelist()
                if re.search(r"word/header\d*\.xml$", n)}


def _header_text(path: str) -> str:
    """页眉里的可见文字 —— 直接用引擎自己的读取口径（与 checker 完全同源）。"""
    return (read_docx(path).header_text or "").strip()


def _has_page_field(path: str) -> bool:
    return any("PAGE" in xml for xml in _header_parts(path).values())


def _content_payload(path: str) -> str:
    """文档**格式内容**的指纹：document.xml + 页眉/页脚，剔掉批注。

    批注不算格式：第二遍修正时「已经没什么可改的了」，于是会把上一遍留下的批注
    清掉（comments.xml 从 4 条变 0 条）。这是正确行为，不该拿来判幂等失败。
    """
    keep = []
    with zipfile.ZipFile(path) as z:
        for n in sorted(z.namelist()):
            if not (n.startswith("word/") and n.endswith(".xml")):
                continue
            if n == "word/comments.xml":
                continue
            keep.append(z.read(n).decode("utf8"))
    blob = "\n".join(keep)
    return _strip_comment_marks(blob)


def _fix(tmp_path, style: str, src: str | None = None, tag: str = "") -> str:
    src = src or _make_paper(str(tmp_path / ("src%s.docx" % tag)))
    out = str(tmp_path / ("fixed-%s%s.docx" % (style, tag)))
    fix_docx(src, out, target_from_spec(style))
    return out


# ------------------------------------------------------------------- 分流铁证
def test_apa_student_and_professional_outputs_differ(tmp_path):
    """🔴 第 ② 条的验收标准本身：两份产物必须真的不同（此前逐字节 100% 相同）。"""
    stu = _fix(tmp_path, "APA-Student")
    pro = _fix(tmp_path, "APA-Professional", tag="-p")
    assert _content_payload(stu) != _content_payload(pro)


def test_apa_professional_writes_uppercase_running_head(tmp_path):
    pro = _fix(tmp_path, "APA-Professional")
    txt = _header_text(pro)
    assert txt == TITLE.upper()
    assert txt == txt.upper()
    assert len(txt) <= 50


def test_apa_student_has_page_number_but_no_running_head(tmp_path):
    """学生版不要求 running head —— 绝不许误加（加了就是凭空多出文字）。"""
    stu = _fix(tmp_path, "APA-Student")
    assert _has_page_field(stu) is True
    assert _header_text(stu) == ""


def test_mla_running_head_is_surname_plus_page_number(tmp_path):
    mla = _fix(tmp_path, "MLA")
    assert _header_text(mla) == "Morgan"
    assert _has_page_field(mla) is True


# --------------------------------------------------------------- 页眉布局细节
def test_running_head_shares_one_paragraph_with_page_number(tmp_path):
    """APA 专业版页眉是「短标题<TAB>页码」**同一段** —— 分两段就成两行了。"""
    pro = _fix(tmp_path, "APA-Professional")
    xml = next(iter(_header_parts(pro).values()))
    paras = [p for p in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S) if "<w:fldChar" in p]
    assert len(paras) == 1
    one = paras[0]
    assert "<w:tab/>" in one
    assert TITLE.upper() in one


def test_running_head_paragraph_is_left_aligned_with_right_tab(tmp_path):
    """必须左对齐 + right 制表位：靠右对齐会把短标题一起甩到右边。"""
    pro = _fix(tmp_path, "APA-Professional")
    xml = next(iter(_header_parts(pro).values()))
    assert re.search(r'<w:tab w:pos="\d+" w:val="right"/>', xml)
    assert '<w:jc w:val="left"/>' in xml


# ---------------------------------------------------------- 检查与修正必须同源
def test_after_fix_no_page_number_or_running_head_warning(tmp_path):
    """修完复检不许再报 warn —— 否则客户看到「修了还是错」。"""
    for style in ("APA-Student", "APA-Professional", "MLA"):
        out = _fix(tmp_path, style, tag="-" + style)
        _prof, findings = run_check(out, target_from_spec(style))
        bad = [f for f in findings
               if f.dimension in ("page_number", "running_head")
               and f.severity in ("warn", "fail")]
        assert not bad, (style, [(f.dimension, f.actual) for f in bad])


def test_page_number_is_fixable_only_on_spec_baseline(tmp_path):
    """学校模板 / 手填 / AI 导入（spec_key=Other）时不许替客户拍板写页码。"""
    src = _make_paper(str(tmp_path / "src.docx"))
    other = target_from_spec("Other")
    _prof, findings = run_check(src, other)
    assert not [f for f in findings if f.dimension == "page_number" and f.fix is not None]
    out = str(tmp_path / "fixed-other.docx")
    fix_docx(src, out, other)
    assert _header_parts(out) == {}


# -------------------------------------------------------------- 文字锁死
def test_body_text_is_never_changed(tmp_path):
    src = _make_paper(str(tmp_path / "src.docx"))
    before = _text_lines(src)
    for style in ("APA-Student", "APA-Professional", "MLA"):
        out = _fix(tmp_path, style, src=src, tag="-" + style)
        assert _text_lines(out) == before, style


def test_existing_header_text_is_never_overwritten(tmp_path):
    """客户自己写过页眉 → 绝不覆盖，一个字都不动。"""
    src = _make_paper(str(tmp_path / "src.docx"), header_text="My own header")
    out = str(tmp_path / "fixed.docx")
    fix_docx(src, out, target_from_spec("APA-Professional"))
    assert _header_text(out) == "My own header"


# -------------------------------------------------------------- 字数上限
def test_short_title_is_cut_at_a_word_boundary(tmp_path):
    doc = Document()
    doc.add_paragraph("A Very Long Title That Definitely Exceeds Fifty Characters In Total Length")
    short = _short_title_from_doc(doc, 50)
    assert len(short) <= 50
    assert short == short.upper()
    assert not short.endswith(" ") and short.endswith(next(
        w for w in reversed(short.split()) if w))
    # 不许切出半个单词：末尾那个词必须在原标题里完整出现
    assert short.split()[-1] in "A Very Long Title That Definitely Exceeds Fifty Characters In Total Length".upper()


def test_missing_title_does_not_crash(tmp_path):
    doc = Document()
    doc.add_paragraph("")
    assert _short_title_from_doc(doc, 50) == ""
    assert _mla_surname_from_doc(doc) == ""


# ------------------------------------------------------ 预览 / 报告如实披露
def test_preview_discloses_the_running_head_text(tmp_path):
    src = _make_paper(str(tmp_path / "src.docx"))
    for lang, needle in (("zh", "页眉写入"), ("en", "Running head written")):
        lines = preview_fix(src, target_from_spec("APA-Professional"), lang=lang)
        hit = [ln for ln in lines if needle in ln]
        assert hit, (lang, lines)
        assert TITLE.upper() in hit[0]


def test_summary_comment_discloses_header_text(tmp_path):
    """摘要批注里必须写明页眉多出了哪几个字 —— 否则「未增删任何文字」会误导客户。"""
    pro = _fix(tmp_path, "APA-Professional")
    with zipfile.ZipFile(pro) as z:
        comments = z.read("word/comments.xml").decode("utf8")
    assert "Running head" in comments
    assert TITLE.upper() in comments


# -------------------------------------------------------------- 幂等
def test_formatting_is_idempotent(tmp_path):
    """对已修文件再修一遍：格式一个字节都不许变（只会清掉已无话可说的旧批注）。"""
    src = _make_paper(str(tmp_path / "src.docx"))
    once = str(tmp_path / "once.docx")
    twice = str(tmp_path / "twice.docx")
    fix_docx(src, once, target_from_spec("APA-Professional"))
    fix_docx(once, twice, target_from_spec("APA-Professional"))
    assert _content_payload(once) == _content_payload(twice)
