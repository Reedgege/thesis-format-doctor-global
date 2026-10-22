# -*- coding: utf-8 -*-
"""v2.3.84 · 文字锁死 + 参考文献粘段拆分（老板 2026-10-21 更新清单第 4、5 条）。

真实样本（`student_paper_unformatted_20pages.docx`）暴露的两个硬问题：

1. **整份文献表被 Shift+Enter 粘在一个段落里** → 引擎按段落采条目，只认成 1 条，
   悬挂缩进也无处落（真实数据：应 3 条，旧版只 1 条）。修法 = 在读/改两侧共用的
   唯一入口处把软换行拆成真实段落，只加段落标记，一个字符都不动。
2. **IEEE 会改写标题文字**（插 "I. " 前缀 + upper()）与**参考文献条目正文**
   （APA→IEEE 引文重排）→ 与第 5 条「100% 锁死文字与标点」直接冲突。
   修法 = `policy.TEXT_IS_LOCKED`，检查器同步把这两项降级为 info 提示，
   绝不出现"报了却改不了"。
"""
import os
import re
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)


def _build(path, glued=3, heading=None):
    """造一份「文献表用软换行粘成一段」的最小论文。"""
    from docx import Document

    d = Document()
    d.add_paragraph("A Study of Glued Reference Paragraphs")
    if heading:
        hp = d.add_paragraph()
        hp.add_run(heading).bold = True      # 手打加粗标题：真实论文的常态
    d.add_paragraph(
        "This body paragraph is long enough to be treated as body text by the "
        "engine, and it contains a few sentences so that it is not mistaken for "
        "a heading of any kind."
    )
    d.add_paragraph("References")
    p = d.add_paragraph()
    for i in range(glued):
        if i:
            p.add_run().add_break()          # <w:br/> —— 学生常用 Shift+Enter
        p.add_run("Author%d, A. (2020). Title number %d. Journal of Tests, "
                  "1(2), 33-40." % (i + 1, i + 1))
    d.save(str(path))
    return path


def _text_lines(path):
    """全文按"行"取文本（软换行也算行），并去掉 [n] 列表标记 —— 用于文字锁死对账。"""
    from docx import Document

    out = []
    for para in Document(str(path)).paragraphs:
        for line in para.text.split("\n"):
            t = re.sub(r"^\s*\[\d+\]\s*", "", line).strip()
            if t:
                out.append(t)
    return out


def _fix(src, out, spec_key):
    from src.engine import checker
    from src.engine.fixer import fix_docx
    from src.engine.questionnaire import build_target

    target = build_target(spec_key)
    prof, findings = checker.run_check(str(src), target)
    summary = fix_docx(str(src), str(out), target,
                       findings=findings, spec=checker.resolve_spec(target))
    return prof, findings, summary


# ------------------------------------------------------------ 粘段拆分
def test_glued_reference_entries_are_split(tmp_path):
    """粘在一段里的 3 条文献，必须被认成 3 条（旧版是 1 条）。"""
    from src.engine.docx_reader import read_docx

    src = _build(tmp_path / "glued.docx", glued=3)
    prof = read_docx(str(src))
    assert len(prof.reference_entries) == 3, prof.reference_entries


def test_split_preserves_text_exactly(tmp_path):
    """拆分只加段落标记 —— 正文与文献的文字必须逐行完全一致。"""
    src = _build(tmp_path / "glued.docx", glued=3)
    before = _text_lines(src)
    _fix(src, tmp_path / "out.docx", "APA-Student")
    after = _text_lines(tmp_path / "out.docx")
    assert before == after, "文字被改动了：\n原 %r\n新 %r" % (before, after)


def test_each_entry_gets_hanging_indent_and_heading_does_not(tmp_path):
    src = _build(tmp_path / "glued.docx", glued=3)
    _fix(src, tmp_path / "out.docx", "APA-Student")
    from docx import Document

    d = Document(str(tmp_path / "out.docx"))
    paras = [p for p in d.paragraphs if p.text.strip()]
    idx = next(i for i, p in enumerate(paras)
               if p.text.strip().lower() == "references")
    head = paras[idx]
    assert head.paragraph_format.first_line_indent in (None, 0) or \
        head.paragraph_format.first_line_indent.inches == 0, \
        "References 标题不该被正文首行缩进"
    entries = paras[idx + 1: idx + 1 + 3]
    assert len(entries) == 3
    for e in entries:
        assert e.paragraph_format.left_indent.inches == pytest.approx(0.5, abs=0.01)
        assert e.paragraph_format.first_line_indent.inches == pytest.approx(-0.5, abs=0.01)


def test_body_paragraphs_with_soft_breaks_are_split(tmp_path):
    """v2.3.87 起正文软换行也拆（老板 2026-10-22 拍板「先拆粘连段」）。

    旧契约（v2.3.84）：正文一律不动，只拆参考文献区。实测证明不够 ——
    正文里「章节标题 + 正文」被 Shift+Enter 粘成一块时，标题行会跟着吃正文的
    0.5″ 首行缩进，客户一眼看出错位。现在两者都拆，但**文字一个字符不许变**。
    """
    from docx import Document

    d = Document()
    d.add_paragraph("Title")
    p = d.add_paragraph()
    p.add_run("First soft line.")
    p.add_run().add_break()
    p.add_run("Second soft line.")
    d.add_paragraph("References")
    d.add_paragraph("Author, A. (2020). Something. Journal, 1(1), 1-2.")
    src = tmp_path / "body_glue.docx"
    d.save(str(src))

    _fix(src, tmp_path / "out.docx", "APA-Student")
    paras = Document(str(tmp_path / "out.docx")).paragraphs
    texts = [p.text for p in paras]
    # 拆成了两个真段落
    assert "First soft line." in texts
    assert "Second soft line." in texts
    # 但没有任何一段还带着软换行（软换行已被拆掉，不是复制一遍）
    assert not any("\n" in t for t in texts if "soft line" in t), texts
    # 🔴 铁律：文字逐行一致
    assert _text_lines(src) == _text_lines(tmp_path / "out.docx")


# ------------------------------------------------------------ 文字锁死
def test_ieee_does_not_rewrite_heading_text(tmp_path):
    """IEEE 不再给标题插 "I. " 前缀、不再全大写（清单第 5 条）。"""
    src = _build(tmp_path / "h.docx", glued=1, heading="Chapter 1: Introduction & Background")
    before = _text_lines(src)
    _fix(src, tmp_path / "out.docx", "IEEE")
    after = _text_lines(tmp_path / "out.docx")
    assert before == after, "IEEE 改写了文字：\n原 %r\n新 %r" % (before, after)


def test_ieee_reference_rewrite_is_off_by_default():
    """APA→IEEE 引文重排属于改正文，默认必须关闭。"""
    from src.engine.questionnaire import target_from_spec

    target = target_from_spec("IEEE")
    assert target.reference_rewrite is False


def test_checker_downgrades_numbering_to_info(tmp_path):
    """编号/大写要求改成 info 提示，不再算"待修正"（避免报了却改不了）。

    注：合成文档很难稳定触发「手打加粗标题」的层级识别（引擎对层级靠字号档推断），
    所以这里直接测负责降级的那段逻辑本身；真实论文链路已由
    `_smoke/audit_real_paper.py` 与 test_ieee_does_not_rewrite_heading_text 覆盖。
    """
    from src.engine import checker

    findings = []
    checker._append_text_lock_note(
        findings, 1, ["numbering (I., II., III.)", "uppercase"], "spec:IEEE")
    assert findings, "没记下 info 提示"
    note = findings[0]
    assert note.severity == "info"
    assert note.dimension == "heading"
    assert "left as-is" in (note.message or "")
    assert "I., II., III." in (note.expected or "")
    assert getattr(note, "fix", None) is None, "不该绑修正（否则就是报了改不了）"


def test_checker_ieee_does_not_fail_on_numbering(tmp_path):
    """IEEE 检查里不该再出现「编号缺失」的 fail 结论。"""
    src = _build(tmp_path / "h.docx", glued=1, heading="Chapter 1: Introduction & Background")
    _, findings, _ = _fix(src, tmp_path / "out.docx", "IEEE")
    fails = [f for f in findings
             if f.severity == "fail" and "numbering" in (f.expected or "")]
    assert not fails, "编号不该再被报成 fail（修正器已不改文字）"


def test_policy_flag_is_single_source():
    """检查器与修正器必须读同一个开关，否则会脱节成"报了改不了"。"""
    from src.engine import checker, fixer, policy

    assert policy.TEXT_IS_LOCKED is True
    assert getattr(checker, "TEXT_IS_LOCKED", None) is True
    assert getattr(fixer, "TEXT_IS_LOCKED", None) is True


def test_page_break_inside_reference_paragraph_is_preserved(tmp_path):
    """拆软换行时**绝不能**把分页符（w:br type=page）当段边界拆掉。

    这是 v2.3.84 开发中真实踩到的坑：第一版把 w:br 一律当软换行，结果客户文档里
    「参考文献另起一页」的分页符被吃掉（tests/test_review_fixes.py::
    test_numbering_preserves_page_break 当场翻红才发现的）。
    """
    from docx import Document
    from docx.enum.text import WD_BREAK

    d = Document()
    d.add_paragraph("Body paragraph for the thesis.")
    d.add_paragraph("References")
    p = d.add_paragraph()
    p.add_run("Author, A. (2020). First entry. Journal, 1(1), 1-2.")
    p.add_run().add_break(WD_BREAK.PAGE)      # 分页符（排版语义，不是软换行）
    p.add_run("Author, B. (2021). Second entry. Journal, 2(2), 3-4.")
    src = tmp_path / "pagebreak.docx"
    d.save(str(src))

    _fix(src, tmp_path / "out.docx", "APA-Student")
    xml = Document(str(tmp_path / "out.docx")).paragraphs[2]._p.xml
    assert 'w:type="page"' in xml, "分页符被拆段吃掉了（客户会丢分页）"


def test_page_break_inside_reference_paragraph_is_preserved(tmp_path):
    """拆软换行时**绝不能**把分页符（w:br type=page）当段边界拆掉。

    这是 v2.3.84 开发中真实踩到的坑：第一版把 w:br 一律当软换行，结果客户文档里
    「参考文献另起一页」的分页符被吃掉（tests/test_review_fixes.py::
    test_numbering_preserves_page_break 当场翻红才发现的）。
    """
    from docx import Document
    from docx.enum.text import WD_BREAK

    d = Document()
    d.add_paragraph("Body paragraph for the thesis.")
    d.add_paragraph("References")
    p = d.add_paragraph()
    p.add_run("Author, A. (2020). First entry. Journal, 1(1), 1-2.")
    p.add_run().add_break(WD_BREAK.PAGE)      # 分页符（排版语义，不是软换行）
    p.add_run("Author, B. (2021). Second entry. Journal, 2(2), 3-4.")
    src = tmp_path / "pagebreak.docx"
    d.save(str(src))

    _fix(src, tmp_path / "out.docx", "APA-Student")
    xml = Document(str(tmp_path / "out.docx")).paragraphs[2]._p.xml
    assert 'w:type="page"' in xml, "分页符被拆段吃掉了（客户会丢分页）"
