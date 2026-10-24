# -*- coding: utf-8 -*-
"""F8 回归：居中加粗 Cambria Math 公式段绝不能被当成伪标题加编号/冒号。

详见 src/engine/docx_reader.py v2.3.104。公式（如 `MTTR = β₀ + … + ε`）整段加粗、
短、不以句读结尾，旧逻辑会误判为伪标题并被加 run-in 冒号 + 编号，违反文字锁死铁律。
判定新增两道闸：① 含符号/数学字体（Cambria Math 等）的 run → 非标题；
② 文本含 `=` 且含希腊字母 → 非标题形状。
"""
import os
import tempfile

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from src.engine import checker, fixer
from src.engine.docx_reader import is_pseudo_heading, is_section_heading_line
from src.engine.questionnaire import build_target


EQUATION = "MTTR = β₀ + β₁(DMI) + β₂(SCV) + ε"


def _build(path):
    d = Document()
    d.add_heading("Introduction", level=1)
    d.add_paragraph("We study the effect of data maturity on model performance.")
    d.add_heading("Method", level=1)
    d.add_paragraph("The regression model is defined as follows.")
    pe = d.add_paragraph()
    pe.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = pe.add_run(EQUATION)
    r.bold = True
    r.font.name = "Cambria Math"
    r.font.size = Pt(11)
    d.add_heading("References", level=1)
    d.add_paragraph("[1] Author, Title, Journal, 2024.")
    d.save(path)


def test_equation_is_not_pseudo_heading():
    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, "eq.docx")
        _build(src)
        doc = Document(src)
        eq_para = next(p for p in doc.paragraphs if EQUATION.split(" = ")[0] in (p.text or ""))
        assert is_pseudo_heading(eq_para) is False
        assert is_section_heading_line(eq_para) is False


def test_equation_text_locked_through_fix():
    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, "eq.docx")
        out1 = os.path.join(td, "eq_out1.docx")
        out2 = os.path.join(td, "eq_out2.docx")
        _build(src)

        target = build_target("IEEE-Conference")
        spec = checker.resolve_spec(target)
        _, f0 = checker.run_check(src, target)
        fixer.fix_docx(src, out1, target, findings=f0, spec=spec)
        _, f2 = checker.run_check(out1, target)
        fixer.fix_docx(out1, out2, target, findings=f2, spec=spec)

        d_out = Document(out1)
        eq_paras = [p for p in d_out.paragraphs if EQUATION.split(" = ")[0] in (p.text or "")]
        # 文字锁死：公式一字未改，也无批注
        assert [p.text for p in eq_paras] == [EQUATION]
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        for p in eq_paras:
            assert p._p.find(f".//{{{ns}}}commentRangeStart") is None

        # 全篇词级文本保全
        w0 = [w for p in Document(src).paragraphs for w in (p.text or "").split() if w]
        w1 = [w for p in d_out.paragraphs for w in (p.text or "").split() if w]
        assert w0 == w1
