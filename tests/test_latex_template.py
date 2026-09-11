"""LaTeX 模板解析器断言式测试（pytest）。

守护：纯正则/结构解析、不跑 TeX 引擎、不解析自定义宏；覆盖主流 80% 模板。
"""

from __future__ import annotations

import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine.latex_template import parse_latex_source, parse_latex_template


def test_ieeetr_extracts_ieee_style_and_no_margins():
    src = r"""
\documentclass[10pt,conference]{IEEEtran}
\usepackage{cite}
\begin{document}
\bibliographystyle{IEEEtran}
\end{document}
"""
    t = parse_latex_source(src)
    assert t.reference_style == "IEEE"
    assert t.reference_numbered is True
    assert t.reference_hanging_indent_in == 0.0
    assert t.page.get("font_size_pt") == 10.0


def test_apa_extracts_margins_and_double_spacing():
    src = r"""
\documentclass[12pt]{article}
\usepackage[margin=1in,top=2cm,bottom=2cm,left=1in,right=1in]{geometry}
\usepackage{setspace}
\doublespacing
\bibliographystyle{apacite}
"""
    t = parse_latex_source(src)
    # font size
    assert t.page.get("font_size_pt") == 12.0
    # margins: top/bottom overridden to 2cm (~0.787in), left/right from 1in
    assert t.page["margin_top_in"] == pytest.approx(2 / 2.54, rel=1e-3)
    assert t.page["margin_bottom_in"] == pytest.approx(2 / 2.54, rel=1e-3)
    assert t.page["margin_left_in"] == pytest.approx(1.0)
    assert t.page["margin_right_in"] == pytest.approx(1.0)
    # line spacing
    assert t.page["line_spacing"] == 2.0
    # bib style
    assert t.reference_style == "APA"


def test_doublespacing_via_setspace_macro():
    src = r"""
\documentclass[11pt]{article}
\usepackage{setspace}
\doublespacing
"""
    t = parse_latex_source(src)
    assert t.page["line_spacing"] == 2.0


def test_setstretch_takes_precedence_over_named():
    src = r"""
\doublespacing
\setstretch{1.75}
"""
    t = parse_latex_source(src)
    assert t.page["line_spacing"] == 1.75


def test_unknown_bibstyle_returns_none():
    src = r"""
\documentclass{article}
\bibliographystyle{weird-unknown-style}
"""
    t = parse_latex_source(src)
    assert t.reference_style is None


def test_no_bibliographystyle_returns_none():
    src = r"""
\documentclass{article}
\begin{document}
No refs here.
\end{document}
"""
    t = parse_latex_source(src)
    assert t.reference_style is None
    # 其他字段未识别也应留空
    assert "font_size_pt" not in t.page
    assert "line_spacing" not in t.page


def test_harvard_agsm_family():
    src = r"""
\documentclass{article}
\bibliographystyle{agsm}
"""
    t = parse_latex_source(src)
    assert t.reference_style == "Harvard"


def test_parse_latex_template_from_file(tmp_path):
    f = tmp_path / "apa.cls"
    f.write_text(
        r"\documentclass[12pt]{article}"
        r"\usepackage[margin=1in]{geometry}"
        r"\onehalfspacing"
        r"\bibliographystyle{apacite}",
        encoding="utf-8",
    )
    t = parse_latex_template(str(f))
    assert t.reference_style == "APA"
    assert t.page["font_size_pt"] == 12.0
    assert t.page["margin_top_in"] == pytest.approx(1.0)
    assert t.page["line_spacing"] == 1.5


def test_no_geometry_no_margins():
    src = r"\documentclass{article}\begin{document}\end{document}"
    t = parse_latex_source(src)
    assert all(not k.startswith("margin_") for k in t.page)