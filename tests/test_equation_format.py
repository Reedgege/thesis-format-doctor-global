# -*- coding: utf-8 -*-
"""公式（Word 原生 OMML）查改回归（v2.3.112）。

为什么单独立文件
----------------
学校模板里「公式居中、编号加括号靠右边界、与正文同字号、段前段后 0」是高频硬要求，
而引擎此前**完全读不到** Word 公式：OMML 的数学 run 是 ``m:r``，python-docx 的
``para.runs`` 只返回 ``w:r`` —— 公式段在引擎眼里是「除了编号一个字都没有的空段」，
任何维度都判不到它。``report.KNOWN_LIMITATIONS`` 里那句「Complex formulas / equations
are NOT processed」说的就是这个空白。

本文件钉死五件事：
1. reader 读得到 OMML（展示 / 行内、对齐、字号、段距、行距、编号位置、表格内）；
2. 模板批注里的公式条款能解析，且**不会串到正文/字体通道**；
3. 批注 / 问卷 → TargetProfile 贯通，且 ``build_target`` 合并时**不丢字段**；
4. 检查 → 一次性修好 → 复检全绿，第二遍零动作（幂等）；
5. 修正只碰版式：公式文本与编号文字**一个字符都不变**。
"""

from __future__ import annotations

import pytest
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt

from src.engine.checker import run_check
from src.engine.comment_spec_en import parse_comments
from src.engine.docx_reader import read_docx
from src.engine.fixer import fix_docx
from src.engine.questionnaire import build_target, target_from_template_docx
from tests._docx_fixtures import inject_comments, mark_comment

EQ_DIMS = ("equation_alignment", "equation_number_position", "equation_size",
           "equation_space_before", "equation_space_after", "equation_line_spacing",
           "equation_font", "equation_format")

# 一套典型的「学校模板公式要求」
EQ_Q = {
    "equation_alignment": "center",
    "equation_number_position": "right",
    "equation_size_pt": 12.0,
    "equation_space_before_pt": 0.0,
    "equation_space_after_pt": 0.0,
    "equation_line_spacing": 2.0,
}


# ---------------------------------------------------------------------------
# OMML 夹具：python-docx 不会造公式，只能注入合法 XML
# ---------------------------------------------------------------------------
def _omath(inner: str) -> str:
    return '<m:oMath %s>%s</m:oMath>' % (nsdecls("m", "w"), inner)


def _disp(inner: str, jc="center") -> str:
    pr = ('<m:oMathParaPr><m:jc m:val="%s"/></m:oMathParaPr>' % jc) if jc else ""
    return ('<m:oMathPara %s>%s<m:oMath>%s</m:oMath></m:oMathPara>'
            % (nsdecls("m", "w"), pr, inner))


def _eq_body(size_half_pt=None) -> str:
    """``E = m c^2``；给 ``size_half_pt`` 时带上显式字号（单位：半磅）。"""
    rpr = ('<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>'
           '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
           % (size_half_pt, size_half_pt)) if size_half_pt else ""
    one = "<m:r>%s<m:t>%s</m:t></m:r>"
    return (one % (rpr, "E") + one % (rpr, "=") + one % (rpr, "m")
            + "<m:sSup><m:e>%s</m:e><m:sup>%s</m:sup></m:sSup>"
            % (one % (rpr, "c"), one % (rpr, "2")))


def _add_eq_para(doc, parts, *, tabs_right_at=None, before_pt=None, after_pt=None,
                 line_spacing=None):
    """插一个含公式的段落（可带右制表位 / 段距 / 行距）。"""
    p_el = doc.add_paragraph()._p
    ppr = p_el.get_or_add_pPr()
    if tabs_right_at is not None:
        ppr.append(parse_xml('<w:tabs %s><w:tab w:val="right" w:pos="%d"/></w:tabs>'
                             % (nsdecls("w"), tabs_right_at)))
    attrs = ""
    if before_pt is not None:
        attrs += ' w:before="%d"' % int(before_pt * 20)
    if after_pt is not None:
        attrs += ' w:after="%d"' % int(after_pt * 20)
    if line_spacing is not None:
        # ⚠️ pPr 只允许一个 w:spacing：段前段后与行距必须写在同一元素上，
        # 分开追加两个 → 只认第一个，行距被静默丢掉。
        attrs += ' w:line="%d" w:lineRule="auto"' % int(line_spacing * 240)
    if attrs:
        ppr.append(parse_xml("<w:spacing %s%s/>" % (nsdecls("w"), attrs)))
    for x in parts:
        p_el.append(parse_xml(x))
    return p_el


def _new_doc():
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(12)
    return doc


def _build(*, compliant=True, with_number=True, tab_stop=True):
    """造一份含展示公式的文档：``compliant`` 决定版式是否已合规。"""
    doc = _new_doc()
    doc.add_heading("Equations", level=1)
    doc.add_paragraph("Body text before the equation.")
    if compliant:
        parts = [_disp(_eq_body(24), "center"),
                 '<w:r %s><w:tab/></w:r>' % nsdecls("w"),
                 '<w:r %s><w:t>(1)</w:t></w:r>' % nsdecls("w")]
        _add_eq_para(doc, parts, tabs_right_at=9360, before_pt=0, after_pt=0,
                     line_spacing=2.0)
    else:
        parts = [_disp(_eq_body(20), "left")]
        if with_number:
            # 有 Tab 但没有右对齐制表位 → 编号没被推到右边界（可安全修正的那一档）
            parts.append('<w:r %s><w:tab/></w:r>' % nsdecls("w"))
            parts.append('<w:r %s><w:t>(2)</w:t></w:r>' % nsdecls("w"))
        if tab_stop:
            _add_eq_para(doc, parts, before_pt=12, after_pt=12, line_spacing=1.5)
        else:
            _add_eq_para(doc, parts, before_pt=12, after_pt=12, line_spacing=1.5)
    doc.add_paragraph("Body text after the equation.")
    return doc


def _eq_findings(path, target):
    _prof, findings = run_check(path, target)
    return [f for f in findings if f.dimension in EQ_DIMS]


def _by_dim(findings):
    out: dict = {}
    for f in findings:
        out.setdefault(f.dimension, []).append(f)
    return out


# ---------------------------------------------------------------------------
# 1. reader：读得到 OMML
# ---------------------------------------------------------------------------
def test_reader_samples_display_and_inline_equations(tmp_path):
    doc = _new_doc()
    # 展示公式：居中 + 右制表位 + Tab + (1)，12pt，段前段后 0，双倍行距
    _add_eq_para(doc, [_disp(_eq_body(24), "center"),
                       '<w:r %s><w:tab/></w:r>' % nsdecls("w"),
                       '<w:r %s><w:t>(1)</w:t></w:r>' % nsdecls("w")],
                 tabs_right_at=9360, before_pt=0, after_pt=0, line_spacing=2.0)
    # 行内公式：正文里混排，不该被当成展示公式
    p = doc.add_paragraph("Inline math ")
    p._p.append(parse_xml(_omath('<m:r><m:t>x</m:t></m:r>')))
    p.add_run(" sits in a sentence.")
    path = str(tmp_path / "eq.docx")
    doc.save(path)

    prof = read_docx(path)
    assert len(prof.equation_samples) == 2, "两个含公式的段落都应被采到"

    disp, inline = prof.equation_samples
    assert disp["display"] is True and disp["display_count"] == 1
    assert disp["alignment"] == "center"
    assert disp["size_pt"] == 12.0
    assert disp["space_before_pt"] == 0.0 and disp["space_after_pt"] == 0.0
    assert disp["line_spacing"] == 2.0
    assert disp["number_text"] == "(1)"
    assert disp["number_at_right"] is True
    assert disp["in_table"] is False

    assert inline["display"] is False and inline["inline_count"] == 1
    assert inline["alignment"] is None


def test_reader_reports_missing_right_tab_stop(tmp_path):
    """有 Tab 但缺右对齐制表位 → number_at_right=False，且 number_by_tab=True（可修）。"""
    doc = _new_doc()
    _add_eq_para(doc, [_disp(_eq_body(20), "center"),
                       '<w:r %s><w:tab/></w:r>' % nsdecls("w"),
                       '<w:r %s><w:t>(2)</w:t></w:r>' % nsdecls("w")])
    path = str(tmp_path / "eq_notab.docx")
    doc.save(path)

    s = read_docx(path).equation_samples[0]
    assert s["number_text"] == "(2)"
    assert s["number_by_tab"] is True
    assert s["number_at_right"] is False


def test_reader_handles_equations_inside_tables(tmp_path):
    """「公式居中 + 编号右单元格」常写成无框表 —— 表格里的公式也必须采到。"""
    doc = _new_doc()
    tbl = doc.add_table(rows=1, cols=2)
    cell_p = tbl.cell(0, 0).paragraphs[0]
    cell_p._p.append(parse_xml(_disp(_eq_body(24), "center")))
    tbl.cell(0, 1).text = "(3)"
    path = str(tmp_path / "eq_table.docx")
    doc.save(path)

    samples = read_docx(path).equation_samples
    assert samples, "表格单元格里的公式被漏采了"
    assert samples[0]["display"] is True
    assert samples[0]["in_table"] is True


# ---------------------------------------------------------------------------
# 2. 模板批注解析
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("text,expect", [
    ("Equations must be centred and numbered consecutively on the right, e.g. (1).",
     {"equation_alignment": "center", "equation_number_position": "right"}),
    ("Equations: Times New Roman 12 pt, centred, no extra space before or after.",
     {"equation_alignment": "center", "equation_size_pt": 12.0,
      "equation_space_before_pt": 0.0, "equation_space_after_pt": 0.0}),
    ("Formulas should be left-aligned, single-spaced, 6 pt space before and 6 pt after.",
     {"equation_alignment": "left", "equation_line_spacing": 1.0,
      "equation_space_before_pt": 6.0, "equation_space_after_pt": 6.0}),
    ("Equation numbering: do not number equations.",
     {"equation_number_position": "none"}),
    ("Equation size: 10 pt, line spacing 1.5",
     {"equation_size_pt": 10.0, "equation_line_spacing": 1.5}),
    ("All equations are numbered flush right.",
     {"equation_number_position": "right"}),
])
def test_equation_comments_are_parsed(text, expect):
    got = parse_comments([{"id": 1, "text": text}])["equation"]
    for k, v in expect.items():
        assert got.get(k) == v, "%r 没解析出 %s（实际 %r）" % (text, k, got)
    assert not parse_comments([{"id": 1, "text": text}])["unparsed"], \
        "这条公式批注被当成「不认识」了：%r" % text


def test_equation_comment_does_not_leak_into_body_font():
    """🔴 回归：「Equation font: …」曾被切成 category=font，把**正文**字体改掉。

    `_split_parts` 原按「位置相差 <2」挡重复标记，挡不住「短标记嵌在长标记内部」：
    equations(0..14) 与 font(9..13) 同时命中 → 前半段成了孤零零的 "Equation"，
    后半段 "font: Cambria Math 12 pt" 落进正文通道 —— 一条公式批注毁一篇稿。
    """
    r = parse_comments([{"id": 1, "text": "Equation font: Cambria Math 12 pt"}])
    assert r["equation"].get("equation_font_family") == "Cambria Math"
    assert r["page"].get("font_family") is None, "公式字体串到正文了：%r" % r["page"]
    assert r["page"].get("font_size_pt") is None, "公式字号串到正文字号了"


def test_british_spelling_centred_is_recognised():
    """🔴 回归：英联邦模板写 "centred"，旧正则认不出（cent+re+**d** 尾部多一个 d）。

    这条不只影响公式 —— 正文/摘要/标题的「居中」判定全都用它。
    """
    for word in ("centered", "centred", "centre", "center", "centering"):
        r = parse_comments([{"id": 1, "text": "Body text: %s" % word}])
        assert r["page"].get("body_alignment") == "center", \
            "%r 没被识别成居中" % word


def test_spacing_value_is_not_mistaken_for_font_size():
    """「6 pt space before」不能被读成「字号 6 pt」——那会把公式缩成 6 磅。"""
    got = parse_comments([{"id": 1, "text":
                           "Formulas: left-aligned, 6 pt space before and 6 pt after."}])["equation"]
    assert "equation_size_pt" not in got, "间距被误读成字号：%r" % got
    assert got.get("equation_space_before_pt") == 6.0
    assert got.get("equation_space_after_pt") == 6.0


# ---------------------------------------------------------------------------
# 3. 贯通：批注 / 问卷 → TargetProfile，build_target 不丢字段
# ---------------------------------------------------------------------------
def _template_with_eq_comments(tmp_path, comments):
    src = str(tmp_path / "_tpl_src.docx")
    doc = _new_doc()
    for i, _txt in enumerate(comments, start=1):
        para = doc.add_paragraph("Anchor %d" % i)
        mark_comment(para, i)
    doc.save(src)
    out = str(tmp_path / "template.docx")
    inject_comments(src, out, [(i, t) for i, t in enumerate(comments, start=1)])
    return out


EQ_COMMENTS = [
    "Equations must be centred and numbered consecutively on the right, e.g. (1).",
    "Equation size: 12 pt, no extra space before or after, double-spaced.",
]


def test_template_comments_reach_target_profile(tmp_path):
    tpl = _template_with_eq_comments(tmp_path, EQ_COMMENTS)
    t = target_from_template_docx(read_docx(tpl), reference_style=None)
    assert t.equation_alignment == "center"
    assert t.equation_number_position == "right"
    assert t.equation_size_pt == 12.0
    assert t.equation_space_before_pt == 0.0
    assert t.equation_space_after_pt == 0.0
    assert t.equation_line_spacing == 2.0
    for k in ("equation_alignment", "equation_number_position", "equation_size_pt"):
        assert t.resolution.get(k), "%s 没登记来源" % k


def test_build_target_keeps_element_level_fields(tmp_path):
    """🔴 回归：``build_target``（CLI 与 GUI 的统一入口）曾把 ``table_lines`` 整条丢掉。

    ``target_from_template_docx`` 明明算出了 ``three_line``，经 ``_overlay_one`` 合并时
    却没搬运 —— 于是「学校模板要求三线表」在真实产品路径上从未生效过，只有直接调
    ``target_from_template_docx`` 的单测看得见。公式字段同属这一类，必须一起守住。
    """
    comments = EQ_COMMENTS + [
        "Tables: use three-line tables with no vertical lines.",
    ]
    tpl = _template_with_eq_comments(tmp_path, comments)

    direct = target_from_template_docx(read_docx(tpl), reference_style=None)
    assert direct.table_lines == "three_line"

    merged = build_target(spec_key=None, template=tpl)
    assert merged.table_lines == "three_line", "build_target 把 table_lines 丢了"
    assert merged.equation_alignment == "center", "build_target 把公式对齐丢了"
    assert merged.equation_number_position == "right"
    assert merged.equation_size_pt == 12.0
    assert merged.resolution.get("table_lines") == "school-comment"


def test_questionnaire_normalises_and_rejects_enum_values():
    """枚举值：大小写不敏感地归一到白名单，不在白名单一律丢弃（不臆造值）。

    前者是好事（AI/手填常写成 "Center"/"RIGHT"）；后者是关键 —— 非法值若静默透传，
    checker 判等永远不成立，客户填了既不报错也不生效，比不填还糟。
    """
    # 大小写归一
    assert build_target(spec_key=None,
                        questionnaire={"equation_alignment": "Center"}
                        ).equation_alignment == "center"
    assert build_target(spec_key=None,
                        questionnaire={"equation_number_position": "RIGHT"}
                        ).equation_number_position == "right"
    # 非法值丢弃（同时留一个合法字段，否则「整份问卷一个有效字段都没有」会按既有
    # 约定抛 ValueError —— 那是另一条护栏，不在这里测）
    t = build_target(spec_key=None, questionnaire={
        "equation_size_pt": 12.0,
        "equation_number_position": "rightish",
        "equation_alignment": "justify",
    })
    assert t.equation_number_position is None
    assert t.equation_alignment is None
    assert t.equation_size_pt == 12.0, "合法字段被连带丢了"


# ---------------------------------------------------------------------------
# 4. 查 → 改 → 复检（幂等）
# ---------------------------------------------------------------------------
def test_fix_then_recheck_passes_and_is_idempotent(tmp_path):
    src = str(tmp_path / "in.docx")
    _build(compliant=False).save(src)
    target = build_target(spec_key=None, questionnaire=EQ_Q)

    before = _by_dim(_eq_findings(src, target))
    # 对齐 / 字号 / 段距 / 行距 / 编号位置 都应报出来，且**每条 warn 都绑了 FixAction**
    for dim in ("equation_alignment", "equation_size", "equation_space_before",
                "equation_space_after", "equation_line_spacing",
                "equation_number_position"):
        assert dim in before, "没检查出 %s" % dim
        warns = [f for f in before[dim] if f.severity == "warn"]
        assert warns, "%s 没报 warn" % dim
        assert all(f.fix is not None for f in warns), \
            "%s 的 warn 没绑 FixAction（报告说可修却修不动）" % dim

    out = str(tmp_path / "fixed.docx")
    summary = fix_docx(src, out, target)
    applied = summary.get("applied") or {}
    for k in ("equation_alignment", "equation_size_pt", "equation_space_before_pt",
              "equation_space_after_pt", "equation_line_spacing",
              "equation_number_position"):
        assert k in applied, "fixer 没实际动手：%s" % k

    after = _by_dim(_eq_findings(out, target))
    leftover = [f for fs in after.values() for f in fs if f.severity in ("warn", "fail")]
    assert not leftover, "修完复检仍有红杠：%r" % [(f.dimension, f.actual) for f in leftover]

    # 幂等：对已修好的文档再修一遍，公式维度必须零动作
    summary2 = fix_docx(out, str(tmp_path / "fixed2.docx"), target)
    applied2 = summary2.get("applied") or {}
    assert not [k for k in applied2 if k.startswith("equation")], \
        "第二遍还在改公式（非幂等）：%r" % applied2


def test_fix_never_touches_equation_text_or_number(tmp_path):
    """🔴 只改版式：公式内容（``m:t``）与编号文字一个字符都不许变。"""
    src = str(tmp_path / "in.docx")
    _build(compliant=False).save(src)
    target = build_target(spec_key=None, questionnaire=EQ_Q)
    out = str(tmp_path / "fixed.docx")
    fix_docx(src, out, target)

    def snapshot(path):
        doc = Document(path)
        texts, numbers = [], []
        for p in doc.paragraphs:
            el = p._p
            if not el.findall(qn("m:oMathPara")):
                continue
            texts.append("|".join(t.text or "" for t in el.findall(".//" + qn("m:t"))))
            numbers.append("".join(t.text or "" for t in el.findall(".//" + qn("w:t"))))
        return texts, numbers

    assert snapshot(src) == snapshot(out), "公式文本或编号被改动了"


def test_equation_number_is_reported_but_never_inserted(tmp_path):
    """补编号 = 往文档里加文字 → 只提示，绝不代写。"""
    src = str(tmp_path / "nonum.docx")
    _build(compliant=False, with_number=False).save(src)
    target = build_target(spec_key=None, questionnaire=EQ_Q)

    found = _by_dim(_eq_findings(src, target)).get("equation_number_position") or []
    assert found, "编号缺失没被报出来"
    assert all(f.severity == "info" for f in found), "编号缺失不该报 warn（修不了）"
    assert all(f.fix is None for f in found), "编号缺失不许绑 FixAction（那是在加文字）"

    out = str(tmp_path / "fixed.docx")
    fix_docx(src, out, target)
    # 修完仍然没有编号 —— 因为工具绝不插入文字
    assert not read_docx(out).equation_samples[0]["number_text"]


def test_document_without_equations_only_gets_info(tmp_path):
    """模板声明了公式规则、但论文里根本没公式 → **只出一条** info，不能平白多红杠。

    也不能每个维度刷一句「没找到公式」：模板写 6 条公式规则就刷 6 行同样的话，
    报告立刻变噪音墙。这里同时守住「合并成一条」与「不报 warn」两件事。
    """
    doc = _new_doc()
    doc.add_heading("Introduction", level=1)
    doc.add_paragraph("A paper with no equations at all.")
    path = str(tmp_path / "noeq.docx")
    doc.save(path)

    target = build_target(spec_key=None, questionnaire=EQ_Q)
    found = _eq_findings(path, target)
    assert len(found) == 1, "应当只出一条合并说明，实际 %d 条：%r" % (
        len(found), [(f.dimension, f.severity) for f in found])
    assert found[0].dimension == "equation_format"
    assert found[0].severity == "info", "没公式却报了 warn"
    assert "No equations were found" in found[0].actual


def test_no_equation_rules_means_no_findings(tmp_path):
    """来源没声明公式规则 → 公式维度整段不判（模板驱动铁律）。"""
    src = str(tmp_path / "in.docx")
    _build(compliant=False).save(src)
    target = build_target(spec_key=None, questionnaire={})
    assert not _eq_findings(src, target), "没声明规则却判了公式"
