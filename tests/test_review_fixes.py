"""round2 codex 审查 P0/P1 修复的回归测试。

守护点（对应 docs/codex_review_2026-09-11_round2.md 的修改清单）：
- P0-1 参考文献编号不再破坏 run 内容（内联图片 / 域代码 / 分页符 / 超链接条目错位）；
        并首次对参考文献段做「逐字一致」断言（仅允许新增 [n] 前缀）。
- P0-2 输出路径按 realpath/normcase/samefile 判定，大小写与相对路径变体不得覆盖原件。
- P1   标题保护覆盖 Title/伪标题/样式级 outlineLvl；不再改 Normal 字体名；
        字体采样排除标题段；rFonts eastAsia 兜底；
        LaTeX 剥离注释、引用维度来源标注正确；五通道优先级与引用维度冲突登记；
        参考文献 Finding 带 source；CLI preview --json 走字节流。
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from io import BytesIO

import pytest

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine import checker
from src.engine.docx_reader import read_docx, is_heading
from src.engine.fixer import compute_changes, fix_docx, preview_fix, same_file
from src.engine.latex_template import parse_latex_source
from src.engine.questionnaire import TargetProfile, build_target
from src.engine.specs import get_spec
from src import main as cli

# 1x1 透明 PNG（最小合法图，用于验证内联图片不被删除）
_PNG_1PX = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
    b"\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01"
    b"\x00\x00\x05\x00\x01\x0d\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82"
)

REF_TEXT_A = "Smith, J. (2020). Title of work. Publisher."
REF_TEXT_B = "Lee, K. (2019). Another title. Journal."


def _p_xml(para) -> str:
    return para._p.xml


# ---------------------------------------------------------------------------
# P0-1 参考文献编号：内容零改动
# ---------------------------------------------------------------------------

def _make_ref_doc(path, decorate_first_entry=None):
    doc = Document()
    doc.add_paragraph("Body paragraph for the thesis.")
    doc.add_paragraph("References")
    p1 = doc.add_paragraph()
    r1 = p1.add_run(REF_TEXT_A)
    if decorate_first_entry:
        decorate_first_entry(r1)
    doc.add_paragraph(REF_TEXT_B)
    doc.save(str(path))
    return str(path)


def _numbered_target():
    return TargetProfile(reference_style="IEEE", reference_hanging_indent_in=0.0,
                         reference_numbered=True)


def test_numbering_preserves_inline_image(tmp_path):
    def add_pic(run):
        run.add_picture(BytesIO(_PNG_1PX), width=Inches(0.3))

    src = _make_ref_doc(tmp_path / "img.docx", add_pic)
    out = tmp_path / "img_out.docx"
    fix_docx(src, str(out), _numbered_target())
    doc = Document(str(out))
    entry = doc.paragraphs[2]
    assert "w:drawing" in _p_xml(entry), "编号不得删除条目里的内联图片"
    assert entry.text.startswith("[1] ")
    assert entry.text[len("[1] "):] == REF_TEXT_A, "除新增前缀外必须逐字一致"


def test_numbering_preserves_page_break(tmp_path):
    def add_break(run):
        run.add_break(WD_BREAK.PAGE)

    src = _make_ref_doc(tmp_path / "br.docx", add_break)
    out = tmp_path / "br_out.docx"
    fix_docx(src, str(out), _numbered_target())
    doc = Document(str(out))
    xml = _p_xml(doc.paragraphs[2])
    assert 'w:type="page"' in xml, "编号不得删除分页符（连逐字比对都检测不到）"


def test_numbering_preserves_field_code(tmp_path):
    def add_field(run):
        # 模拟 EndNote/Zotero 域代码（真实文档里很常见）
        for tag, attrs, text in (
            ("w:fldChar", {"w:fldCharType": "begin"}, None),
            ("w:instrText", {}, " ADDIN EN.CITE "),
            ("w:fldChar", {"w:fldCharType": "end"}, None),
        ):
            el = OxmlElement(tag)
            for k, v in attrs.items():
                el.set(qn(k), v)
            if text is not None:
                el.text = text
            run._r.append(el)

    src = _make_ref_doc(tmp_path / "fld.docx", add_field)
    out = tmp_path / "fld_out.docx"
    fix_docx(src, str(out), _numbered_target())
    doc = Document(str(out))
    xml = _p_xml(doc.paragraphs[2])
    assert "w:fldChar" in xml and "w:instrText" in xml, "编号不得删除域代码"


def test_hyperlink_only_entry_prefix_at_start(tmp_path):
    doc = Document()
    doc.add_paragraph("Body paragraph for the thesis.")
    doc.add_paragraph("References")
    p = doc.add_paragraph()
    # 段落可见文字全部在 w:hyperlink 内 → python-docx 的 Paragraph.runs 为空
    link = OxmlElement("w:hyperlink")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "Smith, J. (2020). Title. https://doi.org/10.1000/x"
    r.append(t)
    link.append(r)
    p._p.append(link)
    src = tmp_path / "link.docx"
    doc.save(str(src))
    out = tmp_path / "link_out.docx"
    fix_docx(str(src), str(out), _numbered_target())
    entry = Document(str(out)).paragraphs[2]
    assert entry.text.startswith("[1] "), "超链接型条目的编号必须插在段首，不能跑到段末"
    assert "w:hyperlink" in _p_xml(entry), "超链接本身必须保留"


def test_reference_rest_text_identical(tmp_path):
    """逐字一致（本测试补上了此前从未覆盖的参考文献维度）。"""
    src = _make_ref_doc(tmp_path / "rest.docx")
    out = tmp_path / "rest_out.docx"
    before = [p.text for p in Document(src).paragraphs[2:4]]
    fix_docx(src, str(out), _numbered_target())
    after = [p.text for p in Document(str(out)).paragraphs[2:4]]
    for i, (b, a) in enumerate(zip(before, after), start=1):
        prefix = f"[{i}] "
        assert a.startswith(prefix)
        assert a[len(prefix):] == b, "除新增 [n] 前缀外，条目文字必须逐字一致"


def test_author_year_mode_never_adds_prefix(tmp_path):
    src = _make_ref_doc(tmp_path / "ay.docx")
    out = tmp_path / "ay_out.docx"
    before = [p.text for p in Document(src).paragraphs[2:4]]
    fix_docx(src, str(out), TargetProfile(reference_style="APA",
                                          reference_hanging_indent_in=0.5,
                                          reference_numbered=False))
    after = [p.text for p in Document(str(out)).paragraphs[2:4]]
    assert after == before, "非编号制不得改动任何字符"


def test_mixed_numbering_no_duplicate(tmp_path):
    doc = Document()
    doc.add_paragraph("Body paragraph for the thesis.")
    doc.add_paragraph("References")
    doc.add_paragraph("[3] Smith, J. (2020). Title of work. Publisher.")
    doc.add_paragraph("Lee, K. (2019). Another title. Journal.")
    src = tmp_path / "mix.docx"
    doc.save(str(src))
    out = tmp_path / "mix_out.docx"
    fix_docx(str(src), str(out), _numbered_target())
    texts = [p.text for p in Document(str(out)).paragraphs[2:4]]
    assert texts[0].startswith("[3] ")
    assert texts[1].startswith("[") and not texts[1].startswith("[3] "), \
        f"混合编号时不得重号，实际：{texts[1]}"


# ---------------------------------------------------------------------------
# P0-2 输出路径：不得以任何形式覆盖原件
# ---------------------------------------------------------------------------

def _tiny_doc(path):
    doc = Document()
    doc.add_paragraph("Body text.")
    doc.save(str(path))
    return str(path)


def test_fix_rejects_relative_variant(tmp_path):
    src = _tiny_doc(tmp_path / "in.docx")
    (tmp_path / "sub").mkdir(exist_ok=True)
    variant = str(tmp_path / "sub" / ".." / "in.docx")
    with pytest.raises(ValueError):
        fix_docx(src, variant, TargetProfile(page={"font_family": "Arial"}))


def test_fix_rejects_case_variant(tmp_path):
    src = _tiny_doc(tmp_path / "in.docx")
    with pytest.raises(ValueError):
        fix_docx(src, str(tmp_path / "IN.DOCX"),
                 TargetProfile(page={"font_family": "Arial"}))


def test_same_file_helper(tmp_path):
    a = _tiny_doc(tmp_path / "a.docx")
    assert same_file(a, str(tmp_path / "A.DOCX")) is True
    assert same_file(a, str(tmp_path / "b.docx")) is False


# ---------------------------------------------------------------------------
# P1 标题保护 / 字体设置范围
# ---------------------------------------------------------------------------

def test_title_style_is_heading(tmp_path):
    doc = Document()
    styles = [s.name for s in doc.styles]
    if "Title" not in styles:
        pytest.skip("当前 python-docx 默认模板无 Title 样式")
    p = doc.add_paragraph("My Thesis Main Title", style="Title")
    assert is_heading(p) is True, "Word 内置 Title 样式必须被识别为标题（否则主标题会被压成正文）"


def test_pseudo_heading_is_heading():
    doc = Document()
    p = doc.add_paragraph()
    r = p.add_run("Chapter One")
    r.bold = True
    r.font.size = Pt(20)
    assert is_heading(p) is True, "短、整段加粗、大字号的手写标题必须被保护"


def test_bold_body_sentence_is_not_heading():
    doc = Document()
    p = doc.add_paragraph()
    r = p.add_run("This is a bold body sentence that ends with a period.")
    r.bold = True
    r.font.size = Pt(12)     # 小字号 → 不是标题
    assert is_heading(p) is False


def test_style_level_outline_lvl_detected():
    doc = Document()
    st = doc.styles.add_style("MyCustomHeading", WD_STYLE_TYPE.PARAGRAPH)
    pPr = st.element.get_or_add_pPr()
    ol = OxmlElement("w:outlineLvl")
    ol.set(qn("w:val"), "0")
    pPr.append(ol)
    p = doc.add_paragraph("Custom Styled Heading", style="MyCustomHeading")
    assert is_heading(p) is True, "样式层 outlineLvl 必须被识别（自定义标题样式常这样定义）"


def test_fix_preserves_title_and_pseudo_heading(tmp_path):
    doc = Document()
    doc.add_paragraph("Body paragraph one.", style="Normal")
    styles = [s.name for s in doc.styles]
    heads = []
    if "Title" in styles:
        heads.append(doc.add_paragraph("Main Title", style="Title"))
    ph = doc.add_paragraph()
    pr = ph.add_run("Handmade Heading")
    pr.bold = True
    pr.font.size = Pt(22)
    heads.append(ph)
    src = tmp_path / "h.docx"
    doc.save(str(src))
    out = tmp_path / "h_out.docx"
    fix_docx(str(src), str(out),
             TargetProfile(page={"font_family": "Arial", "font_size_pt": 11,
                                 "line_spacing": 1.5}))
    after = Document(str(out))
    for hp in heads:
        for r in hp.runs:
            if r.text.strip():
                assert r.font.size != Pt(11), f"标题被压成正文：{hp.text}"
                assert r.font.name != "Arial", f"标题字体被改动：{hp.text}"


def test_fix_does_not_touch_normal_style_font(tmp_path):
    src = _tiny_doc(tmp_path / "n.docx")
    out = tmp_path / "n_out.docx"
    before = Document(src).styles["Normal"].font.name
    fix_docx(src, str(out), TargetProfile(page={"font_family": "Arial"}))
    after = Document(str(out)).styles["Normal"].font.name
    assert after == before, "不得改 Normal 样式字体名（会外溢到标题/表格/页眉）"
    # 正文 run 仍被设为目标字体
    assert Document(str(out)).paragraphs[0].runs[0].font.name == "Arial"


def test_font_sampling_excludes_headings(tmp_path):
    doc = Document()
    doc.add_paragraph("Chapter One", style="Heading 1")
    for t in ("Body one.", "Body two.", "Body three."):
        p = doc.add_paragraph()
        r = p.add_run(t)
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
    src = tmp_path / "s.docx"
    doc.save(str(src))
    prof = read_docx(str(src))
    assert prof.font_family == "Times New Roman", "标题段不应参与「正文字体」统计"


def test_east_asia_font_read(tmp_path):
    doc = Document()
    p = doc.add_paragraph()
    r = p.add_run("正文中文字体测试。")
    rPr = r._r.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:eastAsia"), "SimSun")
    rPr.insert(0, rFonts)
    src = tmp_path / "ea.docx"
    doc.save(str(src))
    prof = read_docx(str(src))
    assert prof.font_family == "SimSun", "只设 eastAsia 的 run 不应读到 None"


# ---------------------------------------------------------------------------
# P1 LaTeX / 合并 / 检查
# ---------------------------------------------------------------------------

def test_latex_comments_are_stripped():
    src = "% \\doublespacing\n\\documentclass{article}\n% \\bibliographystyle{IEEEtran}\n"
    t = parse_latex_source(src)
    assert "line_spacing" not in t.page, "被注释掉的行距声明不得采信"
    assert t.reference_style is None, "被注释掉的引用风格不得采信"


def test_latex_escaped_percent_kept():
    # \% 是转义的百分号，不是注释
    src = "\\documentclass[12pt]{article}  % 12pt 正文\n"
    t = parse_latex_source(src)
    assert t.page.get("font_size_pt") == 12.0


def test_latex_reference_resolution_label(tmp_path):
    f = tmp_path / "s.cls"
    f.write_text("\\bibliographystyle{IEEEtran}", encoding="utf-8")
    t = build_target("APA", latex=str(f))
    assert t.resolution.get("reference_style") == "latex_template", \
        "LaTeX 提供的引用维度不得被标成「学校模板」"


def test_channel_priority_template_beats_latex(tmp_path):
    tpl = Document()
    for s in tpl.sections:
        s.left_margin = Inches(1.5)
    tpl.add_paragraph("Body.")
    tpl_path = tmp_path / "tpl.docx"
    tpl.save(str(tpl_path))
    tex = tmp_path / "t.tex"
    tex.write_text("\\usepackage[margin=1in]{geometry}", encoding="utf-8")
    t = build_target("APA", template=str(tpl_path), latex=str(tex))
    assert t.page.get("margin_left_in") == pytest.approx(1.5), \
        "FR2 规定 template 优先级高于 latex"


def test_reference_conflict_registered(tmp_path):
    ai = tmp_path / "ai.json"
    ai.write_text(json.dumps({"margin_top_in": 1.0, "reference_style": "IEEE"}),
                  encoding="utf-8")
    tex = tmp_path / "t.tex"
    tex.write_text("\\bibliographystyle{apacite}", encoding="utf-8")
    t = build_target("APA", ai=str(ai), latex=str(tex))
    dims = {c["dim"] for c in t.conflicts}
    assert "reference_style" in dims, "引用维度的来源冲突必须登记（此前会静默丢失）"


def test_reference_finding_has_source(tmp_path):
    src = _tiny_doc(tmp_path / "r.docx")
    prof, findings = checker.run_check(src, TargetProfile(reference_style="APA"))
    ref = [f for f in findings if f.dimension == "reference"][0]
    assert ref.source == "spec", "参考文献结论必须标注来源（FR4）"


def test_margin_tolerance_warn_band(tmp_path):
    from src.engine.checker import MARGIN_TOL_FAIL, MARGIN_TOL_WARN
    assert MARGIN_TOL_WARN < MARGIN_TOL_FAIL
    assert checker._cmp_margin(1.03, 1.0)[0] == "pass"
    assert checker._cmp_margin(1.07, 1.0)[0] == "warn"
    assert checker._cmp_margin(1.30, 1.0)[0] == "fail"


# ---------------------------------------------------------------------------
# P1 入口层：已合规不消耗试用 / JSON 走字节流 / 不崩
# ---------------------------------------------------------------------------

def test_compute_changes_empty_when_conforming(tmp_path):
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)
    p = doc.add_paragraph()
    r = p.add_run("Body text.")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    p.paragraph_format.line_spacing = 2.0
    src = tmp_path / "ok.docx"
    doc.save(str(src))
    target = TargetProfile(page={
        "margin_top_in": 1.0, "margin_bottom_in": 1.0,
        "margin_left_in": 1.0, "margin_right_in": 1.0,
        "font_family": "Times New Roman", "font_size_pt": 12.0,
        "line_spacing": 2.0, "body_alignment": "left",
        "first_line_indent_in": 0.0, "space_after_pt": 0.0,
    })
    assert compute_changes(str(src), target) == [], "已合规文档不应报出任何修正项"


def test_preview_reference_line_only_when_needed(tmp_path):
    src = _make_ref_doc(tmp_path / "pv.docx")
    # 悬挂缩进已符合 + 非编号制 → 不应出现参考文献行
    joined = "\n".join(preview_fix(src, TargetProfile(
        reference_style="APA", reference_hanging_indent_in=0.0,
        reference_numbered=False)))
    assert "参考文献" not in joined, f"已符合的参考文献项不应列出：{joined}"
    # 编号制且条目未编号 → 应列出编号改动
    joined2 = "\n".join(preview_fix(src, _numbered_target()))
    assert "参考文献编号" in joined2


def _fix_args(docx, **kw):
    base = dict(docx=docx, spec="APA", preview=True, json=True, out=None,
                template=None, ai=None, questionnaire=None, latex=None)
    base.update(kw)
    return argparse.Namespace(**base)


def test_preview_json_uses_byte_writer(tmp_path, monkeypatch):
    """回归 B13：preview --json 必须走 _write_json（UTF-8 字节流），不能用 print。"""
    src = tmp_path / "p.docx"
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(2.0)
    doc.add_paragraph("Body.")
    doc.save(str(src))
    captured = []
    monkeypatch.setattr(cli, "_write_json", lambda text: captured.append(text))
    from src.license import license as lic
    monkeypatch.setattr(lic, "require_fix_entitlement",
                        lambda: {"allowed": True, "reason": "trial",
                                 "message": "首次免费试用", "remaining_trial": 1})
    rc = cli.cmd_fix(_fix_args(str(src)))
    assert rc == 0
    assert captured, "preview --json 分支必须调用 _write_json"
    payload = json.loads(captured[0])
    assert payload["mode"] == "preview"
    assert payload["changes"], "应列出会变的项"


def test_cli_preview_json_utf8_bytes(tmp_path):
    src = tmp_path / "p2.docx"
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(2.0)
    doc.add_paragraph("Body.")
    doc.save(str(src))
    env = os.environ.copy()
    # 追加 ROOT 而不是覆盖 PYTHONPATH：覆盖会把父进程的 site-packages 一起丢掉，
    # 用系统 Python 跑测试时就找不到 python-docx 了。
    env["PYTHONPATH"] = os.pathsep.join(
        p for p in (ROOT, os.environ.get("PYTHONPATH", "")) if p)
    env["TFD_LICENSE_FILE"] = str(tmp_path / "lic.json")
    # 隔离中台：本测试只关心 stdout 是不是合法 UTF-8 字节流，不该依赖服务器状态
    # （此前跑真实指纹会命中"该机器试用已用"，rc=3 假红）。
    env["TFD_FORCE_OFFLINE"] = "1"
    r = subprocess.run(
        [sys.executable, "-m", "src.main", "fix", str(src), "--spec", "APA",
         "--preview", "--json"],
        cwd=ROOT, env=env, capture_output=True)
    assert r.returncode == 0, r.stderr.decode("utf-8", "replace")
    decoded = r.stdout.decode("utf-8")           # 能解码即证明是合法 UTF-8
    payload = json.loads(decoded)
    assert payload["mode"] == "preview"
    assert "→".encode("utf-8") in r.stdout, "箭头等非 GBK 字符不得被降级成 ?"


def test_cli_noop_does_not_consume_trial(tmp_path, monkeypatch):
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)
    p = doc.add_paragraph()
    r = p.add_run("Body text.")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    p.paragraph_format.line_spacing = 2.0
    src = tmp_path / "noop.docx"
    doc.save(str(src))

    called = []
    from src.license import license as lic
    monkeypatch.setattr(lic, "require_fix_entitlement",
                        lambda: {"allowed": True, "reason": "trial",
                                 "message": "首次免费试用", "remaining_trial": 1})
    monkeypatch.setattr(lic, "record_fix_used", lambda: called.append(1) or True)

    # 用一份「文档已完全符合」的目标画像，验证 CLI 不会落盘、不会消耗试用
    conforming = TargetProfile(page={
        "margin_top_in": 1.0, "margin_bottom_in": 1.0,
        "margin_left_in": 1.0, "margin_right_in": 1.0,
        "font_family": "Times New Roman", "font_size_pt": 12.0,
        "line_spacing": 2.0, "body_alignment": "left",
        "first_line_indent_in": 0.0, "space_after_pt": 0.0,
    })
    monkeypatch.setattr(cli, "_build_target", lambda a: conforming)
    captured = []
    monkeypatch.setattr(cli, "_write_json", lambda text: captured.append(text))
    rc = cli.cmd_fix(_fix_args(str(src), preview=False, json=True))
    assert rc == 0
    assert json.loads(captured[0])["mode"] == "noop"
    assert not called, "已合规文档不得消耗试用额度"
    assert not os.path.exists(str(tmp_path / "noop_fixed.docx")), "不得生成副本"


# ---------------------------------------------------------------------------
# 端到端冒烟暴露的三个问题（入参校验顺序 / 报告章节号 / 数值重复前缀）
# ---------------------------------------------------------------------------

def test_out_same_file_rejected_before_gate(tmp_path, monkeypatch):
    """`--out` 指向原件属**参数错误**，必须在授权门禁之前返回 rc=2。

    回归：此前门禁在前，试用用尽的用户会被报成「免费试用已用完」（rc=3），
    真正的问题（输出路径 = 原件）被掩盖，更糟的是参数校验根本不可达。
    """
    src = tmp_path / "sample.docx"
    doc = Document()
    doc.add_paragraph("Body.")
    doc.save(str(src))

    asked = []
    from src.license import license as lic
    monkeypatch.setattr(lic, "require_fix_entitlement",
                        lambda: asked.append(1) or {
                            "allowed": False, "reason": "trial_exhausted",
                            "message": "免费试用已用完", "remaining_trial": 0})

    # 大小写变体（Windows/macOS 下同一文件）
    before = src.read_bytes()
    rc = cli.cmd_fix(_fix_args(str(src), preview=False, json=False,
                               out=str(tmp_path / "SAMPLE.DOCX")))
    assert rc == 2, "参数错误应返回 rc=2，而不是门禁的 rc=3"
    assert not asked, "入参校验必须先于授权门禁，不得先问授权"
    # 注意：Windows 上 SAMPLE.DOCX 与 sample.docx 是同一个文件，
    # 所以这里不能断言「不存在」，只能断言「字节未变」。
    assert src.read_bytes() == before, "原件不得被改动"


def test_out_same_file_rejected_without_consuming_trial(tmp_path, monkeypatch):
    """试用未用 + 输出=原件 → rc=2 且试用额度不得被消耗。"""
    src = tmp_path / "s2.docx"
    doc = Document()
    doc.add_paragraph("Body.")
    doc.save(str(src))
    called = []
    from src.license import license as lic
    monkeypatch.setattr(lic, "record_fix_used", lambda: called.append(1) or True)
    rc = cli.cmd_fix(_fix_args(str(src), preview=False, json=False,
                               out=str(src)))
    assert rc == 2
    assert not called


def test_report_section_numbers_are_dense(tmp_path):
    """无「提示」无「冲突」时，章节号必须连续（一、二），不得出现「一之二」或跳号。"""
    from src.engine.docx_reader import read_docx
    from src.engine.report import build_report
    from src.engine.checker import run_check

    src = tmp_path / "r.docx"
    doc = Document()
    doc.add_paragraph("Body text.")
    doc.save(str(src))

    target = TargetProfile(page={"margin_top_in": 1.0})   # 无 notes
    prof = read_docx(str(src))
    _prof2, findings = run_check(str(src), target)
    md = build_report(str(src), target, prof, findings).markdown

    heads = [ln for ln in md.splitlines() if ln.startswith("## ")]
    assert heads[0].startswith("## 一、"), heads
    assert heads[1].startswith("## 二、"), heads
    assert "一之二" not in md
    assert "## 三、" not in md, "只有两段时不应出现三"


def test_report_section_numbers_with_notes(tmp_path):
    """有 notes（例如 LaTeX 空结果提示）时，编号顺延且冲突段继续顺延。"""
    from src.engine.docx_reader import read_docx
    from src.engine.report import build_report
    from src.engine.checker import run_check

    src = tmp_path / "r2.docx"
    doc = Document()
    doc.add_paragraph("Body text.")
    doc.save(str(src))

    target = TargetProfile(page={"margin_top_in": 1.0},
                           notes=["LaTeX 未解析出可用的页面维度"])
    prof = read_docx(str(src))
    _p, fs = run_check(str(src), target)
    md = build_report(str(src), target, prof, fs).markdown
    heads = [ln for ln in md.splitlines() if ln.startswith("## ")]
    assert heads[0].startswith("## 一、")
    assert heads[1].startswith("## 二、提示"), heads
    assert heads[2].startswith("## 三、"), heads
    assert "一之二" not in md


def test_margin_actual_has_no_duplicate_prefix():
    """回归「实测：实测 2.0"」：checker 只给数值，拼「实测：」是报告层的活。"""
    for measured, exp in ((1.0, 1.0), (1.07, 1.0), (2.0, 1.0)):
        _sev, actual = checker._cmp_margin(measured, exp)
        assert "实测" not in actual, actual
        assert actual.startswith(f'{measured}"'), actual


def test_report_actual_column_not_duplicated(tmp_path):
    """报告渲染后「实测」只出现一次。"""
    from src.engine.docx_reader import read_docx
    from src.engine.report import build_report
    from src.engine.checker import run_check

    src = tmp_path / "r3.docx"
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(2.0)
    doc.add_paragraph("Body text.")
    doc.save(str(src))

    target = TargetProfile(page={"margin_top_in": 1.0})
    prof = read_docx(str(src))
    _p, fs = run_check(str(src), target)
    md = build_report(str(src), target, prof, fs).markdown
    assert "实测：实测" not in md
    assert "实测：2.0" in md, md


def test_post_fix_message_reflects_consumed_trial():
    """修正成功后展示的是**扣减后**的状态，不再出现「刚用完却显示首次免费」。"""
    from src.license import license as lic

    trial_gate = {"reason": "trial", "message": "首次免费试用（共 1 次）"}
    msg = lic.post_fix_message(trial_gate, True)
    assert "已用尽" in msg and "激活码" in msg

    act_gate = {"reason": "activated", "kind": "lifetime"}
    assert "已激活" in lic.post_fix_message(act_gate, True)

    assert "写入失败" in lic.post_fix_message(trial_gate, False)


def test_cli_fix_prints_post_state(tmp_path, monkeypatch, capsys):
    """CLI 成功输出里必须出现扣减后的试用状态，且不再显示「首次免费试用」。"""
    src = tmp_path / "s3.docx"
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(2.0)
    doc.add_paragraph("Body.")
    doc.save(str(src))

    from src.license import license as lic
    monkeypatch.setattr(lic, "require_fix_entitlement",
                        lambda: {"allowed": True, "reason": "trial",
                                 "message": "首次免费试用（共 1 次）",
                                 "remaining_trial": 1})
    monkeypatch.setattr(lic, "record_fix_used", lambda: True)

    rc = cli.cmd_fix(_fix_args(str(src), preview=False, json=False))
    assert rc == 0
    out = capsys.readouterr().out
    assert "已用尽" in out, out
    assert "首次免费试用" not in out, out
    assert os.path.exists(str(tmp_path / "s3_fixed.docx"))

