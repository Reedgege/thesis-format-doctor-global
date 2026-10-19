"""每级标题参数的端到端护栏（v2.3.65）。

老板原话：「标题层级，各层级的字体字号缩进斜体等等参数你做了吗？ai 导入与手动输入
一样，只是让 ai 帮客户写，你要真正理解」。

这里钉死三件事：
1. **模板/问卷/AI 显式给出的每级参数，必须真的被修正器用上**（历史上被写死的
   「L1 居中加粗」通用规则覆盖，模板驱动铁律形同虚设）；
2. **检查器要看得见每级差异**，且修完再查必须变 pass（check→fix 单一真相源）；
3. **手填与 AI 是同一套字段**：AI 模板里能写的字段，UI 里都能填，反之亦然。
"""
from __future__ import annotations

import os

import pytest
from docx import Document
from docx.shared import Inches, Pt

from src.engine.checker import run_check
from src.engine.fixer import fix_docx
from src.engine.questionnaire import (HEADING_LEVEL_KEYS, QUESTIONNAIRE_SCHEMA,
                                      target_from_dict)
from src.engine.specs import (APA_HEADING, generic_heading_formats,
                              heading_level_format_from_dict,
                              resolve_heading_formats)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_YAML = os.path.join(ROOT, "src", "data", "ai_questionnaire_template.yaml")


# ---------------------------------------------------------------- 解析层
def test_explicit_heading_formats_beat_spec_and_generic():
    """显式来源 > 内置规范 > 通用兜底（顺序错了就等于没做模板驱动）。"""
    explicit = {1: {"alignment": "left", "font_family": "SimHei", "size_pt": 16}}
    got = resolve_heading_formats(explicit, APA_HEADING, 5)
    assert got[1].alignment == "left"          # 模板说左对齐，不是 APA 的居中
    assert got[1].font_family == "SimHei"
    # 有 explicit 就不许拿规范去补它没写的层级（否则等于偷偷换成 APA）
    assert 2 not in got


def test_no_explicit_no_spec_falls_back_to_generic():
    got = resolve_heading_formats({}, {}, 3)
    assert sorted(got) == [1, 2, 3]
    assert got == generic_heading_formats(3)


def test_heading_format_from_dict_keeps_tri_state():
    f = heading_level_format_from_dict({"alignment": "left", "size_pt": "14"})
    assert f.bold is None and f.italic is None     # 没写 = 不规定，不是 False
    f2 = heading_level_format_from_dict({"bold": False, "italic": "true"})
    assert f2.bold is False and f2.italic is True
    # 非法值丢弃，不猜
    f3 = heading_level_format_from_dict({"alignment": "sideways",
                                         "number_prefix": "wat"})
    assert f3.alignment is None and f3.number_prefix is None


def test_heading_desc_is_english_only():
    f = heading_level_format_from_dict({"alignment": "left", "font_family": "SimHei",
                                        "size_pt": 16, "bold": True, "indent_in": 0.5})
    assert f.desc, "必须有说明（会写进文档批注与报告）"
    assert all(ord(ch) < 128 for ch in f.desc), f.desc


# ---------------------------------------------------------------- 修正层
def _doc_with_headings(tmp_path, name="in.docx"):
    src = tmp_path / name
    doc = Document()
    for lvl, text in ((1, "Chapter One"), (2, "Section Two")):
        p = doc.add_paragraph()
        p.style = doc.styles[f"Heading {lvl}"]
        p.add_run(text)
    doc.add_paragraph("Body text one.")
    doc.save(src)
    return src


def test_fixer_applies_per_level_font_size_indent_and_italic(tmp_path):
    src = _doc_with_headings(tmp_path)
    out = tmp_path / "out.docx"
    target = target_from_dict({
        "heading_levels": 2,
        "heading_formats": {
            "1": {"alignment": "left", "bold": True, "font_family": "SimHei",
                  "size_pt": 16},
            "2": {"alignment": "left", "italic": True, "font_family": "KaiTi",
                  "size_pt": 14, "indent_in": 0.25},
        },
    })
    fix_docx(str(src), str(out), target)
    got = {p.text.strip(): p for p in Document(str(out)).paragraphs if p.text.strip()}
    h1, h2 = got["Chapter One"], got["Section Two"]
    assert h1.runs[0].font.name == "SimHei"
    assert h1.runs[0].font.size == Pt(16)
    assert h1.alignment == 0                    # LEFT（不是通用兜底的居中）
    assert h2.runs[0].font.name == "KaiTi"
    assert h2.runs[0].font.size == Pt(14)
    assert h2.runs[0].font.italic is True
    assert h2.paragraph_format.first_line_indent == Inches(0.25)


def test_fixer_leaves_unspecified_bold_alone(tmp_path):
    """没写 bold 就一个字节都不动（学校没说的不臆造）。"""
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    doc = Document()
    p = doc.add_paragraph()
    p.style = doc.styles["Heading 1"]
    r = p.add_run("Keep My Bold")
    r.font.bold = False                         # 作者明确不加粗
    doc.save(src)

    target = target_from_dict({"heading_formats": {
        "1": {"alignment": "left", "font_family": "SimHei"}}})   # 未提 bold
    fix_docx(str(src), str(out), target)
    para = [p for p in Document(str(out)).paragraphs if p.text.strip()][0]
    assert para.runs[0].font.bold is False, "未声明的属性不该被改"


def test_fixer_explicitly_not_bold_strips_bold(tmp_path):
    """写了 bold: false 就必须真的不加粗（否则"各层级的加粗"等于没做）。"""
    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    doc = Document()
    p = doc.add_paragraph()
    p.style = doc.styles["Heading 1"]
    r = p.add_run("Should Lose Bold")
    r.font.bold = True
    doc.save(src)

    target = target_from_dict({"heading_formats": {
        "1": {"alignment": "left", "bold": False}}})
    fix_docx(str(src), str(out), target)
    para = [p for p in Document(str(out)).paragraphs if p.text.strip()][0]
    assert para.runs[0].font.bold is False


def test_fix_is_idempotent_for_per_level_formats(tmp_path):
    src = _doc_with_headings(tmp_path)
    out1 = tmp_path / "o1.docx"
    out2 = tmp_path / "o2.docx"
    target = target_from_dict({"heading_formats": {
        "1": {"alignment": "left", "font_family": "SimHei", "size_pt": 16, "bold": True},
        "2": {"alignment": "left", "font_family": "KaiTi", "size_pt": 14,
              "italic": True, "indent_in": 0.25},
    }})
    s1 = fix_docx(str(src), str(out1), target)
    s2 = fix_docx(str(out1), str(out2), target)
    assert s1["headings_formatted"] > 0
    assert s2["headings_formatted"] == 0, "第二遍不该再有改动（幂等）"


# ---------------------------------------------------------------- 检查层
def test_checker_reports_per_level_differences_then_passes(tmp_path):
    src = _doc_with_headings(tmp_path)
    target = target_from_dict({"heading_formats": {
        "1": {"alignment": "left", "font_family": "SimHei", "size_pt": 16},
        "2": {"alignment": "left", "font_family": "KaiTi", "size_pt": 14},
    }})
    _, findings = run_check(str(src), target)
    lv = [f for f in findings if f.message.startswith("Level ")
          and f.message.endswith("heading format")]
    assert len(lv) == 2, [f.message for f in findings]
    assert all(f.severity == "fail" for f in lv)
    assert all(f.fix is not None for f in lv), "必须可修（check→fix 单一真相源）"
    assert any("16" in f.expected for f in lv), [f.expected for f in lv]

    out = tmp_path / "fixed.docx"
    fix_docx(str(src), str(out), target, findings=findings)
    _, after = run_check(str(out), target)
    lv2 = [f for f in after if f.message.startswith("Level ")
           and f.message.endswith("heading format")]
    assert lv2 and all(f.severity == "pass" for f in lv2), \
        [(f.message, f.severity, f.expected, f.actual) for f in lv2]


def test_school_template_heading_formats_drive_the_fixer(tmp_path):
    """学校模板里的每级格式必须真的驱动修正（模板驱动铁律的端到端实测）。

    历史 bug：``fmt_map = spec.heading_formats or generic_heading_formats(...)`` 完全
    无视 ``target.heading_formats`` —— 学校模板写着「一级左对齐黑体16pt」，修正器却按
    写死的「L1 居中加粗」去改。这里用一份"像学校模板"的 docx 把它钉死。
    """
    from src.engine.questionnaire import build_target

    tpl = tmp_path / "school_template.docx"
    doc = Document()
    p = doc.add_paragraph()
    p.style = doc.styles["Heading 1"]
    r = p.add_run("Chapter Title")
    r.font.name = "SimHei"
    r.font.size = Pt(16)
    r.font.bold = True
    p.alignment = 0                                     # LEFT
    doc.add_paragraph("Body text of the template.")
    doc.save(tpl)

    target = build_target(None, template=str(tpl))
    assert target.heading_formats, "学校模板的每级格式没被读进目标画像"
    assert target.heading_formats[1].get("font_family") == "SimHei"
    assert target.heading_formats[1].get("size_pt") == 16.0

    src = tmp_path / "student.docx"
    sd = Document()
    sp = sd.add_paragraph()
    sp.style = sd.styles["Heading 1"]
    sp.add_run("Student Chapter")
    sd.add_paragraph("Student body text.")
    sd.save(src)

    out = tmp_path / "student_fixed.docx"
    fix_docx(str(src), str(out), target)
    head = [x for x in Document(str(out)).paragraphs if x.text.strip()][0]
    assert head.runs[0].font.name == "SimHei", "没按学校模板的字体改"
    assert head.runs[0].font.size == Pt(16), "没按学校模板的字号改"
    assert head.alignment == 0, "没按学校模板的对齐改（被通用兜底的居中覆盖了）"


def test_preview_lists_heading_changes(tmp_path):
    """预览必须列出每级标题改动。

    ``compute_changes`` 是「要不要出副本 / 要不要扣次」的唯一依据：只改每级标题时
    若这里为空，App 会判「已符合」→ 根本不跑修正（客户填了参数却什么都没发生）。
    """
    from src.engine.fixer import compute_changes

    src = _doc_with_headings(tmp_path)
    target = target_from_dict({"heading_formats": {
        "1": {"alignment": "left", "font_family": "SimHei", "size_pt": 16},
        "2": {"alignment": "left", "font_family": "KaiTi", "size_pt": 14},
    }})
    en = compute_changes(str(src), target, "en")
    zh = compute_changes(str(src), target, "zh")
    assert any(l.startswith("Level 1 heading") for l in en), en
    assert any(l.startswith("Level 2 heading") for l in en), en
    assert any("16pt" in l for l in en), en
    assert any(l.startswith("第 1 级标题") for l in zh), zh
    # 英文界面绝不漏中文（海外版铁律）
    assert all(not any("\u4e00" <= ch <= "\u9fff" for ch in l) for l in en), en


def test_checker_skips_levels_missing_in_document(tmp_path):
    """文档里没有这一级标题时，不凭空判缺失。"""
    src = tmp_path / "one.docx"
    doc = Document()
    p = doc.add_paragraph()
    p.style = doc.styles["Heading 1"]
    p.add_run("Only Level One")
    doc.save(src)
    target = target_from_dict({"heading_formats": {
        "1": {"alignment": "left"},
        "2": {"alignment": "left", "font_family": "KaiTi"},
    }})
    _, findings = run_check(str(src), target)
    msgs = [f.message for f in findings if f.message.endswith("heading format")]
    assert msgs == ["Level 1 heading format"], msgs


# ---------------------------------------------------------------- 两个入口等价
def test_heading_levels_derived_from_heading_formats():
    """只填每级明细、没填层级数 → 层级数取最深一级（不让客户重复填）。"""
    t = target_from_dict({"heading_formats": {
        "1": {"alignment": "left"}, "2": {"italic": True}, "3": {"bold": True}}})
    assert t.heading_levels == 3


def test_ai_template_example_lands_in_heading_formats():
    yaml = pytest.importorskip("yaml")
    with open(TEMPLATE_YAML, encoding="utf-8") as fh:
        tpl = yaml.safe_load(fh)
    t = target_from_dict(tpl["example_filled"])
    assert t.heading_formats, "AI 模板示例必须能被引擎吃成每级标题规则"
    used = set()
    for lvl in t.heading_formats.values():
        used |= set(lvl)
    assert {"font_family", "size_pt", "alignment"} <= used
    assert t.heading_levels == max(t.heading_formats)


def test_ai_template_documents_every_heading_field():
    """AI 模板的说明里必须点到每个字段 —— 否则 AI 根本不会填它。"""
    text = open(TEMPLATE_YAML, encoding="utf-8").read()
    for key in HEADING_LEVEL_KEYS:
        assert key in text, "AI 模板缺少字段说明：%s" % key
    keys = [f["key"] for f in (yaml_load(TEMPLATE_YAML).get("fields") or [])]
    assert "heading_formats" in keys
    assert "heading_levels" in keys


def test_questionnaire_schema_still_documents_heading_levels():
    keys = [f["key"] for f in QUESTIONNAIRE_SCHEMA]
    assert "heading_levels" in keys


def yaml_load(path):
    import yaml
    with open(path, encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}
