"""栏数（单栏 / 双栏）全链路测试：读 → 目标 → 查 → 改 → 预览。

背景：python-docx 1.2.0 的 ``Section`` 没有栏数属性（实测 ``dir(Section)`` 无
``cols_*``），只能直读 / 直写 ``sectPr/w:cols/@w:num``。本文件用真实 docx 造数据，
覆盖「默认单栏 / 双栏识别 / 多分节取众数 / 规范路由（IEEE=2、APA/MLA=1、Other=None）/
模板与问卷来源 / checker 判定与 FixAction / fixer 落盘 / 预览行 / 复核弹窗只改动才升级」。
"""

from __future__ import annotations

import os
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine.docx_reader import read_docx
from src.engine.questionnaire import (
    build_target, target_from_dict, target_from_spec, target_from_ai,
    target_from_template_docx, apply_page_overrides, TargetProfile,
)
from src.engine.checker import run_check
from src.engine.fixer import fix_docx, compute_changes


# --------------------------------------------------------------------- helpers
def _make_paper(path: str) -> str:
    doc = Document()
    doc.add_paragraph("A Study of Column Layout")
    for _ in range(3):
        doc.add_paragraph("This is a body paragraph with enough words to look like real "
                          "content in a student paper, repeated a few times.")
    doc.add_paragraph("References")
    doc.add_paragraph("[1] J. Smith, \"A title,\" J. Name, vol. 1, 2020.")
    doc.save(path)
    return path


def _set_cols(path: str, n: int, extra_sections: int = 0) -> str:
    """把（首个）分节设为 n 栏；extra_sections>0 时追加分节并读回做众数场景。"""
    doc = Document(path)
    if extra_sections:
        doc.add_paragraph("section break")
        doc.add_section()
    for sec in doc.sections:
        cols = sec._sectPr.find(qn("w:cols"))
        if cols is None:
            cols = OxmlElement("w:cols")
            docGrid = sec._sectPr.find(qn("w:docGrid"))
            if docGrid is not None:
                docGrid.addprevious(cols)
            else:
                sec._sectPr.append(cols)
        if n == 1:
            cols.attrib.pop(qn("w:num"), None)
        else:
            cols.set(qn("w:num"), str(n))
    doc.save(path)
    return path


def _cols_of(path: str) -> list:
    doc = Document(path)
    out = []
    for sec in doc.sections:
        c = sec._sectPr.find(qn("w:cols"))
        raw = None if c is None else c.get(qn("w:num"))
        out.append(int(raw) if raw else 1)
    return out


# ------------------------------------------------------------------- 读取层
def test_reader_defaults_to_single_column(tmp_path):
    p = _make_paper(str(tmp_path / "single.docx"))
    assert read_docx(p).columns == 1


def test_reader_detects_two_columns(tmp_path):
    p = _set_cols(_make_paper(str(tmp_path / "two.docx")), 2)
    assert read_docx(p).columns == 2
    assert _cols_of(p) == [2]


def test_reader_takes_mode_across_sections(tmp_path):
    """三分节 [2, 2, 1] → 众数为 2（与页边距同口径）。"""
    p = _make_paper(str(tmp_path / "multi.docx"))
    doc = Document(p)
    doc.add_paragraph("p2")
    doc.add_section()
    doc.add_paragraph("p3")
    doc.add_section()
    doc.add_paragraph("p4")
    for i, sec in enumerate(doc.sections):
        cols = sec._sectPr.find(qn("w:cols"))
        if cols is None:
            cols = OxmlElement("w:cols")
            sec._sectPr.append(cols)
        if i < 2:
            cols.set(qn("w:num"), "2")
        else:
            cols.attrib.pop(qn("w:num"), None)
    doc.save(p)
    assert _cols_of(p) == [2, 2, 1]
    assert read_docx(p).columns == 2


# ------------------------------------------------------------------- 目标层
def test_target_from_spec_columns():
    assert target_from_spec("IEEE").columns == 2
    assert target_from_spec("APA-Student").columns == 1
    assert target_from_spec("APA-Professional").columns == 1
    assert target_from_spec("MLA").columns == 1
    assert target_from_spec("Other").columns is None


def test_target_from_spec_columns_resolution():
    t = target_from_spec("IEEE")
    assert t.resolution.get("columns") == "spec"


def test_target_from_template_reads_columns(tmp_path):
    p = _set_cols(_make_paper(str(tmp_path / "tpl.docx")), 2)
    prof = read_docx(p)
    t = target_from_template_docx(prof)
    assert t.columns == 2
    assert t.resolution.get("columns") == "school"
    assert t.spec_key is None  # 无规范基线模式不变


def test_target_from_dict_reads_columns():
    t = target_from_dict({"columns": 2, "reference_style": "IEEE"})
    assert t.columns == 2
    assert t.resolution.get("columns") == "questionnaire"
    # 只给栏数也应被视为有效输入（不触发空画像报错）
    t2 = target_from_dict({"columns": 1})
    assert t2.columns == 1


def test_build_target_ieee_is_two_column():
    assert build_target("IEEE").columns == 2


def test_build_target_school_columns_from_template(tmp_path):
    p = _set_cols(_make_paper(str(tmp_path / "tpl2.docx")), 2)
    t = build_target(None, template=p)
    assert t.columns == 2
    assert t.resolution.get("columns") == "school"


# ------------------------------------------------------------------- 检查层
def test_checker_flags_single_column_for_ieee(tmp_path):
    p = _make_paper(str(tmp_path / "s.docx"))
    _, findings = run_check(p, build_target("IEEE"))
    col = [f for f in findings if f.dimension == "columns"]
    assert len(col) == 1
    assert col[0].severity == "warn"
    assert col[0].expected == "2 columns"
    assert col[0].actual == "1 column"
    # check→fix 单一真相源：必须绑定 columns 修正，fixer 才能落盘
    assert col[0].fix is not None and col[0].fix.dimension == "columns"
    assert col[0].fix.value == 2


def test_checker_passes_two_column_for_ieee(tmp_path):
    p = _set_cols(_make_paper(str(tmp_path / "t.docx")), 2)
    _, findings = run_check(p, build_target("IEEE"))
    col = [f for f in findings if f.dimension == "columns"]
    assert len(col) == 1
    assert col[0].severity == "pass"


# ------------------------------------------------------------------- 修正层
def test_fixer_applies_two_columns(tmp_path):
    src = _make_paper(str(tmp_path / "in.docx"))
    out = str(tmp_path / "out.docx")
    target = build_target("IEEE")
    _, findings = run_check(src, target)
    res = fix_docx(src, out, target, findings=findings)
    assert _cols_of(out) == [2]
    assert res["applied"].get("columns") == 2
    assert read_docx(out).columns == 2


def test_fixer_reverts_to_single_column(tmp_path):
    src = _set_cols(_make_paper(str(tmp_path / "in2.docx")), 2)
    out = str(tmp_path / "out2.docx")
    target = target_from_spec("APA-Student")   # columns=1
    _, findings = run_check(src, target)
    res = fix_docx(src, out, target, findings=findings)
    assert _cols_of(out) == [1]
    assert res["applied"].get("columns") == 1


def test_fixer_columns_idempotent(tmp_path):
    """已双栏文件再修一遍：不应声称改了栏数（如实返回 changed=False）。"""
    src = _set_cols(_make_paper(str(tmp_path / "in3.docx")), 2)
    out = str(tmp_path / "out3.docx")
    target = build_target("IEEE")
    _, findings = run_check(src, target)
    res = fix_docx(src, out, target, findings=findings)
    assert "columns" not in res["applied"]


# ------------------------------------------------------------------- 预览层
def test_compute_changes_lists_columns(tmp_path):
    p = _make_paper(str(tmp_path / "c.docx"))
    lines = compute_changes(p, build_target("IEEE"))
    assert any(("Column layout" in ln or "栏数" in ln) for ln in lines)


def test_compute_changes_no_columns_when_already_two(tmp_path):
    p = _set_cols(_make_paper(str(tmp_path / "c2.docx")), 2)
    lines = compute_changes(p, build_target("IEEE"))
    assert not any(("Column layout" in ln or "栏数" in ln) for ln in lines)


# --------------------------------------------------- 复核弹窗：仅改动才升级
def test_apply_page_overrides_columns_only_when_changed():
    t = build_target("APA-Student")            # columns=1, resolution=spec
    assert t.columns == 1
    apply_page_overrides(t, {"columns": "1"})   # 未改动 → 保持 spec，不升级
    assert t.columns == 1
    assert t.resolution.get("columns") == "spec"
    apply_page_overrides(t, {"columns": "2"})   # 真改了 → 升级为 user
    assert t.columns == 2
    assert t.resolution.get("columns") == "user"


def test_apply_page_overrides_columns_rejects_garbage():
    t = TargetProfile()
    apply_page_overrides(t, {"columns": "abc"})
    assert t.columns is None
    apply_page_overrides(t, {"columns": "0"})   # <1 非法
    assert t.columns is None


# ------------------------------------------- 手工问卷 / AI 模板两条入口也带栏数
def test_ai_template_includes_columns():
    """导出的 AI 模板必须包含 columns 字段，否则 AI 填了也没处写。"""
    import yaml
    tpl = os.path.join(ROOT, "src", "data", "ai_questionnaire_template.yaml")
    with open(tpl, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    keys = [fld["key"] for fld in data.get("fields", [])]
    assert "columns" in keys
    assert data.get("example_filled", {}).get("columns") == 1


def test_ai_template_covers_all_schema_fields():
    """防漂移：问卷 schema 的每个字段都必须在 AI 模板里出现，否则 AI 填了也传不进来。

    （这条正是「加了字段却忘了改 AI 提示词/模板」的回归护栏。）
    """
    import yaml
    from src.engine.questionnaire import QUESTIONNAIRE_SCHEMA
    tpl = os.path.join(ROOT, "src", "data", "ai_questionnaire_template.yaml")
    with open(tpl, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    tpl_keys = {fld["key"] for fld in data.get("fields", [])}
    schema_keys = {f["key"] for f in QUESTIONNAIRE_SCHEMA}
    missing = schema_keys - tpl_keys
    assert not missing, "AI 模板缺少这些问卷字段：%s" % sorted(missing)


def test_ai_template_prompt_mentions_columns():
    """提示词本身要讲清栏数怎么填（否则 AI 不知道 1/2 的含义）。"""
    tpl = os.path.join(ROOT, "src", "data", "ai_questionnaire_template.yaml")
    with open(tpl, "r", encoding="utf-8") as f:
        text = f.read()
    low = text.lower()
    assert "columns" in low
    assert "single column" in low and "two-column" in low


def test_target_from_ai_reads_columns(tmp_path):
    """AI 导入（YAML，含 filled 包装）能读到 columns。"""
    p = str(tmp_path / "ai.yaml")
    with open(p, "w", encoding="utf-8") as f:
        f.write("filled:\n  columns: 2\n  reference_style: IEEE\n")
    t = target_from_ai(p)
    assert t.columns == 2
    assert t.resolution.get("columns") == "ai_import"


def test_target_from_dict_accepts_string_columns():
    """手工问卷的输入框给的是字符串 "2"，_coerce 必须转成 int 2。"""
    t = target_from_dict({"columns": "2", "reference_style": "IEEE"})
    assert t.columns == 2
