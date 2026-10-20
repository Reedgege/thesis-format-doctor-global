"""「修正必须按检查结果改」的端到端护栏（v2.3.70）。

老板原话（2026-10-19）：
    「修正要按照检查的结果修正，执行逻辑是先检查（按照规范和客户要求这些），
      然后修正，修正一定是根据检查结果进行修正，不是自己瞎修，
      检查结果要仔细，修正时能读取详细检查结果（真正详细的）」

本轮把两处「粗粒度」收紧成「按检查明细执行」，这里逐条钉死：

1. **标题维度细化到级**：检查只报「Level 3 heading format」，修正就只动三级，
   不许整五级重刷；未报错的层级一个字节都不许碰。
2. **预览与实际严格一致**：``_compute_changes`` 与 ``fix_docx`` 共用 ``_heading_scope``，
   预览里列出的层级 = 真会改的层级（绝不出现「预览说要改、修了却没改」或反之）。
3. **参考文献内容改写由检查结论门控**：检查没报 reference 问题（条目本身合规）时，
   哪怕 target 是 IEEE 且 reference_rewrite=True，也一个字都不许改。
"""
from __future__ import annotations

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from src.engine.checker import Finding, FixAction, run_check
from src.engine.docx_reader import read_docx
from src.engine.fixer import (_compute_changes, _heading_scope,
                              _reference_problems_reported, fix_docx)
from src.engine.questionnaire import TargetProfile, target_from_spec


# ---------------------------------------------------------------- 纯函数层
def _heading_finding(lvl=None, severity="warn"):
    f = Finding(dimension="heading", severity=severity,
                message=("Level %d heading format" % lvl) if lvl else "Heading levels",
                expected="x", actual="y", source="spec")
    f.fix = FixAction(dimension="heading", value=True,
                      levels=(lvl,) if lvl else ())
    return f


def test_heading_scope_reads_levels_from_findings():
    """检查结论带 levels=(3,) → 修正范围就是三级，不能是全篇。"""
    assert _heading_scope([_heading_finding(3)]) == frozenset({3})
    assert _heading_scope([_heading_finding(2), _heading_finding(4)]) == frozenset({2, 4})


def test_heading_scope_empty_when_only_dimension_level_conclusion():
    """「层级数不符 / 没有 Heading 样式」是**整维度**结论，不给任何一级授权改格式。

    v2.3.73 改口径（旧断言是 `is None` = 不限级）：这类结论说的是「层级识别/层级数」
    问题，不是「第 N 级的格式不对」。把它当「不限级」→ 全级重刷，实测把论文大标题
    编成 "I."、把章节标题改成 "A. Introduction"，而预览一行都没提。
    现在返回空集合 = 一级都不改；规范基线模式下真有问题时检查器会逐级报（含编号）。
    """
    assert _heading_scope([_heading_finding(None)]) == frozenset()
    # 有具体层级 + 整体结论并存 → 只认具体层级
    assert _heading_scope([_heading_finding(2), _heading_finding(None)]) == frozenset({2})


def test_heading_scope_none_only_for_legacy_no_findings():
    """不传 findings（旧调用方）= 没有检查结果可比 → None（保留旧的全维度 diff 语义）。"""
    assert _heading_scope(None) is None
    # 传了 findings 但里面没有标题结论 → 空集合：没有任何一级被授权修改
    assert _heading_scope([]) == frozenset()
    other = Finding(dimension="margins", severity="warn", message="Top margin")
    other.fix = FixAction(dimension="margins", value=True)
    assert _heading_scope([other]) == frozenset()


def test_reference_problems_helper():
    """只有真实报错（非 pass）的 reference 结论才算问题。"""
    assert _reference_problems_reported(None) is False
    assert _reference_problems_reported([]) is False
    assert _reference_problems_reported(
        [Finding(dimension="reference", severity="pass", message="Reference numbering")]
    ) is False
    assert _reference_problems_reported(
        [Finding(dimension="reference", severity="warn", message="Reference numbering")]
    ) is True
    # 其它维度的报错不得触发参考文献内容改写
    assert _reference_problems_reported(
        [Finding(dimension="reference_size", severity="fail", message="Reference entry font size")]
    ) is False


# ---------------------------------------------------------------- 端到端：标题只修报错层级
def _doc_three_headings(tmp_path, name="in.docx"):
    src = tmp_path / name
    doc = Document()
    for lvl, text in ((1, "Chapter One"), (2, "Section Two"), (3, "Detail Three")):
        p = doc.add_paragraph()
        p.style = doc.styles["Heading %d" % lvl]
        p.add_run(text)
    doc.add_paragraph("Body text one.")
    doc.save(src)
    return src


def _target_all_levels_centered():
    """三级标题都要求居中 —— 若修正不限级，三级会全被改成居中（便于识别越界）。"""
    return TargetProfile(
        heading_levels=3,
        heading_formats={1: {"alignment": "center"},
                         2: {"alignment": "center"},
                         3: {"alignment": "center"}},
        resolution={"heading_formats": "user"},
    )


def _heading_aligns(path):
    doc = Document(str(path))
    return [p.paragraph_format.alignment for p in doc.paragraphs[:3]]


def test_fix_touches_only_reported_heading_levels(tmp_path):
    """检查只报 Level 2 → 只有二级标题被改，一级/三级原样不动。"""
    src = _doc_three_headings(tmp_path)
    out = tmp_path / "out.docx"
    fix_docx(str(src), str(out), _target_all_levels_centered(),
             findings=[_heading_finding(2)])
    aligns = _heading_aligns(out)
    assert aligns[1] == WD_ALIGN_PARAGRAPH.CENTER, "报错的二级必须被修"
    assert aligns[0] is None, "未报错的一级不许动"
    assert aligns[2] is None, "未报错的三级不许动"


def test_fix_touches_nothing_when_check_is_only_a_dimension_conclusion(tmp_path):
    """检查只给了整维度结论（无层级）→ 三级标题一个都不许动（v2.3.73 收紧）。

    旧行为是「三级全部按目标修正」—— 那等于把「层级数不符」当成「每级格式都不对」，
    属于替检查结果加戏；实测后果是把客户手排标题重写成 "A. Introduction"。
    """
    src = _doc_three_headings(tmp_path)
    out = tmp_path / "out.docx"
    fix_docx(str(src), str(out), _target_all_levels_centered(),
             findings=[_heading_finding(None)])
    aligns = _heading_aligns(out)
    assert all(a is None for a in aligns), aligns


def test_preview_lists_exactly_the_levels_that_will_change(tmp_path):
    """预览与实际一致：预览只出现报错层级，且该层级确实被改。"""
    src = _doc_three_headings(tmp_path)
    out = tmp_path / "out.docx"
    target = _target_all_levels_centered()
    prof = read_docx(str(src))
    findings = [_heading_finding(2)]

    lines = _compute_changes(prof, target, "en", findings=findings)
    joined = " | ".join(lines)
    assert "Level 2 heading" in joined, joined
    assert "Level 1 heading" not in joined, joined
    assert "Level 3 heading" not in joined, joined

    fix_docx(str(src), str(out), target, findings=findings)
    aligns = _heading_aligns(out)
    # 预览说了二级会改 → 二级真的变了；预览没说一/三级 → 一/三级真的没变
    assert aligns[1] == WD_ALIGN_PARAGRAPH.CENTER
    assert aligns[0] is None and aligns[2] is None


def test_preview_empty_when_check_is_only_a_dimension_conclusion(tmp_path):
    """与 fix_docx 同口径：整维度结论 → 预览也一行标题改动都不许列。"""
    src = _doc_three_headings(tmp_path)
    prof = read_docx(str(src))
    lines = _compute_changes(prof, _target_all_levels_centered(), "en",
                             findings=[_heading_finding(None)])
    joined = " | ".join(lines)
    for lvl in (1, 2, 3):
        assert "Level %d heading" % lvl not in joined, joined


def test_heading_scope_ignores_passed_umbrella_finding():
    """真实检查结果里必然有一条 pass 的「Heading levels」总览（不带层级）。

    它一旦参与范围计算，就把范围撑成「不限级」→ 又变成全五级重刷。
    这里钉死：pass 的结论不算「要改的理由」，只认真正报错的那几级。
    """
    umbrella = Finding(dimension="heading", severity="pass",
                       message="Heading levels",
                       expected="", actual="Detected 3 levels", source="spec")
    umbrella.fix = FixAction(dimension="heading", value=3)      # levels=() 空
    assert _heading_scope([umbrella]) == frozenset(), "只有 pass 结论 → 没有要改的层级"
    assert _heading_scope([umbrella, _heading_finding(3)]) == frozenset({3})
    # 真正报错的整维度结论（非 pass）同样不给任何一级授权（v2.3.73）
    assert _heading_scope([umbrella, _heading_finding(None, severity="warn")]) == frozenset()


# ---------------------------------------------------------------- 端到端：真实检查结论驱动
def test_real_check_scopes_heading_fix_to_the_failing_level(tmp_path):
    """真实链路：一级/二级合规、三级不合规 → 检查只报三级，修正只动三级、预览只列三级。"""
    src = tmp_path / "mixed.docx"
    doc = Document()
    specs = [(1, WD_ALIGN_PARAGRAPH.CENTER, True),
             (2, WD_ALIGN_PARAGRAPH.LEFT, True),
             (3, WD_ALIGN_PARAGRAPH.LEFT, False)]     # 三级：对齐错 + 不加粗
    for lvl, align, bold in specs:
        p = doc.add_paragraph()
        p.style = doc.styles["Heading %d" % lvl]
        p.paragraph_format.alignment = align
        p.add_run("Heading %d text" % lvl).font.bold = bold
    doc.add_paragraph("Body text.")
    doc.save(src)

    target = TargetProfile(
        heading_levels=3,
        heading_formats={1: {"alignment": "center", "bold": True},
                         2: {"alignment": "left", "bold": True},
                         3: {"alignment": "center", "bold": True}},
        resolution={"heading_formats": "user"},
    )
    _, findings = run_check(str(src), target)
    assert _heading_scope(findings) == frozenset({3}), [
        (f.dimension, f.severity, f.message) for f in findings]

    prof = read_docx(str(src))
    joined = " | ".join(_compute_changes(prof, target, "en", findings=findings))
    assert "Level 3 heading" in joined, joined
    assert "Level 1 heading" not in joined and "Level 2 heading" not in joined, joined

    out = tmp_path / "mixed_out.docx"
    fix_docx(str(src), str(out), target, findings=findings)
    d = Document(str(out))
    assert d.paragraphs[2].paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert d.paragraphs[2].runs[0].font.bold is True


# ---------------------------------------------------------------- 检查器必须「看得见」才能被修正
def test_checker_reports_missing_heading_numbering(tmp_path):
    """IEEE 要求的章节编号（I./A./1)）以前**从不检查**，修正器却在补编号。

    报告说「标题格式合规」、标题却多出 "A." —— 检查看不见的项被改了。
    现在缺编号必须作为结论出现，修正才有依据。
    """
    src = tmp_path / "num.docx"
    doc = Document()
    doc.add_paragraph("Introduction", style="Heading 1")
    doc.add_paragraph("Background", style="Heading 2")
    doc.save(src)
    _, findings = run_check(str(src), target_from_spec("IEEE"))
    lvl2 = [f for f in findings
            if f.dimension == "heading" and f.message == "Level 2 heading format"]
    assert lvl2, [(f.dimension, f.message) for f in findings]
    joined = " ".join(f.expected + " " + f.actual for f in lvl2)
    assert "numbering" in joined.lower(), joined
    # 报出来了 → 才会进修正范围
    assert _heading_scope(findings) is not None
    assert 2 in _heading_scope(findings)


def test_checker_reports_missing_uppercase_once_fixed(tmp_path):
    """IEEE 一级标题要求全大写：没大写要报，修完复检不再报（幂等）。"""
    src = tmp_path / "caps.docx"
    doc = Document()
    doc.add_paragraph("I. Introduction", style="Heading 1")
    doc.save(src)
    target = target_from_spec("IEEE")
    _, before = run_check(str(src), target)
    caps = [f for f in before if f.dimension == "heading"
            and f.message == "Level 1 heading format"]
    assert caps and any("uppercase" in (f.expected + f.actual).lower() for f in caps), \
        [(f.expected, f.actual) for f in caps]

    out = tmp_path / "caps_out.docx"
    fix_docx(str(src), str(out), target, findings=before)
    assert Document(str(out)).paragraphs[0].text.startswith("I. INTRODUCTION")
    _, after = run_check(str(out), target)
    caps2 = [f for f in after if f.dimension == "heading"
             and f.message == "Level 1 heading format" and f.severity != "pass"]
    assert not caps2, [(f.severity, f.expected, f.actual) for f in caps2]


def test_checker_marks_heading_italic_unreadable(tmp_path):
    """APA 三级要斜体，文档读不到斜体 → 必须如实说「读不到」，不许静默跳过。

    旧写法 `if got_b is None: continue` 让「三级该斜体」这条要求在报告里消失，
    修正器却照加斜体 —— 又是「检查看不见、却被改了」。
    """
    src = tmp_path / "ital.docx"
    doc = Document()
    p = doc.add_paragraph(style="Heading 3")
    p.add_run("Detail")
    doc.save(src)
    target = TargetProfile(
        heading_levels=3,
        heading_formats={3: {"alignment": "left", "bold": True, "italic": True}},
        resolution={"heading_formats": "user"},
    )
    _, findings = run_check(str(src), target)
    lvl3 = [f for f in findings
            if f.dimension == "heading" and f.message == "Level 3 heading format"]
    assert lvl3, [(f.dimension, f.message) for f in findings]
    assert any("italic" in f.actual.lower() for f in lvl3), [f.actual for f in lvl3]


# ---------------------------------------------------------------- 端到端：参考文献内容改写门控
def _doc_numbered_apa_refs(tmp_path, name="refs.docx"):
    """参考文献已按 [n] 编号、但条目仍是 APA 作者-年形态。

    这种文档在检查器眼里「编号没问题」，reference 维度不会报错 —— 正是用来验证
    「检查没报错就不许改内容」的场景（旧版只看 reference_style == IEEE 就改写）。
    """
    src = tmp_path / name
    doc = Document()
    doc.add_paragraph("Body paragraph for the thesis.")
    doc.add_paragraph("References")
    doc.add_paragraph("[1] Anderson, K., & Lee, S. (2019). Sleep quality and memory "
                      "consolidation. Journal of Sleep Research, 12(3), 45-67.")
    doc.add_paragraph("[2] Smith, J. A. (2020). Title of Work. Publisher.")
    doc.save(src)
    return src


def test_no_reference_problem_means_no_content_rewrite(tmp_path):
    """检查没报 reference 问题 → IEEE 目标 + reference_rewrite 也不许改条目文字。"""
    src = _doc_numbered_apa_refs(tmp_path)
    out = tmp_path / "out.docx"
    target = target_from_spec("IEEE")
    assert target.reference_rewrite is True, "产品路径默认开启（但必须被检查结论门控）"
    before = [p.text for p in Document(str(src)).paragraphs[2:]]

    # 只给一个「参考文献悬挂缩进」类修正（真实场景：条目文字合规、只是缩进不对）
    hang = Finding(dimension="reference_hanging_indent", severity="warn",
                   message="Reference hanging indent", expected='0.5"', actual='0"')
    hang.fix = FixAction(dimension="reference_hanging_indent", value=True)
    fix_docx(str(src), str(out), target, findings=[hang])

    after = [p.text for p in Document(str(out)).paragraphs[2:]]
    assert after == before, "检查没报错时，条目文字必须逐字不动：%r" % (after,)


def test_reference_problem_reported_opens_the_rewrite(tmp_path):
    """对照：检查确实报了 reference 问题 → 允许 best-effort 改写为 IEEE 形态。"""
    src = _doc_numbered_apa_refs(tmp_path)
    out = tmp_path / "out.docx"
    reported = Finding(dimension="reference", severity="warn",
                       message="Reference sequential numbering")
    reported.fix = FixAction(dimension="reference_numbered", value=True)
    fix_docx(str(src), str(out), target_from_spec("IEEE"), findings=[reported])
    after = [p.text for p in Document(str(out)).paragraphs[2:]]
    assert 'vol. 12, no. 3, pp. 45-67, 2019.' in after[0], after[0]
    assert after[1].startswith("[2] J. A. Smith,"), after[1]


def test_preview_shows_reference_rewrite_when_it_will_happen(tmp_path):
    """改写条目是**改内容**，必须在预览里写出来（客户有权改前知道哪几条被重写）。"""
    src = _doc_numbered_apa_refs(tmp_path)
    prof = read_docx(str(src))
    target = target_from_spec("IEEE")
    reported = Finding(dimension="reference", severity="warn",
                       message="Reference sequential numbering")
    reported.fix = FixAction(dimension="reference_numbered", value=True)

    zh = " | ".join(_compute_changes(prof, target, "zh", findings=[reported]))
    assert "参考文献条目格式" in zh and "改写 2 条" in zh, zh
    en = " | ".join(_compute_changes(prof, target, "en", findings=[reported]))
    assert "Reference entry format" in en and "rewrite 2 entries" in en, en
    # 英文界面不得泄漏中文
    assert all(ord(ch) < 128 for ch in en), en

    # 没有 reference 报错 → 预览里也不许出现这行（否则等于骗客户会改）
    quiet = " | ".join(_compute_changes(prof, target, "en", findings=[]))
    assert "Reference entry format" not in quiet, quiet
