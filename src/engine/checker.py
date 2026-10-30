"""Rule Engine（确定性规则引擎）：把实测文档画像与 TargetProfile 比对。

原则：
- 全部为确定性判断，不调用任何大模型 / 网络。
- 每条 Finding 标注来源（spec / school / user / ai），供对比报告裁决。
- 页面排版来自「规范」时仅作参考（info），不强制 fail；
  页面排版来自学校/问卷/AI 时硬判。参考文献始终以规范硬判。
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Any, Optional

from .docx_reader import (DocumentProfile, read_docx, DocxReadError,
                          describe_paper_size, identify_paper_size,
                          find_abstract_section, describe_table_lines)
from .questionnaire import TargetProfile, PAGE_NUMBER_LAYOUT, origin_of_source
from .policy import TEXT_IS_LOCKED
from .specs import (MARGIN_DIMS, get_spec, resolve_spec,
                    resolve_heading_formats, describe_heading, size_for_font,
                    spec_is_baseline)


@dataclass
class FixAction:
    """checker 产出、fixer 消费的「修正动作」。

    这是 **check→fix 单一真相源** 的载体：checker 在发现某维度需要修正时，
    把「改哪个维度、改成什么值」封进 FixAction 挂到对应 Finding 上；fixer 只消费
    带 fix 的 Finding，绝不再独立重算「要改什么」。这样「检查出来的」与「实际改的」
    必然一致，也杜绝了 fixer 用另一套规范/逻辑重算导致的漂移。

    - dimension: 与 fixer 内部动作对应的维度名（margins/font/.../reference_numbered/heading）。
    - value: 该维度的目标值（供 preview/报告展示与断言；fixer 施加时仍读 target 的同名值，
      两者恒等，故 value 仅作展示与一致性核对）。
    - note: 可选的人类可读说明（编号制条目数等）。
    """
    dimension: str
    value: Any = None
    note: str = ""
    # v2.3.70：**结构化明细**——把「检查具体报了哪一级」带给修正器，
    # 让它只改被报错的对象，而不是整个维度重刷一遍（老板：修正一定按检查结果来，
    # 检查结果要详细、修正要能读到）。标题维度按「级」细化；其它维度留空=整维度。
    levels: tuple = ()


def _finding_heading_level(msg: str) -> Optional[int]:
    """从「Level N heading format」这类 finding 文案里取出**具体层级**。

    只认固定前缀（我们自己产出），取不到就返回 None —— 调用方据此判定
    「这条结论没有细到级」，退回整维度处理，绝不猜。
    """
    m = re.match(r"^Level\s+(\d+)\s", msg or "")
    if not m:
        return None
    try:
        lv = int(m.group(1))
    except (TypeError, ValueError):
        return None
    return lv if lv > 0 else None


# 容差（依据：1/16 英寸内视为排版误差；字号/行距给更宽 warn 带）
# 约定：所有维度的 warn 带必须**宽于**其 fail 带（WARN < FAIL），
#       否则 warn 段几乎不可达（曾出现 MARGIN_TOL_WARN=0.05 < FAIL=0.06 的笔误）。
MARGIN_TOL_FAIL = 0.10      # 英寸
MARGIN_TOL_WARN = 0.05      # 英寸（≈1.27mm）
SIZE_TOL_FAIL = 0.5         # 磅
SIZE_TOL_WARN = 0.3
SPACING_TOL_FAIL = 0.15     # 倍
SPACING_TOL_WARN = 0.1
INDENT_TOL_FAIL = 0.12      # 英寸（正文首行缩进）
INDENT_TOL_WARN = 0.06
# v2.3.100（F 批 F-1 / 裁定 C-1）：Keywords 行缩进的**合格下限**。
# APA §2.10 手册写 0.5″，但官方专业版样例实测 0.0″、design 模板实测 0.431″ ——
# 锁 0.5″ 会把官方合规样例判红。故只要求"有缩进且 ≥0.4″"，不锁具体值。
KEYWORDS_INDENT_MIN_IN = 0.4
HEADER_TOL_FAIL = 0.12      # 英寸（页眉距页顶）
HEADER_TOL_WARN = 0.06
SPACE_TOL_FAIL = 1.5        # 磅（段后间距）
SPACE_TOL_WARN = 0.8
REF_HANG_TOL_FAIL = 0.15    # 英寸（参考文献悬挂缩进；引用维度，容差略宽）
REF_HANG_TOL_WARN = 0.03

SEVERITY_ORDER = {"pass": 0, "info": 1, "warn": 2, "fail": 3}

# Alignment labels (also reused by report.py to avoid duplicate hard-coding)
ALIGN_LABEL = {"left": "Left", "justify": "Justified",
               "center": "Center", "right": "Right"}

# 章节编号形态的人类可读提示（进报告 expected 列；面向客户，一律英文）。
_NUMBER_PREFIX_HINT = {
    "roman": "I., II., III.",
    "upper_letter": "A., B., C.",
    "number": "1., 2., 3.",
    "number_paren": "1), 2), 3)",
    "lower_letter": "a), b), c)",
    "lower_roman": "i), ii), iii)",
}


def _append_text_lock_note(findings: list, lvl: int, text_bits: list,
                           source) -> None:
    """v2.3.84：把「改文字类要求」（章节编号 / 全大写）记成 info 提示，不绑修正。

    老板 2026-10-21 更新清单第 5 条要求「文字内容 100% 锁死，只改 Style 与 Layout」，
    而 IEEE 规范恰好要求标题带 ``I. / A. / 1.`` 编号并全大写 —— 这两件事都是改字符。
    产品选择：**不改**（承诺优先），但必须在报告里如实告诉客户"规范要求这样，请自行处理"。
    与脚注/尾注、目录条目「不修只提示」同一套做法。
    """
    if not text_bits:
        return
    findings.append(Finding(
        dimension="heading", severity="info",
        message="Level %d heading numbering/case (left as-is on purpose)" % lvl,
        expected=", ".join(text_bits),
        actual="Text is never rewritten (format-only tool). Add the numbering or "
               "uppercasing yourself, or use Word multi-level list numbering.",
        source=source,
    ))


def _tp_expected_text(vals: dict) -> str:
    """把标题页三态期望拼成一句英文说明（进报告 expected 列）。"""
    bits = []
    if vals.get("title_page_title_centered") is not None:
        bits.append("title centered" if vals["title_page_title_centered"]
                    else "title left-aligned")
    if vals.get("title_page_title_bold") is not None:
        bits.append("title bold" if vals["title_page_title_bold"] else "title not bold")
    if vals.get("title_page_author_centered") is not None:
        bits.append("author/affiliation centered"
                    if vals["title_page_author_centered"]
                    else "author/affiliation not centered")
    # v2.3.76：标题段的段前 / 段后（磅）
    if vals.get("title_page_space_before_pt") is not None:
        bits.append("title space before %g pt" % float(vals["title_page_space_before_pt"]))
    if vals.get("title_page_space_after_pt") is not None:
        bits.append("title space after %g pt" % float(vals["title_page_space_after_pt"]))
    # v2.3.114：题目字号 / 字体
    if vals.get("title_page_title_size_pt") is not None:
        bits.append("title %g pt" % float(vals["title_page_title_size_pt"]))
    if vals.get("title_page_title_font"):
        bits.append("title font %s" % vals["title_page_title_font"])
    return ", ".join(bits)

# 已带顺序编号的形态：[1] / (1) / 1. / 1)  —— 与 fixer 的 _LEADING_NUMBER_RE 保持一致
_HAS_LEADING_NUMBER_RE = re.compile(r"^\s*(?:\[\s*\d+\s*\]|\(\s*\d+\s*\)|\d+\s*[.)])\s")


def _has_leading_number(text: str) -> bool:
    """条目文本是否已带顺序编号前缀（[n] / (n) / n. / n)）。"""
    return bool(_HAS_LEADING_NUMBER_RE.match(text or ""))


@dataclass
class Finding:
    dimension: str
    severity: str            # pass / warn / fail / info
    message: str
    expected: str = ""
    actual: str = ""
    source: str = ""         # spec / school / user / ai
    fix: Optional[FixAction] = None   # 非 None 表示 check 认为该项需要（且可被）自动修正


# ---------------------------------------------------------------------------
# 确定性比对基元
# ---------------------------------------------------------------------------

def _cmp_margin(measured: Optional[float], expected: float) -> tuple[str, str]:
    if measured is None:
        return "info", "Could not read (document may have no section or be protected)"
    diff = abs(measured - expected)
    # Note: only the value itself is returned here; the report layer appends
    # "Expected: ... | Actual: ..." — so do NOT prepend "Actual" here
    # (otherwise it would read "Actual: Actual 2.0\"").
    if diff <= MARGIN_TOL_WARN:
        return "pass", f'{measured}"'
    if diff <= MARGIN_TOL_FAIL:
        return "warn", f'{measured}" (deviation {diff:.2f}")'
    return "fail", f'{measured}" (deviation {diff:.2f}")'


def _cmp_num(measured: Optional[float], expected: float,
             tol_fail: float, tol_warn: float, unit: str) -> tuple[str, str]:
    if measured is None:
        return "info", "Could not read (unable to determine this dimension from the document)"
    diff = abs(measured - expected)
    if diff <= tol_warn:
        return "pass", f"{measured}{unit}"
    if diff <= tol_fail:
        return "warn", f"{measured}{unit} (deviation {diff:.2f})"
    return "fail", f"{measured}{unit} (deviation {diff:.2f})"


def _downgrade(sev: str, origin: str, is_page_or_heading: bool,
               dim: Optional[str] = None, spec_key: Optional[str] = None) -> str:
    """规范来源的页面/标题维度仅作参考（info），不强制 fail/warn。

    例外（Task B）：当规范为 APA 且其维度属于「不可被模板推翻的硬性正文规则」时，
    即便来自 spec，也至少保留 warn 下限。这些规则在 APA 7th 中明确且不可能由学校模板
    覆盖：
      - alignment（§2.22：不得两端对齐，须左对齐）
      - first_line_indent（§2.24：每段首行缩进 0.5″）
      - heading（§2.27：标题层级≤5，且应使用标题样式）
    其余页面维度（页边距/字体/字号/行距）仍按模板驱动原则降级为 info。
    """
    if is_page_or_heading and origin == "spec" and sev in ("warn", "fail"):
        if spec_key and spec_key.startswith("APA") and dim in (
            "alignment", "first_line_indent", "heading"
        ):
            return "warn"
        return "info"
    return sev


def _page_number_severity(origin: str, target) -> str:
    """页码的严重度（v2.3.90，老板 2026-10-22 拍板「升为可修」）。

    页码**不是版面偏好**，是三家规范的硬要求：APA 7th §2.18 每页右上角；
    MLA 右上角「姓氏 + 页码」；IEEE 页脚居中。此前它和页边距/字体一起被
    「规范来源的页面维度一律降级为 info」吞掉 —— 报告写 info、修正器不动手，
    客户拿到「已按 APA 改好」的论文却一页页码都没有（2026-10-22 实测：四份产物
    zip 里连页眉部件都不存在，headerReference 出现 0 次）。

    新口径：客户**真的在规范下拉里选了内置规范**（``spec_is_baseline``）时，
    页码缺失/错位至少 warn（从而绑上修正动作）；走学校模板 / 手填 / AI 导入时，
    仍按模板说话 —— 显式来源本来就不会被降级，这条分支不受影响。
    """
    sev = _downgrade("warn", origin, True, "page_number",
                     getattr(target, "spec_key", None))
    if sev == "info" and spec_is_baseline(target):
        return "warn"
    return sev


def title_page_spacing_plan(target, spec, prof) -> dict:
    """题名页段距的「期望值 + 允许区间」（磅）。**checker 与 fixer 共用的唯一计算**。

    老板 2026-10-23「规范里面没有要求的我们不要加，仔细看规范」——
    于是把 APA Style 官方 Title Page Setup 的三句原话逐条落成数字：

      ① "Place the title three to four lines down from the top of the title page."
         → **区间** [3, 4] 行。v2.3.91 按精确 3 行判，把合规的 4 行误报成错。
      ② "Place one double-spaced blank line between the paper title and the author names."
         → 标题与作者之间是「一个空行」，**不是 0**。v2.3.91 把它清成 0，反向违规。
      ③ "no extra lines added between elements, except after the paper title"
         → 其余元素之间 0 额外段距。

    一行 = ``字号 × 行距倍数`` 磅（APA 12pt × 2.0 = 24pt）。容差按「行」给
    （``title_page_lines_tolerance``），因为行高换算本来就受字体度量影响。

    返回::

        {"unit_pt": 24.0, "tolerance_lines": 0.25, "baseline": True,
         "before": (66.0, 102.0) | None, "before_lines": (3.0, 4.0) | None,
         "after": (18.0, 30.0) | None,   "after_lines": 1.0 | None,
         "author_zero": True}

    两个键为 None 表示该规范不规定此项 —— 调用方一个字节都不要写/不要判。

    🔴 v2.3.93（老板 2026-10-23「学校模板、手动输入、AI 导入是不是也有类似问题」）：
    ①② 是 **APA 原文的措辞**（"three to four lines" / "one double-spaced blank line"），
    不是普世规则，所以**只在规范基线**下成立。走学校模板 / 手填 / AI 时，
    「标题距顶多少」由客户自己的 ``title_page_space_before_pt`` 表达、
    「标题与作者之间多少」由 ``title_page_space_after_pt`` 表达，
    这两项的 owner 是 ``title_page`` 维度（精确比对自己填的值），
    本函数绝不再拿 APA 的 66–102 磅去判客户的 60 磅 —— 那等于用内置规范推翻客户模板。
    ③（其余元素零额外段距）在非基线下改听 ``title_page_zero_extra_spacing``。
    """
    plan = {"unit_pt": None, "tolerance_lines": 0.0, "baseline": False,
            "before": None, "before_lines": None,
            "after": None, "after_lines": None,
            "author_zero": False}
    baseline = bool(spec_is_baseline(target))
    plan["baseline"] = baseline
    tpf = getattr(spec, "title_page", None)
    # ③ 其余元素零额外段距：基线听规范原话，其他通道听客户显式勾选。
    # 🔴 必须在 `tpf is None` 早退**之前**取值 —— 学校模板 / 手填 / AI 通道下
    # `resolve_spec()` 可能落到压根没有 title_page 的兜底规范，早退会把客户
    # 明确勾了的「元素间不留额外段距」整项吞掉（v2.3.93 实测踩到）。
    if baseline:
        plan["author_zero"] = bool(getattr(tpf, "zero_extra_spacing", False)) \
            if tpf is not None else False
    else:
        plan["author_zero"] = bool(
            getattr(target, "title_page_zero_extra_spacing", False))
    if tpf is None:
        return plan
    page = getattr(target, "page", None) or {}
    try:
        fs = float(page.get("font_size_pt") or getattr(prof, "font_size_pt", None) or 12.0)
    except (TypeError, ValueError):
        fs = 12.0
    try:
        ls = float(page.get("line_spacing") or getattr(prof, "line_spacing", None) or 1.0)
    except (TypeError, ValueError):
        ls = 1.0
    unit = round(fs * ls, 2)
    try:
        tol = float(getattr(tpf, "title_page_lines_tolerance", 0.25) or 0.0)
    except (TypeError, ValueError):
        tol = 0.0
    plan["unit_pt"] = unit
    plan["tolerance_lines"] = tol

    # ①② 只属于规范原文 → 非基线一律不判（客户自己的磅值由 title_page 维度负责）。
    if not baseline:
        return plan

    lo_lines = getattr(tpf, "title_space_before_lines", None)
    if lo_lines is not None:
        hi_lines = getattr(tpf, "title_space_before_lines_max", None)
        hi_lines = lo_lines if hi_lines is None else max(lo_lines, hi_lines)
        plan["before_lines"] = (float(lo_lines), float(hi_lines))
        plan["before"] = (round((float(lo_lines) - tol) * unit, 2),
                          round((float(hi_lines) + tol) * unit, 2))

    sa_lines = getattr(tpf, "title_space_after_lines", None)
    if sa_lines is not None:
        plan["after_lines"] = float(sa_lines)
        plan["after"] = (round((float(sa_lines) - tol) * unit, 2),
                         round((float(sa_lines) + tol) * unit, 2))
    return plan


# —— 题名页「内容完整性」判定的线索（v2.3.92，仅用于提示，绝不代客户写内容）——
# 认定刻意宽松：认出任何一条线索就算「有」，宁可漏提示，也不要把客户已经写好的
# 东西说成缺。论文大标题不在检索范围内（见 docx_reader.tp_author_block_text），
# 否则 "COVID-19" 这类「字母+数字」会被当成课程号。
_TP_COURSE_RE = re.compile(r"(?i)\b[a-z]{2,6}[\s\-]?\d{2,4}[a-z]?\b")
_TP_INSTRUCTOR_RE = re.compile(
    r"(?i)\b(?:dr|prof|professor|instructor|lecturer|teacher)\b\.?")
_TP_DATE_RE = re.compile(
    r"(?i)(?:\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s+\d{1,2}\b"
    r"|\b\d{1,2}\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\b"
    r"|\b\d{1,2}/\d{1,2}/\d{2,4}\b"
    r"|\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b)")


def title_page_missing_elements(prof, spec) -> list:
    """APA **学生版**题名页要求、但文档里找不到的补充元素（v2.3.92，只用于提示）。

    APA Style 学生版题名页共 7 项：标题 / 作者 / 单位 / **课程号与课名** / **教师名** /
    **提交日期** / 页码。前三项由版式检查覆盖，后三项是**内容**，本工具绝不代写
    （文字 100% 锁死），只能如实告诉客户「还差什么，请自己补」。

    只对 ``title_page_kind == "student"`` 生效：APA 专业版题名页要求的是作者注与
    running head，**没有**课程/教师/日期这三项 —— 规范没有要求的绝不多加。
    """
    if getattr(spec, "title_page_kind", "") != "student":
        return []
    text = getattr(prof, "tp_author_block_text", None)
    if not text:
        return []
    missing = []
    if not _TP_COURSE_RE.search(text):
        missing.append("course number and name")
    if not _TP_INSTRUCTOR_RE.search(text):
        missing.append("instructor name")
    if not _TP_DATE_RE.search(text):
        missing.append("assignment due date")
    return missing


_ALIGN_WORD = {"left": "left-aligned", "center": "centered",
               "right": "right-aligned", "justify": "justified"}


def _pt_txt(val) -> str:
    """把磅值安全转成文本（D-06 修复）。

    字号可能是 ``None``（全文字号靠样式继承、无显式 run 字号），此时旧代码
    ``"%g" % None`` 会抛 ``TypeError``，**整份检查直接崩掉**——官方 APA 样例里
    有 3 份就是这样。这里统一兜底为 ``unknown``，保证报告一定能生成。
    """
    try:
        return "%g" % float(val)
    except (TypeError, ValueError):
        return "unknown"


def _describe_abstract_heading(align, bold, italic) -> str:
    """摘要标题期望值 → 客户可读描述（"centered, bold"）。只写规定了的那几项。"""
    bits = []
    if align:
        bits.append(_ALIGN_WORD.get(align, align))
    if bold is not None:
        bits.append("bold" if bold else "not bold")
    if italic is not None:
        bits.append("italic" if italic else "not italic")
    return ", ".join(bits) or "—"


def _describe_abstract_heading_actual(prof, want: dict):
    """摘要标题实测值 → (客户可读描述, 是否有不符)。

    「没显式设置」不算读不到，更不等于没责任：Word 里 run 上没写 w:b / w:i 就是
    不生效（除样式层恰好设了）。故按方向判 —— 规范要求「是」时算不符（该加没加），
    要求「否」时算通过（默认就是否）。这样不会平白多报一条红杠，也放过了真问题。
    """
    bits, bad = [], False
    got_al = getattr(prof, "abstract_heading_alignment", None)
    w_al = want.get("abstract_heading_alignment")
    if w_al:
        if got_al is None:
            bits.append("alignment not set")
            if w_al != "left":          # Word 缺省就是左对齐
                bad = True
        else:
            bits.append(_ALIGN_WORD.get(got_al, got_al))
            if got_al != w_al:
                bad = True
    for key, word in (("abstract_heading_bold", "bold"),
                      ("abstract_heading_italic", "italic")):
        w = want.get(key)
        if w is None:
            continue
        g = getattr(prof, key, None)
        if g is None:
            bits.append("%s not set" % word)
            if bool(w):
                bad = True
            continue
        bits.append(word if g else "not " + word)
        if bool(g) != bool(w):
            bad = True
    return (", ".join(bits) or "—"), bad


def _describe_keywords(italic, indent) -> str:
    """Keywords 行期望值 → 客户可读描述。只写规定了的那几项（没规定的不写）。"""
    bits = []
    if italic is not None:
        bits.append("label italic" if italic else "label not italic")
    if indent is not None:
        try:
            _n = float(indent)
        except (TypeError, ValueError):
            _n = None
        if _n is not None:
            bits.append("flush left" if _n <= 0.01 else "indented %g in" % _n)
    return ", ".join(bits) or "—"


def _describe_keywords_actual(prof, want: dict):
    """Keywords 行实测值 → (客户可读描述, 是否有不符)。

    读不到不算不符（没法判就不冤枉客户），但要在描述里说出来。
    """
    bits, bad = [], False
    w_i = want.get("keywords_label_italic")
    if w_i is not None:
        g = getattr(prof, "keywords_label_italic", None)
        if g is None:
            # 与摘要标题同一口径：要求「是」才算不符（默认就是不斜体）
            bits.append("label italic not set")
            if bool(w_i):
                bad = True
        else:
            bits.append("label italic" if g else "label not italic")
            if bool(g) != bool(w_i):
                bad = True
    w_n = want.get("keywords_indent_in")
    if w_n is not None:
        g = getattr(prof, "keywords_indent_in", None)
        try:
            gn = None if g is None else float(g)
            wn = float(w_n)
        except (TypeError, ValueError):
            gn = wn = None
        if gn is None:
            bits.append("indent unreadable")
        else:
            bits.append("flush left" if gn <= 0.01 else "indented %g in" % gn)
            # 🔴 v2.3.100（F 批 F-1 / 落地裁定 C-1）：只判「**有缩进 ≥ 0.4″**」，
            # **不锁死 0.5″**。依据：APA §2.10 手册写 0.5″，但**官方专业版样例实测
            # 0.0″、design 模板实测 0.431″** —— 按 0.5″ 硬判等于把官方合规样例判红。
            # 手册条文与官方样例冲突时，按"宁缺勿错"取下限（0.4″）作合格线，
            # 写回值仍用规范值（fixer 侧不变），修完复检必然过。
            if gn < KEYWORDS_INDENT_MIN_IN - INDENT_TOL_WARN:
                bad = True
    return (", ".join(bits) or "—"), bad


def _abstract_not_located(prof, suffix: str = "") -> str:
    """「摘要段定位不到」的客户可读解释（v2.3.91 起；v2.3.96 改因）。

    旧文案一律写 "No Abstract section detected"，但真实学生论文里最常见的情形是
    **摘要标题没套标题样式**（实测 student_paper_unformatted_20pages.docx：全文 0 个
    样式标题，连 "Abstract" 都是 Normal）——文档里明明有摘要，报告却说「没检测到」，
    客户只会以为工具坏了。故按「有没有摘要字样」分两种说法。

    🔴 v2.3.96：没套标题样式**已经不是定位不到的原因**了 —— `find_abstract_section`
    补了严格限定的文字回退（整行匹配 Abstract + 必须有 Keywords 行作右边界）。
    所以现在的真因是「找不到 Keywords 行、摘要块的右边界不可知」，
    文案必须照着这个说：否则客户会跑去给 Abstract 套标题样式，套了也不管用。
    """
    if getattr(prof, "has_abstract", False):
        return ("Abstract text is there, but no 'Keywords:' line marks where the "
                "abstract ends, so the section cannot be located" + suffix)
    return "No Abstract section detected" + suffix


# ---------------------------------------------------------------------------
# 参考文献检查：数据驱动（依据 ReferenceRule.checks 分发）
# 每个子检查返回 (severity, message)。
# ---------------------------------------------------------------------------

def _chk_has_year_paren(entries, style) -> tuple:
    text = "\n".join(entries)
    if re.search(r"\(\s*\d{4}\s*\)", text):
        return "pass", f"[{style}] Detected author-year form with '(year)' ({len(entries)} entries)"
    return "warn", f"[{style}] No author-year '(year)' pattern detected; please verify reference format"


def _chk_author_year_order(entries, style) -> tuple:
    text = "\n".join(entries)
    if re.search(r"^[A-Z][a-z]+,\s+[A-Z]\.?", text.strip(), re.M):
        return "pass", f"[{style}] Entry start matches author-year form"
    return "info", f"[{style}] Entry does not start with typical 'Author, Initials'; please verify manually"


def _chk_author_title_order(entries, style) -> tuple:
    text = "\n".join(entries)
    if re.search(r"^[A-Z]", text.strip(), re.M):
        return "pass", f"[{style}] Works Cited format roughly matches ({len(entries)} entries)"
    return "warn", f"[{style}] Entries do not start with a capital letter; please verify Works Cited format"


def _chk_author_title_place_order(entries, style) -> tuple:
    return "info", f"[{style}] {len(entries)} bibliographic entries detected; please verify 'Author. Title. Place: Publisher, Year.'"


def _chk_numbered_ref(entries, style) -> tuple:
    numbered = sum(1 for e in entries if re.match(r"^\s*\[\s*\d+\s*\]", e))
    if numbered == 0:
        return "warn", f"[{style}] No numbered [n] form detected; please verify reference format"
    return "pass", f"[{style}] Numbered form matches ({numbered}/{len(entries)} entries start with [n])"


_REF_CHECKS = {
    "has_year_paren": _chk_has_year_paren,
    "author_year_order": _chk_author_year_order,
    "author_title_order": _chk_author_title_order,
    "author_title_place_order": _chk_author_title_place_order,
    "numbered_ref": _chk_numbered_ref,
}


def _check_reference(style: str, entries: list, spec, source: str = "spec") -> Finding:
    """参考文献格式启发式检查（著录形态，非逐条语义）。数据驱动。

    source: 该结论的来源标注（FR4 要求「并标注该结论来源」）。
    """
    # 客户读的报告里不许出现内部键名：spec_key 缺失时 style 会是 "Other"，
    # 原样写进 expected 会让客户以为「应该写成 Other」（v2.3.71 客户视角实测）。
    _disp = "" if str(style or "").strip().lower() in ("", "other", "none") else style
    style = _disp
    if not entries:
        return Finding("reference", "warn",
                       "No reference entries detected (missing References / Bibliography / Works Cited section)",
                       expected=style, actual="0 entries", source=source)
    tokens = spec.reference.checks if spec and spec.reference.checks else ()
    if not tokens:
        # v2.3.71（客户视角）：这里原来写的是
        #   "Custom/unregistered check items (N entries); please verify against your school template"
        # 客户读不懂（"Custom/unregistered" 是我们的内部概念），而且 expected 直接打出
        # "Other" 这种内部键名，客户会误以为「应该写成 Other」。真正该说清的是两件事：
        #   · 没选引用风格 → 参考文献根本没被检查；选了 APA/MLA/IEEE 才会查；
        #   · 该风格没有内置条目检查 → 需要人工核对。
        if str(style or "").strip().lower() in ("", "other", "none"):
            return Finding("reference", "info",
                           "Reference style not specified — reference entries were NOT checked. "
                           "Choose APA, MLA or IEEE to have them checked",
                           expected="", actual=f"{len(entries)} entries found", source=source)
        return Finding("reference", "info",
                       f"Reference style {style} has no built-in entry checks; "
                       f"{len(entries)} entries were found — please verify them manually "
                       f"against your school template",
                       expected="", actual=f"{len(entries)} entries", source=source)
    results = []
    for tok in tokens:
        fn = _REF_CHECKS.get(tok)
        if fn:
            results.append(fn(entries, style))
    if not results:
        return Finding("reference", "info",
                       f"[{style}] {len(entries)} entries detected; please verify manually.",
                       expected=style, actual=f"{len(entries)} entries", source=source)
    # 取最严重（severity 数值最大）的一项
    worst = max(results, key=lambda r: SEVERITY_ORDER[r[0]])
    return Finding("reference", worst[0],
                   worst[1], expected=style, actual=f"{len(entries)} entries", source=source)


def _apa_reference_content_hints(entries: list, italics: list) -> list:
    """APA 参考文献「内容级」稳健提示（仅 info，绝不 fail/warn，且不误伤官方样例）。

    只针对整份文献的宏观可疑信号给提示，避免对逐条语义做武断判定：
    1) 至少 2 条、且全部无斜体 → 期刊名 / 书名可能未斜体（APA 7th §6.22 / 10.1）。
       官方标注样例含大量斜体期刊名，不会触发；纯网页型条目（罕见）最多给一条 info。
    2) 任一作者段出现 ", and "（逗号后接 and）→ 应用 "&" 连接最后两位作者（APA 7th §9.7）。
       只检查作者段（首个 "(" 或首个 4 位年份之前），避免误伤书名/标题中的 and
       （如 "War and Peace" 出现在年份之后，不计入作者段）。

    说明：官方样例本身存在「无 DOI」「缺期号」的条目（其教学批注亦如此标注），
    故此处刻意不检查 DOI / 期号，否则会误报金标准论文。
    """
    out: list = []
    # 1) 斜体提示：要求至少 2 条且确实存在「像文献」的条目（含 4 位年份），降低噪声
    if len(entries) >= 2 and any(re.search(r"\b(19|20)\d{2}\b", e) for e in entries):
        if not any(italics):
            out.append(Finding(
                dimension="reference",
                severity="info",
                message="Reference italics",
                expected="Journal names and book/report titles should be italicized (APA 7th §6.22 / 10.1)",
                actual="No italicized text found in any reference entry; verify journal names and book titles are italicized",
                source="spec",
            ))
    # 2) & 提示
    and_entries = 0
    for e in entries:
        cut = len(e)
        m_paren = re.search(r"\(", e)
        m_year = re.search(r"\b(19|20)\d{2}\b", e)
        if m_paren:
            cut = min(cut, m_paren.start())
        if m_year:
            cut = min(cut, m_year.start())
        author_seg = e[:cut]
        if re.search(r",\s+and\s+", author_seg, re.IGNORECASE):
            and_entries += 1
    if and_entries:
        out.append(Finding(
            dimension="reference",
            severity="info",
            message="Reference author connector",
            expected="Use '&' before the last author (APA 7th §9.7), e.g. 'Smith, J., & Jones, K.'",
            actual=f"{and_entries} entr{'y' if and_entries == 1 else 'ies'} use 'and' between authors instead of '&'",
            source="spec",
        ))
    return out


# ---------------------------------------------------------------------------
# 主检查
# ---------------------------------------------------------------------------

def _is_sorted_alphabetically(entries: list) -> bool:
    """保守判断参考文献条目是否按首作者姓氏字母序排列（APA 7th 要求）。

    不解析语义，仅比较每条起始 token（去掉前缀序号/标点）的小写形式。
    任一相邻逆序即判为未排序。用于「仅提示、不动内容」。
    """

    def _key(e: str) -> str:
        s = re.sub(r"^[\s\d\.\(\)\-–—]+", "", e.strip())
        return s.lower()

    if len(entries) < 2:
        return True
    keys = [_key(e) for e in entries]
    return all(keys[i] <= keys[i + 1] for i in range(len(keys) - 1))


def run_check(docx_path: str, target: TargetProfile) -> tuple[DocumentProfile, list[Finding]]:
    prof = read_docx(docx_path)
    findings: list[Finding] = []

    # —— 页边距 ——
    for key, side, _label in MARGIN_DIMS:
        if key in target.page:
            exp = target.page[key]
            origin = target.resolution.get(key, "spec")
            sev, actual = _cmp_margin(prof.margins.get(side), exp)
            sev = _downgrade(sev, origin, is_page_or_heading=True,
                             dim="margins", spec_key=target.spec_key)
            findings.append(Finding(
                dimension="margins",
                severity=sev,
                message=f"{side} margin",
                expected=f'{exp}"',
                actual=actual,
                source=origin,
            ))

    # —— 栏数（单栏/双栏；结构性维度）——
    # 与页边距/字体等「规范典型值仅作兜底、spec 来源降级 info」不同：选中某个规范作为
    # **来源**（而非只是引用风格）就是对该规范布局的显式声明，IEEE 双栏更是格式的定义性
    # 特征。故本维度**不**走 _downgrade 的 spec→info 软参考策略，不匹配照常 warn。
    # 学校模板来源（resolution="school"）/ 用户显式设置（"user"）本就是硬判，一致。
    if target.columns is not None:
        exp_cols = int(target.columns)
        origin = target.resolution.get("columns", "spec")
        act_cols = prof.columns if prof.columns is not None else 1
        sev = "pass" if act_cols == exp_cols else "warn"
        findings.append(Finding(
            dimension="columns",
            severity=sev,
            message="Column layout",
            expected="%d column%s" % (exp_cols, "" if exp_cols == 1 else "s"),
            actual="%d column%s" % (act_cols, "" if act_cols == 1 else "s"),
            source=origin,
        ))

    # —— 纸张大小（Letter / A4；**通用维度，不归属任何引用规范**）——
    # 北美（美/加）用 Letter 8.5×11，欧洲/中国/澳洲用 A4 210×297。同一篇 APA 论文
    # 在纽约交是 Letter、在伦敦交是 A4 —— 所以规范给的只是「典型值参考」，spec 来源
    # 一律降级 info（只提示，绝不因纸型给客户判红）；学校模板 / 问卷 / AI 显式给出
    # 时才硬判（那是客户自己声明的意图）。
    if target.paper_size:
        exp_paper = str(target.paper_size)
        origin = target.resolution.get("paper_size", "spec")
        act_name = getattr(prof, "paper_size", None)
        act_desc = describe_paper_size(
            getattr(prof, "paper_width_in", None),
            getattr(prof, "paper_height_in", None))
        if act_name is None and act_desc == "unknown":
            sev, actual = "info", "Could not read the page size"
        elif act_name == exp_paper:
            sev, actual = "pass", act_desc
        else:
            sev, actual = "warn", act_desc
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="paper_size", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="paper_size", severity=sev,
            message="Paper size",
            expected=exp_paper,
            actual=actual,
            source=origin,
        ))

    # —— 字体 ——
    if "font_family" in target.page and target.page["font_family"]:
        exp_font = target.page["font_family"]
        origin = target.resolution.get("font_family", "spec")
        act_font = prof.font_family or "unknown"
        ok = act_font.lower() == exp_font.lower() or exp_font.lower() in act_font.lower()
        # 🔴 APA 7th §2.19 允许 6 款字体（每款有自己的合规字号），spec 里的
        # page.font_family = Times New Roman 只是「该风格典型首选」，**不是唯一正确答案**。
        # 客户用 Calibri / Arial / Georgia 同样合规 → 不该判「字体不符」，更不该在修正时
        # 把客户的字体换成 Times New Roman（换字体是越权，且会毁掉中文文档的 eastAsia 字体）。
        # 字号由 font_size_by_font 按「该字体自己的规范字号」校准 —— 这才是配对表的用处。
        # 只对**规范来源**放宽：学校模板 / 客户明确要求的字体必须照办，绝不放宽。
        if not ok and origin == "spec":
            try:
                _allow = [f.lower() for f in (getattr(resolve_spec(target),
                                                      "font_allowlist", ()) or ())]
            except Exception:
                _allow = []
            if _allow and act_font.lower() in _allow:
                ok = True
        sev = "pass" if ok else "warn"
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="font", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="font", severity=sev,
            message="Body font",
            expected=exp_font, actual=act_font,
            source=origin,
        ))

    # —— 字号 ——
    # 🔴 字体命中规范的允许清单时，**该字体有自己的合规字号**（APA：Calibri 11 / TNR 12 /
    # Georgia 11 …），spec 里的「典型值 12pt」不再是正确答案 → 本维度让位给
    # font_size_by_font（逐 run 按各自字体判），避免同一件事出两条互相矛盾的结论。
    _size_superseded = False
    if "font_size_pt" in target.page:
        _size_origin = target.resolution.get("font_size_pt", "spec")
        if _size_origin == "spec":
            try:
                _pairs_pre = getattr(resolve_spec(target), "font_allowlist_sizes", {}) or {}
            except Exception:
                _pairs_pre = {}
            if size_for_font(prof.font_family, _pairs_pre):
                _size_superseded = True
    if "font_size_pt" in target.page and not _size_superseded:
        exp = target.page["font_size_pt"]
        origin = target.resolution.get("font_size_pt", "spec")
        sev, actual = _cmp_num(prof.font_size_pt, exp, SIZE_TOL_FAIL, SIZE_TOL_WARN, "pt")
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="font_size", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="font_size", severity=sev,
            message="Body font size",
            expected=f"{exp}pt", actual=actual,
            source=origin,
        ))

    # —— 行距 ——
    if "line_spacing" in target.page:
        exp = target.page["line_spacing"]
        origin = target.resolution.get("line_spacing", "spec")
        sev, actual = _cmp_num(prof.line_spacing, exp, SPACING_TOL_FAIL, SPACING_TOL_WARN, "x")
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="line_spacing", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="line_spacing", severity=sev,
            message="Line spacing",
            expected=f"{exp}x", actual=actual,
            source=origin,
        ))

    # —— 正文对齐（段落级；spec 来源降级 info，学校/问卷来源硬判）——
    if target.page.get("body_alignment"):
        exp_al = str(target.page["body_alignment"])
        origin = target.resolution.get("body_alignment", "spec")
        act_al = prof.body_alignment
        if act_al is None:
            sev, actual = "info", "Could not read (all paragraph alignment inherits from styles)"
        elif act_al == exp_al:
            sev, actual = "pass", ALIGN_LABEL.get(act_al, act_al)
        else:
            sev, actual = "warn", ALIGN_LABEL.get(act_al, act_al)
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="alignment", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="alignment", severity=sev, message="Body alignment",
            expected=ALIGN_LABEL.get(exp_al, exp_al), actual=actual, source=origin,
        ))

    # —— 正文首行缩进（段落级）——
    if target.page.get("first_line_indent_in") is not None:
        exp = float(target.page["first_line_indent_in"])
        origin = target.resolution.get("first_line_indent_in", "spec")
        sev, actual = _cmp_num(prof.first_line_indent_in, exp,
                               INDENT_TOL_FAIL, INDENT_TOL_WARN, "″")
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="first_line_indent", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="first_line_indent", severity=sev, message="Body first-line indent",
            expected=f'{exp}″', actual=actual, source=origin,
        ))

    # —— 页眉距页顶（MLA 9th 要求 ½″；页面维度，spec 来源降级 info，绝不 fail）——
    if target.page.get("header_distance_in") is not None:
        exp = float(target.page["header_distance_in"])
        origin = target.resolution.get("header_distance_in", "spec")
        act = prof.header_distance_in
        if act is None:
            sev, actual = "info", "Could not read (document has no section header distance set)"
        else:
            sev, actual = _cmp_num(act, exp, HEADER_TOL_FAIL, HEADER_TOL_WARN, "″")
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim="header_distance", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="header_distance", severity=sev, message="Header distance from top",
            expected=f'{exp}″', actual=actual, source=origin,
        ))

    # —— 段前 / 段后间距（段落级）——
    # v2.3.75：补上「段前」。旧版只有段后，等于段落间距永远只能核一半 ——
    # 学校模板写着「段前 6 磅」时检查器一个字不提，修正器也没法改。
    for _dim, _key, _msg, _attr in (
            ("space_before", "space_before_pt", "Space before paragraph", "space_before_pt"),
            ("space_after", "space_after_pt", "Space after paragraph", "space_after_pt")):
        if target.page.get(_key) is None:
            continue
        exp = float(target.page[_key])
        origin = target.resolution.get(_key, "spec")
        sev, actual = _cmp_num(getattr(prof, _attr, None), exp,
                               SPACE_TOL_FAIL, SPACE_TOL_WARN, "pt")
        sev = _downgrade(sev, origin, is_page_or_heading=True,
                         dim=_dim, spec_key=target.spec_key)
        findings.append(Finding(
            dimension=_dim, severity=sev, message=_msg,
            expected=f"{exp}pt", actual=actual, source=origin,
        ))

    # —— 论文大标题（v2.3.73）——
    # 有摘要的论文由 title_page 维度管；**没有摘要**的论文以前会把大标题当成章节
    # 一级标题，报告于是写「Level 1 heading format: Act Center 16pt」，
    # 修正就把论文大标题改成模板的 14pt 左对齐。现在单独说清楚：认出来了、
    # 它不参与层级比对 —— 客户才不会把「只查到 N 级」误读成漏检。
    if getattr(prof, "doc_title_index", None) is not None \
            and not getattr(prof, "title_page_detected", False):
        _dt_sz = getattr(prof, "doc_title_size_pt", None)
        findings.append(Finding(
            dimension="doc_title", severity="info",
            message="Document title",
            expected="",
            actual=("Found on page 1 (%spt, centered, bold) — it is the paper title, "
                    "so it was not counted as a heading level"
                    % (("%g" % _dt_sz) if _dt_sz else "?")),
            source="document",
        ))

    # —— 标题层级 ——
    if target.heading_levels and target.heading_levels > 0:
        origin = target.resolution.get("heading_levels", "spec")
        _pseudo_lvls = sorted(getattr(prof, "heading_formats", None) or {}) \
            if getattr(prof, "heading_formats_are_pseudo", False) else []
        if prof.heading_levels_used == 0 and not _pseudo_lvls:
            msg = "No Heading styles detected (manual bold may be used instead of styles)"
            sev = _downgrade("warn", origin, is_page_or_heading=True,
                             dim="heading", spec_key=target.spec_key)
        elif prof.heading_levels_used == 0:
            # 无 Heading 样式，但靠「整段加粗 + 字号档」和「章节编号结构」认出了 N 级手排标题。
            # 层级属推断口径，必须如实告知，不能装作读到了标题样式。
            # v2.3.97：不再只靠字号 —— 不加粗的编号小节行（"1. Introduction"）现在也会
            # 按编号序列认进层级，所以文案要把「编号」也说出来，否则客户会以为
            # 他的小节标题是凭字号猜的。
            msg = ("No Heading styles; %d level%s recognised from manual bold headings and "
                   "section numbering (levels are inferred)"
                   % (len(_pseudo_lvls), "" if len(_pseudo_lvls) == 1 else "s"))
            # 🔴 v2.3.95：这是**提醒**不是**违规**。标题已经认出来了、每级格式也按规范
            # 落了盘，唯一的不确定是「层级是按字号推的」。工具又造不出标题样式（那是
            # 改正文结构），判 warn 就永远修不掉 —— 客户点完修正复检还是同一句。
            # 老板 2026-10-23 四份产物实测正是如此（APA 两份修完仍报这条）。
            sev = "info"
        elif prof.heading_levels_used > target.heading_levels:
            msg = f"Detected {prof.heading_levels_used} levels (exceeds ≤{target.heading_levels})"
            sev = _downgrade("warn", origin, is_page_or_heading=True,
                             dim="heading", spec_key=target.spec_key)
        else:
            msg = f"Detected {prof.heading_levels_used} levels (styles: {', '.join(prof.heading_styles) or 'none'})"
            sev = _downgrade("pass", origin, is_page_or_heading=True,
                             dim="heading", spec_key=target.spec_key)
        findings.append(Finding(
            dimension="heading", severity=sev,
            message="Heading levels",
            expected=f"≤{target.heading_levels} levels",
            actual=msg,
            source=origin,
        ))

    # —— 每级标题的具体格式（字体 / 字号 / 对齐 / 加粗 / 斜体 / 缩进量）——
    # v2.3.65 新增维度。背景：旧版只查「层级数 ≤ N」，至于「一级标题该是黑体三号
    # 居中、二级该是楷体四号左对齐」，学校模板读进来、问卷填进来、AI 填进来，
    # 检查器全都不看 —— 客户填了 8 个字段，报告只回一句「层级数对」。
    # 只对**目标显式给出**的层级与属性做比对（None = 不规定 → 不碰），不臆造。
    heading_targets = getattr(target, "heading_formats", None) or {}
    if heading_targets:
        hf_origin = target.resolution.get("heading_formats", target.source or "school")
        want_map = resolve_heading_formats(heading_targets, None, target.heading_levels)
        have_map = prof.heading_formats or {}
        for lvl in sorted(want_map):
            want = want_map[lvl]
            got = have_map.get(lvl)
            if got is None:
                continue        # 文档里没这一级标题 → 不凭空判缺失
            exp_bits: list = []
            act_bits: list = []
            unreadable: list = []

            def _num_diff(label, want_v, got_v, unit=""):
                if want_v is None:
                    return
                if got_v is None:
                    unreadable.append(label)
                    return
                if abs(float(got_v) - float(want_v)) <= 1e-6:
                    return
                exp_bits.append("%s %g%s" % (label, float(want_v), unit))
                act_bits.append("%s %g%s" % (label, float(got_v), unit))

            if want.alignment:
                if got.alignment is None:
                    unreadable.append("alignment")
                elif got.alignment != want.alignment:
                    exp_bits.append(ALIGN_LABEL.get(want.alignment, want.alignment))
                    act_bits.append(ALIGN_LABEL.get(got.alignment, got.alignment))
            if want.font_family:
                if got.font_family is None:
                    unreadable.append("font")
                elif str(got.font_family).strip().lower() != str(want.font_family).strip().lower():
                    exp_bits.append(str(want.font_family))
                    act_bits.append(str(got.font_family))
            _num_diff("size", want.size_pt, got.size_pt, "pt")
            indent_want = want.indent_in if want.indent_in is not None else (
                0.5 if want.indent else None)
            _num_diff("indent", indent_want, got.indent_in, "″")
            for label, want_b, got_b in (("bold", want.bold, got.bold),
                                         ("italic", want.italic, got.italic)):
                if want_b is None:
                    continue
                if got_b is None:
                    # 目标规定了、文档读不到（run 与样式链都没写）→ 如实记「读不到」，
                    # 别静默跳过。旧写法直接 continue，于是「三级标题该斜体」这类要求在
                    # 报告里彻底消失，修正器却照样加斜体 —— 检查看不见的项被改了（v2.3.70）。
                    unreadable.append(label)
                    continue
                if bool(got_b) != bool(want_b):
                    exp_bits.append(label if want_b else "not " + label)
                    act_bits.append(label if got_b else "not " + label)
            # 编号形态（IEEE 的 I. / A. / 1) 全在 number_prefix 里）：v2.3.70 前
            # 检查器**从不检查**它，修正器却一直在补编号 —— 报告说「标题格式合规」，
            # 标题却多出了 "A."。现在缺编号要说出来，修正才有据可依。
            # 只报「缺」不报「形态不同」：作者手打的 "1. Introduction" 若被判成
            # 「应为 I.」，修正器会在前面再插一个前缀（I. 1. Introduction），宁可不碰。
            #
            # v2.3.84：编号与全大写都是**改文字**，与老板清单第 5 条冲突 →
            # policy.TEXT_IS_LOCKED 为 True 时不再计入"待修正"（exp_bits/act_bits），
            # 改走下面的 info 提示。否则就是"报了改不了"（点了修正一个字都不变）。
            _text_bits: list = []
            if want.number_prefix:
                if got.number_prefix is None:
                    _hint = _NUMBER_PREFIX_HINT.get(want.number_prefix,
                                                    want.number_prefix)
                    if TEXT_IS_LOCKED:
                        _text_bits.append("numbering (%s)" % _hint)
                    else:
                        exp_bits.append("numbering (%s)" % _hint)
                        act_bits.append("no numbering")
            # 全大写：只报「该大写却没大写」；反向（作者全大写、目标不要求）
            # 不报也不改 —— 小写化会破坏专有名词。读不到（文本无字母等）也不硬判。
            if want.uppercase and got.uppercase is False:
                if TEXT_IS_LOCKED:
                    _text_bits.append("uppercase")
                else:
                    exp_bits.append("uppercase")
                    act_bits.append("not uppercased")
            # 段前分页（v2.3.67）：三态。want 为 None=不规定；got 为 None=读不到（可能由
            # 样式承载）→ 不硬判，只记 unreadable；两者都非 None 且不一致才判 fail。
            if want.page_break_before is not None:
                if got.page_break_before is None:
                    unreadable.append("page break before")
                elif bool(got.page_break_before) != bool(want.page_break_before):
                    exp_bits.append("page break before" if want.page_break_before
                                    else "no page break before")
                    act_bits.append("page break before" if got.page_break_before
                                    else "no page break before")
            # 段前 / 段后 / 行距（v2.3.75）：三态。读侧永不返回 None（整条链没写就是
            # Word 缺省的 0 磅 / 1.0 倍），所以这里报出来的每一条都是**真实可见的差异**，
            # 修正器按同一条规则落盘。容差 0.1 磅（twips 换磅的舍入余量）。
            for _label, _wv, _gv in (
                    ("space before", getattr(want, "space_before_pt", None),
                     getattr(got, "space_before_pt", None)),
                    ("space after", getattr(want, "space_after_pt", None),
                     getattr(got, "space_after_pt", None))):
                if _wv is None:
                    continue
                if _gv is None:
                    unreadable.append(_label)
                    continue
                if abs(float(_gv) - float(_wv)) > 0.1:
                    exp_bits.append("%s %g pt" % (_label, float(_wv)))
                    act_bits.append("%s %g pt" % (_label, float(_gv)))
            _wls = getattr(want, "line_spacing", None)
            if _wls is not None:
                _gls = getattr(got, "line_spacing", None)
                if _gls is None:
                    # 固定值/最小值的行距无法折算成倍数 → 如实记「读不到」，不硬判
                    unreadable.append("line spacing")
                elif abs(float(_gls) - float(_wls)) > 0.01:
                    exp_bits.append("line spacing %gx" % float(_wls))
                    act_bits.append("line spacing %gx" % float(_gls))

            def _actual_text() -> str:
                bits = list(act_bits)
                if unreadable:
                    bits.append("could not read: " + ", ".join(unreadable))
                return ", ".join(bits) if bits else "Matches"

            if unreadable and not exp_bits:
                sev = _downgrade("info", hf_origin, is_page_or_heading=True,
                                 dim="heading", spec_key=target.spec_key)
                findings.append(Finding(
                    dimension="heading", severity=sev,
                    message="Level %d heading format" % lvl,
                    expected=describe_heading(want),
                    actual="Could not read: " + ", ".join(unreadable),
                    source=hf_origin,
                ))
                _append_text_lock_note(findings, lvl, _text_bits, hf_origin)
                continue
            sev = "pass" if not exp_bits else "fail"
            sev = _downgrade(sev, hf_origin, is_page_or_heading=True,
                             dim="heading", spec_key=target.spec_key)
            findings.append(Finding(
                dimension="heading", severity=sev,
                message="Level %d heading format" % lvl,
                expected=", ".join(exp_bits) if exp_bits else describe_heading(want),
                actual=_actual_text(),
                source=hf_origin,
            ))
            _append_text_lock_note(findings, lvl, _text_bits, hf_origin)

        # —— 超出目标声明范围的更深层级（v2.3.73）——
        # 目标声明了完整多级（APA/IEEE 五级）而文档更深（第 6 级）：检查要**明说**
        # 「第 6 级会按第 5 级规则处理」—— 旧版检查一个字不提、修正器却静默刷了一遍。
        # ⚠️ 只在 len(want_map) >= 2 时报：目标只声明 1 级（学校模板常见）时，
        # 更深层级**没有任何依据**，fixer 根本不会动它（v2.3.72 明确禁止拿一级格式套二级），
        # 这时写「会按第 1 级处理」就是假话。那种情形由上面的「层级数」结论负责说明。
        _declared = max(want_map) if want_map else 0
        _extra = sorted(set(have_map) - set(want_map))
        if _extra and _declared >= 1 and len(want_map) >= 2:
            _deep_want = want_map[_declared]
            for _lv in _extra:
                _got_x = have_map[_lv]
                _cur_bits = []
                for _lbl, _v, _unit in (("size", getattr(_got_x, "size_pt", None), "pt"),
                                        ("align", getattr(_got_x, "alignment", None), "")):
                    if _v:
                        _cur_bits.append("%s %s%s" % (_lbl, _v, _unit))
                findings.append(Finding(
                    dimension="heading",
                    severity=_downgrade("warn", hf_origin, is_page_or_heading=True,
                                        dim="heading", spec_key=target.spec_key),
                    message="Level %d heading format" % _lv,
                    expected=("Falls back to level %d rules (%s)"
                              % (_declared, describe_heading(_deep_want))),
                    actual=(", ".join(_cur_bits) or "Could not read"),
                    source=hf_origin,
                ))

    # —— 标题页格式（标题 / 作者单位的居中与加粗）——
    # 优先级：客户显式给出（学校模板实测 / 问卷 / AI）> 内置规范的 title_page。
    # 过去这条链只有 fixer 一半：spec 有 title_page 就**悄悄改**标题页，检查器从不报告、
    # 预览里也没有 —— 违反「检查→修正单一真相源」。现在先出结论，再允许施加。
    _tp_origin = None
    _tp_vals = {}
    for _k in ("title_page_title_centered", "title_page_title_bold",
               "title_page_author_centered",
               # v2.3.76：标题段的段前 / 段后
               "title_page_space_before_pt", "title_page_space_after_pt",
               # v2.3.114：题目字号 / 字体（模板驱动）
               "title_page_title_size_pt", "title_page_title_font"):
        _v = getattr(target, _k, None)
        if _v is not None:
            _tp_vals[_k] = _v
            _tp_origin = _tp_origin or target.resolution.get(_k, "school")
    # 🔴 v2.3.73：规范的 title_page 只在**规范基线**模式下兜底。非规范基线（学校模板 /
    # 手动 / AI）时 spec_key 为空，resolve_spec 只是按 reference_style 兜底取到整套规范 ——
    # 客户只选了「参考文献按 APA」，不该因此被要求「标题页居中加粗」。
    _spec_tp = (getattr(resolve_spec(target), "title_page", None)
                if spec_is_baseline(target) else None)
    if not _tp_vals and _spec_tp is not None:
        # 规范兜底：spec 里有声明就照它判，来源记 spec（页面维度 → 提示语气）
        if _spec_tp.title_centered is not None:
            _tp_vals["title_page_title_centered"] = bool(_spec_tp.title_centered)
        if _spec_tp.title_bold is not None:
            _tp_vals["title_page_title_bold"] = bool(_spec_tp.title_bold)
        if _spec_tp.author_centered is not None:
            _tp_vals["title_page_author_centered"] = bool(_spec_tp.author_centered)
        # v2.3.114：规范的标题字号也兜底带入（TitlePageFormat.title_size_pt）。
        if getattr(_spec_tp, "title_size_pt", None) is not None:
            _tp_vals.setdefault("title_page_title_size_pt", float(_spec_tp.title_size_pt))
        _tp_origin = "spec"
    if _tp_vals:
        _tp_origin = _tp_origin or "spec"
        if not getattr(prof, "title_page_detected", False):
            # 没标题页不算错：很多学校要求封面另附或直接首页开正文
            findings.append(Finding(
                dimension="title_page", severity="info",
                message="Title page format",
                expected=_tp_expected_text(_tp_vals),
                actual="No title page detected (the first page looks like body text)",
                source=_tp_origin,
            ))
        else:
            _bad, _act = [], []
            _want_center = _tp_vals.get("title_page_title_centered")
            _cur_al = getattr(prof, "tp_title_alignment", None)
            if _want_center is not None:
                if _cur_al is None:
                    _act.append("could not read the title alignment")
                elif (_cur_al == "center") != bool(_want_center):
                    # 文案说「文档哪里不对」，不是说「目标是怎样」：
                    # 要求居中而它没居中 → "not centered"。
                    _bad.append("not centered" if _want_center else "centered")
                    _act.append(ALIGN_LABEL.get(_cur_al, _cur_al))
                else:
                    _act.append(ALIGN_LABEL.get(_cur_al, _cur_al))
            _want_bold = _tp_vals.get("title_page_title_bold")
            _cur_b = getattr(prof, "tp_title_bold", None)
            if _want_bold is not None:
                if _cur_b is None:
                    _act.append("could not read the title weight")
                elif bool(_cur_b) != bool(_want_bold):
                    _bad.append("not bold" if _want_bold else "bolded")
                    _act.append("bold" if _cur_b else "not bold")
                else:
                    _act.append("bold" if _cur_b else "not bold")
            _want_ac = _tp_vals.get("title_page_author_centered")
            _cur_ac = getattr(prof, "tp_author_alignment", None)
            if _want_ac is not None:
                if _cur_ac is None:
                    _act.append("could not read the author-line alignment")
                elif (_cur_ac == "center") != bool(_want_ac):
                    _bad.append("author lines not centered" if _want_ac
                                else "author lines centered")
                    _act.append("authors " + ALIGN_LABEL.get(_cur_ac, _cur_ac))
                else:
                    _act.append("authors " + ALIGN_LABEL.get(_cur_ac, _cur_ac))
            # v2.3.76：标题段的段前 / 段后（磅）。与上面三项同属「Title page format」
            # 一条结论，客户读到的仍是「标题页格式」一行，不额外制造噪音。
            # ⚠️ 必须在算 _sev **之前**判定 —— 旧顺序会把段前段后不一致判成 pass，
            # 报告说「通过」、fixer 拿不到 fix 维度、客户点了修正却什么都没改。
            for _k, _label, _attr in (
                    ("title_page_space_before_pt", "title space before", "tp_title_space_before_pt"),
                    ("title_page_space_after_pt", "title space after", "tp_title_space_after_pt")):
                _w = _tp_vals.get(_k)
                if _w is None:
                    continue
                _got = getattr(prof, _attr, None)
                if _got is None:
                    _act.append("could not read the " + _label)
                    continue
                if abs(float(_got) - float(_w)) > 0.1:
                    _bad.append("%s %g pt (expected %g pt)"
                                % (_label, float(_got), float(_w)))
                else:
                    _act.append("%s %g pt" % (_label, float(_got)))
            # v2.3.114：题目字号（模板驱动）。题目段实测字号来自 prof.doc_title_size_pt；
            # 模板/问卷/AI 显式给了题目字号才判，否则不碰（避免把正文字号当错误）。
            if "title_page_title_size_pt" in _tp_vals:
                _w = _tp_vals["title_page_title_size_pt"]
                _got_sz = getattr(prof, "doc_title_size_pt", None)
                if _got_sz is None:
                    _act.append("could not read the title size")
                elif abs(float(_got_sz) - float(_w)) > 0.5:
                    _bad.append("title %g pt (expected %g pt)"
                                % (float(_got_sz), float(_w)))
                else:
                    _act.append("title %g pt" % float(_got_sz))
            _sev = "fail" if _bad else "pass"
            _sev = _downgrade(_sev, _tp_origin, is_page_or_heading=True,
                              dim="title_page", spec_key=target.spec_key)
            findings.append(Finding(
                dimension="title_page", severity=_sev,
                message="Title page format",
                expected=_tp_expected_text(_tp_vals),
                actual=("Title page: " + ", ".join(_bad)) if _bad else
                       (", ".join(_act) if _act else "Matches"),
                source=_tp_origin,
                # v2.3.114：挂 FixAction 让修正器能施加题目字号/字体（检查→修正单一真相源）。
                # 值取 _tp_vals 全量（含 size_pt/font），修正器按 target 三态幂等施加。
                fix=FixAction(dimension="title_page", value=dict(_tp_vals), note=""),
            ))

    # —— 首页结构（v2.3.91）——
    # 题名页独立成页 / 标题垂直位置 + 题名页零额外段距 / 摘要后正文另起一页。
    # 老板 2026-10-22 拿 v2.3.90 产物核对后拍板「三件一起做，清掉」——
    # 当时题目贴在顶边距第一行、题名页与摘要挤在同一页、还留着原文的 12/18pt 段距。
    # 全部只写 pageBreakBefore 与段距，**不新增段落或文字**（文字 100% 锁死不受影响）。
    # 与页码 / running head 同一口径：只有规范基线（客户真在规范下拉里选了内置规范）
    # 才判 warn 并绑定修正；学校模板 / 手填 / AI 导入一律只提示，模板说了算。
    _fm_spec = resolve_spec(target)
    _fm_base = spec_is_baseline(target)
    _has_tp = bool(getattr(prof, "title_page_detected", False))

    # ① 题名页独立成页
    # v2.3.93：客户显式要求（学校模板实测 / 手填 / AI 导入）优先于内置规范 ——
    # 「模板说了算」在这里同样成立；只有当客户没说话且真选了内置规范时才听规范。
    _tpn_exp = getattr(target, "title_page_new_page", None)
    _tpn_want = (_tpn_exp if _tpn_exp is not None
                 else (getattr(_fm_spec, "title_page_new_page", None)
                       if _fm_base else None))
    _tpn_origin = (target.resolution.get("title_page_new_page", "school")
                   if _tpn_exp is not None else "spec")
    if _tpn_want is not None and _has_tp:
        _fm_want = bool(_tpn_want)
        _fm_cur = getattr(prof, "tp_next_break_before", None)
        if _fm_cur is None:
            findings.append(Finding(
                dimension="title_page_new_page", severity="info",
                message="Title page on its own page",
                expected="The title page occupies its own page",
                actual="Could not read the paragraph that follows the title page",
                source=_tpn_origin,
            ))
        elif bool(_fm_cur) != _fm_want:
            findings.append(Finding(
                dimension="title_page_new_page",
                severity="warn",
                message="Title page on its own page",
                expected=("The title page occupies its own page (a page break follows it)"
                          if _fm_want else "No forced page break after the title page"),
                actual=("A page break already follows the title page" if _fm_cur
                        else "The next section continues on the title page"),
                source=_tpn_origin,
            ))
        else:
            findings.append(Finding(
                dimension="title_page_new_page", severity="pass",
                message="Title page on its own page",
                expected="The title page occupies its own page",
                actual="Matches",
                source=_tpn_origin,
            ))

    # ② 标题垂直位置 + 标题后空行 + 其余元素零额外段距（合成一条结论，客户读一行就够）
    # v2.3.92：判定数字全部来自 title_page_spacing_plan()（与 fixer 同一份计算），
    # 且严格照 APA 原话 —— 距顶 3~4 行是**区间**；标题与作者之间要**一个空行**（不是 0）。
    _plan = title_page_spacing_plan(target, _fm_spec, prof)
    if (_plan["before"] or _plan["after"] or _plan["author_zero"]) and _has_tp:
        _unit = float(_plan["unit_pt"] or 0.0)
        _tpl_bad, _tpl_act, _exp_bits = [], [], []
        # ① 标题距顶 3~4 行（区间）
        if _plan["before"]:
            _lo, _hi = _plan["before"]
            _blo, _bhi = _plan["before_lines"]
            _exp_bits.append("title %g-%g lines below the top margin" % (_blo, _bhi))
            _got_sb = getattr(prof, "tp_title_space_before_pt", None)
            if _got_sb is None:
                _tpl_act.append("could not read the title's space before")
            elif not (_lo <= float(_got_sb) <= _hi):
                _tpl_bad.append(
                    "title sits %.1f pt below the top margin (expected %g-%g lines = "
                    "%.0f-%.0f pt)" % (float(_got_sb), _blo, _bhi, _blo * _unit, _bhi * _unit))
            else:
                _tpl_act.append("title %.1f pt below the top margin" % float(_got_sb))
        # ② 标题与作者之间一个双倍行距空行 —— 真空段与段后间距**两种都认**
        if _plan["after"]:
            _alo, _ahi = _plan["after"]
            _exp_bits.append("one blank line after the title")
            _gap = getattr(prof, "tp_author_gap_paragraphs", None)
            if _gap is not None and int(_gap) >= 1:
                _tpl_act.append("blank line between the title and the author names")
            else:
                _got_sa = getattr(prof, "tp_title_space_after_pt", None)
                if _got_sa is None:
                    _tpl_act.append("could not read the title's space after")
                elif not (_alo <= float(_got_sa) <= _ahi):
                    _tpl_bad.append(
                        "no blank line between the title and the author names "
                        "(title space after %.1f pt, expected about %g line = %.0f pt)"
                        % (float(_got_sa), _plan["after_lines"], _plan["after_lines"] * _unit))
                else:
                    _tpl_act.append("blank line after the title")
        # ③ 其余元素之间 0 额外段距
        if _plan["author_zero"]:
            _exp_bits.append("no extra paragraph spacing")
            _mx = getattr(prof, "tp_author_space_max_pt", None)
            if _mx is None:
                _tpl_act.append("could not read the extra paragraph spacing")
            elif float(_mx) > 0.5:
                _tpl_bad.append(
                    "extra paragraph spacing %.1f pt between title page elements (expected 0)"
                    % float(_mx))
            else:
                _tpl_act.append("no extra paragraph spacing")
        _exp = "; ".join(_exp_bits) or "Matches"
        # v2.3.93：非基线下走到这里的只可能是 ③（其余元素零额外段距），
        # 而 ③ 在非基线下只有一个来源 —— 客户自己勾了「不留额外段距」。
        # 客户明确要求 → 判 warn 并可修（与页码、摘要另起一页同一口径）。
        _lay_origin = ("spec" if _fm_base
                       else target.resolution.get("title_page_zero_extra_spacing",
                                                  "school"))
        findings.append(Finding(
            dimension="title_page_layout",
            severity="warn" if _tpl_bad else "pass",
            message="Title page position and spacing",
            expected=_exp[:1].upper() + _exp[1:],
            actual=("Title page: " + ", ".join(_tpl_bad)) if _tpl_bad
            else (", ".join(_tpl_act) or "Matches"),
            source=_lay_origin,
        ))

    # ③ 摘要之后正文另起一页（v2.3.93：显式通道优先，同 ①）
    _txn_exp = getattr(target, "text_new_page", None)
    _tn_want = (_txn_exp if _txn_exp is not None
                else (getattr(_fm_spec, "text_new_page", None)
                      if _fm_base else None))
    _tn_origin = (target.resolution.get("text_new_page", "school")
                  if _txn_exp is not None else "spec")
    if _tn_want is not None and bool(getattr(prof, "has_abstract", False)):
        _tn_want = bool(_tn_want)
        _tn_cur = getattr(prof, "text_break_before", None)
        if _tn_cur is None:
            findings.append(Finding(
                dimension="text_new_page", severity="info",
                message="Text starts on a new page",
                expected="The body starts on a new page after the abstract",
                actual=_abstract_not_located(
                    prof, " — cannot tell whether the body starts on a new page"),
                source=_tn_origin,
            ))
        elif bool(_tn_cur) != _tn_want:
            findings.append(Finding(
                dimension="text_new_page",
                severity="warn",
                message="Text starts on a new page",
                expected=("The body starts on a new page after the abstract"
                          if _tn_want else "The body may continue right after the abstract"),
                actual=("A page break already follows the abstract" if _tn_cur
                        else "The body continues on the abstract page"),
                source=_tn_origin,
            ))
        else:
            findings.append(Finding(
                dimension="text_new_page", severity="pass",
                message="Text starts on a new page",
                expected="The body starts on a new page after the abstract",
                actual="Matches",
                source=_tn_origin,
            ))

    # ④ 题名页内容完整性（v2.3.92）—— **只提示，永不判红、永不代写**
    # 老板 2026-10-23 拍板「可以做提示」。APA 学生版题名页除版式外还要求课程号+课名、
    # 教师名、提交日期；这三项是**内容**，工具一个字都不会替客户写（文字 100% 锁死），
    # 只能列清楚「还差什么，请自己补」。规范没要求的绝不多加：专业版不查这三项。
    # v2.3.93：这三项是 **APA 学生版这一份规范** 对题名页的清单，不是普世要求 ——
    # 走学校模板 / 手填 / AI 时，客户有没有要求课程号/教师名/日期我们无从得知，
    # 去提示「你缺课程号」等于拿 APA 的清单硬套别人的封面。故仍只在规范基线下提示。
    if _has_tp and _fm_base and getattr(_fm_spec, "title_page_kind", "") == "student":
        _miss = title_page_missing_elements(prof, _fm_spec)
        if _miss:
            findings.append(Finding(
                dimension="title_page_content", severity="info",
                message="Title page content",
                expected=("An APA Style student title page also lists the course number "
                          "and name, the instructor name, and the assignment due date"),
                actual=("Not found on the title page: %s. This tool never adds text, so "
                        "please add these yourself" % ", ".join(_miss)),
                source="spec",
            ))
        else:
            findings.append(Finding(
                dimension="title_page_content", severity="pass",
                message="Title page content",
                expected="Course, instructor and due date are listed on the title page",
                actual="Matches",
                source="spec",
            ))

    # —— 参考文献（始终硬判，引用归规范）——
    ref_style = target.reference_style or "Other"
    # 用 spec_key 取完整规范：学生/专业共用 style_key="APA"，但 tier 标记
    # (running_head/expect_abstract/expect_author_note/running_head_max_chars) 按
    # spec_key 区分。若只按 ref_style("APA") 取，会永远落到学生版，专业版差异项不生效。
    # 单一真相源：checker 与 fixer 共用 resolve_spec，杜绝「检查/修正用不同规范」的漂移
    spec = resolve_spec(target)
    ref_label = "Works Cited" if spec.reference.style_key == "MLA" else "References"
    ref_source = target.resolution.get("reference_style", "spec")
    ref_finding = _check_reference(ref_style, prof.reference_entries, spec, ref_source)
    # 顺序编号制（IEEE 等）且存在未编号条目 → 编号缺失本身是硬问题，须提示并补 [n]，
    # 与 fixer 的 reference_numbered 修正严格一致（check→fix 单一真相源）。
    # 覆盖 _check_reference 可能给出的 pass/info（例如已有 [3] 时整体被误判为「已合规」），
    # 否则 fixer 会因 fix_dims 不含 reference_numbered 而跳过编号，造成「该补却没补」。
    if target.reference_numbered and prof.reference_entries:
        # v2.3.100（C 批 D-16）：Word **自动编号**也算已编号。
        # IEEE 官方会议模板的 references 样式用 `<w:numPr/>` 列表编号，条目文本里
        # 根本没有 "[1]"，旧判定只看文本 → 官方模板 16/16 全被判"未编号"。
        _auto_no = list(getattr(prof, "reference_entries_numbered", None) or [])
        _missing = [
            t for _i, t in enumerate(prof.reference_entries)
            if t.strip() and not _has_leading_number(t)
            and not (_i < len(_auto_no) and _auto_no[_i])
        ]
        if _missing:
            ref_finding = Finding(
                dimension="reference",
                severity="fail" if len(_missing) == len(prof.reference_entries) else "warn",
                message="Reference sequential numbering",
                expected="Each reference entry should carry a sequential [n] code (e.g., [1], [2])",
                actual=f"{len(_missing)} of {len(prof.reference_entries)} entries not numbered",
                source=ref_source,
            )
        elif _auto_no and sum(1 for x in _auto_no if x) >= 2:
            # v2.3.100（C 批 D-16）：条目由 Word **自动编号**（段落级或样式级 numPr）时，
            # 文本里本就没有 "[n]"。_check_reference 的体例检查只看文本，会误报
            # "[IEEE] No numbered [n] form detected" —— 官方 IEEE 模板就是这样被打红的。
            # 自动编号已确认 ≥2 条即视为合规，覆盖体例检查的误报。
            ref_finding = Finding(
                dimension="reference",
                severity="pass",
                message="Reference sequential numbering",
                expected="Entries numbered in order of citation",
                actual=f"{len(prof.reference_entries)} entries, auto-numbered by Word list",
                source=ref_source,
            )
    # 🔴 v2.3.95：反方向同样要判 —— 规范**明确不编号**（APA 7 §9.44：参考文献条目
    # 不加序号；MLA 的 Works Cited 同理，条目按作者姓氏字母序、以姓名开头），
    # 而客户条目带着 `1.` / `[1]` / `(1)` → 学生最常犯的错之一。
    # 旧版只做「该补却没补」，反方向一个字都不说、也一个字不改：老板 2026-10-23
    # 拿 v2.3.93 的四份产物（APA 学生版 / APA 专业版 / IEEE / MLA）核对，APA 与 MLA
    # 两份的条目仍是 `1. Ivanov, D. (2020)…`，报告却全绿 —— 客户按 APA 交上去就是错的。
    # 修正动作只动段首那个列表记号（`_drop_leading_chars`），条目正文一个字不碰。
    elif target.reference_numbered is False and prof.reference_entries:
        _numbered = [t for t in prof.reference_entries
                     if t.strip() and _has_leading_number(t)]
        if _numbered:
            _all = len(_numbered) == len([t for t in prof.reference_entries if t.strip()])
            ref_finding = Finding(
                dimension="reference",
                severity="fail" if _all else "warn",
                message="Reference entries carry list numbers",
                expected=("%s does not number reference entries — each entry starts "
                          "with the author's name" % ref_style),
                actual=("%d of %d entries start with a list number (e.g. %r)"
                        % (len(_numbered), len(prof.reference_entries),
                           _numbered[0].strip()[:34])),
                source=ref_source,
            )
            ref_finding.fix = FixAction(dimension="reference_numbered", value=False)
    findings.append(ref_finding)

    # —— APA 参考文献内容级提示（仅信息，不误伤官方样例；见 _apa_reference_content_hints）——
    if spec.reference.style_key == "APA" and prof.reference_entries:
        findings.extend(_apa_reference_content_hints(
            prof.reference_entries, prof.reference_entries_italic))

    # —— 参考文献悬挂缩进（引用维度，硬判但容差略宽）——
    if target.reference_hanging_indent_in is not None:
        exp_h = float(target.reference_hanging_indent_in)
        act_h = prof.reference_hanging_indent_in
        if act_h is None:
            sev_h, actual_h = "info", "Could not read (no reference section detected or no explicit indent)"
        elif exp_h <= 0.01:
            # v2.3.100（C 批 D-16 / 裁定 C-10）：`hanging_indent_in = 0` 的规范（IEEE）
            # 语义不是"悬挂必须为 0"，而是「编号自成列、悬挂于条目正文之外」
            # —— IEEE Reference Guide 原文 "numbers … form a column of their own,
            # **hanging out beyond the body** of the reference"。
            # 所以只要存在**正向悬挂**（编号比正文更靠左）即合规；0 才算不符。
            if act_h > 0.1:
                sev_h, actual_h = "pass", f'{act_h}″ (number column, hanging beyond the body)'
            else:
                sev_h, actual_h = "warn", f'{act_h}″ (no hanging: number column not separated)'
        else:
            diff = abs(act_h - exp_h)
            if diff <= REF_HANG_TOL_WARN:
                sev_h, actual_h = "pass", f'{act_h}″'
            elif diff <= REF_HANG_TOL_FAIL:
                sev_h, actual_h = "warn", f'{act_h}″ (deviation {diff:.2f}″)'
            else:
                sev_h, actual_h = "fail", f'{act_h}″ (deviation {diff:.2f}″)'
        findings.append(Finding(
            dimension="reference_hanging_indent", severity=sev_h,
            message="Reference hanging indent",
            expected=f'{exp_h}″', actual=actual_h,
            source=target.resolution.get("reference_hanging_indent_in", "spec"),
        ))

    # —— 参考文献条目自身的排版（字体 / 字号 / 行距）——
    # v2.3.65 新增。学校模板里「条目五号宋体、单倍行距」是硬要求，且常与正文不同口径；
    # 这几项以前 target 根本承载不了（只有 style/悬挂/编号），模板读到了也无处放 → 无人检查。
    for _key, _dim, _label, _unit, _tol in (
            ("reference_font_family", "reference_font", "Reference entry font", "", None),
            ("reference_size_pt", "reference_size", "Reference entry font size", "pt",
             (SIZE_TOL_FAIL, SIZE_TOL_WARN)),
            ("reference_line_spacing", "reference_spacing", "Reference entry line spacing", "x",
             (SPACING_TOL_FAIL, SPACING_TOL_WARN)),
            # v2.3.75：条目之间的段前/段后（「条目之间不空行」是学校硬要求里常见的一条）
            ("reference_space_before_pt", "reference_space_before",
             "Reference entry space before", "pt", (SPACE_TOL_FAIL, SPACE_TOL_WARN)),
            ("reference_space_after_pt", "reference_space_after",
             "Reference entry space after", "pt", (SPACE_TOL_FAIL, SPACE_TOL_WARN))):
        _want = getattr(target, _key, None)
        if _want is None:
            continue
        _origin = target.resolution.get(_key, "school")
        if _key == "reference_font_family":
            _got = prof.reference_font_family
            if _got is None:
                # 区分「没有参考文献段」与「条目没显式设字体」——后者不能当缺陷
                _sev = "info"
                _act = ("Could not read (the entries inherit their font, no explicit font set)"
                        if prof.reference_entries else
                        "Could not read (no reference section detected)")
            elif str(_got).strip().lower() == str(_want).strip().lower():
                _sev, _act = "pass", str(_got)
            else:
                _sev, _act = "fail", str(_got)
            _exp = str(_want)
        else:
            _got = getattr(prof, _key, None)
            if _got is None:
                # 读不到就只提示，不判 warn/fail（条目常继承正文字号，读不到不等于不合规）
                _sev = "info"
                _act = ("Could not read (the entries inherit this from the body text)"
                        if prof.reference_entries else
                        "Could not read (no reference section detected)")
            else:
                _sev, _act = _cmp_num(_got, float(_want), _tol[0], _tol[1], _unit)
            _exp = "%g%s" % (float(_want), _unit)
        findings.append(Finding(
            dimension=_dim, severity=_sev, message=_label,
            expected=_exp, actual=_act, source=_origin,
        ))

    # —— 参考文献：另起一页 / 标题加粗 / 标题居中 ——
    # 目标显式给出（学校模板实测 / 问卷 / AI）时硬判并可修正；未给出时保持下面
    # 「仅按 APA/MLA 规范提示」的旧行为，不越界。
    if target.reference_new_page is not None:
        _want_np = bool(target.reference_new_page)
        _got_np = prof.reference_new_page
        _origin = target.resolution.get("reference_new_page", "school")
        if _got_np is None:
            _sev, _act = "info", "No reference section detected"
        elif bool(_got_np) == _want_np:
            _sev, _act = "pass", "Starts on a new page" if _want_np else "Does not force a new page"
        else:
            _sev, _act = "fail", ("Does not begin on a new page" if _want_np
                                  else "Begins on a new page (a page break is present)")
        findings.append(Finding(
            dimension="reference_new_page", severity=_sev,
            message="References on new page",
            expected="begin on a new page" if _want_np else "do not force a new page",
            actual=_act, source=_origin,
        ))

    if target.reference_heading_bold is not None or target.reference_heading_centered is not None:
        _origin = target.resolution.get(
            "reference_heading_bold",
            target.resolution.get("reference_heading_centered", "school"))
        if not prof.reference_entries:
            findings.append(Finding(
                dimension="reference_heading", severity="info",
                message="Reference heading format",
                expected=", ".join(filter(None, [
                    None if target.reference_heading_bold is None else
                    ("bold" if target.reference_heading_bold else "not bold"),
                    None if target.reference_heading_centered is None else
                    ("centered" if target.reference_heading_centered else "left-aligned")])),
                actual="No reference section detected", source=_origin,
            ))
        else:
            _want_txt = ", ".join(filter(None, [
                None if target.reference_heading_bold is None else
                ("bold" if target.reference_heading_bold else "not bold"),
                None if target.reference_heading_centered is None else
                ("centered" if target.reference_heading_centered else "left-aligned")]))
            # 🔴 信任 Heading 样式（v2.3.116 接线后必须保留这条保护）：
            # 官方样例（APA / MLA）用 Heading 1 承载 "References"，粗体/居中来自**样式定义**，
            # python-docx 读不出样式的解析值 → 直接判 pass，绝不把金标准论文误判红。
            # 「只提示」旧分支本来就有这条，接线时漏掉会把官方样例判 fail（负作用）。
            if prof.reference_heading_is_style:
                findings.append(Finding(
                    dimension="reference_heading", severity="pass",
                    message="Reference heading format",
                    expected=_want_txt + " (supplied by the Heading style)",
                    actual="Uses a Heading style", source=_origin,
                ))
            else:
                _bad = []
                if target.reference_heading_bold is not None \
                        and bool(prof.reference_heading_bold) != bool(target.reference_heading_bold):
                    _bad.append("bolded" if prof.reference_heading_bold else "not bold")
                if target.reference_heading_centered is not None \
                        and bool(prof.reference_heading_centered) != bool(target.reference_heading_centered):
                    _bad.append("centered" if prof.reference_heading_centered else "not centered")
                findings.append(Finding(
                    dimension="reference_heading",
                    severity="fail" if _bad else "pass",
                    message="Reference heading format",
                    expected=_want_txt,
                    actual=("The heading is " + " and ".join(_bad)) if _bad else "Matches",
                    source=_origin,
                ))

    # —— APA 专属保守检查（仅当 spec 显式声明时生效；其余风格取中性默认值自动跳过）——
    # 这些检查只做「提示」，不产生 fail，也不修改文档内容（符合「模板驱动 / 只提示」原则）。
    # 页码 / running head 属页面维度，来源为 spec → 按代码约定降级为 info（不强制）。
    apa_spec = spec  # 当前 spec 已按 target.reference_style 取出（非 APA 时各字段为中性默认值）
    # —— 页码位置（v2.3.66 重写）——
    # 三处旧错：
    #   ① 只判「含页码的段落是否右对齐」，**不分页眉/页脚** → 页脚右对齐也判符合 APA top-right；
    #   ② 客户手打「第 1 页」被判「没有页码」→ 一修就出现两个页码；
    #   ③ 对齐读不到（继承样式）时落进 else 报 pass → 假通过。
    # 现在按 PAGE_NUMBER_LAYOUT 把「位置」与「对齐」一起比，且区分手打/自动。
    _pn_pos = target.page_number_position or apa_spec.page_number_position
    _pn_res = origin_of_source(target.resolution.get("page_number_position"))
    _pn_explicit = _pn_res in ("user", "school", "ai_import")
    _pn_origin = _pn_res or "spec"
    if _pn_pos:
        _want_in, _want_align = PAGE_NUMBER_LAYOUT.get(_pn_pos, ("header", "right"))
        _pos_label = {
            "top-right": "top-right of the header",
            "bottom-center": "bottom-center of the footer",
            "bottom-right": "bottom-right of the footer",
        }.get(_pn_pos, _pn_pos)
        if not prof.page_number_present:
            findings.append(Finding(
                dimension="page_number",
                severity=_page_number_severity(_pn_origin, target),
                message="Page number",
                expected=f"Page number at {_pos_label}",
                actual="No page number found in the header or footer",
                source=_pn_origin,
            ))
        else:
            _bad = []
            if prof.page_number_in and prof.page_number_in != _want_in:
                _bad.append("placed in the " + prof.page_number_in)
            if prof.page_number_alignment and prof.page_number_alignment != _want_align:
                _bad.append(prof.page_number_alignment + "-aligned")
            if _bad:
                findings.append(Finding(
                    dimension="page_number",
                    severity=_page_number_severity(_pn_origin, target),
                    message="Page number position",
                    expected=_pos_label,
                    actual="Page number is " + " and ".join(_bad),
                    source=_pn_origin,
                ))
            else:
                findings.append(Finding(
                    dimension="page_number", severity="pass",
                    message="Page number",
                    expected=_pos_label,
                    actual="Present at " + _pos_label,
                    source=_pn_origin,
                ))
        # 手打页码单独给一条提示：修正时会**原地**把数字换成 PAGE 域（保留「第…页」前后缀），
        # 不是再加一条 —— 客户必须知道自己那串数字被动过。
        if prof.page_number_present and prof.page_number_is_static:
            findings.append(Finding(
                dimension="page_number", severity="info",
                message="Page number is typed, not automatic",
                expected="An automatic PAGE field (so Word renumbers every page by itself)",
                actual='Typed page number "%s" — the fix replaces only the digits with a PAGE '
                       "field and keeps the surrounding text" % prof.page_number_static_text,
                source="doc",
            ))

    # 目标已显式给出「另起一页」时，上面已经硬判并可修，这里不再重复提示
    if apa_spec.reference_new_page and target.reference_new_page is None:
        if prof.reference_new_page is None:
            findings.append(Finding(
                dimension="reference_new_page", severity="info",
                message=f"{ref_label} on new page",
                expected=f"{ref_label} begin on a new page",
                actual="No reference section detected",
                source="spec",
            ))
        elif not prof.reference_new_page:
            findings.append(Finding(
                dimension="reference_new_page", severity="info",
                message=f"{ref_label} on new page",
                expected=f"{ref_label} begin on a new page",
                actual=f"{ref_label} heading does not begin on a new page (insert a page break before {ref_label})",
                source="spec",
            ))
        else:
            findings.append(Finding(
                dimension="reference_new_page", severity="pass",
                message=f"{ref_label} on new page",
                expected=f"{ref_label} begin on a new page",
                actual="Starts on a new page",
                source="spec",
            ))

    # —— 字体白名单（APA 7th §2.19）：**字体 + 字号配对**，不是一串字体名 ——
    # 旧版只比字体名：客户用 Calibri 12pt（APA 要求 Calibri 11pt）会被判「合格」。
    # 修正策略（老板 2026-10-19 拍板）：**保留客户选的字体，只把字号校正到该字体的
    # 规范值** —— 替客户换字体是越权（6 款都合规，没有唯一正确答案），改字号才有。
    # 白名单之外的字体绝不猜字号（只提示），客户显式给了正文字号时也不抢（由 font_size 管）。
    allow_pairs = getattr(apa_spec, "font_allowlist_sizes", None) or {}
    font_size_by_font: dict = {}          # 字体 → 该字体应有的规范字号（只含要改的）
    if apa_spec.font_allowlist and prof.font_family:
        allowed = [f.lower() for f in apa_spec.font_allowlist]
        # 按「每种字体各自的众数字号」判，而不是只看全篇众数：
        # 正文里混用 Times New Roman（APA 要 12pt）与 Calibri（要 11pt）时，
        # 只看众数会把 Calibri 那几段彻底漏掉。
        # 🔴 D-06：全文字号若都靠样式继承，font_size_by_family 会是空、font_size_pt 会是
        # None。旧代码直接 {字体: None} 落到下面的 "%g" → TypeError，整份检查崩掉
        # （官方 APA 样例 3 份实测崩溃）。这里**过滤掉 None**，宁可"读不到"也不崩。
        _fam_sizes = {_k: _v for _k, _v in
                      (getattr(prof, "font_size_by_family", None) or {}).items()
                      if _v is not None}
        if not _fam_sizes and prof.font_size_pt is not None:
            _fam_sizes = {prof.font_family: prof.font_size_pt}
        _outside = [f for f in _fam_sizes if f.lower() not in allowed]
        if _outside:
            findings.append(Finding(
                dimension="font_allowlist", severity="info",
                message="Body font (APA allowlist)",
                expected="One of: " + ", ".join(
                    f"{k} {_pt_txt(v)}pt" for k, v in allow_pairs.items()) if allow_pairs
                else "One of: " + ", ".join(apa_spec.font_allowlist),
                actual=", ".join(_outside) + " (APA permits a range of readable fonts; "
                       "verify your school's requirement)",
                source="spec",
            ))
        # 「客户显式给了正文字号」时不抢：那种情况由 font_size 维度硬判硬改。
        # 注意必须看 resolution —— 选 APA 规范时 spec 自带 font_size_pt=12，
        # 那是「该风格典型值」，不属于客户显式要求，不能拿它挡掉字体配对的字号修正。
        _size_owned = origin_of_source(
            target.resolution.get("font_size_pt")) in ("user", "school", "ai_import")
        _cur_map: dict = {}
        if not _size_owned:
            for _fam, _cur in _fam_sizes.items():
                if _fam.lower() not in allowed:
                    continue
                _want = size_for_font(_fam, allow_pairs)
                if not _want or not _cur:
                    continue
                if abs(float(_cur) - float(_want)) > 0.01:
                    font_size_by_font[_fam] = float(_want)
                    _cur_map[_fam] = float(_cur)
        if font_size_by_font:
            # v2.3.100（落地裁定 C-4）：字号与 §2.19 白名单配对不符时**降为 info、不自动修**。
            # 原因：APA §2.19 允许 6 款字体各自配对字号（TNR 12 / Calibri 11 / Arial 11 …），
            # 而**官方学生样例自己就是 TNR 11pt** —— 按 warn 判等于把官方合规样例判红
            # （实测 13 份里 5 份命中本维度）。规则保留（仍算不符），只把严重度降下来，
            # 并明确告知客户「换字体就能合规」，让他自己选。
            # 🔴 C-4：**不挂 FixAction**。旧实现挂了一个 value=整张 6 款字体配对表的
            # FixAction，fixer 据此把客户字号刷成配对值 —— 那既替客户做了字体决定，
            # 也正中「官方样例 TNR 11pt 被判红」。现在只提示，客户自己决定。
            findings.append(Finding(
                dimension="font_size_by_font", severity="info",
                message="Body text size for "
                        + ("this font" if len(font_size_by_font) == 1 else "these fonts"),
                expected=", ".join("%s %g pt" % (f, s) for f, s in font_size_by_font.items()),
                actual=", ".join("%s %g pt" % (f, s) for f, s in _cur_map.items())
                       + " (APA pairs each allowed font with a size, e.g. Times New Roman "
                         "12 pt / Calibri, Arial or Georgia 11 pt; switching font is an "
                         "option too)",
                source="spec",
            ))
        elif not _outside:
            findings.append(Finding(
                dimension="font_allowlist", severity="pass",
                message="Body font (APA allowlist)",
                expected="One of: " + ", ".join(
                    f"{k} {_pt_txt(v)}pt" for k, v in allow_pairs.items()) if allow_pairs
                else "One of: " + ", ".join(apa_spec.font_allowlist),
                # D-06：_fam_sizes 可能为空（字号全靠样式继承）→ 空字符串会让报告出现
                # 一个没有 actual 的 pass 条目，客户看不懂。给明确兜底文案。
                actual=(", ".join("%s %s pt" % (f, _pt_txt(s))
                                  for f, s in _fam_sizes.items())
                        or "font size is inherited from styles (no explicit run size)"),
                source="spec",
            ))

    if apa_spec.running_head and getattr(apa_spec, "running_head_kind", "apa") == "apa":
        # 仅专业版 running_head=True 且 kind="apa" 才进此分支；学生版 running_head=False 完全跳过
        # （学生版不要求页眉短标题，绝不被误报「缺 running head」）。
        # 🔴 v2.3.90：规范基线时升为 warn 且**可修**。此前恒为 info，于是 apa-stu 与
        # apa-p 的产物 document.xml 逐字节 100% 相同（实测 95272 字节全等）——
        # 清单第②条「首页按 APA-Stu/APA-Pro 分别处理」等于没做。
        # 现在基线模式下由 fixer 从**客户自己的论文标题**派生短标题写入页眉（正文不动）。
        _rh_sev = "warn" if spec_is_baseline(target) else "info"
        rh = prof.header_text or ""
        if not rh:
            findings.append(Finding(
                dimension="running_head", severity=_rh_sev,
                message="Running head",
                expected=f"Short title in the header (≤ {apa_spec.running_head_max_chars} characters) + page number",
                actual="No running-head text detected in the header (APA Professional requires a short title)",
                source="spec",
            ))
        elif apa_spec.running_head_max_chars and len(rh) > apa_spec.running_head_max_chars:
            findings.append(Finding(
                dimension="running_head", severity=_rh_sev,
                message="Running head length",
                expected=f"≤ {apa_spec.running_head_max_chars} characters",
                actual=f"{len(rh)} characters: {rh[:60]}{'…' if len(rh) > 60 else ''}",
                source="spec",
            ))
        else:
            findings.append(Finding(
                dimension="running_head", severity="pass",
                message="Running head length",
                expected=f"≤ {apa_spec.running_head_max_chars} characters",
                actual=f"{len(rh)} characters",
                source="spec",
            ))

    if spec.running_head and getattr(spec, "running_head_kind", "apa") == "mla":
        # MLA running head：姓氏 + 空格 + 页码，每页右上（Purdue OWL general_format）。
        # 页码常以 PAGE 域形式存在（header_text 中不含数字），故同时看 header_text 与 page_number_present。
        # 🔴 v2.3.90：与 APA 专业版同理，规范基线时可修 —— 姓氏从客户**自己的姓名行**派生。
        _mla_sev = "warn" if spec_is_baseline(target) else "info"
        rh = prof.header_text or ""
        name_part = re.sub(r"\d+", "", rh).strip()
        has_name = bool(re.search(r"[A-Za-z]", name_part)) or (rh.strip() != "" and not rh.strip().isdigit())
        has_page = bool(prof.page_number_present) or bool(re.search(r"\d", rh))
        if not rh.strip() and not prof.page_number_present:
            findings.append(Finding(
                dimension="running_head", severity=_mla_sev,
                message="Running head",
                expected="Your last name followed by the page number, in the upper-right corner of every page",
                actual="No header text detected (MLA requires a running head with your last name + page number)",
                source="spec",
            ))
        elif not has_name:
            findings.append(Finding(
                dimension="running_head", severity=_mla_sev,
                message="Running head",
                expected="Your last name + page number (e.g., 'Smith 1')",
                actual=f"Header text detected: {rh[:60]!r}; MLA expects your last name followed by the page number",
                source="spec",
            ))
        elif not has_page:
            findings.append(Finding(
                dimension="running_head", severity=_mla_sev,
                message="Running head",
                expected="Your last name + page number",
                actual="Running head has a name but no page number detected",
                source="spec",
            ))
        else:
            findings.append(Finding(
                dimension="running_head", severity="pass",
                message="Running head",
                expected="Your last name + page number",
                actual=rh[:60] or "(name + page number field)",
                source="spec",
            ))

    # —— 摘要页特化（v2.3.67）——
    # 摘要正文口径与正文相反：APA 摘要**不首行缩进**（顶格）、专业版摘要**另起一页**。
    # 以前「正文首行缩进 0.5″」会把摘要段一起刷成 0.5″（越权），「参考文献另起一页」也不管摘要。
    # 目标显式给出（spec/school/user/ai）才判；None = 不规定 → 不碰。
    _abs_fli = target.abstract_first_line_indent_in
    _abs_np = target.abstract_new_page
    # 目标没显式给摘要字段、但规范带了（如 APA 专业版 abstract_new_page=True）→ 用规范值。
    # 🔴 v2.3.96：这层兜底**只在规范基线下**成立（客户真在规范下拉里选了内置规范）。
    # 学校模板 / 手填 / AI 导入三个通道里，客户只说了「参考文献体例用 APA」就足以让
    # resolve_spec 取回整套 APA —— 摘要顶格、摘要另起一页随之被判，可它们的来源是 spec
    # （非基线不绑 FixAction）→ 客户看到一条红杠、点了修正却不消失，纯属冤枉。
    # v2.3.96 之前这条被「摘要定位不到 → 只给 info」掩盖着，摘要能定位后立刻暴露。
    # 与 page_number / title_page_new_page 同口径：谁说话就听谁的，没提就闭嘴。
    if _abs_fli is None and spec_is_baseline(target):
        _abs_fli = apa_spec.abstract_first_line_indent_in
    if _abs_np is None and spec_is_baseline(target):
        _abs_np = apa_spec.abstract_new_page
    if _abs_fli is not None or _abs_np is not None:
        # 直接用 prof 已采样值（不重复读文件）；摘要段是否存在由采样值是否 None 判断
        _abs_fli_origin = target.resolution.get("abstract_first_line_indent_in") \
            if target.abstract_first_line_indent_in is not None else "spec"
        _abs_np_origin = target.resolution.get("abstract_new_page") \
            if target.abstract_new_page is not None else "spec"
        if _abs_fli is not None:
            _cur_fli = prof.abstract_first_line_indent_in
            if _cur_fli is None:
                # 无摘要段 → 不判缺失（很多论文摘要与正文同页/无独立摘要段）
                findings.append(Finding(
                    dimension="abstract_indent", severity="info",
                    message="Abstract first-line indent",
                    expected="Abstract text is flush left (no first-line indent)"
                    if abs(float(_abs_fli)) < 1e-3 else
                    f"Abstract text first-line indent = {_abs_fli:g}″",
                    actual=_abstract_not_located(prof, " (nothing to change)"),
                    source=_abs_fli_origin,
                ))
            elif abs(float(_cur_fli) - float(_abs_fli)) > 0.01:
                findings.append(Finding(
                    dimension="abstract_indent", severity="warn",
                    message="Abstract first-line indent",
                    expected=f"{_abs_fli:g}″",
                    actual=f"{_cur_fli:g}″",
                    source=_abs_fli_origin,
                ))
            else:
                findings.append(Finding(
                    dimension="abstract_indent", severity="pass",
                    message="Abstract first-line indent",
                    expected=f"{_abs_fli:g}″",
                    actual=f"{_cur_fli:g}″",
                    source=_abs_fli_origin,
                ))
        if _abs_np is not None:
            _cur_np = prof.abstract_new_page
            if _cur_np is None:
                findings.append(Finding(
                    dimension="abstract_new_page", severity="info",
                    message="Abstract on new page",
                    expected="Abstract begins on a new page"
                    if _abs_np else "Abstract does not start on a new page",
                    actual=_abstract_not_located(prof),
                    source=_abs_np_origin,
                ))
            elif bool(_cur_np) != bool(_abs_np):
                findings.append(Finding(
                    dimension="abstract_new_page", severity="warn",
                    message="Abstract on new page",
                    expected="begins on a new page" if _abs_np else "continues on the same page",
                    actual="begins on a new page" if _cur_np else "continues on the same page",
                    source=_abs_np_origin,
                ))
            else:
                findings.append(Finding(
                    dimension="abstract_new_page", severity="pass",
                    message="Abstract on new page",
                    expected="begins on a new page" if _abs_np else "continues on the same page",
                    actual="Matches",
                    source=_abs_np_origin,
                ))

    # —— 摘要正文段的段前 / 段后（v2.3.76）——
    # 学校模板常写死「摘要上方留 12 磅、下方留 6 磅」；此前只有缩进与另起一页，
    # 这一维度既读不到也改不了。目标显式给出才判，None = 不规定。
    for _k, _dim, _label, _attr in (
            ("abstract_space_before_pt", "abstract_space_before",
             "Abstract space before", "abstract_space_before_pt"),
            ("abstract_space_after_pt", "abstract_space_after",
             "Abstract space after", "abstract_space_after_pt")):
        _w = getattr(target, _k, None)
        if _w is None:
            continue
        _o = target.resolution.get(_k, "school")
        _got = getattr(prof, _attr, None)
        if _got is None:
            findings.append(Finding(
                dimension=_dim, severity="info", message=_label,
                expected="%g pt" % float(_w),
                actual=_abstract_not_located(prof),
                source=_o,
            ))
            continue
        _sev, _act = _cmp_num(_got, float(_w), SPACE_TOL_FAIL, SPACE_TOL_WARN, "pt")
        sev = _downgrade(_sev, _o, is_page_or_heading=True,
                         dim=_dim, spec_key=target.spec_key)
        findings.append(Finding(
            dimension=_dim, severity=sev, message=_label,
            expected="%g pt" % float(_w), actual=_act, source=_o,
        ))

    # —— 摘要**标题**自身的格式（v2.3.97）——
    # 老板 2026-10-23 打开 APA 产物一眼看出「摘要标题没居中」。此前压根没有这个维度，
    # 而 `Abstract` 段还被伪标题兜底采样认成二级章节标题、套上 Level 2 的左对齐 ——
    # 与 APA §2.9「标签居中 + 加粗」正好相反。
    # 规范驱动，三家口径差别很大：APA = 居中加粗；MLA = 不要求摘要；IEEE = 各模板不一。
    _ah_want: dict = {}
    for _k in ("abstract_heading_alignment", "abstract_heading_bold",
               "abstract_heading_italic"):
        _v = getattr(target, _k, None)
        if _v is None and _fm_base:
            _v = getattr(_fm_spec, _k, None)
        _ah_want[_k] = _v
    if any(v is not None for v in _ah_want.values()):
        _ah_origin = (target.resolution.get("abstract_heading_alignment", "school")
                      if getattr(target, "abstract_heading_alignment", None) is not None
                      else "spec")
        _ah_got_al = getattr(prof, "abstract_heading_alignment", None)
        if _ah_got_al is None:
            # 没有摘要标题 → 不判缺失（很多论文根本没有摘要页）
            findings.append(Finding(
                dimension="abstract_heading", severity="info",
                message="Abstract heading",
                expected=_describe_abstract_heading(
                    _ah_want["abstract_heading_alignment"],
                    _ah_want["abstract_heading_bold"],
                    _ah_want["abstract_heading_italic"]),
                actual=_abstract_not_located(prof, " (nothing to change)"),
                source=_ah_origin,
            ))
        else:
            _eh = _describe_abstract_heading(
                _ah_want["abstract_heading_alignment"],
                _ah_want["abstract_heading_bold"],
                _ah_want["abstract_heading_italic"])
            _gh, _ah_bad = _describe_abstract_heading_actual(prof, _ah_want)
            if _ah_bad:
                # 直接在发现时挂 FixAction（与 v2.3.95 的参考文献反向编号同一写法）：
                # 摘要标题的对齐/加粗属于**格式**、不碰一个字，规范基线或有显式要求就该修。
                _f = Finding(
                    dimension="abstract_heading", severity="warn",
                    message="Abstract heading",
                    expected=_eh, actual=_gh, source=_ah_origin,
                )
                _f.fix = FixAction(dimension="abstract_heading", value=True)
                findings.append(_f)
            else:
                findings.append(Finding(
                    dimension="abstract_heading", severity="pass",
                    message="Abstract heading",
                    expected=_eh, actual=_gh, source=_ah_origin,
                ))

    # —— Keywords 行自身的格式（v2.3.97）——
    # APA 7 §2.10：标签 "Keywords:" 斜体、整行像正文首行那样缩进 0.5″、词条本身不斜体。
    # 规范驱动：MLA 根本不要求摘要、IEEE 各模板自定 → 两者该两项均为 None，一个字都不说。
    _kw_want: dict = {}
    for _k in ("keywords_label_italic", "keywords_indent_in"):
        _v = getattr(target, _k, None)
        if _v is None and _fm_base:
            _v = getattr(_fm_spec, _k, None)
        _kw_want[_k] = _v
    if any(v is not None for v in _kw_want.values()):
        _kw_origin = (target.resolution.get("keywords_indent_in", "school")
                      if getattr(target, "keywords_indent_in", None) is not None
                      else "spec")
        _kw_has = (getattr(prof, "keywords_indent_in", None) is not None
                   or getattr(prof, "keywords_label_italic", None) is not None)
        _ek = _describe_keywords(_kw_want["keywords_label_italic"],
                                 _kw_want["keywords_indent_in"])
        if not _kw_has:
            # 没有 Keywords 行 → 不判缺失（我们不改写文字、更不能替客户编关键词）。
            # 但要分清「文档里就没有」与「摘要定位不到」——后者客户需要知道去补什么。
            _kw_why = ("Abstract found, but it has no 'Keywords:' line"
                       if getattr(prof, "has_abstract", False)
                       else _abstract_not_located(prof))
            findings.append(Finding(
                dimension="keywords", severity="info",
                message="Keywords line",
                expected=_ek, actual=_kw_why + " (nothing to change)",
                source=_kw_origin,
            ))
        else:
            _gk, _kw_bad = _describe_keywords_actual(prof, _kw_want)
            if _kw_bad:
                _f = Finding(
                    dimension="keywords", severity="warn",
                    message="Keywords line",
                    expected=_ek, actual=_gk, source=_kw_origin,
                )
                _f.fix = FixAction(dimension="keywords", value=True)
                findings.append(_f)
            else:
                findings.append(Finding(
                    dimension="keywords", severity="pass",
                    message="Keywords line",
                    expected=_ek, actual=_gk, source=_kw_origin,
                ))


    # 客户/学校/AI 显式声明「要 running head」或「页码含作者姓氏」时硬判；
    # 否则回落到已有 spec 驱动的 info 提示（APA 专业版 / MLA）。文本内容（短标题/姓氏）
    # 涉及「改内容」且无法可靠自动生成 → 只提示客户手动补；页码位置由 page_number 维度可修。
    _rh_req = target.running_head_required
    _rh_name = target.page_number_include_author_name
    _rh_origin = target.resolution.get("running_head_required") or \
        target.resolution.get("page_number_include_author_name") or "spec"
    _rh_origin = origin_of_source(_rh_origin) or "spec"
    if _rh_req is not None or _rh_name is not None:
        _rh_text = prof.header_text or ""
        _rh_has_text = bool(_rh_text.strip())
        _rh_has_page = bool(prof.page_number_present) or bool(re.search(r"\d", _rh_text))
        # ① 是否要求 running head
        if _rh_req is True:
            if not _rh_has_text and not _rh_has_page:
                findings.append(Finding(
                    dimension="running_head", severity="warn",
                    message="Running head",
                    expected="A running head in the header (short title or last name + page number)",
                    actual="No running-head text or page number detected in the header",
                    source=_rh_origin,
                ))
            else:
                findings.append(Finding(
                    dimension="running_head", severity="pass",
                    message="Running head",
                    expected="A running head in the header",
                    actual=_rh_text[:60] or "(page number in header)",
                    source=_rh_origin,
                ))
        elif _rh_req is False:
            # 明确不要求 running head：不判缺失，仅当有多余页眉文本时温和提示
            if _rh_has_text:
                findings.append(Finding(
                    dimension="running_head", severity="info",
                    message="Running head (not required)",
                    expected="No running head required",
                    actual=f"Header text detected: {_rh_text[:60]!r}",
                    source=_rh_origin,
                ))
        # ② 页码是否含作者姓氏（MLA 的「Smith 1」形态）
        if _rh_name is True:
            _name_part = re.sub(r"\d+", "", _rh_text).strip()
            _has_name = bool(re.search(r"[A-Za-z]", _name_part))
            if not _has_name:
                findings.append(Finding(
                    dimension="running_head_name", severity="warn",
                    message="Running head author name",
                    expected="Your last name followed by the page number (e.g., 'Smith 1')",
                    actual="No author name detected in the header",
                    source=_rh_origin,
                ))
            else:
                findings.append(Finding(
                    dimension="running_head_name", severity="pass",
                    message="Running head author name",
                    expected="Last name + page number",
                    actual=_rh_text[:60],
                    source=_rh_origin,
                ))

    # —— 图表 caption（v2.3.67）——
    # 字号/行距可查可修（识别 Caption 段落）；表题/图题位置只提示（图片 caption 位置
    # 无法可靠判定）。目标显式给出才判；None = 不规定。
    _cap_size = target.caption_font_size_pt
    _cap_spacing = target.caption_line_spacing
    _cap_tbl_pos = target.caption_position_table
    _cap_fig_pos = target.caption_position_figure
    # 目标没显式给、但规范带了 → 用规范值（来源 spec → 只提示语气）
    if _cap_size is None:
        _cap_size = apa_spec.caption_font_size_pt
    if _cap_spacing is None:
        _cap_spacing = apa_spec.caption_line_spacing
    if _cap_tbl_pos is None:
        _cap_tbl_pos = apa_spec.caption_position_table
    if _cap_fig_pos is None:
        _cap_fig_pos = apa_spec.caption_position_figure
    _has_captions = bool(getattr(prof, "caption_count", 0))
    # ① 字号（可修）
    if _cap_size is not None:
        _cap_size_origin = target.resolution.get("caption_font_size_pt", "spec")
        _cur_cap_size = prof.caption_font_size_pt
        if not _has_captions or _cur_cap_size is None:
            findings.append(Finding(
                dimension="caption_size", severity="info",
                message="Figure/table caption font size",
                expected=f"{_cap_size:g} pt",
                actual="No captions detected" if not _has_captions else "Could not read caption size",
                source=_cap_size_origin,
            ))
        elif abs(float(_cur_cap_size) - float(_cap_size)) > 0.01:
            findings.append(Finding(
                dimension="caption_size", severity="warn",
                message="Figure/table caption font size",
                expected=f"{_cap_size:g} pt",
                actual=f"{_cur_cap_size:g} pt",
                source=_cap_size_origin,
            ))
        else:
            findings.append(Finding(
                dimension="caption_size", severity="pass",
                message="Figure/table caption font size",
                expected=f"{_cap_size:g} pt",
                actual=f"{_cur_cap_size:g} pt",
                source=_cap_size_origin,
            ))
    # —— E-6 / D-22：MLA 论文标题的格式（v2.3.100）——
    # MLA 9：标题**居中**、**不加粗、不加斜体**、不额外加空行（与正文同字号）。
    # ⚠️ **只对 MLA 档判**（APA 标题是要加粗的，IEEE 有自己的口径）。
    # ⚠️ 置信度分两级（reader 的 sample_mla_title）：
    #   · certain=True（标题段居中）→ **可自动修**（去粗 / 去斜，改格式不改文字）；
    #   · certain=False（客户没居中，只能靠"信息块之后第一段"猜）→ **只提示不自动改**，
    #     因为那一段可能其实是正文首段，改错就是改坏客户的文档。
    if str(getattr(target, "spec_key", "") or "").upper().startswith("MLA"):
        _mi = getattr(prof, "mla_title_index", None)
        if _mi is not None:
            _mcertain = bool(getattr(prof, "mla_title_certain", False))
            _mal = getattr(prof, "mla_title_alignment", None)
            _mb = getattr(prof, "mla_title_bold", None)
            _mit = getattr(prof, "mla_title_italic", None)
            if _mb is True:
                _f = Finding(
                    dimension="mla_title_bold",
                    severity="warn" if _mcertain else "info",
                    message="MLA paper title: bold",
                    expected="The MLA title is not bold",
                    actual="The title paragraph is bold",
                    source="spec",
                )
                if _mcertain:
                    _f.fix = FixAction(dimension="mla_title_bold", value=False)
                findings.append(_f)
            if _mit is True:
                _f = Finding(
                    dimension="mla_title_italic",
                    severity="warn" if _mcertain else "info",
                    message="MLA paper title: italic",
                    expected="The MLA title is not italic",
                    actual="The title paragraph is italic",
                    source="spec",
                )
                if _mcertain:
                    _f.fix = FixAction(dimension="mla_title_italic", value=False)
                findings.append(_f)
            if not _mcertain:
                findings.append(Finding(
                    dimension="mla_title_alignment", severity="info",
                    message="MLA paper title: placement and alignment",
                    expected=("The title is centred and is neither bold nor italic; it sits "
                              "after the upper-left name/instructor/course/date block"),
                    actual=("Paragraph %s looks like the title (alignment: %s, bold: %s, "
                            "italic: %s) - centre it yourself if that is the title"
                            % (_mi, _mal, _mb, _mit)),
                    source="doc",
                ))

    # —— E-2 / D-18：APA caption 的**字形**（v2.3.100）——
    # APA 7：图表**号**（"Table 1"）**粗体**，**题**（描述）**斜体**。
    # ⚠️ **只对 APA 档判**：IEEE 的 caption 用 small caps 且不斜体，MLA 不强制 ——
    #   同一套规则套到 IEEE 上就是把官方模板判红（正是本轮一直在消灭的那类错）。
    # 实测 6 份有 caption 的官方 APA 样例**全部合规**（号段粗、题段斜）→ 不会误判。
    if str(getattr(target, "spec_key", "") or "").upper().startswith("APA"):
        _cap_cnt = int(getattr(prof, "caption_typography_count", 0) or 0)
        if _cap_cnt:
            _lb = getattr(prof, "caption_label_bold", None)
            if _lb is False:
                _f = Finding(
                    dimension="caption_label_bold", severity="warn",
                    message="Caption label bold",
                    expected="The figure/table number (e.g. 'Table 1') is bold",
                    actual="The caption number is not bold",
                    source="spec",
                )
                _f.fix = FixAction(dimension="caption_label_bold", value=True)
                findings.append(_f)
            _ti = getattr(prof, "caption_title_italic", None)
            if _ti is False:
                _f = Finding(
                    dimension="caption_title_italic", severity="warn",
                    message="Caption title italic",
                    expected="The figure/table title is italic",
                    actual="The caption title is not italic",
                    source="spec",
                )
                _f.fix = FixAction(dimension="caption_title_italic", value=True)
                findings.append(_f)

    # —— E 批：MLA + APA 剩余的**只提示**观测（v2.3.100）——
    # 🔴 全部 info、不挂 FixAction：涉及"改文字/改结构"的一律只提示。

    # E-1（D-17）：正文页顶部重复了论文标题。
    # ⚠️ 只做**中性提示**，绝不判它错 —— APA 稿的正文页顶部本来就要重复论文标题，
    # 官方 APA 专业版样例自己就是这样（实测 body_title_repeat=True）。
    # 若按"重复=问题"去判 warn，等于把官方合规样例判红（正是本轮一直在消灭的那类错）。
    if getattr(prof, "body_title_repeat", False):
        findings.append(Finding(
            dimension="body_title_repeat", severity="info",
            message="Paper title repeated on the first page of the text",
            expected="Many styles repeat the title at the top of the first text page",
            actual="The paper title also appears in the body - keep it only if your "
                   "style asks for it",
            source="doc",
        ))

    # E-3 / E-4（D-19 / D-20）：编号标题。APA 与 MLA 都不给章节标题编号
    #（IEEE 才用罗马数字编号）→ 提示客户改，但**不代改**（改文字 = 越权）。
    _nbh = list(getattr(prof, "numbered_body_heads", None) or [])
    if _nbh:
        findings.append(Finding(
            dimension="numbered_body_heads", severity="info",
            message="Numbered section headings",
            expected="APA and MLA do not number section headings (only IEEE does)",
            actual="%d numbered heading(s): %s" % (
                len(_nbh), ", ".join(repr(x.get("text")) for x in _nbh[:4])),
            source="spec",
        ))

    # E-5（D-21）：MLA 首页左上四行（姓名 / 教授 / 课程 / 日期）。
    # ⚠️ **只增强既有 first_page_heading 的文案，不新增 finding** —— 既有维度是
    # 「检测不到四行 → info 提示」，我若在"检测到"时再报一条，会与它语义相反
    # （`test_mla_first_page_block_present_no_finding` 守护的正是"有信息块就不报"）。
    # 读到的行数与内容放进既有那条的 actual 里，供客户核对缺了哪几项。

    # E-7 / E-8（D-40 / D-41）：APA 专业版 Author Note 的位置与上标编号。
    _an_idx = getattr(prof, "author_note_index", None)
    if _an_idx is not None:
        findings.append(Finding(
            dimension="author_note_position", severity="info",
            message="Author Note placement",
            expected="The Author Note sits on the title page (APA professional papers)",
            actual="An Author Note was found at paragraph %s" % _an_idx,
            source="spec",
        ))
        _an_sup = getattr(prof, "author_note_has_superscript", None)
        if _an_sup is False:
            findings.append(Finding(
                dimension="author_note_superscript", severity="info",
                message="Author Note numbering",
                expected="Author Note entries are marked with superscript numbers",
                actual="No superscript marker was found in the Author Note",
                source="spec",
            ))

    # —— D-2~D-11：IEEE 系列的**只提示**观测（v2.3.100）——
    # 🔴 D 批门禁明确要求：「新增的 info 类维度不得产生任何 warn/fail」。
    # 这十项要么涉及**改文字**（引用顺序、拼写、编号形态），要么只是"看有没有"，
    # 一律 **info + 不挂 FixAction**：提示客户「IEEE 的口径是什么」，改不改他自己决定。
    _spec_key_is_ieee = str(getattr(target, "spec_key", "") or "").upper().startswith("IEEE")
    if getattr(spec, "expect_index_terms", False) or _spec_key_is_ieee:
        # D-2：表/图标签形态（IEEE 表题用罗马数字大写「TABLE I」，图题用「Fig. N.」）
        _caps = list(getattr(prof, "ieee_caption_labels", None) or [])
        if _caps:
            _bad_cap = []
            for _c in _caps:
                if _c.get("kind") == "table":
                    if _c.get("numeral") != "roman" or not _c.get("uppercase"):
                        _bad_cap.append("%r (IEEE table captions use Roman numerals, "
                                        "e.g. 'TABLE I')" % _c.get("label"))
                else:
                    if _c.get("numeral") != "arabic":
                        _bad_cap.append("%r (IEEE figure captions use 'Fig. N.')"
                                        % _c.get("label"))
            if _bad_cap:
                findings.append(Finding(
                    dimension="ieee_caption_label", severity="info",
                    message="Table/figure caption labels",
                    expected="Tables: 'TABLE I' (Roman numerals); figures: 'Fig. 1.'",
                    actual="; ".join(_bad_cap[:4]),
                    source="spec",
                ))
            # D-3：IEEE 表题末尾**不加句点**（ESM：表题描述末尾无句点）
            _per = [c for c in _caps
                    if c.get("kind") == "table" and c.get("ends_with_period")]
            if _per:
                findings.append(Finding(
                    dimension="ieee_caption_period", severity="info",
                    message="Table caption punctuation",
                    expected="No full stop at the end of an IEEE table caption",
                    actual="%d table caption(s) end with a period: %s" % (
                        len(_per), ", ".join(repr(c.get("text")) for c in _per[:3])),
                    source="spec",
                ))

        # D-4：图表在正文里首次引用必须按数字顺序
        _pos = getattr(prof, "ieee_caption_position", None) or {}
        if _pos.get("figure_ref_order_ok") is False:
            findings.append(Finding(
                dimension="ieee_figure_ref_order", severity="info",
                message="Order of first references to figures",
                expected="Figures are first mentioned in ascending number order",
                actual="First mentions are not in ascending order",
                source="spec",
            ))

        # D-5：首脚注（IEEE 惯例：首脚注不编号、用上标符号）
        _ff = getattr(prof, "ieee_first_footnote", None) or {}
        if _ff:
            findings.append(Finding(
                dimension="ieee_first_footnote", severity="info",
                message="First footnote",
                expected="The first footnote carries an unnumbered marker",
                actual=("The first footnote is numbered"
                       if _ff.get("numbered")
                       else "The first footnote marker is not numbered"),
                source="spec",
            ))

        # D-6 / D-7：致谢拼写与利益冲突声明的位置
        _ack = getattr(prof, "ieee_acknowledgment", None) or {}
        if _ack:
            findings.append(Finding(
                dimension="ieee_acknowledgment", severity="info",
                message="Acknowledgment heading",
                expected="IEEE uses 'Acknowledgment' (US spelling)",
                actual="Found %r" % _ack.get("spelling"),
                source="spec",
            ))
        _coi = getattr(prof, "ieee_coi", None) or {}
        if _coi:
            findings.append(Finding(
                dimension="ieee_coi", severity="info",
                message="Conflict of interest statement",
                expected="A conflict-of-interest statement, placed near the end matter",
                actual="Found a conflict-of-interest statement (paragraph %s)"
                       % _coi.get("para"),
                source="spec",
            ))

        # D-8：作者传记里姓名加粗
        _bio = getattr(prof, "ieee_biography", None) or {}
        if _bio and _bio.get("name_bold") is False:
            findings.append(Finding(
                dimension="ieee_biography", severity="info",
                message="Biography entry",
                expected="The author's name in a biography entry is bold",
                actual="The name after the biography heading is not bold",
                source="spec",
            ))

        # D-9：Index Terms 必须是摘要的**最后一段**
        _it = getattr(prof, "ieee_index_terms", None) or {}
        if _it:
            if not _it.get("is_last_in_abstract"):
                findings.append(Finding(
                    dimension="ieee_index_terms", severity="info",
                    message="Index Terms placement",
                    expected="Index Terms form the last paragraph of the abstract",
                    actual="The Index Terms paragraph is not the last one in the abstract",
                    source="spec",
                ))
        else:
            findings.append(Finding(
                dimension="ieee_index_terms", severity="info",
                message="Index Terms placement",
                expected="Index Terms as the last paragraph of the abstract",
                actual="No Index Terms paragraph was detected",
                source="spec",
            ))

        # D-10：摘要里不得出现编号引用或编号公式
        _aref = list(getattr(prof, "ieee_abstract_numbered_refs", None) or [])
        if _aref:
            findings.append(Finding(
                dimension="ieee_abstract_numbered_refs", severity="info",
                message="Numbered citations or equations in the abstract",
                expected="The abstract carries no numbered citations or numbered equations",
                actual="%d found: %s" % (
                    len(_aref),
                    ", ".join(repr(a.get("text")) for a in _aref[:4])),
                source="spec",
            ))

        # D-11：IEEE 惯例章节标题用**复数**（Introduction / References）
        _sing = list(getattr(prof, "ieee_numbered_singular_heads", None) or [])
        if _sing:
            findings.append(Finding(
                dimension="ieee_singular_heads", severity="info",
                message="Plural section headings",
                expected="IEEE uses plural headings (e.g. 'Introduction', 'References')",
                actual="Numbered singular heading(s): %s" % (
                    ", ".join(repr(x.get("text")) for x in _sing[:4])),
                source="spec",
            ))

    # —— F-9：三类「只提示、不改字」的引用 / 标点观测（裁定 C-12 C-13 C-14）——
    # 🔴 三条全部 info、**绝不挂 FixAction**：改标点/改引用写法就是改客户的文字
    #（P-2 文字锁死）。只告诉客户「IEEE 现在的口径是什么」，改不改由客户自己决定。

    # C-12：文内连续引用里的 en dash。IEEE Reference Guide V3.28.2025 原文：
    # 「Reference ranges in text **will not include an en dash** … '[1]–[4]' will now be
    # '[1], [2], [3], [4]'」（Purdue OWL 的旧写法已作废）。
    _en_dash = list(getattr(prof, "citation_range_en_dash", None) or [])
    if _en_dash:
        findings.append(Finding(
            dimension="citation_range_style", severity="info",
            message="Citation ranges in text",
            expected=("IEEE no longer uses an en dash for a range of references: write "
                      "[1], [2], [3], [4] instead of [1]-[4]"),
            actual="%d range(s) with a dash: %s" % (
                len(_en_dash),
                "; ".join("%r (paragraph %s)" % (h.get("text"), h.get("para"))
                          for h in _en_dash[:4])),
            source="spec",
        ))

    # C-13：Jr. / Sr. / III 前面的逗号。REF 原文：「**Do not use commas** around Jr.,
    # Sr., and III：Michael Smith Jr.」
    _sfx = list(getattr(prof, "suffix_name_comma", None) or [])
    if _sfx:
        findings.append(Finding(
            dimension="name_suffix_style", severity="info",
            message="Name suffixes (Jr., Sr., III)",
            expected="No comma before Jr., Sr., or III (e.g. 'Michael Smith Jr.')",
            actual="%d occurrence(s) with a comma: %s" % (
                len(_sfx),
                "; ".join("%r (paragraph %s)" % (h.get("text"), h.get("para"))
                          for h in _sfx[:4])),
            source="spec",
        ))

    # C-14：公式编号**全文一致性**（ESM：形式随作者，但必须一致；数学指南另说
    # 连字符/句点通常去掉 —— 两者不冲突的交集就是"一致"）。我们**只判混用**，
    # 绝不判某种形式为错（ESM 明说 (1a)/(1.1)/(1-1) 都可）。
    _eq_styles = list(getattr(prof, "equation_number_styles", None) or [])
    if len(_eq_styles) > 1:
        findings.append(Finding(
            dimension="equation_number_style", severity="info",
            message="Equation numbering style",
            expected="One consistent style for all equation numbers",
            actual="Mixed styles found: %s (IEEE asks for consistency, not one form)"
                   % ", ".join(_eq_styles),
            source="spec",
        ))

    # —— 官方模板**指导文字残留**（v2.3.100，D-1 / 裁定 C-11）——
    # IEEE 官方会议模板下载页原文：「Please ensure that all guidance text is removed
    # from your conference paper prior to submission… **Failure to remove template text
    # from your paper may result in your paper not being published.**」
    # 客户把模板直接交上去 = 论文不发表，而我们此前报"全部合规"——这是本轮最值钱的发现。
    # 🔴 只报位置、**绝不挂 FixAction**：挂上"一键修正"就会去删客户的文字（越权 +
    #    违反 P-2 文字锁死）。删除必须客户自己做。
    _bp_hits = list(getattr(prof, "boilerplate_hits", None) or [])
    if _bp_hits:
        _bp_where = {"body": "body paragraph", "header": "header",
                     "footer": "footer"}
        _bp_list = "; ".join(
            "%s %s: %r" % (_bp_where.get(h.get("where"), h.get("where")),
                           h.get("index"), h.get("text"))
            for h in _bp_hits[:6])
        if len(_bp_hits) > 6:
            _bp_list += "; …and %d more" % (len(_bp_hits) - 6)
        findings.append(Finding(
            dimension="template_boilerplate", severity="fail",
            message="Template guidance text still in the document",
            expected=("No template placeholder or guidance text (IEEE: 'Failure to "
                      "remove template text … may result in your paper not being "
                      "published')"),
            actual="%d place(s): %s" % (len(_bp_hits), _bp_list),
            source="spec",
        ))
    else:
        findings.append(Finding(
            dimension="template_boilerplate", severity="pass",
            message="Template guidance text still in the document",
            expected="No template placeholder or guidance text",
            actual="None found",
            source="spec",
        ))

    # —— byline（作者/单位/城市/邮箱）字号 + 结构（v2.3.100，F-5 / 裁定 C-6）——
    # 只对**显式给出 byline 字号**的风格判（IEEE = 11pt，模板 `Author` 样式实测）。
    # 结构（几行、缺不缺邮箱/单位）**只提示**：补行就是替客户写内容（P-2）。
    _bl_size = getattr(target, "byline_size_pt", None)
    if _bl_size is None and _fm_base:
        _bl_size = getattr(_fm_spec, "byline_size_pt", None)
    if _bl_size is not None:
        _bl_origin = (target.resolution.get("byline_size_pt", "school")
                      if getattr(target, "byline_size_pt", None) is not None else "spec")
        _cur_bl = prof.byline_size_pt
        _bl_lines = int(getattr(prof, "byline_line_count", 0) or 0)
        if _cur_bl is None:
            findings.append(Finding(
                dimension="byline_size", severity="info",
                message="Author block font size",
                expected=f"{_bl_size:g} pt",
                actual=("No author block detected (no Author/Affiliation style and no "
                        "title-page author lines)"),
                source=_bl_origin,
            ))
        elif abs(float(_cur_bl) - float(_bl_size)) > 0.01:
            findings.append(Finding(
                dimension="byline_size", severity="warn",
                message="Author block font size",
                expected=f"{_bl_size:g} pt",
                actual=f"{_cur_bl:g} pt",
                source=_bl_origin,
            ))
        else:
            findings.append(Finding(
                dimension="byline_size", severity="pass",
                message="Author block font size",
                expected=f"{_bl_size:g} pt",
                actual=f"{_cur_bl:g} pt",
                source=_bl_origin,
            ))
        # 结构只提示：IEEE 惯例四行（姓名 / 单位 / 城市+国家 / 邮箱或 ORCID）。
        # 缺项**不代写**，只告诉客户去哪补。
        if _bl_lines:
            findings.append(Finding(
                dimension="byline_structure", severity="info",
                message="Author block contents",
                expected=("Name / affiliation / city and country / e-mail or ORCID "
                          "(IEEE convention)"),
                actual=("%d author line(s) found; check that every author has all four "
                        "parts" % _bl_lines),
                source=_bl_origin,
            ))

    # ② 行距（可修）
    if _cap_spacing is not None:
        _cap_spacing_origin = target.resolution.get("caption_line_spacing", "spec")
        _cur_cap_sp = prof.caption_line_spacing
        if not _has_captions or _cur_cap_sp is None:
            findings.append(Finding(
                dimension="caption_spacing", severity="info",
                message="Figure/table caption line spacing",
                expected=f"{_cap_spacing:g}x",
                actual="No captions detected" if not _has_captions else "Could not read caption spacing",
                source=_cap_spacing_origin,
            ))
        elif abs(float(_cur_cap_sp) - float(_cap_spacing)) > 0.01:
            findings.append(Finding(
                dimension="caption_spacing", severity="warn",
                message="Figure/table caption line spacing",
                expected=f"{_cap_spacing:g}x",
                actual=f"{_cur_cap_sp:g}x",
                source=_cap_spacing_origin,
            ))
        else:
            findings.append(Finding(
                dimension="caption_spacing", severity="pass",
                message="Figure/table caption line spacing",
                expected=f"{_cap_spacing:g}x",
                actual=f"{_cur_cap_sp:g}x",
                source=_cap_spacing_origin,
            ))
    # ③ 段前 / 段后（可修，v2.3.76）
    # 表题在表格上方、图题在图片下方 ——「题与其相邻内容之间留多少」是排版硬要求。
    for _k, _dim, _label, _attr in (
            ("caption_space_before_pt", "caption_space_before",
             "Figure/table caption space before", "caption_space_before_pt"),
            ("caption_space_after_pt", "caption_space_after",
             "Figure/table caption space after", "caption_space_after_pt")):
        _w = getattr(target, _k, None)
        if _w is None:
            continue
        _o = target.resolution.get(_k, "school")
        _got = getattr(prof, _attr, None)
        if not _has_captions or _got is None:
            findings.append(Finding(
                dimension=_dim, severity="info", message=_label,
                expected="%g pt" % float(_w),
                actual="No captions detected" if not _has_captions
                else "Could not read caption spacing",
                source=_o,
            ))
            continue
        _sev, _act = _cmp_num(_got, float(_w), SPACE_TOL_FAIL, SPACE_TOL_WARN, "pt")
        sev = _downgrade(_sev, _o, is_page_or_heading=True,
                         dim=_dim, spec_key=target.spec_key)
        findings.append(Finding(
            dimension=_dim, severity=sev, message=_label,
            expected="%g pt" % float(_w), actual=_act, source=_o,
        ))
    # ④ 表题位置（只提示）
    if _cap_tbl_pos is not None:
        _cap_tbl_origin = target.resolution.get("caption_position_table", "spec")
        _tbl_positions = getattr(prof, "caption_table_positions", None) or []
        if not _tbl_positions:
            findings.append(Finding(
                dimension="caption_position_table", severity="info",
                message="Table caption position",
                expected="table caption " + ("above the table" if _cap_tbl_pos == "top" else "below the table"),
                actual="No table caption position could be determined",
                source=_cap_tbl_origin,
            ))
        else:
            _mode_pos = Counter(_tbl_positions).most_common(1)[0][0]
            _mode_label = "above" if _mode_pos == "top" else "below"
            _want_label = "above" if _cap_tbl_pos == "top" else "below"
            if _mode_pos != _cap_tbl_pos:
                findings.append(Finding(
                    dimension="caption_position_table", severity="info",
                    message="Table caption position",
                    expected="table caption " + _want_label + " the table",
                    actual="table caption " + _mode_label + " the table",
                    source=_cap_tbl_origin,
                ))
            else:
                findings.append(Finding(
                    dimension="caption_position_table", severity="pass",
                    message="Table caption position",
                    expected="table caption " + _want_label + " the table",
                    actual="Matches",
                    source=_cap_tbl_origin,
                ))
    # ⑤ 图题位置（只提示）：图片 caption 位置无法可靠判定，仅当有 caption 时给说明性提示
    if _cap_fig_pos is not None:
        _cap_fig_origin = target.resolution.get("caption_position_figure", "spec")
        _fig_want = "above the figure" if _cap_fig_pos == "top" else "below the figure"
        findings.append(Finding(
            dimension="caption_position_figure", severity="info",
            message="Figure caption position",
            expected="figure caption " + _fig_want,
            actual="Figure caption position is not auto-checked (figures are embedded in paragraphs); "
                   "verify it manually",
            source=_cap_fig_origin,
        ))

    # —— 表格线（学术三线表，v2.3.85）——
    # 老板 2026-10-21 清单第③条：「表格去四周边框做学术三线表」。审计实测：这条
    # **完全没做** —— 4 张表在四种输出里边框一个字节都没变，引擎源码里连 tblBorders
    # 都没有。而真实样本的 4 张表**根本没有边框**（0 处 tblBorders / 0 处 tcBorders），
    # 所以「去掉四边」不够，还得把顶线 / 表头下线 / 底线**画上去**。
    # 与栏数同为**版式结构维度**：选中 APA/MLA/IEEE 就是为其表格样式做了显式声明
    # （三家都只允许横线、禁止竖线 —— APA 7th §7.10 / IEEEtran open table），
    # 故本维度**不**做 spec→info 软参考降级；学校模板/问卷显式给出时本就硬判。
    # 只判「线」，绝不碰表格里的文字（文字 100% 锁死，见 policy.TEXT_IS_LOCKED）。
    if getattr(target, "table_lines", None) == "three_line":
        _lines_origin = target.resolution.get("table_lines", "spec")
        _states = list(getattr(prof, "table_line_states", None) or [])
        _lines_msg = "Table rules (three-line academic table)"
        _lines_exp = ("three-line tables: top rule + rule under the header + bottom rule, "
                      "no vertical rules, no inner horizontal rules")
        if not _states:
            findings.append(Finding(
                dimension="table_lines", severity="info", message=_lines_msg,
                expected=_lines_exp,
                actual="No tables were found in the document",
                source=_lines_origin,
            ))
        else:
            _bad = [(i, s) for i, s in enumerate(_states) if not s.get("three_line")]
            if not _bad:
                findings.append(Finding(
                    dimension="table_lines", severity="pass", message=_lines_msg,
                    expected=_lines_exp,
                    actual="%d table(s), all three-line" % len(_states),
                    source=_lines_origin,
                ))
            else:
                _detail = "; ".join(
                    "Table %d: %s" % (s.get("index", i + 1), describe_table_lines(s))
                    for i, s in _bad)
                findings.append(Finding(
                    dimension="table_lines", severity="warn", message=_lines_msg,
                    expected=_lines_exp,
                    actual=_detail,
                    source=_lines_origin,
                ))
    # —— 表格与前后文字的间距（v2.3.98）——
    # 老板 2026-10-23 截图：表格下面紧贴「Chapter 23」一级标题。取证发现原文该段
    # 段前 14pt，被我们「正文段前段后应为 0」的规则清成了 0 —— 因为此前**没有任何
    # 维度**管「表格与文字之间的空隙」。
    # APA 7（Publication Manual 7th, Tables and Figures）原文：「If embedded, place the
    # table or figure at either the top or the bottom of the page with an extra
    # double-spaced line between the table or the figure and any text.」
    # → 表格是「段前段后不加额外空」的**例外**，必须与任何文字之间额外空一条双倍行距。
    _tg_want = getattr(target, "table_gap_pt", None)
    if _tg_want is None and _fm_base:
        _tg_want = getattr(_fm_spec, "table_gap_pt", None)
    if _tg_want is not None:
        _tg_origin = (target.resolution.get("table_gap_pt", "school")
                      if getattr(target, "table_gap_pt", None) is not None else "spec")
        _tg_msg = "Spacing between tables and the text"
        # v2.3.100（F-2 / 裁定 C-2）：三档阈值。
        #   · 硬下限 ``_tg_min``：低于它才 warn（可修）—— 默认取 min_pt，没配则退回推荐值；
        #   · 推荐线 ``_tg_good`` = 推荐值 × ratio（默认 0.8）—— 达到它即算合格。
        # 官方样例实测 12pt 落在 [12, 19.2) → pass，正合裁定「12pt 放行」。
        _tg_min_spec = getattr(target, "table_gap_min_pt", None)
        if _tg_min_spec is None and _fm_base:
            _tg_min_spec = getattr(_fm_spec, "table_gap_min_pt", None)
        _tg_ratio_spec = getattr(target, "table_gap_ratio", None)
        if _tg_ratio_spec is None and _fm_base:
            _tg_ratio_spec = getattr(_fm_spec, "table_gap_ratio", None)
        _tg_min = float(_tg_min_spec) if _tg_min_spec is not None else float(_tg_want)
        _tg_ratio = float(_tg_ratio_spec) if _tg_ratio_spec is not None else 1.0
        _tg_good = float(_tg_want) * _tg_ratio
        _tg_exp = ("at least %g pt (one double-spaced line) above and below each table"
                   % float(_tg_want))
        if _tg_min < float(_tg_want):
            _tg_exp += ("; %g pt is accepted (below the recommended %g pt)"
                        % (_tg_min, float(_tg_want)))
        _tg_samples = list(getattr(prof, "table_gap_samples", None) or [])
        if not _tg_samples:
            findings.append(Finding(
                dimension="table_gap", severity="info", message=_tg_msg,
                expected=_tg_exp, actual="No tables were found in the document",
                source=_tg_origin,
            ))
        else:
            _tg_short = []
            _tg_judged = 0
            _tg_soft = []          # 达硬下限但没到推荐值 → 通过，附一句推荐
            for _s in _tg_samples:
                _bits = []
                _soft_bits = []
                for _side, _key, _blank, _same in (
                        ("above", "before_pt", "before_blank", "before_same_page"),
                        ("below", "after_pt", "after_blank", "after_same_page")):
                    # v2.3.100（F-2 / 裁定 C-2）：**跨页侧不判**。表格在页顶/页底时，
                    # 那一侧量出来的空隙其实跨了两页，判它毫无意义（且 docx 无绝对坐标，
                    # 只能用邻接段 pageBreakBefore 近似 → 不确定时宁可不判）。
                    if _s.get(_same) is False:
                        continue
                    if _s.get(_blank):
                        continue            # 已有真空行 → 这一侧算够
                    _g = _s.get(_key)
                    if _g is None:
                        continue            # 这一侧不判（邻接表格/文档边界）
                    _tg_judged += 1
                    _gv = float(_g)
                    # 三档：< 硬下限 → warn（可修）；[下限, 推荐×比例) → pass+推荐提示；
                    #        ≥ 推荐×比例 → pass。
                    if _gv + 1e-6 < _tg_min:
                        _bits.append("%s %g pt" % (_side, _gv))
                    elif _gv + 1e-6 < _tg_good:
                        _soft_bits.append("%s %g pt" % (_side, _gv))
                if _bits:
                    _tg_short.append("Table %d: %s" % (_s.get("index", "?"),
                                                       ", ".join(_bits)))
                elif _soft_bits:
                    _tg_soft.append("Table %d: %s" % (_s.get("index", "?"),
                                                      ", ".join(_soft_bits)))
            if not _tg_short and _tg_judged:
                _act = "%d table(s), enough space around each" % len(_tg_samples)
                if _tg_soft:
                    _act += (" (%s — below the recommended %g pt, but acceptable)"
                             % ("; ".join(_tg_soft), float(_tg_want)))
                findings.append(Finding(
                    dimension="table_gap", severity="pass", message=_tg_msg,
                    expected=_tg_exp,
                    actual=_act,
                    source=_tg_origin,
                ))
            elif not _tg_short:
                # 每一侧都已是真空行（或紧邻另一张表）→ 这份文档无从判定，
                # 不能说「间距够」（那是假绿），只说「表旁边本来就空着，没动」。
                findings.append(Finding(
                    dimension="table_gap", severity="info", message=_tg_msg,
                    expected=_tg_exp,
                    actual=("The tables are already separated from the text by a blank "
                            "line, so no spacing was added"),
                    source=_tg_origin,
                ))
            else:
                _f = Finding(
                    dimension="table_gap", severity="warn", message=_tg_msg,
                    expected=_tg_exp,
                    actual="; ".join(_tg_short),
                    source=_tg_origin,
                )
                # 只动段前/段后间距，不碰一个字 → 可修。
                _f.fix = FixAction(dimension="table_gap", value=float(_tg_want))
                findings.append(_f)

    # —— 公式（Word 原生 OMML，v2.3.112）——
    # 学校模板里「公式居中、编号加括号靠右边界、与正文同字号、段前段后 0」是高频硬要求。
    # 引擎此前**完全读不到** Word 公式：数学 run 是 ``m:r``，而 python-docx 的
    # ``para.runs`` 只返回 ``w:r`` —— 公式段在引擎眼里是「除了编号一个字都没有的空段」，
    # 任何维度都判不到它。report 里那句「Complex formulas / equations are NOT processed」
    # 说的就是这个空白。
    #
    # 🔴 只判**版式**（对齐 / 编号位置 / 字号 / 段距 / 行距），公式内容一个字都不碰
    #    （policy.TEXT_IS_LOCKED）。来源没声明公式规则的文档，本块整段不判。
    # 与 table_lines 同为**元素级版式维度**，故不做 spec→info 软参考降级。
    _eq_samples = list(getattr(prof, "equation_samples", None) or [])
    # 只有**展示公式**（独立成行）吃对齐 / 编号 / 字号 / 段距；行内公式随正文跑，
    # 它的字号归「正文字号」维度，在这里再判一次就是两套 owner 打架。
    _eq_disp = [s for s in _eq_samples if s.get("display")]

    def _eq_ids(items) -> str:
        return ", ".join("equation %d" % s.get("index", 0) for s in items)

    # 有公式规则、却没有可判的展示公式 → **只出一条**说明。
    # 若放任每个维度各报一句「没找到公式」，模板写了几条公式规则就刷几行同样的话，
    # 报告立刻变成噪音墙（6 条规则 = 6 行废话）。
    _eq_rules_declared = any(
        getattr(target, _k, None) is not None
        for _k in ("equation_alignment", "equation_number_position", "equation_size_pt",
                   "equation_font_family", "equation_line_spacing",
                   "equation_space_before_pt", "equation_space_after_pt"))
    _eq_judge = bool(_eq_disp)
    if _eq_rules_declared and not _eq_judge:
        findings.append(Finding(
            dimension="equation_format", severity="info", message="Equations",
            expected="Display equations follow the equation rules in your template",
            actual=("No equations were found in the document, so the equation rules your "
                    "template states do not apply to anything here."
                    if not _eq_samples else
                    "The document has inline equations only (maths inside a sentence). "
                    "There is no displayed equation to align, number or space, so those "
                    "rules are not judged."),
            source=target.resolution.get("equation_alignment", "school"),
        ))

    _eq_align_want = getattr(target, "equation_alignment", None)
    if _eq_align_want and _eq_judge:
        _o = target.resolution.get("equation_alignment", "school")
        _exp = ("Display equations centred in the text column"
                if _eq_align_want == "center" else "Display equations left-aligned")
        _bad = [s for s in _eq_disp
                if (s.get("alignment") or "center") != _eq_align_want]
        if not _bad:
            findings.append(Finding(
                dimension="equation_alignment", severity="pass",
                message="Equation alignment", expected=_exp,
                actual="%d display equation(s), all %s" % (len(_eq_disp), _eq_align_want),
                source=_o,
            ))
        else:
            _f = Finding(
                dimension="equation_alignment", severity="warn",
                message="Equation alignment", expected=_exp,
                actual="; ".join("%s: %s" % (_eq_ids([s]),
                                             s.get("alignment") or "alignment not set")
                                 for s in _bad),
                source=_o,
            )
            # 只改 m:oMathParaPr/m:jc（段落级展示对齐），不动任何数学内容 → 可修。
            _f.fix = FixAction(dimension="equation_alignment", value=_eq_align_want)
            findings.append(_f)

    _eq_num_want = getattr(target, "equation_number_position", None)
    if _eq_num_want and _eq_judge:
        _o = target.resolution.get("equation_number_position", "school")
        # 「公式 + 编号」放进一张无框表（公式居中单元格、编号右单元格）是**合法版式**，
        # 此时编号根本不在公式段里 —— 拿「同段有没有编号」去判会把它误报成「没编号」。
        # 表格里的公式只判对齐/字号/段距，编号位置不判（如实说明，不猜）。
        _eq_tbl = [s for s in _eq_disp if s.get("in_table")]
        _eq_flat = [s for s in _eq_disp if not s.get("in_table")]
        if _eq_tbl:
            findings.append(Finding(
                dimension="equation_number_position", severity="info",
                message="Equation numbering",
                actual=("%s sit inside a table. A formula-plus-number layout built as a "
                        "table is legitimate, so the number position was not judged."
                        % _eq_ids(_eq_tbl)),
                expected="The equation number sits in parentheses, flush with the right margin",
                source=_o,
            ))
        _with_num = [s for s in _eq_flat if s.get("number_text")]
        _without = [s for s in _eq_flat if not s.get("number_text")]
        if not _eq_flat:
            pass
        elif _eq_num_want == "none":
            # 「不许编号」→ 删编号就是删客户的文字，只提示（P-2 铁律）。
            if _with_num:
                findings.append(Finding(
                    dimension="equation_number_position", severity="info",
                    message="Equation numbering",
                    expected="Display equations are not numbered",
                    actual=("%s carry a number (%s). Removing it would delete your own "
                            "text, so it is reported only."
                            % (_eq_ids(_with_num),
                               ", ".join(str(s.get("number_text")) for s in _with_num))),
                    source=_o,
                ))
            else:
                findings.append(Finding(
                    dimension="equation_number_position", severity="pass",
                    message="Equation numbering",
                    expected="Display equations are not numbered",
                    actual="%d display equation(s), none numbered" % len(_eq_flat),
                    source=_o,
                ))
        else:
            _bad = [s for s in _with_num if not s.get("number_at_right")]
            _fixable = [s for s in _bad if s.get("number_by_tab")]
            _stuck = [s for s in _bad if not s.get("number_by_tab")]
            if _fixable:
                _f = Finding(
                    dimension="equation_number_position", severity="warn",
                    message="Equation numbering",
                    expected="The equation number sits in parentheses, flush with the "
                             "right margin",
                    actual="%s: the number is not pushed to the right margin" % _eq_ids(_fixable),
                    source=_o,
                )
                # 已有 Tab 分隔 → 只需把右对齐制表位补到右边界，不加任何字符 → 可修。
                _f.fix = FixAction(dimension="equation_number_position", value="right")
                findings.append(_f)
            if _stuck:
                # 没有 Tab 分隔：补一个 Tab 就是**插入字符**，属于改客户的文档内容 → 只提示。
                findings.append(Finding(
                    dimension="equation_number_position", severity="info",
                    message="Equation numbering",
                    expected="The equation number sits in parentheses, flush with the "
                             "right margin",
                    actual=("%s: the number follows the formula directly. Pushing it right "
                            "needs a tab character, which would mean inserting something "
                            "into your text - please add it in Word (type the number, then "
                            "Tab, and set a right tab stop at the margin)." % _eq_ids(_stuck)),
                    source=_o,
                ))
            if _without:
                # 补编号同样是加文字 → 只提示，绝不代写。
                findings.append(Finding(
                    dimension="equation_number_position", severity="info",
                    message="Equation numbering",
                    expected="Every display equation carries a number",
                    actual=("%s have no equation number. Adding one would insert text into "
                            "your document, so it is reported only." % _eq_ids(_without)),
                    source=_o,
                ))
            if not _bad and not _without:
                findings.append(Finding(
                    dimension="equation_number_position", severity="pass",
                    message="Equation numbering",
                    expected="The equation number sits in parentheses, flush with the "
                             "right margin",
                    actual="%d display equation(s), all numbered at the right margin"
                           % len(_eq_flat),
                    source=_o,
                ))

    _eq_size_want = getattr(target, "equation_size_pt", None)
    if _eq_size_want is not None and _eq_judge:
        _o = target.resolution.get("equation_size_pt", "school")
        _bad = []
        for s in _eq_disp:
            _sev, _act = _cmp_num(s.get("size_pt"), float(_eq_size_want),
                                  SIZE_TOL_FAIL, SIZE_TOL_WARN, " pt")
            if _sev != "pass":
                _bad.append((s, _act))
        if not _bad:
            findings.append(Finding(
                dimension="equation_size", severity="pass", message="Equation font size",
                expected="%g pt" % float(_eq_size_want),
                actual="%d display equation(s), all at %g pt"
                       % (len(_eq_disp), float(_eq_size_want)),
                source=_o,
            ))
        else:
            _f = Finding(
                dimension="equation_size", severity="warn", message="Equation font size",
                expected="%g pt" % float(_eq_size_want),
                actual="; ".join("%s: %s" % (_eq_ids([s]), _act) for s, _act in _bad),
                source=_o,
            )
            # 只写 m:r/w:rPr/w:sz（字号），绝不换数学字体、不碰公式内容 → 可修。
            _f.fix = FixAction(dimension="equation_size", value=float(_eq_size_want))
            findings.append(_f)

    # ⚠️ 注意三列分别是「target 字段名 / 采样字典里的键名 / 维度名」——
    # 采样用的是 ``space_before_pt``（无 equation 前缀，因为它描述的是「公式所在段落」），
    # 别拿 target 的字段名去取采样值，否则一律读成 None → 报「读不到」（实测踩过）。
    for _k, _sk, _dim, _label in (
            ("equation_space_before_pt", "space_before_pt",
             "equation_space_before", "Space before equations"),
            ("equation_space_after_pt", "space_after_pt",
             "equation_space_after", "Space after equations")):
        _w = getattr(target, _k, None)
        if _w is None or not _eq_judge:
            continue
        _o = target.resolution.get(_k, "school")
        _bad = []
        for s in _eq_disp:
            _sev, _act = _cmp_num(s.get(_sk), float(_w),
                                  SPACE_TOL_FAIL, SPACE_TOL_WARN, " pt")
            if _sev != "pass":
                _bad.append((s, _act))
        if not _bad:
            findings.append(Finding(
                dimension=_dim, severity="pass", message=_label,
                expected="%g pt" % float(_w),
                actual="%d display equation(s), all at %g pt" % (len(_eq_disp), float(_w)),
                source=_o,
            ))
        else:
            _f = Finding(
                dimension=_dim, severity="warn", message=_label,
                expected="%g pt" % float(_w),
                actual="; ".join("%s: %s" % (_eq_ids([s]), _act) for s, _act in _bad),
                source=_o,
            )
            # 只写段落段前/段后（并清掉 w:beforeLines 干扰），不动文字 → 可修。
            _f.fix = FixAction(dimension=_dim, value=float(_w))
            findings.append(_f)

    _eq_ls_want = getattr(target, "equation_line_spacing", None)
    if _eq_ls_want is not None and _eq_judge:
        _o = target.resolution.get("equation_line_spacing", "school")
        _bad = []
        for s in _eq_disp:
            _sev, _act = _cmp_num(s.get("line_spacing"), float(_eq_ls_want),
                                  SPACING_TOL_FAIL, SPACING_TOL_WARN, " lines")
            if _sev != "pass":
                _bad.append((s, _act))
        if not _bad:
            findings.append(Finding(
                dimension="equation_line_spacing", severity="pass",
                message="Equation line spacing",
                expected="%g lines" % float(_eq_ls_want),
                actual="%d display equation(s), all at %g lines"
                       % (len(_eq_disp), float(_eq_ls_want)),
                source=_o,
            ))
        else:
            _f = Finding(
                dimension="equation_line_spacing", severity="warn",
                message="Equation line spacing",
                expected="%g lines" % float(_eq_ls_want),
                actual="; ".join("%s: %s" % (_eq_ids([s]), _act) for s, _act in _bad),
                source=_o,
            )
            _f.fix = FixAction(dimension="equation_line_spacing", value=float(_eq_ls_want))
            findings.append(_f)

    # 公式字体：**只提示、绝不自动替换**。
    # 数学区换字体在 Word 里不是「改个 w:rFonts」这么简单 —— 字形映射一变，符号可能
    # 直接渲染成空框（客户打开就是一堆方框，比不对齐严重得多）。宁可不改，如实告知。
    _eq_font_want = getattr(target, "equation_font_family", None)
    if _eq_font_want and _eq_samples:
        findings.append(Finding(
            dimension="equation_font", severity="info", message="Equation font",
            expected="%s, as specified in your template" % _eq_font_want,
            actual="Not changed automatically: substituting the font inside Word equations "
                   "can make symbols render as empty boxes. If it matters, set it in Word "
                   "(Equation Tools -> Conversions) after applying the other fixes.",
            source=target.resolution.get("equation_font_family", "school"),
        ))

    # —— 表格内段落：段前 / 段后（v2.3.77）——
    # 表格里「不要多余空行」是学校硬要求里很常见的一条；表格段落与正文是两套口径
    # （表格通常 0 磅、正文常有 6–12 磅），所以单独立维度、单独采样。
    for _k, _dim, _label in (
            ("table_space_before_pt", "table_space_before",
             "Table cell paragraph space before"),
            ("table_space_after_pt", "table_space_after",
             "Table cell paragraph space after")):
        _w = getattr(target, _k, None)
        if _w is None:
            continue
        _o = target.resolution.get(_k, "school")
        _got = getattr(prof, _k, None)
        if not getattr(prof, "table_paragraph_count", 0) or _got is None:
            findings.append(Finding(
                dimension=_dim, severity="info", message=_label,
                expected="%g pt" % float(_w),
                actual="No table text detected" if not getattr(prof, "table_paragraph_count", 0)
                else "Could not read the table cell spacing",
                source=_o,
            ))
            continue
        _sev, _act = _cmp_num(_got, float(_w), SPACE_TOL_FAIL, SPACE_TOL_WARN, "pt")
        sev = _downgrade(_sev, _o, is_page_or_heading=True,
                         dim=_dim, spec_key=target.spec_key)
        findings.append(Finding(
            dimension=_dim, severity=sev, message=_label,
            expected="%g pt" % float(_w), actual=_act, source=_o,
        ))

    # —— 目录（TOC，v2.3.77）——
    # 🔴 目录条目是**域段落**：Word「更新目录(F9)」时用 TOC 1..9 样式整体重排，
    # 逐条改直接格式会被冲掉 —— 所以条目只核对、不修改；能安全改的只有目录标题那一段。
    _toc_on = bool(getattr(prof, "toc_detected", False))
    for _k, _dim, _label in (
            ("toc_heading_space_before_pt", "toc_heading_space_before",
             "Table of contents heading space before"),
            ("toc_heading_space_after_pt", "toc_heading_space_after",
             "Table of contents heading space after")):
        _w = getattr(target, _k, None)
        if _w is None:
            continue
        _o = target.resolution.get(_k, "school")
        _got = getattr(prof, _k, None)
        if not _toc_on or _got is None:
            findings.append(Finding(
                dimension=_dim, severity="info", message=_label,
                expected="%g pt" % float(_w),
                actual="No table of contents detected" if not _toc_on
                else "The TOC heading paragraph could not be identified",
                source=_o,
            ))
            continue
        _sev, _act = _cmp_num(_got, float(_w), SPACE_TOL_FAIL, SPACE_TOL_WARN, "pt")
        sev = _downgrade(_sev, _o, is_page_or_heading=True,
                         dim=_dim, spec_key=target.spec_key)
        findings.append(Finding(
            dimension=_dim, severity=sev, message=_label,
            expected="%g pt" % float(_w), actual=_act, source=_o,
        ))
    if _toc_on:
        # 条目间距：只核对、不修改。目标给了值就报差异（让客户知道差多少），
        # 但 severity 恒为 info、fix 恒为 None —— 我们绝不替客户改域段落。
        _e_bits = []
        _cur_sb = getattr(prof, "toc_entry_space_before_pt", None)
        _cur_sa = getattr(prof, "toc_entry_space_after_pt", None)
        _want_sb = getattr(target, "toc_entry_space_before_pt", None)
        _want_sa = getattr(target, "toc_entry_space_after_pt", None)
        if _cur_sb is not None:
            _e_bits.append("space before %g pt" % float(_cur_sb))
        if _cur_sa is not None:
            _e_bits.append("space after %g pt" % float(_cur_sa))
        _want_bits = []
        if _want_sb is not None:
            _want_bits.append("space before %g pt" % float(_want_sb))
        if _want_sa is not None:
            _want_bits.append("space after %g pt" % float(_want_sa))
        findings.append(Finding(
            dimension="toc_entry_spacing", severity="info",
            message="Table of contents entry spacing",
            expected=(", ".join(_want_bits) if _want_bits else ""),
            actual=("%d entries: %s. Entry spacing is generated by Word from the "
                    "TOC 1-9 styles, so it is NOT auto-fixed here — update the TOC "
                    "styles in Word if it needs to change."
                    % (int(getattr(prof, "toc_entry_count", 0) or 0),
                       ", ".join(_e_bits) if _e_bits else "could not be read")),
            source=target.resolution.get("toc_entry_space_before_pt", ""),
        ))

    # —— 脚注 / 尾注：本工具不检查也不修改（v2.3.77）——
    # 客户文档里常有大量页下注。报告全绿时如果不说明白，客户会以为「脚注也查过了」。
    # 这里只给一条 info（fix 恒为 None），把数量报出来、把边界讲清楚。
    _fn = int(getattr(prof, "footnote_count", 0) or 0)
    _en = int(getattr(prof, "endnote_count", 0) or 0)
    if _fn or _en:
        _parts = []
        if _fn:
            _parts.append("%d footnote%s" % (_fn, "" if _fn == 1 else "s"))
        if _en:
            _parts.append("%d endnote%s" % (_en, "" if _en == 1 else "s"))
        findings.append(Finding(
            dimension="footnote", severity="info",
            message="Footnotes / endnotes are not checked",
            expected="",
            actual="%s detected — their font, size, line spacing and paragraph spacing "
                   "were left untouched; please verify them manually"
                   % " and ".join(_parts),
            source="",
        ))

    # —— 前置材料：按各自 tier 分流（学生/专业差异项，不混用）——
    # 专业版（expect_abstract / expect_author_note = True）才提示「应有」；
    # 学生版（均 False）不要求，仅在「多写了作者注」时温和提示非必需。
    # v2.3.100（C 批 D-08 / 裁定 C-7）：当规范**显式规定了摘要版式**（IEEE 会议档给了
    # abstract_font_size_pt 等）时，把"必须有摘要"升为 **hard**，并逐项判版式：
    #   9pt + 加粗 + 两端对齐 + 段后 10pt（依据：模板 Abstract 样式
    #   `<w:sz w:val="18"/>` + `<w:b/>` + `jc=both` + `<w:spacing w:after="10pt"/>`；
    #   OWL「All papers must begin with an abstract」）。
    # APA 专业版没有这些字号字段 → 仍是原来的 info，行为不变。
    # v2.3.114：模板驱动优先。学校模板批注写了摘要正文版式（如 12 pt、不加粗）就按模板，
    # 规范硬编码值只在目标没显式给时兜底。此前直接取 apa_spec（= 当前 spec，IEEE 会议档
    # 硬编码 9pt 加粗），导致学校模板模式下摘要被错误套成 9pt 加粗、还加了粗。
    _t_abs_bold = getattr(target, "abstract_bold", None)
    _abs_want = (
        getattr(target, "abstract_font_size_pt", None)
        or getattr(apa_spec, "abstract_font_size_pt", None),
        _t_abs_bold if _t_abs_bold is not None
        else getattr(apa_spec, "abstract_bold", None),
        getattr(target, "abstract_alignment", None)
        or getattr(apa_spec, "abstract_alignment", None),
        getattr(target, "abstract_space_after_pt", None)
        if getattr(target, "abstract_space_after_pt", None) is not None
        else getattr(apa_spec, "abstract_space_after_pt", None),
    )
    _abs_hard = apa_spec.expect_abstract and any(v is not None for v in _abs_want)
    if apa_spec.expect_abstract and not prof.has_abstract:
        findings.append(Finding(
            dimension="abstract",
            severity="fail" if _abs_hard else "info",
            message="Abstract & Keywords" if not _abs_hard else "Abstract section",
            expected=("An Abstract with Keywords is required (APA Professional)" if not _abs_hard
                      else "An Abstract section is required"),
            actual="No Abstract section detected",
            source="spec",
        ))
    elif _abs_hard and prof.has_abstract:
        _w_sz, _w_bold, _w_al, _w_sa = _abs_want
        _bad, _act = [], []
        if _w_sz is not None:
            _cur = getattr(prof, "abstract_font_size_pt", None)
            if _cur is None:
                _act.append("size could not be read")
            elif abs(float(_cur) - float(_w_sz)) > 0.5:
                _bad.append("size %g pt" % float(_cur))
            _act.append("size %g pt" % float(_cur)) if _cur is not None else None
        if _w_bold is not None:
            _cur_b = getattr(prof, "abstract_bold", None)
            if _cur_b is not None and bool(_cur_b) != bool(_w_bold):
                _bad.append("bold" if _cur_b else "not bold")
        if _w_al is not None:
            _cur_al = getattr(prof, "abstract_alignment", None)
            if _cur_al and _cur_al != _w_al:
                _bad.append("%s-aligned" % _cur_al)
        if _w_sa is not None:
            _cur_sa = getattr(prof, "abstract_space_after_pt", None)
            if _cur_sa is not None and abs(float(_cur_sa) - float(_w_sa)) > 0.5:
                _bad.append("space after %g pt" % float(_cur_sa))
        if _bad:
            findings.append(Finding(
                dimension="abstract", severity="warn",
                message="Abstract format",
                # v2.3.100（D-08）：挂 FixAction，把「规范要求」原样带给 fixer ——
                # 不挂的话客户点「修正」这条红杠不会消失（检查说能改、实际没改）。
                fix=FixAction(
                    dimension="abstract",
                    value={"size_pt": _w_sz, "bold": _w_bold,
                           "alignment": _w_al, "space_after_pt": _w_sa},
                    note=""),
                expected=", ".join(filter(None, [
                    ("%g pt" % float(_w_sz)) if _w_sz is not None else "",
                    "bold" if _w_bold else ("not bold" if _w_bold is False else ""),
                    str(_w_al) if _w_al else "",
                    ("space after %g pt" % float(_w_sa)) if _w_sa is not None else ""])),
                actual=", ".join(_bad),
                source="spec",
            ))
        else:
            findings.append(Finding(
                dimension="abstract", severity="pass",
                message="Abstract format",
                expected="Abstract formatting as specified",
                actual=", ".join(_act) or "matches",
                source="spec",
            ))
    if apa_spec.expect_author_note and not prof.has_author_note:
        findings.append(Finding(
            dimension="author_note", severity="info",
            message="Author Note",
            expected="An Author Note is expected (APA Professional)",
            actual="No Author Note section detected",
            source="spec",
        ))
    if (not apa_spec.expect_author_note) and apa_spec.reference.style_key == "APA" and prof.has_author_note:
        # 仅「APA 学生版」触发（running_head=False 且 style_key="APA"）：学生论文通常不要求作者注，
        # 写了也仅作温和提示。注意：MLA/IEEE/Other 的 expect_author_note 同为 False，但绝不应触发此提示
        # （它们不是 APA 学生版），故用 style_key=="APA" 收窄，避免误伤其他风格。
        findings.append(Finding(
            dimension="author_note", severity="info",
            message="Author Note (not required for student papers)",
            expected="Usually not required for APA Student papers",
            actual="An Author Note section was found; include it only if your instructor asks",
            source="spec",
        ))

    if apa_spec.reference_alpha_sort_hint:
        # 参考文献字母序仅做提示，不改动内容（用户明确：排序涉及改内容，只提示）。
        if prof.reference_entries and len(prof.reference_entries) > 1:
            if _is_sorted_alphabetically(prof.reference_entries):
                findings.append(Finding(
                    dimension="reference_sort", severity="pass",
                    message="Reference list order",
                    expected="Entries sorted alphabetically by first author's surname",
                    actual="Appears alphabetically ordered",
                    source="spec",
                ))
            else:
                findings.append(Finding(
                    dimension="reference_sort", severity="info",
                    message="Reference list order",
                    expected="Entries sorted alphabetically by first author's surname",
                    actual="Entries do not appear to be in alphabetical order (MLA/APA require alphabetical order)",
                    source="spec",
                ))

    # —— 参考文献标题格式（APA 7th：'References' 须粗体+居中于页顶；MLA 9th：'Works Cited' 居中、不加粗）——
    # 安全处理：官方样例用 Heading 1 样式承载，粗体/居中来自样式（python-docx 读不出样式解析值），
    # 故 is_style 为真时「信任样式」直接判 pass，绝不误伤金标准论文；仅当标题是「手动段落」
    # （非 Heading 样式）时才检查直接格式，缺失则给 info 提示（不 fail、不改内容）。
    if prof.reference_entries and spec.reference.style_key in ("APA", "MLA") \
            and target.reference_heading_bold is None \
            and target.reference_heading_centered is None:
        if prof.reference_heading_is_style:
            findings.append(Finding(
                dimension="reference_heading", severity="pass",
                message="Reference heading format",
                expected=f"{ref_label} uses a Heading style"
                          + (" (bold/centered supplied by the style)" if spec.reference_heading_bold_mode == "require"
                             else " (centered, not bold supplied by the style)"),
                actual="Uses a Heading style",
                source="spec",
            ))
        else:
            problems = []
            if spec.reference_heading_bold_mode == "forbid":
                if prof.reference_heading_bold:
                    problems.append("bolded")
            elif spec.reference_heading_bold and not prof.reference_heading_bold:
                problems.append("not bold")
            if not prof.reference_heading_centered:
                problems.append("not centered")
            if problems:
                expected = f"{ref_label} should be centered" + \
                           (" and not bold" if spec.reference_heading_bold_mode == "forbid"
                            else " and bold" if spec.reference_heading_bold else "") + \
                           " at the top of a new page"
                findings.append(Finding(
                    dimension="reference_heading", severity="info",
                    message="Reference heading format",
                    expected=expected,
                    actual=f"The '{ref_label}' heading is " + " and ".join(problems),
                    source="spec",
                ))

    # —— MLA 专属版式检查（仅信息/提示级，绝不改动内容；引文内容不查）——
    bq: list = []      # 待修块引用候选（供下方 FixAction 绑定用；非 MLA 时为 [];）
    if spec.reference.style_key == "MLA":
        # G3 标题格式：居中、不加粗、不斜体、不下划线（Purdue OWL general_format）
        if prof.title_text:
            t_problems = []
            if prof.title_alignment != "center":
                t_problems.append("not centered")
            if prof.title_bold:
                t_problems.append("bolded")
            if prof.title_italic:
                t_problems.append("italicized")
            if prof.title_underline:
                t_problems.append("underlined")
            if t_problems:
                findings.append(Finding(
                    dimension="title_format", severity="info",
                    message="Title format",
                    expected="Centered, not bold / italic / underlined (MLA)",
                    actual="The title is " + " and ".join(t_problems),
                    source="spec",
                ))
            else:
                findings.append(Finding(
                    dimension="title_format", severity="pass",
                    message="Title format",
                    expected="Centered, not bold/italic/underlined",
                    actual="Centered, unformatted",
                    source="spec",
                ))
        # G2 首页信息块：标题之前应有左上对齐的 姓名/教授/课程/日期（双行距）共约 4 行
        if prof.title_text and prof.first_page_left_lines < 3:
            _fpl_txt = list(getattr(prof, "first_page_left_texts", None) or [])
            _why = ("Could not clearly detect the 4-line MLA heading block before the title"
                    + (" (%d left-aligned line(s) detected: %s)"
                       % (len(_fpl_txt),
                          "; ".join(repr(t) for t in _fpl_txt[:4]))
                       if _fpl_txt else ""))
            findings.append(Finding(
                dimension="first_page_heading", severity="info",
                message="First-page heading block",
                expected="Upper left: your name, instructor, course, date (double-spaced)",
                actual=_why,
                source="spec",
            ))
        # G5 块引用：**停判**（v2.3.100，E-0 / 裁定 D-09）。
        # 🔴 老板原话：「按你建议——[材料缺失]，代码里 0.5″ 硬编码停判（不降 info，
        # 直接不判），等 MLA 手册条文到位再恢复。理由：无依据的规则判了就是误报。」
        # 本轮 MLA 材料（8 份学生论文 + OWL 带批注样例）里**没有任何块引用缩进条文**，
        # 那句「0.5″」是我凭印象写的 —— 按 P-3「材料缺失一律标 [材料缺失]，不许造规则」，
        # 这条判据整个撤掉（含 checker 的 finding 与 fixer 的修正值），不降级保留。
        # 恢复条件：拿到 MLA Handbook 的块引用缩进条文后，把下面的代码原样加回。
        pass

    # —— 把「需要修正」的 Finding 绑定 FixAction（check→fix 单一真相源）——
    # fixer 只消费带 fix 的 Finding，绝不再独立重算「要改什么」，从而确保
    # 「检查出来的问题」与「实际改的内容」必然一致。
    _fix_value = {
        "margins": target.page.get("margin_top_in"),
        "font": target.page.get("font_family"),
        "font_size": target.page.get("font_size_pt"),
        "line_spacing": target.page.get("line_spacing"),
        "alignment": target.page.get("body_alignment"),
        "first_line_indent": target.page.get("first_line_indent_in"),
        # v2.3.75：正文段前（此前只有段后，段落间距只能核一半）
        "space_before": target.page.get("space_before_pt"),
        "space_after": target.page.get("space_after_pt"),
        "header_distance": target.page.get("header_distance_in"),
        "columns": target.columns,
        "paper_size": target.paper_size,
        "reference_hanging_indent": target.reference_hanging_indent_in,
    }
    # 参考文献细项只在「目标显式给出」时才允许修正：值来自规范软参考（spec）时保持
    # 「仅提示」，避免把提示类维度悄悄升级成会改文档的动作。
    for _k, _dim in (("reference_font_family", "reference_font"),
                     ("reference_size_pt", "reference_size"),
                     ("reference_line_spacing", "reference_spacing"),
                     ("reference_space_before_pt", "reference_space_before"),
                     ("reference_space_after_pt", "reference_space_after"),
                     ("reference_new_page", "reference_new_page"),
                     ("reference_heading", "reference_heading")):
        if _k == "reference_heading":
            _v = (target.reference_heading_bold
                  if target.reference_heading_bold is not None
                  else target.reference_heading_centered)
            if target.reference_heading_bold is None \
                    and target.reference_heading_centered is None:
                continue
        else:
            _v = getattr(target, _k, None)
        if _v is not None:
            _fix_value[_dim] = _v
    # 标题页 / 块引用：这两项**会实际改动文档**，必须由检查结论驱动。
    # 旧版这两处是 fixer 里不受 gating 的自动动作（客户看不到、报告里也没有）——
    # 违反「检查→修正单一真相源」。现在补成正常维度：检查出结论 → 绑 FixAction →
    # 修正器才动手 → 修正预览里也能看见。
    if _tp_vals:
        _fix_value["title_page"] = True
    # v2.3.91：首页结构三件套 —— 可修。v2.3.93：**三个通道齐平**，
    # 客户显式要求（学校模板 / 手填 / AI）与选了内置规范同等对待，都能落地。
    # 全都不新增段落或文字，只写 pageBreakBefore 与段距。
    # 只绑「读得出来且确实不符」的情形：读不到（None）时给的是 info「无法判定」，
    # 此时绑一个什么都不做的 FixAction 只会让报告说「可自动修正」却修不动。
    _fm_st = resolve_spec(target)
    _fm_tpf = getattr(_fm_st, "title_page", None)
    _fm_prof = prof
    if getattr(_fm_prof, "title_page_detected", False):
        # ① 题名页独立成页：显式值优先，其次规范基线
        _tpn_bind = getattr(target, "title_page_new_page", None)
        if _tpn_bind is None and spec_is_baseline(target):
            _tpn_bind = getattr(_fm_st, "title_page_new_page", None)
        _tpn_cur = getattr(_fm_prof, "tp_next_break_before", None)
        if _tpn_bind is not None and _tpn_cur is not None \
                and bool(_tpn_cur) != bool(_tpn_bind):
            _fix_value["title_page_new_page"] = bool(_tpn_bind)
        # ② 题名页版式：规范基线听规范原话；非基线只在客户勾了
        # 「元素间不留额外段距」时才动手（①② 的磅值由 title_page 维度各自负责）
        _lay_bind = bool(
            spec_is_baseline(target) and _fm_tpf is not None
            and (getattr(_fm_tpf, "title_space_before_lines", None) is not None
                 or getattr(_fm_tpf, "title_space_after_lines", None) is not None
                 or getattr(_fm_tpf, "zero_extra_spacing", False)))
        if not _lay_bind and getattr(target, "title_page_zero_extra_spacing", False):
            _lay_bind = True
        if _lay_bind:
            _fix_value["title_page_layout"] = True
    # ③ 正文另起一页
    _txn_bind = getattr(target, "text_new_page", None)
    if _txn_bind is None and spec_is_baseline(target):
        _txn_bind = getattr(_fm_st, "text_new_page", None)
    _txn_cur = getattr(_fm_prof, "text_break_before", None)
    if _txn_bind is not None and bool(getattr(_fm_prof, "has_abstract", False)) \
            and _txn_cur is not None and bool(_txn_cur) != bool(_txn_bind):
        _fix_value["text_new_page"] = bool(_txn_bind)
    # v2.3.100（E-0 / D-09）：块引用缩进**停判** → 不再给 fixer 传修正值。
    # （原先恒为 0.5″，是无材料依据的硬编码；材料缺失，按 P-3 不许造规则。）
    # _fix_value["block_quote"] = 0.5
    # 页码位置：**只有客户/学校/AI 显式要求时才动手**（来源 spec 时只提示）。
    # 理由与纸型一致 —— 国内学校普遍要「页脚居中」，把 APA 的 top-right 当硬要求去改，
    # 会把学生本来正确的页脚居中改成页眉右上。客户显式选了才改。
    if _pn_pos and _pn_explicit and prof.page_number_present:
        _fix_value["page_number"] = _pn_pos
    elif _pn_pos and _pn_explicit and not prof.page_number_present:
        _fix_value["page_number"] = _pn_pos      # 完全没有页码时，显式要求就补一个
    elif _pn_pos and spec_is_baseline(target):
        # 🔴 v2.3.90（老板 2026-10-22 拍板「升为可修」）：客户真的在规范下拉里
        # 选了内置规范 = 全篇按该规范走，页码是它的硬要求，必须落地。
        # 这里**不区分**「完全没有页码」与「页码在别处/对齐不对」：两者都按规范重写
        # （_apply_page_number 会先写目标位置、再清另一处，绝不会出现两个页码）。
        # 保持「检查报了 → 修正就能改」，不留 v2.3.71 那类报了却修不了的缺口。
        _fix_value["page_number"] = _pn_pos
    # 🔴 v2.3.100（落地裁定 C-4）：**不再为 font_size_by_font 挂修正值**。
    # 旧代码在这里把整张 6 款字体配对表塞进 _fix_value，fixer 据此逐 run 刷字号 ——
    # 既替客户做了字体决定，也正中「官方学生样例 TNR 11pt 被判红」（13 份里 5 份命中）。
    # 现在该维度只提示（checker 侧已降 info），_fix_value 不再给它传值。
    # if font_size_by_font:
    #     _fix_value["font_size_by_font"] = dict(allow_pairs)
    # 摘要页特化：**只有客户/学校/AI 显式要求时才动手**（来源 spec 时只提示）。
    # 理由与页码位置一致 —— 摘要是否缩进/另起一页，不同学校要求不同，APA 的
    # 「顶格 + 另起一页」只是典型值，不能替客户拍板去改摘要页。
    if target.abstract_first_line_indent_in is not None \
            and origin_of_source(target.resolution.get("abstract_first_line_indent_in")) \
            in ("user", "school", "ai_import"):
        _fix_value["abstract_indent"] = float(target.abstract_first_line_indent_in)
    # v2.3.91：规范基线（客户真在规范下拉里选了内置规范）时，摘要另起一页也可修。
    # 与页码同口径 —— APA 7 §2.9 对两档都要求「摘要自成一页」，这不该被
    # 「页面维度降级 info」吞掉；学校模板 / 手填 / AI 导入不受影响。
    if target.abstract_new_page is not None and (
            origin_of_source(target.resolution.get("abstract_new_page"))
            in ("user", "school", "ai_import") or spec_is_baseline(target)):
        _fix_value["abstract_new_page"] = bool(target.abstract_new_page)
    # 图表 caption 字号/行距可修：**只有客户/学校/AI 显式要求时才动手**（spec 只提示）。
    # caption 位置只提示（不绑 fix）。
    if target.caption_font_size_pt is not None \
            and origin_of_source(target.resolution.get("caption_font_size_pt")) \
            in ("user", "school", "ai_import"):
        _fix_value["caption_size"] = float(target.caption_font_size_pt)
    if target.caption_line_spacing is not None \
            and origin_of_source(target.resolution.get("caption_line_spacing")) \
            in ("user", "school", "ai_import"):
        _fix_value["caption_spacing"] = float(target.caption_line_spacing)
    # v2.3.76：摘要 / 图表标题的段前段后 —— 同样只在**显式来源**时才动手
    # （规范没给这两项，全部由模板 / 问卷 / AI 驱动；来源 spec 时只提示）。
    for _key, _dim in (("abstract_space_before_pt", "abstract_space_before"),
                       ("abstract_space_after_pt", "abstract_space_after"),
                       ("caption_space_before_pt", "caption_space_before"),
                       ("caption_space_after_pt", "caption_space_after")):
        _v = getattr(target, _key, None)
        if _v is not None and origin_of_source(target.resolution.get(_key)) \
                in ("user", "school", "ai_import"):
            _fix_value[_dim] = float(_v)
    # v2.3.77：表格内段落段前段后（可修）+ 目录标题段前段后（可修）。
    # ⚠️ 目录**条目**间距有意**不**进 _fix_value：条目是域段落，改直接格式会被
    # 「更新目录」冲掉，只核对不修改（toc_entry_spacing 恒为 info、fix=None）。
    for _key, _dim in (("table_space_before_pt", "table_space_before"),
                       ("table_space_after_pt", "table_space_after"),
                       ("toc_heading_space_before_pt", "toc_heading_space_before"),
                       ("toc_heading_space_after_pt", "toc_heading_space_after")):
        _v = getattr(target, _key, None)
        if _v is not None and origin_of_source(target.resolution.get(_key)) \
                in ("user", "school", "ai_import"):
            _fix_value[_dim] = float(_v)

    for f in findings:
        if f.severity == "pass":
            continue
        if f.dimension in _fix_value:
            f.fix = FixAction(dimension=f.dimension, value=_fix_value[f.dimension])
        elif f.dimension == "reference" and target.reference_numbered:
            # IEEE 顺序编码制：references 维度发现偏差（缺 [n]）→ 绑定编号修正
            f.fix = FixAction(dimension="reference_numbered", value=True)
        elif f.dimension == "table_lines" and target.table_lines == "three_line":
            # v2.3.85：三线表可修。与栏数同理，规范来源就是显式版式意图
            # （APA/IEEE/MLA 三家的表格都不许有竖线），故 spec 来源也绑定修正。
            f.fix = FixAction(dimension="table_lines", value="three_line")
        elif f.dimension == "running_head" and spec_is_baseline(target) \
                and getattr(resolve_spec(target), "running_head", False):
            # 🔴 v2.3.90（老板 2026-10-22 拍板「做」）：running head 可修。
            # 文字从**客户自己的论文标题**（APA 短标题）或**自己的姓名行**（MLA 姓氏）
            # 派生，只写页眉、正文一字不动；派生结果在预览与报告里逐字列出。
            f.fix = FixAction(
                dimension="running_head",
                value=getattr(resolve_spec(target), "running_head_kind", "apa"))
        # reference_new_page / font_allowlist（字体本身，不换字体）/
        # reference_heading / reference_sort / title_format / first_page_heading /
        # abstract / author_note / 内容提示类：均为「仅提示、改内容」项，fix=None → fixer 绝不碰。
        # （page_number 与 font_size_by_font 已在 _fix_value 里，可修。）

    # 标题层级编号/格式：由 target.heading_levels 驱动（而非 finding 严重度），
    # 只要目标要求标题层级 > 0 即绑定 heading 修正，与 fixer 既有行为一致。
    # v2.3.65：每级标题格式的 finding 也挂在 heading 维度上，这里一并绑定（不再 break 一条）。
    # v2.3.70：把**具体层级**写进 FixAction.levels —— 「Level 3 heading format」这类
    # 检查结果本来就精确到级，修正器据此只改被报错的那级（否则全五级重刷＝不按结果修）。
    if (target.heading_levels and target.heading_levels > 0) or heading_targets:
        _lvl_value = target.heading_levels or (max(heading_targets) if heading_targets else 0)
        for f in findings:
            if f.dimension == "heading":
                # 🔴 v2.3.95：「Heading levels」这一条讲的是**标题层级结构**
                # （有没有套标题样式 / 层级数超没超），修正器**造不出标题、也合并不了层级**
                # —— 标题是客户自己写的内容。旧版无条件给它绑 heading 修正，于是
                # 「报告说可自动修正 → 客户点了 → 同一个 warn 还在」（老板 2026-10-23
                # 四份产物实测：APA 两份修完复检仍报 `No Heading styles … levels are
                # inferred from font size`）。改正文结构的活只能客户自己做，故只提示。
                # 每级**格式**的 finding 不在此列，照旧精确到级地绑定修正。
                if f.message == "Heading levels":
                    continue
                _lv = _finding_heading_level(f.message)
                f.fix = FixAction(dimension="heading", value=_lvl_value,
                                  levels=(_lv,) if _lv else ())

    return prof, findings


def summarize(findings: list[Finding]) -> dict:
    counts = {"pass": 0, "info": 0, "warn": 0, "fail": 0}
    for f in findings:
        counts[f.severity] = counts.get(f.severity, 0) + 1
    return counts
