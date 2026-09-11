"""LaTeX 模板解析器（第⑤种目标画像来源）。

从 .cls / .sty / .tex 文件中抽取格式要求，产出 partial TargetProfile，
与 docx 模板 / 问卷 / AI 填表 JSON 走**同一套** build_target 合并。

设计铁律（边界明确，不是 bug，是上限）：
- **不跑 TeX 引擎**：只做正则/结构解析，不编译。
- **不解析自定义宏**：\\newcommand 重定义、条件分支只读字面声明。
- **先剥离注释**：`%` 之后（含被注释掉的 \\doublespacing / \\bibliographystyle）一律不采信。
- **覆盖率目标 = 主流模板**：IEEEtran / apa6-7 / article+geometry+setspace / revtex 等；
  引用风格只映射到已登记的五种（APA/MLA/Chicago/IEEE/Harvard），
  未登记的（如 ACM 数字制）**不猜**，留空回落规范。
- 拿不全的字段留空，由规范(spec)兜底；不构造幻觉值。
- 与海外版其他模块一样：纯本地、零网络、零 LLM。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

from .questionnaire import TargetProfile
from .specs import REFERENCE_STYLE_KEYS, get_spec

SOURCE_LABEL = "latex_template"


# ---------------------------------------------------------------------------
# 引用样式映射表（未知→留空回退规范）
# ---------------------------------------------------------------------------

_BIB_STYLE_MAP = {
    # IEEE 家族
    "ieeetr": "IEEE", "ieee": "IEEE", "ieeetran": "IEEE", "ieeeconf": "IEEE",
    # APA 家族
    "apacite": "APA", "apalike": "APA", "apa": "APA", "apa6": "APA", "apa7": "APA",
    # MLA
    "mla": "MLA", "mla7": "MLA", "mla9": "MLA",
    # Chicago
    "chicago": "Chicago", "chicago-authordate": "Chicago",
    # Harvard 家族
    "agsm": "Harvard", "dcu": "Harvard", "kuleuven": "Harvard", "harvard": "Harvard",
}


def _map_bib_style(name: str) -> Optional[str]:
    n = name.lower().strip()
    if n in _BIB_STYLE_MAP:
        return _BIB_STYLE_MAP[n]
    # 模糊匹配（包含）
    for k, v in _BIB_STYLE_MAP.items():
        if k in n or n in k:
            return v
    return None


def _to_inches(val: float, unit: str) -> float:
    u = unit.lower()
    if u == "in":
        return val
    if u == "cm":
        return val / 2.54
    if u == "mm":
        return val / 25.4
    if u == "pt":
        return val / 72.0
    if u == "bp":
        return val / 72.0
    return val  # 未知单位原样返回，由调用方识别


def _strip_comments(text: str) -> str:
    """剥离 LaTeX 注释：`%` 到行尾（保留 `\\%` 这类转义百分号）。

    LaTeX 规则：`%` 前若有**奇数个**连续反斜杠，则是被转义的百分号（`\\%`），
    不算注释起点；偶数个则 `%` 起始注释（`\\\\%` = 换行后注释）。
    不剥离注释会让 `% \\doublespacing`、`% \\bibliographystyle{apacite}` 这类
    被注释掉的声明被当成真实格式要求。
    """
    out: list = []
    for line in text.splitlines():
        cut = None
        i = 0
        while True:
            j = line.find("%", i)
            if j < 0:
                break
            bs = 0
            k = j - 1
            while k >= 0 and line[k] == "\\":
                bs += 1
                k -= 1
            if bs % 2 == 0:
                cut = j
                break
            i = j + 1
        out.append(line[:cut] if cut is not None else line)
    return "\n".join(out)


# ---------------------------------------------------------------------------
# 抽取基元（每个负责一件事，宽松容错）
# ---------------------------------------------------------------------------

def _extract_font_size(text: str) -> Optional[float]:
    """\\documentclass[10pt,...]{...} 或 \\documentclass[12pt,...]{...}。"""
    m = re.search(r"\\documentclass\s*\[([^\]]*)\]", text)
    if not m:
        return None
    size = re.search(r"(\d+(?:\.\d+)?)\s*pt", m.group(1))
    if size:
        return float(size.group(1))
    return None


def _extract_font_family_hint(text: str) -> Optional[str]:
    """根据 \\documentclass 选项或字体包给个粗略提示。"""
    m = re.search(r"\\documentclass\s*\[([^\]]*)\]", text)
    opts = m.group(1).lower() if m else ""
    if "times" in opts or re.search(r"\\usepackage\s*\{?\s*(times|newtxtext|mathptmx)", text):
        return "Times New Roman"
    if "helvetica" in opts or re.search(r"\\usepackage\s*\{?\s*helvet", text):
        return "Arial"
    return None


def _extract_geometry_margins(text: str) -> dict:
    """geometry 包参数：\\geometry{options} 或 \\usepackage[options]{geometry}。"""
    out: dict = {}
    opts: str = ""
    # 形式一：\geometry{top=...,left=...}
    m = re.search(r"\\geometry\s*\{([^}]*)\}", text)
    if m:
        opts = m.group(1)
    else:
        # 形式二：\usepackage[options]{geometry}
        m = re.search(r"\\usepackage\s*\[\s*([^\]]*)\s*\]\s*\{geometry\}", text)
        if m:
            opts = m.group(1)
        else:
            return out
    # 整体 margin=X 作为基线
    full = re.search(r"\bmargin\s*=\s*([\d.]+)\s*(cm|in|mm|pt|bp)", opts)
    if full:
        inches = round(_to_inches(float(full.group(1)), full.group(2)), 3)
        for k in ("margin_top_in", "margin_bottom_in", "margin_left_in", "margin_right_in"):
            out[k] = inches
    # 分边覆盖整体（典型 LaTeX 行为：per-side 优先）
    for side in ("top", "bottom", "left", "right"):
        ms = re.search(rf"\b{side}\s*=\s*([\d.]+)\s*(cm|in|mm|pt|bp)", opts)
        if ms:
            inches = round(_to_inches(float(ms.group(1)), ms.group(2)), 3)
            out[f"margin_{side}_in"] = inches
    return out


def _extract_line_spacing(text: str) -> Optional[float]:
    """setspace 包：\\setstretch{X} 优先（更具体），否则 \\onehalfspacing / \\doublespacing / \\singlespacing。"""
    m = re.search(r"\\setstretch\s*\{([\d.]+)\}", text)
    if m:
        return float(m.group(1))
    if re.search(r"\\onehalfspacing", text):
        return 1.5
    if re.search(r"\\doublespacing", text):
        return 2.0
    if re.search(r"\\singlespacing", text):
        return 1.0
    return None


def _extract_bib_style(text: str) -> Optional[str]:
    """\\bibliographystyle{...} 或 biblatex \\usepackage[style=...]{biblatex}。"""
    m = re.search(r"\\bibliographystyle\s*\{([^}]+)\}", text)
    if m:
        return _map_bib_style(m.group(1))
    m = re.search(r"\\usepackage\s*\[[^\]]*style\s*=\s*([\w-]+)", text)
    if m and "biblatex" in text:
        return _map_bib_style(m.group(1))
    return None


# ---------------------------------------------------------------------------
# 主解析入口
# ---------------------------------------------------------------------------

def parse_latex_source(text: str) -> TargetProfile:
    """从 LaTeX 源码字符串抽取格式要求 → partial TargetProfile。"""
    text = _strip_comments(text)      # 注释一律不采信（先剥离再解析）
    page: dict = {}
    resolution: dict = {}

    size = _extract_font_size(text)
    if size is not None:
        page["font_size_pt"] = float(size)
        resolution["font_size_pt"] = SOURCE_LABEL

    family = _extract_font_family_hint(text)
    if family:
        page["font_family"] = family
        resolution["font_family"] = SOURCE_LABEL

    margins = _extract_geometry_margins(text)
    for k, v in margins.items():
        page[k] = v
        resolution[k] = SOURCE_LABEL

    sp = _extract_line_spacing(text)
    if sp is not None:
        page["line_spacing"] = float(sp)
        resolution["line_spacing"] = SOURCE_LABEL

    ref = _extract_bib_style(text)
    ref_hang = None
    ref_num = None
    if ref and ref in REFERENCE_STYLE_KEYS:
        spec = get_spec(ref)
        ref_hang = spec.reference.hanging_indent_in
        ref_num = spec.reference.numbered
        # 引用维度的来源必须显式登记，否则合并时会按缺省兜底，
        # 在报告里被误标成「学校模板」（FR2/FR4 要求来源标注准确）。
        resolution["reference_style"] = SOURCE_LABEL
        resolution["reference_hanging_indent_in"] = SOURCE_LABEL
        resolution["reference_numbered"] = SOURCE_LABEL

    return TargetProfile(
        source=SOURCE_LABEL,
        spec_key=ref if ref in REFERENCE_STYLE_KEYS else None,
        page=page,
        reference_style=ref,
        reference_hanging_indent_in=ref_hang,
        reference_numbered=ref_num,
        resolution=resolution,
    )


def parse_latex_template(path: str) -> TargetProfile:
    """从 .cls / .sty / .tex 文件读取并解析。"""
    text = Path(path).read_text(encoding="utf-8", errors="ignore")
    return parse_latex_source(text)