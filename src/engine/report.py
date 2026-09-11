"""对比裁决报告：把检查结果整理成「以谁为准」的可读报告。

核心价值：把 spec / 学校模板 / 问卷 / AI 四类来源在「每个维度」上的裁决显式列出，
让用户一眼看清哪些按通用规范、哪些按学校要求，冲突项高亮。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from .checker import ALIGN_LABEL, Finding, summarize
from .docx_reader import DocumentProfile
from .questionnaire import TargetProfile
from .specs import get_spec

ORIGIN_LABEL = {
    "spec": "通用规范",
    "school": "学校模板",
    "template": "学校模板",
    "user": "格式问卷",
    "ai_import": "AI 填表",
    "ai": "AI 填表",
    "latex_template": "LaTeX 模板",
    "latex": "LaTeX 模板",
}


@dataclass
class Report:
    meta: dict = field(default_factory=dict)
    resolution: list = field(default_factory=list)   # (维度, 来源标签)
    conflicts: list = field(default_factory=list)     # (维度, 左描述, 右描述)
    findings: list = field(default_factory=list)
    summary: dict = field(default_factory=dict)
    markdown: str = ""


_DIM_LABEL = {
    "margin_top_in": "上页边距", "margin_bottom_in": "下页边距",
    "margin_left_in": "左页边距", "margin_right_in": "右页边距",
    "font_family": "正文字体", "font_size_pt": "正文字号",
    "line_spacing": "行距", "heading_levels": "标题层级",
    "body_alignment": "正文对齐", "first_line_indent_in": "正文首行缩进",
    "space_after_pt": "段后间距",
    "reference_style": "参考文献格式", "reference_hanging_indent_in": "参考文献悬挂缩进",
    "reference_numbered": "参考文献编码方式",
}

_ALIGN_LABEL = ALIGN_LABEL  # 对齐标签单一来源在 checker


def _fmt_val(key: str, v) -> str:
    if v is None:
        return "—"
    if key in ("margin_top_in", "margin_bottom_in", "margin_left_in",
               "margin_right_in", "reference_hanging_indent_in",
               "first_line_indent_in"):
        return f'{v}"'
    if key == "font_size_pt":
        return f"{v}pt"
    if key == "line_spacing":
        return f"{v}x"
    if key == "space_after_pt":
        return f"{v}pt"
    if key == "body_alignment":
        return _ALIGN_LABEL.get(str(v), str(v))
    if key == "reference_numbered":
        return "顺序编码" if v else "著者-出版年"
    if key == "reference_style":
        return str(v)
    return str(v)


def _source_label(origin: str) -> str:
    return ORIGIN_LABEL.get(origin, origin)


def _friendly_source(target: TargetProfile) -> str:
    origins = sorted({v for v in target.resolution.values() if v})
    if not origins:
        return "未设置"
    return " + ".join(_source_label(o) for o in origins)


def _collect_conflicts(target: TargetProfile) -> list:
    """规范默认 vs 显式来源 + 显式来源之间 的冲突。"""
    out: list = []
    spec_key = target.spec_key
    spec = get_spec(spec_key) if spec_key and spec_key != "Other" else None

    # 1) 规范默认 vs 显式来源
    if spec:
        spec_page = spec.page.__dict__
        spec_ref = spec.reference
        for key, origin in target.resolution.items():
            if origin in ("spec",):
                continue
            if key in TargetProfile.PAGE_KEYS:
                tval = target.page.get(key)
                sval = spec_page.get(key)
            elif key == "heading_levels":
                tval = target.heading_levels
                sval = spec.heading_levels
            elif key == "reference_style":
                tval = target.reference_style
                sval = spec_ref.style_key
            elif key == "reference_hanging_indent_in":
                tval = target.reference_hanging_indent_in
                sval = spec_ref.hanging_indent_in
            elif key == "reference_numbered":
                tval = target.reference_numbered
                sval = spec_ref.numbered
            else:
                continue
            if tval is not None and sval is not None and tval != sval:
                out.append((_DIM_LABEL.get(key, key),
                            f"通用规范 {_fmt_val(key, sval)}",
                            f"{_source_label(origin)} {_fmt_val(key, tval)}"))

    # 2) 显式来源之间
    for c in target.conflicts:
        dim = _DIM_LABEL.get(c["dim"], c["dim"])
        chain = " → ".join(
            f"{_source_label(lbl)} {_fmt_val(c['dim'], val)}" for lbl, val in c["entries"]
        )
        out.append((dim, "来源冲突", chain))

    return out


def build_report(docx_path: str, target: TargetProfile,
                 prof: DocumentProfile, findings: list[Finding]) -> Report:
    rep = Report()
    rep.meta = {
        "docx": docx_path,
        "source": _friendly_source(target),
        "spec_key": target.spec_key or "Other",
    }

    # —— 来源裁决表 ——
    for dim, origin in target.resolution.items():
        rep.resolution.append((_DIM_LABEL.get(dim, dim), _source_label(origin)))

    # —— 冲突 ——
    rep.conflicts = _collect_conflicts(target)

    rep.findings = findings
    rep.summary = summarize(findings)

    # —— Markdown ——
    lines: list = []
    lines.append("# Thesis Format Doctor Global · 格式体检报告")
    lines.append("")
    lines.append(f"- 文档：`{docx_path}`")
    lines.append(f"- 规范：{rep.meta['spec_key']}　来源：{rep.meta['source']}")
    s = rep.summary
    lines.append(f"- 结论：通过 {s.get('pass',0)} · 提示 {s.get('info',0)} · 警告 {s.get('warn',0)} · 不达标 {s.get('fail',0)}")
    lines.append("")

    # 章节号**动态分配**：只有实际渲染的段落才占号，避免出现「一、」直接跳「三、」。
    _CN = "一二三四五六七八九十"
    _sec = [0]

    def _sec_no() -> str:
        _sec[0] += 1
        n = _sec[0]
        return _CN[n - 1] if n <= len(_CN) else str(n)

    lines.append(f"## {_sec_no()}、维度裁决（以谁为准）")
    lines.append("")
    if rep.resolution:
        # 按来源分组：模板存在时来源会分散，分组后比逐行罗列更易读。
        groups: dict = {}
        for dim, origin in rep.resolution:
            groups.setdefault(origin, []).append(dim)
        for origin, dims in groups.items():
            lines.append(f"- **{origin}**（{len(dims)} 项）：{'、'.join(dims)}")
    else:
        lines.append("- 未设置具体目标值（请上传学校模板或填写问卷）")
    lines.append("")
    lines.append("> 说明：标注「通用规范」的页面维度（页边距/字体/行距等）仅作参考提示，"
                 "不强制判不达标；页面排版以学校模板/问卷为准。参考文献维度始终以规范为准。")
    lines.append("")

    if getattr(target, "notes", None):
        lines.append(f"## {_sec_no()}、提示")
        lines.append("")
        for n in target.notes:
            lines.append(f"- ℹ️ {n}")
        lines.append("")

    if rep.conflicts:
        lines.append(f"## {_sec_no()}、规范 vs 学校模板 / 来源间 冲突项")
        lines.append("")
        lines.append("> 以下维度不同来源要求不一致，已**以具体来源（学校/问卷/AI）为准**检查：")
        lines.append("")
        for dim, left, right in rep.conflicts:
            lines.append(f"- **{dim}**：{left} → {right}")
        lines.append("")

    lines.append(f"## {_sec_no()}、详细检查项")
    lines.append("")
    order = {"fail": 0, "warn": 1, "info": 2, "pass": 3}
    for f in sorted(findings, key=lambda x: order.get(x.severity, 9)):
        icon = {"pass": "✅", "warn": "⚠️", "fail": "❌", "info": "ℹ️"}.get(f.severity, "•")
        line = f"- {icon} **[{f.dimension}]** {f.message}"
        if f.expected or f.actual:
            line += f"　（期望：{f.expected or '—'} ｜ 实测：{f.actual or '—'}）"
        lines.append(line)
    lines.append("")

    rep.markdown = "\n".join(lines)
    return rep
