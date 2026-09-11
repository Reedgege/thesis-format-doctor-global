"""格式问卷（Format Questionnaire）—— 四输入统一的「中间层」。

Thesis Format Doctor Global 的三种目标来源，全部收敛到同一个 TargetProfile：
  A) 选中的内置规范（spec）
  B) 用户手填的格式问卷（questionnaire / user）
  C) 用户把学校模板丢给任意 AI，让 AI 按本 schema 填出的 JSON（ai_import）
  D) 解析学校 Word 模板得到的结构（school / template）

设计原则（产品铁律）：
- 页面排版（页边距/字体/字号/行距/标题层级）**以学校/问卷为准**；
  规范里的 page 值仅作「兜底参考」，检查时不强制 fail（checker 对 spec 来源的页面维度降级为 info）。
- 参考文献格式**以规范为准**，除非学校/问卷显式给出引用字段。
- 多通道可组合：build_target() 统一合并并记录冲突，避免 GUI/CLI 各写一套。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional

from .specs import get_spec, REFERENCE_STYLE_KEYS, SPEC_ORDER
from .docx_reader import read_docx

# 来源标签映射（resolution 里缺省时按 TargetProfile.source 兜底；单一真相）
_SOURCE_ORIGIN = {
    "template": "school",
    "school": "school",
    "ai_import": "ai_import",
    "ai": "ai_import",
    "user": "user",
    "questionnaire": "user",
    "latex": "latex_template",
    "latex_template": "latex_template",
    "spec": "spec",
}

# 叠加顺序（FR2 明文优先级 template > ai > questionnaire > latex）：
# _overlay_one 是「后写覆盖先写」，故按优先级**从低到高**依次叠加。
_CHANNEL_ORDER = ("latex", "questionnaire", "ai", "template")


# ---------------------------------------------------------------------------
# 问卷 schema：驱动「手填 UI」与「AI 填表 JSON」的字段定义。
# 每个字段：key / label / group / type / options / default / unit / help
#          （可选）range=(min,max) —— 供 UI 做越界提示，避免 UI 侧再抄一份范围表
# group 用于 UI 分组与报告维度归类。本列表为字段唯一来源。
# ---------------------------------------------------------------------------

QUESTIONNAIRE_SCHEMA: list[dict] = [
    # —— 页面排版 ——
    {"key": "margin_top_in", "label": "上页边距", "group": "page", "type": "float",
     "default": 1.0, "unit": "英寸", "range": (0.2, 3.0),
     "help": "顶部页边距，单位英寸（1 英寸≈2.54 厘米）"},
    {"key": "margin_bottom_in", "label": "下页边距", "group": "page", "type": "float",
     "default": 1.0, "unit": "英寸", "range": (0.2, 3.0), "help": "底部页边距"},
    {"key": "margin_left_in", "label": "左页边距", "group": "page", "type": "float",
     "default": 1.0, "unit": "英寸", "range": (0.2, 3.0),
     "help": "左侧页边距（装订侧常更大）"},
    {"key": "margin_right_in", "label": "右页边距", "group": "page", "type": "float",
     "default": 1.0, "unit": "英寸", "range": (0.2, 3.0), "help": "右侧页边距"},
    {"key": "font_family", "label": "正文字体", "group": "page", "type": "str",
     "default": "Times New Roman", "unit": "", "help": "如 Times New Roman / Arial / Calibri"},
    {"key": "font_size_pt", "label": "正文字号", "group": "page", "type": "float",
     "default": 12.0, "unit": "磅", "range": (6.0, 72.0), "help": "正文推荐 10–12 磅"},
    {"key": "line_spacing", "label": "行距", "group": "page", "type": "float",
     "default": 2.0, "unit": "倍", "range": (0.8, 4.0),
     "help": "2.0=双倍，1.5=一倍半，1.0=单倍"},
    {"key": "body_alignment", "label": "正文对齐", "group": "page", "type": "choice",
     "options": ["left", "justify", "center", "right"], "default": "left", "unit": "",
     "help": "left=左对齐（APA/MLA 等），justify=两端对齐（IEEE 等）"},
    {"key": "first_line_indent_in", "label": "正文首行缩进", "group": "page", "type": "float",
     "default": 0.5, "unit": "英寸", "range": (0.0, 2.0),
     "help": "常见 0.5 英寸；0 = 不缩进（IEEE）"},
    {"key": "space_after_pt", "label": "段后间距", "group": "page", "type": "float",
     "default": 0.0, "unit": "磅", "range": (0.0, 72.0),
     "help": "常见 0 = 段间不留空行，用首行缩进区分段落"},

    # —— 标题层级 ——
    {"key": "heading_levels", "label": "标题层级数", "group": "heading", "type": "int",
     "default": 0, "unit": "级", "range": (0, 9), "help": "0=不强制；APA 常 5 级，其余多为学校定"},

    # —— 参考文献 ——
    {"key": "reference_style", "label": "参考文献格式", "group": "reference", "type": "choice",
     "options": list(REFERENCE_STYLE_KEYS), "default": "APA",
     "unit": "", "help": "引用风格；与上方规范选择保持一致即可"},
    {"key": "reference_hanging_indent_in", "label": "参考文献悬挂缩进", "group": "reference",
     "type": "float", "default": 0.5, "unit": "英寸", "range": (0.0, 2.0),
     "help": "悬挂缩进，常见 0.5 英寸；IEEE 为 0"},
    {"key": "reference_numbered", "label": "是否顺序编码", "group": "reference", "type": "bool",
     "default": False, "unit": "", "help": "True=如 IEEE [1]；False=著者-出版年制"},
]

# 数值字段的合理范围（由 schema 派生，避免 UI 侧再抄一份）
FIELD_RANGES: dict = {f["key"]: f["range"] for f in QUESTIONNAIRE_SCHEMA if f.get("range")}


@dataclass
class TargetProfile:
    """检查器消费的唯一目标格式画像。

    resolution: 维度 -> 来源标签，供对比报告标注「以规范为准 / 以学校模板为准」。
    conflicts: 显式来源之间互相冲突的维度记录（由 build_target 填充）。
    """

    source: str = "spec"                 # spec / school / user / ai_import 的组合
    spec_key: Optional[str] = None
    page: dict = field(default_factory=dict)        # 仅含已设置的页面值
    heading_levels: Optional[int] = None
    reference_style: Optional[str] = None
    reference_hanging_indent_in: Optional[float] = None
    reference_numbered: Optional[bool] = None
    resolution: dict = field(default_factory=dict)  # 维度 -> 来源
    conflicts: list = field(default_factory=list)   # 显式来源间冲突：{"dim","entries":[(label,val)]}
    notes: list = field(default_factory=list)       # 提示（如「LaTeX 未识别到字段，已回落规范」）

    PAGE_KEYS = ("margin_top_in", "margin_bottom_in", "margin_left_in",
                 "margin_right_in", "font_family", "font_size_pt", "line_spacing",
                 "body_alignment", "first_line_indent_in", "space_after_pt")

    def to_dict(self) -> dict:
        return {
            "source": self.source,
            "spec_key": self.spec_key,
            "page": self.page,
            "heading_levels": self.heading_levels,
            "reference_style": self.reference_style,
            "reference_hanging_indent_in": self.reference_hanging_indent_in,
            "reference_numbered": self.reference_numbered,
            "resolution": self.resolution,
            "conflicts": self.conflicts,
            "notes": self.notes,
        }


def target_from_spec(key: str) -> TargetProfile:
    """仅用内置规范生成目标画像。

    页面值标 resolution="spec"（checker 会将其降级为 info，不强制 fail）；
    引用维度标 resolution="spec"（引用归规范，硬判）。
    """
    spec = get_spec(key)
    page = {k: v for k, v in spec.page.__dict__.items() if v is not None}
    ref = spec.reference
    resolution = {k: "spec" for k in page}
    if spec.heading_levels:
        resolution["heading_levels"] = "spec"
    resolution["reference_style"] = "spec"
    resolution["reference_hanging_indent_in"] = "spec"
    resolution["reference_numbered"] = "spec"
    return TargetProfile(
        source="spec",
        spec_key=key,
        page=page,
        heading_levels=spec.heading_levels or None,
        reference_style=ref.style_key,
        reference_hanging_indent_in=ref.hanging_indent_in,
        reference_numbered=ref.numbered,
        resolution=resolution,
    )


def _coerce(schema: dict, raw: dict) -> dict:
    """按 schema 把任意来源（手填/AI-JSON）的原始值做类型校正与裁剪。"""
    out: dict = {}
    valid = {f["key"]: f for f in schema}
    for k, f in valid.items():
        if k not in raw:
            continue
        v = raw[k]
        if v in (None, ""):
            continue
        t = f["type"]
        try:
            if t == "float":
                out[k] = float(v)
            elif t == "int":
                out[k] = int(v)
            elif t == "bool":
                out[k] = bool(v) if not isinstance(v, str) else v.strip().lower() in ("1", "true", "yes", "y")
            elif t == "choice":
                s = str(v).strip()
                opts = f.get("options") or []
                if opts:
                    # 大小写不敏感匹配，归一化到白名单原始写法；不在白名单则丢弃（不臆造值）
                    hit = next((o for o in opts if str(o).lower() == s.lower()), None)
                    if hit is None:
                        continue
                    out[k] = hit
                else:
                    out[k] = s
            else:
                out[k] = str(v).strip()
        except (ValueError, TypeError):
            continue
    return out


def target_from_dict(data: dict, source: str = "questionnaire") -> TargetProfile:
    """手填问卷 / AI 填表 JSON -> TargetProfile。

    修复点：
    - 兼容 AI 产出的 {"filled": {...}} / {"data": {...}} 等包装，不再静默丢弃；
    - 若解析结果为空（无任何有效字段），显式抛出 ValueError，而非返回空画像。
    - 仅给 reference_style 时，自动反填规范的悬挂缩进/编码方式，保持四通道结构一致。
    """
    raw = dict(data)
    # 解包常见 AI 包装
    for wrapper in ("filled", "data", "answers", "responses"):
        if wrapper in raw and isinstance(raw[wrapper], dict):
            raw = raw[wrapper]
            break
    # 兼容嵌套 page 结构
    if "page" in raw and isinstance(raw["page"], dict):
        for k, v in raw["page"].items():
            raw.setdefault(k, v)

    clean = _coerce(QUESTIONNAIRE_SCHEMA, raw)
    page = {k: clean[k] for k in TargetProfile.PAGE_KEYS if k in clean}
    heading = clean.get("heading_levels")
    ref_style = clean.get("reference_style")
    ref_hang = clean.get("reference_hanging_indent_in")
    ref_num = clean.get("reference_numbered")

    resolution = {k: source for k in page}
    if heading is not None:
        resolution["heading_levels"] = source
    if ref_style is not None:
        resolution["reference_style"] = source
    if ref_hang is not None:
        resolution["reference_hanging_indent_in"] = source
    if ref_num is not None:
        resolution["reference_numbered"] = source

    # 仅给 reference_style 时，反填规范引用字段，保证结构完整
    if ref_style in REFERENCE_STYLE_KEYS:
        sp = get_spec(ref_style)
        if ref_hang is None:
            ref_hang = sp.reference.hanging_indent_in
            resolution["reference_hanging_indent_in"] = "spec"
        if ref_num is None:
            ref_num = sp.reference.numbered
            resolution["reference_numbered"] = "spec"

    # 空校验：避免静默得到空画像
    if not page and ref_style is None and heading is None:
        raise ValueError(
            "未从输入解析出任何有效格式字段（页边距/字体/参考文献风格/标题层级均为空）。"
            "请检查 JSON 是否为合法格式，或是否包含 'filled' 等包装字段。"
        )

    return TargetProfile(
        source=source,
        spec_key=ref_style if ref_style in REFERENCE_STYLE_KEYS else None,
        page=page,
        heading_levels=heading,
        reference_style=ref_style,
        reference_hanging_indent_in=ref_hang,
        reference_numbered=ref_num,
        resolution=resolution,
    )


def target_from_ai_json(path: str) -> TargetProfile:
    """导入 AI 填好的 JSON（用户把学校模板交给任意 AI 产出的表单）。"""
    import json
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return target_from_dict(data, source="ai_import")


def target_from_template_docx(doc_profile, spec_key: str) -> TargetProfile:
    """把学校模板 docx 的实测画像转成 school 目标（页面排版以学校为准）。

    修复点：学校模板通道**不设置任何 reference_* 字段**（引用归规范），
    避免把规范的引用值错标成「学校模板」。
    引用维度由 spec 目标在合并时提供（硬判）。
    """
    page: dict = {}
    if doc_profile.margins:
        for k_src, k_dst in (("top", "margin_top_in"), ("bottom", "margin_bottom_in"),
                             ("left", "margin_left_in"), ("right", "margin_right_in")):
            v = doc_profile.margins.get(k_src)
            if v is not None:
                page[k_dst] = round(float(v), 3)
    if doc_profile.font_family:
        page["font_family"] = doc_profile.font_family
    if doc_profile.font_size_pt:
        page["font_size_pt"] = round(float(doc_profile.font_size_pt), 1)
    if doc_profile.line_spacing:
        page["line_spacing"] = round(float(doc_profile.line_spacing), 2)
    if doc_profile.body_alignment:
        page["body_alignment"] = doc_profile.body_alignment
    if doc_profile.first_line_indent_in is not None:
        page["first_line_indent_in"] = round(float(doc_profile.first_line_indent_in), 3)
    if doc_profile.space_after_pt is not None:
        page["space_after_pt"] = round(float(doc_profile.space_after_pt), 1)
    if doc_profile.margins_inconsistent:
        # 多分节页边距不一致：页面维度不强行以模板为准，留给用户注意
        pass

    resolution = {k: "school" for k in page}
    return TargetProfile(
        source="template",
        spec_key=spec_key if spec_key in REFERENCE_STYLE_KEYS else None,
        page=page,
        heading_levels=doc_profile.heading_levels_used or None,
        reference_style=None,
        reference_hanging_indent_in=None,
        reference_numbered=None,
        resolution=resolution,
    )


def _overlay_one(merged: TargetProfile, extra: TargetProfile, conflicts: list) -> TargetProfile:
    """把一个显式来源叠加到 merged 上，并登记来源间冲突。

    来源标签优先取 extra.resolution 里的显式标注；缺失时按 extra.source 映射
    （**不再硬编码 "school"** —— LaTeX 来源曾被误标成「学校模板」）。
    """
    fallback = _SOURCE_ORIGIN.get(extra.source, extra.source)

    def _origin(key: str) -> str:
        return extra.resolution.get(key) or fallback

    def _register(key, prev_val, prev_origin, new_val, new_origin):
        if (prev_val is not None and prev_origin not in (None, "spec")
                and prev_val != new_val):
            conflicts.append({
                "dim": key,
                "entries": [
                    (prev_origin, prev_val),
                    (new_origin, new_val),
                ],
            })

    for k, v in extra.page.items():
        _register(k, merged.page.get(k), merged.resolution.get(k), v, _origin(k))
        merged.page[k] = v
        merged.resolution[k] = _origin(k)

    if extra.heading_levels is not None:
        _register("heading_levels", merged.heading_levels,
                  merged.resolution.get("heading_levels"),
                  extra.heading_levels, _origin("heading_levels"))
        merged.heading_levels = extra.heading_levels
        merged.resolution["heading_levels"] = _origin("heading_levels")

    if extra.reference_style is not None:
        _register("reference_style", merged.reference_style,
                  merged.resolution.get("reference_style"),
                  extra.reference_style, _origin("reference_style"))
        merged.reference_style = extra.reference_style
        merged.resolution["reference_style"] = _origin("reference_style")
    if extra.reference_hanging_indent_in is not None:
        _register("reference_hanging_indent_in", merged.reference_hanging_indent_in,
                  merged.resolution.get("reference_hanging_indent_in"),
                  extra.reference_hanging_indent_in, _origin("reference_hanging_indent_in"))
        merged.reference_hanging_indent_in = extra.reference_hanging_indent_in
        merged.resolution["reference_hanging_indent_in"] = _origin("reference_hanging_indent_in")
    if extra.reference_numbered is not None:
        _register("reference_numbered", merged.reference_numbered,
                  merged.resolution.get("reference_numbered"),
                  extra.reference_numbered, _origin("reference_numbered"))
        merged.reference_numbered = extra.reference_numbered
        merged.resolution["reference_numbered"] = _origin("reference_numbered")
    return merged


def build_target(spec_key: str, template: Optional[str] = None,
                 ai: Optional[str] = None,
                 questionnaire: Optional[dict] = None,
                 latex: Optional[str] = None) -> TargetProfile:
    """五通道统一合并入口（FR2：任选其一或组合）。

    - 规范为基线（页面软参考 / 引用硬判）；
    - 显式来源优先级 **template(学校docx) > ai > questionnaire > latex**（FR2 明文），
      实现上按「从低到高依次叠加」达到高优先级覆盖低优先级；
    - 引用仅当显式来源提供才覆盖规范；
    - 来源间冲突（含引用维度）记录在 conflicts 中供报告高亮；

    template: 学校 docx 模板路径
    ai: AI 填表 JSON 路径
    questionnaire: 用户手填问卷 dict
    latex: LaTeX 模板路径（.cls / .sty / .tex）
    """
    base = target_from_spec(spec_key)
    built: dict = {}
    if template:
        built["template"] = ("school", target_from_template_docx(read_docx(template), spec_key))
    if ai:
        built["ai"] = ("ai_import", target_from_ai_json(ai))
    if questionnaire:
        built["questionnaire"] = ("user", target_from_dict(questionnaire, "user"))
    if latex:
        from .latex_template import parse_latex_template  # 局部导入避免循环
        latex_target = parse_latex_template(latex)
        built["latex"] = ("latex_template", latex_target)

    merged = base
    conflicts: list = []
    notes: list = []
    applied_labels: list = []
    for channel in _CHANNEL_ORDER:              # 从低优先级到高优先级依次叠加
        item = built.get(channel)
        if item is None:
            continue
        t = item[1]
        merged = _overlay_one(merged, t, conflicts)
        applied_labels.append(_SOURCE_ORIGIN.get(t.source, t.source))
        if channel == "latex" and not t.page and not t.reference_style:
            notes.append(
                "已读取 LaTeX 模板，但未能从中识别出格式字段（可能用了自定义宏或非常规模板）；"
                "相关维度已回落到所选规范。"
            )

    merged.source = "+".join(["spec"] + applied_labels)
    merged.conflicts = conflicts
    merged.notes = notes
    return merged


def merge_school(spec_target: TargetProfile, school: TargetProfile) -> TargetProfile:
    """向后兼容的二元合并（页面以学校为准，引用以规范为准）。

    保留 school 自带的来源标签（不再统一覆盖成 "school"）。
    """
    conflicts: list = []
    merged = _overlay_one(spec_target, school, conflicts)
    merged.conflicts = conflicts
    merged.source = "spec+school"
    return merged
