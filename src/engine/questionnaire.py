"""格式问卷（Format Questionnaire）—— 四输入统一的「中间层」。

PaperFormat Pro 的三种目标来源，全部收敛到同一个 TargetProfile：
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
    "spec": "spec",
}

# 叠加顺序（FR2 明文优先级 template > ai > questionnaire）：
# _overlay_one 是「后写覆盖先写」，故按优先级**从低到高**依次叠加。
_CHANNEL_ORDER = ("questionnaire", "ai", "template")


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
    # 页眉距页顶：多数格式不规定，故默认留空（填写后才参与检查/修正，不臆造）。
    # MLA 9th 要求 ½″（0.5）；学校模板能读到时以模板为准。
    {"key": "header_distance_in", "label": "页眉距页顶", "group": "page", "type": "float",
     "default": "", "unit": "英寸", "range": (0.0, 3.0),
     "help": "页眉上边缘到页面顶部的距离；多数格式不规定，留空即不检查（MLA 9 = 0.5）"},
    # 栏数：单栏=1 / 双栏=2。结构性页面维度，不进 PAGE_KEYS；复核弹窗用单选下拉。
    # 标签双语：zh 走此处的 label，en 走 i18n.Q_LABELS_EN（英文界面绝不泄漏中文）。
    {"key": "columns", "label": "栏数", "group": "page", "type": "int",
     "default": 1, "unit": "", "range": (1, 4),
     "help": "1=单栏（多数格式）；2=双栏（IEEE 会议/期刊）"},

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
    # 分节栏数（单栏=1 / 双栏=2）。None = 不检查（Other/自定义且用户未声明）。
    # 属于**结构性页面维度**，与 page 里的标量分开：它不进 PAGE_KEYS（复核弹窗里用
    # 单选下拉而非文本输入），checker 也不对它做「spec→info」软参考降级。
    columns: Optional[int] = None
    # 每级标题格式样本（学校模板驱动 / 问卷 / AI 显式给出）：level(int) -> dict
    # {alignment, bold, italic, indent, size_pt, number_prefix, uppercase}。
    # 仅当来源显式提供时才填入；其余入口留空、不臆造。
    heading_formats: dict = field(default_factory=dict)
    # 是否允许把参考文献**内容**改写为规范引文格式（best-effort，如 APA→IEEE）。
    # 默认 False：产品铁律是「只改格式、不改内容」，参考文献默认逐字不动（仅加 [n]
    # 前缀 + 悬挂缩进）。只有走 target_from_spec() 的真实产品路径、且该规范显式
    # 声明支持时才为 True —— 保证纯格式承诺在其它入口（含单测/单参调用）恒成立。
    reference_rewrite: bool = False
    resolution: dict = field(default_factory=dict)  # 维度 -> 来源
    conflicts: list = field(default_factory=list)   # 显式来源间冲突：{"dim","entries":[(label,val)]}
    notes: list = field(default_factory=list)       # 提示（如「学校模板未识别到字段，已回落规范」）

    PAGE_KEYS = ("margin_top_in", "margin_bottom_in", "margin_left_in",
                 "margin_right_in", "font_family", "font_size_pt", "line_spacing",
                 "body_alignment", "first_line_indent_in", "space_after_pt",
                 "header_distance_in")

    def to_dict(self) -> dict:
        return {
            "source": self.source,
            "spec_key": self.spec_key,
            "page": self.page,
            "heading_levels": self.heading_levels,
            "reference_style": self.reference_style,
            "reference_hanging_indent_in": self.reference_hanging_indent_in,
            "reference_numbered": self.reference_numbered,
            "columns": self.columns,
            "heading_formats": {str(k): v for k, v in self.heading_formats.items()},
            "reference_rewrite": self.reference_rewrite,
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
    if spec.columns is not None:
        resolution["columns"] = "spec"
    return TargetProfile(
        source="spec",
        spec_key=key,
        page=page,
        heading_levels=spec.heading_levels or None,
        columns=spec.columns,
        reference_style=ref.style_key,
        reference_hanging_indent_in=ref.hanging_indent_in,
        reference_numbered=ref.numbered,
        # 只有「顺序编码制 + IEEE」才允许 best-effort 引文内容改写；APA/MLA 等
        # 著者-出版年制改写风险大于收益，保持「只改格式」承诺。
        reference_rewrite=(ref.style_key == "IEEE" and bool(ref.numbered)),
        resolution=resolution,
        # 把所选规范的「期望说明」（含学生/专业分层的前置材料差异）带入画像，
        # 供报告展示「正确说明」。空说明不塞，避免报告出现空行。
        notes=[spec.notes] if spec.notes else [],
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
    cols = clean.get("columns")            # 栏数（int；不在 PAGE_KEYS 里，单独承载）
    ref_style = clean.get("reference_style")
    ref_hang = clean.get("reference_hanging_indent_in")
    ref_num = clean.get("reference_numbered")

    # 每级标题格式（问卷 / AI 显式给出时纳入；否则留空，不臆造）
    hf = raw.get("heading_formats") or raw.get("headingLevels")
    heading_formats: dict = {}
    if isinstance(hf, dict):
        for _k, _v in hf.items():
            try:
                _lvl = int(_k)
            except (ValueError, TypeError):
                continue
            if isinstance(_v, dict):
                heading_formats[_lvl] = {
                    _kk: _vv for _kk, _vv in _v.items()
                    if _kk in ("alignment", "bold", "italic", "indent",
                               "size_pt", "number_prefix", "uppercase")
                }

    resolution = {k: source for k in page}
    if heading is not None:
        resolution["heading_levels"] = source
    if cols is not None:
        resolution["columns"] = source
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
    if not page and ref_style is None and heading is None and cols is None:
        raise ValueError(
            "No valid format fields could be parsed from the input "
            "(margins, font, reference style, and heading levels are all empty). "
            "Check that the YAML/JSON is well-formed, or that it contains wrapper fields such as 'filled'."
        )

    if heading_formats:
        resolution["heading_formats"] = source
    return TargetProfile(
        source=source,
        spec_key=ref_style if ref_style in REFERENCE_STYLE_KEYS else None,
        page=page,
        heading_levels=heading,
        columns=cols,
        heading_formats=heading_formats,
        reference_style=ref_style,
        reference_hanging_indent_in=ref_hang,
        reference_numbered=ref_num,
        resolution=resolution,
    )


def target_from_ai(path: str) -> TargetProfile:
    """导入 AI 填好的配置文件（YAML 或旧 JSON）。

    两种结构都支持：① 顶层扁平键值；② 包在 filled/data 等对象里。
    文件可以是 .yaml/.yml（推荐）或旧版 .json；按内容自动识别，用户无需关心扩展名。
    """
    import json
    try:
        from yaml import safe_load
    except ImportError:  # pragma: no cover - pyyaml 应随包分发
        safe_load = None
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()
    if not text.strip():
        raise ValueError("The AI config file is empty.")
    # JSON 是 YAML 的子集；先按 JSON 试（报错信息更精准），失败再回落 YAML。
    try:
        data = json.loads(text)
    except (ValueError, json.JSONDecodeError):
        if safe_load is None:
            raise ValueError("PyYAML is not installed and the YAML template cannot be parsed; use JSON instead, or install pyyaml.")
        data = safe_load(text)
    if not isinstance(data, dict):
        raise ValueError("The top level of the AI config file must be an object (key-value pairs); got: %s" % type(data).__name__)
    return target_from_dict(data, source="ai_import")


def target_from_template_docx(doc_profile, reference_style: Optional[str] = None) -> TargetProfile:
    """把学校模板 docx 的实测画像转成 school 目标（页面排版以学校为准）。

    无规范基线模式（产品铁律）：spec_key 恒为 None，不套任何内置规范的页面值，
    避免把规范值错标成「学校模板」、或抢占用户在高级面板里显式选的引用风格。
    引用风格由「学校模板」子面板的下拉框驱动（reference_style 参数）：
      - 留空（None）→ 不检查参考文献（也不带任何风格专属开关）；
      - 选 APA/MLA/IEEE → 自动带入该风格的引用开关（悬挂缩进 / 编码方式，以及
        running head / 页码位置 / 参考文献另起页 / 字体白名单 / 标题页 / 标题加粗规则，
        这些开关由 checker 经 resolve_spec(target) 按 reference_style 取出，无需在此落盘）。
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
    # 页眉距页顶（MLA 9th 要求 ½″）：学校模板读出则硬判，读不到则留空（不臆造）
    if doc_profile.header_distance_in is not None:
        page["header_distance_in"] = round(float(doc_profile.header_distance_in), 3)
    if doc_profile.margins_inconsistent:
        # 多分节页边距不一致：页面维度不强行以模板为准，留给用户注意
        pass

    resolution = {k: "school" for k in page}

    # 栏数：学校模板实测值（多分节取众数；读不到则为 None，不臆造）——学校为准，硬判
    tpl_columns = getattr(doc_profile, "columns", None)
    if tpl_columns is not None:
        resolution["columns"] = "school"

    # 每级标题格式样本（学校模板驱动）：层级 → dict
    heading_formats: dict = {}
    for lvl, sample in (doc_profile.heading_formats or {}).items():
        heading_formats[lvl] = sample.to_dict() if hasattr(sample, "to_dict") else dict(sample)

    ref_style = reference_style if reference_style in REFERENCE_STYLE_KEYS else None
    ref_hang = None
    ref_num = None
    if ref_style is not None:
        sp = get_spec(ref_style)
        ref_hang = sp.reference.hanging_indent_in
        ref_num = sp.reference.numbered
        resolution["reference_style"] = "school"
        resolution["reference_hanging_indent_in"] = "school"
        resolution["reference_numbered"] = "school"

    return TargetProfile(
        source="template",
        spec_key=None,
        page=page,
        heading_levels=doc_profile.heading_levels_used or None,
        columns=tpl_columns,
        heading_formats=heading_formats,
        reference_style=ref_style,
        reference_hanging_indent_in=ref_hang,
        reference_numbered=ref_num,
        resolution=resolution,
    )


def _overlay_one(merged: TargetProfile, extra: TargetProfile, conflicts: list) -> TargetProfile:
    """把一个显式来源叠加到 merged 上，并登记来源间冲突。

    来源标签优先取 extra.resolution 里的显式标注；缺失时按 extra.source 经
    _SOURCE_ORIGIN 映射（不再硬编码 "school"）。
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

    if extra.columns is not None:
        _register("columns", merged.columns, merged.resolution.get("columns"),
                  extra.columns, _origin("columns"))
        merged.columns = extra.columns
        merged.resolution["columns"] = _origin("columns")

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
    if extra.heading_formats:
        merged.heading_formats = dict(extra.heading_formats)
        merged.resolution["heading_formats"] = _origin("heading_formats")
    return merged


def build_target(spec_key: Optional[str] = None, template: Optional[str] = None,
                 ai: Optional[str] = None,
                 questionnaire: Optional[dict] = None,
                 reference_style: Optional[str] = None) -> TargetProfile:
    """四通道统一合并入口（FR2：任选其一或组合）。

    两种基线模式：
    - 规范基线：spec_key 为真实规范（APA-Student / APA-Professional / MLA / IEEE，
      含别名 "APA"）时，以 target_from_spec 为基线（页面软参考 / 引用硬判）。
    - 无规范基线：spec_key 为 None 或 "Other"（学校模板 / 自定义来源）时，
      **不套任何内置规范页面值**，页面以学校模板 / 问卷 / AI 为准；引用由
      reference_style 参数（学校模板子下拉）或问卷 / AI 数据驱动，spec_key 恒为 None。

    template: 学校 docx 模板路径
    ai: AI 填表 JSON 路径
    questionnaire: 用户手填问卷 dict
    reference_style: 学校模板子下拉选中的引用风格（仅无规范基线模式生效；
                    留空表示「不检查参考文献」）
    """
    use_spec = spec_key is not None and spec_key != "Other"
    if use_spec:
        base = target_from_spec(spec_key)
        ref_for_template = None  # 规范已权威，学校模板子下拉不再覆盖引用风格
    else:
        base = TargetProfile(source="")  # 空基线：页面 / 引用全由显式来源提供
        ref_for_template = reference_style if reference_style in REFERENCE_STYLE_KEYS else None

    built: dict = {}
    if template:
        built["template"] = ("school", target_from_template_docx(read_docx(template), ref_for_template))
    if ai:
        built["ai"] = ("ai_import", target_from_ai(ai))
    if questionnaire:
        built["questionnaire"] = ("user", target_from_dict(questionnaire, "user"))

    merged = base
    conflicts: list = []
    notes: list = list(base.notes)  # 继承规范自带的「期望说明」（学生/专业分层等）
    applied_labels: list = []
    for channel in _CHANNEL_ORDER:              # 从低优先级到高优先级依次叠加
        item = built.get(channel)
        if item is None:
            continue
        t = item[1]
        merged = _overlay_one(merged, t, conflicts)
        applied_labels.append(_SOURCE_ORIGIN.get(t.source, t.source))

    if use_spec:
        merged.source = "+".join(["spec"] + applied_labels)
    else:
        merged.source = "+".join([s for s in applied_labels] or ["custom"])
    merged.conflicts = conflicts
    merged.notes = notes
    return merged


# 页面可数值化的字段（录入弹窗里按 float 解析；其余按字符串原样写回）
PAGE_FLOAT_KEYS = {
    "margin_top_in", "margin_bottom_in", "margin_left_in", "margin_right_in",
    "font_size_pt", "line_spacing", "first_line_indent_in",
    "space_after_pt", "header_distance_in",
}


def apply_page_overrides(target: "TargetProfile", edits: dict) -> None:
    """复核弹窗：仅把用户「真正改动过」的 Page Layout 字段写回 target。

    edits: {page_key: raw_string}. 仅当 raw 与 target.page 当前值**不同**才写回，
    并把 resolution 标 "user"（显式覆盖）。未改动 / 仅回显的字段保持原 resolution，
    避免把规范的软参考（spec 来源 → checker 降级为 info）误升级为硬规则（warn/fail）。

    纯函数、不依赖 GUI，便于 headless 单测。
    """
    for key, raw in edits.items():
        # 栏数是结构性维度（不在 PAGE_KEYS 里），单独处理：int 归一、仅改动才升级。
        if key == "columns":
            raw = (raw or "").strip()
            if raw == "":
                continue
            try:
                iv = int(raw)
            except (ValueError, TypeError):
                continue
            if iv < 1:
                continue
            if iv == target.columns:
                continue  # 与提取/规范值一致 → 不覆盖、不升级
            target.columns = iv
            target.resolution["columns"] = "user"
            continue
        if key not in TargetProfile.PAGE_KEYS:
            continue
        raw = (raw or "").strip()
        if raw == "":
            continue
        orig = target.page.get(key)
        if raw == ("" if orig is None else str(orig)):
            continue  # 与提取/规范值一致 → 不覆盖、不升级
        if key in PAGE_FLOAT_KEYS:
            try:
                fv = float(raw)
            except (ValueError, TypeError):
                continue
            target.page[key] = fv
        else:
            target.page[key] = raw
        target.resolution[key] = "user"


def merge_school(spec_target: TargetProfile, school: TargetProfile) -> TargetProfile:
    """向后兼容的二元合并（页面以学校为准，引用以规范为准）。

    保留 school 自带的来源标签（不再统一覆盖成 "school"）。
    """
    conflicts: list = []
    merged = _overlay_one(spec_target, school, conflicts)
    merged.conflicts = conflicts
    merged.source = "spec+school"
    return merged
