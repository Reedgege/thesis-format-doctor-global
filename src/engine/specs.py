"""内置论文格式规范（离线、硬编码、零网络、零 LLM）。

Thesis Format Doctor Global —— 海外版规则引擎的规范底本。

设计铁律（模板驱动原则）：
- 这些规范只定义「该风格自身的文稿格式指引」+「参考文献格式规则」。
- 页面排版（页边距 / 字体 / 字号 / 行距）**由学校模板或格式问卷决定**，
  规范里的 page 默认值仅作为「该风格典型文稿」的兜底参考（检查时不强制 fail）。
- 参考文献格式规则以所选引用风格为准，除非学校显式覆盖。

本文件不依赖任何外部网络或大模型，全部写死，保证离线安全。
"""

from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import Optional


@dataclass
class PageLayout:
    """页面排版默认值（英寸 / 磅）。值为 None 表示该风格不强制。

    注意（模板驱动铁律）：这些都是「该风格典型文稿」的兜底参考值，
    检查时若来源为 spec 会降级为 info（不强制 fail）；只有学校模板 / 问卷 / AI
    显式给出时才硬判。
    """

    margin_top_in: Optional[float] = 1.0
    margin_bottom_in: Optional[float] = 1.0
    margin_left_in: Optional[float] = 1.0
    margin_right_in: Optional[float] = 1.0
    font_family: Optional[str] = "Times New Roman"
    font_size_pt: Optional[float] = 12.0
    line_spacing: Optional[float] = 2.0  # 2.0 = 双倍行距
    # —— 正文段落级排版（段落属性，非页面属性）——
    # 默认 None = 不强制；具体各风格典型值见下方 _BODY_LAYOUT 表（单一真相）。
    body_alignment: Optional[str] = None            # left / justify / center / right
    first_line_indent_in: Optional[float] = None    # 正文首行缩进（英寸）；0 = 不缩进
    space_after_pt: Optional[float] = None          # 段后间距（磅）


@dataclass
class ReferenceRule:
    """参考文献 / 引用格式规则。

    checks: 确定性检查标识元组，供 checker 数据驱动地分发引用检查。
    可选值见 checker._REF_CHECKS。
    """

    style_key: str                 # APA / MLA / Chicago / IEEE / Harvard / Other
    hanging_indent_in: Optional[float] = 0.5
    numbered: bool = False         # True=顺序编码制（如 IEEE），False=著者-出版年制
    order_hint: str = ""           # 人类可读的顺序说明
    pattern_hint: str = ""         # 给用户的示例
    checks: tuple = ()             # 确定性检查标识，供 checker 使用


@dataclass
class StyleSpec:
    """一个完整的引用风格规范。"""

    key: str
    name: str
    page: PageLayout
    reference: ReferenceRule
    heading_levels: int = 0       # 0 = 该风格不强制标题层级
    notes: str = ""

    def to_dict(self) -> dict:
        return {
            "key": self.key,
            "name": self.name,
            "page": asdict(self.page),
            "reference": asdict(self.reference),
            "heading_levels": self.heading_levels,
            "notes": self.notes,
        }


# ---------------------------------------------------------------------------
# 五个内置规范 + Other 兜底。数值均来自各风格官方手册的典型文稿要求。
# 注意：IEEE 会议模板常为 2 栏 10pt，学位论文多为 12pt/1.5 倍，这里取论文常用值，
#       且仅作为「参考」而非对页面排版的硬性要求（页面以学校模板/问卷为准）。
# ---------------------------------------------------------------------------

SPECS: dict[str, StyleSpec] = {
    "APA": StyleSpec(
        key="APA",
        name="APA (7th ed.) 美国心理学会",
        page=PageLayout(1.0, 1.0, 1.0, 1.0, "Times New Roman", 12.0, 2.0),
        reference=ReferenceRule(
            style_key="APA",
            hanging_indent_in=0.5,
            numbered=False,
            order_hint="著者-出版年：Author, A. A. (Year). Title. Publisher.",
            pattern_hint="Smith, J. A. (2020). Title of work. Publisher.",
            checks=("has_year_paren", "author_year_order"),
        ),
        heading_levels=5,
        notes="学生论文含标题页；标题层级最多 5 级；参考文献悬挂缩进 0.5 英寸。",
    ),
    "MLA": StyleSpec(
        key="MLA",
        name="MLA (9th ed.) 现代语言协会",
        page=PageLayout(1.0, 1.0, 1.0, 1.0, "Times New Roman", 12.0, 2.0),
        reference=ReferenceRule(
            style_key="MLA",
            hanging_indent_in=0.5,
            numbered=False,
            order_hint="著者-题名-出版项：Author. Title. Publisher, Year.",
            pattern_hint="Smith, John. Title of Work. Publisher, 2020.",
            checks=("author_title_order",),
        ),
        heading_levels=0,
        notes="通常无独立标题页；Works Cited 悬挂缩进 0.5 英寸；不强制编号标题。",
    ),
    "Chicago": StyleSpec(
        key="Chicago",
        name="Chicago (17th ed.) 注脚-书目制",
        page=PageLayout(1.0, 1.0, 1.0, 1.0, "Times New Roman", 12.0, 2.0),
        reference=ReferenceRule(
            style_key="Chicago",
            hanging_indent_in=0.5,
            numbered=False,
            order_hint="书目：Author. Title. Place: Publisher, Year.",
            pattern_hint="Smith, John. Title of Work. Chicago: Publisher, 2020.",
            checks=("author_title_place_order",),
        ),
        heading_levels=0,
        notes="注脚-书目制；Bibliography 悬挂缩进 0.5 英寸；不强制编号标题。",
    ),
    "IEEE": StyleSpec(
        key="IEEE",
        name="IEEE 电气电子工程师学会",
        page=PageLayout(1.0, 1.0, 1.0, 1.0, "Times New Roman", 12.0, 1.5),
        reference=ReferenceRule(
            style_key="IEEE",
            hanging_indent_in=0.0,
            numbered=True,
            order_hint="顺序编码：[1] A. Author, \"Title,\" Journal, vol., no., pp., Year.",
            pattern_hint='[1] J. A. Smith, "Title," J. Name, vol. 1, no. 2, pp. 3-4, 2020.',
            checks=("numbered_ref",),
        ),
        heading_levels=0,
        notes="顺序编码制；会议模板常 2 栏 10pt，学位论文常 12pt/1.5 倍；以学校模板为准。",
    ),
    "Harvard": StyleSpec(
        key="Harvard",
        name="Harvard 哈佛体系（著者-出版年）",
        page=PageLayout(1.0, 1.0, 1.0, 1.0, "Times New Roman", 12.0, 2.0),
        reference=ReferenceRule(
            style_key="Harvard",
            hanging_indent_in=0.5,
            numbered=False,
            order_hint="著者-出版年-题名：Author, A. A. (Year) Title. Place: Publisher.",
            pattern_hint="Smith, J. A. (2020) Title of Work. London: Publisher.",
            checks=("has_year_paren", "author_year_order"),
        ),
        heading_levels=0,
        notes="无单一官方，各大学自有细则；以学校模板/问卷为准；悬挂缩进 0.5 英寸常见。",
    ),
    "Other": StyleSpec(
        key="Other",
        name="Other / 自定义（请上传学校模板或手填问卷）",
        page=PageLayout(None, None, None, None, None, None, None),
        reference=ReferenceRule(
            style_key="Other",
            hanging_indent_in=None,
            numbered=None,
            order_hint="由学校模板或格式问卷决定",
            pattern_hint="",
            checks=(),
        ),
        heading_levels=0,
        notes="规范不内置具体值，必须上传学校模板或填写格式问卷才能检查页面排版。",
    ),
}

# ---------------------------------------------------------------------------
# 正文段落级排版典型值（单一真相）：(对齐方式, 首行缩进英寸, 段后间距磅)
# 说明：均为「该风格典型文稿」的兜底参考，检查时 spec 来源会降级为 info；
#       学校模板 / 问卷 / AI 显式给出时才硬判。
# 依据：APA 7 / MLA 9 / Chicago 17 / Harvard 均为左对齐 + 首行缩进 0.5″ + 段间无空行；
#       IEEE 正文两端对齐、首行不缩进、段间无空行（IEEEtran 默认）。
# ---------------------------------------------------------------------------
_BODY_LAYOUT: dict[str, tuple] = {
    "APA":     ("left",    0.5,  0.0),
    "MLA":     ("left",    0.5,  0.0),
    "Chicago": ("left",    0.5,  0.0),
    "IEEE":    ("justify", 0.0,  0.0),
    "Harvard": ("left",    0.5,  0.0),
    "Other":   (None,      None, None),
}

for _k, (_al, _fi, _sa) in _BODY_LAYOUT.items():
    if _k in SPECS:
        SPECS[_k].page.body_alignment = _al
        SPECS[_k].page.first_line_indent_in = _fi
        SPECS[_k].page.space_after_pt = _sa

# 正文段落级排版的字段名（供 questionnaire / checker / report 复用，避免各处硬编码）
BODY_LAYOUT_KEYS = ("body_alignment", "first_line_indent_in", "space_after_pt")

# 页边距维度单一真相：(TargetProfile.page 字段名, docx 分节属性名, 显示标签)
# 供 checker / fixer / report 共用，避免三处各抄一份映射。
MARGIN_DIMS = (
    ("margin_top_in", "top", "上页边距"),
    ("margin_bottom_in", "bottom", "下页边距"),
    ("margin_left_in", "left", "左页边距"),
    ("margin_right_in", "right", "右页边距"),
)

# 下拉框展示顺序（白名单单一来源，新增规范只需在此增加 key 并补 SPECS）
SPEC_ORDER = ["APA", "MLA", "Chicago", "IEEE", "Harvard", "Other"]

# 引用风格白名单（与 SPEC_ORDER 中可引用者保持一致）
REFERENCE_STYLE_KEYS = ("APA", "MLA", "Chicago", "IEEE", "Harvard")


def get_spec(key: str) -> StyleSpec:
    """按 key 取规范；找不到回退到 Other。"""
    return SPECS.get(key, SPECS["Other"])


def list_specs() -> list[StyleSpec]:
    return [SPECS[k] for k in SPEC_ORDER]
