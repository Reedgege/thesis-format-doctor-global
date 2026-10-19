"""内置论文格式规范（离线、硬编码、零网络、零 LLM）。

PaperFormat Pro —— 海外版规则引擎的规范底本。

设计铁律（模板驱动原则）：
- 这些规范只定义「该风格自身的文稿格式指引」+「参考文献格式规则」。
- 页面排版（页边距 / 字体 / 字号 / 行距）**由学校模板或格式问卷决定**，
  规范里的 page 默认值仅作为「该风格典型文稿」的兜底参考（检查时不强制 fail）。
- 参考文献格式规则以所选引用风格为准，除非学校显式覆盖。

本文件不依赖任何外部网络或大模型，全部写死，保证离线安全。
"""

from __future__ import annotations

from dataclasses import dataclass, asdict, field
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
    header_distance_in: Optional[float] = None       # 页眉距页顶（英寸）；MLA 9th 要求 ½″


@dataclass
class ReferenceRule:
    """参考文献 / 引用格式规则。

    checks: 确定性检查标识元组，供 checker 数据驱动地分发引用检查。
    可选值见 checker._REF_CHECKS。
    """

    style_key: str                 # APA / MLA / IEEE / Other
    hanging_indent_in: Optional[float] = 0.5
    numbered: bool = False         # True=顺序编码制（如 IEEE），False=著者-出版年制
    order_hint: str = ""           # 人类可读的顺序说明
    pattern_hint: str = ""         # 给用户的示例
    checks: tuple = ()             # 确定性检查标识，供 checker 使用


@dataclass
class HeadingLevelFormat:
    """单级标题的格式规则（直接格式化，绝不写 styleId）。

    值直接落在段落 / run 属性上，故天然幂等，且不会写出文档里不存在的样式 id
    （国内版坑 1：styleId 孤儿 → 非幂等重标，海外版从根上规避）。

    number_prefix: 章节编号前缀样式（仅 IEEE 等「顺序编号制标题」风格使用）：
        None=不加编号；"roman"=大写罗马数字 "I. "；"upper_letter"="A. "；
        "number"="1. "；"lower_letter"="a) "；"lower_roman"="i) "。
        编号由 fixer 按文档顺序维护计数（同层 +1，遇到更浅层重置更深层），幂等
        （已带正确前缀则跳过、不重复叠加）。编号属 IEEE 等风格的显式约定，而非正文改写。
    uppercase: 标题文本全大写（IEEE 章节标题约定；天然幂等）。
    """

    alignment: Optional[str] = None      # "left" / "center" / "right"
    bold: bool = False
    italic: bool = False
    run_in: bool = False                 # L4/L5：标题与正文接排（同行）
    indent: bool = False                 # 缩进式标题（L4/L5 首行缩进）
    size_pt: Optional[float] = None      # None = 与正文同字号
    number_prefix: Optional[str] = None  # 见上；None = 不加章节编号
    uppercase: bool = False              # 标题文本全大写（IEEE 约定）
    desc: str = ""                       # 给用户的英文批注描述


@dataclass
class TitlePageFormat:
    """标题页（title page）标题 / 作者 / 单位的直接格式化规则。

    与 HeadingLevelFormat 同理：值直接落在段落 / run 属性，绝不写 styleId，天然幂等。
    字段为 None 表示该风格不规定（修正时不触碰该属性），而非「强制为 False」——
    这样「Other / 自定义模板」可整体留空、由模板驱动接管，而内置风格各自给出明确约定。

    各风格标题页约定（来自官方手册的典型文稿）：
    - APA 7：标题居中 + 加粗（title case）；作者 / 单位居中（不加粗）。
    - IEEE：标题居中 + 加粗；作者 / 单位居中（IEEEtran 默认）。
    - MLA 9：标题居中（不加粗）；作者块在首页左上方，不居中。
    """

    title_centered: Optional[bool] = None
    title_bold: Optional[bool] = None
    author_centered: Optional[bool] = None


@dataclass
class StyleSpec:
    """一个完整的引用风格规范。"""

    key: str
    name: str
    page: PageLayout
    reference: ReferenceRule
    heading_levels: int = 0       # 0 = 该风格不强制标题层级
    # 分节栏数（单栏=1 / 双栏=2）。None = 该风格不规定（由学校模板 / 用户决定）。
    # 仅 IEEE 标准两栏=2；APA/MLA 单栏=1；Other=None（模板驱动，不强制）。
    # 栏数是结构性页面维度：选中某规范作为**来源**即为其显式布局意图，checker 对此
    # 不做「spec→info」降级（与页边距/字体等软参考不同），不匹配时照常 warn。
    columns: Optional[int] = 1
    heading_formats: dict = field(default_factory=dict)  # 每级直接格式化规则（不写 styleId）
    title_page: Optional[TitlePageFormat] = None          # 标题页规则；None = 不自动处理
    # —— 学生/专业分层（APA 7th 专用，其余风格留默认中性值）——
    # 这些标记只为「前置/后置材料期望」服务：正文/标题/参考文献规则学生与专业完全一致，
    # 差异只在 running head、摘要、作者注、标题页字段（详见两份官方标注样例）。
    # 当前 checker 尚未硬查这些维度，标记先种入引擎，供报告「正确说明」+ 未来硬检查复用。
    running_head: bool = False            # 专业版要求页眉短标题+页码；学生版不需要
    expect_abstract: bool = False         # 专业版须含 Abstract+Keywords；学生版仅当老师要求
    expect_author_note: bool = False      # 专业版须含 Author Note；学生版无
    title_page_kind: str = ""             # "student" / "professional" / ""
    front_matter_note: str = ""           # 给用户的「前置材料期望」说明（报告展示）
    # —— APA 7th 专属格式约束（仅 APA 两档赋值；其余风格默认中性值，零影响）——
    # 以下维度 APA 7 手册明规定义、此前引擎未覆盖；均为保守检查（warn/info），
    # 不改变文档内容。其他引用风格（MLA/IEEE/Other）因取默认中性值，完全不受影响。
    page_number_position: Optional[str] = None   # 页码位置，如 "top-right"；None=不检查。APA=top-right
    reference_new_page: bool = False             # 参考文献须另起一页；APA=True
    font_allowlist: tuple = ()                   # 允许字体清单；空=不限制（模板驱动）。APA=多种
    running_head_max_chars: Optional[int] = None # 页眉短标题最大字符数；None=不检查。APA专业版=50
    reference_alpha_sort_hint: bool = False      # 参考文献字母序仅做提示（不动内容）；APA=True
    running_head_kind: str = "apa"               # running head 形态："apa"=短标题(专业版) / "mla"=姓氏+页码；仅 running_head=True 时生效
    reference_heading_bold: bool = True          # 参考文献标题是否须加粗：APA=True（粗体居中）、MLA=False（居中不加粗）
    reference_heading_bold_mode: str = "require"  # "require"=必须加粗(APA) / "forbid"=禁止加粗(MLA)；仅影响提示文案与判定
    notes: str = ""

    def to_dict(self) -> dict:
        return {
            "key": self.key,
            "name": self.name,
            "page": asdict(self.page),
            "reference": asdict(self.reference),
        "heading_levels": self.heading_levels,
        "columns": self.columns,
        "heading_formats": {str(k): asdict(v) for k, v in self.heading_formats.items()},
        "title_page": asdict(self.title_page) if self.title_page else None,
        "running_head": self.running_head,
        "expect_abstract": self.expect_abstract,
        "expect_author_note": self.expect_author_note,
        "title_page_kind": self.title_page_kind,
        "front_matter_note": self.front_matter_note,
        "page_number_position": self.page_number_position,
        "reference_new_page": self.reference_new_page,
        "font_allowlist": list(self.font_allowlist),
        "running_head_max_chars": self.running_head_max_chars,
        "reference_alpha_sort_hint": self.reference_alpha_sort_hint,
        "notes": self.notes,
        }


# ---------------------------------------------------------------------------
# 五个内置规范 + Other 兜底。数值均来自各风格官方手册的典型文稿要求。
# 注意：IEEE 标准两栏（会议/期刊）文稿常为 2 栏 10pt 单倍行距，这里取两栏典型值；
#       这些页面值仅作为「参考」而非硬性要求（页面以学校模板/问卷为准）。
# ---------------------------------------------------------------------------

# APA 7 五级的精确直接格式化规则（官方手册 §2.26–2.27）。绝不写 styleId。
APA_HEADING: dict[int, HeadingLevelFormat] = {
    1: HeadingLevelFormat("center", True, False, False, False, None,
                          "centered, bold, title case"),
    2: HeadingLevelFormat("left", True, False, False, False, None,
                          "flush-left, bold, title case"),
    3: HeadingLevelFormat("left", True, True, False, False, None,
                          "flush-left, bold italic, title case"),
    4: HeadingLevelFormat("left", True, False, True, True, None,
                          "indented, bold, title case, ending with a period (run-in)"),
    5: HeadingLevelFormat("left", True, True, True, True, None,
                          "indented, bold italic, title case, ending with a period (run-in)"),
}


# IEEE 章节标题编号方案（对齐 IEEE_Paper_Format_Rules V1.0 §3 + IEEEtran 实际渲染）：
# - L1：罗马数字 "I." + 全大写 + 加粗 + 居中（§3 示例 "I. INTRODUCTION"，与 IEEEtran 一致）；
# - L2：大写字母 "A." + 加粗 + 左对齐，标题保留作者原大小写（§3 示例 "A. Formal Framework"
#   为 title case，且 §3「自动修正建议」明确「不直接覆盖原文字内容，仅调整编号和格式」）；
# - L3/L4/L5：数字/小写字母编号（"1)" / "a)" / "i)"）+ 斜体 + 左对齐（IEEEtran subsubsection/paragraph）。
# 仅 L1 做全大写（编号 + 样式），L2–L5 不改文字大小写 —— 严守「不改内容」原则；
# 罗马数字递增由 fixer 按文档顺序维护（同层 +1、遇浅层重置更深层），幂等。
IEEE_HEADING: dict[int, HeadingLevelFormat] = {
    1: HeadingLevelFormat("center", True, False, False, False, None, "roman",
                          True, "centered, bold, roman numeral, uppercase"),
    2: HeadingLevelFormat("left", True, False, False, False, None, "upper_letter",
                          False, "flush-left, bold, capital letter"),
    3: HeadingLevelFormat("left", False, True, False, False, None, "number_paren",
                          False, "flush-left, italic, arabic numeral + paren"),
    4: HeadingLevelFormat("left", False, True, False, False, None, "lower_letter",
                          False, "flush-left, italic, lowercase letter"),
    5: HeadingLevelFormat("left", False, True, False, False, None, "lower_roman",
                          False, "flush-left, italic, lowercase roman"),
}


def generic_heading_formats(levels: int) -> dict:
    """学校自定义标题层级的保守通用方案（仅当规范本身不强制、且用户模板/问卷显式给出 heading_levels 时才用）。

    层级越深视觉权重越低：L1 居中加粗，L2 左对齐加粗，更深一级加斜体。绝不写 styleId。
    """
    out: dict = {}
    for lv in range(1, max(int(levels), 1) + 1):
        if lv == 1:
            out[lv] = HeadingLevelFormat("center", True, False, False, False, None,
                                         "centered, bold")
        elif lv == 2:
            out[lv] = HeadingLevelFormat("left", True, False, False, False, None,
                                         "flush-left, bold")
        else:
            out[lv] = HeadingLevelFormat("left", True, True, False, False, None,
                                         "flush-left, bold italic")
    return out


def _apa_spec(key: str, name: str, *, running_head: bool, expect_abstract: bool,
               expect_author_note: bool, title_page_kind: str, notes: str,
               page_number_position: Optional[str] = "top-right",
               reference_new_page: bool = True,
               # APA 7th (§2.19) 官方推荐的 6 款「无障碍、可读」字体：
               # 11pt Calibri / 11pt Arial / 10pt Lucida Sans Unicode /
               # 12pt Times New Roman / 11pt Georgia / 10pt Computer Modern。
               # 旧版误含 Lucida Console、Courier New（等宽字体，非 APA 推荐）且漏了 Computer Modern，
               # 会导致「用 Courier New 反而判定通过、用 Computer Modern 反而提示」的反向错误；此处修正。
               font_allowlist: tuple = ("Times New Roman", "Calibri", "Arial",
                                        "Georgia", "Lucida Sans Unicode",
                                        "Computer Modern"),
               running_head_max_chars: Optional[int] = None,
               reference_alpha_sort_hint: bool = True) -> StyleSpec:
    """构造一个 APA 7th 规范档（学生/专业共享同一套正文/标题/参考文献规则，仅前置材料不同）。

    新增的 APA 专属维度（页码位置 / 参考文献另起一页 / 字体允许清单 / 页眉≤50字符 /
    参考文献字母序提示）在此统一赋值，确保两档都带上；其余引用风格取默认中性值不受影响。
    """
    return StyleSpec(
        key=key,
        name=name,
        # 页面四边距 / 字体 / 字号 / 行距：典型 APA 文稿值，仅作参考（spec 来源检查降级 info）。
        # body_alignment="left" 与 first_line_indent_in=0.5 是 APA 7th 的硬性正文规则
        # （§2.22 不得两端对齐 / §2.24 每段首行缩进 0.5″）：此前规范未设这两个值，
        # 导致 checker 与 fixer 对这两项「零覆盖」。现补入 —— checker 出 info 提示、
        # fixer 真正落盘（标题/标题页/摘要/块引用已在对齐与缩进环节被保护，不会误伤）。
        page=PageLayout(margin_top_in=1.0, margin_bottom_in=1.0, margin_left_in=1.0,
                        margin_right_in=1.0, font_family="Times New Roman",
                        font_size_pt=12.0, line_spacing=2.0,
                        body_alignment="left", first_line_indent_in=0.5),
        reference=ReferenceRule(
            style_key="APA",
            hanging_indent_in=0.5,
            numbered=False,
            order_hint="Author-Year: Author, A. A. (Year). Title. Publisher.",
            pattern_hint="Smith, J. A. (2020). Title of work. Publisher.",
            checks=("has_year_paren", "author_year_order"),
        ),
        heading_levels=5,
        heading_formats=APA_HEADING,
        title_page=TitlePageFormat(True, True, True),
        running_head=running_head,
        expect_abstract=expect_abstract,
        expect_author_note=expect_author_note,
        title_page_kind=title_page_kind,
        page_number_position=page_number_position,
        reference_new_page=reference_new_page,
        font_allowlist=font_allowlist,
        running_head_max_chars=running_head_max_chars,
        reference_alpha_sort_hint=reference_alpha_sort_hint,
        notes=notes,
    )


# APA 7th 学生版（课程作业）：官方标注样例 student-paper-annotated.docx
# 无 running head / 作者注 / 摘要（除非老师要求）；标题页=标题(粗体居中)+姓名+院系+
# 课程号:课名+教师名+提交日期。正文与参考文献规则与专业版完全一致。
# 差异项：running_head/expect_abstract/expect_author_note 全 False；running_head_max_chars=None
# （学生版不要求页眉短标题，故不参与 running head 检查）。
_APA_STUDENT = _apa_spec(
    "APA-Student", "APA 7th - Student (Word)",
    running_head=False, expect_abstract=False, expect_author_note=False,
    title_page_kind="student",
    running_head_max_chars=None,
    notes="Student edition: the title page is Title (bold, centered) + Name + Department + "
          "Course No.:Course Name + Instructor + Submission date; no running head, author note, "
          "or abstract is required (unless your instructor asks). Body and reference rules are "
          "identical to the Professional edition.",
)
# APA 7th 专业版（投稿/发表）：官方标注样例 professional-paper-annotated.docx
# 需 running head（页眉短标题+页码）；标题页=作者+单位(上标编号)+作者注；须含 Abstract+Keywords；
# 可用内容脚注(2.13)；正文与参考文献规则与学生版完全一致。
# 差异项：running_head/expect_abstract/expect_author_note 全 True；running_head_max_chars=50（专业版硬性上限）。
_APA_PRO = _apa_spec(
    "APA-Professional", "APA 7th - Professional (Word)",
    running_head=True, expect_abstract=True, expect_author_note=True,
    title_page_kind="professional",
    running_head_max_chars=50,
    notes="Professional edition: requires a running head (short title in the header + page number); "
          "the title page is Author + Affiliation (superscript numbers) + Author Note; must include "
          "an Abstract + Keywords; content footnotes are allowed (2.13). Body and reference rules are "
          "identical to the Student edition.",
)

SPECS: dict[str, StyleSpec] = {
    # 「APA」保留为「学生版」的兼容别名：保证旧调用 / 旧测试（get_spec("APA")、
    # build_target("APA")、TargetProfile(spec_key="APA")）不破；UI/CLI 下拉只展示下方两档。
    "APA": _APA_STUDENT,
    "APA-Student": _APA_STUDENT,
    "APA-Professional": _APA_PRO,
    "MLA": StyleSpec(
        key="MLA",
        name="MLA 9th - Standard (Word)",
        page=PageLayout(1.0, 1.0, 1.0, 1.0, "Times New Roman", 12.0, 2.0, header_distance_in=0.5),
        reference=ReferenceRule(
            style_key="MLA",
            hanging_indent_in=0.5,
            numbered=False,
            order_hint="Author-Title-Publication: Author. Title. Publisher, Year.",
            pattern_hint="Smith, John. Title of Work. Publisher, 2020.",
            checks=("author_title_order",),
        ),
        heading_levels=0,
        title_page=TitlePageFormat(True, False, False),
        # —— MLA 9th 版式检查开关（Purdue OWL general_format + 样例论文，均经逐字核对）——
        page_number_position="top-right",          # 页码右上（MLA 要求每页右上姓氏+页码）
        reference_new_page=True,                   # Works Cited 另起一页
        reference_alpha_sort_hint=True,            # 按作者姓氏字母序（仅提示）
        running_head=True,                         # 要求 running head
        running_head_kind="mla",                  # = 姓氏 + 页码（非 APA 短标题）
        reference_heading_bold=False,             # Works Cited 标题不加粗
        reference_heading_bold_mode="forbid",    # MLA 明确禁止加粗（OWL: without bolding）
        notes="Usually no separate title page; the Works Cited list begins on a new page, titled "
              "'Works Cited' (centered, not bold), in alphabetical order by author's last name, with a "
              "0.5-inch hanging indent. Running head = your last name + page number, top-right, every page. "
              "First-page heading block (name / instructor / course / date) sits in the upper left, double-spaced. "
              "Numbered headings are not required (use a consistent format of your own).",
    ),
    "IEEE": StyleSpec(
        key="IEEE",
        name="IEEE - Standard Two-Column (Word)",
        page=PageLayout(1.0, 1.0, 1.0, 1.0, "Times New Roman", 10.0, 1.0),
        reference=ReferenceRule(
            style_key="IEEE",
            hanging_indent_in=0.0,
            numbered=True,
            order_hint="Sequential: [1] A. Author, \"Title,\" Journal, vol., no., pp., Year.",
            pattern_hint='[1] J. A. Smith, "Title," J. Name, vol. 1, no. 2, pp. 3-4, 2020.',
            checks=("numbered_ref",),
        ),
        heading_levels=5,
        columns=2,   # IEEE 标准两栏（会议/期刊）：选此规范即声明双栏布局意图
        heading_formats=IEEE_HEADING,
        title_page=TitlePageFormat(True, True, True),
        notes="Numeric (sequential) citation system [1]; section headings use Roman/Arabic "
              "numbering (I. / A. / 1.): level 1 centered + bold + UPPERCASE, level 2 flush-left "
              "+ bold, deeper levels italic. This option targets the standard IEEE two-column "
              "article format (conference & journal, e.g. IEEEtran two-column). Column layout and "
              "exact margins/font/size come from your school or IEEE template - follow it.",
    ),
    "Other": StyleSpec(
        key="Other",
        name="Other / Custom (upload your school template or fill the questionnaire)",
        page=PageLayout(None, None, None, None, None, None, None),
        reference=ReferenceRule(
            style_key="Other",
            hanging_indent_in=None,
            numbered=None,
            order_hint="Determined by your school template or questionnaire",
            pattern_hint="",
            checks=(),
        ),
        heading_levels=0,
        columns=None,   # 无内置栏数规定：由学校模板 / 问卷 / 复核弹窗决定
        notes="This style carries no built-in values; you must upload a school template or fill in "
              "the questionnaire before page layout can be checked.",
    ),
}

# ---------------------------------------------------------------------------
# 正文段落级排版典型值（单一真相）：(对齐方式, 首行缩进英寸, 段后间距磅)
# 说明：均为「该风格典型文稿」的兜底参考，检查时 spec 来源会降级为 info；
#       学校模板 / 问卷 / AI 显式给出时才硬判。
# 依据：APA 7 / MLA 9 均为左对齐 + 首行缩进 0.5″ + 段间无空行；
#       IEEE 正文两端对齐、首行不缩进、段间无空行（IEEEtran 默认）。
# ---------------------------------------------------------------------------
_BODY_LAYOUT: dict[str, tuple] = {
    "APA-Student":     ("left",    0.5,  0.0),
    "APA-Professional": ("left",   0.5,  0.0),
    "MLA":     ("left",    0.5,  0.0),
    "IEEE":    ("justify", 0.0,  0.0),
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
    ("margin_top_in", "top", "Top margin"),
    ("margin_bottom_in", "bottom", "Bottom margin"),
    ("margin_left_in", "left", "Left margin"),
    ("margin_right_in", "right", "Right margin"),
)

# 下拉框展示顺序（白名单单一来源，新增规范只需在此增加 key 并补 SPECS）。
# APA 7th 拆成学生/专业两档；「APA」别名不进此列表（避免下拉出现三个 APA 条目）。
SPEC_ORDER = ["APA-Student", "APA-Professional", "MLA", "IEEE", "Other"]

# 引用风格白名单（与 SPEC_ORDER 中可引用者保持一致）
REFERENCE_STYLE_KEYS = ("APA", "MLA", "IEEE")

# 规范展示名（面向用户，下拉 / 报告共用单一来源）。
# 与 SPEC_ORDER 一一对应；未列出者回退到原始 key（如 "Other" / 别名 "APA"）。
SPEC_DISPLAY = {
    "APA": "APA 7th - Student (Word)",
    "APA-Student": "APA 7th - Student (Word)",
    "APA-Professional": "APA 7th - Professional (Word)",
    "MLA": "MLA 9th - Standard (Word)",
    "IEEE": "IEEE - Standard Two-Column (Word)",
}


def spec_display(key: str) -> str:
    """规范 key → 面向用户的展示名（下拉 / 报告共用单一真相源）。"""
    return SPEC_DISPLAY.get(key, key)


def get_spec(key: str) -> StyleSpec:
    """按 key 取规范；找不到回退到 Other。"""
    return SPECS.get(key, SPECS["Other"])


def resolve_spec(target) -> StyleSpec:
    """从 TargetProfile 解析出应使用的完整规范（单一真相源）。

    这是 checker 与 fixer 共用的规范解析函数：两者都走同一表达式，
    杜绝「检查用一套规范、修正用另一套规范」的漂移（曾出现 fixer 在
    target.spec_key 为 None 时回退到 "Other"、而 checker 按 reference_style
    解析出 IEEE，导致检查与修正结论不一致的 bug）。

    解析优先级：spec_key（最权威，含学生/专业 tier 差异）> reference_style
    （仅当 spec_key 缺失时兜底，例如纯引用风格场景）。
    """
    key = getattr(target, "spec_key", None) or (getattr(target, "reference_style", None) or "Other")
    return get_spec(key)


def list_specs() -> list[StyleSpec]:
    return [SPECS[k] for k in SPEC_ORDER]
