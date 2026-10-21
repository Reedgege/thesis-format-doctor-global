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
    # v2.3.75：段前 + 段后（磅）。此前只有段后，等于「段落间距」永远只能修一半 ——
    # 学校模板写着「段前 6 磅」时既读不到也改不了。两者语义一致，三态 None=不规定。
    space_before_pt: Optional[float] = None         # 段前间距（磅）
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
    # v2.3.65：加粗 / 斜体改**三态**（None=本级不规定 / True=必须 / False=必须不是）。
    # 旧写法是纯 bool、默认 False，分不清「未规定」与「明确不加粗」；一旦据此修正，
    # 会去主动清除作者原有的加粗（越界）。三态下 None 一律不触碰该属性。
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    run_in: bool = False                 # L4/L5：标题与正文接排（同行）
    indent: bool = False                 # 缩进式标题（L4/L5 首行缩进）——旧的开关式写法
    # v2.3.65：缩进量可填（英寸）。学校模板里大量是「标题左缩进 0.5 英寸 / 首行缩进 0.25」
    # 这种具体值，只有一个开关表达不了。为 None 时回退到旧的 indent 开关（0.5 英寸）。
    indent_in: Optional[float] = None
    # v2.3.65：每级标题可以有自己的字体（很多学校要求「一级黑体、二级楷体」）。
    # None = 跟随正文字体（不臆造），这是旧行为，保持兼容。
    font_family: Optional[str] = None
    size_pt: Optional[float] = None      # None = 与正文同字号
    number_prefix: Optional[str] = None  # 见上；None = 不加章节编号
    uppercase: bool = False              # 标题文本全大写（IEEE 约定）
    # v2.3.67：段前分页（pageBreakBefore）。三态：None=本级不规定 / True=该级标题必须
    # 另起一页 / False=明确不另起一页。海外学术市场高频诉求：Journal Article 章首不分页、
    # Dissertation 每章（通常是 L1）强制另起一页。落在段落 pPr/pageBreakBefore（纯格式、
    # 幂等、不写 styleId）。
    page_break_before: Optional[bool] = None
    # v2.3.75：**段前 / 段后 / 行距** —— 三态，None=本级不规定（一个字节都不写）。
    # 老板 2026-10-20 追问「调整格式之中，段前段后有没有考虑？是核心的东西」。
    # 事实核查：此前整条链路只有正文段后一项（PageLayout.space_after_pt），标题级
    # 完全没有 spacing 字段 —— 学校模板写着「章标题段前 12 磅」，读不到、查不出、
    # 也改不了。规范侧确有硬要求：APA 7th §2.21「Do not add extra space before or
    # after headings」+「Double-space the entire paper」，故 APA 五级显式给
    # before=0 / after=0 / line=2.0；IEEE 官方未规定具体磅值（IEEEtran 模板各自为政），
    # 一律留 None，绝不臆造。
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None
    line_spacing: Optional[float] = None     # 倍数（2.0=双倍）；None=不规定
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
    # 纸张大小（Letter / A4 / Legal / A5）。None = 该风格不规定（由学校模板 / 用户决定）。
    # 🔴 这是**通用维度，不属于任何引用规范**：北美（美/加）多用 Letter（8.5×11 in），
    # 欧洲、中国、澳洲等用 A4（210×297 mm）。同一份 APA 论文在伦敦交就是 A4、
    # 在纽约交就是 Letter —— 所以真正说了算的是学校模板/客户，规范的只是「典型值参考」。
    # 因此 spec 来源固定降级为 info（只提示「APA 常用 Letter，你的是 A4，请确认学校要求」），
    # 学校模板 / 问卷 / AI 显式给出时才硬判。
    paper_size: Optional[str] = None
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
    # 字体 → 该字体在规范下的**合规字号**（像素级配对）。APA 7th §2.19 的允许清单不是
    # 「一串字体名」，而是「字体+字号」配对：TNR 12 / Calibri 11 / Arial 11 / Georgia 11 /
    # Lucida Sans Unicode 10 / Computer Modern 10。只比字体名不比字号 → Calibri 12pt 会被
    # 误判为合格。键为字体名（原样大小写），值为磅值；空 dict = 不做配对检查。
    # 🔴 有了这张表，修正就不必「替客户换字体」（换字体是越权：6 种都合规），
    #    只需**保留客户选的字体、把字号校正到该字体的规范值**（只动 w:sz，不动 w:rFonts）。
    font_allowlist_sizes: dict = field(default_factory=dict)
    running_head_max_chars: Optional[int] = None # 页眉短标题最大字符数；None=不检查。APA专业版=50
    reference_alpha_sort_hint: bool = False      # 参考文献字母序仅做提示（不动内容）；APA=True
    running_head_kind: str = "apa"               # running head 形态："apa"=短标题(专业版) / "mla"=姓氏+页码；仅 running_head=True 时生效
    reference_heading_bold: bool = True          # 参考文献标题是否须加粗：APA=True（粗体居中）、MLA=False（居中不加粗）
    reference_heading_bold_mode: str = "require"  # "require"=必须加粗(APA) / "forbid"=禁止加粗(MLA)；仅影响提示文案与判定
    # —— 摘要页特化（v2.3.67）——
    # APA 7th 摘要正文**不首行缩进**（顶格）且专业版摘要**另起一页**；其它风格默认中性值。
    # None = 该风格不规定（由学校模板 / 问卷 / AI 决定），不臆造。
    abstract_first_line_indent_in: Optional[float] = None   # 摘要正文首行缩进（英寸）；None=不规定
    abstract_new_page: Optional[bool] = None                 # 摘要是否另起一页；None=不规定
    # —— 图表 caption（v2.3.67）——
    # 图表标题的字号/行距可修（识别 Caption 样式段落）；表题/图题的「上方还是下方」位置
    # 只能提示（python-docx 无高层 API 可靠判定图片 caption 位置，图片内嵌在段落 drawing 里）。
    # None = 该风格不规定（由学校模板 / 问卷 / AI 决定）。APA/IEEE 表题在上、图题在下。
    caption_font_size_pt: Optional[float] = None             # 图表标题字号（磅）
    caption_line_spacing: Optional[float] = None             # 图表标题行距（倍数）
    caption_position_table: Optional[str] = None             # "top" / "bottom"（表题在表格上方/下方）
    caption_position_figure: Optional[str] = None            # "top" / "bottom"（图题在图上方/下方）
    # —— 表格线（v2.3.85）——
    # "three_line" = 学术三线表：只保留**顶线 / 表头下线 / 底线**，四边竖线与内部横线全部去掉。
    # "boxed"      = 允许（或要求）全框线。
    # None         = 该风格不规定（由学校模板 / 问卷 / 客户决定），不检查也不改动。
    # 依据：APA 7th §7.10（表格只允许必要的横线，禁止竖线）、IEEEtran 的 open table、
    # MLA 的表格惯例，三家都是「无竖线、只用横线分区」= 三线表。
    # 与栏数同为**版式结构维度**：选中某规范即为其显式版式意图，checker 不把它降级成 info。
    table_lines: Optional[str] = None
    notes: str = ""

    def to_dict(self) -> dict:
        return {
            "key": self.key,
            "name": self.name,
            "page": asdict(self.page),
            "reference": asdict(self.reference),
        "heading_levels": self.heading_levels,
        "columns": self.columns,
        "paper_size": self.paper_size,
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
        "font_allowlist_sizes": dict(self.font_allowlist_sizes),
        "running_head_max_chars": self.running_head_max_chars,
        "reference_alpha_sort_hint": self.reference_alpha_sort_hint,
        "abstract_first_line_indent_in": self.abstract_first_line_indent_in,
        "abstract_new_page": self.abstract_new_page,
        "caption_font_size_pt": self.caption_font_size_pt,
        "caption_line_spacing": self.caption_line_spacing,
        "caption_position_table": self.caption_position_table,
        "caption_position_figure": self.caption_position_figure,
        "table_lines": self.table_lines,
        "notes": self.notes,
        }


# ---------------------------------------------------------------------------
# 五个内置规范 + Other 兜底。数值均来自各风格官方手册的典型文稿要求。
# 注意：IEEE 标准两栏（会议/期刊）文稿常为 2 栏 10pt 单倍行距，这里取两栏典型值；
#       这些页面值仅作为「参考」而非硬性要求（页面以学校模板/问卷为准）。
# ---------------------------------------------------------------------------

# APA 7th §2.19 官方推荐的 6 款「无障碍、可读」字体，**每种字体绑定自己的合规字号**。
# 这张表是「字体白名单」的单一真相源：清单本身 = 键序列，配对字号 = 值。
#   Times New Roman 12pt / Calibri 11pt / Arial 11pt /
#   Georgia 11pt / Lucida Sans Unicode 10pt / Computer Modern 10pt
# 🔴 只比字体名不比字号是错的（Calibri 12pt 会被误判合格）；旧版还误含 Lucida Console、
#    Courier New（等宽，非 APA 推荐）、漏了 Computer Modern，已在 v2.3.65 修正。
APA_FONT_SIZES: dict[str, float] = {
    "Times New Roman": 12.0,
    "Calibri": 11.0,
    "Arial": 11.0,
    "Georgia": 11.0,
    "Lucida Sans Unicode": 10.0,
    "Computer Modern": 10.0,
}


def size_for_font(family: Optional[str],
                  pairs: Optional[dict] = None) -> Optional[float]:
    """给定字体名，返回该字体在``pairs``表中的合规字号（磅）；查不到返回 None。

    归一规则：忽略大小写、折叠空白、忽略引号（Word 里 "Times New Roman" 可能带
    撇号或用 ``TimesNewRoman`` 变体）。查不到时**绝不猜**——宁可只提示，也不按
    别的字体的字号去改客户的文档。
    """
    if not family or not pairs:
        return None

    def _norm(s: str) -> str:
        return "".join(str(s).replace("'", "").replace("'", "")
                       .replace('"', "").split()).lower()

    want = _norm(family)
    for k, v in pairs.items():
        if _norm(k) == want:
            return float(v)
    return None


# APA 7 五级的精确直接格式化规则（官方手册 §2.26–2.27）。绝不写 styleId。
# ⚠️ 一律**关键字传参**：v2.3.65 给 HeadingLevelFormat 增补 indent_in / font_family
# 两个字段，插在 size_pt 之前；旧的位置写法会让第 6~9 位参数整体错位（"roman" 会
# 落到 font_family、"desc" 会落到 number_prefix → 编号前缀被塞进标题正文）。
#
# v2.3.75：**段前/段后/行距**按 APA 7th 原文给死 ——
#   §2.21 "Do not add blank lines before or after headings"（标题上下都不许空行）
#   §2.21 "Double-space the entire paper"（含标题，故行距 2.0）
# 这三项只在「规范基线」模式（客户选的/系统判定为 APA）下生效，模板/手动/AI 一律以
# 客户提供的为准（spec_is_baseline 门控）—— 绝不拿 APA 去覆盖学校的段前磅值。
_APA_SPACING = {"space_before_pt": 0.0, "space_after_pt": 0.0, "line_spacing": 2.0}

APA_HEADING: dict[int, HeadingLevelFormat] = {
    1: HeadingLevelFormat(alignment="center", bold=True, **_APA_SPACING,
                          desc="centered, bold, title case, no extra space above/below"),
    2: HeadingLevelFormat(alignment="left", bold=True, **_APA_SPACING,
                          desc="flush-left, bold, title case, no extra space above/below"),
    3: HeadingLevelFormat(alignment="left", bold=True, italic=True, **_APA_SPACING,
                          desc="flush-left, bold italic, title case, no extra space above/below"),
    4: HeadingLevelFormat(alignment="left", bold=True, run_in=True,
                          indent=True, indent_in=0.5, **_APA_SPACING,
                          desc="indented, bold, title case, ending with a period (run-in)"),
    5: HeadingLevelFormat(alignment="left", bold=True, italic=True, run_in=True,
                          indent=True, indent_in=0.5, **_APA_SPACING,
                          desc="indented, bold italic, title case, ending with a period (run-in)"),
}


# IEEE 章节标题编号方案（对齐 IEEE_Paper_Format_Rules V1.0 §3 + IEEEtran 实际渲染）：
# - L1：罗马数字 "I." + 全大写 + 加粗 + 居中（§3 示例 "I. INTRODUCTION"，与 IEEEtran 一致）；
# - L2：大写字母 "A." + 加粗 + 左对齐，标题保留作者原大小写（§3 示例 "A. Formal Framework"
#   为 title case，且 §3「自动修正建议」明确「不直接覆盖原文字内容，仅调整编号和格式」）；
# - L3/L4/L5：数字/小写字母编号（"1)" / "a)" / "i)"）+ 斜体 + 左对齐（IEEEtran subsubsection/paragraph）。
# 仅 L1 做全大写（编号 + 样式），L2–L5 不改文字大小写 —— 严守「不改内容」原则；
# 罗马数字递增由 fixer 按文档顺序维护（同层 +1、遇浅层重置更深层），幂等。
IEEE_HEADING: dict[int, HeadingLevelFormat] = {
    1: HeadingLevelFormat(alignment="center", bold=True, number_prefix="roman",
                          uppercase=True,
                          desc="centered, bold, roman numeral, uppercase"),
    2: HeadingLevelFormat(alignment="left", bold=True, number_prefix="upper_letter",
                          desc="flush-left, bold, capital letter"),
    3: HeadingLevelFormat(alignment="left", italic=True, number_prefix="number_paren",
                          desc="flush-left, italic, arabic numeral + paren"),
    4: HeadingLevelFormat(alignment="left", italic=True, number_prefix="lower_letter",
                          desc="flush-left, italic, lowercase letter"),
    5: HeadingLevelFormat(alignment="left", italic=True, number_prefix="lower_roman",
                          desc="flush-left, italic, lowercase roman"),
}


def generic_heading_formats(levels: int) -> dict:
    """学校自定义标题层级的保守通用方案（仅当规范本身不强制、且用户模板/问卷显式给出 heading_levels 时才用）。

    层级越深视觉权重越低：L1 居中加粗，L2 左对齐加粗，更深一级加斜体。绝不写 styleId。
    """
    out: dict = {}
    for lv in range(1, max(int(levels), 1) + 1):
        if lv == 1:
            out[lv] = HeadingLevelFormat(alignment="center", bold=True,
                                         desc="centered, bold")
        elif lv == 2:
            out[lv] = HeadingLevelFormat(alignment="left", bold=True,
                                         desc="flush-left, bold")
        else:
            out[lv] = HeadingLevelFormat(alignment="left", bold=True, italic=True,
                                         desc="flush-left, bold italic")
    return out


# ---------------------------------------------------------------------------
# 每级标题规则的统一解析（v2.3.65）
#
# 背景：每级标题参数有四个可能的来源 —— 学校模板（docx 采样）、手填问卷、
# AI 填表、内置规范。历史上只有内置规范这一路真的被引擎消费：
#   ``fmt_map = spec.heading_formats or generic_heading_formats(...)``
# 于是学校模板明明写着「一级标题左对齐、黑体三号」，修正器却按写死的
# 「L1 居中加粗」去改 —— 模板驱动的铁律被绕过。这里把「谁说了算」收成一处：
#   显式来源（模板/问卷/AI，即 TargetProfile.heading_formats）
#     > 内置规范规则 > 最深层回退。
# ---------------------------------------------------------------------------

_HEADING_ALIGN_VALUES = ("left", "center", "right", "justify")
# 与 fixer._section_number_prefix 支持的形态一一对应
_NUMBER_PREFIX_VALUES = ("number", "roman", "upper_letter", "lower_letter",
                         "lower_roman", "number_paren")


def describe_heading(f: "HeadingLevelFormat") -> str:
    """给一级标题规则拼一句英文说明（用于报告 / 文档批注，绝不外泄中文）。"""
    bits: list = []
    if f.alignment:
        bits.append({"left": "flush-left", "center": "centered",
                     "right": "flush-right", "justify": "justified"}.get(f.alignment, f.alignment))
    if f.bold:
        bits.append("bold")
    if f.italic:
        bits.append("italic")
    if f.uppercase:
        bits.append("uppercase")
    if f.font_family:
        bits.append(str(f.font_family))
    if f.size_pt:
        bits.append(("%g pt" % float(f.size_pt)))
    ind = f.indent_in if f.indent_in is not None else (0.5 if f.indent else None)
    if ind:
        bits.append(("first-line indent %g in" % float(ind)))
    if f.number_prefix:
        bits.append("numbered (%s)" % f.number_prefix)
    if f.run_in:
        bits.append("run-in with body text")
    if f.page_break_before is True:
        bits.append("page break before")
    elif f.page_break_before is False:
        bits.append("no page break before")
    # 段前 / 段后：两者都为 0 时合并成一句（APA 的原话就是 "no extra space"），
    # 其余情形如实分开报 —— 报告里必须让客户看得见磅值。
    _sb, _sa = f.space_before_pt, f.space_after_pt
    if _sb is not None and _sa is not None and float(_sb) == 0.0 and float(_sa) == 0.0:
        bits.append("no extra space above or below")
    else:
        if _sb is not None:
            bits.append("%g pt before" % float(_sb))
        if _sa is not None:
            bits.append("%g pt after" % float(_sa))
    if f.line_spacing is not None:
        bits.append("%gx line spacing" % float(f.line_spacing))
    return ", ".join(bits) if bits else "format not specified"


def heading_level_format_from_dict(raw: dict) -> "HeadingLevelFormat":
    """把手填 / AI / 学校模板给出的**一级标题 dict** 转成 HeadingLevelFormat。

    只认白名单字段；值不合法就丢掉（不臆造）。缺省字段保持 None/False 语义：
    ``bold=None`` / ``italic=None`` 表示「本级不规定该属性」，修正器一个字节都不写。
    """
    raw = raw if isinstance(raw, dict) else {}

    def _num(key):
        v = raw.get(key)
        if v is None or v == "":
            return None
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    def _tri(key):
        v = raw.get(key)
        if v is None or v == "":
            return None
        if isinstance(v, str):
            return v.strip().lower() in ("1", "true", "yes", "y")
        return bool(v)

    al = raw.get("alignment")
    al = str(al).strip().lower() if al is not None else None
    if al not in _HEADING_ALIGN_VALUES:
        al = None
    prefix = raw.get("number_prefix")
    if prefix is not None:
        prefix = str(prefix).strip()
        if prefix not in _NUMBER_PREFIX_VALUES:
            prefix = None          # 不认识就不编号，绝不照着乱塞前缀
    font = raw.get("font_family")
    font = str(font).strip() if font else None

    f = HeadingLevelFormat(
        alignment=al,
        bold=_tri("bold"),
        italic=_tri("italic"),
        run_in=bool(_tri("run_in")),
        indent=bool(_tri("indent")),
        indent_in=_num("indent_in"),
        font_family=font,
        size_pt=_num("size_pt"),
        number_prefix=prefix,
        uppercase=bool(_tri("uppercase")),
        page_break_before=_tri("page_break_before"),
        space_before_pt=_num("space_before_pt"),
        space_after_pt=_num("space_after_pt"),
        line_spacing=_num("line_spacing"),
        desc=str(raw.get("desc") or "").strip(),
    )
    if not f.desc:
        f.desc = describe_heading(f)
    return f


def resolve_heading_formats(explicit: Optional[dict] = None,
                            spec_formats: Optional[dict] = None,
                            levels: Optional[int] = None) -> dict:
    """决定「每级标题用哪套规则」，单一真相（修正器与检查器共用）。

    优先级：``explicit``（学校模板 / 手填问卷 / AI 填表）> ``spec_formats``（内置规范）
    > ``generic_heading_formats(levels)``（只在两层都空时才兜底的通用方案）。

    ⚠️ 有 ``explicit`` 时**只用 explicit**，绝不拿内置规范去补它没写的层级 ——
    学校模板就是学校模板，混进 APA 的层级规则等于偷偷换了一套格式（模板驱动铁律）。
    explicit 没覆盖的深层级，由调用方按「回退到最深层规则」处理。
    """
    explicit = explicit or {}
    spec_formats = spec_formats or {}
    if not explicit:
        if spec_formats:
            return dict(spec_formats)
        return generic_heading_formats(levels or 0)
    out: dict = {}
    top = max(int(k) for k in explicit)
    for lv in range(1, top + 1):
        if lv in explicit:
            out[lv] = heading_level_format_from_dict(explicit[lv])
    return out


def _apa_spec(key: str, name: str, *, running_head: bool, expect_abstract: bool,
               expect_author_note: bool, title_page_kind: str, notes: str,
               page_number_position: Optional[str] = "top-right",
               reference_new_page: bool = True,
               # APA 7th (§2.19) 的 6 款推荐字体与各自字号见模块级 ``APA_FONT_SIZES``
               # （单一真相源；清单由该表的键序列派生，避免两处各写一份走样）。
               font_allowlist: tuple = tuple(APA_FONT_SIZES),
               running_head_max_chars: Optional[int] = None,
               reference_alpha_sort_hint: bool = True,
               abstract_first_line_indent_in: Optional[float] = 0.0,
               abstract_new_page: Optional[bool] = None,
               caption_position_table: Optional[str] = "top",
               caption_position_figure: Optional[str] = "bottom") -> StyleSpec:
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
        paper_size="Letter",     # APA 7th §2.1：8.5 × 11 in（北美 Letter）
        title_page=TitlePageFormat(True, True, True),
        running_head=running_head,
        expect_abstract=expect_abstract,
        expect_author_note=expect_author_note,
        title_page_kind=title_page_kind,
        page_number_position=page_number_position,
        reference_new_page=reference_new_page,
        font_allowlist=font_allowlist,
        # 字体→合规字号配对表（APA 6 款）。修正时只改字号、绝不改客户选的字体。
        font_allowlist_sizes=dict(APA_FONT_SIZES),
        running_head_max_chars=running_head_max_chars,
        reference_alpha_sort_hint=reference_alpha_sort_hint,
        abstract_first_line_indent_in=abstract_first_line_indent_in,
        abstract_new_page=abstract_new_page,
        caption_position_table=caption_position_table,
        caption_position_figure=caption_position_figure,
        # APA 7th §7.10：表格只用横线，**不许有竖线**（就要求是顶线/表头下线/底线）。
        table_lines="three_line",
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
    abstract_new_page=True,   # 专业版摘要须另起一页（紧跟标题页之后独立一页）
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
        paper_size="Letter",     # MLA 9th / Purdue OWL：8.5 × 11 in
        title_page=TitlePageFormat(True, False, False),
        # —— MLA 9th 版式检查开关（Purdue OWL general_format + 样例论文，均经逐字核对）——
        page_number_position="top-right",          # 页码右上（MLA 要求每页右上姓氏+页码）
        reference_new_page=True,                   # Works Cited 另起一页
        reference_alpha_sort_hint=True,            # 按作者姓氏字母序（仅提示）
        running_head=True,                         # 要求 running head
        running_head_kind="mla",                  # = 姓氏 + 页码（非 APA 短标题）
        reference_heading_bold=False,             # Works Cited 标题不加粗
        reference_heading_bold_mode="forbid",    # MLA 明确禁止加粗（OWL: without bolding）
        table_lines="three_line",                 # MLA 表格惯例：无竖线，只用横线分区
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
        paper_size=None,   # IEEEtran 官方模板 A4 / US Letter 两版并存 → 不规定，由模板决定
        heading_formats=IEEE_HEADING,
        title_page=TitlePageFormat(True, True, True),
        # 图表 caption（v2.3.67）：IEEE 表题在表格上方、图题在图下方（IEEEtran 约定）。
        caption_position_table="top",
        caption_position_figure="bottom",
        table_lines="three_line",   # IEEEtran 表格为 open table：无竖线，顶线/表头线/底线
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
# 正文段落级排版典型值（单一真相）：(对齐方式, 首行缩进英寸, 段前磅, 段后磅)
# 说明：均为「该风格典型文稿」的兜底参考，检查时 spec 来源会降级为 info；
#       学校模板 / 问卷 / AI 显式给出时才硬判。
# 依据：APA 7 / MLA 9 均为左对齐 + 首行缩进 0.5″ + 段间无空行（§2.21：段前段后都不许
#       加空行）；IEEE 正文两端对齐、首行不缩进、段间无空行（IEEEtran 默认）。
# ---------------------------------------------------------------------------
_BODY_LAYOUT: dict[str, tuple] = {
    "APA-Student":     ("left",    0.5,  0.0,  0.0),
    "APA-Professional": ("left",   0.5,  0.0,  0.0),
    "MLA":     ("left",    0.5,  0.0,  0.0),
    "IEEE":    ("justify", 0.0,  0.0,  0.0),
    "Other":   (None,      None, None, None),
}

for _k, (_al, _fi, _sb, _sa) in _BODY_LAYOUT.items():
    if _k in SPECS:
        SPECS[_k].page.body_alignment = _al
        SPECS[_k].page.first_line_indent_in = _fi
        SPECS[_k].page.space_before_pt = _sb
        SPECS[_k].page.space_after_pt = _sa

# 正文段落级排版的字段名（供 questionnaire / checker / report 复用，避免各处硬编码）
BODY_LAYOUT_KEYS = ("body_alignment", "first_line_indent_in",
                    "space_before_pt", "space_after_pt")

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

# 「按我校要求（自定义）」：不套任何内置规范的引用体例。
# 🔴 v2.3.73：客户走手动输入 / AI 导入，本来就是因为有学校自己的要求或自己的个性化值，
# 旧版把参考文献定死成三选一（默认 APA）＝替客户拍板。现在它是**默认值**：
# 只执行客户填的/AI 读出的条目细则，不拿 APA/MLA/IEEE 去核条目内容。
# 注意：它不影响 REFERENCE_STYLE_KEYS 的语义（那一组才代表「真的选了某个规范」）。
REFERENCE_STYLE_CUSTOM = "custom"
REFERENCE_STYLE_CHOICES = (REFERENCE_STYLE_CUSTOM,) + REFERENCE_STYLE_KEYS

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


def spec_is_baseline(target) -> bool:
    """客户是否**真的在「规范」下拉里选了内置规范**（= 全篇按规范走）。

    🔴 与 ``resolve_spec()`` 的分工（v2.3.73，老板 2026-10-20 定的口径）：
    ``resolve_spec`` 在 ``spec_key`` 缺失时会按 ``reference_style`` 兜底取**整套规范**，
    这是为「参考文献归规范」设计的 —— 客户只选了「参考文献按 IEEE」，检查参考文献
    条目就该用 IEEE 的引用规则。但整套规范里还带着**五级标题格式、标题页、页面**，
    这些**绝不能**跟着外溢到学校模板 / 手动输入 / AI 导入这三条通道：

        客户走这三条路，恰恰是因为现行规范满足不了他的需求
        （学校有自己的排版要求 / 手填自己的个性化值 / AI 已经读出了他文档的真实值）。
        这时再把全篇按 APA、IEEE 重排，等于把他上传的东西全废掉。

    所以「拿规范兜底页面 / 标题 / 标题页」的地方，一律先问本函数。
    实测事故（v2.3.72 及以前）：手动输入只填页面 + 参考文献选 IEEE + 标题层级填 2，
    点「修正」后论文大标题变成 ``I. THE IMPACT OF AI ON ACADEMIC WRITING``、
    章节标题变成 ``A. Introduction / B. Method / C. Results``（预览里一个字都没提）。
    """
    key = getattr(target, "spec_key", None)
    return bool(key) and str(key) != "Other"


def list_specs() -> list[StyleSpec]:
    return [SPECS[k] for k in SPEC_ORDER]
