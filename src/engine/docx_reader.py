"""docx 读取器：把 Word 文档解析成 DocumentProfile（确定性、纯本地）。

只负责「读」，不负责「判」。判定在 checker.py。

能力：
- 多分节页边距取众数 + 不一致标记；
- 正文字体 / 字号 / 行距（段落 → 样式 → 文档默认三级回退）；
- 正文段落级排版：对齐方式 / 首行缩进 / 段后间距（同样三级回退）；
- 字体/字号/行距**只采样正文段落**（排除标题段与参考文献段），表格内文字另行纳入
  （cell 去重，避免合并单元格重复计数），避免「正文字体/字号」众数被带偏；
- 标题识别：本地化样式名（英/西/法/德/中/日/韩/荷/波/俄/意）+ 段落级/样式级
  outlineLvl + 伪标题启发式（短、整段加粗、字号 ≥13pt）；
- 参考文献段识别加固（单一实现，两处复用，避免正文误判）；
- 受保护 / 损坏 / 加密文档一律抛出 DocxReadError（含「能打开但内部 XML 异常」）。
"""

from __future__ import annotations

import io
import re
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from typing import Optional

from docx import Document
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from copy import deepcopy

# 软换行（Shift+Enter）：真实学生论文用它分隔标题/正文/参考文献，必须当段边界看待。
# 注意 w:br 还有 page/column 两个变体，属于排版语义，不能当段边界 —— 见 _is_soft_break。
def _is_soft_break(el) -> bool:
    """是不是「软换行」（Shift+Enter）。

    ⚠️ ``<w:br>`` 有三个变体：``textWrapping``（默认，软换行）、``page``（分页符）、
    ``column``（分栏符）。后两者是**排版语义**，绝不能被当段边界拆掉 ——
    把分页符拆没会让客户丢分页（tests/test_review_fixes.py::test_numbering_preserves_page_break
    就是这条守卫）。只认 ``w:cr`` 与「无 type / type=textWrapping」的 ``w:br``。
    """
    if el.tag == qn("w:cr"):
        return True
    if el.tag == qn("w:br"):
        return el.get(qn("w:type")) in (None, "textWrapping")
    return False

# 参考文献标题常见写法（含前缀/包含匹配）
_REF_HEADING_MARKERS = (
    "references", "bibliography", "works cited", "reference list",
    "参考文献", "参考资料", "引用文献",
)
# 标题段识别的最大长度（超过则不太可能是纯标题行）
REF_HEADING_MAX_LEN = 40

# 标题样式关键词（本地化，覆盖主要语种）：英/西/法/德/中/日/韩/荷/波/俄/意
# 这些关键词后面通常跟层级数字（Heading 1 / 标题 2 / Kop 3 / Заголовок 1）。
_HEADING_KEYWORDS = (
    "heading", "título", "titulo", "titre", "überschrift", "titel",
    "标题", "標題", "题目", "題目", "見出し", "제목", "kop", "nagłówek", "naglowek",
    "заголовок", "intestazione",
)

# 「整段即标题」（无层级编号）的样式名：Word 内置 Title / Subtitle 与各语种对应写法。
# 漏判会让论文主标题（24–28pt）被一键修正压成正文 12pt，视觉层级崩塌。
_TITLE_STYLE_KEYWORDS = (
    "title", "subtitle", "标题", "標題", "题目", "題目", "見出し",
    "제목", "titel", "título", "titulo", "titre", "заголовок",
)

# 伪标题启发式参数（保守取值，避免把加粗正文句误判成标题）
_PSEUDO_HEADING_MAX_LEN = 80
_PSEUDO_HEADING_MIN_PT = 13.0          # 保护用途阈值（高，避免误伤）
# 识别用途阈值：目标规范显式声明 heading_levels>0 时，用更低的阈值把正文字号
# (11–12pt) 加粗的章节标题也识别出来（IEEE/APA 常见这种标题，没有 13pt 那么大）
_PSEUDO_HEADING_FORMAT_MIN_PT = 10.0

# run 字体名读取顺序：先西文（ascii/hAnsi），再 eastAsia/cs 兜底
_FONT_ATTRS = ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs")


def _build_align_map() -> dict:
    """WD_ALIGN_PARAGRAPH 枚举 → 归一化标签（不同 python-docx 版本枚举名有差异，用 getattr 容错）。"""
    pairs = (
        ("LEFT", "left"), ("CENTER", "center"), ("RIGHT", "right"),
        ("JUSTIFY", "justify"), ("DISTRIBUTE", "justify"),
        ("JUSTIFY_LOW", "justify"), ("JUSTIFY_MED", "justify"),
        ("JUSTIFY_HI", "justify"), ("THAI_JUSTIFY", "justify"),
    )
    out: dict = {}
    for name, label in pairs:
        val = getattr(WD_ALIGN_PARAGRAPH, name, None)
        if val is not None:
            out[val] = label
    return out


_ALIGN_MAP = _build_align_map()


class DocxReadError(Exception):
    """文档读取失败（非 docx / 损坏 / 加密 / 受保护）。"""


@dataclass
class DocumentProfile:
    margins: dict = field(default_factory=dict)        # top/right/bottom/left (英寸)
    margins_inconsistent: bool = False                 # 多分节页边距是否不一致
    # 分节栏数（单栏=1 / 双栏=2 …）：读 sectPr/w:cols/@w:num，缺省视为 1。
    # python-docx 1.2.0 的 Section 没有栏数属性，故直读 sectPr XML（见 _section_columns）。
    # 多分节取众数；读不到任何分节时保持 None（不臆造）。
    columns: Optional[int] = None
    # v2.3.100：body 里的分节数量。>1 表示文档是"标题节/作者节/正文节"结构
    # （IEEE 会议模板典型 5~7 节），此时页边距/栏数只按**正文节**判（见 _body_section）。
    body_section_count: int = 0
    # 纸张大小（Letter / A4 / Legal / A5 / None=非标准或读不到）。**通用维度**：
    # 北美常用 Letter，欧洲/中国常用 A4，与引用规范无关，由学校模板/客户决定。
    # 多分节取众数；横向文档按「短边×长边」归一后识别（不会因宽高对调漏判）。
    paper_size: Optional[str] = None
    paper_width_in: Optional[float] = None             # 实测宽（英寸），写实值便于报告展示
    paper_height_in: Optional[float] = None            # 实测高（英寸）
    font_family: Optional[str] = None                  # 正文最常见字体（含表格文字）
    # 正文**每种字体各自的众数字号**：{字体名: 磅值}。
    # 为什么不能只看「正文字体众数 + 正文字号众数」：一篇稿子正文里混用 Times New Roman
    # 与 Calibri 时，APA 对两者的合规字号不同（12 / 11）。只看众数会把混用当成单一字体，
    # Calibri 那几段永远判不出来、也就永远改不到位。
    font_size_by_family: dict = field(default_factory=dict)
    font_size_pt: Optional[float] = None               # 正文最常见字号（含表格文字）
    line_spacing: Optional[float] = None               # 倍数（固定值/最小值返回 None）
    body_alignment: Optional[str] = None               # left / justify / center / right
    first_line_indent_in: Optional[float] = None       # 正文首行缩进（英寸）
    space_after_pt: Optional[float] = None             # 段后间距（磅）
    space_before_pt: Optional[float] = None            # 段前间距（磅，v2.3.75）
    table_paragraph_count: int = 0                     # 纳入采样的表格段落数
    # v2.3.77：表格内单元格段落的段前 / 段后（磅，众数）。表格里「不要多余空行」是
    # 学校硬要求里很常见的一条，此前完全读不到（表格只被采样了字体/字号/行距）。
    table_space_before_pt: Optional[float] = None
    table_space_after_pt: Optional[float] = None
    # v2.3.98：每张表与**前后文字**之间的空隙（见 table_gap_samples）。
    # 每条：{index, before_pt, before_blank, after_pt, after_blank}；磅，None=不判/已有空行。
    table_gap_samples: list = field(default_factory=list)
    # —— 表格边框（三线表，v2.3.85）——
    # 每张表一条 dict，见 sample_table_lines()：has_top / has_bottom / has_header_line /
    # has_vertical / has_inside_h / three_line。空列表 = 文档里没有表格（不是"合规"）。
    table_line_states: list = field(default_factory=list)
    # —— 公式（Word 原生 OMML，v2.3.112）——
    # 每个「含公式的段落」一条 dict，见 sample_equations()：display / alignment /
    # size_pt / space_before_pt / space_after_pt / line_spacing / number_text /
    # number_at_right / in_table。空列表 = 文档里没有 OMML 公式（不是"合规"）。
    equation_samples: list = field(default_factory=list)
    # —— 目录（TOC）探测（v2.3.77）——
    # 目录条目是**域段落**，fixer 一律不碰（改直接格式会被「更新目录」冲掉）；
    # 这里只负责认出「有没有目录 / 有几条 / 目录标题与条目的实际间距」，
    # 供 checker 出结论：「目录标题间距」可修，「条目间距」只提示（改样式才是长久之计）。
    toc_detected: bool = False
    toc_entry_count: int = 0
    toc_heading_text: str = ""                          # 目录标题段文本（"Contents" / "目录" …）
    toc_heading_space_before_pt: Optional[float] = None  # 目录标题段段前（磅）；无目录=None
    toc_heading_space_after_pt: Optional[float] = None   # 目录标题段段后（磅）；无目录=None
    toc_entry_space_before_pt: Optional[float] = None    # 目录条目段前（众数，只提示）
    toc_entry_space_after_pt: Optional[float] = None     # 目录条目段后（众数，只提示）
    # —— 脚注 / 尾注计数（v2.3.77）——
    # 本工具**不检查也不修改**脚注/尾注格式（见 report.KNOWN_LIMITATIONS）；
    # 这两个计数只用于给客户一条「你的文档有 N 条脚注，本工具未检查」的提示，
    # 避免客户误以为「报告全绿 = 脚注也看过了」。计数跳过 id<=0 的分隔线条目。
    footnote_count: int = 0
    endnote_count: int = 0
    # —— 软换行（Shift+Enter）拆段计数（v2.3.87）——
    # 读数第一步就把正文级段落里的软换行拆成真实段落（见 split_body_soft_breaks），
    # 这里记录「拆出了几个新段落」，供预览与批注告诉客户「结构动了哪、文字没动」。
    soft_break_splits: int = 0
    heading_styles: list = field(default_factory=list) # 用到的标题样式名
    heading_levels_used: int = 0
    reference_text: str = ""                           # 参考文献段原始文本
    reference_entries: list = field(default_factory=list)
    # 每个条目是否自带 Word 自动编号（`w:numPr`）；与 reference_entries 等长
    reference_entries_numbered: list = field(default_factory=list)
    # 与 reference_entries 一一对应：该条目是否含有任意斜体 run（用于 APA 等风格的
    # 「期刊名/书名应斜体」提示检查）。直接格式化的斜体均可被 python-docx 读到；
    # 仅当整份参考文献「一条斜体都没有」时才作为可疑信号（避免误伤纯网页型条目）。
    reference_entries_italic: list = field(default_factory=list)
    # 参考文献标题（"References" 那一段）的格式探针，用于 APA 等风格的
    # 「标题应粗体+居中」提示检查。注意：官方样例用 Heading 1 样式承载，粗体/居中来自
    # 样式而非直接格式（python-docx 读不出样式解析后的属性），故用 is_style 标记「信任样式」。
    reference_heading_is_style: bool = False
    reference_heading_bold: bool = False               # 直接格式：是否有任意 run 粗体
    reference_heading_centered: bool = False           # 直接格式：段落是否居中
    reference_hanging_indent_in: Optional[float] = None  # 参考文献实测悬挂缩进（英寸）
    # v2.3.65：参考文献条目自己的字体 / 字号 / 行距（条目常与正文不同口径，
    # 如「正文小四双倍、条目五号单倍」）。无参考文献段或读不到 → None（不臆造）。
    reference_font_family: Optional[str] = None
    reference_size_pt: Optional[float] = None
    reference_line_spacing: Optional[float] = None
    reference_space_before_pt: Optional[float] = None   # 条目段前（磅，v2.3.75）
    reference_space_after_pt: Optional[float] = None    # 条目段后（磅，v2.3.75）
    header_text: str = ""                              # 页眉可见文本（剔除 PAGE 域结果，用于 running head 长度/内容检查）
    page_number_present: bool = False                 # 页眉/页脚中是否存在页码（PAGE 域 **或** 手打静态页码）
    page_number_right_aligned: Optional[bool] = None  # 含页码段落是否右对齐（APA 要求 top-right）；无法判定=None
    # —— 页码位置细项（v2.3.66）：区分页眉/页脚 + 有效对齐 + 是否手打 ——
    # 旧版只判「含 PAGE 域的段落是否右对齐」，导致①页脚右对齐被误判为符合 APA top-right
    # （top-right 指页眉右上）②客户手打「第 1 页」被判「没有页码」→⑧补页码时出现两个页码。
    page_number_in: Optional[str] = None              # "header" / "footer" / None
    page_number_alignment: Optional[str] = None       # "left" / "center" / "right" / "both"（已按样式继承归一）
    page_number_is_static: bool = False               # True=手打页码（无 PAGE 域），修正时需原地替换
    page_number_static_text: str = ""                 # 手打页码原文（报告展示 + 前后缀保留）
    header_distance_in: Optional[float] = None       # 页眉距页顶（英寸）；MLA 9th 要求 ½″，python-docx 读 section.header_distance
    reference_new_page: Optional[bool] = None         # 参考文献标题是否另起一页（分页符/段前分页）；无参考文献段=None
    has_abstract: bool = False                        # 文档是否含 Abstract / 摘要 段落（专业版要求）
    has_author_note: bool = False                     # 文档是否含 Author Note / 作者注 段落（专业版要求）
    # —— 摘要页特化（v2.3.67）：摘要正文段的口径与正文不同（APA 摘要不首行缩进、另起一页）——
    abstract_first_line_indent_in: Optional[float] = None  # 摘要正文段实测首行缩进（英寸）；无摘要段=None
    # v2.3.100（C 批 D-08）：摘要正文的字号 / 加粗 / 对齐（IEEE 会议档判定用）
    abstract_font_size_pt: Optional[float] = None
    abstract_bold: Optional[bool] = None
    abstract_alignment: Optional[str] = None
    # v2.3.76：摘要正文段的段前 / 段后（磅）；无摘要段 = None
    abstract_space_before_pt: Optional[float] = None
    abstract_space_after_pt: Optional[float] = None
    abstract_new_page: Optional[bool] = None          # 摘要是否另起一页（段前分页/上一段末尾分页符）；无摘要段=None
    # —— 图表 caption 采样（v2.3.67）——
    caption_count: int = 0                            # 识别到的 Caption 样式段落数
    # v2.3.100（F-5 / 裁定 C-6）：IEEE byline（作者/单位/城市/邮箱那几行）实测。
    # 官方会议模板用 `Author` 样式承载，**样式字号 11pt**（run 上的 9pt 是模板占位，
    # 不代表规范要求）→ 故 byline 字号优先取**样式字号**，客户没用 byline 样式时才取 run。
    byline_size_pt: Optional[float] = None      # byline 字号（磅）；无 byline=None
    byline_line_count: int = 0                  # byline 行数（结构提示用）
    # v2.3.100（D-1 / 裁定 C-11）：官方模板**指导文字残留**的命中位置。
    # 每条：{"where": "body"/"footer"/"header", "index": 段号, "text": 摘要}。
    # 空列表 = 没有残留。IEEE 官方原文：「Failure to remove template text from your
    # paper may result in your paper not being published.」——客户直接交就是"论文不
    # 发表"，而我们此前完全不查。**只报位置，绝不自动删**（删客户文字越权）。
    boilerplate_hits: list = field(default_factory=list)
    # v2.3.100（F-9 / 裁定 C-12 C-13 C-14）：三类「只提示、不改字」的引用/标点观测。
    # 三条都是**改文字**类 → 绝不能挂 FixAction（P-2 文字锁死）。
    # C-12：文内连续引用里的 en dash（IEEE Reference Guide V3.28.2025 已不用）
    citation_range_en_dash: list = field(default_factory=list)
    # C-13：Jr. / Sr. / III 前面的逗号（REF：无逗号）
    suffix_name_comma: list = field(default_factory=list)
    # C-14：全文公式编号出现过哪些形态（只判「是否一致」，不判某种形式为错）
    equation_number_styles: list = field(default_factory=list)
    # v2.3.100（D-2~D-11）：IEEE 系列的**只提示/只判**观测（ESM + OWL + 数学指南）。
    # 每项都只读、不改；涉及"改文字"的一律只提示（见各 checker 维度的说明）。
    ieee_caption_labels: list = field(default_factory=list)   # 表/图标签形态与句点
    ieee_caption_position: dict = field(default_factory=dict)  # 表题在上/图题在下
    ieee_figure_ref_order: list = field(default_factory=list)  # 正文首次引用顺序
    ieee_first_footnote: dict = field(default_factory=dict)    # 首脚注形态
    ieee_acknowledgment: dict = field(default_factory=dict)   # 致谢拼写与位置
    ieee_coi: dict = field(default_factory=dict)               # 利益冲突声明位置
    ieee_biography: dict = field(default_factory=dict)         # 作者传记姓名加粗
    ieee_index_terms: dict = field(default_factory=dict)       # Index Terms 位置
    ieee_abstract_numbered_refs: list = field(default_factory=list)  # 摘要里的编号引用
    ieee_numbered_singular_heads: list = field(default_factory=list) # 被编号的单数标题
    # v2.3.100（E 批）：MLA + APA 剩余的只读观测。
    body_title_repeat: bool = False        # 正文首段是否重复了论文标题（APA）
    numbered_body_heads: list = field(default_factory=list)  # 被编号的正文标题
    author_note_index: Optional[int] = None   # Author Note 段下标（APA 专业版）
    author_note_has_superscript: Optional[bool] = None  # Author Note 里是否有上标编号
    first_page_left_texts: list = field(default_factory=list)  # MLA 首页左上各行文本
    # v2.3.100（E-2 / D-18）：APA caption 的**字形**实测。
    # APA 7 §2.5/§2.6：图表**号**（"Table 1" / "Figure 1"）用**粗体**，
    # 图/表的**题**（描述文字）用**斜体**；注（Note.）两者都不用。
    caption_label_bold: Optional[bool] = None   # 号段是否粗体（众数）
    caption_title_italic: Optional[bool] = None  # 题段是否斜体（众数）
    caption_typography_count: int = 0            # 采到多少组「号+题」
    # v2.3.100（E-6 / D-22）：MLA 论文标题段的实测（居中 / 不粗 / 不斜）。
    # index=None → 定位不到（不判）；certain=False → 定位到但**不能确定是标题**。
    mla_title_index: Optional[int] = None
    mla_title_alignment: Optional[str] = None
    mla_title_bold: Optional[bool] = None
    mla_title_italic: Optional[bool] = None
    mla_title_certain: bool = False
    caption_font_size_pt: Optional[float] = None      # 图表标题实测字号（众数）；无 caption=None
    caption_line_spacing: Optional[float] = None      # 图表标题实测行距（众数）；无 caption=None
    # v2.3.76：图表标题段的段前 / 段后（磅，众数）；无 caption = None
    caption_space_before_pt: Optional[float] = None
    caption_space_after_pt: Optional[float] = None
    # 表题位置：caption 段落与最近表格的相对位置（"top"/"bottom"/None=无表格 caption）
    caption_table_positions: list = field(default_factory=list)
    # —— MLA 专用探测字段（Purdue OWL general_format + 样例论文）——
    title_text: str = ""                              # 首页标题段文本（首个居中段落）
    title_alignment: Optional[str] = None             # 标题对齐标签（center/left/...）
    title_bold: bool = False                          # 标题是否含任意粗体 run
    title_italic: bool = False                        # 标题是否含任意斜体 run
    title_underline: bool = False                     # 标题是否含任意下划线 run
    first_page_left_lines: int = 0                    # 标题之前左上对齐行数（MLA 首页信息块：姓名/教授/课程/日期 共 4 行）
    block_quote_candidates: list = field(default_factory=list)  # 长段落候选：{words, left_indent_in, has_quote}
    # —— 标题页（title page）实测字段（v2.3.65）——
    # 检测规则见 detect_title_page（checker 与 fixer 共用，单一真相源）。
    # 以前标题页改动由 fixer 里一条不受检查驱动的流程悄悄施加：客户看不到、报告里也没有，
    # 违反「检查→修正单一真相源」。现在先采样、由 checker 出结论，才允许修正。
    title_page_detected: bool = False                 # 是否识别到标题页（有 Abstract/Keywords 正面信号）
    tp_title_alignment: Optional[str] = None          # 标题页标题的对齐（center/left/...）
    tp_title_bold: Optional[bool] = None              # 标题页标题是否加粗
    tp_author_alignment: Optional[str] = None         # 作者/单位行的对齐
    # v2.3.76：标题页标题段的段前 / 段后（磅）。无标题页或读不到 → None。
    tp_title_space_before_pt: Optional[float] = None
    tp_title_space_after_pt: Optional[float] = None
    # —— v2.3.91：首页结构（题名页独立成页 / 垂直位置 / 额外段距）——
    # 老板 2026-10-22 核对 v2.3.90 产物后拍板「三件一起做，清掉」：
    # 题名页必须独立成页、标题在顶边距下 3–4 行、题名页整页不留额外段距。
    # 此前只有 title_page（居中/加粗/作者居中），垂直位置与分页完全没人管。
    tp_next_break_before: Optional[bool] = None   # 题名页块之后那段是否另起一页；无题名页=None
    tp_author_space_max_pt: Optional[float] = None  # 作者/单位行段前段后的最大值（磅）；应为 0
    tp_author_gap_paragraphs: Optional[int] = None  # 标题与第一行作者之间夹了几个空段落
    # ↑ v2.3.92：APA 要求「标题与作者之间留一个双倍行距空行」。这个空行有两种合法实现：
    #   ① 真有一个空段落（gap ≥ 1）；② 用标题的段后间距撑出来（gap = 0）。
    #   不区分就会出现「明明有空行却被判缺空行」的误报，所以两种都要能读出来。
    tp_author_block_text: Optional[str] = None
    # ↑ 题名页里**标题以外**那块文本（作者 / 单位 / 课程 / 教师 / 日期），换行拼接。
    #   v2.3.92：APA 学生版题名页还要求课程号+课名、教师名、提交日期 —— 这些内容
    #   我们绝不代客户编造，只能提示。判定必须在标题之外的文本里做：论文大标题本身
    #   就常含 "COVID-19" 这类「字母+数字」，混在一起查会把标题误当成课程号。
    text_break_before: Optional[bool] = None      # 摘要之后正文首段是否另起一页；无摘要=None
    # 摘要**标题**（"Abstract" 这一行）本身的格式（v2.3.97）：
    # 对齐 / 加粗 / 斜体。None = 读不到或没有摘要标题。
    abstract_heading_alignment: Optional[str] = None   # "left" / "center" / "right" / "justify"
    abstract_heading_bold: Optional[bool] = None
    abstract_heading_italic: Optional[bool] = None
    # 关键词行（"Keywords: ..."）的格式（v2.3.97）：标签斜体 + 整行首行缩进。
    # None = 文档里没有关键词行（很多论文确实没有）。
    keywords_label_italic: Optional[bool] = None
    keywords_indent_in: Optional[float] = None
    # 每级标题格式样本（学校模板驱动：层级 → 代表性样本）。仅记录可从 docx 读出的属性；
    # number_prefix 为编号形态的最佳猜测（读不到为 None）。
    heading_formats: dict = field(default_factory=dict)         # level(int) -> HeadingFormatSample
    # heading_formats 是否来自「伪标题」兜底采样（无 Heading 样式、靠整段加粗+字号档识别）。
    # True 时检查报告需说明「层级是推断的」——这是口径说明，不是错误。
    heading_formats_are_pseudo: bool = False
    # v2.3.73：识别到「论文大标题」（首页居中加粗、全篇最大字号）时记下它的下标与字号。
    # 它不属于章节标题层级，检查报告要说清「没把它当标题比」，否则客户看到
    # 「只查到 N 级」会以为漏检。None = 没识别到。
    doc_title_index: Optional[int] = None
    doc_title_size_pt: Optional[float] = None
    # 模板批注（Word comments.xml）—— 海外「学校模板」模式的**一手规范来源**。
    # 每条：{"id": int, "author": str, "text": str, "anchor": int|None}
    # 空列表 = 模板没有批注（此时退回纯实测格式，行为与旧版一致）。
    # 优先级由 questionnaire.target_from_template_docx 决定：批注 > 实测。
    template_comments: list = field(default_factory=list)


@dataclass
class HeadingFormatSample:
    """从 Word 模板抽出的单级标题直接格式化样本（仅可读属性，绝不写 styleId）。"""

    alignment: Optional[str] = None      # left / center / right / justify
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    indent: Optional[bool] = None        # 标题段是否缩进（左缩进或首行缩进 > 0）
    # v2.3.65：缩进的**具体英寸数**（左缩进优先、否则首行缩进）。学校模板里写的是
    # 「标题左缩进 0.5 英寸」这种值，只有一个 bool 表达不了，也无法据此修正。
    indent_in: Optional[float] = None
    # v2.3.65：本级标题自己的字体（run → 段落样式 → 文档默认 三级回退）。
    # 很多学校要求「一级黑体、二级楷体」，只记正文字体会把标题字体统一抹平。
    font_family: Optional[str] = None
    size_pt: Optional[float] = None      # 字号（磅）
    number_prefix: Optional[str] = None  # 编号形态：number/roman/upper_letter/lower_letter/lower_roman/number_paren/None
    uppercase: Optional[bool] = None     # 标题文本是否全大写
    # v2.3.67：段前分页（pageBreakBefore）。读 pPr/w:pageBreakBefore；三态 None=读不到
    # 或未显式设置（段前分页由样式承载时读不到直接值，此时留 None 由 checker 决定不硬判）。
    page_break_before: Optional[bool] = None
    # v2.3.75：段前 / 段后（磅）与行距倍数。学校模板里「章标题段前 12 磅、段后 6 磅、
    # 单倍行距」是最常见的硬要求之一；此前引擎连字段都没有，读了也存不下。
    # 读不到一律按 Word 缺省（0 磅 / 1.0 倍）→ 永不返回 None，避免「假读不到」。
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None
    line_spacing: Optional[float] = None

    def to_dict(self) -> dict:
        return {
            "alignment": self.alignment, "bold": self.bold, "italic": self.italic,
            "indent": self.indent, "indent_in": self.indent_in,
            "font_family": self.font_family,
            "size_pt": self.size_pt,
            "number_prefix": self.number_prefix, "uppercase": self.uppercase,
            "page_break_before": self.page_break_before,
            "space_before_pt": self.space_before_pt,
            "space_after_pt": self.space_after_pt,
            "line_spacing": self.line_spacing,
        }


def _is_quoted_paragraph(text: str) -> bool:
    """段落是不是**以引文开头**的长段落（块引用候选的判定口径，v2.3.86）。

    🔴 旧口径是 `any(q in text for q in ('"', "“", "”", "'"))` ——把**英文撇号**
    当成了引号。真实论文里 "the firm's ability"、"company's"、"don't" 满地都是，
    于是任何 50 词以上、带一个撇号的正文段都被判成「块引用」：
      - 检查器给客户报一条假问题「1 处长引文未按块引用缩进」；
      - 修正器把这些段落从正文缩进里**豁免**掉 → 每章第一段丢了 0.5″ 首行缩进。
    老板那份 20 页样本上一共误伤 4 段（每份各 1 段），这是清单第①条的漏项真因。

    新口径两条同时满足才算：
      ① **段首就是开引号**（" “ ‘ « 「）—— "the firm's" 这种永远不可能满足；
      ② 段内确实存在收引号（" ” ’ » 」），排除「以引号开头但其实是排版残留」的情况。
    句中的撇号、句中夹一段引文都不再触发；而「整段被引起来」「引文开头 + 出处说明」
    这两种真实的块引用形态都还能认出来（MLA 的 (Danhof 5) 那种）。
    """
    t = (text or "").strip()
    if len(t) < 2:
        return False
    openers = {'"', "“", "‘", "«", "「", "'"}
    closers = {'"', "”", "’", "»", "」", "'"}
    if t[0] not in openers:
        return False
    return any(ch in closers for ch in t[1:])


def _inch(emu) -> Optional[float]:
    if emu is None:
        return None
    try:
        return round(float(emu.inches), 3)
    except Exception:
        return None


def _pt(length) -> Optional[float]:
    if length is None:
        return None
    try:
        return round(float(length.pt), 1)
    except Exception:
        return None


def _run_font_name(run) -> Optional[str]:
    """读 run 的字体名：run.font.name（映射 w:ascii/hAnsi）→ w:rFonts 各属性兜底。

    中文 Word 文档常见「只设中文字体」(w:eastAsia) 的 run，只读 font.name 会得到
    None 并错误回落到 Normal → 字体误判或白报「未能读取」。
    """
    try:
        n = run.font.name
    except Exception:
        n = None
    if n:
        return n
    try:
        rPr = run._r.rPr
        if rPr is None:
            return None
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            return None
        for attr in _FONT_ATTRS:
            v = rFonts.get(qn(attr))
            if v:
                return v
    except Exception:
        pass
    return None


def _line_spacing_of(pf) -> Optional[float]:
    """从段落格式推断行距倍数（确定性）。

    python-docx 在「固定值/最小值」(EXACTLY/AT_LEAST) 时，line_spacing
    返回的是 Length(EMU)，不能直接当倍数。此时返回 None，交由 checker 以
    「未能读取」处理，避免误判 fail。
    仅当规则为 MULTIPLE（或明确的单/1.5/双倍命名规则）才返回倍数。
    """
    if pf is None:
        return None
    try:
        rule = pf.line_spacing_rule
    except Exception:
        rule = None
    if rule == WD_LINE_SPACING.MULTIPLE:
        try:
            return round(float(pf.line_spacing), 3)
        except Exception:
            return None
    # 命名规则可近似为倍数
    named = {
        WD_LINE_SPACING.DOUBLE: 2.0,
        WD_LINE_SPACING.ONE_POINT_FIVE: 1.5,
        WD_LINE_SPACING.SINGLE: 1.0,
    }
    if rule in named:
        return named[rule]
    # EXACTLY / AT_LEAST / None：无法可靠转倍数 → 返回 None
    return None


def _raw_line_pt(para):
    """从 段落 → 样式链 取 ``w:spacing/@w:line`` 的原始值与 rule。

    为什么必须直读 XML：IEEE 会议模板把 ``w:line="11.4pt" w:lineRule="auto"`` 写在
    **BodyText 样式**上，python-docx 会按「1/240 行」把它算成 0.95 倍（比字号还小，
    渲染上不可能），而作者的本意是"11.4pt 行距"。要判定是否合规必须能看见原始单位。
    """
    def _from_ppr(ppr):
        if ppr is None:
            return None
        sp = ppr.find(qn("w:spacing"))
        if sp is None:
            return None
        raw = sp.get(qn("w:line"))
        if raw is None:
            return None
        return (raw, sp.get(qn("w:lineRule")))
    try:
        hit = _from_ppr(para._p.find(qn("w:pPr")))
        if hit:
            return hit
    except Exception:
        pass
    try:
        st = para.style
    except Exception:
        st = None
    while st is not None:
        try:
            el = getattr(st, "element", None)
            hit = _from_ppr(el.find(qn("w:pPr")) if el is not None else None)
            if hit:
                return hit
        except Exception:
            pass
        st = getattr(st, "base_style", None)
    return None


def _doc_default_size_pt(doc) -> Optional[float]:
    """文档默认字号（``styles.xml`` → ``w:docDefaults/w:rPrDefault/w:rPr/w:sz``，磅）。

    ⚠️ ``doc.styles["Normal"].font.size`` **不等于** docDefaults —— Normal 样式没写 sz 时
    它返回 None，而真正决定继承字号的是 docDefaults。旧代码把两者混用，导致
    "字号读不到 → 硬编码 10pt 兜底"，把 10pt 混进正文众数（MLA 实测众数变成 10.0）。
    """
    try:
        root = doc.styles.element
        dd = root.find(qn("w:docDefaults"))
        if dd is None:
            return None
        rprd = dd.find(qn("w:rPrDefault"))
        if rprd is None:
            return None
        rpr = rprd.find(qn("w:rPr"))
        if rpr is None:
            return None
        sz = rpr.find(qn("w:sz"))
        if sz is None:
            return None
        return round(float(sz.get(qn("w:val"))) / 2.0, 1)
    except Exception:
        return None


def _style_chain_size_pt(para) -> Optional[float]:
    """沿段落样式链（含 base_style）找字号。"""
    try:
        st = para.style
    except Exception:
        return None
    while st is not None:
        try:
            if st.font is not None and st.font.size is not None:
                return round(float(st.font.size.pt), 1)
        except Exception:
            pass
        st = getattr(st, "base_style", None)
    return None


def _font_size_hint(para) -> Optional[float]:
    """该段的字号提示（run → 样式链），仅用于把「pt 行距」换算成倍数。"""
    try:
        for r in para.runs:
            if (r.text or "").strip() and r.font.size is not None:
                return round(float(r.font.size.pt), 1)
    except Exception:
        pass
    try:
        st = para.style
    except Exception:
        st = None
    while st is not None:
        try:
            if st.font is not None and st.font.size is not None:
                return round(float(st.font.size.pt), 1)
        except Exception:
            pass
        st = getattr(st, "base_style", None)
    return None


def _line_pt_multiple(para) -> Optional[float]:
    """``w:line`` 带 **pt 单位** 且 ``lineRule=auto`` 时，按 ``pt ÷ 字号`` 换算倍数。

    IEEE 会议模板实测：BodyText 样式 ``w:line="11.4pt" lineRule="auto"``，字号 10pt
    → 1.14 倍（C 批裁定 D-13 的口径）。不带 pt 的普通写法（``w:line="480"``）不在此
    处理，仍走 python-docx 的老路径（480/240 = 2.0 倍），影响面为零。
    """
    hit = _raw_line_pt(para)
    if not hit:
        return None
    raw, rule = hit
    s = str(raw).strip()
    if not s.lower().endswith("pt"):
        return None
    if rule not in (None, "", "auto"):
        return None
    try:
        pt = float(s[:-2])
    except (TypeError, ValueError):
        return None
    size = _font_size_hint(para) or 10.0
    if size <= 0:
        return None
    return round(pt / size, 2)


def _line_spacing_of_para(para, doc_pf=None) -> Optional[float]:
    """行距倍数，带**样式继承**回退（段落直接 → 段落样式链 → 文档默认/Normal）。

    v2.3.71 客户视角实测发现：绝大多数 Word 文档与学校模板把行距写在 **Normal 样式**
    上（模板制作者的正常做法），段落本身没有直接格式。旧实现只读
    ``para.paragraph_format``，于是行距一律读成 None —— 检查器的「行距」维度整个
    失效：把单倍行距的论文交给本工具，报告照样全绿（假通过，且是我们最该管的一项）。

    语义约定（与「没写 w:jc 就是左对齐」同口径）：
    - 段落自己是「固定值/最小值」→ 倍数无意义，如实返回 None（**不再往样式回退**，
      否则会把「固定值 20 磅」误读成样式里的双倍）；
    - 整条链都没写行距 → Word 渲染就是单倍，返回 1.0（不是 None）。
    """
    pf0 = None
    try:
        pf0 = para.paragraph_format
    except Exception:
        pf0 = None
    # v2.3.100：pt 单位的 auto 行距（IEEE 模板写法）先换算成倍数，见 _line_pt_multiple
    try:
        _pt_mul = _line_pt_multiple(para)
        if _pt_mul is not None:
            return _pt_mul
    except Exception:
        pass
    if pf0 is not None:
        try:
            rule0 = pf0.line_spacing_rule
        except Exception:
            rule0 = None
        if rule0 in (WD_LINE_SPACING.EXACTLY, WD_LINE_SPACING.AT_LEAST):
            return None

    candidates: list = []
    if pf0 is not None:
        candidates.append(pf0)
    try:
        st = para.style
    except Exception:
        st = None
    while st is not None:                      # 沿 base_style 一路向上（Heading→Normal）
        try:
            candidates.append(st.paragraph_format)
        except Exception:
            pass
        st = getattr(st, "base_style", None)
    if doc_pf is not None:
        candidates.append(doc_pf)
    for pf in candidates:
        v = _line_spacing_of(pf)
        if v is not None:
            return v
    return 1.0


def _heading_level_from_name(name: Optional[str]) -> Optional[int]:
    """按样式名推断标题层级。

    两类：
    1) 「关键词 + 数字」形态：Heading 1 / 标题 2 / Kop 3 / Überschrift 4；
    2) 「整段标题」形态（无编号）：Title / Subtitle / 题目 … → 视为第 1 级。
    """
    if not name:
        return None
    norm = re.sub(r"\s+", "", name).lower()
    for kw in _HEADING_KEYWORDS:
        if norm.startswith(kw):
            digits = re.search(r"\d+", norm[len(kw):])
            if digits:
                n = int(digits.group())
                if 1 <= n <= 9:
                    return n
    for kw in _TITLE_STYLE_KEYWORDS:
        key = re.sub(r"\s+", "", kw).lower()
        if key and norm.startswith(key):
            return 1
    return None


# 「文档标题」样式的关键词（Title / 题目 / Titel / Título …），**不含 Subtitle**：
# 副标题不是论文大标题，不能顶替首页标题的位置。
_DOC_TITLE_STYLE_KEYS = tuple(
    k for k in _TITLE_STYLE_KEYWORDS if "sub" not in k.lower()
)


def _is_doc_title_style_name(name: Optional[str]) -> bool:
    """样式名是否属于「文档标题」类（Word 内置 Title / 题目 / Titel …）。

    为什么不复用 ``_heading_level_from_name``：后者故意把 Title 视为 **1 级标题**
    （保护主标题不被压成正文，见 tests/test_review_fixes.py::test_title_style_is_heading）。
    但「是标题」与「是章节标题层级」是两件事 —— 论文大标题绝不能占掉 L1 字号档。
    """
    if not name:
        return False
    norm = re.sub(r"\s+", "", name).lower()
    for kw in _DOC_TITLE_STYLE_KEYS:
        key = re.sub(r"\s+", "", kw).lower()
        if key and norm.startswith(key):
            return True
    return False


def _style_outline_level(para) -> Optional[int]:
    """样式定义里的 outlineLvl。

    自定义标题样式（样式名不含 heading 关键词）常把 outlineLvl 写在**样式层**
    而不是段落层，只读段落 pPr 会漏判 —— 既漏保护，也漏统计标题层级。
    """
    try:
        style = para.style
        if style is None:
            return None
        el = getattr(style, "element", None)
        if el is None:
            return None
        pPr = el.find(qn("w:pPr"))
        if pPr is None:
            return None
        ol = pPr.find(qn("w:outlineLvl"))
        if ol is None:
            return None
        return int(ol.get(qn("w:val"))) + 1
    except Exception:
        return None


def _heading_level(para) -> Optional[int]:
    """标题层级：样式名 → 段落级 outlineLvl → 样式级 outlineLvl（兼容自定义样式）。"""
    lvl = _heading_level_from_name(para.style.name if para.style else None)
    if lvl:
        return lvl
    try:
        pPr = para._p.pPr
        if pPr is not None:
            ol = pPr.find(qn("w:outlineLvl"))
            if ol is not None:
                return int(ol.get(qn("w:val"))) + 1  # outlineLvl 0 = 第 1 级
    except Exception:
        pass
    return _style_outline_level(para)


def _max_run_size_pt(para) -> Optional[float]:
    best: Optional[float] = None
    try:
        runs = list(para.runs)
    except Exception:
        return None
    for r in runs:
        try:
            sz = r.font.size
        except Exception:
            sz = None
        if sz is None:
            continue
        try:
            pt = float(sz.pt)
        except Exception:
            continue
        if best is None or pt > best:
            best = pt
    return best


def _detect_number_prefix(text: str) -> Optional[str]:
    """最佳猜测标题的编号前缀形态（仅首段文本，读不到返回 None）。

    返回值与 HeadingLevelFormat.number_prefix 保持一致（fixer 可复用）：
    number / roman / upper_letter / lower_letter / lower_roman / number_paren。
    """
    import re as _re
    s = (text or "").strip()
    if not s:
        return None
    # 多级阿拉伯数字（1.1 / 1.1.）优先于单级（1.）
    if _re.match(r"^\d+(?:\.\d+)+\.?\s", s):
        return "number"
    if _re.match(r"^\d+\)\s", s):
        return "number_paren"
    if _re.match(r"^\d+\.\s", s):
        return "number"
    if _re.match(r"^[IVXLCDM]+\.\s", s):
        return "roman"
    if _re.match(r"^[A-Z]\.\s", s):
        return "upper_letter"
    if _re.match(r"^[ivxlcdm]+\.\s", s):
        return "lower_roman"
    if _re.match(r"^[a-z]\)\s", s):
        return "lower_letter"
    if _re.match(r"^\(?[a-z]\)?\s", s):
        return "lower_letter"
    return None


def _sample_heading_format(para, doc_pf, doc_default_font=None) -> "HeadingFormatSample":
    """采样一个代表性标题段落的直接格式化（用于学校模板驱动）。

    对齐 / 加粗 / 斜体 / 字体 / 字号均三级回退（run → 段落样式 → 文档默认），
    与正文采样口径一致；编号前缀 / 全大写按标题文本启发式识别。
    """
    # 对齐
    al_raw = para.alignment
    al = _ALIGN_MAP.get(al_raw) if al_raw is not None else None
    if al is None:
        al_raw2 = _pf_attr(para, doc_pf, "alignment")
        al = _ALIGN_MAP.get(al_raw2) if al_raw2 is not None else None
    # v2.3.70：段落 / 样式链 / 文档默认**都没写 w:jc** 时，Word 的渲染结果就是左对齐 ——
    # 这是 OOXML 语义（缺省 jc = left），不是「读不到」。旧写法留 None，造成三处不一致：
    #   ① 检查器对正常标题报 "Could not read: alignment"（等于检查有洞）；
    #   ② 预览因 got.alignment is None 一个字都不列；
    #   ③ 修正器却照样把对齐改掉 → 客户在报告里看到「读不到」，文档却被悄悄改了。
    # 与页码对齐的口径一致：样式继承解析完后**永不返回 None**，否则出假通过/假空白。
    if al is None:
        al = "left"

    # 加粗 / 斜体：run 全部显式 True / 全部显式 False 才下结论；有未显式(None)或
    # 混合时回退段落样式。旧写法用 `all(x is True)`，把「未显式设置」直接判成 False，
    # 于是「粗体由 Heading 样式承载」的模板被读成「不加粗」——新加的每级标题检查会
    # 据此误报（v2.3.65 修正）。
    runs = [r for r in para.runs if (r.text or "").strip()]
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    if runs:
        b_raw = [getattr(r.font, "bold", None) for r in runs]
        i_raw = [getattr(r.font, "italic", None) for r in runs]
        if all(v is True for v in b_raw):
            bold = True
        elif all(v is False for v in b_raw):
            bold = False
        if all(v is True for v in i_raw):
            italic = True
        elif all(v is False for v in i_raw):
            italic = False
    if bold is None or italic is None:
        try:
            st = para.style
            if st is not None and st.font is not None:
                if bold is None and st.font.bold is not None:
                    bold = st.font.bold is True
                if italic is None and st.font.italic is not None:
                    italic = st.font.italic is True
        except Exception:
            pass

    # 字号：run 显式值优先（直接格式覆盖样式，这才是用户看到的值），读不到再回退样式。
    # 旧写法反过来（先样式），模板 run 写着 16pt、样式 14pt 时会读成 14pt。
    size: Optional[float] = None
    size = _max_run_size_pt(para)
    if size is None:
        try:
            st = para.style
            if st is not None and st.font is not None and st.font.size is not None:
                size = round(float(st.font.size.pt), 1)
        except Exception:
            size = None

    # 字体：run → 段落样式 → 文档默认（与正文口径一致）。取第一个有字体的 run
    # 作代表 —— 标题段通常整段同字体，混合时按首个非空 run 记为样本，不臆造。
    font: Optional[str] = None
    for r in runs:
        fn = _run_font_name(r)
        if fn:
            font = fn
            break
    if font is None:
        try:
            st = para.style
            if st is not None and st.font is not None and st.font.name:
                font = st.font.name
        except Exception:
            pass
    if font is None and doc_default_font:
        font = doc_default_font

    # 缩进：左缩进或首行缩进 > 0 视为缩进式标题；同时记下具体英寸数
    li = _inch(_pf_attr(para, doc_pf, "left_indent"))
    fli = _inch(_pf_attr(para, doc_pf, "first_line_indent"))
    indent_in: Optional[float] = None
    if li is not None and li > 0:
        indent_in = li
    elif fli is not None and fli > 0:
        indent_in = fli
    indent = indent_in is not None

    # 全大写：只看**字母**的大小写。
    # 旧写法 `txt.replace(" ", "").isalpha()` 会把带编号/标点的标题（IEEE 的
    # "I. INTRODUCTION"、"A. BACKGROUND"）判成「读不到」——而那恰恰是最需要判大写的场景，
    # 于是修完之后复检依旧「读不到」，永远无法收敛（v2.3.70 修正）。
    txt = para.text.strip()
    uppercase: Optional[bool] = None
    _letters = [ch for ch in txt if ch.isalpha()]
    if _letters:
        # 汉字等无大小写的字符 islower() 恒为 False，不干扰判定
        uppercase = all(not ch.islower() for ch in _letters)

    # 段前分页（v2.3.67）：读段落直接格式 pPr/pageBreakBefore；读不到（None 或样式承载）
    # 保持 None，由 checker 决定不硬判。三态：True/False/None。
    pbb: Optional[bool] = None
    try:
        _pf_pbb = getattr(para.paragraph_format, "page_break_before", None)
        if _pf_pbb is not None:
            pbb = bool(_pf_pbb)
    except Exception:
        pbb = None

    # 编号前缀：启发式识别标题起始编号形态
    number_prefix = _detect_number_prefix(txt)

    # 段前 / 段后 / 行距（v2.3.75）：三级回退 + 「行」单位换算。
    # 段前段后走 _spacing_pair_of_para（直读 XML，认 w:beforeLines）；行距走
    # _line_spacing_of_para（固定值/最小值如实返回 None，由 checker 记「读不到」）。
    sb, sa = _spacing_pair_of_para(para, base_size_pt=size)
    ls = _line_spacing_of_para(para)

    return HeadingFormatSample(
        alignment=al, bold=bold, italic=italic, indent=indent, indent_in=indent_in,
        font_family=font, size_pt=size, number_prefix=number_prefix, uppercase=uppercase,
        page_break_before=pbb,
        space_before_pt=sb, space_after_pt=sa, line_spacing=ls,
    )


def is_pseudo_heading(para, min_pt: Optional[float] = None) -> bool:
    """伪标题判定：无标题样式/outline 标记，但形似标题（短 + 整段加粗 + 字号大于正文）。

    规格第 10 条要求保护「手写大字号标题」。判定刻意保守（整段显式加粗 +
    ≤80 字 + 不以句读结尾），宁可少改一段，也不能把主标题压成正文。

    ``min_pt`` 控制字号下限：
    - 默认 ``_PSEUDO_HEADING_MIN_PT``(13pt) 用于*保护*用途（只判断「这是不是标题，
      别把它当正文刷」），刻意高阈值避免误伤；
    - 当目标规范显式声明 heading_levels>0（用户就是要给标题套编号/格式）时，
      fixer 用更低的阈值（见 _PSEUDO_HEADING_FORMAT_MIN_PT）来*识别*正文字号加粗的
      章节标题（IEEE/APA 常见 11–12pt 加粗标题，没有 13pt 那么大）。
    """
    try:
        text = para.text.strip()
    except Exception:
        return False
    if not text or len(text) > _PSEUDO_HEADING_MAX_LEN:
        return False
    if text.endswith((".", "。", ",", "，", ";", "；", ":", "：", "?", "？", "!", "！")):
        return False
    try:
        runs = [r for r in para.runs if (r.text or "").strip()]
    except Exception:
        return False
    if not runs:
        return False
    # v2.3.104（F8 修复）：符号 / 数学字体（Cambria Math / Symbol / …）的段不是标题。
    # 真实论文里用 Cambria Math 居中的公式（如 `MTTR = β₀ + … + ε`）整段加粗、短、
    # 不以句读结尾，会被误判成伪标题并被加编号 + run-in 冒号 —— 违反文字锁死铁律
    # （实测：输出变成 `... + ε: ` 并在段尾插批注）。公式段一律豁免。
    if any(_is_symbol_font(_run_font_name(r)) for r in runs):
        return False
    # 兜底：用普通字体手打的公式行（含 `=` 且含希腊字母）同样不是标题形状。
    if _looks_like_equation(text):
        return False
    for r in runs:
        try:
            if r.bold is not True:
                return False
        except Exception:
            return False
    threshold = _PSEUDO_HEADING_MIN_PT if min_pt is None else min_pt
    sz = _max_run_size_pt(para)
    return sz is not None and sz >= threshold


_GREEK_LETTERS = set(
    "αβγδεζηθικλμνξοπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ"
)


def _looks_like_equation(text: str) -> bool:
    """粗略识别「公式文本」（非 OMath、用普通字体手打的公式行）。

    含 `=` 且含希腊字母（如 `β`/`ε`）即判定为公式 —— 标题段几乎不会同时出现
    这两个特征，误伤正文 / 标题的风险极低；反过来公式是论文里最常见的
    「居中加粗短段」伪标题误判源，必须挡掉（见 v2.3.104 F8）。
    """
    return "=" in text and any(ch in _GREEK_LETTERS for ch in text)


def is_heading(para) -> bool:
    """公开的标题判定（fixer / 测试复用，避免各自重写判定逻辑）。

    命中样式/outline，或形似标题（伪标题）即视为标题 —— 标题段的一切格式维度
    在一键修正中都不允许改动。
    """
    return _heading_level(para) is not None or is_pseudo_heading(para)


# ——「章节标题行」判定（v2.3.87）——
# 真实学生论文的章节标题常常是一行纯文本：无标题样式、不加粗、与正文同字号，
# 只靠编号表明身份（"1. Introduction" / "2.1 Sample" / "一、引言"）。
# `is_pseudo_heading` 认不出它们（它要求整段加粗且字号更大），于是这些行会被当成
# 正文套上 0.5″ 首行缩进 —— 客户看到的就是「章节标题缩进了，正文反倒没缩进」。
# 本判定只回答「这一行是不是标题形状」，刻意不判断字号/加粗：
# 条件要求「编号开头 + 短 + 不以句读结尾」，普通正文段落不会同时满足。
_SECTION_HEADING_RE = re.compile(
    r"^\s*(?:"
    r"\d+(?:\.\d+)*\.?\s+\S"                    # 1. Introduction / 2.1 Sample / 3.2.1 Detail
    r"|[IVXLC]+\.\s+\S"                         # I. / IV.
    r"|[A-Z]\.\s+\S"                            # A. / B.
    r"|[一二三四五六七八九十]+[、.．]\s*\S"        # 一、引言 / 二.方法
    r"|(?i:chapter|section|part)\s+(?:\d+|[IVXLC]+)\b"   # Chapter 3 / Part II
    r")"
)
_SENTENCE_END = (".", "。", "?", "？", "!", "！", ";", "；", ",", "，", "：")


def is_section_heading_line(para) -> bool:
    """段落是否是「只有编号、没有加粗/大字号」的章节标题行（v2.3.87）。

    用途很窄：只给**首行缩进**这一个维度做豁免 —— 标题行不该像正文那样缩进半英寸。
    行距等「整篇属性」照常施加（APA §2.21 要求整篇双倍行距，标题也在内），
    所以本函数不参与行距判定。

    判定条件（三条全中才算，宁可漏判也不误伤正文）：
      ① 编号开头（见 ``_SECTION_HEADING_RE``）；
      ② 短（≤ ``_PSEUDO_HEADING_MAX_LEN`` 字）；
      ③ 不以任何句读结尾（分号/逗号/冒号/问号/感叹号/句号都不行）。
    普通正文段落几乎不可能三条全中：它会长、而且以句号结尾。
    """
    try:
        text = (para.text or "").strip()
    except Exception:
        return False
    if not text or len(text) > _PSEUDO_HEADING_MAX_LEN:
        return False
    if text.endswith(_SENTENCE_END):
        return False
    # v2.3.104：公式行（含 `=` 且含希腊字母）一律不是标题行，避免 `1. f(x) = β₀`
    # 这类编号公式被当成章节标题行（见 _looks_like_equation）。
    if _looks_like_equation(text):
        return False
    return bool(_SECTION_HEADING_RE.match(text))


def is_standalone_heading_shape(para) -> bool:
    """段落是否呈「独立标题行」形状：短（≤ ``_PSEUDO_HEADING_MAX_LEN`` 字）+ 不以句读结尾。

    v2.3.100（幂等收尾）：判定**刻意不依赖加粗 / 居中 / 编号** —— 这三样都会被修正器
    写掉：IEEE 一级标题按裁定 C-5 是「不加粗 + 居中 + small caps」，于是第一遍修好后
    第二遍 ``is_pseudo_heading``（要求整段显式加粗）就不再命中它，12pt 的标题重新混进
    正文字号众数 → 众数被判 12pt、而修正侧只改正文 → 每遍都判「字号不符」、
    每遍都记 applied → 客户点完「修正」那条红杠永远不消失
    （实测 ``test_fix_ieee_realistic_doc_is_idempotent`` 第二遍 applied 非空）。
    判据只保留「短 + 无句读结尾」这两个修正器改不掉的形状特征。
    """
    try:
        text = (para.text or "").strip()
    except Exception:
        return False
    if not text or len(text) > _PSEUDO_HEADING_MAX_LEN:
        return False
    if text.endswith(_SENTENCE_END):
        return False
    return True


def heading_level(para) -> Optional[int]:
    """公开的标题层级访问器（fixer 套用每级格式规则用）。

    返回 1-based 层级（1=最高级）或 None（非标题 / 伪标题无明确层级）。
    判定链路：样式名 → 段落级 outlineLvl → 样式级 outlineLvl，兼容自定义标题样式。
    """
    return _heading_level(para)


# ---------------------------------------------------------------------------
# 编号结构行 → 标题层级（v2.3.97）
#
# 老板 2026-10-23 实测指出两处，都在这条链上：
#   ① APA Level 2 = **Flush Left, Bold, Title Case**（Table 2.3, p.48）→ 要加粗；
#   ② 真实论文里 "1. Introduction" / "2. Literature Review" 这类小节行既没加粗、
#      也压根没被认成标题（v2.3.87 只让它「豁免首行缩进」，格式一个字节没动），
#      于是报告不提、修正不改 —— 而「各层级标题的修正」正是格式修正的主干。
#
# 🔴 控制原则：**宁可漏判，绝不误伤正文**。四道闸门全过才认：
#   ① 先过 `is_section_heading_line`（编号开头 + 够短 + 不以任何句读结尾）；
#   ② 同一层级上的编号必须构成**从 1 开始、严格递增、至少两项**的序列 ——
#      正文里那种 "1. First step / 2. Second step" 的列表偶尔也短、也带编号，
#      但很少同时做到「从 1 连续、互不重复、彼此都是整行编号」；
#   ③ 层级**自顶向下**成立：第 1 层的编号序列不成立时，更深层一律不认
#      （不跳级认人 —— 没有顶层序列就说明这不是章节编号体系）；
#   ④ 认出的层级 = 已有锚点层级 + 编号深度（无锚点则 = 编号深度），上限 5 级。
#      ④ 是给「文档里已有 `Chapter 1/2/3` 加粗大章」这种结构用的：
#      此时 "1. Introduction" 必须排在它**下面**一级，不能跟大章抢 L1。
# ---------------------------------------------------------------------------

_ORD_CN = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
           "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}
_ROMAN_VAL = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100}


def _roman_to_int(s: str) -> Optional[int]:
    """罗马数字转整数（只支持 I/V/X/L/C 这套章节常用档，非法返回 None）。"""
    s = (s or "").upper()
    if not s or any(ch not in _ROMAN_VAL for ch in s):
        return None
    total, prev = 0, 0
    for ch in reversed(s):
        val = _ROMAN_VAL[ch]
        total = total - val if val < prev else total + val
        prev = max(prev, val)
    return total or None


def _cn_to_int(s: str) -> Optional[int]:
    """中文数字转整数（覆盖「一」到「九十九」这类章节编号，够用即可）。"""
    if not s:
        return None
    if s == "十":
        return 10
    if "十" in s:
        head, _, tail = s.partition("十")
        tens = _ORD_CN.get(head, 1) if head else 1
        ones = _ORD_CN.get(tail, 0) if tail else 0
        return tens * 10 + ones
    return _ORD_CN.get(s)


def parse_section_number(text: str) -> Optional[tuple]:
    """章节编号 → 数字元组；解析不出返回 None。

    "1. Introduction" → (1,)      "2.1 Sample" → (2, 1)     "3.2.1 Detail" → (3, 2, 1)
    "Chapter 3 ..."  → (3,)       "IV. Method" → (4,)      "A. Framework" → (1,)
    "一、引言"        → (1,)

    元组长度即**编号深度**：长度 1 = 顶层小节，长度 2 = 它的子节 …
    """
    t = (text or "").strip()
    if not t:
        return None
    m = re.match(r"^(\d+(?:\.\d+)*)\.?\s+\S", t)
    if m:
        return tuple(int(x) for x in m.group(1).split("."))
    m = re.match(r"(?i)^(?:chapter|section|part)\s+(\d+)\b", t)
    if m:
        return (int(m.group(1)),)
    m = re.match(r"(?i)^(?:chapter|section|part)\s+([IVXLC]+)\b", t)
    if m:
        v = _roman_to_int(m.group(1))
        return (v,) if v else None
    m = re.match(r"^([IVXLC]+)\.\s+\S", t)
    if m:
        v = _roman_to_int(m.group(1))
        return (v,) if v else None
    m = re.match(r"^([A-Z])\.\s+\S", t)
    if m:
        return (ord(m.group(1)) - 64,)
    m = re.match(r"^([一二三四五六七八九十]+)[、.．]", t)
    if m:
        v = _cn_to_int(m.group(1))
        return (v,) if v else None
    return None


def _norm_matter_title(text: str) -> str:
    """归一化标题文字（去空格、小写、去尾部冒号），用于比对前置/后置材料标题。"""
    t = (text or "").strip().lower().rstrip(":：")
    return t.replace(" ", "")


# 前置 / 后置**材料**的标题：它们是「文档部件标题」，不属于章节层级里的某一级。
# 🔴 v2.3.97 实测链条：`Abstract` 段因「整段加粗 + 字号居中档」被伪标题兜底采样认成
# **二级章节标题**，随后被套上 APA Level 2 的「左对齐 + 加粗」—— 而 APA §2.9 要的是
# 「**居中** + 加粗」。老板 2026-10-23 打开产物一眼看出「摘要标题没居中」，根子在这里。
# 它们各归各管：摘要标题 → abstract_heading 维度；参考文献标题 → reference_heading 维度。
_MATTER_TITLES = {
    "abstract", "摘要", "abstractpage",
    "references", "reference", "bibliography", "workscited", "参考文献",
    "acknowledgments", "acknowledgements", "acknowledgment", "acknowledgement", "致谢",
    "appendix", "appendices", "附录",
    "keywords", "关键词",
    "tableofcontents", "contents", "目录",
}


def structural_heading_map(paras, min_pt: Optional[float] = None,
                           exclude_ps: Optional[set] = None) -> dict:
    """文档里所有可认出层级的标题段 → {w:p 元素: level}（v2.3.97）。

    ⚠️ **单一真相源**：reader 的每级格式采样与 fixer 的层级指派共用本函数。
    以前只有 `pseudo_heading_level_map`（整段加粗 + 字号分层），于是「不加粗的编号
    小节行」两头落空 —— 采样看不见它们（L2 要求无从比对），修正也不碰它们。

    组成：`pseudo_heading_level_map` 的全部结果 + 通过四道闸门的编号结构行。
    闸门细节见本函数上方的注释；本函数**只负责认，不负责改**。
    """
    paras = list(paras)
    exclude_ps = exclude_ps or set()
    out = dict(pseudo_heading_level_map(paras, min_pt=min_pt, exclude_ps=exclude_ps))

    # 前置 / 后置材料标题（Abstract / References / 致谢 / 附录 …）**不进章节层级**：
    # 它们一旦占掉某一级，既会把真正的章节标题挤到更深一级（`1. Introduction` 被认成
    # L3、套上 APA Level 3 的「加粗+斜体」），又会被套成错的格式（摘要标题被套成
    # Level 2 的左对齐）。见 _MATTER_TITLES 的说明。
    _text_by_p = {para._p: (para.text or "") for para in paras}
    _kept = {p: lv for p, lv in out.items()
             if _norm_matter_title(_text_by_p.get(p, "")) not in _MATTER_TITLES}
    # 🔴 v2.3.100：过滤后**紧凑重编号**（消除跳级）。
    # 旧写法直接留下空洞：一篇论文里「Abstract」若是 12pt 加粗，它会占掉字号档的 L2，
    # 而它作为材料标题又被上面过滤掉 → 真实的 11pt 小节行从 **L3** 起步，
    # 于是分层变成「有 L1、有 L3、没有 L2」的跳级形态。
    # 后果（实测）：APA Level 3 要求斜体，而修正器第一遍只修**检查当时报出的**层级
    # （原稿还没有 L3，只报了 L1/L2）→ 修完层级漂移 → 复检报「Level 3 not italic」
    # → 第二遍又去改 24 个标题（headings_formatted=24、comments_added=24），
    # 客户点一次「修正」以为是干净的，其实第二遍还在动文档。
    # 紧凑化后层级连续，漂移从根上消失。分层本就是推断值，连续比保真跳级更重要。
    if _kept:
        _order = sorted(set(_kept.values()))
        _compact = {old: i + 1 for i, old in enumerate(_order)}
        out = {p: _compact[lv] for p, lv in _kept.items()}
    else:
        out = {}

    # 已有锚点的最高层级：编号小节要排在它下面，而不是跟它抢同一级。
    base = max(out.values()) if out else 0

    cands: list = []          # [(index, para, nums)]
    for i, para in enumerate(paras):
        if para._p in exclude_ps or para._p in out:
            continue
        if not is_section_heading_line(para):
            continue
        nums = parse_section_number(para.text or "")
        if nums is None:
            continue
        cands.append((i, para, nums))
    if not cands:
        return out

    by_depth: dict = {}
    for row in cands:
        by_depth.setdefault(len(row[2]), []).append(row)

    # 闸门 ②③：顶层（深度 1）必须构成**可信的编号序列**，且至少两项。
    #   合法形态 = 每一项「比上一项大 1」**或**「重新从 1 开始」——
    #   后者是真实论文的常态：多数学校「每章小节各自从 1 编号」（Ch1→1., Ch2→1.），
    #   整篇按「第 1 章 1. 2. 3. / 第 2 章 1. 2.」走。要求全局严格递增会把这类
    #   文档整批漏掉（实测：老板这份 20 页论文每 6 章循环一次，严格递增直接判空）。
    #   反过来，正文里的编号列表一旦跳号（1. 2. 5.）就不认 —— 宁可漏，不误伤。
    top = by_depth.get(1, [])
    seq = [n[0] for _, _, n in top]
    if len(top) < 2 or seq[0] != 1:
        return out
    for a, b in zip(seq, seq[1:]):
        if b != a + 1 and b != 1:
            return out
    for _i, para, _n in top:
        out[para._p] = min(5, base + 1)

    # 更深层的编号：其首段编号必须落在顶层的章号里（不跳级认人）
    top_prefixes = set(seq)
    for depth in sorted(d for d in by_depth if d > 1):
        rows = [r for r in by_depth[depth] if r[2][0] in top_prefixes]
        if not rows:
            continue
        for _i, para, _n in rows:
            out[para._p] = min(5, base + depth)
    return out


def pseudo_heading_level_map(paras, min_pt: Optional[float] = None,
                             exclude_ps: Optional[set] = None) -> dict:
    """把「无标题样式但形似标题」的段落按字号档分层，返回 {w:p 元素: level}。

    ⚠️ **单一真相源**：fixer 的层级推断与 docx_reader 的伪标题采样共用本函数，
    否则会出现「检查说没问题、修正却改了一批」这类检查/修正脱节（v2.3.65 实测）。

    分层规则（best-effort，无法更可靠时的既定口径）：
    - 出现 ≥ 2 种不同字号 → 最大档 = L1、次大档 = L2 …
    - 只有 1 种字号（均匀加粗、无样式）→ 全部当作 L1（真实「大节/小节」场景可正确分层）

    ``min_pt``：字号下限，默认用 _PSEUDO_HEADING_FORMAT_MIN_PT（识别用途的宽松阈值）。
    ``exclude_ps``：需排除的 w:p 集合（标题页标题/作者，由标题页流程独占）。
    """
    threshold = _PSEUDO_HEADING_FORMAT_MIN_PT if min_pt is None else min_pt
    exclude_ps = exclude_ps or set()
    pseudo = []
    for para in paras:
        if para._p in exclude_ps:
            continue
        if is_pseudo_heading(para, min_pt=threshold):
            pseudo.append(para)
    if not pseudo:
        return {}
    out: dict = {}
    sizes = [round(s, 1) for p in pseudo if (s := _max_run_size_pt(p)) is not None]
    tiers = sorted(set(sizes)) if sizes else []
    if len(tiers) >= 2:
        tier_to_lvl = {sz: i + 1 for i, sz in enumerate(reversed(tiers))}
        for p in pseudo:
            s = _max_run_size_pt(p)
            out[p._p] = tier_to_lvl.get(round(s, 1) if s is not None else tiers[0], 1)
    else:
        for p in pseudo:
            out[p._p] = 1
    return out


def _pf_attr(para, doc_pf, attr):
    """段落格式取值，三级回退：run 所在段落 → 段落样式 → 文档默认(Normal)样式。

    返回 None 表示「无法确定」（而非 0），由调用方决定如何提示。
    """
    try:
        pf = para.paragraph_format
    except Exception:
        pf = None
    if pf is not None:
        v = getattr(pf, attr, None)
        if v is not None:
            return v
    try:
        spf = para.style.paragraph_format if para.style is not None else None
        if spf is not None:
            v = getattr(spf, attr, None)
            if v is not None:
                return v
    except Exception:
        pass
    if doc_pf is not None:
        try:
            return getattr(doc_pf, attr, None)
        except Exception:
            return None
    return None


# ---------------------------------------------------------------------------
# 段前 / 段后（spacing before/after）—— v2.3.75
#
# 为什么不能用 `_pf_attr(para, doc_pf, "space_before")`：python-docx 的属性只映射
# **`w:before` / `w:after`（twips，1/20 磅）**，而中文学校模板里极其常见的写法是
# 「段前 0.5 行 / 段后 1 行」，落盘成 **`w:beforeLines` / `w:afterLines`（单位 1/100 行）**。
# 直接读属性会得到 None → 被当成「段前 0 磅」→ 模板里写死的要求读不到、检查报 pass、
# 客户点修正也不会改（又一个「假通过」）。这里直读 XML 并做「行 → 磅」换算。
# ---------------------------------------------------------------------------
# 「1 行」折成多少磅：Word 单倍行距下正文行高 ≈ 1.2 × 字号（不同字体略有浮动，
# 中文宋体约 1.3）。取 1.2 是 Word 自身的换算基准，误差 <1 磅，用于比对足够。
_LINE_TO_PT = 1.2


def _ppr_spacing_pair(pPr, base_size_pt=None):
    """从一个 pPr 元素读段前/段后（磅）。读不到返回 None（不是 0）。

    OOXML 语义：`w:beforeLines` 与 `w:before` 同时存在时 **`*Lines` 优先**，
    这里遵循同一优先级。
    """
    if pPr is None:
        return None, None
    try:
        sp = pPr.find(qn("w:spacing"))
    except Exception:
        sp = None
    if sp is None:
        return None, None
    size = float(base_size_pt) if base_size_pt else 12.0
    out: list = []
    for pt_attr, ln_attr in (("w:before", "w:beforeLines"),
                             ("w:after", "w:afterLines")):
        v: Optional[float] = None
        raw_lines = sp.get(qn(ln_attr))
        if raw_lines is not None:
            try:
                v = round(float(raw_lines) / 100.0 * _LINE_TO_PT * size, 1)
            except (TypeError, ValueError):
                v = None
        if v is None:
            raw = sp.get(qn(pt_attr))
            if raw is not None:
                try:
                    # v2.3.100：`w:before/after` 除 twips 十进制数外，**还允许带单位**
                    # （官方 IEEE 模板的 Abstract 样式写的就是 `<w:spacing w:after="10pt"/>`）。
                    # 旧代码一律 `float(raw)/20` → "10pt" 抛 ValueError → 落到 0 磅，
                    # 于是摘要"段后 10pt"被判成 0，判定与修正都跟着错。
                    _s = str(raw).strip()
                    if _s.lower().endswith("pt"):
                        v = round(float(_s[:-2]), 1)
                    elif _s.lower().endswith("in"):
                        v = round(float(_s[:-2]) * 72.0, 1)
                    else:
                        v = round(float(_s) / 20.0, 1)      # twips → 磅
                except (TypeError, ValueError):
                    v = None
        out.append(v)
    return out[0], out[1]


def _spacing_pair_of_para(para, base_size_pt=None):
    """段前/段后（磅），三级回退：段落 → 段落样式链 → 文档默认（**逐属性**回退）。

    语义与行距/对齐同口径：整条链都没写 → Word 渲染就是 0 磅，返回 ``(0.0, 0.0)``
    而不是 ``(None, None)``。否则「模板要求段前 6 磅、文档一个字节没写」会被判成
    「读不到」→ 检查出 info、修正器却照改，两处口径不一致。
    """
    if base_size_pt is None:
        try:
            base_size_pt = _max_run_size_pt(para)
        except Exception:
            base_size_pt = None
    elems: list = []
    try:
        if para._p.pPr is not None:
            elems.append(para._p.pPr)
    except Exception:
        pass
    try:
        st = para.style
    except Exception:
        st = None
    while st is not None:                     # Heading 2 → Heading 1 → Normal
        try:
            if st.element.pPr is not None:
                elems.append(st.element.pPr)
        except Exception:
            pass
        st = getattr(st, "base_style", None)
    before: Optional[float] = None
    after: Optional[float] = None
    for el in elems:
        b, a = _ppr_spacing_pair(el, base_size_pt)
        if before is None:
            before = b
        if after is None:
            after = a
        if before is not None and after is not None:
            break
    return (0.0 if before is None else before), (0.0 if after is None else after)


def _looks_like_reference_entry(text: str) -> bool:
    """启发式判断一段文字是否「像参考文献条目」（用于给标题候选做佐证）。"""
    t = text.strip()
    if not t:
        return False
    if re.match(r"^\[\s*\d+\s*\]", t):                       # IEEE [1]
        return True
    if re.search(r"\((?:19|20)\d{2}[a-z]?\)", t):            # APA (2020)
        return True
    if re.search(r"\b(?:19|20)\d{2}\b", t) and len(t) > 25:  # 含年份的长句
        return True
    if re.search(r"\bdoi\b|https?://|Retrieved from", t, re.I):
        return True
    if len(t) > 20 and t.count(".") >= 2 and t[0].isupper():  # 作者. 题名. 出版项.
        return True
    return False


def find_reference_section(paras: list) -> tuple[Optional[int], list[int]]:
    """定位参考文献段：返回 (标题段下标, 条目段下标列表)。

    加固点（避免正文误判）：候选标题段之后必须紧跟至少一条「像参考文献条目」
    的段落；否则视为正文里出现的普通字样（如 "see the References section"），
    不认定为参考文献段。若标题段本身是 Heading 样式，则放宽该佐证要求。

    本函数为**唯一**的参考文献段落定位实现，docx_reader 与 fixer 共用。
    """
    for i, para in enumerate(paras):
        text = para.text.strip()
        if not text or len(text) > REF_HEADING_MAX_LEN:
            continue
        low = text.lower()
        if not any(m in low for m in _REF_HEADING_MARKERS):
            continue
        # 跳过空段，找到标题后第一条有内容的段落
        j = i + 1
        while j < len(paras) and not paras[j].text.strip():
            j += 1
        if j >= len(paras):
            continue
        if not _looks_like_reference_entry(paras[j].text) and _heading_level(para) is None:
            continue  # 佐证不足且非标题样式 → 不认定
        # 认定：从 j 起收集条目，直到遇到下一个标题段
        idxs: list = []
        k = j
        while k < len(paras):
            t = paras[k].text.strip()
            if t and _heading_level(paras[k]) is not None:
                break
            if t:
                idxs.append(k)
            k += 1
        if not idxs:
            continue
        return i, idxs
    return None, []


def _iter_table_paragraphs(doc):
    """遍历表格单元格内的段落（按 tc 元素去重，合并单元格不重复计数）。"""
    seen: set = set()
    try:
        tables = list(doc.tables)
    except Exception:
        return
    for tbl in tables:
        try:
            rows = list(tbl.rows)
        except Exception:
            continue
        for row in rows:
            try:
                cells = list(row.cells)
            except Exception:
                continue
            for cell in cells:
                try:
                    tc = cell._tc
                except Exception:
                    continue
                if tc in seen:
                    continue
                seen.add(tc)
                try:
                    for p in cell.paragraphs:
                        yield p
                except Exception:
                    continue


def iter_table_paragraphs(doc):
    """公开入口：遍历表格单元格内的段落（checker / fixer 共用同一实现）。"""
    return _iter_table_paragraphs(doc)


# ---------------------------------------------------------------------------
# 表格边框采样（三线表，v2.3.85）—— 单一真相源，checker 与 fixer 共用
# ---------------------------------------------------------------------------
# 为什么要读「有效边框」而不是只看 tblPr/tblBorders：
#   Word 里表格线的来源有三层，优先级 单元格 tcBorders > 表格 tblBorders > 表格样式
#   （tblStyle 指向的 style 里那份 tblBorders，例如 "Table Grid"）。只看中间一层会有
#   两种误判：①客户用了 Table Grid 样式、表格级没写边框 → 看着有框却说"没框"；
#   ②单元格上带着残留的 tcBorders → 表格级已经改干净了，客户打开还是一条不少。
# 所以这里逐格解析「格 → 表 → 样式」三级，得到客户肉眼所见的那套线。
_TBL_PR_AFTER_BORDERS = ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
                         "w:tblCaption", "w:tblDescription")


def _read_border(box, side):
    """从 <w:tblBorders>/<w:tcBorders> 容器读某一边：返回 (val, sz)；无该边返回 None。"""
    if box is None:
        return None
    el = box.find(qn("w:" + side))
    if el is None:
        return None
    try:
        sz = int(el.get(qn("w:sz")) or 0)
    except Exception:
        sz = 0
    return (el.get(qn("w:val")) or "single", sz)


def _border_visible(border) -> bool:
    """边框是否真的画得出来（nil / none = 不画）。"""
    return bool(border) and str(border[0]).lower() not in ("nil", "none")


def _table_style_borders(doc, tbl):
    """表格所引用样式（tblStyle）里定义的 tblBorders 容器；没有返回 None。"""
    try:
        tblPr = tbl._tbl.tblPr
    except Exception:
        return None
    sid = None
    if tblPr is not None:
        st = tblPr.find(qn("w:tblStyle"))
        if st is not None:
            sid = st.get(qn("w:val"))
    if not sid:
        return None
    try:
        for style in doc.styles.element.findall(qn("w:style")):
            if style.get(qn("w:styleId")) == sid:
                spPr = style.find(qn("w:tblPr"))
                if spPr is not None:
                    return spPr.find(qn("w:tblBorders"))
    except Exception:
        pass
    return None


def _effective_border(tblPr, style_box, tcPr, side):
    """有效边框：单元格 → 表格 → 表格样式，第一个有定义的就赢。"""
    b = _read_border(tcPr.find(qn("w:tcBorders")) if tcPr is not None else None, side)
    if b is not None:
        return b
    b = _read_border(tblPr.find(qn("w:tblBorders")) if tblPr is not None else None, side)
    if b is not None:
        return b
    return _read_border(style_box, side)


def _row_tc_prs(row):
    """一行里每格的 tcPr（合并单元格会重复返回同一个 tc，这里去重）。"""
    out, seen = [], set()
    try:
        cells = list(row.cells)
    except Exception:
        return out
    for cell in cells:
        try:
            tc = cell._tc
        except Exception:
            continue
        if tc in seen:
            continue
        seen.add(tc)
        out.append(tc.tcPr if tc is not None else None)
    return out


_BYLINE_STYLE_KEYWORDS = ("author", "affiliation", "byline", "authors")


def sample_byline(doc) -> tuple:
    """采样 byline（作者信息块）→ (字号磅值或 None, 行数)。

    v2.3.100（F-5 / 落地裁定 C-6）。IEEE 会议模板的作者块用 ``Author`` 样式承载，
    OWL 与官方模板对字号的说法不一致（OWL 10pt / 模板 11pt），老板裁定**以模板 11pt
    为准。本函数只负责读，判与修在 checker/fixer。

    取字号口径（宁可不判，不臆造）：
      · **样式字号优先** —— byline 的字号由 Author/byline 样式定义，模板里 run 上的
        9pt 是"占位提示"（同一批 run 的文字是 "line 1: 1st Given Name Surname" 这种
        指导文字），拿它当规范值会把官方模板判红；
      · 没有任何 byline 样式时（客户手打）才退回 run 显式字号的众数；
      · 两类都读不到 → None（不判）。
    行数：连续 byline 段的条数（结构只提示，不判好坏 —— 删/补行都是改文字）。
    """
    paras = []
    for para in doc.paragraphs:
        try:
            nm = (para.style.name or "").strip().lower()
        except Exception:
            continue
        if not (para.text or "").strip():
            continue
        if any(nm == k or nm.startswith(k) for k in _BYLINE_STYLE_KEYWORDS):
            paras.append(para)
    if not paras:
        # 退回标题页探测到的作者段（detect_title_page 只抓 ≤3 段，作兜底）
        try:
            tp = detect_title_page(doc) or {}
            for idx in (tp.get("authors") or ()):  # noqa: E501
                if 0 <= idx < len(doc.paragraphs):
                    paras.append(doc.paragraphs[idx])
        except Exception:
            paras = []
    if not paras:
        return None, 0
    style_sizes, run_sizes = [], []
    for para in paras:
        try:
            st = para.style
        except Exception:
            st = None
        if st is not None and st.font is not None and st.font.size is not None:
            style_sizes.append(round(float(st.font.size.pt), 1))
        for r in para.runs:
            if (r.text or "").strip() and r.font.size is not None:
                run_sizes.append(round(float(r.font.size.pt), 1))
    if style_sizes:
        from collections import Counter
        return Counter(style_sizes).most_common(1)[0][0], len(paras)
    if run_sizes:
        from collections import Counter
        return Counter(run_sizes).most_common(1)[0][0], len(paras)
    return None, len(paras)


def sample_template_boilerplate(doc) -> list:
    """扫出官方模板**指导文字**残留（v2.3.100，D-1 / 裁定 C-11）。

    IEEE 官方会议模板下载页原文：「Please ensure that all **guidance text is removed**
    from your conference paper prior to submission… **Failure to remove template text
    from your paper may result in your paper not being published.**」
    本地两份 2024-07-16 官方模板实测仍含：
      · 正文：「This electronic document is a "live" template…」
      · 正文：「line 1: 1st Given Name Surname」「line 2: dept. name of organization」
      · 正文：「Identify applicable funding agency here」
      · 页脚：「XXX-X-XXXX-XXXX-X/XX/$XX.00 ©20XX IEEE」
    客户把模板直接交上去 = 论文不发表，而我们此前报"全部合规"。

    只读不改：返回命中位置清单，删除由客户自己做（删文字越权，且 P-2 锁死文字）。
    """
    hits: list = []

    def _scan(text: str, where: str, index, limit: int = 60):
        t = (text or "").strip()
        if not t:
            return
        if is_template_boilerplate_text(t):
            hits.append({"where": where, "index": index,
                         "text": t[:limit] + ("…" if len(t) > limit else "")})

    for i, para in enumerate(doc.paragraphs):
        _scan(para.text or "", "body", i)
    # 页眉 / 页脚：官方模板的论文编号/版权行在**页脚**，不在正文
    try:
        for part, kind, _scope in _iter_header_footer_parts(doc):
            for j, para in enumerate(_part_paragraphs(part)):
                _scan(para.text or "", kind, j)
    except Exception:
        pass
    return hits


# —— F-9：三类「只提示」的引用 / 标点观测（v2.3.100）——
# C-12：IEEE Reference Guide V3.28.2025 原文：「Reference ranges in text **will not
# include an en dash** … '[1]–[4]' will now be '[1], [2], [3], [4]'」。
# ⚠️ 形状是「**两个方括号**中间一个破折号」：[3]–[5] —— 不是 [3-5]。
# 我第一版写成 `\[\d+\s*-\s*\d+\]`（一个方括号里含范围），实测一个都匹配不到。
_CITATION_RANGE_RE = re.compile(
    r"\[\s*\d+\s*\]\s*[\u2013\u2014-]\s*\[\s*\d+\s*\]")
# C-13：同指南：「**Do not use commas** around Jr., Sr., and III：Michael Smith Jr.」。
# 只认「逗号 + 空白 + Jr./Sr./II/III/IV」这一种形状，不碰 "Smith, J."（那是姓名缩写）。
_SUFFIX_COMMA_RE = re.compile(r",\s+(?:Jr|Sr|II|III|IV|V)\.?(?=\s|$|[),.;])")
# C-14：公式编号形态。IEEE Mathematics Guide：「Hyphens and periods are usually
# removed from equation numbers, i.e., (1a) rather than (1-a) … consistently」。
# ESM 则说 (1a)/(1.1)/(1-1) 都行、只要全文一致 → 我们**只采集形态、不判好坏**。
_EQ_NUM_STYLE_RES = (
    ("plain", re.compile(r"\(\s*\d+\s*\)")),
    ("letter_suffix", re.compile(r"\(\s*\d+\s*[a-z]\s*\)")),
    ("dotted", re.compile(r"\(\s*\d+(?:\.\d+)+\s*\)")),
    ("hyphenated", re.compile(r"\(\s*\d+-\d+[a-z]?\s*\)")),
    ("appendix", re.compile(r"\(\s*[A-Z]\d+[a-z]?\s*\)")),
)


def sample_text_style_notes(doc) -> dict:
    """扫全文引用/标点/公式编号的三类**只提示**观测（v2.3.100，F-9）。

    返回 ``{"citation_range_en_dash": [...], "suffix_name_comma": [...],
    "equation_number_styles": [...]}``。每条命中带 ``para``（段下标）与 ``text``（摘要）。

    🔴 三类全部**只提示**：改标点就是改客户的文字（P-2 锁死），所以 checker 侧一律
    info、**绝不挂 FixAction**。C-14 更特殊 —— ESM 明说 (1a)/(1.1)/(1-1) 均可、
    只要求全文一致，所以我们只判「混用了几种形态」，绝不判某种形式为错。
    """
    out = {"citation_range_en_dash": [], "suffix_name_comma": [],
           "equation_number_styles": []}
    styles_seen: set = set()

    def _note(bucket: list, i: int, text: str, limit: int = 48):
        t = (text or "").strip()
        if not t:
            return
        bucket.append({"para": i,
                       "text": t[:limit] + ("…" if len(t) > limit else "")})

    for i, para in enumerate(doc.paragraphs):
        t = para.text or ""
        if not t.strip():
            continue
        for m in _CITATION_RANGE_RE.finditer(t):
            _note(out["citation_range_en_dash"], i, m.group(0))
        for m in _SUFFIX_COMMA_RE.finditer(t):
            _note(out["suffix_name_comma"], i, m.group(0))
        if "(" in t:
            for name, rx in _EQ_NUM_STYLE_RES:
                if rx.search(t):
                    styles_seen.add(name)
                    break
    out["equation_number_styles"] = sorted(styles_seen)
    return out


# —— D 批：IEEE 系列的只读观测（v2.3.100）——
# 表标签：IEEE 表题是「TABLE I」（罗马数字、大写），图题是「Fig. N.」（阿拉伯数字 + 句点）。
_CAP_TABLE_LABEL_RE = re.compile(r"^\s*TABLE\s+([IVXLCDM]+|\d+)\b", re.IGNORECASE)
_CAP_FIG_LABEL_RE = re.compile(r"^\s*Fig(?:ure)?\.?\s+(\d+)\b", re.IGNORECASE)
# 正文里对图表的引用（用于"首次引用必须按数字序"）
_FIG_REF_RE = re.compile(r"\bFig(?:ure)?\.?\s+(\d+)\b", re.IGNORECASE)
_TAB_REF_RE = re.compile(r"\bTABLE\s+(\d+|I|II|III|IV|V|VI|VII|VIII|IX|X)\b")
# 章节标题里的"单数"形式（IEEE 惯例用复数 Introduction / References）
_SINGULAR_HEADS = {
    "introduction", "references", "acknowledgment", "acknowledgement",
    "appendix", "biography", "abstract", "conclusion", "conclusions",
}


def _roman_to_int(s: str) -> int:
    """罗马数字 → 整数（认不了就返回 0）。"""
    vals = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
    s = (s or "").strip().upper()
    if not s:
        return 0
    total, prev = 0, 0
    for ch in reversed(s):
        v = vals.get(ch)
        if v is None:
            return 0
        if v < prev:
            total -= v
        else:
            total += v
            prev = v
    return total


def sample_ieee_extras(doc) -> dict:
    """采集 IEEE 系列的只读观测（D-2~D-11）。

    🔴 本函数**只读**：所有结果都只用于出提示或判定"形态是否一致"，
    绝不在这里改客户文字。涉及"改文字"的项目（引用顺序、拼写）在 checker 侧一律 info。
    """
    out = {
        "caption_labels": [], "caption_position": {}, "figure_ref_order": [],
        "first_footnote": {}, "acknowledgment": {}, "coi": {},
        "biography": {}, "index_terms": {}, "abstract_numbered_refs": [],
        "numbered_singular_heads": [],
    }
    paras = list(doc.paragraphs)
    texts = [(p.text or "") for p in paras]

    # —— D-2 / D-3：表/图标签形态与句点 ——
    for i, t in enumerate(texts):
        ts = t.strip()
        if not ts:
            continue
        m = _CAP_TABLE_LABEL_RE.match(ts)
        if m:
            tok = m.group(1)
            out["caption_labels"].append({
                "kind": "table", "para": i, "label": ts.split("\n")[0][:40],
                "number": _roman_to_int(tok) if tok.isalpha() else int(tok),
                "numeral": "roman" if tok.isalpha() else "arabic",
                "uppercase": ts.strip().startswith("TABLE"),
                "ends_with_period": ts.rstrip().endswith("."),
                "text": ts[:60],
            })
            continue
        m = _CAP_FIG_LABEL_RE.match(ts)
        if m:
            out["caption_labels"].append({
                "kind": "figure", "para": i, "label": ts.split("\n")[0][:40],
                "number": int(m.group(1)), "numeral": "arabic",
                "uppercase": False,
                "ends_with_period": ts.rstrip().endswith("."),
                "text": ts[:60],
            })

    # —— D-2 续：表题在表**上方**、图题在图**下方**（IEEE 惯例）——
    try:
        body = doc.element.body
        seq = [c for c in body.iterchildren()
               if c.tag in (qn("w:p"), qn("w:tbl"))]
        tbl_idx = [k for k, c in enumerate(seq) if c.tag == qn("w:tbl")]
        for lab in out["caption_labels"]:
            if lab["kind"] != "table":
                continue
            # 找该 caption 段在 body 序列中的位置
            for k, c in enumerate(seq):
                if c.tag != qn("w:p"):
                    continue
                if paras and _same_p(c, paras[lab["para"]]):
                    nxt_tbl = next((t for t in tbl_idx if t > k), None)
                    if nxt_tbl is not None:
                        lab["caption_above_table"] = (nxt_tbl - k) <= 2
                    break
    except Exception:
        pass

    # —— D-4：正文首次引用图表的数字序 ——
    seen_fig, seen_tab = [], []
    for i, t in enumerate(texts):
        for m in _FIG_REF_RE.finditer(t):
            n = int(m.group(1))
            if n not in seen_fig:
                seen_fig.append(n)
                out["figure_ref_order"].append({"para": i, "number": n,
                                               "text": m.group(0)})
        for m in _TAB_REF_RE.finditer(t):
            tok = m.group(1)
            n = _roman_to_int(tok) if tok.isalpha() else int(tok)
            if n and n not in seen_tab:
                seen_tab.append(n)
    out["figure_ref_order_ok"] = (seen_fig == sorted(seen_fig)) if seen_fig else None
    out["table_ref_order_ok"] = (seen_tab == sorted(seen_tab)) if seen_tab else None

    # —— D-5：首脚注 ——
    # ⚠️ python-docx 的 `Package` **没有** `.blob` 属性（实测），脚注 part 要从
    # `package.iter_parts()` 里按 partname 找 —— 这是唯一稳定的取法。
    try:
        _blob = None
        try:
            for _part in doc.part.package.iter_parts():
                if str(getattr(_part, "partname", "")) == "/word/footnotes.xml":
                    _blob = _part.blob
                    break
        except Exception:
            _blob = None
        if _blob:
            from lxml import etree as _et
            _root = _et.fromstring(_blob)
            _ids, _ftxts = [], []
            for _fn in _root.findall(qn("w:footnote")):
                _fid = _fn.get(qn("w:id"))
                try:
                    _fid = int(_fid)
                except Exception:
                    continue
                if _fid <= 0:
                    continue          # 分隔线条目
                _ftxt = "".join(t.text or "" for t in _fn.iter(qn("w:t")))
                _ids.append(_fid)
                _ftxts.append(_ftxt.strip())
            if _ftxts:
                out["first_footnote"] = {
                    "count": len(_ftxts),
                    "first_text": _ftxts[0][:80],
                    "numbered": bool(_ids and sorted(_ids) == list(range(1, len(_ids) + 1))),
                }
    except Exception:
        pass

    # —— D-6 / D-7 / D-8 / D-9：致谢、COI、传记、Index Terms ——
    abs_head, abs_body = find_abstract_section(paras)
    abs_last = max(abs_body) if abs_body else None
    for i, t in enumerate(texts):
        low = t.strip().lower()
        if not low:
            continue
        if low.startswith("acknowledg"):
            out["acknowledgment"] = {
                "para": i,
                "spelling": ("acknowledgement" if low.startswith("acknowledgement")
                             else "acknowledgment"),
                "text": t.strip()[:70],
            }
        elif low.startswith("conflict of interest") or low.startswith("conflicts of interest"):
            out["coi"] = {"para": i, "text": t.strip()[:70]}
        elif low.startswith("index terms") or low.startswith("indexterm"):
            out["index_terms"] = {
                "para": i,
                "text": t.strip()[:70],
                "in_abstract_block": (abs_last is not None and i <= abs_last),
                "is_last_in_abstract": (abs_last is not None and i == abs_last),
            }
        elif low.startswith("biography") or low.startswith("biographies"):
            # 作者传记：姓名一般加粗
            _name_bold = None
            _nxt = paras[i + 1] if i + 1 < len(paras) else None
            if _nxt is not None:
                _rs = [r for r in _nxt.runs if (r.text or "").strip()]
                if _rs:
                    _name_bold = any(r.font.bold is True for r in _rs)
            out["biography"] = {"para": i, "name_bold": _name_bold,
                                "text": t.strip()[:60]}

    # —— D-10：摘要里出现编号引用 / 编号公式 ——
    if abs_body:
        for j in abs_body:
            _t = texts[j] if j < len(texts) else ""
            for _m in re.finditer(r"\[\s*\d+\s*\]", _t):
                out["abstract_numbered_refs"].append(
                    {"para": j, "text": _m.group(0)})
            for _m in re.finditer(r"\(\s*\d+\s*[a-z]?\s*\)", _t):
                out["abstract_numbered_refs"].append(
                    {"para": j, "text": _m.group(0)})

    # —— D-11：被编号的**单数**章节标题（IEEE 惯例用复数）——
    try:
        _cand = [p for p in paras if heading_level(p) is not None]
        _lvl = structural_heading_map(_cand)
        for _p, _lv in _lvl.items():
            _txt = (_p.text or "").strip().lower()
            _core = re.sub(r"^[\s\d.ivxlcdm()\-–—]+", "", _txt).strip(" .:")
            if _core in _SINGULAR_HEADS:
                _has_prefix = bool(re.match(r"^\s*(?:[\d]+|[IVXLCDM]+|[A-Z])\.?\s", _txt))
                if _has_prefix:
                    out["numbered_singular_heads"].append(
                        {"level": _lv, "text": (_p.text or "").strip()[:50]})
    except Exception:
        pass
    return out


def sample_mla_apa_extras(doc) -> dict:
    """MLA + APA 剩余的只读观测（v2.3.100，E 批 E-1/E-3/E-4/E-5/E-7/E-8）。

    全部**只读**：判定结果只用于出 info 提示，涉及"改文字"的一律不改。
    """
    out = {"body_title_repeat": False, "numbered_body_heads": [],
           "author_note_index": None, "author_note_has_superscript": None,
           "first_page_left_texts": []}
    paras = list(doc.paragraphs)
    texts = [(p.text or "") for p in paras]
    _norm = lambda t: re.sub(r"\s+", " ", (t or "").strip().lower())

    # —— E-1：正文首段是否重复了论文标题（APA 稿里标题只该出现在题名页）——
    _title = ""
    for t in texts[:12]:
        if t.strip():
            _title = _norm(t)
            break
    if _title:
        for i, t in enumerate(texts):
            if i < 6 or not t.strip():
                continue
            if _norm(t) == _title:
                out["body_title_repeat"] = True
                break

    # —— E-3 / E-4：正文标题带编号（APA / MLA 都不用编号标题）——
    try:
        _cand = [p for p in paras if heading_level(p) is not None]
        _lvl = structural_heading_map(_cand)
        for _p, _lv in _lvl.items():
            _txt = (_p.text or "").strip()
            if not _txt:
                continue
            if re.match(r"^\s*(?:\d+(?:\.\d+)*[.)]?|[IVXLCDM]+[.)]?|[A-Z][.)])\s+\S", _txt):
                out["numbered_body_heads"].append(
                    {"level": _lv, "text": _txt[:50]})
    except Exception:
        pass

    # —— E-7 / E-8：APA 专业版 Author Note 的位置与上标编号 ——
    for i, t in enumerate(texts):
        if _norm(t).startswith("author note"):
            out["author_note_index"] = i
            # 该段或其紧邻段里是否有上标 run（APA 的作者注编号用上标）
            _sup = False
            for _j in (i, i + 1):
                if 0 <= _j < len(paras):
                    for _r in paras[_j].runs:
                        if not ( _r.text or "").strip():
                            continue
                        try:
                            if _r.font.superscript is True:
                                _sup = True
                        except Exception:
                            pass
            out["author_note_has_superscript"] = _sup
            break

    # —— E-5：MLA 首页左上四行（姓名/教授/课程/日期）——
    try:
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _WA
        _cnt = 0
        for i, p in enumerate(paras[:12]):
            _t = (p.text or "").strip()
            if not _t:
                continue
            if p.alignment == _WA.LEFT:
                _cnt += 1
                out["first_page_left_texts"].append(_t[:50])
            else:
                break
    except Exception:
        pass
    return out


# v2.3.100（E-2 / D-18）：纯标签段（"Table 1" / "Figure 1."）——只含编号、没有描述文字。
_CAPTION_PURE_LABEL_RE = re.compile(
    r"^\s*(?:Table|Figure|Fig\.?)\s*\d+\s*\.?\s*$", re.IGNORECASE)


def sample_caption_typography(doc) -> dict:
    """采样 APA caption 的字形：号段是否粗体、题段是否斜体（v2.3.100，E-2 / D-18）。

    APA 7：号（"Table 1"）**粗体**，题（描述）**斜体**。
    「题段」的位置约定：紧跟在号段之后的那一段（APA 的标准排法就是号、题各占一段）。
    只在**确实采到「号 + 题」这一对**时才给结论，否则 None（不判）。

    ⚠️ 字形取「run 直接格式 → 样式链」三级回退（与摘要标题同一口径），
    否则官方样例（caption 用 Normal + 直接格式）会读成"读不到"。
    """
    out = {"caption_label_bold": None, "caption_title_italic": None, "count": 0}
    try:
        paras = list(doc.paragraphs)
    except Exception:
        return out
    label_bolds, title_italics = [], []
    for i, p in enumerate(paras):
        if not is_caption_paragraph(p):
            continue
        text = (p.text or "").strip()
        if not _CAPTION_PURE_LABEL_RE.match(text):
            continue                      # 不是纯号段（可能是号+题同段）
        runs = [r for r in p.runs if (r.text or "").strip()]
        if not runs:
            continue
        _b = _run_bool_with_style(runs, p, "bold")
        # ⚠️ None（run 与样式链都没写）= 渲染出来就是**不加粗** → 记 False。
        # 与摘要标题那边不同：那里 None 表示"读不到、别冤枉客户"；这里要求是
        # "必须是粗体"，未设置就是不满足 → 如实报出来（可修）。
        label_bolds.append(bool(_b))
        # 题段 = 紧跟其后的一段
        if i + 1 < len(paras):
            nxt = paras[i + 1]
            ntxt = (nxt.text or "").strip()
            if ntxt and not is_caption_paragraph(nxt):
                nruns = [r for r in nxt.runs if (r.text or "").strip()]
                if nruns:
                    _it = _run_bool_with_style(nruns, nxt, "italic")
                    title_italics.append(bool(_it))
    out["count"] = len(label_bolds)
    if label_bolds:
        out["caption_label_bold"] = (sum(label_bolds) * 2 >= len(label_bolds))
    if title_italics:
        out["caption_title_italic"] = (sum(title_italics) * 2 >= len(title_italics))
    return out


def sample_mla_title(doc) -> dict:
    """定位并实测 MLA 的论文标题段（v2.3.100，E-6 / D-22）。

    MLA 9：标题**居中**、**不加粗、不加斜**、不额外加空行，字号与正文相同。

    🔴 定位的难点（计划里点名了）：MLA 稿里"论文标题"与"章节标题"长得一样，
    而"首页左上四行信息块"与"标题"也可能都左对齐 —— **定位错就会改错客户的段**。
    故分两级置信度：
      · **certain=True**：首页首个**居中**段（`doc_title_index`）—— 绝无可能是信息块；
      · **certain=False**：首页左上信息块之后的第一段非空段（客户没居中时兜底）——
        可能其实是正文首段，所以只给**提示**，**不自动改**。

    返回 ``{"index":…, "alignment":…, "bold":…, "italic":…, "certain":…}``。
    """
    out = {"index": None, "alignment": None, "bold": None, "italic": None,
           "certain": False}
    try:
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _WA
        paras = list(doc.paragraphs)
    except Exception:
        return out
    if not paras:
        return out

    # —— 定位：跳过首页开头的**左对齐连续块**（MLA 的姓名/教授/课程/日期信息块），
    #    取其后第一段非空段。
    # ⚠️ 不能用 doc_title_index：它要求「**第一个非空段**就是标题」，而 MLA 稿的
    #   第一个非空段是信息块 → 永远认不出（实测）。
    i = 0
    while i < len(paras):
        p = paras[i]
        t = (p.text or "").strip()
        if not t:
            i += 1
            continue
        if p.alignment == _WA.LEFT:
            i += 1
            continue
        break
    if i >= len(paras):
        return out
    p = paras[i]
    text = (p.text or "").strip()
    if not text:
        return out
    runs = [r for r in p.runs if (r.text or "").strip()]
    out["index"] = i
    # 🔴 `_effective_para_alignment` 已经返回**字符串**（"left"/"center"/…），
    #    不要再过 `_ALIGN_MAP`（那会把值映成 None）。
    out["alignment"] = _effective_para_alignment(p)
    if runs:
        out["bold"] = _run_bool_with_style(runs, p, "bold")
        out["italic"] = _run_bool_with_style(runs, p, "italic")
    # 置信度：**居中**才是确定无疑的标题（信息块一定是左对齐）；
    # 不居中时可能其实是正文首段 → 只提示、不自动改。
    out["certain"] = (out["alignment"] == "center")
    return out


def _same_p(el, para) -> bool:
    """body 序列里的元素是否就是某个段落对象（用于定位 caption 段的下标）。"""
    try:
        return el is para._p
    except Exception:
        return False


def caption_block_top(seq, j: int, doc) -> int:
    """表格**上方**紧邻段 ``j`` → 承载「表格与文字之间那一条空隙」的段落下标。

    表题常常是**两段**：「Table 1」+ 斜体标题（APA 标准写法），而按文本只能识别前者。
    若把 24pt 加在紧邻那段（= 标题行）的 space_after 上，等于把表题推离自己的表格 ——
    比不改更难看。故：
      · ``j`` 本身是表题 → 上溯到**连续表题块**的最上一段；
      · ``j`` 不是表题，但 ``j-1`` 是表题、且 ``j`` 短（≤120 字，像标题不像正文句）
        → 认为 ``j`` 是表题续行，把整块算进来；
      · 否则原样返回 ``j``（普通正文，空隙就是它的 space_after）。
    reader 的 ``table_gap_samples`` 与 fixer 的 ``_apply_table_gap`` **必须共用**本函数，
    否则「检查说 24pt、实际改在别处」。
    """
    def _cap(k: int) -> bool:
        if k < 0 or k >= len(seq) or seq[k].tag != qn("w:p"):
            return False
        try:
            from docx.text.paragraph import Paragraph
            return is_caption_paragraph(Paragraph(seq[k], doc))
        except Exception:
            return False

    k = j
    if _cap(k):
        while _cap(k - 1):
            k -= 1
        return k
    # 续行：紧邻段短、且它上面那段是表题
    if _cap(k - 1):
        try:
            from docx.text.paragraph import Paragraph
            txt = (Paragraph(seq[j], doc).text or "").strip()
        except Exception:
            txt = ""
        if txt and len(txt) <= 120:
            k -= 1
            while _cap(k - 1):
                k -= 1
    return k


def _is_table_note(para) -> bool:
    """段落是否是表格的**表注**（APA 惯例：紧贴表格下方，以 "Note." 开头）。

    v2.3.100（F-2 / 裁定 C-2）。APA 7 表注规范：表注置于表格下方、与表格紧贴，
    所以「表格与表注之间」本来就不该有额外空隙 —— 旧写法把表注当普通文字，
    量到 0pt 就判"below 0 pt"，把官方样例判红。
    只认 **"Note." 开头的短段**（≤200 字）：正文里 "Note that ..." 是句子，不误伤。
    """
    try:
        text = (para.text or "").strip()
    except Exception:
        return False
    if not text or len(text) > 200:
        return False
    return bool(re.match(r"(?i)^note\.\s", text))


# v2.3.100（F-7）：Word 的**样式提示 run** —— 文本就是 "(references)"、"(body text)"
# 这种，Word 用它在页面显示当前样式名，自带 SimSun 10pt。混进参考文献字号众数会把
# 模板的 8pt 顶成 10pt（实测）。只认"纯字母串"，条目里的 "(2020)" 不受影响。
_STYLE_HINT_RE = re.compile(r"^\([A-Za-z][A-Za-z ]{0,28}\)$")


# v2.3.100（F-7 先建、D-1 扩展为 template_boilerplate 维度）：官方模板**指导文字**特征串。
# 来源：IEEE 官方会议模板下载页原文「Please ensure that all guidance text is removed from
# your conference paper prior to submission… **Failure to remove template text from your
# paper may result in your paper not being published.**」—— 客户直接交就是"论文不发表"，
# 而我们此前完全不查。这里先取**完整短语**（宁窄勿宽：只匹配 "Given Name" 这种碎片会
# 误命中正常论文的姓名栏）。
TEMPLATE_BOILERPLATE_CUES = (
    "this electronic document is a",
    "given name surname",
    "dept. name of organization",
    "identify applicable funding agency",
    "name of organization",
    "ieee conference templates contain guidance text",
    "xxx-x-xxxx-xxxx-x/xx",
    "use style: paper title",
    "are not captured in xplore",
)


def is_template_boilerplate_text(text: str) -> bool:
    """文本是否是 IEEE 官方模板的**指导文字**（v2.3.100，C-11 / D-12）。

    只认完整短语，避免把客户论文里的正常内容误判成模板残留。
    """
    try:
        low = (text or "").strip().lower()
    except Exception:
        return False
    if not low:
        return False
    return any(cue in low for cue in TEMPLATE_BOILERPLATE_CUES)


def _is_real_ref_entry(paras, i: int) -> bool:
    """``paras[i]`` 是不是**真正的参考文献条目**（v2.3.100，F-7 / 裁定 C-10）。

    采用**排除法**而非准入法（踩过）：早期版本要求"样式含 reference 或带 numPr"
    才算真条目，结果把普通论文里用 Normal 样式手打的参考文献**整段滤光**，
    ``reference_font`` / ``reference_size`` / ``reference_spacing`` 三个维度全部
    读成 None（``test_param_pipeline`` roundtrip 实测回归）。普通客户的条目才是
    常态，官方模板的特殊样式不是 —— 准入法会把常态判成"读不到"。

    现在只排除一类**确定不是条目**的段：文本命中官方模板**指导文字**特征串的
    （``is_template_boilerplate_text``）。官方模板 "References" 标题下夹 4 段
    Body Text 指导文字、末尾又夹 1 段，不排除就会把它们的 SimSun 10pt 混进条目
    字号众数（实测把模板 references 样式的 8pt 顶成 10pt）。
    只认形状、不猜内容：模板残留由 D 批 ``template_boilerplate`` 维度单独报告，
    只提示不删。
    """
    try:
        p = paras[i]
    except Exception:
        return False
    return not is_template_boilerplate_text(p.text or "")


def table_gap_samples(doc) -> tuple:
    """每张表与前后**紧邻文字**之间的空隙（磅）→ (样本列表, 紧邻段落的 w:p 集合)。

    APA 7（Publication Manual 7th, Tables and Figures）原文：
    「If embedded, place the table or figure at either the top or the bottom of the page
    **with an extra double-spaced line between the table or the figure and any text**.」
    即：表格是「段落段前段后不加额外空」（§2.21）这条规则的**例外**，它必须和任何
    文字之间额外空一条双倍行距。v2.3.97 之前没有任何维度管这件事，而「正文段前段后
    应为 0」还会把原文里 14pt 的段前一起清掉 → 修完表格紧贴标题（老板 2026-10-23
    截图）。本函数只负责读，不负责判/改。

    规则（宁可不判，绝不臆造）：
      · 紧邻的是**空段落** → 这条空行已经存在（gap 记为 None、blank=True）；
      · 紧邻的是正文段 → 空隙 = 该段 space_after（表格在前）/ space_before（表格在后）；
      · 紧邻的是另一张表 / 没有元素 → 不判（gap=None）。
      · **v2.3.100（F-2 / 裁定 C-2）跨页侧不判**：docx 没有绝对页坐标，只能用近似 ——
        邻接段落带 `w:pageBreakBefore`（它是新页的第一段）说明表格与该段之间隔着分页，
        这一侧量出来的"空隙"其实跨了两页，判它毫无意义。该侧记 `same_page=False`，
        checker 见到 False 就跳过（宁可不判，不冤枉客户）。
    返回的 `adjacent`（**紧邻且非空**的段落）必须从「正文段前段后=0」里豁免，
    否则两个维度互相清，客户修完复检又报红（自伤）。采样侧同样要剔掉它们，
    否则那 24pt 会抬高「正文段前段后」的众数。
    """
    samples: list = []
    adjacent: set = set()
    try:
        from docx.text.paragraph import Paragraph
        body = doc.element.body
        seq = [c for c in body.iterchildren()
               if c.tag in (qn("w:p"), qn("w:tbl"))]
    except Exception:
        return samples, adjacent

    def _neighbour(j: int, side: str):
        """返回 (空隙磅值, 是否真空行, 承载间隙的那个 w:p)。"""
        if j < 0 or j >= len(seq) or seq[j].tag != qn("w:p"):
            return None, False, None
        para = Paragraph(seq[j], doc)
        if not (para.text or "").strip():
            return None, True, None          # 已有真空行
        # 🔴 表格**上方**紧邻的是表题（APA 表题紧贴表格）→ 那一条额外行距必须落在
        # **表题块之上**（正文与表题之间），不是表题与表格之间；否则等于把表题推离
        # 自己的表格。故上溯到表题块最上一段，读它的 space_before。
        if side == "after":
            k = caption_block_top(seq, j, doc)
            if k != j or is_caption_paragraph(para):
                # 🔴 v2.3.100（F-2 / 裁定 C-2）：上溯到表题块最上一段后，还要看
                # **表题块之上**是不是空段落。官方 APA 样例就是「空段 + "Table 1" +
                # 表题描述 + 表格」四件套 —— 空隙其实已经由那个空段提供了，
                # 旧写法只读到表题自身的 space_before=0 → 判"上方 0pt"→ **把官方
                # 合规样例判红**。有空行就是有空行，这一侧不判。
                if k - 1 >= 0 and seq[k - 1].tag == qn("w:p"):
                    _above = Paragraph(seq[k - 1], doc)
                    if not (_above.text or "").strip():
                        return None, True, None
                _sb, _ = _spacing_pair_of_para(Paragraph(seq[k], doc))
                return float(0.0 if _sb is None else _sb), False, seq[k]
        sb, sa = _spacing_pair_of_para(para)
        # 🔴 v2.3.100（F-2）：表格**下方**紧邻的是表注（APA 惯例 "Note. ..."）时，
        # 这一侧不判 —— 表注紧贴自己的表格是 APA 标准写法（官方 mixed-methods 样例
        # Table 3 正是如此），旧写法把它当普通文字 → 判"below 0 pt"→ 把官方样例判红。
        if side == "before" and _is_table_note(para):
            return None, True, None
        val = sa if side == "after" else sb
        return float(0.0 if val is None else val), False, seq[j]

    n = 0
    for i, child in enumerate(seq):
        if child.tag != qn("w:tbl"):
            continue
        n += 1
        gap_b, blank_b, holder_b = _neighbour(i - 1, "after")
        gap_a, blank_a, holder_a = _neighbour(i + 1, "before")
        if holder_b is not None:
            adjacent.add(holder_b)
        if holder_a is not None:
            adjacent.add(holder_a)
        samples.append({
            "index": n,
            "before_pt": gap_b, "before_blank": blank_b,
            "after_pt": gap_a, "after_blank": blank_a,
            # v2.3.100（F-2 / 裁定 C-2）：两侧是否与表格**同页**。
            # 邻接段带 pageBreakBefore → 它与表格之间隔着分页 → 该侧跨页，不判。
            "before_same_page": not _starts_new_page(seq[i - 1]) if i > 0 else False,
            "after_same_page": not _starts_new_page(seq[i + 1]) if i + 1 < len(seq) else False,
        })
    return samples, adjacent


def _starts_new_page(el) -> bool:
    """元素（段落）是否是「新页的第一段」—— 带 `w:pageBreakBefore`。

    v2.3.100（F-2）：docx 没有绝对页坐标，判「表格与邻接文字是否同页」只能用这个近似。
    宁可漏判（把同页的当跨页而跳过），也不能把跨页的当同页去判一个没有意义的空隙。
    """
    try:
        if el is None or el.tag != qn("w:p"):
            return False
        ppr = el.find(qn("w:pPr"))
        if ppr is None:
            return False
        if ppr.find(qn("w:pageBreakBefore")) is not None:
            return True
        # 段内分页符（`<w:br w:type="page"/>`）同样说明它前面换了页
        for br in el.iter(qn("w:br")):
            if br.get(qn("w:type")) == "page":
                return True
        return False
    except Exception:
        return False


def sample_table_lines(doc) -> list:
    """采样每张表的边框现状，供 checker 判「是不是学术三线表」。

    三线表 = ① 有顶线 ② 有底线 ③ 表头下方有一条横线（≥2 行时）
             ④ **没有**任何竖线（left / right / insideV）
             ⑤ **没有**内部横线（insideH）。
    每张表返回一个 dict；读不出来就少给一条，绝不臆造。
    """
    out: list = []
    try:
        tables = list(doc.tables)
    except Exception:
        return out
    for idx, tbl in enumerate(tables):
        try:
            rows = list(tbl.rows)
        except Exception:
            continue
        if not rows:
            continue
        if len(rows) < 2:
            # 单行（含 1×1 布局/摘要框）无法形成「表头线」，不属于数据表。
            # 参与三线表判定会永远 warn 且修正器永远修不动 → 非幂等。跳过（N/A）。
            continue
        try:
            tblPr = tbl._tbl.tblPr
        except Exception:
            tblPr = None
        style_box = _table_style_borders(doc, tbl)
        first = _row_tc_prs(rows[0])
        last = _row_tc_prs(rows[-1])
        has_top = any(_border_visible(_effective_border(tblPr, style_box, p, "top"))
                      for p in first)
        has_bottom = any(_border_visible(_effective_border(tblPr, style_box, p, "bottom"))
                         for p in last)
        has_vertical = False
        has_inside_h = False
        for row in rows:
            for p in _row_tc_prs(row):
                if (_border_visible(_effective_border(tblPr, style_box, p, "left"))
                        or _border_visible(_effective_border(tblPr, style_box, p, "right"))
                        or _border_visible(_effective_border(tblPr, style_box, p, "insideV"))):
                    has_vertical = True
                if _border_visible(_effective_border(tblPr, style_box, p, "insideH")):
                    has_inside_h = True
        has_header_line = False
        if len(rows) >= 2:
            # v2.3.100（B-6 / 裁定 C-3）：表头行**任一**单元格有 bottom 即算有表头线。
            # 旧写法用 ``all`` —— 官方 APA 样例（qualitative / mixed-methods）把表头线
            # 写在**单元格级** tcBorders 上，且常常只给部分单元格写（合并单元格、
            # 首列不写），``all`` 于是判"没有表头线"→ 报 table_lines 不符（2 份连报）。
            # 裁定原文：「合并表级→行级→单元格级边框；表头行任一单元格有 bottom 即算有」。
            has_header_line = bool(first) and any(
                _border_visible(_effective_border(tblPr, style_box, p, "bottom"))
                for p in first)
        out.append({
            "index": idx + 1,                      # 给客户看的表号（从 1 起）
            "rows": len(rows),
            "cols": len(first),
            "has_top": has_top,
            "has_bottom": has_bottom,
            "has_header_line": has_header_line,
            "has_vertical": has_vertical,
            "has_inside_h": has_inside_h,
            "three_line": bool(has_top and has_bottom and not has_vertical
                               and not has_inside_h and has_header_line),
        })
    return out


def describe_table_lines(state: dict) -> str:
    """把一张表的边框现状说成人话（进报告 / 复核弹窗给客户看）。"""
    if not state:
        return "unknown"
    problems = []
    if state.get("has_vertical"):
        problems.append("vertical rules")
    if state.get("has_inside_h"):
        problems.append("inner horizontal rules")
    if not state.get("has_top"):
        problems.append("no top rule")
    if not state.get("has_bottom"):
        problems.append("no bottom rule")
    if not state.get("has_header_line"):
        problems.append("no rule under the header")
    if not problems:
        return "three-line (top / header / bottom, no vertical rules)"
    return ", ".join(problems)


# ---------------------------------------------------------------------------
# 公式（Word 原生 OMML）采样（v2.3.112）
# ---------------------------------------------------------------------------
# 为什么需要
# --------
# 学校模板里「公式居中、编号靠右、与正文同字号、段前段后 0」是高频率硬要求，但引擎
# 此前**完全读不到** Word 公式：OMML 的数学 run 是 ``m:r``，而 python-docx 的
# ``para.runs`` 只返回 ``w:r`` —— 公式段在引擎眼里永远是「除了编号一个字都没有」的
# 空段，任何维度都判不到它，修正器更无从下手。``report.KNOWN_LIMITATIONS`` 里那句
# 「Complex formulas / equations are NOT processed」说的就是这个空白。
#
# 🔴 只读**版式**：对齐 / 字号 / 段距 / 行距 / 编号位置。数学内容在 ``m:t`` 里，
#    本函数一个字节都不读进输出、更不写（policy.TEXT_IS_LOCKED：文字 100% 锁死）。

# 「这一段里除公式外只有一串编号」的形态：(1) / [1] / (2-1) / (3.4a)。
# 与 C-14 那套 _EQ_NUM_STYLE_RES 不同：那边收集的是「全文用过哪些编号形态」，
# 用来判「有没有混用」；这里要的是「这个公式段里剩下的文字是不是就一个编号」。
_EQ_NUMBER_ONLY_RE = re.compile(
    r"^[\s\u00a0]*[\(\[]\s*\d+(?:\s*[.\-\u2013]\s*\d+)*\s*[a-zA-Z]?\s*[\)\]][\s\u00a0]*$")

# m:jc 的取值 → 客户能理解的三档。centerGroup 是 Word 的「整组居中」，
# 版面上与居中完全一样，判等时必须归一，否则「明明居中却报未居中」。
_MATH_JC_ALIASES = {
    "center": "center",
    "centergroup": "center",
    "left": "left",
    "right": "right",
}


def _iter_all_paragraphs(doc):
    """按文档顺序产出 ``(Paragraph, in_table)``，含表格（及嵌套表格）单元格段落。

    公式常常出现在表格里 —— 「公式 + 编号」的真实版式很多时候就是一张 1×2 无框表
    （居中单元格放公式、右单元格放编号）。只走 ``doc.paragraphs`` 会漏掉这些，
    于是客户报「我的公式明明有问题，你却查不出来」。

    🔴 合并单元格会让同一个 ``w:tc`` 在多个 grid 列上重复出现，必须按 ``id(tc)``
    去重：否则一张合并表里的公式会被数成好几份（「共 3 个公式」显示成 9 个）。
    """
    try:
        for p in doc.paragraphs:
            yield p, False
    except Exception:
        pass
    seen: set = set()

    def _walk(tbl):
        try:
            rows = list(tbl.rows)
        except Exception:
            return
        for row in rows:
            try:
                cells = list(row.cells)
            except Exception:
                continue
            for cell in cells:
                try:
                    key = id(cell._tc)
                except Exception:
                    key = id(cell)
                if key in seen:
                    continue
                seen.add(key)
                try:
                    for p in cell.paragraphs:
                        yield p, True
                except Exception:
                    pass
                try:
                    for inner in cell.tables:
                        for item in _walk(inner):
                            yield item
                except Exception:
                    pass

    try:
        for tbl in list(doc.tables):
            for item in _walk(tbl):
                yield item
    except Exception:
        pass


def _eq_display_alignment(p_el) -> str:
    """展示公式（``m:oMathPara``）的**生效**对齐：center / left / right。

    优先级：``m:oMathParaPr/m:jc`` > ``w:pPr/w:jc`` > Word 缺省（居中）。

    ⚠️ 为什么 ``m:jc`` 必须优先：Word 里把一段展示公式改成左对齐，写的是 ``m:jc``，
    而 ``w:jc`` 往往还原样留着。若先读 ``w:jc``，「已经左对齐」会被读回「还居中」，
    修正器于是反复改、报告永远点不掉（自伤型误报）。
    """
    raw = None
    try:
        for op in p_el.findall(qn("m:oMathPara")):
            pr = op.find(qn("m:oMathParaPr"))
            if pr is None:
                continue
            jc = pr.find(qn("m:jc"))
            if jc is None:
                continue
            val = jc.get(qn("m:val"))
            if val:
                raw = str(val).strip().lower()
                break
    except Exception:
        raw = None
    if raw:
        return _MATH_JC_ALIASES.get(raw, raw)
    # 回退到段落对齐：只认能映射的三档；both/distribute（两端对齐）对单行公式没有
    # 实际意义，Word 仍按居中渲染 → 不采纳。
    try:
        ppr = p_el.find(qn("w:pPr"))
    except Exception:
        ppr = None
    if ppr is not None:
        try:
            jc = ppr.find(qn("w:jc"))
        except Exception:
            jc = None
        if jc is not None:
            val = str(jc.get(qn("w:val")) or "").strip().lower()
            if val in ("left", "start"):
                return "left"
            if val in ("right", "end"):
                return "right"
            if val == "center":
                return "center"
    # m:oMathParaPr 与 w:jc 都没写 → Word 对展示公式的渲染就是居中
    return "center"


def _eq_has_right_tab_stop(p_el) -> bool:
    """段落里是否有**右对齐制表位**。

    Word 的「公式 + 编号」标准版式 = 展示公式居中 + 右对齐制表位 + Tab + ``(1)``。
    有了这个制表位，编号才会被推到正文右边界；没有它，编号只是紧跟在公式后面。
    """
    try:
        ppr = p_el.find(qn("w:pPr"))
    except Exception:
        return False
    if ppr is None:
        return False
    try:
        tabs = ppr.find(qn("w:tabs"))
    except Exception:
        return False
    if tabs is None:
        return False
    for tab in tabs.findall(qn("w:tab")):
        val = str(tab.get(qn("w:val")) or "").strip().lower()
        if val in ("right", "end"):
            return True
    return False


def _eq_math_size_pt(p_el) -> Optional[float]:
    """公式内数学 run 的**显式**字号（众数，磅）。读不到 → None。

    只认 ``m:r/w:rPr/w:sz``：公式没写显式字号时会继承段落/样式，这一层回退交给
    调用方（它有 Paragraph 对象，能走样式链），本函数不猜。
    """
    sizes: Counter = Counter()
    try:
        math_runs = p_el.findall(".//" + qn("m:r"))
    except Exception:
        return None
    for mr in math_runs:
        try:
            rpr = mr.find(qn("w:rPr"))
        except Exception:
            continue
        if rpr is None:
            continue
        sz = rpr.find(qn("w:sz"))
        if sz is None:
            continue
        try:
            sizes[round(float(sz.get(qn("w:val"))) / 2.0, 1)] += 1
        except Exception:
            continue
    if not sizes:
        return None
    return sizes.most_common(1)[0][0]


def sample_equations(doc) -> list:
    """采样全文公式（OMML）的版式现状，供 checker 判 / fixer 修。

    每个「含公式的段落」一条记录：

    - ``index``            第几个公式段（从 1 起，给客户看的序号）
    - ``display``          是否含独立成行的展示公式（``m:oMathPara``）
    - ``display_count``    该段展示公式个数
    - ``inline_count``     该段行内公式个数（``m:oMath`` 直接挂在 ``w:p`` 下）
    - ``alignment``        展示公式生效对齐：center / left / right
    - ``size_pt``          公式字号（数学 run 显式 → 样式链 → 文档默认）
    - ``space_before_pt``  所在段段前（磅，``w:beforeLines`` 已折算）
    - ``space_after_pt``   所在段段后
    - ``line_spacing``     所在段行距倍数
    - ``number_text``      同段除公式外**只有**一串编号时的编号文本（如 "(1)"）
    - ``number_at_right``  编号是否已靠右（段右对齐，或右制表位 + Tab）
    - ``number_by_tab``    编号前是否已有 Tab 分隔（决定能否安全改成靠右）
    - ``in_table``         是否在表格单元格里

    🔴 只读版式，不读公式内容：``m:t`` 里的数学文本一个字节都不进结果。
    """
    out: list = []
    idx = 0
    for para, in_table in _iter_all_paragraphs(doc):
        try:
            p_el = para._p
        except Exception:
            continue
        if p_el is None:
            continue
        try:
            n_display = len(p_el.findall(qn("m:oMathPara")))
            n_inline = len(p_el.findall(qn("m:oMath")))
        except Exception:
            continue
        if not n_display and not n_inline:
            continue
        idx += 1

        # 段内除公式外的可见文字：python-docx 的 ``w:t`` **不含** ``m:t``，所以这里
        # 天然只捞到「编号」这类普通文字；正文段即使混排行内公式也不会被误当编号。
        try:
            plain = "".join(t.text or "" for t in p_el.findall(".//" + qn("w:t")))
        except Exception:
            plain = ""
        plain = (plain or "").strip()
        number_text = plain if plain and _EQ_NUMBER_ONLY_RE.match(plain) else None

        size_pt = _eq_math_size_pt(p_el)
        if size_pt is None:
            # 公式没写显式字号 → 走与正文完全相同的三级回退（段落样式链 → 文档默认）。
            # 少了这一层，「模板要求 12pt、公式继承 12pt」会被判成「读不到字号」。
            size_pt = (_style_chain_size_pt(para) or _font_size_hint(para)
                       or _doc_default_size_pt(doc))

        try:
            sb, sa = _spacing_pair_of_para(para)
        except Exception:
            sb, sa = None, None
        try:
            ls = _line_spacing_of_para(para)
        except Exception:
            ls = None
        try:
            align_para = _effective_para_alignment(para)
        except Exception:
            align_para = "left"

        try:
            has_tab = bool(p_el.findall(".//" + qn("w:tab")))
        except Exception:
            has_tab = False
        right_tab = _eq_has_right_tab_stop(p_el)
        at_right = (align_para == "right") or (has_tab and right_tab)

        out.append({
            "index": idx,
            "display": bool(n_display),
            "display_count": n_display,
            "inline_count": n_inline,
            "alignment": _eq_display_alignment(p_el) if n_display else None,
            "align_para": align_para,
            "size_pt": size_pt,
            "space_before_pt": sb,
            "space_after_pt": sa,
            "line_spacing": ls,
            "number_text": number_text,
            "number_at_right": bool(at_right),
            "number_by_tab": bool(has_tab),
            "in_table": bool(in_table),
        })
    return out


def describe_equation(state: dict) -> str:
    """把一个公式段的现状说成人话（进报告给客户看）。"""
    if not state:
        return "unknown"
    bits = []
    if state.get("display"):
        bits.append("display equation, %s" % (state.get("alignment") or "alignment unset"))
    else:
        bits.append("inline equation")
    if state.get("size_pt") is not None:
        bits.append("%s pt" % state["size_pt"])
    num = state.get("number_text")
    if num:
        bits.append("number %s%s" % (num, "" if state.get("number_at_right") else " (not at the right margin)"))
    else:
        bits.append("no equation number")
    return ", ".join(bits)


# ---------------------------------------------------------------------------
# 目录（TOC）识别（v2.3.77）—— 单一真相源，checker 与 fixer 共用
# ---------------------------------------------------------------------------
# 目录条目是**域段落**（每个条目带指向 _Toc 书签的超链接、段落里还有 PAGEREF 域），
# Word 在「更新目录(F9)」时会用 TOC 1..9 样式整体重排 —— 所以：
#   · 条目间距**不能**逐段改直接格式（改了下次更新目录就还原）；
#   · 能安全改的只有「目录标题那一段」（它不是域段落）；
#   · 条目间距只做提示，并明确告诉客户「要改请改 TOC 样式」。
_TOC_TITLE_RE = re.compile(r"(?i)^\s*(table\s+of\s+contents|contents|目\s*录|目\s*次|目\s*錄)\s*$")
_TOC_STYLE_ID_RE = re.compile(r"(?i)^toc\s*(\d)$")
# 目录标题样式：Word 内置 styleId（中文界面名是「目录标题」，但 styleId 恒为 TOCHeading）
_TOC_HEADING_STYLE_IDS = ("tocheading",)


def _style_id_of(para) -> str:
    """段落所用样式的 styleId（小写）。**绝不用 style.name** —— 中文 Word 会把
    「TOC 1」显示成「目录 1」，按名字识别必然漏判。"""
    try:
        st = para.style
        if st is None:
            return ""
        return str(getattr(st, "style_id", "") or "").strip().lower()
    except Exception:
        return ""


def toc_entry_level(para) -> Optional[int]:
    """该段是否为目录条目：TOC 1..9 样式 / 指向 _Toc 书签的超链接 / PAGEREF _Toc 域。"""
    m = _TOC_STYLE_ID_RE.match(_style_id_of(para))
    if m:
        return int(m.group(1))
    try:
        p = para._p
    except Exception:
        return None
    try:
        for h in p.iter(qn("w:hyperlink")):
            if (h.get(qn("w:anchor")) or "").startswith("_Toc"):
                return 1
        for it in p.iter(qn("w:instrText")):
            up = (it.text or "").upper()
            if "_TOC" in up or up.strip().startswith("TOC"):
                return 1
    except Exception:
        pass
    return None


def is_toc_heading(para) -> bool:
    """该段是否为「目录」标题段（Contents / 目录 …）。样式优先，文本兜底。"""
    if _style_id_of(para) in _TOC_HEADING_STYLE_IDS:
        return True
    if toc_entry_level(para) is not None:
        return False                      # 条目本身不算标题（防止把「目录 1」当标题）
    return bool(_TOC_TITLE_RE.match((para.text or "").strip()))


def detect_toc(doc):
    """识别目录结构，返回 ``{'heading': idx|None, 'entries': [idx, ...]}``；无目录返回 None。

    只认正面信号（TOC 样式 / _Toc 书签超链接 / PAGEREF 域），绝不按「像目录」瞎猜。
    """
    try:
        paras = list(doc.paragraphs)
    except Exception:
        return None
    entries = [i for i, p in enumerate(paras) if toc_entry_level(p) is not None]
    if not entries:
        return None
    heading = None
    # 目录标题：从第一条条目往前最多找 4 段（中间常夹一行「目录」或一个空段）
    for j in range(entries[0] - 1, max(-1, entries[0] - 5), -1):
        if is_toc_heading(paras[j]):
            heading = j
            break
    return {"heading": heading, "entries": entries}


def _sample_toc(doc, prof: "DocumentProfile") -> None:
    """采样目录：标题段与条目段的段前/段后（只读不判；判定在 checker）。"""
    info = detect_toc(doc)
    if not info:
        # 文档没有目录 → 所有字段保持 None（**不是 0**），检查据此给 info 而非 fail
        return
    prof.toc_detected = True
    prof.toc_entry_count = len(info["entries"])
    try:
        paras = list(doc.paragraphs)
    except Exception:
        return
    hi = info.get("heading")
    if isinstance(hi, int) and 0 <= hi < len(paras):
        prof.toc_heading_text = (paras[hi].text or "").strip()
        sb, sa = _spacing_pair_of_para(paras[hi])
        prof.toc_heading_space_before_pt = sb
        prof.toc_heading_space_after_pt = sa
    sb_c: Counter = Counter()
    sa_c: Counter = Counter()
    for i in info["entries"]:
        if not (0 <= i < len(paras)):
            continue
        b, a = _spacing_pair_of_para(paras[i])
        sb_c[0.0 if b is None else b] += 1
        sa_c[0.0 if a is None else a] += 1
    if sb_c:
        prof.toc_entry_space_before_pt = sb_c.most_common(1)[0][0]
    if sa_c:
        prof.toc_entry_space_after_pt = sa_c.most_common(1)[0][0]


# ---------------------------------------------------------------------------
# 脚注 / 尾注计数（v2.3.77）—— 只数不改
# ---------------------------------------------------------------------------
def count_note_entries(doc, tag: str) -> int:
    """数脚注（tag="w:footnote"）/ 尾注（tag="w:endnote"）的真实条目数。

    ⚠️ 必须跳过 ``w:id <= 0`` 与带 ``w:type`` 的特殊条目 —— Word 在 footnotes.xml
    里额外放三条不是正文脚注的东西（separator 分隔线 / continuationSeparator /
    continuationNotice），把它们数成「脚注」会让提示数量虚高。
    ⚠️ python-docx **没有脚注模型**：`part_related_by(RT.FOOTNOTES)` 拿到的是裸 Part
    （没有 ``.element``），只能 ``parse_xml(part.blob)`` 读 —— 我们只读不写，故不必回写。
    """
    rel = RT.FOOTNOTES if tag == "w:footnote" else RT.ENDNOTES
    try:
        part = doc.part.part_related_by(rel)
    except Exception:
        return 0
    if part is None:
        return 0
    try:
        root = parse_xml(part.blob)
    except Exception:
        return 0
    n = 0
    try:
        for el in root.findall(qn(tag)):
            if el.get(qn("w:type")):
                continue
            try:
                nid = int(el.get(qn("w:id")) or 0)
            except (TypeError, ValueError):
                nid = 0
            if nid <= 0:
                continue
            n += 1
    except Exception:
        return n
    return n


def _sample_notes(doc, prof: "DocumentProfile") -> None:
    prof.footnote_count = count_note_entries(doc, "w:footnote")
    prof.endnote_count = count_note_entries(doc, "w:endnote")


# ---------------------------------------------------------------------------
# 页眉 / 页码 / 参考文献分页 轻量读取（仅读不判；判定在 checker.py）
# ---------------------------------------------------------------------------
_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _para_has_page_field(para) -> bool:
    """段落内是否存在 PAGE 域（页码）。兼顾简单域 fldSimple 与复杂域 instrText。"""
    try:
        el = para._p
    except Exception:
        return False
    for fs in el.iter(f"{{{_W}}}fldSimple"):
        instr = fs.get(f"{{{_W}}}instr") or ""
        if "PAGE" in instr.upper():
            return True
    for it in el.iter(f"{{{_W}}}instrText"):
        if it.text and "PAGE" in it.text.upper():
            return True
    return False


def _para_text_excl_fields(para) -> str:
    """段落可见文本，但剔除域（PAGE 等）的结果文本，避免页码污染 running head 长度判断。

    实现：用 lxml 遍历——fldChar(begin)…fldChar(end) 区间内的 w:t，
    以及 fldSimple 内部的 w:t，全部跳过。
    """
    try:
        el = para._p
    except Exception:
        return para.text or ""
    # 标记区间内/域内的 w:t，避免使用 namespace 前缀等长字符串作为 key
    in_field = False
    for node in el.iter():
        tag = node.tag
        if tag == f"{{{_W}}}fldChar":
            ftype = node.get(f"{{{_W}}}fldCharType")
            if ftype == "begin":
                in_field = True
            elif ftype == "end":
                in_field = False
        elif tag == f"{{{_W}}}fldSimple":
            for t in node.iter(f"{{{_W}}}t"):
                t.set("data-skip", "1")
        elif tag == f"{{{_W}}}t" and in_field:
            node.set("data-skip", "1")
    parts = []
    for t in el.iter(f"{{{_W}}}t"):
        if t.get("data-skip") != "1":
            parts.append(t.text or "")
    return "".join(parts).strip()


# 手打静态页码的识别式：整段只有一个数字，允许「1」「Page 1」「第 1 页」「- 1 -」等前后缀。
# 页码最多三位（999 页），四年份（2020）不会被误当页码。
# 🔴 刻意**不认**「Page 1 of 10」这类带总页数的：就地替换要同时把总页数换成 NUMPAGES 域，
#    超出保守范围 —— 宁可不自动改（只提示客户手改），也不能把总页数写成死数字。
_STATIC_PAGE_RE = re.compile(
    r"^[\s\u00a0]*"
    r"(?P<pre>(?:(?:page|p\.|第)\s*)?[-–—]?\s*)"
    r"(?P<num>\d{1,3})"
    r"(?P<suf>\s*(?:页|of|/)?\s*[-–—]?)"
    r"[\s\u00a0]*$",
    re.IGNORECASE,
)

_ALIGN_NAMES = {
    WD_ALIGN_PARAGRAPH.LEFT: "left",
    WD_ALIGN_PARAGRAPH.CENTER: "center",
    WD_ALIGN_PARAGRAPH.RIGHT: "right",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "both",
}


def detect_static_page_number(text: str) -> Optional[dict]:
    """整段是否只是「手打页码」；是则返回 ``{"prefix", "number", "suffix"}``，否则 None。

    只在页眉/页脚里调用。要求整段**只有一个数字**（见 ``_STATIC_PAGE_RE`` 注释）。
    返回的前后缀用于修正时**原地替换**：把数字换成 PAGE 域、前后缀原样保留，
    这样「第 1 页」会变成「第 {PAGE} 页」，既不会出现两个页码，也不丢客户格式。
    """
    if not text:
        return None
    t = text.strip()
    if not t:
        return None
    if sum(ch.isdigit() for ch in t) > 3:
        return None
    m = _STATIC_PAGE_RE.match(t)
    if not m:
        return None
    return {"prefix": m.group("pre"), "number": m.group("num"),
            "suffix": m.group("suf")}


def _effective_para_alignment(para) -> str:
    """段落的**有效**对齐：直接格式 → 样式 → 默认左。

    🔴 旧版在直接格式为空时把对齐读成 None，checker 的 ``elif ... is False`` 分支因此
    不成立、直接落进 else 报 pass —— 客户页码在页脚居中却被判「右上，符合」。这里统一
    归一到具体值：一个声明都没有时就是左对齐（Word 的默认），不存在「无法判定」。
    """
    al = None
    try:
        al = para.alignment
    except Exception:
        al = None
    if al is None:
        try:
            st = para.style
            if st is not None:
                al = st.paragraph_format.alignment
        except Exception:
            al = None
    if al is None:
        return "left"
    return _ALIGN_NAMES.get(al, "left")


@dataclass
class PageNumberInfo:
    """页码读取结果（页眉/页脚全口径）。"""

    has_field: bool = False                 # 存在 PAGE 域（自动页码）
    has_static: bool = False                # 手打静态页码（无域）
    location: Optional[str] = None          # "header" / "footer"
    alignment: Optional[str] = None         # "left"/"center"/"right"/"both"（有效对齐）
    static_text: str = ""                   # 手打页码原文

    @property
    def present(self) -> bool:
        return self.has_field or self.has_static


def _part_paragraphs(part) -> list:
    """页眉/页脚 part 的段落列表，**含被 ``w:sdt`` 内容控件包住的段落**（v2.3.100，B-1）。

    🔴 为什么需要它：Word 的「插入页码」会把页码段落整个包进一个**内容控件**
    （``<w:sdt>``），python-docx 的 ``part.paragraphs`` 只认直接子级的 ``<w:p>``，
    于是这些页码段落**完全看不见**。实测 13 份官方 APA 样例里 9 份的页码就是这样
    被包着的 → 我们报「No page number found in the header or footer」；更糟的是
    「一键修正」会**再插一个页码**，客户文档里出现两个页码。

    官方 IEEE 模板的页眉同理（``header1.xml`` 顶层就是 ``<w:sdt>``）。

    实现：按文档序遍历 part 的**直接子级** ``w:p``，以及 ``w:sdt/w:sdtContent``
    里的 ``w:p``。不递归进 ``w:tbl``（表格单元格里的文字不是页眉文字）；
    不去重（同一段落不会被两条路径同时命中 —— sdt 与直接子级互斥）。
    """
    out: list = []
    try:
        root = part._element if hasattr(part, "_element") else part
    except Exception:
        return list(getattr(part, "paragraphs", []) or [])
    try:
        from docx.text.paragraph import Paragraph
        for child in root.iterchildren():
            tag = child.tag
            if tag == qn("w:p"):
                out.append(Paragraph(child, part))
            elif tag == qn("w:sdt"):
                # 内容控件：取其 sdtContent 下的段落（只一层，不递归嵌套 sdt）
                content = child.find(qn("w:sdtContent"))
                if content is None:
                    continue
                for sub in content.iterchildren():
                    if sub.tag == qn("w:p"):
                        out.append(Paragraph(sub, part))
    except Exception:
        return list(getattr(part, "paragraphs", []) or [])
    return out


def _iter_header_footer_parts(doc):
    """产出 ``(part, kind, scope)``：kind="header"/"footer"，scope="main"/"first"/"even"。

    🔴 首页 / 偶数页只在**文档真正启用**时才读：
    - 首页：``section.different_first_page_header_footer`` 为 True；
    - 偶数页：``settings.xml`` 里存在 ``<w:evenAndOddHeaders/>``。
    否则那些 part 即使留着内容也不会被渲染，读了会误报「有页码」。
    （旧版完全没读首页 → APA 标题页页码放在首页页眉时被报「没有页码」。）
    """
    try:
        sections = list(doc.sections)
    except Exception:
        return
    even_on = False
    try:
        even_on = doc.settings.element.find(qn("w:evenAndOddHeaders")) is not None
    except Exception:
        even_on = False
    for sec in sections:
        try:
            diff = bool(sec.different_first_page_header_footer)
        except Exception:
            diff = False
        for kind in ("header", "footer"):
            main = getattr(sec, kind, None)
            if main is not None:
                yield main, kind, "main"
            if diff:
                fp = getattr(sec, "first_page_" + kind, None)
                if fp is not None:
                    yield fp, kind, "first"
            if even_on:
                ev = getattr(sec, "even_page_" + kind, None)
                if ev is not None:
                    yield ev, kind, "even"


def _page_field_pushed_right_by_tab(para) -> bool:
    """页码是否靠「右制表位」被推到右边界（APA 专业版页眉的标准写法）。

    APA 7th 专业版的页眉是**同一段**里「短标题<TAB>页码」：段落对齐必须是
    left（否则左制表位上的短标题会被一起甩到右边），页码靠 ``w:tab w:val="right"``
    定位到正文右边界。若只看 ``w:jc``，这种**完全合规**的页眉会被判成
    ``left-aligned`` —— v2.3.90 实测：修正器刚写好的 APA 专业版页眉，复检立刻报
    「Page number is left-aligned」，客户会看到「修了还是错」。

    判定：段内有 right 制表位，且 PAGE 域之前出现过 ``<w:tab/>``。
    """
    # 🔴 v2.3.100（B-1）：**递归**找制表位与 run，并且制表位要沿**样式链**找。
    # 实测官方 APA 专业版样例的页眉段落：Word 把「短标题 + 若干 run + 页码」整段包进
    # 段落内的 ``<w:sdt>``（内容控件），run 全在 ``sdtContent`` 里 ——
    # ``para._p.findall("w:r")`` 只认直接子级，**一个都找不到**；
    # 而右制表位根本不写在段落上，而是写在 **Header 样式**里（Word 内置样式自带
    # center + right 两个制表位）。两处都按"直接子级"找 → 恒判 False →
    # 完全合规的 APA 专业版页眉被报「Page number is left-aligned」（5 份样例连报）。
    def _tabs_in(el) -> list:
        try:
            ppr = el.find(qn("w:pPr"))
        except Exception:
            return []
        if ppr is None:
            return []
        tabs = ppr.find(qn("w:tabs"))
        if tabs is None:
            return []
        return list(tabs.findall(qn("w:tab")))

    tab_els: list = []
    try:
        tab_els.extend(_tabs_in(para._p))
        # 样式链：段落样式 → base_style（右制表位常定义在 Header 样式上）
        st = para.style
        seen = set()
        while st is not None and st.style_id not in seen:
            seen.add(st.style_id)
            el = getattr(st, "element", None)
            if el is not None:
                tab_els.extend(_tabs_in(el))
            st = getattr(st, "base_style", None)
    except Exception:
        pass
    has_right_tab = any((t.get(qn("w:val")) or "") == "right" for t in tab_els)
    # 逐个走 run（**递归**，含 sdtContent 内的 run）：
    # 见到 <w:tab/> 记一笔，见到 PAGE 域的 fldChar(begin) 就结算。
    # ⚠️ v2.3.100：两条路径都要走到底，不能因为有/没有 right 制表位就提前 return ——
    # 「连续多个 tab」这条路在**没有** right 制表位时也成立（见下方注释）。
    seen_tab = False
    tab_count = 0
    try:
        for r in para._p.iter(qn("w:r")):
            for el in r:
                if el.tag == qn("w:tab"):
                    seen_tab = True
                    tab_count += 1
                elif el.tag == qn("w:fldChar") \
                        and (el.get(qn("w:fldCharType")) or "") == "begin":
                    if has_right_tab and seen_tab:
                        return True
                    # v2.3.100（B-1 收尾）：**连续多个 tab 也是右推**。
                    # Word 的「Page Numbers (Top of Page)」构件常写成
                    # 「短标题 + 5 个以上 <w:tab/> + 页码」（官方 mixed-methods 样例即此：
                    # 段内 6 个 tab、段落与 Header 样式里都没有 right 制表位），
                    # 渲染上页码确实落在右端 → 旧判据漏掉它，报「left-aligned」。
                    # 阈值取 2：单个 tab 在没有右制表位时只跳到默认制表位（0.5"），
                    # 不足以断定"到右边界"；≥2 个连续 tab 才是一路顶到右边的写法。
                    return tab_count >= 2
    except Exception:
        return False
    return False


def _page_number_alignment(para) -> str:
    """页码段的「有效对齐」——左对齐但靠右制表位推到右边界，等同于右对齐。"""
    al = _effective_para_alignment(para)
    if al != "right" and _page_field_pushed_right_by_tab(para):
        return "right"
    return al


def read_page_number_info(doc) -> PageNumberInfo:
    """读页码：有无（域/手打）+ 在页眉还是页脚 + 有效对齐。

    页眉与页脚同时有页码时以**页眉**为准 —— page_number_position 的 top-right 语义
    指的是页眉右上（APA 7th / MLA 9th 都是）。
    """
    info = PageNumberInfo()
    for part, kind, _scope in _iter_header_footer_parts(doc):
        try:
            para_list = _part_paragraphs(part)
        except Exception:
            continue
        for p in para_list:
            if _para_has_page_field(p):
                if not info.has_field or (kind == "header" and info.location != "header"):
                    info.has_field = True
                    info.location = kind
                    info.alignment = _page_number_alignment(p)
                continue
            txt = _para_text_excl_fields(p)
            st = detect_static_page_number(txt)
            if st is not None:
                if not info.has_static:
                    info.has_static = True
                    info.static_text = txt.strip()
                if not info.has_field and info.location is None:
                    info.location = kind
                    info.alignment = _effective_para_alignment(p)
    return info


def _read_header_footer_info(doc) -> tuple[str, bool, Optional[bool]]:
    """返回 (running_head_text, page_number_present, page_number_right_aligned)。

    running_head_text：拼接主分节页眉文本并剔除 PAGE 域结果（页脚不计入）；
    page_number_present：页眉/页脚存在 PAGE 域**或**手打页码；
    page_number_right_aligned：页码段落是否右对齐（已按样式继承归一，不再出现 None）。
    """
    texts: list = []
    for part, kind, scope in _iter_header_footer_parts(doc):
        if kind != "header" or scope != "main":
            continue
        try:
            para_list = _part_paragraphs(part)
        except Exception:
            continue
        for p in para_list:
            t = _para_text_excl_fields(p)
            if t:
                texts.append(t)
    info = read_page_number_info(doc)
    right: Optional[bool] = None
    if info.present:
        right = (info.alignment == "right")
    return " ".join(texts).strip(), info.present, right


def _para_page_break_before(para) -> bool:
    """当前段落是否「段前分页」（pPr/pageBreakBefore）。"""
    try:
        pf = para.paragraph_format
        if pf is not None and getattr(pf, "page_break_before", None):
            return True
    except Exception:
        pass
    return False


def _para_ends_with_page_break(para) -> bool:
    """上一段末尾是否带分页符（w:br w:type='page'）。"""
    try:
        for br in para._p.iter(qn("w:br")):
            if br.get(qn("w:type")) == "page":
                return True
    except Exception:
        pass
    return False


def _run_bool_with_style(runs: list, para, attr: str):
    """run 级三态布尔（bold / italic）取值：**run 直接格式 → 段落样式链**（v2.3.100，B-5）。

    🔴 为什么需要它：官方 APA 样例的 "Abstract" 标题用 **Heading 1 样式**承载，
    加粗/斜体写在**样式**里，run 上什么都没有。旧写法只读 ``r.font.bold`` →
    全部读成 None → 判定"读不到" → 报「Abstract heading: could not read bold」
    （7 份官方样例连报同一条），而修正器其实按样式判得了、能改 —— 检查与修正脱节。

    三态语义必须保住：``None`` = 真的没设置（run 与样式链都没写），不能与
    ``False`` 混 —— 混了就会把"未规定"当成"规定不要加粗"，从而改错客户文档。

    返回：``True`` / ``False`` / ``None``。
    """
    vals = []
    for r in runs:
        try:
            vals.append(getattr(r.font, attr, None))
        except Exception:
            vals.append(None)
    if not vals:
        return None
    if any(v is True for v in vals):
        return True
    if all(v is False for v in vals):
        return False
    # 全部 None → 查**字符样式**（rStyle），再查段落样式链（含 base_style）。
    # v2.3.100（E-2 续）：B-5 当时留的尾巴「rStyle（字符样式）要一起解析，
    # 否则'用字符样式加粗'仍读不到」—— 这次补上。官方样例里就有用字符样式
    # （如 Strong / Emphasis）做加粗斜体的段落，只查段落样式链会读不到。
    try:
        for r in runs:
            rpr = r._r.find(qn("w:rPr"))
            if rpr is None:
                continue
            rstyle = rpr.find(qn("w:rStyle"))
            if rstyle is None:
                continue
            sid = rstyle.get(qn("w:val"))
            if not sid:
                continue
            _doc = getattr(getattr(r, "part", None), "document", None)
            if _doc is None:
                continue
            _cstyle = _doc.styles.element.find(
                qn("w:style") + "[@" + qn("w:styleId") + "='%s']" % sid)
            if _cstyle is None:
                continue
            _rpr = _cstyle.find(qn("w:rPr"))
            if _rpr is None:
                continue
            _el = _rpr.find(qn("w:b") if attr == "bold" else qn("w:i"))
            if _el is None:
                continue
            _val = _el.get(qn("w:val"))
            if _val is None:
                return True
            if str(_val).strip().lower() in ("0", "false", "off"):
                return False
            return True
    except Exception:
        pass
    # 全部 None → 查段落样式链（含 base_style）
    try:
        st = para.style
    except Exception:
        return None
    seen = set()
    while st is not None and st.style_id not in seen:
        seen.add(st.style_id)
        try:
            v = getattr(st.font, attr, None) if st.font is not None else None
        except Exception:
            v = None
        if v is True:
            return True
        if v is False:
            return False
        st = getattr(st, "base_style", None)
    return None


def _para_starts_with_page_break(para) -> bool:
    """段落**自身开头**是否带分页符（v2.3.100，B-3 / D-04）。

    🔴 实测官方 APA 专业版样例（``qualitative-professional-paper.docx`` 的 "Abstract"
    段）的分页写法是：

        ``<w:p><w:r><w:br w:type="page"/></w:r><w:r>Abstract</w:r></w:p>``

    即**段首一个分页符 run，然后才是标题文字**。这种写法三个旧判据全部落空：
      · 不是 ``pageBreakBefore``（``_para_page_break_before`` 判不到）；
      · 分页符不在**上一段**末尾（``_para_ends_with_page_break(prev)`` 判不到）；
      · 段落自身有文字，不会被当成"空段"跳过。
    → 明明另起了一页，我们报「Abstract continues on the same page」，
      而"修正"还会再插一个分页 → **客户文档多出一张空白页**（B-4/D-05）。

    判定：按文档序找第一个含 ``w:t`` 的 run，若在它之前出现过 ``w:br w:type="page"``
    即算。纯空段（没有任何文字）不算 —— 那种由 ``_para_ends_with_page_break`` 管。
    """
    try:
        seen_text = False
        for r in para._p.iter(qn("w:r")):
            has_text = any(t.text for t in r.iter(qn("w:t")) if t.text)
            if has_text:
                seen_text = True
                continue
            if seen_text:
                # 文字都出现了还往后找就没有意义（段首才叫"另起一页"）
                break
            for br in r.iter(qn("w:br")):
                if br.get(qn("w:type")) == "page":
                    return True
    except Exception:
        return False
    return False


def _para_rendered_at_page_top(para) -> bool:
    """段落是否**上次渲染时就在页首**（v2.3.100，B-3 收尾）。

    Word 会在每个"上次保存时恰好位于页首"的段落第一个 run 里写一个
    ``<w:lastRenderedPageBreak/>`` 标记。它是**渲染痕迹**、不是分页控制
    （客户之后编辑排版，Word 会重算它），但对我们回答的"这一页是不是另起的"
    恰恰是最直接的证据 —— 实测官方 ``quantitative-professional-paper.docx`` 的
    "Abstract" 段就是这样：没有 pageBreakBefore、没有 br type=page，
    只有一个 lastRenderedPageBreak，而它**确实**渲染在新页上。
    没有这条证据时报「continues on the same page」= 把官方合规样例判红。

    判定：文字出现之前有 ``w:lastRenderedPageBreak``。
    """
    try:
        for r in para._p.iter(qn("w:r")):
            for el in r:
                if el.tag == qn("w:t") and (el.text or "").strip():
                    return False          # 文字出现了还没见到标记 → 不在页首
                if el.tag == qn("w:lastRenderedPageBreak"):
                    return True
    except Exception:
        return False
    return False


def _reference_starts_new_page(prev_para, cur_para) -> bool:
    """参考文献标题 / 摘要标题是否另起一页。

    四条路径（v2.3.100 补齐后两条）：
      ① 当前段 ``pageBreakBefore``；
      ② 上一段末尾分页符（分页符单独占一个空段的写法）；
      ③ **当前段自身开头**带分页符（官方 APA 样例的标准写法）；
      ④ 当前段**上次渲染就在页首**（``lastRenderedPageBreak``，无分页控制但
         确实另起了一页的官方样例 —— quantitative-professional 即此）。
    """
    if cur_para is not None:
        if _para_page_break_before(cur_para):
            return True
        if _para_starts_with_page_break(cur_para):
            return True
        if _para_rendered_at_page_top(cur_para):
            return True
    if prev_para is not None and _para_ends_with_page_break(prev_para):
        return True
    return False


def _scan_front_matter(paras) -> tuple[bool, bool]:
    """保守检测文档是否含 Abstract/摘要 与 Author Note/作者注 段落（专业版要求）。

    启发式：短段落（<40 字，通常作为标题/标签）中包含关键词即判定。
    仅用于 info 级提示，误判不修改内容、不影响其他维度。
    """
    has_abs = False
    has_auth = False
    for p in paras:
        t = (p.text or "").strip().lower()
        if not t:
            continue
        if len(t) <= 40 and ("abstract" in t or "摘要" in t):
            has_abs = True
        else:
            # v2.3.100（C 批 D-08）：**期刊式同段摘要** `Abstract—正文…`（IEEE 官方
            # 会议模板就是这种），整段很长（>40 字），旧规则直接跳过 → has_abstract=False
            # → 摘要维度只能给 info「No Abstract section detected」，9pt/加粗规则无从生效。
            # 只认正面信号：文本以 Abstract / 摘要 起头，或段落样式名含 abstract。
            if re.match(r"^(?:abstract|摘\s*要)\b", t) or t.startswith(("abstract", "摘要")):
                has_abs = True
            try:
                if "abstract" in (p.style.name or "").lower():
                    has_abs = True
            except Exception:
                pass
        if len(t) <= 40 and ("author note" in t or "作者注" in t):
            has_auth = True
    return has_abs, has_auth


# —— 摘要页关键词行（v2.3.96）——
# APA 摘要页的 "Keywords: ..." 行属于摘要页块（分页锚点必须落在它**之后**），
# 但它的**样式口径与摘要正文相反**：摘要正文顶格、Keywords 行缩进 0.5″。
# 所以：定位摘要块时它算在内，采样/施加「摘要首行缩进」时必须排除它，
# 否则取众数会把 Keywords 行一起刷成顶格（真实论文里踩到过）。
_KEYWORDS_LINE_RE = re.compile(
    r"^\s*(?:key\s*words?|关键词)\s*[:：]", re.IGNORECASE)
_KEYWORDS_LINE_MAX_CHARS = 400       # 关键词行可能很长，但绝不是整段正文


def is_keywords_line(para) -> bool:
    """该段是否是摘要页的关键词行（"Keywords: ..." / "Key words: ..." / "关键词：..."）。

    只认行首标签 + 冒号，不做子串匹配（正文里提到 "keywords" 的句子不算）。
    """
    if para is None:
        return False
    text = (para.text or "").strip()
    if not text or len(text) > _KEYWORDS_LINE_MAX_CHARS:
        return False
    return bool(_KEYWORDS_LINE_RE.match(text))


# —— 摘要标题文字回退的严格边界（v2.3.96）——
# 真实学生论文里 "Abstract" 极常见是 Normal 手排、压根没套标题样式；只认样式级标题
# 就会一律「摘要定位不到」，连带「摘要另起一页 / 正文另起一页 / 摘要顶格」三件事
# 全都只能提示、永远改不了。故补一条**严格限定**的文字回退：
#   · 只在前 15 段内找（摘要标题不可能出现在正文深处）；
#   · 整行匹配（去空格、小写后**完全等于** abstract，不做子串匹配）；
#   · 长度上限 40 字符（长句即便以 Abstract 开头也不认）；
#   · **必须**在后继段落里找到 Keywords 行作为摘要块的右边界。
# 最后一条是安全底线：没有右边界的「摘要」无法确定正文从哪一段开始，
# 而正文另起一页是**往某一段写分页**，写错就是把客户版式改坏 → 宁可不修。
_ABS_FALLBACK_WINDOW = 15            # 只在前 15 段内找摘要标题
_ABS_FALLBACK_MAX_CHARS = 40         # 摘要标题行长度上限
_ABS_KEYWORDS_LOOKAHEAD = 12         # 从标题往后最多看 12 段找 Keywords 行


def find_abstract_section(paras) -> tuple[Optional[int], list[int]]:
    """定位摘要段：返回 (摘要标题段下标, 摘要正文段下标列表)。

    **单一真相源**：checker（判摘要格式）与 fixer（施加摘要格式）共用本函数，
    避免「检查判这一段、修正改另一段」的漂移（与 find_reference_section 同理）。

    识别规则（保守，仅认正面信号）：
    - ① 摘要标题 = 样式级标题（heading_level 非 None）且文本（去空格、小写）等于
      "abstract" / "摘要" / "abstract page"（APA 常见 "Abstract" 一级标题）。
      摘要正文 = 标题之后、下一个样式级标题之前的所有非空段。
    - ② v2.3.96 文字回退：① 找不到时，才在前 15 段内找**整行等于** "Abstract"
      且没套标题样式的段，并**必须**在其后找到 Keywords 行作为右边界；
      摘要正文 = 标题之后、Keywords 行（含）之前的所有非空段。
    - ①② 都找不到时返回 (None, [])，绝不把加粗正文误判成摘要标题。

    返回的 body **包含** Keywords 行（分页锚点取 ``max(body) + 1``，必须落在它之后）；
    需要「摘要正文」的样式口径时请用 :func:`abstract_body_style_indices` 去掉关键词行。
    """
    paras = list(paras)

    # ① 样式级标题（首选；行为与 v2.3.67 起完全一致）
    for i, para in enumerate(paras):
        text = (para.text or "").strip()
        low = text.lower().replace(" ", "")
        if low in ("abstract", "摘要", "abstractpage") and heading_level(para) is not None:
            body: list[int] = []
            for j in range(i + 1, len(paras)):
                pj = paras[j]
                if heading_level(pj) is not None:
                    break
                if (pj.text or "").strip():
                    body.append(j)
            return i, body

    # ①-b v2.3.100（C 批 D-08）：**期刊式同段摘要** —— `Abstract—正文…` 写在一段里，
    # 且用专属段落样式（IEEE 官方会议模板就是 `<w:pStyle w:val="Abstract"/>` +
    # `Abstract—This electronic document…`）。旧逻辑要求"整行等于 Abstract"，认不出
    # 这类格式 → 整份 IEEE 模板被判「没有摘要」，D-08 的 9pt/加粗规则无从落地。
    # 只认**正面信号**：样式名含 abstract 且该段文本以 Abstract 起头（不靠"像摘要"瞎猜）。
    for i, para in enumerate(paras):
        text = (para.text or "").strip()
        if not text:
            continue
        try:
            _sn = (para.style.name or "").strip().lower()
        except Exception:
            _sn = ""
        if "abstract" not in _sn:
            continue
        if not re.match(r"(?i)^abstract\b", text):
            continue
        body = [i]                     # 标题与正文同段：这一段既是标题也是正文
        for j in range(i + 1, len(paras)):
            pj = paras[j]
            if heading_level(pj) is not None:
                break
            if re.match(r"(?i)^(index\s*terms|keywords?)\b", (pj.text or "").strip()):
                body.append(j)         # 含关键词行（分页锚点要用 max(body)+1）
                break
            if (pj.text or "").strip():
                body.append(j)
        return i, body

    # ② 文字回退（严格限定，见上方常量注释）
    for i, para in enumerate(paras[:_ABS_FALLBACK_WINDOW]):
        text = (para.text or "").strip()
        if not text or len(text) > _ABS_FALLBACK_MAX_CHARS:
            continue
        if heading_level(para) is not None:
            continue                     # ① 已覆盖，这里只处理没套样式的
        if text.lower().replace(" ", "") not in ("abstract", "摘要", "abstractpage"):
            continue
        body = []
        real_body = 0                    # Keywords 行之前真正的摘要正文段数
        _end = min(len(paras), i + 1 + _ABS_KEYWORDS_LOOKAHEAD)
        for j in range(i + 1, _end):
            pj = paras[j]
            if heading_level(pj) is not None:
                break                    # 撞到下一个样式标题 → 右边界不可知，放弃
            if not (pj.text or "").strip():
                continue
            body.append(j)
            if is_keywords_line(pj):
                # 找到右边界。摘要正文为空（标题后直接跟 Keywords）→ 不可信，放弃
                return (i, body) if real_body else (None, [])
            real_body += 1
        # 这一段候选没找到 Keywords 行 → 右边界不可知，继续看后面还有没有候选
    return None, []


def abstract_body_style_indices(paras, body) -> list[int]:
    """从摘要块下标里剔掉 Keywords 行，只留真正的摘要正文段（供样式口径的采样/施加用）。

    APA：摘要正文顶格、Keywords 行缩进 0.5″。两者口径相反，混在一起取众数/一起改
    都会错，故采样（reader）与修正（fixer）统一走本函数。
    若剔除后什么都不剩（异常文档），退回原列表，宁可少判也不凭空丢段。
    """
    paras = list(paras)
    kept = [j for j in body
            if 0 <= j < len(paras) and not is_keywords_line(paras[j])]
    return kept or list(body)


# 图表 caption 识别（v2.3.67）：Word 内置 "Caption" 样式 + 文本前缀启发式双保险。
# 很多学生用「Table 1」「Figure 2」「表 1」「图 2」作为标题段，但没套 Caption 样式；
# 只认样式会漏掉一大半。前缀式兜底（保守：前缀后必须跟数字或标题文字）。
_CAPTION_STYLE_KEYWORDS = ("caption", "图注", "表注", "图表标题", "captions",
                           # v2.3.100（C 批 D-15）：IEEE 会议模板用
                           # `<w:pStyle w:val="tablehead"/>`（表题）与 `figurecaption`（图题）
                           "tablehead", "table head", "table title")
_CAPTION_PREFIX_RE = re.compile(
    # v2.3.100（F-2）：允许 **字母+数字**的编号（"Table D1"、"Fig. A2"）——
    # APA 官方 design 模板的表题就是 "Table D1"，旧正则只认纯数字 → 认不出表题
    # → 表格上方的空隙量到表题自身（0pt）→ 把官方合规样例判红。
    # 前瞻 `(?=[0-9a-z]*\d)` 保证编号里**至少有一个数字**，否则 "Tables are useful"、
    # "Table of Contents" 也会被当表题（实测踩过）。
    r"^(?:table|figure|fig\.|tab\.|表|图)\s*(?=[0-9a-z]*\d)[0-9a-z]+\b", re.IGNORECASE)

# v2.3.98：正文**提及句**不是表题。
# 真实论文里 "Table 1 below shows the detailed parameter estimates and statistical
# significance tests." 这种句子以 "Table 1" 开头，旧判定一律当表题 →
# ① `caption_position_table` 对一张**根本没有表号表题**的表格报「位置正确」（假绿）；
# ② 修正器还会去改这句正文的字号/行距（把正文当表题刷）。
# 故：标签后紧跟「下面的/上面的/展示了…」这类谓语 → 判为正文提及句，不是表题。
_CAPTION_MENTION_RE = re.compile(
    r"^\s*(?:table|figure|fig\.|tab\.|表|图)\s*\d+\s*"
    r"(?:below|above|shows?|lists?|presents?|summari[sz]es?|reports?|"
    r"displays?|provides?|illustrates?|contains?|gives?)\b",
    re.IGNORECASE)


def is_caption_paragraph(para) -> bool:
    """判断一段是否是图表标题（Caption）。

    双保险：① 样式名含 "caption" 关键词；② 文本以 "Table N" / "Figure N" / "表 N" /
    "图 N" 开头（学生常不套 Caption 样式）。保守：前缀后必须跟数字，避免把普通句子误判。
    """
    try:
        style_name = (para.style.name or "").lower() if para.style else ""
    except Exception:
        style_name = ""
    if any(k in style_name for k in _CAPTION_STYLE_KEYWORDS):
        return True
    text = (para.text or "").strip()
    if not text:
        return False
    if not _CAPTION_PREFIX_RE.match(text):
        return False
    # v2.3.98：以 "Table 1" 开头的**句子**（"Table 1 below shows…"）是正文提及句，
    # 不是表题 —— 否则"没有表号表题"会被判成"表题位置正确"，还会被当 caption 改字号。
    return not _CAPTION_MENTION_RE.match(text)


def _sample_captions(doc, prof: "DocumentProfile") -> None:
    """采样图表标题（Caption 段落）的字号 / 行距 / 与表格的相对位置。

    仅读不判；判定在 checker.py。字号/行距取众数（多张图表标题通常同口径）；
    表题位置遍历 body 元素顺序，看 caption 段落与最近 w:tbl 的前后关系。
    """
    try:
        paras = list(doc.paragraphs)
    except Exception:
        return
    cap_paras = [p for p in paras if is_caption_paragraph(p)]
    if not cap_paras:
        return
    prof.caption_count = len(cap_paras)
    sizes: Counter = Counter()
    spacings: list = []
    cap_sb: Counter = Counter()
    cap_sa: Counter = Counter()
    for p in cap_paras:
        sz = _max_run_size_pt(p)
        if sz is None:
            # v2.3.100（C 批 D-15）：run 没有显式字号时回退到**样式字号**。
            # IEEE 模板的 `tablehead` / `figurecaption` 样式把 `<w:sz w:val="16"/>`（8pt）
            # 写在样式里，run 上一片空白 → 旧逻辑读不到 → caption_size 只能给 info
            # 「Could not read caption size」，8pt 规则形同虚设。
            try:
                _st = p.style
                if _st is not None and _st.font is not None and _st.font.size is not None:
                    sz = round(float(_st.font.size.pt), 1)
            except Exception:
                sz = None
        if sz is not None:
            sizes[sz] += 1
        sp = _line_spacing_of_para(p)
        if sp is not None:
            spacings.append(sp)
        # v2.3.76：题与图表之间、题与正文之间的间距（认「行」单位）
        _b, _a = _spacing_pair_of_para(p)
        cap_sb[0.0 if _b is None else _b] += 1
        cap_sa[0.0 if _a is None else _a] += 1
    if sizes:
        prof.caption_font_size_pt = sizes.most_common(1)[0][0]
    if spacings:
        prof.caption_line_spacing = Counter(
            round(s, 2) for s in spacings).most_common(1)[0][0]
    if cap_sb:
        prof.caption_space_before_pt = cap_sb.most_common(1)[0][0]
    if cap_sa:
        prof.caption_space_after_pt = cap_sa.most_common(1)[0][0]
    # 表题位置：遍历 body 子元素，caption 段落（w:p）与最近 w:tbl 的前后关系。
    # 只记「caption 段落紧邻表格」的情况（中间无其它段落），否则位置判不准就不记。
    try:
        from docx.text.paragraph import Paragraph
        body = doc.element.body
        seq = list(body.iterchildren())
        positions: list = []
        for i, child in enumerate(seq):
            if child.tag != qn("w:p"):
                continue
            para = Paragraph(child, doc)
            if not is_caption_paragraph(para):
                continue
            # 找前后最近的非空 w:p 之外最近的 w:tbl（跳过紧邻的空段）
            prev_tbl = False
            next_tbl = False
            j = i - 1
            while j >= 0 and seq[j].tag == qn("w:p"):
                _pt = (Paragraph(seq[j], doc).text or "").strip()
                if _pt:
                    break
                j -= 1
            if j >= 0 and seq[j].tag == qn("w:tbl"):
                prev_tbl = True
            k = i + 1
            while k < len(seq) and seq[k].tag == qn("w:p"):
                _pt = (Paragraph(seq[k], doc).text or "").strip()
                if _pt:
                    break
                k += 1
            if k < len(seq) and seq[k].tag == qn("w:tbl"):
                next_tbl = True
            if prev_tbl and not next_tbl:
                positions.append("bottom")   # 表格在上、caption 在下 → 表题在表格下方
            elif next_tbl and not prev_tbl:
                positions.append("top")      # caption 在上、表格在下 → 表题在表格上方
        if positions:
            prof.caption_table_positions = positions
    except Exception:
        prof.caption_table_positions = []


# ---------------------------------------------------------------------------
# Strict OOXML (ISO/IEC 29500) -> Transitional (ECMA-376) namespace normalisation.
# Some official templates (e.g. the IEEE conference template) are saved as
# Strict OOXML, whose namespace URIs python-docx cannot parse (KeyError on
# officeDocument). We rewrite the strict URIs to the transitional ones before
# opening so those documents can be read. Normal (transitional) documents are
# detected and passed through untouched (zero risk of regression).
# ---------------------------------------------------------------------------
_STRICT_TO_TRANSITIONAL: dict = {
    # ---------------------------------------------------------------------
    # 🔴 v2.3.100（D-01 修复）：整张表按**官方 IEEE 会议模板（Strict OOXML）实测**
    # 的命名空间清单重写。旧表的键集体多写了 "/2006"（例如
    #   http://purl.oclc.org/ooxml/wordprocessingml/2006/main
    # 而 Strict 真实值是没有 2006 的 .../wordprocessingml/main），
    # 于是除关系类型外**一条都没匹配上** —— 归一化形同虚设，官方模板永远打不开。
    # 教训：这类映射必须以实产物的 xmlns 原文为准，不能凭印象写。
    # ---------------------------------------------------------------------
    # ① 关系类型（rel Type）前缀：必须先于裸 officeDocument 命中，
    #    否则会改写成 .../2006/relationships/**relationships**/xxx（重复）。
    b"http://purl.oclc.org/ooxml/officeDocument/relationships":
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    # ② 文档属性命名空间（Strict 用驼峰，Transitional 用连字符）
    b"http://purl.oclc.org/ooxml/officeDocument/docProps/core-properties":
        b"http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    b"http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes":
        b"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    b"http://purl.oclc.org/ooxml/officeDocument/extendedProperties":
        b"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    b"http://purl.oclc.org/ooxml/officeDocument/custom-properties":
        b"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    b"http://purl.oclc.org/ooxml/officeDocument/customXml":
        b"http://schemas.openxmlformats.org/officeDocument/2006/customXml",
    b"http://purl.oclc.org/ooxml/officeDocument/math":
        b"http://schemas.openxmlformats.org/officeDocument/2006/math",
    # 裸 officeDocument URI 只在上面的长前缀都没命中时才生效（长度排序保证）
    b"http://purl.oclc.org/ooxml/officeDocument":
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    # ③ WordprocessingML（Strict 无 /2006）
    b"http://purl.oclc.org/ooxml/wordprocessingml/main":
        b"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    # ④ DrawingML（Strict 无 /2006）
    b"http://purl.oclc.org/ooxml/drawingml/main":
        b"http://schemas.openxmlformats.org/drawingml/2006/main",
    b"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing":
        b"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    b"http://purl.oclc.org/ooxml/drawingml/picture":
        b"http://schemas.openxmlformats.org/drawingml/2006/picture",
    b"http://purl.oclc.org/ooxml/drawingml/chart":
        b"http://schemas.openxmlformats.org/drawingml/2006/chart",
    b"http://purl.oclc.org/ooxml/drawingml/chartDrawing":
        b"http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
    b"http://purl.oclc.org/ooxml/drawingml/diagram":
        b"http://schemas.openxmlformats.org/drawingml/2006/diagram",
    b"http://purl.oclc.org/ooxml/drawingml/lockedCanvas":
        b"http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas",
    # ⑤ package / Markup Compatibility（包层无 strict 变体，仍是自映射，保留兜底）
    b"http://purl.oclc.org/ooxml/package/2006/relationships":
        b"http://schemas.openxmlformats.org/package/2006/relationships",
    b"http://purl.oclc.org/ooxml/package/2006/content-types":
        b"http://schemas.openxmlformats.org/package/2006/content-types",
    b"http://purl.oclc.org/ooxml/markup-compatibility/2006":
        b"http://schemas.openxmlformats.org/markup-compatibility/2006",
}

# Replace longest keys first so a base URI is not partially rewritten
# (e.g. the longer ".../officeDocument/docProps/core-properties" must be handled
# before the bare ".../officeDocument").
_STRICT_NS_SORTED = sorted(
    _STRICT_TO_TRANSITIONAL.items(), key=lambda kv: len(kv[0]), reverse=True
)


def _rewrite_strict_ns(data: bytes) -> bytes:
    """Replace every Strict OOXML namespace URI with its Transitional equivalent."""
    for strict, trans in _STRICT_NS_SORTED:
        if strict in data:
            data = data.replace(strict, trans)
    return data


def _normalize_strict_ooxml(path: str):
    """If *path* contains Strict OOXML namespaces, return a ``BytesIO`` of the
    same archive with those namespaces rewritten to Transitional so python-docx
    can open it. Returns ``None`` when no Strict namespace is found (nothing to
    do) or when preprocessing fails (caller keeps the original open error).

    Detection scans the *decompressed* XML parts: a namespace URI is compressed
    inside the zip, so it is invisible in the raw file bytes and must not be
    tested there.
    """
    try:
        with open(path, "rb") as fh:
            raw = fh.read()
    except Exception:
        return None
    try:
        src = io.BytesIO(raw)
        with zipfile.ZipFile(src, "r") as zin:
            entries = [(it, zin.read(it.filename)) for it in zin.infolist()]
        found_strict = any(
            fn.lower().endswith((".xml", ".rels")) and b"purl.oclc.org/ooxml" in data
            for fn, data in ((it.filename, d) for it, d in entries)
        )
        if not found_strict:
            return None  # already transitional - leave untouched
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for it, data in entries:
                if it.filename.lower().endswith((".xml", ".rels")):
                    data = _rewrite_strict_ns(data)
                zout.writestr(it, data)
        out.seek(0)
        return out
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 模板批注（Word comments.xml）—— 海外「学校模板」模式的规范来源
# ---------------------------------------------------------------------------
# 国内版（文档格式先生）的模板驱动优先级是「批注 > 样式定义 > 通用规范」；海外版
# 此前只读实测格式（样式 + 直接格式），模板里写的 Word 批注**从未被读取**。
# 这里补上「读」的一侧：把批注原文取出来，交给 comment_spec_en 解析成英文规范，
# 再由 target_from_template_docx 以「批注 > 实测」的顺序合并。
# 🔴 只读，绝不修改模板文档本身（不建 comments 部件、不清任何批注）。


def read_template_comments(document) -> list:
    """读取 docx 里的 Word 批注（comments.xml），只读。

    返回 ``[{"id": int, "author": str, "text": str, "anchor": int|None}, ...]``：
    - ``text``   : 批注正文（多段以空格拼接；空白批注直接跳过）
    - ``anchor`` : 该批注锚定在正文第几段（0 基，对应 ``document.paragraphs``）；
                   找不到锚点（如批注锚在表格/页眉里）→ ``None``

    无批注 / 无 comments 部件 / 部件异常 → 返回 ``[]``，**绝不抛错**：
    读模板属于锦上添花，不能因为它让整个检查流程失败。
    """
    root = None
    try:
        part = document.part.part_related_by(RT.COMMENTS)
    except Exception:
        part = None
    if part is not None:
        try:
            root = getattr(part, "element", None)
            if root is None:                  # 裸 Part（非 XmlPart）才需要解析 blob
                root = parse_xml(part.blob)
        except Exception:
            root = None
    if root is None:
        # 兜底：关系缺失/异常时，从已加载的包部件里按 partname 找 comments.xml。
        # 有些模板生成工具注入批注时漏写 relationship（Word 仍能打开），
        # 只走 part_related_by 会把这类模板的批注整片漏掉。
        try:
            for p in document.part.package.iter_parts():
                if str(getattr(p, "partname", "")).lower() == "/word/comments.xml":
                    root = getattr(p, "element", None)
                    if root is None:
                        root = parse_xml(p.blob)
                    break
        except Exception:
            root = None
    if root is None:
        return []

    out: list = []
    for c in root.findall(qn("w:comment")):
        try:
            cid = int(c.get(qn("w:id")))
        except (TypeError, ValueError):
            continue
        # 批注正文可能有多段（w:p），逐段取 w:t 后拼接（保留 Word 里的换行语义）
        chunks = []
        for p in c.findall(qn("w:p")):
            seg = "".join(t.text or "" for t in p.iter(qn("w:t"))).strip()
            if seg:
                chunks.append(seg)
        text = " ".join(chunks).strip()
        if not text:
            continue
        out.append({
            "id": cid,
            "author": (c.get(qn("w:author")) or "").strip(),
            "text": text,
            "anchor": None,
        })
    if not out:
        return []

    # 锚点：正文段里以 w:commentRangeStart/@w:id 标记该批注的作用范围。
    # 只扫 body 段落（document.paragraphs 不含表格内段落）—— 模板规范批注通常挂在正文段。
    try:
        wanted = {c["id"] for c in out}
        by_id = {c["id"]: c for c in out}
        for idx, p in enumerate(document.paragraphs):
            for cs in p._p.iter(qn("w:commentRangeStart")):
                try:
                    pid = int(cs.get(qn("w:id")))
                except (TypeError, ValueError):
                    continue
                if pid in wanted and by_id[pid]["anchor"] is None:
                    by_id[pid]["anchor"] = idx
    except Exception:
        pass
    return out


def read_docx(path: str) -> DocumentProfile:
    """读取 docx → DocumentProfile。

    任何读取/解析异常（非 docx、损坏、加密、能打开但内部 XML 异常、受保护文档）
    一律转成 DocxReadError，保证 CLI/GUI 给用户友好提示而不是抛 traceback。
    """
    try:
        doc = Document(path)
    except Exception as first_err:
        # Could be a Strict OOXML document python-docx cannot parse, or a
        # genuinely broken / corrupt / encrypted file. Try the strict
        # normaliser; if it rewrites something, open that; otherwise keep the
        # original error.
        normalized = _normalize_strict_ooxml(path)
        if normalized is None:
            raise DocxReadError(
                f"Cannot read the Word document: {first_err} "
                f"(make sure it is a .docx file, not corrupted, and not encrypted)"
            ) from first_err
        try:
            doc = Document(normalized)
        except Exception as e:  # strict rewrite still failed to parse
            raise DocxReadError(
                f"Cannot read the Word document: {e} "
                f"(make sure it is a .docx file, not corrupted, and not encrypted)"
            ) from e
    try:
        return _read_docx_impl(doc)
    except DocxReadError:
        raise
    except Exception as e:  # opens but internal structure is abnormal / protected
        raise DocxReadError(
            f"Cannot parse the Word document content: {e} (the document may be protected or structurally corrupted)"
        ) from e


# ---------------------------------------------------------------------------
# 纸张大小（Letter / A4 / Legal / A5）
# ---------------------------------------------------------------------------
# 🔴 这是**通用维度，不属于任何引用规范**：北美（美/加）多用 Letter（8.5×11 in），
# 欧洲、中国、澳洲等用 A4（210×297 mm）。同一篇 APA 论文在伦敦交是 A4、在纽约交是
# Letter —— 所以真正说了算的是学校模板/客户，规范的只是「典型值参考」。
# 尺寸写成 **短边 × 长边**（英寸），比对时把文档实测宽高也归一成短边×长边，
# 这样横向（landscape）文档不会因为宽高对调被判成「不是 A4」。
PAPER_SIZES: dict = {
    "Letter": (8.5, 11.0),      # 215.9 × 279.4 mm —— 美国/加拿大/墨西哥等
    "A4": (8.27, 11.69),        # 210 × 297 mm —— 欧洲/中国/澳洲/日韩等
    "Legal": (8.5, 14.0),       # 215.9 × 355.6 mm
    "A5": (5.83, 8.27),         # 148 × 210 mm
}
PAPER_KEYS = tuple(PAPER_SIZES)
# 容差（英寸）：Word 里常见 8.27/8.28 这种舍入，0.06″ ≈ 1.5 mm，够容下各种导出差异
PAPER_TOL_IN = 0.06


def normalize_paper_dims(width_in: float, height_in: float) -> tuple:
    """把任意宽高归一成 (短边, 长边)，横向/纵向统一口径。"""
    w, h = float(width_in), float(height_in)
    return (min(w, h), max(w, h))


def identify_paper_size(width_in, height_in) -> Optional[str]:
    """由实测宽高识别纸型名；不在已知表内返回 None（绝不臆造）。"""
    if width_in is None or height_in is None:
        return None
    short, long_ = normalize_paper_dims(width_in, height_in)
    for name, (w, h) in PAPER_SIZES.items():
        if abs(short - w) <= PAPER_TOL_IN and abs(long_ - h) <= PAPER_TOL_IN:
            return name
    return None


def describe_paper_size(width_in, height_in) -> str:
    """给报告用的英文描述：认出就报名字，认不出就报实测尺寸。"""
    name = identify_paper_size(width_in, height_in)
    if name:
        return name
    if width_in is None or height_in is None:
        return "unknown"
    return "non-standard (%.2f x %.2f in)" % (float(width_in), float(height_in))


def detect_title_page(doc):
    """识别标题页结构，返回 {'title': int, 'authors': list[int]} 或 None。

    **单一真相源**：checker（判断客户要求是否满足）与 fixer（施加改动）共用本函数，
    否则会出现「检查说没问题、修正却动了另一段」的漂移。

    判定刻意保守 —— 必须有**标题页正面信号**才认，绝不对无信号文档盲猜：
    ① 首段不是已识别的**样式**标题（是样式标题就交给 heading 流程，不重复处理）；
    ② 前 40 段内出现 Abstract / Keywords 标签（标题页型论文的强特征）。
    满足后：首段 = 标题；紧随其后 ≤3 段内的短行（≤90 字、非标题、未到摘要）= 作者/单位。

    v2.3.71（客户视角实测）：首段判定从「is_heading（含手排伪标题）」收紧为
    「样式/outline 标题」+ 「手排大标题必须居中」。旧写法下，一篇**手排标题**的论文
    首段（16pt 居中加粗）会因 is_heading 为真而直接放弃标题页识别，连锁后果是：
    该大标题被当成伪标题占掉 L1 字号档 → 学生手排的章节标题（14pt）被推成 L2 →
    模板只声明 1 级 → 检查不报、修正跳过 → 客户点完「修正」章节标题一个字没变。
    手排大标题「居中」这道保险是为了避免把左对齐的章节标题误当论文标题摘掉。
    """
    paras = doc.paragraphs
    idx0 = None
    for i, p in enumerate(paras):
        if (p.text or "").strip():
            idx0 = i
            break
    if idx0 is None:
        return None
    _lvl0 = heading_level(paras[idx0])
    if _lvl0 is not None:
        # Word 内置「Title」样式被 _heading_level_from_name 故意算作 1 级；但它其实
        # 是**论文大标题**，不能因此放弃标题页识别 —— 否则标题页（标题 + 作者 + 单位）
        # 整块漏识别，标题还会占掉 L1 字号档（实测：heading_formats 只剩 {1: 24pt}）。
        _st_name0 = paras[idx0].style.name if paras[idx0].style is not None else None
        if not _is_doc_title_style_name(_st_name0):
            return None
    # v2.3.100（C 批 D-14）：官方 IEEE 会议模板的论文标题用 **`paper title` 样式**
    # （`<w:pStyle w:val="papertitle"/>`，居中写在样式里、段落本身没有 w:jc）。
    # 旧逻辑在这里走 is_heading 分支后又因"读不到段落级居中"而放弃 → 整份模板被判
    # 「没有标题页」，title_page 维度只能给 info。样式名是模板给的正面信号，直接认。
    try:
        _st0 = (paras[idx0].style.name or "").strip().lower()
    except Exception:
        _st0 = ""
    _is_paper_title_style = ("paper title" in _st0) or (_st0 in ("papertitle", "title"))
    if not _is_paper_title_style and is_heading(paras[idx0]):
        # 标题页型论文的大标题常是「左对齐 + 加粗 + 大字」的手排样式（尚未被居中）。
        # 这里仅在首段像**章节标题**（带编号前缀 "1." / "3.1"）时才放弃 ——
        # 章节标题永远不是首页首段；否则仍按论文大标题处理。原写法「不居中就放弃」
        # 会造成死锁：没居中→认不出→标题页版式不修正→永远不居中（v2.3.101）。
        # 论文大标题被认出后，标题页版式会先把它居中，彼此不冲突。
        if re.match(r'^\s*\d+(\.\d+)*\s*[\.\)]', (paras[idx0].text or "")):
            return None
        _al_raw = paras[idx0].alignment
        if _al_raw is None:
            _al_raw = _pf_attr(paras[idx0], None, "alignment")
        # 不再因「未居中」放弃：让标题页版式把它居中，打破死锁。
    blob = "\n".join((p.text or "") for p in paras[:min(len(paras), 60)])
    # v2.3.101：摘要常写在表格里（真实样本用 1×1 表格当摘要框）。只扫段落文字会
    # 漏掉 Abstract 信号 → 整篇被判「无标题页」→ 作者块没被排除 → 被卷进章节标题
    # 判定、误当成 run-in 标题还被注入句点（TEXT_IS_LOCKED 红线）。把表格文字也并入。
    try:
        for _t in doc.tables:
            for _row in _t.rows:
                for _cell in _row.cells:
                    blob += "\n" + (_cell.text or "")
    except Exception:
        pass
    # v2.3.100（C 批 D-14）：窗口 40 → **60 段**。IEEE 会议模板的 Abstract/Index Terms
    # 落在第 46 段附近（前面是标题、作者、摘要前的说明段），40 段窗口看不见 →
    # 整份模板被判「没有标题页」。60 段仍只是首页范围，不会把正文吞进来。
    if not re.search(r"(?i)\babstract\b|\bkeywords\b", blob):
        return None
    title_idx = idx0
    authors = []
    # v2.3.92：窗口从 3 段放宽到 7 段。原因是真实 APA 学生版题名页在「作者 + 单位」
    # 之后还有 **课程号与课名 / 教师名 / 提交日期** 三行，原来的 3 段窗口看不见它们，
    # 后果有两个：① 内容完整性提示把明明写了的元素说成「缺」；② 这几行不被居中、
    # 段距也不被清 —— 同一次修正里一半居中一半左对齐，客户一眼看出没做完。
    # 守卫一个没减：遇到摘要/关键词、标题、超长行（>90 字）立刻停，绝不吞正文。
    for i in range(idx0 + 1, min(len(paras), idx0 + 8)):
        p = paras[i]
        t = (p.text or "").strip()
        if not t:
            continue
        if re.search(r"(?i)\b(abstract|keywords)\b", t):
            break                       # 已经到摘要/关键词，停止
        if is_heading(p):
            break
        if len(t) <= 90:                # 作者/单位通常是短行
            authors.append(i)
        else:
            break
    return {"title": title_idx, "authors": authors}


def doc_title_index(doc) -> Optional[int]:
    """首页首段的「论文大标题」下标；判不出返回 None。

    为什么需要（v2.3.73，客户视角实测）：``detect_title_page`` 要求文档里有
    Abstract / Keywords 才认标题页。**没有摘要的论文非常常见** —— 这类文档里，
    16pt 居中加粗的大标题会落进「伪标题」判定、占掉 L1 字号档，学生手排的章节标题
    被推成 L2。检查结果于是张冠李戴：模板写着「L1 = Times New Roman 14pt 左对齐」，
    报告却说「Level 1 heading format: Exp Left 14pt / Act Center 16pt」，
    点修正就把**论文大标题**改成 14pt 左对齐（比漏检更伤）。

    判定刻意保守（全部满足才认，宁可判不出也不乱摘）：
      ① 第一个非空段，且不是样式 / outline 标题；
      ② 居中（段落直接格式 → 样式链回退，与 checker 的采样同口径）；
      ③ 整段加粗；
      ④ 字号是**全篇最大**，且**严格大于**它之后的每一段（章节标题不会比论文标题更大）；
      ⑤ 它之后至少还有 2 个非空段（单段 / 两段文档不猜）。

    本函数只负责「认出它是论文标题、别当章节标题」，**不**决定任何格式施加 ——
    要不要给标题页套格式仍由 title_page 维度（客户三态 > 规范）说了算。
    """
    paras = list(doc.paragraphs)
    idx0 = None
    for i, p in enumerate(paras):
        if (p.text or "").strip():
            idx0 = i
            break
    if idx0 is None:
        return None
    _lvl0 = heading_level(paras[idx0])
    if _lvl0 is not None:
        # 与 detect_title_page 同口径：Word 内置 Title 样式虽被算作 1 级标题，
        # 但它是论文大标题，仍需走「居中 + 加粗 + 全篇最大字号」的判定。
        _st_name0 = paras[idx0].style.name if paras[idx0].style is not None else None
        if not _is_doc_title_style_name(_st_name0):
            return None
    first = paras[idx0]
    runs = [r for r in first.runs if (r.text or "").strip()]
    if not runs:
        return None
    # 加粗 / 字号都必须**样式感知**：Word 内置 Title 样式把 bold / sz 写在样式层，
    # run 级读不到（r.bold is None），只认直接格式会把最标准的论文标题判成"不是标题"。
    if _run_bool_with_style(runs, first, "bold") is not True:
        return None
    _al_raw = first.alignment
    if _al_raw is None:
        _al_raw = _pf_attr(first, None, "alignment")
    if _ALIGN_MAP.get(_al_raw) != "center":
        return None
    title_sz = _max_run_size_pt(first) or _style_chain_size_pt(first)
    if title_sz is None:
        return None
    later = [p for p in paras[idx0 + 1:] if (p.text or "").strip()]
    if len(later) < 2:
        return None
    for p in later:
        sz = _max_run_size_pt(p) or _style_chain_size_pt(p)
        if sz is not None and sz >= title_sz:
            return None
    return idx0


def _sample_title_page(doc, prof: "DocumentProfile") -> None:
    """把标题页的实测格式写进画像（标题对齐/加粗、作者行对齐）。

    读不到的属性留 None（不臆造），由 checker 决定「无法判定」时只提示不判红。
    """
    try:
        tp = detect_title_page(doc)
    except Exception:
        tp = None
    if not tp:
        return
    prof.title_page_detected = True
    try:
        title_p = doc.paragraphs[tp["title"]]
    except Exception:
        return

    # 对齐：直接格式 → 段落样式 → Word 缺省（左）。
    # ⚠️ 「没有显式对齐」不等于「读不到」——Word 里未设 alignment 就是左对齐。
    # 旧写法在这里返回 None，checker 于是把「本该居中却没设」判成「无法判定」→ 不修，
    # 客户点完修正发现标题还是左对齐（实测打红 test_fix_apa_title_page_centers_and_bolds_title）。
    al_raw = title_p.alignment
    if al_raw is None:
        al_raw = _pf_attr(title_p, None, "alignment")
    prof.tp_title_alignment = _ALIGN_MAP.get(al_raw, "left") if al_raw is not None else "left"

    runs = [r for r in title_p.runs if (r.text or "").strip()]
    b_raw = [getattr(r.font, "bold", None) for r in runs]
    if b_raw and all(v is True for v in b_raw):
        prof.tp_title_bold = True
    elif b_raw and all(v is False for v in b_raw):
        prof.tp_title_bold = False
    else:
        # run 未显式声明 → 看段落样式；样式也没声明 → Word 缺省即「不加粗」。
        bold = False
        try:
            st = title_p.style
            if st is not None and st.font is not None and st.font.bold is not None:
                bold = st.font.bold is True
        except Exception:
            pass
        prof.tp_title_bold = bold

    for ai in tp["authors"]:
        try:
            ap = doc.paragraphs[ai]
        except Exception:
            continue
        a_al = ap.alignment
        if a_al is None:
            a_al = _pf_attr(ap, None, "alignment")
        prof.tp_author_alignment = _ALIGN_MAP.get(a_al, "left") if a_al is not None else "left"
        break                            # 取第一行作者/单位作代表（通常整组同格式）

    # 标题页标题段的段前 / 段后（v2.3.76）：认「行」单位，三级回退，读不到按 0 磅。
    sb, sa = _spacing_pair_of_para(title_p)
    prof.tp_title_space_before_pt = sb
    prof.tp_title_space_after_pt = sa

    # —— v2.3.91 首页结构 ——
    # ① 作者/单位行的段前段后（APA 要求整页双倍行距、不留额外段距 → 应为 0）。
    _mx: Optional[float] = None
    _last_tp_idx = tp["title"]
    for ai in tp["authors"]:
        try:
            ap = doc.paragraphs[ai]
        except Exception:
            continue
        _last_tp_idx = max(_last_tp_idx, ai)
        _b, _a = _spacing_pair_of_para(ap)
        for _v in (_b, _a):
            if _v is None:
                continue
            _mx = _v if _mx is None else max(_mx, _v)
    prof.tp_author_space_max_pt = _mx
    # ③ 标题与第一行作者之间夹了几个空段落（v2.3.92）。
    #    APA 的「一个双倍行距空行」既可能是真空段，也可能是标题段后间距撑出来的，
    #    判定必须两种都认，否则合规文档会被误报「缺空行」。
    try:
        prof.tp_author_gap_paragraphs = (
            max(0, min(tp["authors"]) - tp["title"] - 1) if tp["authors"] else None)
    except Exception:
        prof.tp_author_gap_paragraphs = None
    # ④ 标题以外那块文本（作者 / 单位 / 课程 / 教师 / 日期），供内容完整性提示用。
    #    刻意不含标题：论文大标题本身常带「字母+数字」（如 COVID-19），混进来会被
    #    误判成课程号，反而漏掉真正该提示的客户。
    try:
        _end = max([tp["title"], *tp["authors"]])
        _blk = []
        for _i in range(tp["title"] + 1, _end + 1):
            _t = (doc.paragraphs[_i].text or "").strip()
            if _t:
                _blk.append(_t)
        prof.tp_author_block_text = "\n".join(_blk) if _blk else None
    except Exception:
        prof.tp_author_block_text = None
    # ② 题名页块之后的段落是否另起一页（复用参考文献同一口径：段前分页 或 上一段末尾分页符）。
    try:
        _nxt = doc.paragraphs[_last_tp_idx + 1] if _last_tp_idx + 1 < len(doc.paragraphs) else None
        if _nxt is not None:
            prof.tp_next_break_before = _reference_starts_new_page(
                doc.paragraphs[_last_tp_idx], _nxt)
    except Exception:
        prof.tp_next_break_before = None


def _section_paper_size(sec) -> Optional[str]:
    """读单个分节的纸张大小：``sectPr/w:pgSz/@w:w`` 与 ``@w:h``（unit = twip）。"""
    try:
        w = sec.page_width
        h = sec.page_height
        if w is None or h is None:
            return None
        return identify_paper_size(_inch(w), _inch(h))
    except Exception:
        return None


def _section_columns(sec) -> int:
    """读单个分节的栏数：``sectPr/w:cols/@w:num``，缺省视为 1。

    python-docx 1.2.0 的 ``Section`` **没有**栏数属性（已实测 ``dir(Section)`` 无
    ``cols_*``），只能直读 ``sectPr`` 的 XML：``<w:cols w:num="2" w:space="425"/>``。
    ``w:num`` 缺失即 Word 默认单栏。任何读取异常都回落 1，绝不抛错阻断整体画像。
    """
    try:
        cols = sec._sectPr.find(qn("w:cols"))
        if cols is None:
            return 1
        raw = cols.get(qn("w:num"))
        if raw is None:
            return 1
        return max(1, int(raw))
    except Exception:
        return 1


def _body_sections(doc) -> list:
    """按 body 顺序产出 ``(sectPr 元素, 该节覆盖的块数)``。

    OOXML 没有「节起始」标记，只有「节结束」：段落 ``pPr/w:sectPr`` 表示**这一节
    在这段结束**；body 末尾的 ``sectPr`` 表示最后一节。所以只能顺序反推。

    🔴 为什么需要它（v2.3.100，C 批前置 X-1/X-2）：
    官方 IEEE 会议模板有 **5~7 个分节**（标题节 top=27pt 单栏、作者节 3~4 栏、
    正文节 top=54pt 两栏、尾节单栏…）。旧代码把**所有节的页边距/栏数取众数**，
    结果拿到的是标题节/作者节的值 —— A4 模板读出 top=0.312″、columns=1，
    而正文其实要求 0.75″/2 栏。判据错在「取哪一节」，不在「规范值」。
    """
    out: list = []
    try:
        body = doc.element.body
    except Exception:
        return out
    pending = 0
    for child in body.iterchildren():
        tag = child.tag
        if tag == qn("w:p"):
            pending += 1
            ppr = child.find(qn("w:pPr"))
            sp = ppr.find(qn("w:sectPr")) if ppr is not None else None
            if sp is not None:
                out.append((sp, pending))
                pending = 0
        elif tag == qn("w:tbl"):
            pending += 1
        elif tag == qn("w:sectPr"):
            out.append((child, pending))
            pending = 0
    return out


def _body_section(doc):
    """取**正文节**的 sectPr：覆盖块数最多者；并列取靠后的（正文在标题/作者节之后）。"""
    secs = _body_sections(doc)
    if not secs:
        return None
    best = secs[0]
    for sp, cnt in secs[1:]:
        if cnt >= best[1]:        # >= ：并列时向后取（后面的是正文）
            best = (sp, cnt)
    return best[0] if best[1] > 0 else None


def _twips_to_inch(raw) -> Optional[float]:
    """``w:pgMar`` / ``w:pgSz`` 的度量值 → 英寸。

    值可能是 twip 十进制数（``1440``），也可能带单位（Strict 的 IEEE 模板写的是
    ``54pt``）——两种都必须认，否则 Strict 文档的页边距会读成 0.0375″。
    """
    if raw is None:
        return None
    s = str(raw).strip()
    try:
        if s.lower().endswith("pt"):
            return round(float(s[:-2]) / 72.0, 4)
        if s.lower().endswith("in"):
            return round(float(s[:-2]), 4)
        return round(float(s) / 1440.0, 4)
    except (TypeError, ValueError):
        return None


def _sect_attr(sectPr, tag: str) -> dict:
    """安全读 sectPr 下的 pgMar/pgSz/cols 属性字典（键已去命名空间）。"""
    try:
        el = sectPr.find(qn(tag))
    except Exception:
        return {}
    if el is None:
        return {}
    return {k.split("}")[-1]: v for k, v in el.attrib.items()}


def _sect_columns(sectPr) -> int:
    """sectPr 的栏数：``w:cols/@w:num``，缺省 1（与 ``_section_columns`` 同口径）。"""
    raw = _sect_attr(sectPr, "w:cols").get("num")
    try:
        return max(1, int(raw))
    except (TypeError, ValueError):
        return 1


# 🔴 v2.3.100（C 批前置 X-3）：符号/公式字体不该参与「正文字体 + 字号」众数。
# 官方 IEEE 会议模板里 Symbol 的 run 数量压过正文 run，导致 font_family 读成
# "Symbol"、font_size 读成 9.0（正文其实是 Times New Roman 10pt）。
# 判定会因此全错，且这不是"规范值"问题、是"取样口径"问题。
_SYMBOL_FONT_NAMES = (
    "symbol", "wingdings", "webdings", "cambria math", "mt extra",
    "euclid math", "stix", "asana math", "latin modern math", "xits math",
)


def _is_symbol_font(name) -> bool:
    """是否为符号/数学字体（这类字体的字号不代表正文字号）。"""
    n = str(name or "").strip().lower()
    if not n:
        return False
    return any(n == s or n.startswith(s) for s in _SYMBOL_FONT_NAMES)


def _read_docx_impl(doc) -> DocumentProfile:
    prof = DocumentProfile()

    # v2.3.87：一切采样之前，先把软换行（Shift+Enter）拆成真实段落。
    # 位置必须**最早** —— 否则标题/正文/行距/参考文献的采样全都还在「粘块」视角上：
    # 粘在一起的章节标题不会被认成标题（检查器漏检），整块还会被当正文吃 0.5″ 缩进。
    try:
        prof.soft_break_splits = split_body_soft_breaks(doc)
    except Exception:
        prof.soft_break_splits = 0

    # 模板批注（Word comments）：必须在软换行拆分**之后**读 —— 拆分会产生新段落对象，
    # 先读会让 anchor 段号与后续 body_paras 的下标错位。
    try:
        prof.template_comments = read_template_comments(doc)
    except Exception:
        prof.template_comments = []

    # —— 页边距：遍历所有分节取众数；不一致则标记 ——
    per_side: dict = {s: [] for s in ("top", "bottom", "left", "right")}
    col_samples: list = []
    paper_samples: list = []
    try:
        sections = list(doc.sections)
    except Exception:
        sections = []
    for sec in sections:
        per_side["top"].append(_inch(sec.top_margin))
        per_side["bottom"].append(_inch(sec.bottom_margin))
        per_side["left"].append(_inch(sec.left_margin))
        per_side["right"].append(_inch(sec.right_margin))
        col_samples.append(_section_columns(sec))
        paper_samples.append(_section_paper_size(sec))
        try:
            if prof.paper_width_in is None and sec.page_width is not None:
                prof.paper_width_in = round(_inch(sec.page_width), 2)
                prof.paper_height_in = round(_inch(sec.page_height), 2)
        except Exception:
            pass

    def _mode(vals):
        vals = [v for v in vals if v is not None]
        if not vals:
            return None
        c = Counter(round(v, 3) for v in vals)
        return c.most_common(1)[0][0]

    prof.margins = {side: _mode(per_side[side]) for side in per_side}
    # 不一致判定：任一侧面出现 ≥2 个不同值
    prof.margins_inconsistent = any(
        len({round(v, 3) for v in vals if v is not None}) > 1
        for vals in per_side.values()
    )
    # 栏数：多分节取众数（与页边距同口径）；无分节则留 None（不臆造）
    if col_samples:
        prof.columns = Counter(col_samples).most_common(1)[0][0]
        # 纸张大小：多分节取众数（横向/纵向都按「短边×长边」归一后比对，
        # 否则一份横向 A4 会被判成「不是 A4」）
        paper_samples = [p for p in paper_samples if p is not None]
        if paper_samples:
            prof.paper_size = Counter(paper_samples).most_common(1)[0][0]

    # 🔴 v2.3.100（C 批前置 X-1/X-2）：**正文节优先**。
    # 上面按"全节众数"取到的往往是标题节/作者节的值（A4 官方模板 5 个分节：
    # 标题节 top=27pt 单栏、作者节 3 栏、正文节 top=54pt 双栏），拿它去判
    # 「页边距/栏数」必错。这里改用**覆盖块数最多的那一节**（= 正文节）的值覆盖，
    # 若拿不到正文节则保留众数兜底（旧行为，不倒退）。
    try:
        _sec = _body_section(doc)
    except Exception:
        _sec = None
    if _sec is not None:
        _mar = _sect_attr(_sec, "w:pgMar")
        _body_margins = {
            "top": _twips_to_inch(_mar.get("top")),
            "bottom": _twips_to_inch(_mar.get("bottom")),
            "left": _twips_to_inch(_mar.get("left")),
            "right": _twips_to_inch(_mar.get("right")),
        }
        for _side, _v in _body_margins.items():
            if _v is not None:
                prof.margins[_side] = _v
        prof.columns = _sect_columns(_sec)
        _sz = _sect_attr(_sec, "w:pgSz")
        _w_in, _h_in = _twips_to_inch(_sz.get("w")), _twips_to_inch(_sz.get("h"))
        if _w_in and _h_in:
            _ps = identify_paper_size(_w_in, _h_in)
            if _ps:
                prof.paper_size = _ps
        # 记录"这份文档是分节结构"，供报告说明"只按正文节判"（避免客户误解）
        prof.body_section_count = len(_body_sections(doc))


    # —— 文档默认（Normal）样式：字体 + 段落格式，作为三级回退的兜底 ——
    doc_default_font: Optional[str] = None
    doc_default_size: Optional[float] = None
    doc_pf = None
    try:
        normal = doc.styles["Normal"]
        if normal.font is not None:
            doc_default_font = normal.font.name
            # X-3 后续：字号也要三级回退（run → 样式 → 文档默认）。
            # 官方 IEEE 模板的正文 BodyText 样式**不写字号**（继承 OOXML 缺省 10pt），
            # 旧代码把这类 run 直接跳过，于是"有显式字号的 run"（摘要/作者信息 9pt）
            # 成了众数 → 正文字号被判成 9pt。这是取样口径问题，不是规范值问题。
            if normal.font.size is not None:
                doc_default_size = round(float(normal.font.size.pt), 1)
        doc_pf = normal.paragraph_format
    except Exception:
        pass


    # 参考文献区的软换行拆分已提到 `_read_docx_impl` 第一行（v2.3.87 起全文档统一拆），
    # 这里保留一次幂等调用兜底（无软换行时直接返回 0），防止将来有人单独调用本函数。
    split_reference_soft_breaks(doc)
    paras = list(doc.paragraphs)

    # —— 参考文献段（先定位，正文统计时要排除）——
    ref_start, ref_indices = find_reference_section(paras)
    # 🔴 v2.3.100（C 批 D-16）：若区内存在**专属参考文献样式**（IEEE 模板用
    # `<w:pStyle w:val="references"/>`，带 `<w:numPr/>` 自动编号），只认这些段落。
    # 官方模板在 "References" 标题与真条目之间还夹着 4 段**指导文字**（Body Text 样式），
    # 旧逻辑把它们也算成条目 → 15/16 被判"未编号"，连官方模板都被判红。
    # 只在该样式段落 ≥2 条时才生效，APA/MLA（条目多为 Normal/Body Text）完全不受影响。
    try:
        _ref_style_idx = [
            i for i in ref_indices
            if (paras[i].style.name or "").strip().lower()
            in ("references", "reference", "bibliography", "works cited", "workscited")
        ]
        if len(_ref_style_idx) >= 2:
            ref_indices = _ref_style_idx
    except Exception:
        pass
    prof.reference_entries = [paras[i].text.strip() for i in ref_indices]
    # v2.3.100（C 批 D-16）：条目是否**自带编号**。
    # ① 段落级 `<w:numPr/>`；② **样式级** numPr（IEEE 模板把编号写在 `references`
    # 样式里，段落本身不带 numPr —— 只看段落会 11/12 全判"未编号"）；
    # ③ 文本里已有 [n]（由 checker 的 _has_leading_number 判，不在这里重复）。
    try:
        def _has_num(para) -> bool:
            try:
                ppr = para._p.find(qn("w:pPr"))
                if ppr is not None and ppr.find(qn("w:numPr")) is not None:
                    return True
            except Exception:
                pass
            try:
                st = para.style
            except Exception:
                st = None
            while st is not None:
                try:
                    el = getattr(st, "element", None)
                    ppr = el.find(qn("w:pPr")) if el is not None else None
                    if ppr is not None and ppr.find(qn("w:numPr")) is not None:
                        return True
                except Exception:
                    pass
                st = getattr(st, "base_style", None)
            return False
        prof.reference_entries_numbered = [_has_num(paras[i]) for i in ref_indices]
    except Exception:
        prof.reference_entries_numbered = []
    prof.reference_entries_italic = [
        any(getattr(r.font, "italic", None) is True for r in paras[i].runs)
        for i in ref_indices
    ]
    prof.reference_text = "\n".join(prof.reference_entries)
    # 参考文献标题格式探针（APA 要求 "References" 粗体+居中于页顶；Purdue OWL general_format）
    if ref_start is not None:
        rh = paras[ref_start]
        prof.reference_heading_is_style = bool(is_heading(rh))
        prof.reference_heading_bold = any(
            getattr(r.font, "bold", None) is True for r in rh.runs)
        prof.reference_heading_centered = (rh.alignment == WD_ALIGN_PARAGRAPH.CENTER)

    # 正文段落 = 非标题（含伪标题）、非参考文献段、非空
    # v2.3.77：目录标题段与目录条目也排除 —— 它们是 Word 生成的目录内容（条目里还有域），
    # 若混进正文采样，短文档里「正文段前段后/行距/对齐」的众数会被目录口径带偏；
    # 而且 fixer 本来就不改条目（域守卫），检查按目录口径判、修正不碰，就是「检查/修正不同源」。
    skip = set(ref_indices)
    if ref_start is not None:
        skip.add(ref_start)
    # 🔴 v2.3.90：标题页段落（标题 / 作者 / 单位）**不进正文采样**。
    # 它们的排版由 ``spec.title_page`` 独占管辖（见 fixer._apply_title_page_format 的
    # 同一句注释），混进正文样本就是「一个属性两个 owner」——而且是自伤：
    # 修正器按标题页要求把「作者/单位」行居中后，检查器又把它们当正文样本，
    # 于是**修完复检反报「正文对齐 = 居中」**。短文档里更狠：
    # 正文段 2 条 + 居中作者行 2 条 → Counter 2:2 平票，most_common 直接翻成 center。
    try:
        _tp_probe = detect_title_page(doc)
    except Exception:
        _tp_probe = None
    if _tp_probe:
        if _tp_probe.get("title") is not None:
            skip.add(_tp_probe["title"])
        skip.update(_tp_probe.get("authors") or ())
    # 🔴 v2.3.100（C 批前置 X-3 续）：再按**样式名**排除前置部分。
    # 官方 IEEE 会议模板里 Author 占位段有 41 段、正文 Body Text 只有 17 段，且作者段
    # 全是 9pt + 居中 → 正文字号被判 9pt、对齐被判 center、首行缩进被判 0，**三项全错**。
    # detect_title_page 只能抓到「标题 + 紧随其后的 ≤3 段」，兜不住整块作者信息，
    # 所以这里按模板自带的样式名再兜一层（Author / paper title / Abstract / Keywords…）。
    _FRONT_STYLE_KEYS = ("author", "affiliation", "paper title", "papertitle",
                         "paper subtitle", "papersubtitle", "abstract", "keyword",
                         "index term", "sponsor", "cover page", "title")

    def _is_front_style(p) -> bool:
        try:
            nm = (p.style.name or "").strip().lower()
        except Exception:
            return False
        return any(nm == k or nm.startswith(k) for k in _FRONT_STYLE_KEYS)

    # 🔴 v2.3.100（幂等收尾）：把「fixer 会当成标题的段」也从正文字号/行距众数里排除。
    # 单一真相源：与 fixer._assign_heading_levels 共用 structural_heading_map。
    # 旧口径只排 is_heading()（**带样式**的标题），漏了「12pt 加粗、样式 Normal」的伪标题：
    # 官方 IEEE 样例里 7 个伪标题 12pt > 5 个正文 10pt → 众数被判 12pt，
    # 而修正侧只改正文段（标题有自己的字号口径）→ 每遍都判「字号不符」、
    # 每遍都记 applied → 客户点完「修正」那条红杠永远不消失。
    # ⚠️ 必须**先筛候选**（heading_level 为 None 的段）再喂给 structural_heading_map ——
    # 与 fixer 的调用方式保持一致。直接传全量 paras 会认出 **0 个**标题（实测：全量时
    # pseudo_heading_level_map 的字号档推断被正文段稀释，整张表落空），排除等于没做。
    _excl_p = {paras[i]._p for i in skip if 0 <= i < len(paras)}
    _pseudo_cand = [p for p in paras if _heading_level(p) is None]
    _struct_ps = set(structural_heading_map(_pseudo_cand, exclude_ps=_excl_p).keys())

    _body_core = [p for i, p in enumerate(paras)
                  if i not in skip and p.text.strip() and not is_heading(p)
                  and toc_entry_level(p) is None and not is_toc_heading(p)
                  and not _is_front_style(p)
                  # v2.3.100：**手排伪标题**（"1. Introduction"、"Chapter 2: …"，
                  # 样式是 Normal、但显然是标题）也要排除出行距众数 ——
                  # 否则「手排标题 1.0 倍 + 正文 1.14 倍」打平/被拉低，判定说"行距不符"，
                  # 而修正侧（fixer 的 no_heading）本就不改标题 → **每跑一遍都判不符**，
                  # 客户点完「修正」那条红杠永远不消失（幂等缺陷的真凶）。
                  and not is_section_heading_line(p)
                  # v2.3.100：fixer 认得出的标题段（含上面 is_section_heading_line 兜不住的
                  # 「无编号加粗标题」）一并排除，保证「检查所依据的正文」与
                  # 「修正所施加的正文」是同一批段落。
                  and p._p not in _struct_ps
                  # v2.3.100：**形状像标题的段**也排除。上面两条都依赖「加粗/编号」，
                  # 而 IEEE 一级标题按 C-5 是「不加粗 + 居中」—— 第一遍修好后，
                  # is_pseudo_heading 与 is_section_heading_line 同时失效，12pt 标题
                  # 重新混进众数。改用修正器改不掉的「短 + 不以句读结尾」。
                  and not is_standalone_heading_shape(p)]
    # 兜底：**只要还剩 1 段就用**。原阈值是 >=3，本意是防「极端文档全用前置样式 →
    # 采样搞空」，但它把「排除标题后只剩 2 段真正文」的文档也打回了含标题的旧口径 ——
    # 官方 IEEE 样例正文只有 5 段、伪标题 7 段，排除后剩 5 段（没事）；而合成样例
    # 排除后只剩 2 段 → 回退 → 7 个 12pt 标题重新压过 2 个 10pt 正文 →
    # 众数判 12pt，而修正侧只改正文 → 每遍都判「字号不符」→ 红杠永远不消失。
    # 兜底要防的是「空」，不是「少」：剩 1 段时那 1 段就是正文，取它比取标题更接近真相。
    body_paras = _body_core if _body_core else [
        p for i, p in enumerate(paras)
        if i not in skip and p.text.strip() and not is_heading(p)
        and toc_entry_level(p) is None and not is_toc_heading(p)]

    # —— 标题段探测（MLA/APA：标题居中；MLA 不加粗/不斜体/不下划线）——
    # 首页首个居中段落即视为标题（MLA 标题居中；参考文献标题已排除在 skip 外，另有专门检查）。
    title_idx = None
    for i, p in enumerate(paras):
        if i in skip or not p.text.strip():
            continue
        if p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            title_idx = i
            break
    if title_idx is not None:
        tp = paras[title_idx]
        prof.title_text = tp.text.strip()
        prof.title_alignment = "center"
        prof.title_bold = any(getattr(r.font, "bold", None) is True for r in tp.runs)
        prof.title_italic = any(getattr(r.font, "italic", None) is True for r in tp.runs)
        prof.title_underline = any(getattr(r.font, "underline", None) is True for r in tp.runs)
        # 首页信息块：标题之前、左上对齐、非标题、非空行数（MLA 应为 姓名/教授/课程/日期 共 4 行）
        cnt = 0
        for j in range(title_idx):
            pp = paras[j]
            if j in skip or not pp.text.strip() or is_heading(pp):
                continue
            al = _pf_attr(pp, doc_pf, "alignment")
            aln = _ALIGN_MAP.get(al) if al is not None else "left"
            if aln == "left":
                cnt += 1
        prof.first_page_left_lines = cnt

    # —— 标题页实测（APA/IEEE 等「有独立标题页」的论文）——
    # 与上面的「首页居中标题」是两件事：那个服务于 MLA 无标题页形态，这里服务于
    # 有 Abstract/Keywords 的标题页形态。检测函数与 fixer 共用（单一真相源）。
    _sample_title_page(doc, prof)

    # —— 论文大标题（无摘要的论文也要认得出来）——
    # 它不属于章节标题层级：混进去会占掉 L1 字号档，让报告把「论文标题」说成
    # 「一级标题格式不符」，修正时把大标题改小（详见 doc_title_index 说明）。
    try:
        _dti = doc_title_index(doc)
    except Exception:
        _dti = None
    if _dti is not None:
        prof.doc_title_index = _dti
        prof.doc_title_size_pt = _max_run_size_pt(doc.paragraphs[_dti])
    else:
        # v2.3.114 加固（题目字号读取盲区）：doc_title_index 因「题目不是全篇最大字号」
        # 等保守条件判不出时（典型：题目被错改小到与正文同号，不再最大），题目字号会
        # 读不到 → title_page 维度误报 pass。此处退化用 detect_title_page 定位题名页首段
        # 读字号，让检查器仍能标红「题目字号不符」。
        # 守卫：仅认「居中」或「Title/papertitle 样式」的题名页首段，避免把左对齐的章节
        # 标题误读成论文题目字号（负作用）。绝不改 prof.doc_title_index —— 章节标题层级
        # 排除逻辑（checker / docx_reader 扫描 / fixer 均依赖它）保持原样，零影响。
        try:
            _tp = detect_title_page(doc)
        except Exception:
            _tp = None
        if _tp and isinstance(_tp.get("title"), int):
            _ti = _tp["title"]
            _tp0 = doc.paragraphs[_ti]
            _al = _tp0.alignment
            if _al is None:
                _al = _pf_attr(_tp0, None, "alignment")
            _st = (_tp0.style.name or "").strip().lower() if _tp0.style is not None else ""
            _title_like = (_ALIGN_MAP.get(_al) == "center"
                           or "paper title" in _st or _st in ("papertitle", "title")
                           or _is_doc_title_style_name(_tp0.style.name if _tp0.style is not None else None))
            if _title_like:
                _sz = _max_run_size_pt(_tp0) or _style_chain_size_pt(_tp0)
                if _sz is not None:
                    prof.doc_title_size_pt = _sz

    # —— 块引用候选收集（MLA：超 4 行引文应缩进 0.5″、不加引号；此处只收集长段落，由 checker 按风格判断）——
    # idx 一并记下：fixer 要按**同一批候选**施加缩进，否则「检查按长引文判、修正只认
    # Quote 样式」，客户看到「3 处块引用不达标」，点完修正一处没动（检查/修正不同源）。
    # 键用 p._p（w:p 元素），绝不用 id(para)：python-docx 每次迭代都新建包装对象（铁律）
    _body_idx = {p._p: i for i, p in enumerate(paras)}
    for p in body_paras:
        words = len(p.text.split())
        if words < 50:
            continue
        left_indent = _inch(_pf_attr(p, doc_pf, "left_indent"))
        has_quote = _is_quoted_paragraph(p.text)
        prof.block_quote_candidates.append({
            "idx": _body_idx.get(p._p),
            "words": words,
            "left_indent_in": left_indent,
            "has_quote": has_quote,
        })

    # —— 段落级排版：对齐 / 首行缩进 / 段后间距（正文段落众数）——
    # 三级回退都取不到时，按 Word 规范默认值计入（对齐缺省=左对齐、缩进缺省=0、
    # 段后缺省=0），否则会对绝大多数正常文档误报「未能读取」。
    align_counter: Counter = Counter()
    indent_counter: Counter = Counter()
    space_after_counter: Counter = Counter()
    space_before_counter: Counter = Counter()
    # v2.3.98：表格紧邻段落不参与「正文段前段后」众数（它们的间距归 table_gap 管），
    # 否则给表格留的那 24pt 会把众数抬高，检查器反过来说「正文段前应 0」→ 自伤。
    # 但退化文档（正文段落本来就少、且全挨着表格）会把众数算空 → 回退统计全量。
    try:
        prof.table_gap_samples, _tbl_adj_ps = table_gap_samples(doc)
    except Exception:
        prof.table_gap_samples, _tbl_adj_ps = [], set()
    _spacing_all = []                       # [(sb, sa)] 全部正文段
    _spacing_plain = []                     # 不含表格紧邻段
    for para in body_paras:
        al_raw = _pf_attr(para, doc_pf, "alignment")
        al = _ALIGN_MAP.get(al_raw) if al_raw is not None else "left"
        if al:
            align_counter[al] += 1
        fli = _inch(_pf_attr(para, doc_pf, "first_line_indent"))
        indent_counter[0.0 if fli is None else fli] += 1
        # 段前 / 段后：直读 XML（认 w:beforeLines / w:afterLines「行」单位），
        # 不再只认 w:after。旧写法读 python-docx 的 space_after，模板写「段后 0.5 行」
        # 时读到 None → 计入 0.0 → 检查报「段后 0 磅，合规」，客户看到的空行
        # 比要求大一倍也不管。
        sb, sa = _spacing_pair_of_para(para)
        _spacing_all.append((sb, sa))
        if para._p not in _tbl_adj_ps:
            _spacing_plain.append((sb, sa))

    for _sb, _sa in (_spacing_plain or _spacing_all):
        space_after_counter[0.0 if _sa is None else _sa] += 1
        space_before_counter[0.0 if _sb is None else _sb] += 1

    if align_counter:
        prof.body_alignment = align_counter.most_common(1)[0][0]
    if indent_counter:
        prof.first_line_indent_in = indent_counter.most_common(1)[0][0]
    if space_after_counter:
        prof.space_after_pt = space_after_counter.most_common(1)[0][0]
    if space_before_counter:
        prof.space_before_pt = space_before_counter.most_common(1)[0][0]

    # —— 字体 / 字号 / 行距 / 标题（正文段落 + 表格段落）——
    font_counter: Counter = Counter()
    size_counter: Counter = Counter()
    # X-3：非符号字体的字号（正文字号的真实众数来源）
    plain_size_counter: Counter = Counter()
    # v2.3.100：表格段的字体/字号**单独计数**。旧写法把表格 run 并进正文 counter，
    # 于是一篇「正文 12pt、表格 10pt」的论文里，40 个表格 run 会压过 34 个正文 run →
    # 正文字号被判 10pt，而修正只改正文 → 复检仍读 10 → **永远点不掉的红杠**
    # （v2.3.97 只在 family_size_counter 上挡过这一刀，font/size 两个 counter 仍被污染）。
    # 现在表格只在「正文完全没有样本」时兜底。
    table_font_counter: Counter = Counter()
    table_size_counter: Counter = Counter()
    family_size_counter: dict = {}          # 字体 → 该字体出现的字号计数（判混用）
    spacing_samples: list = []
    heading_styles: list = []
    heading_levels: set = set()

    def _sample_paragraph(para, pair_by_family: bool = True, into_body: bool = True):
        """把一段文字的字体/字号纳入采样（run → 段落样式 → 文档默认）。

        ``pair_by_family``：是否同时记「字体 → 字号」配对。表格段落传 False ——
        配对表是 `font_size_by_font` 维度的唯一依据，而这个维度只改正文段
        （`_apply_size_by_font_family(no_heading, ...)` 从不碰表格）。
        表格若混进来，一篇表格文字 10pt 的论文里 TNR 的众数会被 10 顶掉 →
        检查说「应为 12pt」、修正只改正文、复检仍读 10 → **一条永远点不掉的红杠**
        （v2.3.97 实测：真实论文 40 个表格 run 压过 34 个正文 run）。

        ``into_body``：是否计入**正文字体/字号众数**。表格段落传 False ——
        表格有自己的字号口径（常见比正文小 2pt），压过正文会让「正文字号」判错
        （v2.3.100 实测：MLA 真实论文 80 个表格 10pt run 压过 65 个正文 12pt run，
        第二遍恒记 `font_size_pt` 改动）。行距/段距同理早已分开（见下方注释）。
        """
        for run in para.runs:
            if not run.text.strip():
                continue
            fn = _run_font_name(run)
            if not fn and para.style is not None and para.style.font is not None:
                fn = para.style.font.name
            if not fn:
                fn = doc_default_font
            fs = run.font.size
            if fs is None and para.style is not None and para.style.font is not None:
                fs = para.style.font.size
            # X-3 后续：第三级回退到文档默认；再取不到按 OOXML 缺省 w:sz=20（10pt）。
            # 少这一级，正文"靠样式继承字号"的文档会把只有显式字号的 run（题注 9pt、
            # 作者信息 9pt）当成正文众数 —— IEEE 模板实测就因此读成 9pt。
            _fs_pt: Optional[float] = None
            if fs is not None:
                try:
                    _fs_pt = round(float(fs.pt), 1)
                except Exception:
                    _fs_pt = None
            if _fs_pt is None:
                _fs_pt = doc_default_size if doc_default_size is not None else 10.0
            if fn:
                if into_body:
                    font_counter[fn] += 1
                else:
                    table_font_counter[fn] += 1
            if _fs_pt is not None:
                if into_body:
                    size_counter[_fs_pt] += 1
                    if not _is_symbol_font(fn):
                        plain_size_counter[_fs_pt] += 1
                else:
                    table_size_counter[_fs_pt] += 1
            # 字体 → 字号配对计数（混用字体时每种字体各取自己的众数字号）
            if pair_by_family and fn and _fs_pt is not None:
                family_size_counter.setdefault(fn, Counter())[_fs_pt] += 1

    # 标题统计（样式名 / outlineLvl）
    for para in paras:
        lvl = _heading_level(para)
        if lvl is not None and para.text.strip():
            heading_styles.append(para.style.name if para.style else f"level{lvl}")
            heading_levels.add(lvl)

    # 字体/字号/行距**只统计正文段落**：标题段与参考文献段的口径与正文不同，
    # 计入会把「正文字体/字号」的众数带偏（短文档里可能整段判错）。
    # 表格段落另行采样（论文三线表常见，漏掉同样会让判断偏）。
    for para in body_paras:
        sp = _line_spacing_of_para(para)
        if sp is not None:
            spacing_samples.append(sp)
        _sample_paragraph(para)

    # 表格内文字一并采样（论文三线表常见，漏掉会让「正文最常见字体/字号」判断偏）
    table_count = 0
    tbl_sb: Counter = Counter()
    tbl_sa: Counter = Counter()
    for para in _iter_table_paragraphs(doc):
        if not para.text.strip():
            continue
        table_count += 1
        _sample_paragraph(para, pair_by_family=False, into_body=False)
        # v2.3.100：**表格段不参与正文行距众数**。表格有自己的行距口径（常见单倍），
        # 混进来会把正文众数拉低 → 判定说"行距不符 1.14"，而修正侧只改正文段
        # （表格不动）→ 每遍都判不符、每遍都记 applied → **永远修不掉的红杠**。
        # 字体/字号仍然纳入（防表格字号压过正文，那两条改的是正文，必须看得见表格）。
        # v2.3.77：表格单元格段落的段前 / 段后单独立计数器 —— 绝不能并进正文众数
        # （表格里通常是 0，正文却常有 6–12 磅；混在一起两边都会判错）。
        _tb, _ta = _spacing_pair_of_para(para)
        tbl_sb[0.0 if _tb is None else _tb] += 1
        tbl_sa[0.0 if _ta is None else _ta] += 1
    prof.table_paragraph_count = table_count
    if tbl_sb:
        prof.table_space_before_pt = tbl_sb.most_common(1)[0][0]
    if tbl_sa:
        prof.table_space_after_pt = tbl_sa.most_common(1)[0][0]
    # v2.3.85：表格边框现状（三线表检查用）。读「格→表→样式」三级有效边框，
    # 见 sample_table_lines()。空列表 = 没有表格，checker 会说明「没找到表格」。
    try:
        prof.table_line_states = sample_table_lines(doc)
    except Exception:
        prof.table_line_states = []

    # v2.3.112：公式（Word 原生 OMML）现状。python-docx 的 para.runs 只返回 w:r，
    # 数学 run 是 m:r —— 不单独采一遍，公式段在引擎眼里就是「一个字都没有的空段」，
    # 任何维度都判不到它。空列表 = 没有公式，checker 会如实说明。
    try:
        prof.equation_samples = sample_equations(doc)
    except Exception:
        prof.equation_samples = []

    # v2.3.100（F-5 / 裁定 C-6）：byline（作者/单位/城市/邮箱）采样。
    try:
        prof.byline_size_pt, prof.byline_line_count = sample_byline(doc)
    except Exception:
        prof.byline_size_pt, prof.byline_line_count = None, 0
    # v2.3.100（D-1 / 裁定 C-11）：官方模板指导文字残留（只报位置，不自动删）。
    try:
        prof.boilerplate_hits = sample_template_boilerplate(doc)
    except Exception:
        prof.boilerplate_hits = []
    # v2.3.100（F-9 / 裁定 C-12 C-13 C-14）：三类只提示观测。
    try:
        _tsn = sample_text_style_notes(doc)
        prof.citation_range_en_dash = _tsn["citation_range_en_dash"]
        prof.suffix_name_comma = _tsn["suffix_name_comma"]
        prof.equation_number_styles = _tsn["equation_number_styles"]
    except Exception:
        prof.citation_range_en_dash = []
        prof.suffix_name_comma = []
        prof.equation_number_styles = []
    # v2.3.100（D-2~D-11）：IEEE 系列只读观测。
    try:
        _ie = sample_ieee_extras(doc)
        prof.ieee_caption_labels = _ie["caption_labels"]
        prof.ieee_caption_position = {
            "figure_ref_order_ok": _ie.get("figure_ref_order_ok"),
            "table_ref_order_ok": _ie.get("table_ref_order_ok"),
        }
        prof.ieee_figure_ref_order = _ie["figure_ref_order"]
        prof.ieee_first_footnote = _ie["first_footnote"]
        prof.ieee_acknowledgment = _ie["acknowledgment"]
        prof.ieee_coi = _ie["coi"]
        prof.ieee_biography = _ie["biography"]
        prof.ieee_index_terms = _ie["index_terms"]
        prof.ieee_abstract_numbered_refs = _ie["abstract_numbered_refs"]
        prof.ieee_numbered_singular_heads = _ie["numbered_singular_heads"]
    except Exception:
        prof.ieee_caption_labels = []
        prof.ieee_caption_position = {}
        prof.ieee_figure_ref_order = []
        prof.ieee_first_footnote = {}
        prof.ieee_acknowledgment = {}
        prof.ieee_coi = {}
        prof.ieee_biography = {}
        prof.ieee_index_terms = {}
        prof.ieee_abstract_numbered_refs = []
        prof.ieee_numbered_singular_heads = []
    # v2.3.100（E 批）：MLA + APA 剩余只读观测。
    try:
        _me = sample_mla_apa_extras(doc)
        prof.body_title_repeat = _me["body_title_repeat"]
        prof.numbered_body_heads = _me["numbered_body_heads"]
        prof.author_note_index = _me["author_note_index"]
        prof.author_note_has_superscript = _me["author_note_has_superscript"]
        prof.first_page_left_texts = _me["first_page_left_texts"]
    except Exception:
        prof.body_title_repeat = False
        prof.numbered_body_heads = []
        prof.author_note_index = None
        prof.author_note_has_superscript = None
        prof.first_page_left_texts = []
    # v2.3.100（E-6 / D-22）：MLA 论文标题段实测。
    try:
        _mt = sample_mla_title(doc)
        prof.mla_title_index = _mt["index"]
        prof.mla_title_alignment = _mt["alignment"]
        prof.mla_title_bold = _mt["bold"]
        prof.mla_title_italic = _mt["italic"]
        prof.mla_title_certain = bool(_mt["certain"])
    except Exception:
        prof.mla_title_index = None
        prof.mla_title_alignment = None
        prof.mla_title_bold = None
        prof.mla_title_italic = None
        prof.mla_title_certain = False
    # v2.3.100（E-2 / D-18）：APA caption 字形。
    try:
        _ct = sample_caption_typography(doc)
        prof.caption_label_bold = _ct["caption_label_bold"]
        prof.caption_title_italic = _ct["caption_title_italic"]
        prof.caption_typography_count = _ct["count"]
    except Exception:
        prof.caption_label_bold = None
        prof.caption_title_italic = None
        prof.caption_typography_count = 0

    # v2.3.100：**正文优先，表格只在正文无样本时兜底**。
    # 旧写法把表格 run 并进同一个 counter，一篇「正文 12pt / 表格 10pt」的论文里
    # 表格 run 数常常更多 → 正文字号判成 10pt，而修正只改正文 → 复检仍读 10
    # → 客户点完「修正」那条红杠永远不消失（实测 MLA 第二遍恒记 font_size_pt）。
    _font_pool = font_counter or table_font_counter
    if _font_pool:
        # X-3：符号/数学字体（Symbol、Cambria Math…）不代表正文字体，排除后再取众数；
        # 若全篇只有符号字体（极端情况）则回退全量，不制造 None。
        _plain_fonts = Counter({k: v for k, v in _font_pool.items()
                                if not _is_symbol_font(k)})
        prof.font_family = (_plain_fonts or _font_pool).most_common(1)[0][0]
    if family_size_counter:
        prof.font_size_by_family = {
            str(f): float(c.most_common(1)[0][0])
            for f, c in family_size_counter.items() if c and not _is_symbol_font(f)
        }
    # X-3：优先用「非符号字体」的字号众数；全篇皆为符号字体时回退全量。
    # v2.3.100：正文样本永远优先于表格样本（见 _font_pool 处注释）。
    _size_pool = plain_size_counter or size_counter or table_size_counter
    if _size_pool:
        prof.font_size_pt = _size_pool.most_common(1)[0][0]
    if spacing_samples:
        prof.line_spacing = Counter(round(s, 2) for s in spacing_samples).most_common(1)[0][0]

    prof.heading_styles = list(dict.fromkeys(heading_styles))
    prof.heading_levels_used = max(heading_levels) if heading_levels else 0

    # —— 每级标题格式采样（学校模板驱动：取每级首个标题段为代表性样本）——
    # 每级只取第一个样本，避免同层多个标题互相平均掩盖差异。
    level_samples: dict = {}
    seen_levels: set = set()
    for _i, para in enumerate(paras):
        # 🔴 前置块不是章节标题，绝不能占层级字号档（实测真 bug）：
        #   · skip            —— detect_title_page 认出的标题页标题 + 作者行；
        #   · _is_front_style —— 样式名属前置类（含 Word 内置 "Title"）；
        #   · doc_title_index —— 无摘要论文里认出的论文大标题。
        # 不排除的后果：Word "Title" 样式被 _heading_level_from_name 判为 1 级，
        # 先占掉 L1，seen_levels 再把真正的 Heading 1 跳过 → heading_formats 只剩
        # {1: 标题字号}，学校模板目标 L1 变成标题字号，修正会把章节标题放大成标题号。
        # 与下方「伪标题兜底」路径（4789-4793 已排除 skip / doc_title_index）同一口径。
        if _i in skip or _i == prof.doc_title_index or _is_front_style(para):
            continue
        lvl = _heading_level(para)
        if lvl is None or not para.text.strip():
            continue
        if lvl in seen_levels:
            continue
        seen_levels.add(lvl)
        level_samples[lvl] = _sample_heading_format(para, doc_pf, doc_default_font)
    if level_samples:
        prof.heading_formats = {lvl: level_samples[lvl] for lvl in sorted(level_samples)}

    # —— 伪标题兜底采样（v2.3.65）——
    # 学生论文最常见的形态是「章节标题就是整段加粗的 Normal 段，压根没有 Heading 样式」。
    # 旧版只采真标题样式 → 客户填了「一级 16pt 居中加粗」，检查器只回一句
    # 「No Heading styles detected」、修正预览一行不出，而修正器其实会照改（fixer 走
    # _assign_heading_levels 认伪标题）——检查/预览/修正三方脱节。
    # 这里补一份伪标题样本，分层规则与 fixer 共用 pseudo_heading_level_map（单一真相源）。
    # 只在「一个真标题样式都没有」时兜底：有样式就按样式走，绝不把两套口径混在一起。
    if not level_samples:
        # 排除参考文献标题与条目：模板里的 "References" 常是「12pt 居中加粗」，
        # 会因字号档差异被当成「二级标题」凭空造出一级要求（实测：学校模板画像里
        # 多出一条 L2 = 12pt 居中加粗，其实是参考文献标题）。后置部分标题不属于
        # 章节标题层级，必须排除。
        _skip_ps = {paras[i]._p for i in skip}
        # v2.3.73：论文大标题（首页居中加粗、全篇最大字号）不是章节标题，
        # 若混进采样，画像里的 L1 就变成「16pt 居中加粗」，把真实章节标题挤成 L2，
        # 检查和修正会一起对着错误的分层跑（详见 doc_title_index 的说明）。
        _dt = doc_title_index(doc)
        if _dt is not None:
            _skip_ps.add(paras[_dt]._p)
        # 🔴 v2.3.100：**必须先筛候选再调**，与 fixer._assign_heading_levels 完全一致。
        # 旧写法直接传全量 paras → structural_heading_map 内部 pseudo_heading_level_map
        # 的字号档推断被正文段稀释，**一张表全落空**（实测认出 0 个）；而 fixer 传的是
        # 「heading_level 为 None 的候选子集」，认得出 7 个。于是同一份文档：
        # 检查器按 reader 的分层（产物实测 L1+L3）判「三级标题该斜体」，
        # 修正器按自己的分层（L1+L2+L4）压根不碰 L3 → **报了却修不掉的红杠**
        # （C-4 停掉 font_size_by_font 自动修后，这条在 APA 真实论文上显形）。
        _pseudo_cand2 = [p for p in paras if _heading_level(p) is None]
        _pseudo = structural_heading_map(_pseudo_cand2, exclude_ps=_skip_ps)
        _first_of_level: dict = {}
        for para in paras:
            lvl = _pseudo.get(para._p)
            if lvl is None or not para.text.strip():
                continue
            _first_of_level.setdefault(lvl, para)
        if _first_of_level:
            prof.heading_formats = {
                lvl: _sample_heading_format(p, doc_pf, doc_default_font)
                for lvl, p in sorted(_first_of_level.items())
            }
            prof.heading_formats_are_pseudo = True

    # —— 参考文献悬挂缩进实测（条目段落里 first_line_indent 为负 → 悬挂）——
    # 三级回退取不到显式缩进时按 0 计入（Word 缺省无悬挂），与正文缩进口径一致。
    if ref_indices:
        hangs: list = []
        for i in ref_indices:
            fli_emu = _pf_attr(paras[i], doc_pf, "first_line_indent")
            fli = _inch(fli_emu)
            hangs.append(round(-fli, 3) if (fli is not None and fli < 0) else 0.0)
        if hangs:
            prof.reference_hanging_indent_in = Counter(hangs).most_common(1)[0][0]

        # —— 参考文献条目的字体 / 字号 / 行距（v2.3.65）——
        # 学校模板里「参考文献用五号宋体、单倍行距」是硬要求，且常与正文不同口径
        # （正文小四双倍、条目五号单倍）。以前只采正文，条目一律按正文处理 → 模板
        # 里这条要求既读不到也改不了。取众数（条目之间可能有混排）。
        ref_fonts: Counter = Counter()
        ref_sizes: Counter = Counter()
        ref_spacings: list = []
        ref_sb: list = []
        ref_sa: list = []
        _ref_style_sizes: list = []
        for i in ref_indices:
            if not _is_real_ref_entry(paras, i):
                continue
            _st = paras[i].style
            while _st is not None:
                if _st.font is not None and _st.font.size is not None:
                    _ref_style_sizes.append(round(float(_st.font.size.pt), 1))
                    break
                _st = getattr(_st, "base_style", None)
        for i in ref_indices:
            # v2.3.100（F-7 / 裁定 C-10）：只统计**真条目**。
            # 官方模板 "References" 标题下先夹 4 段 `Body Text` 样式的**指导文字**
            # （"The template will number citations…"），末尾又夹 1 段 —— 旧写法把它们
            # 一起算进众数，字號被 Body Text 的 10pt 顶掉（实测 reference_size=10.0，
            # 而模板 references 样式明明是 8pt）→ 判"应为 8pt"→ 把官方模板判红。
            # 真条目判据：样式名含 reference/bibliography，或该段带 Word 自动编号。
            if not _is_real_ref_entry(paras, i):
                continue
            para_i = paras[i]
            for r in para_i.runs:
                if not (r.text or "").strip():
                    continue
                # v2.3.100：段内还夹着 **Word 的样式提示 run** —— 文本就是 "(references)"、
                # "(body text)" 这种（Word 用它在页边显示当前样式名），自带 SimSun 10pt。
                # 不排除的话 3 个提示 run 就能把整条参考文献的字号众数顶成 10pt
                # （实测：模板 references 样式明明 8pt，却被判 10pt）。
                # 判据：纯括号包住的**字母串**（≤30 字）—— 条目里的 "(2020)"、"(IEEE)"
                # 之类含数字或是专有名词的仍照常统计。
                if _STYLE_HINT_RE.match(r.text.strip()):
                    continue
                fn = _run_font_name(r)
                if fn:
                    ref_fonts[fn] += 1
                sz = _pt(r.font.size)
                if sz:
                    ref_sizes[sz] += 1
            sp_i = _line_spacing_of_para(para_i)
            if sp_i is not None:
                ref_spacings.append(sp_i)
            # 条目之间的段前/段后（v2.3.75）：认「行」单位，读不到按 0 磅计入
            _b, _a = _spacing_pair_of_para(para_i)
            ref_sb.append(0.0 if _b is None else _b)
            ref_sa.append(0.0 if _a is None else _a)
        if ref_fonts:
            prof.reference_font_family = ref_fonts.most_common(1)[0][0]
        if ref_sizes:
            prof.reference_size_pt = ref_sizes.most_common(1)[0][0]
        elif _ref_style_sizes:
            # v2.3.100（F-7）：run 没写字号时回退**样式链字号** ——
            # 官方模板的 references 条目正是如此（样式 8pt、run 无 sz）。
            # 没有这一级回退就会读成 None → 该维度恒"无结论"，等于没判。
            prof.reference_size_pt = Counter(_ref_style_sizes).most_common(1)[0][0]
        if ref_spacings:
            prof.reference_line_spacing = Counter(
                round(s, 2) for s in ref_spacings).most_common(1)[0][0]
        if ref_sb:
            prof.reference_space_before_pt = Counter(ref_sb).most_common(1)[0][0]
        if ref_sa:
            prof.reference_space_after_pt = Counter(ref_sa).most_common(1)[0][0]

    # —— 页眉 / 页码 / 参考文献另起一页（轻量读取，仅用于保守提示，不修改文档）——
    try:
        rh_text, page_present, page_right = _read_header_footer_info(doc)
        prof.header_text = rh_text
        prof.page_number_present = page_present
        prof.page_number_right_aligned = page_right
        # 页码细项：在页眉还是页脚、有效对齐、是否手打（已含首页/偶数页口径）
        _pn = read_page_number_info(doc)
        prof.page_number_in = _pn.location
        prof.page_number_alignment = _pn.alignment
        prof.page_number_is_static = _pn.has_static and not _pn.has_field
        prof.page_number_static_text = _pn.static_text
        # 页眉距页顶（MLA 9th 要求 ½″）：取首个有显式值的分节；读不到则为 None
        try:
            _hd_val = None
            for _sec in doc.sections:
                _hd = _inch(getattr(_sec, "header_distance", None))
                if _hd is not None:
                    _hd_val = _hd
                    break
            prof.header_distance_in = _hd_val
        except Exception:
            prof.header_distance_in = None
    except Exception:
        # 读取页眉失败不应中断主流程；其余维度照常
        prof.header_text = ""
        prof.page_number_present = False
        prof.page_number_right_aligned = None
        prof.page_number_in = None
        prof.page_number_alignment = None
        prof.page_number_is_static = False
        prof.page_number_static_text = ""
    if ref_start is not None:
        prev = paras[ref_start - 1] if ref_start > 0 else None
        try:
            prof.reference_new_page = _reference_starts_new_page(prev, paras[ref_start])
        except Exception:
            prof.reference_new_page = None
    else:
        prof.reference_new_page = None

    # —— 前置材料检测（Abstract / Author Note）：仅用于专业版提示，不修改内容 ——
    try:
        prof.has_abstract, prof.has_author_note = _scan_front_matter(paras)
    except Exception:
        prof.has_abstract = False
        prof.has_author_note = False

    # —— 摘要页特化采样（v2.3.67）：摘要正文段的首行缩进 + 是否另起一页 ——
    # 摘要正文（APA 不首行缩进）与正文（0.5″ 首行缩进）口径相反，必须单独采样，
    # 否则「正文首行缩进」会把摘要段一起刷成 0.5″（越权）。
    try:
        _abs_head, _abs_body = find_abstract_section(paras)
        if _abs_head is not None and _abs_body:
            # v2.3.96：样式口径的采样**排除 Keywords 行**（摘要正文顶格、Keywords 行
            # 缩进 0.5″，混在一起取众数会互相污染）；分页锚点仍用完整 body。
            _abs_style = abstract_body_style_indices(paras, _abs_body)
            # 首行缩进：取摘要正文段的众数（与正文缩进口径一致，三级回退）
            _abs_indents: Counter = Counter()
            for _j in _abs_style:
                _fli = _inch(_pf_attr(paras[_j], doc_pf, "first_line_indent"))
                _abs_indents[0.0 if _fli is None else _fli] += 1
            if _abs_indents:
                prof.abstract_first_line_indent_in = _abs_indents.most_common(1)[0][0]
            # v2.3.76：摘要正文段的段前 / 段后（磅），取众数；读不到按 0 磅计入
            _abs_sb: Counter = Counter()
            _abs_sa: Counter = Counter()
            for _j in _abs_style:
                _b, _a = _spacing_pair_of_para(paras[_j])
                _abs_sb[0.0 if _b is None else _b] += 1
                _abs_sa[0.0 if _a is None else _a] += 1
            if _abs_sb:
                prof.abstract_space_before_pt = _abs_sb.most_common(1)[0][0]
            if _abs_sa:
                prof.abstract_space_after_pt = _abs_sa.most_common(1)[0][0]
            # v2.3.100（C 批 D-08）：摘要正文的字号 / 加粗 / 对齐（IEEE 会议模板
            # Abstract 样式 `<w:sz w:val="18"/>`(9pt) + `<w:b/>` + `jc=both`）。
            # 采样口径与正文一致：run → 样式 → OOXML 缺省 10pt；对齐走三级回退。
            _abs_sizes: Counter = Counter()
            _abs_bold: Counter = Counter()
            _abs_align: Counter = Counter()
            for _j in _abs_style:
                _pj = paras[_j]
                _sz_pt = None
                for _r in _pj.runs:
                    if not (_r.text or "").strip():
                        continue
                    if _r.font.size is not None:
                        _sz_pt = round(float(_r.font.size.pt), 1)
                        break
                if _sz_pt is None:
                    try:
                        _stf = _pj.style.font
                        if _stf is not None and _stf.size is not None:
                            _sz_pt = round(float(_stf.size.pt), 1)
                    except Exception:
                        _sz_pt = None
                if _sz_pt is None:
                    _sz_pt = doc_default_size if doc_default_size is not None else 10.0
                _abs_sizes[_sz_pt] += 1
                for _r in _pj.runs:
                    if (_r.text or "").strip():
                        _bl = _r.font.bold
                        if _bl is None:
                            # v2.3.100：run 未写 bold 时回退**样式链**。
                            # IEEE 模板把 `<w:b/>` 写在 Abstract 样式里，run 上一片空白，
                            # 只看 run 会判成 "not bold"，把合规摘要报成 warn（R2 同类病根）。
                            try:
                                _stb = _pj.style
                                while _stb is not None:
                                    if _stb.font is not None and _stb.font.bold is not None:
                                        _bl = _stb.font.bold
                                        break
                                    _stb = getattr(_stb, "base_style", None)
                            except Exception:
                                _bl = None
                        _abs_bold[bool(_bl)] += 1
                        break
                _al = _pf_attr(_pj, doc_pf, "alignment")
                if _al is not None:
                    _abs_align[_ALIGN_MAP.get(_al) or str(_al)] += 1
            if _abs_sizes:
                prof.abstract_font_size_pt = _abs_sizes.most_common(1)[0][0]
            if _abs_bold:
                prof.abstract_bold = _abs_bold.most_common(1)[0][0]
            if _abs_align:
                prof.abstract_alignment = _abs_align.most_common(1)[0][0]
            # 是否另起一页：段前分页 或 上一段末尾分页符（复用参考文献同一口径）
            _prev_abs = paras[_abs_head - 1] if _abs_head > 0 else None
            prof.abstract_new_page = _reference_starts_new_page(
                _prev_abs, paras[_abs_head])
            # v2.3.97：摘要**标题**自身的格式（对齐 / 加粗 / 斜体）。
            # 采集时必须三级回退取值（段落 → 样式链 → 文档默认），否则「样式里写着居中、
            # 段上没有直接格式」的文档会读成左对齐，检查器随即报一条修不掉的错。
            _ah = paras[_abs_head]
            _ah_al = _pf_attr(_ah, doc_pf, "alignment")
            prof.abstract_heading_alignment = (
                _ALIGN_MAP.get(_ah_al) if _ah_al is not None else None)
            _ah_runs = [r for r in _ah.runs if (r.text or "").strip()]
            if _ah_runs:
                # v2.3.100（B-5）：走「run 直接 → 样式链」三级回退。
                # 官方样例的 Abstract 标题用 Heading 1 样式承载，加粗写在样式里，
                # 只读 run 会全部得到 None → 报「could not read bold」这种修不掉的错。
                prof.abstract_heading_bold = _run_bool_with_style(
                    _ah_runs, _ah, "bold")
                prof.abstract_heading_italic = _run_bool_with_style(
                    _ah_runs, _ah, "italic")
            # v2.3.97：关键词行（"Keywords: ..."）的标签斜体 + 整行首行缩进（APA §2.10）。
            _kw_idx = next((j for j in _abs_body if is_keywords_line(paras[j])), None)
            if _kw_idx is not None:
                _kwp = paras[_kw_idx]
                _kw_runs = [r for r in _kwp.runs if (r.text or "").strip()]
                if _kw_runs:
                    # v2.3.100（B-5）：与摘要标题同一口径（run → 样式链）。
                    # 官方样例的 "Keywords:" 标签斜体常写在字符/段落样式上，
                    # 只读 run 会得到 None → 判定"label italic not set" → 误报。
                    prof.keywords_label_italic = _run_bool_with_style(
                        _kw_runs, _kwp, "italic")
                _kw_fli = _inch(_pf_attr(_kwp, doc_pf, "first_line_indent"))
                prof.keywords_indent_in = 0.0 if _kw_fli is None else _kw_fli
            # v2.3.91：摘要之后**正文首段**是否另起一页（APA：正文从摘要后新一页开始）。
            # 摘要正文可能跨多段，取摘要区最后一段为前一段。
            _abs_last = max(_abs_body)
            _after = paras[_abs_last + 1] if _abs_last + 1 < len(paras) else None
            if _after is not None:
                prof.text_break_before = _reference_starts_new_page(
                    paras[_abs_last], _after)
    except Exception:
        prof.abstract_first_line_indent_in = None
        prof.abstract_new_page = None
        prof.abstract_space_before_pt = None
        prof.abstract_space_after_pt = None
        prof.text_break_before = None
        prof.abstract_heading_alignment = None
        prof.abstract_heading_bold = None
        prof.abstract_heading_italic = None
        prof.keywords_label_italic = None
        prof.keywords_indent_in = None

    # —— 图表 caption 采样（v2.3.67）——
    try:
        _sample_captions(doc, prof)
    except Exception:
        prof.caption_count = 0
        prof.caption_font_size_pt = None
        prof.caption_line_spacing = None
        prof.caption_space_before_pt = None
        prof.caption_space_after_pt = None
        prof.caption_table_positions = []

    # —— 目录采样（v2.3.77）：只读不判；没有目录时全部字段保持 None ——
    try:
        _sample_toc(doc, prof)
    except Exception:
        prof.toc_detected = False
        prof.toc_entry_count = 0
        prof.toc_heading_text = ""
        prof.toc_heading_space_before_pt = None
        prof.toc_heading_space_after_pt = None
        prof.toc_entry_space_before_pt = None
        prof.toc_entry_space_after_pt = None

    # —— 脚注 / 尾注计数（v2.3.77）：只数不改，供「本工具未检查脚注」提示用 ——
    try:
        _sample_notes(doc, prof)
    except Exception:
        prof.footnote_count = 0
        prof.endnote_count = 0

    return prof


def reference_paragraph_indices(doc) -> list:
    """返回参考文献段落在 doc.paragraphs 中的下标列表。

    复用与 read_docx 完全一致的 find_reference_section（唯一实现），
    供 fixer 与测试精确区分「正文段落」与「参考文献段落」。

    v2.3.84：调用前先把参考文献区的软换行拆成真实段落 —— 否则「整份文献表被
    Shift+Enter 粘成一个段落」的文档（真实学生论文极常见）只会被认成 1 条，
    悬挂缩进无从落地（真实样本实测：应有 12 条，引擎只认 1 条）。
    """
    split_reference_soft_breaks(doc)
    return find_reference_section(list(doc.paragraphs))[1]


def _run_pieces(run_el):
    """把 <w:r> 按软换行切段，返回 [[子元素...], ...]（已剔除 rPr；长度>1 表示有换行）。"""
    pieces = [[]]
    for child in run_el:
        if child.tag == qn("w:rPr"):
            continue
        if _is_soft_break(child):
            pieces.append([])
        else:
            pieces[-1].append(deepcopy(child))
    return pieces


def _new_run(like_run_el, children):
    """按参考 run 的 rPr 造一个新 run，装载给定子元素。"""
    run = OxmlElement("w:r")
    rpr = like_run_el.find(qn("w:rPr"))
    if rpr is not None:
        run.append(deepcopy(rpr))
    for el in children:
        run.append(el)
    return run


def split_paragraph_on_soft_breaks(para) -> int:
    """把一段里的软换行（``<w:br/>`` / ``<w:cr>``）拆成真正的段落。返回拆出的段数。

    为什么必须拆：真实学生论文大量用 Shift+Enter 分隔「章节标题 / 每条参考文献」，
    于是标题与正文、或十几条文献全挤在一个 ``<w:p>`` 里。引擎按段落工作，
    粘在一起就既不能逐条悬挂缩进、也不能分别套标题格式。

    铁律：**一个字符都不动**（老板 2026-10-21 清单第 5 条）。这里只做结构性操作：
    复制 ``pPr``（段落格式原样）与 ``rPr``（字符格式原样），把原来的子元素按换行
    位置重新分配到多个 ``<w:p>`` 里；超链接、域、书签等容器元素整体搬到所属段落，
    内容与顺序完全保持。空行（连续换行）也如实保留。
    """
    p = para._p
    body = p.getparent()
    if body is None:
        return 0
    # 只有真的含软换行才动手
    if not any(_is_soft_break(ch) or
               (ch.tag == qn("w:r") and any(_is_soft_break(x) for x in ch))
               for ch in p):
        return 0
    ppr = p.find(qn("w:pPr"))
    lines = [[]]
    for child in list(p):
        if child.tag == qn("w:pPr"):
            continue
        if _is_soft_break(child):
            lines.append([])
            continue
        if child.tag == qn("w:r"):
            pieces = _run_pieces(child)
            if len(pieces) == 1:
                lines[-1].append(deepcopy(child))
            else:
                for i, children in enumerate(pieces):
                    if i > 0:
                        lines.append([])
                    if children:
                        lines[-1].append(_new_run(child, children))
            continue
        # 超链接 / 域 / 书签等容器：整块放进当前行（内部软换行不拆，保守）
        lines[-1].append(deepcopy(child))

    if len(lines) < 2:
        return 0

    # 第一行复用原段落，其余新建段落插在它后面
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)
    for el in lines[0]:
        p.append(el)
    prev = p
    for line in lines[1:]:
        new_p = OxmlElement("w:p")
        if ppr is not None:
            new_p.append(deepcopy(ppr))
        for el in line:
            new_p.append(el)
        prev.addnext(new_p)
        prev = new_p
    return len(lines) - 1


def split_reference_soft_breaks(doc) -> int:
    """只拆参考文献区里的软换行（保守：正文段落一律不动）。返回拆出的段数。"""
    try:
        ref_start, ref_indices = find_reference_section(list(doc.paragraphs))
    except Exception:
        return 0
    if not ref_indices:
        return 0
    paras = list(doc.paragraphs)
    targets = [paras[i] for i in ref_indices if i < len(paras)]
    n = 0
    for para in targets:
        try:
            n += split_paragraph_on_soft_breaks(para)
        except Exception:
            continue
    return n


def split_body_soft_breaks(doc) -> int:
    """把**全文档**正文级段落里的软换行拆成真实段落。返回拆出的段落数。

    v2.3.87。v2.3.84 只拆了参考文献区（当时保守），正文里「章节标题 + 正文」被
    Shift+Enter 粘成一整块的问题被漏了下来。真实样本实测：87 个段落里 30 个含软换行，
    其中 30 个第一行就是章节标题。后果有两层：

      ① 粘块整体被当成一个正文段落 → 标题行跟着吃到 0.5″ 首行缩进
         （客户一眼看出「章节标题缩进了、正文没缩进」）；
      ② 检查器压根看不见这些标题（它们不是独立段落），标题维度整块漏检。

    所以拆必须做在**读数最前面**（``_read_docx_impl`` 第一行），让检查器和修正器
    看到同一份段落结构 —— 否则两边 paragraph 下标会错位。

    安全性：只拆 ``<w:br/>``（无 type / ``textWrapping``）与 ``<w:cr>``；
    ``<w:br w:type="page|column">`` 是分页/分栏符，属于排版语义，绝不当段边界
    （拆了会把客户的分页符吃掉）。文字一个字符都不动，只增加段落标记。
    """
    try:
        paras = list(doc.paragraphs)
    except Exception:
        return 0
    n = 0
    for para in paras:
        try:
            n += split_paragraph_on_soft_breaks(para)
        except Exception:
            continue
    return n
