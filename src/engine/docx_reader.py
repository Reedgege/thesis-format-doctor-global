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
    # —— 表格边框（三线表，v2.3.85）——
    # 每张表一条 dict，见 sample_table_lines()：has_top / has_bottom / has_header_line /
    # has_vertical / has_inside_h / three_line。空列表 = 文档里没有表格（不是"合规"）。
    table_line_states: list = field(default_factory=list)
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
    # v2.3.76：摘要正文段的段前 / 段后（磅）；无摘要段 = None
    abstract_space_before_pt: Optional[float] = None
    abstract_space_after_pt: Optional[float] = None
    abstract_new_page: Optional[bool] = None          # 摘要是否另起一页（段前分页/上一段末尾分页符）；无摘要段=None
    # —— 图表 caption 采样（v2.3.67）——
    caption_count: int = 0                            # 识别到的 Caption 样式段落数
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
    for r in runs:
        try:
            if r.bold is not True:
                return False
        except Exception:
            return False
    threshold = _PSEUDO_HEADING_MIN_PT if min_pt is None else min_pt
    sz = _max_run_size_pt(para)
    return sz is not None and sz >= threshold


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
    return bool(_SECTION_HEADING_RE.match(text))


def heading_level(para) -> Optional[int]:
    """公开的标题层级访问器（fixer 套用每级格式规则用）。

    返回 1-based 层级（1=最高级）或 None（非标题 / 伪标题无明确层级）。
    判定链路：样式名 → 段落级 outlineLvl → 样式级 outlineLvl，兼容自定义标题样式。
    """
    return _heading_level(para)


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
                    v = round(float(raw) / 20.0, 1)
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
            has_header_line = bool(first) and all(
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


def read_page_number_info(doc) -> PageNumberInfo:
    """读页码：有无（域/手打）+ 在页眉还是页脚 + 有效对齐。

    页眉与页脚同时有页码时以**页眉**为准 —— page_number_position 的 top-right 语义
    指的是页眉右上（APA 7th / MLA 9th 都是）。
    """
    info = PageNumberInfo()
    for part, kind, _scope in _iter_header_footer_parts(doc):
        try:
            para_list = part.paragraphs
        except Exception:
            continue
        for p in para_list:
            if _para_has_page_field(p):
                if not info.has_field or (kind == "header" and info.location != "header"):
                    info.has_field = True
                    info.location = kind
                    info.alignment = _effective_para_alignment(p)
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
            para_list = part.paragraphs
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


def _reference_starts_new_page(prev_para, cur_para) -> bool:
    """参考文献标题是否另起一页：段前分页 或 上一段末尾分页符。"""
    if cur_para is not None and _para_page_break_before(cur_para):
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
        if not t or len(t) > 40:
            continue
        if "abstract" in t or "摘要" in t:
            has_abs = True
        if "author note" in t or "作者注" in t:
            has_auth = True
    return has_abs, has_auth


def find_abstract_section(paras) -> tuple[Optional[int], list[int]]:
    """定位摘要段：返回 (摘要标题段下标, 摘要正文段下标列表)。

    **单一真相源**：checker（判摘要格式）与 fixer（施加摘要格式）共用本函数，
    避免「检查判这一段、修正改另一段」的漂移（与 find_reference_section 同理）。

    识别规则（保守，仅认正面信号）：
    - 摘要标题 = 样式级标题（heading_level 非 None）且文本（去空格、小写）等于
      "abstract" / "摘要" / "abstract page"（APA 常见 "Abstract" 一级标题）。
    - 摘要正文 = 标题之后、下一个样式级标题之前的所有非空段。
    - 找不到样式级标题的 "Abstract" 时返回 (None, [])，绝不把加粗正文误判成摘要标题。
    """
    paras = list(paras)
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
    return None, []


# 图表 caption 识别（v2.3.67）：Word 内置 "Caption" 样式 + 文本前缀启发式双保险。
# 很多学生用「Table 1」「Figure 2」「表 1」「图 2」作为标题段，但没套 Caption 样式；
# 只认样式会漏掉一大半。前缀式兜底（保守：前缀后必须跟数字或标题文字）。
_CAPTION_STYLE_KEYWORDS = ("caption", "图注", "表注", "图表标题", "captions")
_CAPTION_PREFIX_RE = re.compile(
    r"^(?:table|figure|fig\.|tab\.|表|图)\s*\d+", re.IGNORECASE)


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
    return bool(_CAPTION_PREFIX_RE.match(text))


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
    b"http://purl.oclc.org/ooxml/officeDocument/docProps/core-properties":
        b"http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    b"http://purl.oclc.org/ooxml/officeDocument/extended-properties":
        b"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    b"http://purl.oclc.org/ooxml/officeDocument/custom-properties":
        b"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    b"http://purl.oclc.org/ooxml/officeDocument/math":
        b"http://schemas.openxmlformats.org/officeDocument/2006/math",
    # Bare officeDocument URI must be mapped AFTER the longer /docProps/... keys
    # above (handled by the length-sorted rewrite below).
    b"http://purl.oclc.org/ooxml/officeDocument":
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    b"http://purl.oclc.org/ooxml/wordprocessingml/2006/main":
        b"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    b"http://purl.oclc.org/ooxml/drawingml/2006/main":
        b"http://schemas.openxmlformats.org/drawingml/2006/main",
    b"http://purl.oclc.org/ooxml/drawingml/2006/wordprocessingDrawing":
        b"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    b"http://purl.oclc.org/ooxml/drawingml/2006/picture":
        b"http://schemas.openxmlformats.org/drawingml/2006/picture",
    b"http://purl.oclc.org/ooxml/drawingml/2006/chart":
        b"http://schemas.openxmlformats.org/drawingml/2006/chart",
    b"http://purl.oclc.org/ooxml/drawingml/2006/chartDrawing":
        b"http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
    b"http://purl.oclc.org/ooxml/drawingml/2006/diagram":
        b"http://schemas.openxmlformats.org/drawingml/2006/diagram",
    b"http://purl.oclc.org/ooxml/drawingml/2006/lockedCanvas":
        b"http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas",
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
    if heading_level(paras[idx0]) is not None:
        return None
    if is_heading(paras[idx0]):
        # 无样式但形似标题（大字 + 加粗）：只有居中才按「论文大标题」处理
        _al_raw = paras[idx0].alignment
        if _al_raw is None:
            _al_raw = _pf_attr(paras[idx0], None, "alignment")
        if _ALIGN_MAP.get(_al_raw) != "center":
            return None
    blob = "\n".join((p.text or "") for p in paras[:min(len(paras), 40)])
    if not re.search(r"(?i)\babstract\b|\bkeywords\b", blob):
        return None
    title_idx = idx0
    authors = []
    for i in range(idx0 + 1, min(len(paras), idx0 + 4)):
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
    if idx0 is None or heading_level(paras[idx0]) is not None:
        return None
    first = paras[idx0]
    runs = [r for r in first.runs if (r.text or "").strip()]
    if not runs:
        return None
    for r in runs:
        try:
            if r.bold is not True:
                return None
        except Exception:
            return None
    _al_raw = first.alignment
    if _al_raw is None:
        _al_raw = _pf_attr(first, None, "alignment")
    if _ALIGN_MAP.get(_al_raw) != "center":
        return None
    title_sz = _max_run_size_pt(first)
    if title_sz is None:
        return None
    later = [p for p in paras[idx0 + 1:] if (p.text or "").strip()]
    if len(later) < 2:
        return None
    for p in later:
        sz = _max_run_size_pt(p)
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


def _read_docx_impl(doc) -> DocumentProfile:
    prof = DocumentProfile()

    # v2.3.87：一切采样之前，先把软换行（Shift+Enter）拆成真实段落。
    # 位置必须**最早** —— 否则标题/正文/行距/参考文献的采样全都还在「粘块」视角上：
    # 粘在一起的章节标题不会被认成标题（检查器漏检），整块还会被当正文吃 0.5″ 缩进。
    try:
        prof.soft_break_splits = split_body_soft_breaks(doc)
    except Exception:
        prof.soft_break_splits = 0

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

    # —— 文档默认（Normal）样式：字体 + 段落格式，作为三级回退的兜底 ——
    doc_default_font: Optional[str] = None
    doc_pf = None
    try:
        normal = doc.styles["Normal"]
        if normal.font is not None:
            doc_default_font = normal.font.name
        doc_pf = normal.paragraph_format
    except Exception:
        pass

    # 参考文献区的软换行拆分已提到 `_read_docx_impl` 第一行（v2.3.87 起全文档统一拆），
    # 这里保留一次幂等调用兜底（无软换行时直接返回 0），防止将来有人单独调用本函数。
    split_reference_soft_breaks(doc)
    paras = list(doc.paragraphs)

    # —— 参考文献段（先定位，正文统计时要排除）——
    ref_start, ref_indices = find_reference_section(paras)
    prof.reference_entries = [paras[i].text.strip() for i in ref_indices]
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
    body_paras = [p for i, p in enumerate(paras)
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
        space_after_counter[0.0 if sa is None else sa] += 1
        space_before_counter[0.0 if sb is None else sb] += 1

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
    family_size_counter: dict = {}          # 字体 → 该字体出现的字号计数（判混用）
    spacing_samples: list = []
    heading_styles: list = []
    heading_levels: set = set()

    def _sample_paragraph(para):
        """把一段文字的字体/字号纳入采样（run → 段落样式 → 文档默认）。"""
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
            if fn:
                font_counter[fn] += 1
            if fs is not None:
                try:
                    size_counter[round(float(fs.pt), 1)] += 1
                except Exception:
                    pass
            # 字体 → 字号配对计数（混用字体时每种字体各取自己的众数字号）
            if fn and fs is not None:
                try:
                    family_size_counter.setdefault(fn, Counter())[
                        round(float(fs.pt), 1)] += 1
                except Exception:
                    pass

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
        _sample_paragraph(para)
        sp = _line_spacing_of_para(para)
        if sp is not None:
            spacing_samples.append(sp)
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

    if font_counter:
        prof.font_family = font_counter.most_common(1)[0][0]
    if family_size_counter:
        prof.font_size_by_family = {
            str(f): float(c.most_common(1)[0][0])
            for f, c in family_size_counter.items() if c
        }
    if size_counter:
        prof.font_size_pt = size_counter.most_common(1)[0][0]
    if spacing_samples:
        prof.line_spacing = Counter(round(s, 2) for s in spacing_samples).most_common(1)[0][0]

    prof.heading_styles = list(dict.fromkeys(heading_styles))
    prof.heading_levels_used = max(heading_levels) if heading_levels else 0

    # —— 每级标题格式采样（学校模板驱动：取每级首个标题段为代表性样本）——
    # 每级只取第一个样本，避免同层多个标题互相平均掩盖差异。
    level_samples: dict = {}
    seen_levels: set = set()
    for para in paras:
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
        _pseudo = pseudo_heading_level_map(paras, exclude_ps=_skip_ps)
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
        for i in ref_indices:
            para_i = paras[i]
            for r in para_i.runs:
                if not (r.text or "").strip():
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
            # 首行缩进：取摘要正文段的众数（与正文缩进口径一致，三级回退）
            _abs_indents: Counter = Counter()
            for _j in _abs_body:
                _fli = _inch(_pf_attr(paras[_j], doc_pf, "first_line_indent"))
                _abs_indents[0.0 if _fli is None else _fli] += 1
            if _abs_indents:
                prof.abstract_first_line_indent_in = _abs_indents.most_common(1)[0][0]
            # v2.3.76：摘要正文段的段前 / 段后（磅），取众数；读不到按 0 磅计入
            _abs_sb: Counter = Counter()
            _abs_sa: Counter = Counter()
            for _j in _abs_body:
                _b, _a = _spacing_pair_of_para(paras[_j])
                _abs_sb[0.0 if _b is None else _b] += 1
                _abs_sa[0.0 if _a is None else _a] += 1
            if _abs_sb:
                prof.abstract_space_before_pt = _abs_sb.most_common(1)[0][0]
            if _abs_sa:
                prof.abstract_space_after_pt = _abs_sa.most_common(1)[0][0]
            # 是否另起一页：段前分页 或 上一段末尾分页符（复用参考文献同一口径）
            _prev_abs = paras[_abs_head - 1] if _abs_head > 0 else None
            prof.abstract_new_page = _reference_starts_new_page(
                _prev_abs, paras[_abs_head])
    except Exception:
        prof.abstract_first_line_indent_in = None
        prof.abstract_new_page = None
        prof.abstract_space_before_pt = None
        prof.abstract_space_after_pt = None

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
