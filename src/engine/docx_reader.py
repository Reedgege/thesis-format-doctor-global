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
from docx.oxml.ns import qn

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
    font_family: Optional[str] = None                  # 正文最常见字体（含表格文字）
    font_size_pt: Optional[float] = None               # 正文最常见字号（含表格文字）
    line_spacing: Optional[float] = None               # 倍数（固定值/最小值返回 None）
    body_alignment: Optional[str] = None               # left / justify / center / right
    first_line_indent_in: Optional[float] = None       # 正文首行缩进（英寸）
    space_after_pt: Optional[float] = None             # 段后间距（磅）
    table_paragraph_count: int = 0                     # 纳入采样的表格段落数
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
    header_text: str = ""                              # 页眉可见文本（剔除 PAGE 域结果，用于 running head 长度/内容检查）
    page_number_present: bool = False                 # 页眉/页脚中是否存在 PAGE 域（页码）
    page_number_right_aligned: Optional[bool] = None  # 含页码段落是否右对齐（APA 要求 top-right）；无法判定=None
    header_distance_in: Optional[float] = None       # 页眉距页顶（英寸）；MLA 9th 要求 ½″，python-docx 读 section.header_distance
    reference_new_page: Optional[bool] = None         # 参考文献标题是否另起一页（分页符/段前分页）；无参考文献段=None
    has_abstract: bool = False                        # 文档是否含 Abstract / 摘要 段落（专业版要求）
    has_author_note: bool = False                     # 文档是否含 Author Note / 作者注 段落（专业版要求）
    # —— MLA 专用探测字段（Purdue OWL general_format + 样例论文）——
    title_text: str = ""                              # 首页标题段文本（首个居中段落）
    title_alignment: Optional[str] = None             # 标题对齐标签（center/left/...）
    title_bold: bool = False                          # 标题是否含任意粗体 run
    title_italic: bool = False                        # 标题是否含任意斜体 run
    title_underline: bool = False                     # 标题是否含任意下划线 run
    first_page_left_lines: int = 0                    # 标题之前左上对齐行数（MLA 首页信息块：姓名/教授/课程/日期 共 4 行）
    block_quote_candidates: list = field(default_factory=list)  # 长段落候选：{words, left_indent_in, has_quote}
    # 每级标题格式样本（学校模板驱动：层级 → 代表性样本）。仅记录可从 docx 读出的属性；
    # number_prefix 为编号形态的最佳猜测（读不到为 None）。
    heading_formats: dict = field(default_factory=dict)         # level(int) -> HeadingFormatSample


@dataclass
class HeadingFormatSample:
    """从 Word 模板抽出的单级标题直接格式化样本（仅可读属性，绝不写 styleId）。"""

    alignment: Optional[str] = None      # left / center / right / justify
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    indent: Optional[bool] = None        # 标题段是否缩进（左缩进或首行缩进 > 0）
    size_pt: Optional[float] = None      # 字号（磅）
    number_prefix: Optional[str] = None  # 编号形态：number/roman/upper_letter/lower_letter/lower_roman/number_paren/None
    uppercase: Optional[bool] = None     # 标题文本是否全大写

    def to_dict(self) -> dict:
        return {
            "alignment": self.alignment, "bold": self.bold, "italic": self.italic,
            "indent": self.indent, "size_pt": self.size_pt,
            "number_prefix": self.number_prefix, "uppercase": self.uppercase,
        }


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


def _sample_heading_format(para, doc_pf) -> "HeadingFormatSample":
    """采样一个代表性标题段落的直接格式化（用于学校模板驱动）。

    对齐 / 加粗 / 斜体 / 字号均三级回退（run → 段落样式 → 文档默认），
    与正文采样口径一致；编号前缀 / 全大写按标题文本启发式识别。
    """
    # 对齐
    al_raw = para.alignment
    al = _ALIGN_MAP.get(al_raw) if al_raw is not None else None
    if al is None:
        al_raw2 = _pf_attr(para, doc_pf, "alignment")
        al = _ALIGN_MAP.get(al_raw2) if al_raw2 is not None else None

    # 加粗 / 斜体：整段 run 全为某状态才记为 True/False；有混合则回退样式
    runs = [r for r in para.runs if (r.text or "").strip()]
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    if runs:
        bold = all(getattr(r.font, "bold", None) is True for r in runs)
        italic = all(getattr(r.font, "italic", None) is True for r in runs)
    if bold is None or italic is None:
        try:
            st = para.style
            if st is not None and st.font is not None:
                if bold is None:
                    bold = st.font.bold is True
                if italic is None:
                    italic = st.font.italic is True
        except Exception:
            pass

    # 字号：先样式（标题格式常由样式承载），再 run 最大
    size: Optional[float] = None
    try:
        st = para.style
        if st is not None and st.font is not None and st.font.size is not None:
            size = round(float(st.font.size.pt), 1)
    except Exception:
        size = None
    if size is None:
        size = _max_run_size_pt(para)

    # 缩进：左缩进或首行缩进 > 0 视为缩进式标题
    li = _inch(_pf_attr(para, doc_pf, "left_indent"))
    fli = _inch(_pf_attr(para, doc_pf, "first_line_indent"))
    indent = bool((li is not None and li > 0) or (fli is not None and fli > 0))

    # 全大写：标题文本去掉空格后全为字母且等于大写
    txt = para.text.strip()
    uppercase: Optional[bool] = None
    if txt and txt.replace(" ", "").isalpha():
        uppercase = txt == txt.upper()

    # 编号前缀：启发式识别标题起始编号形态
    number_prefix = _detect_number_prefix(txt)

    return HeadingFormatSample(
        alignment=al, bold=bold, italic=italic, indent=indent,
        size_pt=size, number_prefix=number_prefix, uppercase=uppercase,
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


def heading_level(para) -> Optional[int]:
    """公开的标题层级访问器（fixer 套用每级格式规则用）。

    返回 1-based 层级（1=最高级）或 None（非标题 / 伪标题无明确层级）。
    判定链路：样式名 → 段落级 outlineLvl → 样式级 outlineLvl，兼容自定义标题样式。
    """
    return _heading_level(para)


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


def _read_header_footer_info(doc) -> tuple[str, bool, Optional[bool]]:
    """返回 (running_head_text, page_number_present, page_number_right_aligned)。

    running_head_text：拼接所有分节页眉文本并剔除 PAGE 域结果；
    page_number_present：任意页眉/页脚含 PAGE 域；
    page_number_right_aligned：含 PAGE 域的段落中是否有右对齐者（None=未能判定）。
    """
    texts: list = []
    page_present = False
    right_aligned: Optional[bool] = None
    try:
        sections = list(doc.sections)
    except Exception:
        sections = []
    for sec in sections:
        for part in (getattr(sec, "header", None), getattr(sec, "footer", None)):
            if part is None:
                continue
            try:
                para_list = part.paragraphs
            except Exception:
                continue
            for p in para_list:
                if _para_has_page_field(p):
                    page_present = True
                    al = getattr(p, "alignment", None)
                    if al == WD_ALIGN_PARAGRAPH.RIGHT:
                        right_aligned = True
                    elif al is not None and al != WD_ALIGN_PARAGRAPH.RIGHT and right_aligned is None:
                        right_aligned = False
                t = _para_text_excl_fields(p)
                if t:
                    texts.append(t)
    return " ".join(texts).strip(), page_present, right_aligned


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

    # —— 页边距：遍历所有分节取众数；不一致则标记 ——
    per_side: dict = {s: [] for s in ("top", "bottom", "left", "right")}
    col_samples: list = []
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
    skip = set(ref_indices)
    if ref_start is not None:
        skip.add(ref_start)
    body_paras = [p for i, p in enumerate(paras)
                  if i not in skip and p.text.strip() and not is_heading(p)]

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

    # —— 块引用候选收集（MLA：超 4 行引文应缩进 0.5″、不加引号；此处只收集长段落，由 checker 按风格判断）——
    for p in body_paras:
        words = len(p.text.split())
        if words < 50:
            continue
        left_indent = _inch(_pf_attr(p, doc_pf, "left_indent"))
        has_quote = any(q in p.text for q in ('"', "“", "”", "'"))
        prof.block_quote_candidates.append({
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
    for para in body_paras:
        al_raw = _pf_attr(para, doc_pf, "alignment")
        al = _ALIGN_MAP.get(al_raw) if al_raw is not None else "left"
        if al:
            align_counter[al] += 1
        fli = _inch(_pf_attr(para, doc_pf, "first_line_indent"))
        indent_counter[0.0 if fli is None else fli] += 1
        sa = _pt(_pf_attr(para, doc_pf, "space_after"))
        space_after_counter[0.0 if sa is None else sa] += 1

    if align_counter:
        prof.body_alignment = align_counter.most_common(1)[0][0]
    if indent_counter:
        prof.first_line_indent_in = indent_counter.most_common(1)[0][0]
    if space_after_counter:
        prof.space_after_pt = space_after_counter.most_common(1)[0][0]

    # —— 字体 / 字号 / 行距 / 标题（正文段落 + 表格段落）——
    font_counter: Counter = Counter()
    size_counter: Counter = Counter()
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
        sp = _line_spacing_of(para.paragraph_format)
        if sp is not None:
            spacing_samples.append(sp)
        _sample_paragraph(para)

    # 表格内文字一并采样（论文三线表常见，漏掉会让「正文最常见字体/字号」判断偏）
    table_count = 0
    for para in _iter_table_paragraphs(doc):
        if not para.text.strip():
            continue
        table_count += 1
        _sample_paragraph(para)
        sp = _line_spacing_of(para.paragraph_format)
        if sp is not None:
            spacing_samples.append(sp)
    prof.table_paragraph_count = table_count

    if font_counter:
        prof.font_family = font_counter.most_common(1)[0][0]
    if size_counter:
        prof.font_size_pt = size_counter.most_common(1)[0][0]
    if spacing_samples:
        prof.line_spacing = Counter(round(s, 2) for s in spacing_samples).most_common(1)[0][0]

    prof.heading_styles = list(dict.fromkeys(heading_styles))
    prof.heading_levels_used = max(heading_levels) if heading_levels else 0

    # —— 每级标题格式采样（学校模板驱动：取每级首个标题段为代表性样本）——
    # 仅真实标题层级参与（伪标题无明确层级，跳过）。每级只取第一个样本，
    # 避免同层多个标题互相平均掩盖差异。
    level_samples: dict = {}
    seen_levels: set = set()
    for para in paras:
        lvl = _heading_level(para)
        if lvl is None or not para.text.strip():
            continue
        if lvl in seen_levels:
            continue
        seen_levels.add(lvl)
        level_samples[lvl] = _sample_heading_format(para, doc_pf)
    if level_samples:
        prof.heading_formats = {lvl: level_samples[lvl] for lvl in sorted(level_samples)}

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

    # —— 页眉 / 页码 / 参考文献另起一页（轻量读取，仅用于保守提示，不修改文档）——
    try:
        rh_text, page_present, page_right = _read_header_footer_info(doc)
        prof.header_text = rh_text
        prof.page_number_present = page_present
        prof.page_number_right_aligned = page_right
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

    return prof


def reference_paragraph_indices(doc) -> list:
    """返回参考文献段落在 doc.paragraphs 中的下标列表。

    复用与 read_docx 完全一致的 find_reference_section（唯一实现），
    供 fixer 与测试精确区分「正文段落」与「参考文献段落」。
    """
    return find_reference_section(list(doc.paragraphs))[1]
