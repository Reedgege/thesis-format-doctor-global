# -*- coding: utf-8 -*-
"""comment_spec_en.py —— 英文模板 Word 批注 → 结构化格式规范（纯本地、确定性）。

为什么需要
----------
海外版「学校模板」模式此前只读**实测格式**（样式 + 直接格式）。Word 批注里写明的
规范要求（"Heading 1: Times New Roman 18 pt, bold, centered"）从未被读取。
本模块补上「解析」这一侧：把 ``docx_reader.read_template_comments()`` 取出的英文
批注解析成与 ``TargetProfile`` 同名的字段，供 ``target_from_template_docx`` 以
「批注 > 实测」的顺序合并。

设计约束
--------
- **纯英文**：只认英文类别词与英文单位（pt / inch / cm / mm / line / spaced）。
  中文批注不解析（海外版面向英文用户，不做双语，避免半吊子映射）。
- **类别词门控**：只有「批注开头是已知类别词（可带冒号）」或「命中强特征模式」
  （three-line table / after each table / APA 7th …）才当规范；其余进 ``unparsed``，
  绝不把 "fix this later" 这类编辑性批注误读成格式规则。
- **多段批注**：一条批注里可以用 "Abstract heading: … Abstract body: …" 写两条规则，
  按「类别词 + 冒号」切片逐段解析。
- **只输出 TargetProfile 认识的键**，不臆造字段（否则 ``TargetProfile(**extra)`` 会炸）。

输出契约
--------
``parse_comments(comments)`` 返回::

    {
      "page":      {margin_top_in, font_family, font_size_pt, line_spacing,
                    body_alignment, first_line_indent_in, space_before_pt,
                    space_after_pt, page_number_position, ...},
      "title_page":{title_page_title_centered, title_page_title_bold,
                    title_page_space_before_pt, title_page_space_after_pt},
      "abstract":  {abstract_heading_alignment/bold/italic, abstract_first_line_indent_in,
                    abstract_space_before_pt/after_pt, keywords_label_italic,
                    keywords_indent_in},
      "caption":   {caption_font_size_pt, caption_line_spacing,
                    caption_space_before_pt/after_pt,
                    caption_position_table/figure},
      "table":     {table_space_before_pt/after_pt, table_gap_pt},
      "equation":  {equation_alignment, equation_number_position, equation_size_pt,
                    equation_font_family, equation_line_spacing,
                    equation_space_before_pt/after_pt},
      "structural":{"columns": int, "paper_size": str, "table_lines": "three_line"},
      "heading_formats": {1: {...与 HeadingFormatSample.to_dict() 同键...}},
      "title_format":    {...论文大标题的 attrs，供报告/标题页用...},
      "reference": {reference_style, reference_hanging_indent_in, reference_numbered,
                    reference_font_family, reference_size_pt, reference_line_spacing,
                    reference_new_page, reference_heading_bold,
                    reference_heading_centered},
      "count": int,          # 参与解析的批注条数
      "unparsed": [{"id", "text"}],   # 未识别（诊断用）
      "notes": [str],        # 英文提示（给用户看的）
    }

所有「未提到」的字段**不出现**在结果里（而不是写 None）——合并侧按
「批注有的键才覆盖实测」处理，与国内版「批注 override、样式兜底」同一口径。
"""

from __future__ import annotations

import re
from typing import Optional

# ---------------------------------------------------------------------------
# 基础词表
# ---------------------------------------------------------------------------
# 常见西文字体（模板里几乎只用这些）。按长度倒序匹配，避免 "Times" 抢在
# "Times New Roman" 前面命中半截。
_FONT_NAMES = (
    "Times New Roman", "Book Antiqua", "Courier New", "Century Schoolbook",
    "Palatino Linotype", "Trebuchet MS", "Comic Sans MS", "Cambria Math",
    "Helvetica Neue", "Lucida Sans", "Segoe UI", "Georgia", "Garamond",
    "Cambria", "Calibri", "Arial", "Helvetica", "Verdana", "Tahoma",
    "Palatino", "Century", "Courier", "Consolas", "Times", "SimSun",
)

_PAPER_ALIASES = {
    "letter": "Letter", "a4": "A4", "legal": "Legal", "a5": "A5",
}

_ALIGN_PATTERNS = (
    # 🔴 必须认英式拼写 "centred"。旧写法 ``cent(?:er|re)(?:ed|ing)?`` 只覆盖
    # "centered / centering / center / centre" —— "centred" = cent+re+**d**，
    # 末尾那个孤零零的 d 让 ``\b`` 落在两个字母之间，整条匹配失败。英联邦（英/澳/加/
    # 印）大学的模板大量写 "centred"，于是「批注说居中」被静默丢掉，客户照旧居中不了。
    (re.compile(r"\bcent(?:er|re)(?:ed|d|ing)?\b", re.I), "center"),
    (re.compile(r"\bjustified\b|\bjustify\b|\bflush\s+left\s+and\s+right\b", re.I), "justify"),
    (re.compile(r"\bleft[-\s]?aligned\b|\bflush\s+left\b", re.I), "left"),
    (re.compile(r"\bright[-\s]?aligned\b|\bflush\s+right\b", re.I), "right"),
)

_LS_PATTERNS = (
    (re.compile(r"\bsingle[-\s]?spaced?\b|\bsingle\s+line\s+spacing\b", re.I), 1.0),
    (re.compile(r"\bdouble[-\s]?spaced?\b|\bdouble\s+line\s+spacing\b", re.I), 2.0),
    (re.compile(r"\b1\.5[-\s]?(?:spaced?|lines?)\b|\bone\s+and\s+a\s+half\s+(?:line\s+)?spac", re.I), 1.5),
    (re.compile(r"\btriple[-\s]?spaced?\b", re.I), 3.0),
)
_LS_VALUE_RE = re.compile(
    r"([0-3](?:\.\d+)?)\s*(?:x\s*)?line\s+spacing"
    r"|line\s+spacing\s*(?:of|:|=)?\s*([0-3](?:\.\d+)?)", re.I)

_SIZE_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(?:pt|points?)\b", re.I)
_SIZE_HYPHEN_RE = re.compile(r"(\d+(?:\.\d+)?)[-\s]point\b", re.I)

_BOLD_NEG_RE = re.compile(r"\b(?:not\s+bold|non[-\s]?bold|unbold|no\s+bold)\b", re.I)
_BOLD_RE = re.compile(r"\bbold(?:face|ened)?\b", re.I)
_ITAL_NEG_RE = re.compile(r"\b(?:not\s+italic|non[-\s]?italic|no\s+italic)\b", re.I)
_ITAL_RE = re.compile(r"\bitalic(?:s|ized)?\b", re.I)
_UPPER_RE = re.compile(r"\ball\s+caps\b|\bupper[-\s]?case\b", re.I)

_SP_BEFORE_RE = re.compile(
    r"(?:space\s+(?:before|above)\s*(?:of\s*)?([\d.]+)\s*(?:pt|points?))"
    # "6 pt space before" —— 数字前置、中间隔一个 space 的写法，模板里同样高频。
    r"|(?:([\d.]+)\s*(?:pt|points?)\s+(?:of\s+)?(?:space|spacing)\s+(?:before|above)\b)"
    r"|(?:([\d.]+)\s*(?:pt|points?)\s+(?:before|above)\b)", re.I)
_SP_AFTER_RE = re.compile(
    r"(?:space\s+(?:after|below)\s*(?:of\s*)?([\d.]+)\s*(?:pt|points?))"
    r"|(?:([\d.]+)\s*(?:pt|points?)\s+(?:of\s+)?(?:space|spacing)\s+(?:after|below)\b)"
    r"|(?:([\d.]+)\s*(?:pt|points?)\s+(?:after|below)\b)", re.I)

# 长度单位：必须同时认「inch」与「inches」。
# 🔴 写成 ``inches?`` 是错的（= "inche" + 可选 s），单数 inch 匹配不到 —— 实测让
# 页边距/首行缩进/悬挂缩进整片漏解析。统一走这个常量。
_UNIT = r"(inch(?:es)?|in\b|\"|”|cm|mm)"

_NO_INDENT_RE = re.compile(
    r"\b(?:no|without\s+(?:a|any))\s+(?:first[-\s]line\s+)?indent(?:ation)?\b", re.I)
_HANGING_RE = re.compile(
    r"\bhanging\s+indent(?:ation)?\s*(?:of\s*)?([\d.]+)\s*" + _UNIT, re.I)
_FIRST_INDENT_RE = re.compile(
    r"\b(?:first[-\s]line\s+indent(?:ation)?|indented?)\s*(?:of\s*)?"
    r"([\d.]+)\s*" + _UNIT, re.I)
_INDENT_CHARS_RE = re.compile(
    r"\b(?:first[-\s]line\s+)?indent(?:ation)?\s*(?:of\s*)?([\d.]+)\s*(?:characters?|chars?)\b", re.I)

_MARGIN_SIDE_RE = re.compile(
    r"\b(top|bottom|left|right)\s*(?:margin)?\s*(?:of\s*|:\s*)?([\d.]+)\s*" + _UNIT, re.I)
_MARGIN_ALL_RE = re.compile(
    r"\b(?:all\s+)?margins?\s*(?:of|:|are)?\s*([\d.]+)\s*" + _UNIT, re.I)

_COL_TWO_RE = re.compile(
    r"\b(?:two|2)[-\s]column|\bdouble[-\s]column|\btwo\s+columns?\b", re.I)
_COL_ONE_RE = re.compile(
    r"\b(?:single|one|1)[-\s]column|\bone\s+column\b", re.I)

_THREE_LINE_RE = re.compile(
    r"\bthree[-\s]line\b|\bno\s+vertical\s+lines?\b|\bwithout\s+vertical\s+lines?\b"
    r"|\bhorizontal\s+(?:rules?|lines?)\s+only\b|\bformal\s+table\b", re.I)
# 纸型「纯尺寸」写法："210 x 297 mm" / "8.5 x 11 inches"。
_PAPER_DIM_RE = re.compile(
    r"([\d.]+)\s*(?:x|×|by)\s*([\d.]+)\s*" + _UNIT, re.I)
_GAP_PT_RE = re.compile(
    r"([\d.]+)\s*(?:pt|points?)\s+(?:of\s+)?(?:space\s+)?(?:after|before)\s+"
    r"(?:each\s+|the\s+)?table\b", re.I)
_GAP_LINES_RE = re.compile(
    r"\b(one|two|three|\d+)\s+(?:blank|empty)\s+lines?\s+(?:of\s+(?:space|spacing)\s+)?"
    r"(?:after|before)\s+(?:each\s+|the\s+)?table\b", re.I)

# —— 公式（v2.3.112）——
# 类别词。**必须能吃掉紧跟的限定词**（"Equation font:" / "Equation numbering:"）：
# 否则 `_split_parts` 会把 "Equation font: Times New Roman 12 pt" 切成 category=font，
# 于是这条公式批注被当成**正文**字体规则，把客户全篇正文改成公式字体 —— 一个批注
# 毁一篇稿。加限定词后标记命中位置提前到 "Equation font"，冒号紧随其后，切对了。
_EQ_WORD_RE = re.compile(
    r"\b(?:equations?|formulas?)"
    r"(?:\s*(?:and|&|/)\s*(?:equations?|formulas?))?"
    r"(?:\s+(?:fonts?|typeface|numbering|numbers?|alignment|size|sizes|spacing|"
    r"layout|format|formatting))?\b", re.I)
# 「编号靠右」：先认「编号词 + 靠右说法」**同时出现**，而不是把两句写成一条紧邻正则。
# 模板的实际写法太散 —— "numbered consecutively on the right" / "are numbered flush
# right" / "numbers at the right margin" / "numbering right-aligned"，穷举紧邻形态
# 一定会漏。这里只要求两个特征在同一段里共现，覆盖度远高于穷举。
_EQ_NUM_WORD_RE = re.compile(r"\bnumber(?:s|ed|ing)?\b", re.I)
_EQ_RIGHT_HINT_RE = re.compile(
    r"\b(?:on|at|to)\s+the\s+right\b|\bright[-\s]?aligned\b|\bflush\s+right\b"
    r"|\bright(?:[-\s]hand)?\s+margin\b", re.I)
# 「不许编号」：not numbered / no numbering / unnumbered / do not number / without numbers
_EQ_NO_NUMBER_RE = re.compile(
    r"\b(?:do\s+not|don'?t|never)\s+number\b|\bwithout\s+(?:an?\s+)?numbers?\b", re.I)
# 「段前段后都不要额外空」：no extra space / without extra spacing / no blank lines
_EQ_NO_SPACE_RE = re.compile(
    r"\b(?:no|without|zero)\s+(?:extra\s+|additional\s+)?"
    r"(?:space|spacing|blank\s+lines?|gap)\b", re.I)

_APA_RE = re.compile(r"\bAPA\b", re.I)
_MLA_RE = re.compile(r"\bMLA\b", re.I)
_IEEE_RE = re.compile(r"\bIEEE\b", re.I)
_NUMBERED_RE = re.compile(r"\bnumbered\b|\bnumbering\b|\[\d+\]", re.I)
_NOT_NUMBERED_RE = re.compile(r"\b(?:not|no)\s+number(?:ed|ing)?\b|\bunnumbered\b", re.I)

_ABOVE_RE = re.compile(r"\b(?:above|on\s+top\s+of)\b", re.I)
_BELOW_RE = re.compile(r"\b(?:below|under|beneath)\b", re.I)

_PAGE_NUM_RE = re.compile(r"\bpage\s+numbers?\b", re.I)

# ---------------------------------------------------------------------------
# 类别词（顺序即优先级：越具体的越靠前）
# ---------------------------------------------------------------------------
_NUM_WORDS = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
    "first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5, "sixth": 6,
}


def _num(token: str) -> Optional[int]:
    t = (token or "").strip().lower()
    if t.isdigit():
        return int(t)
    return _NUM_WORDS.get(t)


# (类别, 正则) —— 用于「类别词 + 冒号」切片；顺序敏感
# 标题层级的三种英文写法（**必须全部覆盖**，缺一种就整条批注漏解析）：
#   · "Heading 1" / "Level 3 headings"      → 数字在后
#   · "H2"                                  → 缩写
#   · "First-level heading" / "second level heading" → 数字在前
_HEADING_AFTER_RE = re.compile(
    r"(?:headings?|level)\s*#?\s*(\d+|one|two|three|four|five|six|"
    r"first|second|third|fourth|fifth|sixth)\b", re.I)
_HEADING_H_RE = re.compile(r"\bh([1-6])\b", re.I)
_HEADING_BEFORE_RE = re.compile(
    r"\b(\d+|one|two|three|four|five|six|first|second|third|fourth|fifth|sixth)"
    r"(?:st|nd|rd|th)?[-\s]*(?:level|order)?[-\s]*headings?\b", re.I)

_PART_MARKERS = (
    ("caption_table", re.compile(r"tables?\s+(?:captions?|titles?|headings?)", re.I)),
    ("caption_figure", re.compile(r"figures?\s+(?:captions?|titles?|headings?)", re.I)),
    ("abstract_heading", re.compile(r"abstract\s+(?:heading|title)", re.I)),
    ("abstract_body", re.compile(r"abstract\s+(?:body|text|paragraphs?)", re.I)),
    ("title", re.compile(r"(?:paper\s+)?title", re.I)),
    ("heading", _HEADING_AFTER_RE),
    ("heading", _HEADING_H_RE),
    ("heading", _HEADING_BEFORE_RE),
    ("keywords", re.compile(r"key\s*words?", re.I)),
    # v2.3.112：公式（equation / formula）。放在 body 之前 —— `_category_of` 按本元组
    # 顺序取**第一个**命中的类别，公式比"正文"更具体，不能被 body 抢走。
    ("equations", _EQ_WORD_RE),
    ("body", re.compile(r"body\s+(?:text|paragraphs?)|main\s+text", re.I)),
    ("margins", re.compile(r"margins?", re.I)),
    ("paper_size", re.compile(r"(?:paper|page)\s+size", re.I)),
    ("columns", re.compile(r"columns?|column\s+layout|\blayout\b", re.I)),
    ("tables", re.compile(r"tables?", re.I)),
    ("references", re.compile(r"references?|reference\s+list|bibliography|works\s+cited", re.I)),
    ("toc", re.compile(r"table\s+of\s+contents|\bcontents\b", re.I)),
    ("page_number", re.compile(r"page\s+numbers?", re.I)),
    ("line_spacing", re.compile(r"line\s+spacing", re.I)),
    ("indent", re.compile(r"(?:first[-\s]line\s+)?indent(?:ation)?", re.I)),
    ("font", re.compile(r"fonts?", re.I)),
)


def _heading_dict() -> dict:
    """与 ``HeadingFormatSample.to_dict()`` 同键的空样本（只填批注提到的值）。"""
    return {
        "alignment": None, "bold": None, "italic": None, "indent": None,
        "indent_in": None, "font_family": None, "size_pt": None,
        "number_prefix": None, "uppercase": None, "page_break_before": None,
        "space_before_pt": None, "space_after_pt": None, "line_spacing": None,
    }


# ---------------------------------------------------------------------------
# 单位换算
# ---------------------------------------------------------------------------
def _to_inches(value: float, unit: str) -> Optional[float]:
    u = (unit or "").strip().lower().lstrip()
    if u.startswith("in") or u in ('"', "”"):
        return round(float(value), 3)
    if u.startswith("cm"):
        return round(float(value) / 2.54, 3)
    if u.startswith("mm"):
        return round(float(value) / 25.4, 3)
    return None


# ---------------------------------------------------------------------------
# 属性解析（字体 / 字号 / 粗斜 / 对齐 / 行距 / 段前段后 / 缩进）
# ---------------------------------------------------------------------------
def _find_font(text: str) -> Optional[str]:
    low = text.lower()
    for name in sorted(_FONT_NAMES, key=len, reverse=True):
        if re.search(r"\b" + re.escape(name.lower()) + r"\b", low):
            return name
    return None


def _find_size(text: str) -> Optional[float]:
    m = _SIZE_RE.search(text) or _SIZE_HYPHEN_RE.search(text)
    if not m:
        return None
    try:
        v = float(m.group(1))
    except (TypeError, ValueError):
        return None
    return v if 4.0 <= v <= 72.0 else None


def _find_alignment(text: str) -> Optional[str]:
    for rx, label in _ALIGN_PATTERNS:
        if rx.search(text):
            return label
    return None


def _find_line_spacing(text: str) -> Optional[float]:
    for rx, val in _LS_PATTERNS:
        if rx.search(text):
            return val
    m = _LS_VALUE_RE.search(text)
    if m:
        raw = next((g for g in m.groups() if g), None)
        if raw is None:
            return None
        try:
            return round(float(raw), 2)
        except (TypeError, ValueError):
            return None
    return None


def _find_bool(text: str, pos_re, neg_re) -> Optional[bool]:
    """粗体/斜体三态：明确否定 → False；明确肯定 → True；未提 → None。"""
    if neg_re.search(text):
        return False
    if pos_re.search(text):
        return True
    return None


def _find_space_pt(text: str, rx) -> Optional[float]:
    m = rx.search(text)
    if not m:
        return None
    # 正则每多一条备选就多一个捕获组：一律扫「第一个非空组」，别再写死 group(1)/group(2)
    # （加第三条 "6 pt space before" 时，写死分组会把新写法静默读成 None）。
    raw = next((g for g in m.groups() if g), None)
    if raw is None:
        return None
    try:
        v = float(raw)
    except (TypeError, ValueError):
        return None
    return round(v, 1) if 0.0 <= v <= 200.0 else None


def _find_first_line_indent_in(text: str, font_size_pt: float) -> Optional[float]:
    if _NO_INDENT_RE.search(text):
        return 0.0
    m = _FIRST_INDENT_RE.search(text)
    if m:
        return _to_inches(float(m.group(1)), m.group(2))
    m = _INDENT_CHARS_RE.search(text)
    if m:
        # 英文模板偶尔用 "2 characters"；按「1 字符 ≈ 0.5 em」折算（12pt → 0.083 in/字符），
        # 结果与 Word 的「首行缩进 N 字符」同量级。折算不确定，故只在没有绝对值时使用。
        try:
            chars = float(m.group(1))
        except (TypeError, ValueError):
            return None
        return round(chars * 0.5 * float(font_size_pt) / 72.0, 3)
    return None


def _parse_attrs(text: str, font_size_hint: float = 12.0) -> dict:
    """通用属性：只返回**文本里确实提到**的键。"""
    out: dict = {}
    fam = _find_font(text)
    if fam:
        out["font_family"] = fam
    size = _find_size(text)
    if size is not None:
        out["size_pt"] = size
    bold = _find_bool(text, _BOLD_RE, _BOLD_NEG_RE)
    if bold is not None:
        out["bold"] = bold
    ital = _find_bool(text, _ITAL_RE, _ITAL_NEG_RE)
    if ital is not None:
        out["italic"] = ital
    al = _find_alignment(text)
    if al:
        out["alignment"] = al
    ls = _find_line_spacing(text)
    if ls is not None:
        out["line_spacing"] = ls
    sb = _find_space_pt(text, _SP_BEFORE_RE)
    if sb is not None:
        out["space_before_pt"] = sb
    sa = _find_space_pt(text, _SP_AFTER_RE)
    if sa is not None:
        out["space_after_pt"] = sa
    ind = _find_first_line_indent_in(text, size or font_size_hint)
    if ind is not None:
        out["first_line_indent_in"] = ind
    if _UPPER_RE.search(text):
        out["uppercase"] = True
    return out


# ---------------------------------------------------------------------------
# 切片：一条批注里可能写多条规则（"Abstract heading: … Abstract body: …"）
# ---------------------------------------------------------------------------
def _split_parts(text: str) -> list:
    """按「类别词 + 冒号」把批注切成 [(category, segment), ...]。

    没有冒号标记时整条作为一段，类别交给 ``_category_of`` 判定。
    """
    marks = []
    for cat, rx in _PART_MARKERS:
        for m in rx.finditer(text):
            # 只认「标记后紧跟冒号」的形态（"Heading 1: …"），避免把正文里的
            # 普通名词（"… after each table …"）当成新规则起点。
            tail = text[m.end():m.end() + 4]
            if not re.match(r"\s*:", tail):
                continue
            # 同一位置只留最具体的类别（_PART_MARKERS 已按优先级排序）
            #
            # 🔴 必须按**区间重叠**挡，不能只挡「位置相差 <2」。旧写法挡不住「短标记
            # 嵌在长标记内部」：``"Equation font: Cambria Math 12 pt"`` 会同时命中
            # equations(0..14) 与 font(9..13)，于是前半段被切成孤零零的 "Equation"
            # （没有任何规则可解析），后半段 "font: Cambria Math 12 pt" 落进**正文**
            # 字体通道 —— 一条关于公式的批注把整篇正文改成公式字体。
            if any(p <= m.start() < e for p, _c, e in marks):
                continue
            marks.append((m.start(), cat, m.end()))
    if not marks:
        return [(_category_of(text), text)]

    marks.sort(key=lambda x: x[0])
    parts = []
    # 冒号之前的引言（如 "For all headings: …"）归入第一段的类别
    head = text[:marks[0][0]].strip()
    for i, (pos, cat, end) in enumerate(marks):
        nxt = marks[i + 1][0] if i + 1 < len(marks) else len(text)
        seg = text[pos:nxt].strip()
        if head and i == 0:
            seg = (head + " " + seg).strip()
        parts.append((cat, seg))
    return parts


def _category_of(text: str) -> str:
    """没有冒号标记时的类别判定（整条批注一个类别）。"""
    low = text.lower()
    # 强特征优先：这些说法一旦出现，语义无歧义
    if _THREE_LINE_RE.search(text) or _GAP_PT_RE.search(text) or _GAP_LINES_RE.search(text):
        return "tables"
    for cat, rx in _PART_MARKERS:
        if rx.search(text):
            return cat
    if re.search(r"\bsingle[-\s]column\b|\btwo[-\s]column\b|\bdouble[-\s]column\b", low):
        return "columns"
    if _APA_RE.search(text) or _MLA_RE.search(text) or _IEEE_RE.search(text):
        return "references"
    return "unknown"


# ---------------------------------------------------------------------------
# 各维度落位
# ---------------------------------------------------------------------------
def _parse_margins(text: str, page: dict) -> bool:
    hit = False
    for m in _MARGIN_SIDE_RE.finditer(text):
        side = m.group(1).lower()
        inches = _to_inches(float(m.group(2)), m.group(3))
        if inches is not None:
            page["margin_%s_in" % side] = inches
            hit = True
    if not hit:
        m = _MARGIN_ALL_RE.search(text)
        if m:
            inches = _to_inches(float(m.group(1)), m.group(2))
            if inches is not None:
                for side in ("top", "bottom", "left", "right"):
                    page["margin_%s_in" % side] = inches
                hit = True
    return hit


def _parse_columns(text: str, structural: dict) -> bool:
    if _COL_TWO_RE.search(text):
        structural["columns"] = 2
        return True
    if _COL_ONE_RE.search(text):
        structural["columns"] = 1
        return True
    return False


def _parse_paper(text: str, structural: dict) -> bool:
    # ① 先认纸型名（Letter / A4 / A5 / Legal）
    m = re.search(r"\b(letter|a4|a5|legal)\b", text, re.I)
    if m:
        structural["paper_size"] = _PAPER_ALIASES[m.group(1).lower()]
        return True
    # ② 只写了尺寸（"210 x 297 mm" / "8.5 x 11 inches"）→ 复用引擎的识别器归一成纸型名。
    #    不自己写一套尺寸表：容差/横排归一必须与 docx_reader.identify_paper_size 同口径，
    #    否则「批注说 A4、实测也是 A4」会因为两套容差不一致而互相打架。
    m = _PAPER_DIM_RE.search(text)
    if m:
        unit = m.group(3)
        w = _to_inches(float(m.group(1)), unit)
        h = _to_inches(float(m.group(2)), unit)
        if w and h:
            from .docx_reader import identify_paper_size  # 局部导入：避免模块级循环依赖
            name = identify_paper_size(w, h)
            if name:
                structural["paper_size"] = name
                return True
    return False


def _parse_table_rules(text: str, table: dict, structural: dict, page: dict) -> bool:
    hit = False
    if _THREE_LINE_RE.search(text):
        structural["table_lines"] = "three_line"
        hit = True
    m = _GAP_PT_RE.search(text)
    if m:
        table["table_gap_pt"] = round(float(m.group(1)), 1)
        hit = True
    else:
        m = _GAP_LINES_RE.search(text)
        if m:
            lines = _num(m.group(1))
            if lines:
                # 「N 个空行」→ 磅值：一行 ≈ 字号 × 行距。用已解析的页面值，缺省 12pt × 2.0。
                size = float(page.get("font_size_pt") or 12.0)
                ls = float(page.get("line_spacing") or 2.0)
                table["table_gap_pt"] = round(size * ls * float(lines), 1)
                hit = True
    # 表格单元格内的段前/段后（"no extra space inside table cells"）
    if re.search(r"\b(?:inside|within)\s+(?:the\s+)?(?:table\s+)?cells?\b", text, re.I):
        if _NO_INDENT_RE.search(text) or re.search(r"\bno\s+(?:extra\s+)?spac", text, re.I):
            table["table_space_before_pt"] = 0.0
            table["table_space_after_pt"] = 0.0
            hit = True
    return hit


def _scrub_space_phrases(text: str) -> str:
    """把「间距」短语先抠掉，再交给字号识别。

    `_find_size` 只认「数字 + pt」，分不清 12 pt 说的是**字号**还是**间距**。公式批注里
    "6 pt space before and 6 pt after" 会被读成「公式字号 6 pt」——修正器于是把公式缩到
    6 磅，一眼可见的破坏。凡是要判字号的地方都必须先洗一遍。

    只在公式解析里用，不动 `_parse_attrs` 的既有行为：那边「字号写在前面」是常态
    （`_SIZE_RE.search` 取第一个匹配），贸然全局改会牵动已上线的正文/摘要/标题口径。
    """
    out = re.sub(r"[\d.]+\s*(?:pt|points?)\s+(?:of\s+)?(?:space|spacing|gap)\b",
                 " ", text, flags=re.I)
    out = re.sub(r"\b(?:space|spacing|gap)\s+(?:before|above|after|below)\s*(?:of\s*)?"
                 r"[\d.]+\s*(?:pt|points?)", " ", out, flags=re.I)
    # "6 pt after" / "12 pt before" —— 省掉 space 一词的写法同样会被误当字号。
    out = re.sub(r"[\d.]+\s*(?:pt|points?)\s+(?:before|above|after|below)\b",
                 " ", out, flags=re.I)
    return out


def _parse_equation_rules(text: str, equation: dict, page: dict) -> bool:
    """公式类批注 → 结构化规则。

    只落 ``TargetProfile`` 认识的键。三条纪律：
    1. **对齐只认 center / left**。段内出现 "right-aligned" 时它说的是**编号**的位置，
       不是公式本身 —— 顺手拿 `_find_alignment` 的结果会把「编号靠右」误读成
       「公式右对齐」，修正器于是把公式推到右边，越改越错。
    2. **"no extra space" 只在没写具体磅值时才兜底成 0**；写了 "6 pt before" 就用 6。
    3. 字体只**记录**、不自动改（Word 的数学区一旦换字体可能丢字形，见 fixer 说明）。
    """
    hit = False

    # ① 编号：靠右 / 不许编号
    if _NOT_NUMBERED_RE.search(text) or _EQ_NO_NUMBER_RE.search(text):
        equation["equation_number_position"] = "none"
        hit = True
    elif _EQ_NUM_WORD_RE.search(text) and _EQ_RIGHT_HINT_RE.search(text):
        equation["equation_number_position"] = "right"
        hit = True

    # ② 展示公式对齐：只认 center / left（right 归编号，见上述纪律 1）
    al = _find_alignment(text)
    if al in ("center", "left"):
        equation["equation_alignment"] = al
        hit = True

    # ③ 字号 / 字体（字号必须洗掉间距短语，见 _scrub_space_phrases）
    size = _find_size(_scrub_space_phrases(text))
    if size is not None:
        equation["equation_size_pt"] = size
        hit = True
    fam = _find_font(text)
    if fam:
        equation["equation_font_family"] = fam
        hit = True

    # ④ 行距
    ls = _find_line_spacing(text)
    if ls is not None:
        equation["equation_line_spacing"] = ls
        hit = True

    # ⑤ 段前 / 段后
    got_space = False
    sb = _find_space_pt(text, _SP_BEFORE_RE)
    if sb is not None:
        equation["equation_space_before_pt"] = sb
        got_space = True
    sa = _find_space_pt(text, _SP_AFTER_RE)
    if sa is not None:
        equation["equation_space_after_pt"] = sa
        got_space = True
    if not got_space and _EQ_NO_SPACE_RE.search(text):
        equation["equation_space_before_pt"] = 0.0
        equation["equation_space_after_pt"] = 0.0
        hit = True
    return hit or got_space


def _parse_reference(text: str, reference: dict) -> bool:
    hit = False
    style = None
    if _APA_RE.search(text):
        style = "apa"
    elif _MLA_RE.search(text):
        style = "mla"
    elif _IEEE_RE.search(text):
        style = "ieee"
    if style:
        reference["reference_style"] = style
        hit = True
    m = _HANGING_RE.search(text)
    if m:
        inches = _to_inches(float(m.group(1)), m.group(2))
        if inches is not None:
            reference["reference_hanging_indent_in"] = inches
            hit = True
    if _NOT_NUMBERED_RE.search(text):
        reference["reference_numbered"] = False
        hit = True
    elif _NUMBERED_RE.search(text):
        reference["reference_numbered"] = True
        hit = True
    attrs = _parse_attrs(text)
    if "font_family" in attrs:
        reference["reference_font_family"] = attrs["font_family"]
        hit = True
    if "size_pt" in attrs:
        reference["reference_size_pt"] = attrs["size_pt"]
        hit = True
    if "line_spacing" in attrs:
        reference["reference_line_spacing"] = attrs["line_spacing"]
        hit = True
    if re.search(r"\bnew\s+page\b|\bseparate\s+page\b|\bstart(?:s|ing)?\s+on\s+a\s+new\s+page\b", text, re.I):
        reference["reference_new_page"] = True
        hit = True
    return hit


def _parse_caption(text: str, caption: dict, which: str) -> bool:
    hit = False
    attrs = _parse_attrs(text)
    if "size_pt" in attrs:
        caption["caption_font_size_pt"] = attrs["size_pt"]
        hit = True
    if "line_spacing" in attrs:
        caption["caption_line_spacing"] = attrs["line_spacing"]
        hit = True
    if "space_before_pt" in attrs:
        caption["caption_space_before_pt"] = attrs["space_before_pt"]
        hit = True
    if "space_after_pt" in attrs:
        caption["caption_space_after_pt"] = attrs["space_after_pt"]
        hit = True
    if _ABOVE_RE.search(text):
        caption["caption_position_%s" % which] = "top"
        hit = True
    elif _BELOW_RE.search(text):
        caption["caption_position_%s" % which] = "bottom"
        hit = True
    return hit


def _parse_page_number(text: str, page: dict) -> bool:
    low = text.lower()
    if not _PAGE_NUM_RE.search(text):
        return False
    in_header = "header" in low
    in_footer = "footer" in low
    al = _find_alignment(text)
    if in_header and al in (None, "right", "center"):
        page["page_number_position"] = {"right": "top-right", "center": "top-center"}.get(al, "top-right")
        return True
    if in_footer and al:
        page["page_number_position"] = {"center": "bottom-center", "right": "bottom-right"}.get(al, "bottom-center")
        return True
    return False


def _apply_part(out: dict, cat: str, seg: str, meta: dict) -> bool:
    """把一段（类别 + 文本）落进输出结构；返回是否真的解析出东西。"""
    page = out["page"]
    font_size_hint = float(page.get("font_size_pt") or 12.0)

    if cat == "heading":
        # 三种写法按「后置数字 → H 缩写 → 前置数字」依次尝试；绝不用索引取正则
        # （_PART_MARKERS 顺序一调整就会静默取错，踩过）。
        lvl = None
        for rx in (_HEADING_AFTER_RE, _HEADING_H_RE, _HEADING_BEFORE_RE):
            m = rx.search(seg)
            if m and m.groups():
                lvl = _num(m.group(1))
                if lvl:
                    break
        if lvl is None:
            return False
        attrs = _parse_attrs(seg, font_size_hint)
        if not attrs:
            return False
        hf = _heading_dict()
        hf.update(attrs)
        # indent（布尔）由 indent_in 派生，保持与 HeadingFormatSample.to_dict() 同口径
        if hf.get("indent_in") is not None:
            hf["indent"] = float(hf["indent_in"]) > 0
        out["heading_formats"][lvl] = hf
        return True

    if cat == "title":
        attrs = _parse_attrs(seg, font_size_hint)
        if not attrs:
            return False
        out["title_format"] = attrs
        tp = out["title_page"]
        if "alignment" in attrs:
            tp["title_page_title_centered"] = (attrs["alignment"] == "center")
        if "bold" in attrs:
            tp["title_page_title_bold"] = bool(attrs["bold"])
        if "space_before_pt" in attrs:
            tp["title_page_space_before_pt"] = attrs["space_before_pt"]
        if "space_after_pt" in attrs:
            tp["title_page_space_after_pt"] = attrs["space_after_pt"]
        return True

    if cat == "abstract_heading":
        attrs = _parse_attrs(seg, font_size_hint)
        absd = out["abstract"]
        hit = False
        if "alignment" in attrs:
            absd["abstract_heading_alignment"] = attrs["alignment"]
            hit = True
        if "bold" in attrs:
            absd["abstract_heading_bold"] = bool(attrs["bold"])
            hit = True
        if "italic" in attrs:
            absd["abstract_heading_italic"] = bool(attrs["italic"])
            hit = True
        return hit

    if cat == "abstract_body":
        attrs = _parse_attrs(seg, font_size_hint)
        absd = out["abstract"]
        hit = False
        if "first_line_indent_in" in attrs:
            absd["abstract_first_line_indent_in"] = attrs["first_line_indent_in"]
            hit = True
        if "space_before_pt" in attrs:
            absd["abstract_space_before_pt"] = attrs["space_before_pt"]
            hit = True
        if "space_after_pt" in attrs:
            absd["abstract_space_after_pt"] = attrs["space_after_pt"]
            hit = True
        return hit

    if cat == "keywords":
        attrs = _parse_attrs(seg, font_size_hint)
        absd = out["abstract"]
        hit = False
        if attrs.get("italic") is True and re.search(r"\blabel\b", seg, re.I):
            absd["keywords_label_italic"] = True
            hit = True
        if "first_line_indent_in" in attrs:
            absd["keywords_indent_in"] = attrs["first_line_indent_in"]
            hit = True
        return hit

    if cat == "body":
        attrs = _parse_attrs(seg, font_size_hint)
        hit = False
        for k_src, k_dst in (("font_family", "font_family"), ("size_pt", "font_size_pt"),
                             ("line_spacing", "line_spacing"), ("alignment", "body_alignment"),
                             ("first_line_indent_in", "first_line_indent_in"),
                             ("space_before_pt", "space_before_pt"),
                             ("space_after_pt", "space_after_pt")):
            if k_src in attrs:
                page[k_dst] = attrs[k_src]
                hit = True
        return hit

    if cat == "margins":
        return _parse_margins(seg, page)

    if cat == "paper_size":
        return _parse_paper(seg, out["structural"])

    if cat == "columns":
        return _parse_columns(seg, out["structural"])

    if cat == "tables":
        return _parse_table_rules(seg, out["table"], out["structural"], page)

    if cat == "equations":
        return _parse_equation_rules(seg, out["equation"], page)

    if cat == "references":
        return _parse_reference(seg, out["reference"])

    if cat == "caption_table":
        return _parse_caption(seg, out["caption"], "table")

    if cat == "caption_figure":
        return _parse_caption(seg, out["caption"], "figure")

    if cat == "page_number":
        return _parse_page_number(seg, page)

    if cat in ("font", "line_spacing", "indent"):
        attrs = _parse_attrs(seg, font_size_hint)
        hit = False
        for k_src, k_dst in (("font_family", "font_family"), ("size_pt", "font_size_pt"),
                             ("line_spacing", "line_spacing"),
                             ("first_line_indent_in", "first_line_indent_in")):
            if k_src in attrs:
                page[k_dst] = attrs[k_src]
                hit = True
        return hit

    if cat == "toc":
        attrs = _parse_attrs(seg, font_size_hint)
        hit = False
        if "space_before_pt" in attrs:
            out["toc"]["toc_heading_space_before_pt"] = attrs["space_before_pt"]
            hit = True
        if "space_after_pt" in attrs:
            out["toc"]["toc_heading_space_after_pt"] = attrs["space_after_pt"]
            hit = True
        return hit

    return False


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
def parse_comments(comments, *, default_font_pt: float = 12.0) -> dict:
    """把英文模板批注解析成结构化规范（见模块 docstring 的输出契约）。"""
    out = {
        "page": {},
        "title_page": {},
        "abstract": {},
        "caption": {},
        "table": {},
        "equation": {},
        "toc": {},
        "structural": {},
        "heading_formats": {},
        "title_format": {},
        "reference": {},
        "count": 0,
        "unparsed": [],
        "notes": [],
    }
    for c in (comments or []):
        text = (c.get("text") or "").strip()
        if not text:
            continue
        out["count"] += 1
        matched = False
        for cat, seg in _split_parts(text):
            if cat == "unknown":
                continue
            if _apply_part(out, cat, seg, c):
                matched = True
        if not matched:
            out["unparsed"].append({"id": c.get("id"), "text": text})

    # 正文没写字号时，给「N characters」缩进折算留一个基准（不落到 page 里，避免臆造）
    if "font_size_pt" not in out["page"] and default_font_pt:
        pass

    # 结构维度：把空字典清掉，方便合并侧判断「批注是否提到」
    out["structural"] = {k: v for k, v in out["structural"].items() if v is not None}
    return out
