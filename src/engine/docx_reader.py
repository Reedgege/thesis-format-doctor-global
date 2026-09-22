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

import re
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
_PSEUDO_HEADING_MIN_PT = 13.0

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
    reference_hanging_indent_in: Optional[float] = None  # 参考文献实测悬挂缩进（英寸）


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


def is_pseudo_heading(para) -> bool:
    """伪标题判定：无标题样式/outline 标记，但形似标题（短 + 整段加粗 + 字号大于正文）。

    规格第 10 条要求保护「手写大字号标题」。判定刻意保守（整段显式加粗 +
    字号 ≥13pt + ≤80 字 + 不以句读结尾），宁可少改一段，也不能把主标题压成正文。
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
    sz = _max_run_size_pt(para)
    return sz is not None and sz >= _PSEUDO_HEADING_MIN_PT


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


def read_docx(path: str) -> DocumentProfile:
    """读取 docx → DocumentProfile。

    任何读取/解析异常（非 docx、损坏、加密、能打开但内部 XML 异常、受保护文档）
    一律转成 DocxReadError，保证 CLI/GUI 给用户友好提示而不是抛 traceback。
    """
    try:
        doc = Document(path)
    except Exception as e:  # 非 docx / 损坏 / 加密
        raise DocxReadError(
            f"无法读取 Word 文档：{e}（请确认是 .docx 文件、未损坏且未加密）"
        ) from e
    try:
        return _read_docx_impl(doc)
    except DocxReadError:
        raise
    except Exception as e:  # 能打开但内部结构异常 / 受保护
        raise DocxReadError(
            f"无法解析 Word 文档内容：{e}（文档可能受保护或结构损坏）"
        ) from e


def _read_docx_impl(doc) -> DocumentProfile:
    prof = DocumentProfile()

    # —— 页边距：遍历所有分节取众数；不一致则标记 ——
    per_side: dict = {s: [] for s in ("top", "bottom", "left", "right")}
    try:
        sections = list(doc.sections)
    except Exception:
        sections = []
    for sec in sections:
        per_side["top"].append(_inch(sec.top_margin))
        per_side["bottom"].append(_inch(sec.bottom_margin))
        per_side["left"].append(_inch(sec.left_margin))
        per_side["right"].append(_inch(sec.right_margin))

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
    prof.reference_text = "\n".join(prof.reference_entries)

    # 正文段落 = 非标题（含伪标题）、非参考文献段、非空
    skip = set(ref_indices)
    if ref_start is not None:
        skip.add(ref_start)
    body_paras = [p for i, p in enumerate(paras)
                  if i not in skip and p.text.strip() and not is_heading(p)]

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

    return prof


def reference_paragraph_indices(doc) -> list:
    """返回参考文献段落在 doc.paragraphs 中的下标列表。

    复用与 read_docx 完全一致的 find_reference_section（唯一实现），
    供 fixer 与测试精确区分「正文段落」与「参考文献段落」。
    """
    return find_reference_section(list(doc.paragraphs))[1]
