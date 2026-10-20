"""一键修正引擎：纯格式修正，绝不改动正文内容。

PaperFormat Pro —— 海外版「一键修正」核心。

设计铁律（与产品「只改格式、不改内容」原则一致）：
- 只改格式层：页边距 / 正文字体 / 字号 / 行距 / 正文对齐 / 首行缩进 / 段后间距 /
  参考文献悬挂缩进与顺序编码前缀。
- 绝不插入、删除、改写任何正文文本；正文段落文本修正前后逐字一致。
- 参考文献段落允许：① 设置悬挂缩进（纯格式，不碰文本）；
  ② 规范为顺序编码制(IEEE 等)且条目尚未带 [n] 时，补顺序前缀（结构标记，不删内容）。
  补前缀采用「在段落首位插入一个纯文本 run」的方式，**保留段落内原有的一切子元素**
  （内联图片 w:drawing、脚注/尾注引用、域代码 fldChar/instrText、分页符 w:br、
  超链接 w:hyperlink 等），不使用 Run.text setter（它会 clear_content() 清掉这些内容）。
- **标题段一律不碰**（字号/字体/行距/对齐/缩进/段距全部跳过），避免把标题压成正文 12pt。
- **目录/书签/交叉引用等「域段落」一律不碰**：这些段落通常是 Normal 样式而非标题，
  标题守卫拦不住；若对其逐 run 改字体/字号/对齐，会把自动生成目录改花、交叉引用错位。
  修正前先识别 w:fldChar/w:instrText/w:hyperlink/w:bookmarkStart 域再决定是否下手
  （国内版 Pit 2/3 的同类风险，海外版此前缺此守卫）。
- **不修改 Normal 样式的字体名**（否则 basedOn Normal 的标题/表格/页眉会间接被改）；
  字体只逐 run 设置，并同时写 rFonts 的 ascii/hAnsi/cs/eastAsia 四个属性。
- **首行缩进只作用于正文段落**（跳过标题与参考文献），避免破坏参考文献悬挂缩进。
- 永远写入新文件，绝不覆盖原件；输出路径按 realpath+normcase+samefile 判定，
  大小写/相对路径变体一律拒绝（见 same_file）。

注：本模块不含任何授权/网络逻辑，授权门禁由 CLI/GUI 在上层调用前裁决。
"""

from __future__ import annotations

import os
import re
from typing import Optional

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from .checker import (ALIGN_LABEL, run_check, Finding, FixAction,
                      _NUMBER_PREFIX_HINT)  # 对齐/编号提示文案单一来源（避免两处硬编码）
from .docx_reader import (
    DocxReadError, heading_level, is_heading, is_pseudo_heading, read_docx,
    reference_paragraph_indices, find_reference_section, find_abstract_section,
    pseudo_heading_level_map, PAPER_SIZES, describe_paper_size, detect_title_page,
    doc_title_index,
    _max_run_size_pt, _PSEUDO_HEADING_FORMAT_MIN_PT,
    detect_static_page_number, _para_has_page_field, _run_font_name,
    _para_text_excl_fields, is_caption_paragraph,
)
from .questionnaire import TargetProfile, PAGE_NUMBER_LAYOUT
from .specs import (MARGIN_DIMS, get_spec, resolve_spec, generic_heading_formats,
                    resolve_heading_formats, HeadingLevelFormat, TitlePageFormat,
                    size_for_font, spec_is_baseline)
from .comments import add_comment, strip_our_comments

_ALIGN_ENUM = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

_MARGIN_KEYS = MARGIN_DIMS     # (page 字段, 分节属性, 标签)

# 已带顺序编号的形态：[1] / (1) / 1. / 1)
_LEADING_NUMBER_RE = re.compile(r"^\s*(?:\[\s*\d+\s*\]|\(\s*\d+\s*\)|\d+\s*[.)])\s")


# ---------------------------------------------------------------------------
# 路径安全
# ---------------------------------------------------------------------------

def same_file(a: str, b: str) -> bool:
    """判断两个路径是否指向同一文件。

    覆盖：Windows 大小写不敏感（THESIS.DOCX vs thesis.docx）、
    相对/绝对混写（./paper.docx）、软链、以及已存在的同一 inode。
    """
    if not a or not b:
        return False
    try:
        pa = os.path.realpath(os.path.abspath(a))
        pb = os.path.realpath(os.path.abspath(b))
    except Exception:
        pa, pb = a, b
    if os.path.normcase(pa) == os.path.normcase(pb):
        return True
    try:
        return bool(os.path.exists(pa) and os.path.exists(pb) and os.path.samefile(pa, pb))
    except Exception:
        return False


def unique_output_path(path: str) -> str:
    """若目标文件已存在，追加 -2 / -3 … 序号，避免覆盖上一次的修正稿。"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 2
    while True:
        cand = f"{base}-{n}{ext}"
        if not os.path.exists(cand):
            return cand
        n += 1


# ---------------------------------------------------------------------------
# 段落选择（标题保护是铁律）
# ---------------------------------------------------------------------------

def _para_has_field(para) -> bool:
    """段落内是否含域元素（TOC/HYPERLINK/PAGEREF/REF/书签/交叉引用等）。

    这些段落是 Word 自动生成或结构页内容，格式应由其自身样式 / 域定义决定，
    一键修正中必须跳过，否则会把自动目录、交叉引用、书签改花
    （国内版 Pit 2/3 的同类风险：域段落被当正文改）。
    """
    p = getattr(para, "_p", None)
    if p is None:
        return False
    for tag in ("w:fldChar", "w:instrText", "w:hyperlink", "w:bookmarkStart",
                "w:footnoteReference", "w:endnoteReference"):
        if p.find(qn(tag)) is not None:
            return True
    for run in para.runs:
        r = getattr(run, "_r", None)
        if r is None:
            continue
        if (r.find(qn("w:fldChar")) is not None
                or r.find(qn("w:instrText")) is not None
                or r.find(qn("w:hyperlink")) is not None):
            return True
    return False


def _select_paragraphs(doc, ref_set: set, include_reference: bool = True,
                       heading_ps: set | None = None) -> list:
    """选取可被格式修正的段落：一律跳过标题段与域段落；可选跳过参考文献段。

    域段落（TOC/书签/交叉引用）通常不是标题样式，标题守卫拦不住，故单独排除；
    但参考文献条目本身可能含 URL 超链接（w:hyperlink），仍要参与格式修正，
    因此仅当该段不在 ref_set 时才跳过域段落。

    ``heading_ps``：已判定为标题的 w:p 元素集合（由 _assign_heading_levels 产出）。
    必须与 _apply_heading_format 用**同一份**判定 —— 否则「正文档」会把标题段
    当正文刷（首行缩进等），标题流程再改回来，两边每遍互相打架：
    明明没变却恒报「已改 N 处」，且标题被正文格式污染（实测 12pt 加粗标题
    因 is_heading 的 13pt 阈值漏判而被 justify + 首行缩进污染）。
    """
    out: list = []
    for i, para in enumerate(doc.paragraphs):
        if heading_ps is not None:
            if para._p in heading_ps:
                continue
        elif is_heading(para):
            continue
        if _para_has_field(para) and i not in ref_set:
            continue
        if not include_reference and i in ref_set:
            continue
        out.append(para)
    return out


def _abstract_body_indices(doc) -> set:
    """摘要正文段索引：从 'Abstract'/'摘要' 标题之后，到下一个「真实标题」之前。

    摘要正文首行必须顶格（不缩进），与正文首行缩进 0.5″ 相反，故需从首行缩进目标中剔除。
    仅认「样式级标题」（heading_level 非 None）作为摘要标题与结束边界，避免把加粗正文误判为边界。
    """
    idxs: set = set()
    paras = doc.paragraphs
    in_abs = False
    for i, para in enumerate(paras):
        if not in_abs:
            if (para.text or "").strip().lower() in ("abstract", "摘要") and heading_level(para) is not None:
                in_abs = True
                continue
        else:
            if heading_level(para) is not None:
                break
            if (para.text or "").strip():
                idxs.add(i)
    return idxs


def _abstract_section(doc) -> tuple[Optional[int], list]:
    """摘要段定位（复用 reader 的单一真相源）：返回 (标题段下标, 正文段下标列表)。

    找不到摘要标题时返回 (None, [])。checker 与 fixer 共用同一识别逻辑，
    避免「检查判这段、修正改另一段」的漂移。
    """
    return find_abstract_section(list(doc.paragraphs))


def _apply_abstract_indent(doc, indent_in: float) -> bool:
    """设置摘要正文段的首行缩进（APA 摘要顶格 = 0）。

    只动摘要正文段（标题之后、下一标题之前的段落），不碰摘要标题与其它正文。
    摘要正文段本来就在正文首行缩进的豁免集合里，这里再正向设置目标缩进值。
    """
    try:
        want = float(indent_in)
    except (ValueError, TypeError):
        return False
    _head, body_idx = _abstract_section(doc)
    if not body_idx:
        return False
    changed = False
    for j in body_idx:
        if 0 <= j < len(doc.paragraphs):
            pf = doc.paragraphs[j].paragraph_format
            cur = pf.first_line_indent
            if cur is None or abs(float(cur.inches) - want) > 1e-3:
                try:
                    pf.first_line_indent = Inches(want)
                    changed = True
                except Exception:
                    pass
    return changed


def _apply_abstract_new_page(doc, new_page: bool) -> bool:
    """摘要标题是否另起一页（段前分页 pageBreakBefore）。

    True = 摘要标题前强制分页；False = 明确不另起（清段前分页，不动上一段末尾的分页符）。
    只动摘要标题段的 paragraph_format.page_break_before，幂等。
    """
    _head, _body = _abstract_section(doc)
    if _head is None:
        return False
    para = doc.paragraphs[_head]
    pf = para.paragraph_format
    try:
        cur = bool(pf.page_break_before)
    except Exception:
        cur = False
    want = bool(new_page)
    if cur == want:
        return False
    try:
        pf.page_break_before = want
        return True
    except Exception:
        return False


def _caption_paragraphs(doc) -> list:
    """收集图表标题（Caption）段落：复用 reader 的单一真相源识别。"""
    out = []
    for p in doc.paragraphs:
        if is_caption_paragraph(p):
            out.append(p)
    return out


def _apply_caption_size(doc, size_pt: float) -> bool:
    """设置图表标题字号（逐 run，不碰字体）。识别 Caption 段落。"""
    try:
        want = Pt(float(size_pt))
    except (ValueError, TypeError):
        return False
    caps = _caption_paragraphs(doc)
    if not caps:
        return False
    changed = False
    for para in caps:
        for run in para.runs:
            if _is_comment_run(run):
                continue
            try:
                if run.font.size != want:
                    run.font.size = want
                    changed = True
            except Exception:
                pass
    return changed


def _apply_caption_spacing(doc, spacing: float) -> bool:
    """设置图表标题行距（倍数制）。识别 Caption 段落。"""
    try:
        sp = float(spacing)
    except (ValueError, TypeError):
        return False
    caps = _caption_paragraphs(doc)
    if not caps:
        return False
    changed = False
    for para in caps:
        cur = para.paragraph_format.line_spacing
        try:
            same = cur is not None and abs(float(cur) - sp) <= 1e-6
        except (TypeError, ValueError):
            same = False
        if not same:
            para.paragraph_format.line_spacing = sp
            changed = True
    return changed


def _is_block_quote(para) -> bool:
    """块引用判定：样式名含 'quote'（Word 内置 Quote / Intense Quote）。"""
    try:
        name = (para.style.name or "").lower() if para.style else ""
    except Exception:
        return False
    return "quote" in name


def _apply_block_quote_indent(paras: list) -> bool:
    """块引用（APA ≥40 词长引用）：左缩进 0.5″、首行不缩进。仅动格式，不动文本。

    与正文首行缩进互斥——若不加此豁免，块引用会被正文流程误加 0.5″ 首行缩进且左缩进清零。
    """
    changed = False
    for para in paras:
        pf = para.paragraph_format
        if pf.left_indent != Inches(0.5):
            pf.left_indent = Inches(0.5)
            changed = True
        if pf.first_line_indent is None or pf.first_line_indent != Inches(0):
            pf.first_line_indent = Inches(0)
            changed = True
    return changed


# ---------------------------------------------------------------------------
# 单维度修正基元（全部只动格式属性，不动文本）
# ---------------------------------------------------------------------------

def _is_comment_run(run) -> bool:
    """批注标记 run（含 w:commentReference）不是正文，绝不能当正文套格式。

    不跳过的话：上一遍加的批注标记 run 会在下一遍被当成正文重新设字体/字号 ——
    既污染我们自己的标记，又让「是否真改了」恒为真（明明没改动却仍声称 font set，
    汇总批注于是满嘴假话）。国内版 Pit 已踩过同类坑，这里照抄其结论。
    """
    try:
        return run._r.find(qn("w:commentReference")) is not None
    except Exception:
        return False


def _set_run_font(run, font: Optional[str], size: Optional[float]) -> bool:
    """给单个 run 设字体/字号，**如实返回是否真改了**。

    字体同时写 w:rFonts 的 ascii/hAnsi/cs/eastAsia —— run.font.name 只映射
    ascii/hAnsi，只设它会让「仅设中文字体」的文档看起来没改到。
    """
    changed = False
    if font:
        try:
            if run.font.name != font:
                run.font.name = font
                changed = True
        except Exception:
            pass
        try:
            rPr = run._r.get_or_add_rPr()
            rFonts = rPr.find(qn("w:rFonts"))
            if rFonts is None:
                rFonts = OxmlElement("w:rFonts")
                rPr.insert(0, rFonts)
                changed = True
            for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                if rFonts.get(qn(attr)) != font:
                    rFonts.set(qn(attr), font)
                    changed = True
        except Exception:
            pass
    if size is not None:
        try:
            want = Pt(float(size))
            if run.font.size != want:
                run.font.size = want
                changed = True
        except Exception:
            pass
    return changed


def _apply_margins(doc, page: dict):
    """按 TargetProfile.page 设置各分节页边距。

    MARGIN_DIMS 的第二个元素是「侧面名」(top/bottom/left/right)：读取侧对应
    prof.margins 的键，写入侧对应 docx 分节的 `<side>_margin` 属性。

    如实返回「是否真改了」：原来无返回值，汇总批注于是无条件写 "margins set to X"，
    对本来页边距就合规的文档也会凭空冒出一条"我改了页边距"的批注。
    """
    changed = False
    for pk, side, _label in _MARGIN_KEYS:
        if page.get(pk) is None:
            continue
        val = Inches(float(page[pk]))
        attr = f"{side}_margin"
        for sec in doc.sections:
            if getattr(sec, attr, None) != val:
                setattr(sec, attr, val)
                changed = True
    return changed


def _apply_paper_size(doc, target: TargetProfile) -> bool:
    """按 ``target.paper_size`` 设置各分节的纸张大小（``sectPr/w:pgSz``）。

    🔴 **通用维度**：Letter（北美 8.5×11 in）与 A4（欧洲/中国 210×297 mm）不是某个
    引用规范的一部分。客户选「A4」就整篇改成 A4，与选 APA 还是 MLA 无关。

    实现用 python-docx 的 ``Section.page_width / page_height``（内部写 ``w:pgSz/@w:w/@w:h``），
    单位 EMU。**保留原有横向/纵向**：原文档宽 > 高 视为横向，改尺寸后仍保持横向，
    否则一份横向 A4 会被悄悄转成纵向（版面全乱）。
    幂等：已是目标尺寸则不改、返回 False。
    """
    name = getattr(target, "paper_size", None)
    if not name:
        return False
    dims = PAPER_SIZES.get(str(name))
    if not dims:
        return False
    short_in, long_in = dims
    changed = False
    for sec in doc.sections:
        try:
            cur_w, cur_h = sec.page_width, sec.page_height
        except Exception:
            continue
        landscape = False
        if cur_w is not None and cur_h is not None and cur_w > cur_h:
            landscape = True
        new_w, new_h = (long_in, short_in) if landscape else (short_in, long_in)
        try:
            want_w, want_h = Inches(new_w), Inches(new_h)
            if cur_w is not None and cur_h is not None \
                    and abs(int(cur_w) - int(want_w)) < 1000 \
                    and abs(int(cur_h) - int(want_h)) < 1000:
                continue          # 已达标（1000 EMU ≈ 0.001 in 容差）
            sec.page_width = want_w
            sec.page_height = want_h
            changed = True
        except Exception:
            continue
    return changed


# ---------------------------------------------------------------------------
# 字体白名单的「字体 + 字号」配对修正（v2.3.66）
# ---------------------------------------------------------------------------

def _apply_size_by_font_family(paras: list, pairs: dict) -> tuple:
    """按**各 run 自己的字体**把字号校正到该字体的规范值；只动 ``w:sz``，不动 ``w:rFonts``。

    🔴 为什么不统一设一个字号：APA 允许的 6 款字体各有自己的合规字号（TNR 12 / Calibri 11 /
    Arial 11 / Georgia 11 / Lucida Sans Unicode 10 / Computer Modern 10）。一篇正文里
    混用 TNR 与 Calibri 时，统一设成 12pt 必然把 Calibri 那一半改错。逐 run 判定才对。
    🔴 为什么不换字体：6 款都合规，替客户换成哪一款都是越权；而且 ``_set_run_font`` 会连
    ``w:eastAsia`` 一起写 —— 中文文档的「宋体 + Times New Roman」混排会被一并破坏。
    白名单之外的字体一律跳过（绝不猜字号）。

    返回 ``(改动的 run 数, {字体: 目标字号})``。
    """
    changed = 0
    detail: dict = {}
    if not pairs:
        return 0, {}
    for para in paras:
        for run in para.runs:
            try:
                if _is_comment_run(run):
                    continue
            except Exception:
                pass
            if not (run.text or "").strip():
                continue
            fam = _run_font_name(run)
            want = size_for_font(fam, pairs)
            if not want:
                continue
            try:
                cur = run.font.size
                if cur is not None and abs(float(cur.pt) - float(want)) <= 0.01:
                    continue
                run.font.size = Pt(float(want))
                changed += 1
                detail[str(fam)] = float(want)
            except Exception:
                continue
    return changed, detail


# ---------------------------------------------------------------------------
# 页码位置（v2.3.66 新能力）：fixer 第一次操作页眉 / 页脚 part
# ---------------------------------------------------------------------------

def _write_page_field(para):
    """在段落末尾追加一个标准 PAGE 复杂域，返回承载该域的 run。

    ``fldChar(begin) → instrText(" PAGE ") → fldChar(separate) → 缓存文本 → fldChar(end)``
    是 Word 自己保存时用的写法，Office / WPS / LibreOffice 都认。
    python-docx 没有域写入的高层 API，只能手写 OOXML（但结构完全标准）。
    """
    run = para.add_run()
    r = run._r
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE  \\* MERGEFORMAT "
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t")
    txt.text = "1"                     # 缓存结果：未更新域时 Word 显示这个
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr, sep, txt, end):
        r.append(el)
    return run


def _restore_run_fmt(run, name, size, bold) -> None:
    """把「替换前那个 run」的字体/字号/加粗套回新 run（不覆盖已有值）。"""
    try:
        if name:
            run.font.name = name
            rPr = run._r.get_or_add_rPr()
            rFonts = rPr.find(qn("w:rFonts"))
            if rFonts is not None:
                for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                    if rFonts.get(qn(attr)) is None:
                        rFonts.set(qn(attr), name)
        if size is not None:
            run.font.size = size
        if bold is not None:
            run.font.bold = bold
    except Exception:
        pass


def _set_para_align(para, align: str) -> bool:
    """设置段落对齐，**如实返回是否真改了**（幂等：已对就不动）。"""
    enum = _ALIGN_ENUM.get(str(align))
    if enum is None:
        return False
    try:
        if para.alignment == enum:
            return False
        para.alignment = enum
        return True
    except Exception:
        return False


def _replace_static_page_number(para) -> bool:
    """策略 C：把客户手打的页码**原地**换成 PAGE 域，前后缀与字符格式原样保留。

    「第 1 页」→「第 {PAGE} 页」；「Page 1」→「Page {PAGE}」；「- 1 -」→「- {PAGE} -」。
    这样既不会出现两个页码，也不会把客户的「第…页」格式丢掉。
    整段只有一个数字才处理（"Page 1 of 10" 不动 —— 总页数得换成 NUMPAGES，超出保守范围）。
    """
    info = detect_static_page_number(para.text)
    if not info:
        return False
    pre = info.get("prefix") or ""
    suf = info.get("suffix") or ""
    font_name = font_size = font_bold = None
    for r in para.runs:
        if (r.text or "").strip():
            try:
                font_name = r.font.name
                font_size = r.font.size
                font_bold = r.font.bold
            except Exception:
                pass
            break
    # 只清 run 元素，保留 pPr（段落对齐 / 间距 / 制表位一概不动）
    for r in list(para.runs):
        try:
            r._r.getparent().remove(r._r)
        except Exception:
            pass
    if pre:
        _restore_run_fmt(para.add_run(pre), font_name, font_size, font_bold)
    _restore_run_fmt(_write_page_field(para), font_name, font_size, font_bold)
    if suf:
        _restore_run_fmt(para.add_run(suf), font_name, font_size, font_bold)
    return True


def _ensure_page_number_in_part(part, align: str) -> tuple:
    """确保某个页眉/页脚 part 里有页码且对齐正确。

    返回 ``(动作, 是否改动)``，动作为 ``"kept"`` / ``"replaced"`` / ``"added"`` / ``""``。

    没有页码时怎么写：页眉里若已有非空文本（running head 短标题）→ **另起一段**放页码，
    绝不把标题挤到右边；页眉/页脚为空 → 复用第一个空段落。
    """
    try:
        paras = list(part.paragraphs)
    except Exception:
        return "", False
    # ① 已有 PAGE 域：只校对齐
    for p in paras:
        if _para_has_page_field(p):
            return "kept", _set_para_align(p, align)
    # ② 手打静态页码：原地替换 + 校对齐
    for p in paras:
        if detect_static_page_number(p.text):
            _replace_static_page_number(p)
            _set_para_align(p, align)
            return "replaced", True
    # ③ 完全没有：挑一个空白段（没有就新建）
    target = None
    for p in paras:
        if not (p.text or "").strip():
            target = p
            break
    if target is None:
        try:
            target = part.add_paragraph()
        except Exception:
            return "", False
    _write_page_field(target)
    _set_para_align(target, align)
    return "added", True


def _strip_page_numbers_in_part(part) -> int:
    """删掉某个页眉/页脚里已有的页码（自动域或整段手打数字），返回清理的段数。

    🔴 为什么必须做：客户原来页码在页脚、要求改成页眉右上时，若只往页眉加一条，
    页脚那条还在 —— 页面上会**同时出现两个页码**。
    只删「页码本身」：段落里若还有别的文字（如页脚写着 "Confidential" 另起一段），
    那些段落一个字节都不动；同一段里除页码外还有实质文字时，只删承载域/数字的 run。
    """
    removed = 0
    try:
        paras = list(part.paragraphs)
    except Exception:
        return 0
    for p in paras:
        txt = p.text or ""
        is_static = detect_static_page_number(txt) is not None
        has_field = _para_has_page_field(p)
        if not (is_static or has_field):
            continue
        if is_static:
            # 整段就是手打页码（「第 1 页」）→ 清空该段（前后缀一起去掉，否则留下「第  页」）
            for r in list(p.runs):
                try:
                    r._r.getparent().remove(r._r)
                except Exception:
                    pass
            removed += 1
            continue
        # 含 PAGE 域：看这段除了页码还有没有别的实质文字
        reside = _para_text_excl_fields(p)
        reside = re.sub(r"(?i)\bpage\b|第|页", "", reside)
        reside = re.sub(r"[\s\-–—./:：]+", "", reside)
        if not reside:
            for r in list(p.runs):
                try:
                    r._r.getparent().remove(r._r)
                except Exception:
                    pass
        else:
            for r in list(p.runs):
                try:
                    if r._r.find(qn("w:fldChar")) is not None \
                            or r._r.find(qn("w:instrText")) is not None:
                        r._r.getparent().remove(r._r)
                except Exception:
                    pass
        removed += 1
    return removed


def _apply_page_number(doc, position: str) -> dict:
    """按 ``position`` 写入自动页码：top-right=页眉右上 / bottom-center / bottom-right=页脚。

    三条硬约束（都是实测踩过的坑）：
    1. **原地替换**手打页码，绝不叠加第二条（否则客户页面上出现两个页码）。
    2. **多分节按链接状态处理**：``is_linked_to_previous=True`` 的节继承前节定义，不重复写；
       取消链接的节必须单独写，否则那几页没有页码。
    3. 首页不同（``different_first_page_header_footer``）时首页页眉/页脚是独立 part，
       必须一起处理，否则标题页没有页码。

    返回空 dict 表示一个字没改（幂等），非空表示真改了（供 ``_compute_changes`` 预览）。
    """
    kind, align = PAGE_NUMBER_LAYOUT.get(str(position), ("header", "right"))
    stats = {"location": kind, "alignment": align,
             "added": 0, "replaced": 0, "aligned_only": 0, "removed_elsewhere": 0}
    changed = False
    try:
        sections = list(doc.sections)
    except Exception:
        sections = []
    even_on = False
    try:
        even_on = doc.settings.element.find(qn("w:evenAndOddHeaders")) is not None
    except Exception:
        even_on = False
    other_kind = "footer" if kind == "header" else "header"

    def _parts_of(sec, k, i):
        """某分节某类（页眉/页脚）需要处理的 part 列表（尊重「链接到前节」）。"""
        out = []
        main = getattr(sec, k, None)
        if main is not None:
            linked = False
            if i > 0:
                try:
                    linked = bool(main.is_linked_to_previous)
                except Exception:
                    linked = False
            if not linked:
                out.append(main)
        try:
            diff = bool(sec.different_first_page_header_footer)
        except Exception:
            diff = False
        if diff:
            fp = getattr(sec, "first_page_" + k, None)
            if fp is not None:
                out.append(fp)
        if even_on:
            ev = getattr(sec, "even_page_" + k, None)
            if ev is not None:
                out.append(ev)
        return out

    # ① 先确保目标位置有页码（先写、后清 —— 万一写入失败，客户原来的页码不会先被删掉）
    for i, sec in enumerate(sections):
        for part in _parts_of(sec, kind, i):
            act, did = _ensure_page_number_in_part(part, align)
            if act == "added" and did:
                stats["added"] += 1
                changed = True
            elif act == "replaced" and did:
                stats["replaced"] += 1
                changed = True
            elif act == "kept" and did:
                stats["aligned_only"] += 1
                changed = True
    # ② 再清掉另一处（页眉↔页脚）的页码，否则页面上会出现两个页码
    for i, sec in enumerate(sections):
        for part in _parts_of(sec, other_kind, i):
            n = _strip_page_numbers_in_part(part)
            if n:
                stats["removed_elsewhere"] += n
                changed = True
    return stats if changed else {}


def _apply_columns(doc, target: TargetProfile) -> bool:
    """按 ``target.columns`` 设置各分节栏数（``sectPr/w:cols/@w:num``）。

    python-docx 1.2.0 的 ``Section`` 没有栏数属性，只能直写 sectPr XML：
    - 双栏：``w:num="2"``，并确保 ``w:space`` 有值（默认 425 twips ≈ 0.3" 两栏间距）；
    - 单栏：移除 ``w:num``（Word 默认即 1 栏），不动已有 ``w:space``。
    幂等：已是目标栏数则不改、返回 False —— 避免凭空声称「我改了栏数」。
    """
    n = getattr(target, "columns", None)
    if n is None:
        return False
    try:
        n = max(1, int(n))
    except (ValueError, TypeError):
        return False
    changed = False
    for sec in doc.sections:
        try:
            sectPr = sec._sectPr
        except Exception:
            continue
        cols = sectPr.find(qn("w:cols"))
        if cols is None:
            cols = OxmlElement("w:cols")
            # schema 顺序：cols 应位于 docGrid 之前、pgMar 之后（见 probe_cols 实测）
            dg = sectPr.find(qn("w:docGrid"))
            if dg is not None:
                dg.addprevious(cols)
            else:
                sectPr.append(cols)
        cur = cols.get(qn("w:num"))
        try:
            cur_n = int(cur) if cur is not None else 1
        except (ValueError, TypeError):
            cur_n = 1
        if cur_n != n:
            if n == 1:
                cols.attrib.pop(qn("w:num"), None)   # 单栏=Word 默认，去掉 num 更干净
            else:
                cols.set(qn("w:num"), str(n))
            changed = True
        if n >= 2 and cols.get(qn("w:space")) is None:
            cols.set(qn("w:space"), "425")           # 两栏间距 ≈ 0.3"
            changed = True
    return changed


def _apply_header_distance(doc, page: dict) -> bool:
    """按 ``page["header_distance_in"]`` 设置各分节页眉距页顶（英寸）。

    python-docx 的 ``Section.header_distance`` 是可直接写的 Length 属性，故这不是
    「只提示」项——可以真正落盘修正（与页边距同类）。幂等：已等于目标值则返回 False。
    """
    val = page.get("header_distance_in")
    if val is None:
        return False
    try:
        want = Inches(float(val))
    except (ValueError, TypeError):
        return False
    changed = False
    for sec in doc.sections:
        try:
            if getattr(sec, "header_distance", None) != want:
                sec.header_distance = want
                changed = True
        except Exception:
            continue
    return changed


def _apply_font_and_size(page: dict, paras: list) -> bool:
    """逐 run 设置正文字体/字号。

    **不改 Normal 样式**：Word 中 Heading/Title/表格/页眉等大量样式 basedOn Normal
    且不显式设字体，改 Normal 会把这些内容的字体一起改掉，违反「标题字体不得改动」。
    """
    font = page.get("font_family")
    size = page.get("font_size_pt")
    if font is None and size is None:
        return False
    changed = False
    for para in paras:
        for run in para.runs:
            if _is_comment_run(run):
                continue                      # 批注标记不是正文，跳过
            if _set_run_font(run, font, size):
                changed = True
    return changed


def _apply_line_spacing(page: dict, paras: list) -> bool:
    """设置行距（倍数制）。

    只设 line_spacing 数值，由 python-docx 自动归到对应规则
    （1.0/1.5/2.0 为命名规则，其余为 MULTIPLE），均为倍数制，非固定值/最小值。
    """
    if page.get("line_spacing") is None:
        return False
    sp = float(page["line_spacing"])
    # 如实返回「是否真改了」：原来恒返回 True，导致汇总批注即便一个字没动也声称
    # "line spacing set to X" —— 客户按批注去找改动会扑空（违背「让客户知道改了哪儿」）。
    changed = False
    for para in paras:
        cur = para.paragraph_format.line_spacing
        try:
            same = cur is not None and abs(float(cur) - sp) <= 1e-6
        except (TypeError, ValueError):
            same = False        # 固定值/最小值行距（Length）一律视为需要归一化为倍数制
        if not same:
            para.paragraph_format.line_spacing = sp
            changed = True
    return changed


def _apply_alignment(page: dict, paras: list) -> bool:
    wanted = page.get("body_alignment")
    if not wanted:
        return False
    enum = _ALIGN_ENUM.get(str(wanted))
    if enum is None:
        return False
    changed = False
    for para in paras:
        if para.paragraph_format.alignment != enum:
            para.paragraph_format.alignment = enum
            changed = True
    return changed


def _apply_first_line_indent(page: dict, paras: list) -> bool:
    """正文段落首行缩进（只作用于正文段落，不含标题与参考文献）。

    如实返回「是否真改了」（原来无返回值，调用方只能无条件记进 applied → 假报改动）。
    """
    val = page.get("first_line_indent_in")
    if val is None:
        return False
    inches = Inches(float(val))
    changed = False
    for para in paras:
        pf = para.paragraph_format
        if pf.first_line_indent != inches:
            pf.first_line_indent = inches
            changed = True
        # 首行缩进要求整体不缩进（悬挂由参考文献单独处理）
        if pf.left_indent != Inches(0):
            pf.left_indent = Inches(0)
            changed = True
    return changed


def _apply_space_after(page: dict, paras: list) -> bool:
    """段后间距。如实返回「是否真改了」（原来无返回值 → 汇总批注假报）。"""
    val = page.get("space_after_pt")
    if val is None:
        return False
    pt = Pt(float(val))
    changed = False
    for para in paras:
        if para.paragraph_format.space_after != pt:
            para.paragraph_format.space_after = pt
            changed = True
    return changed


# ---------------------------------------------------------------------------
# 参考文献：格式层（悬挂缩进 + 顺序编码前缀）
# ---------------------------------------------------------------------------

def _has_leading_number(text: str) -> bool:
    return bool(_LEADING_NUMBER_RE.match(text or ""))


def _existing_numbers(entries: list) -> set:
    nums: set = set()
    for para in entries:
        m = re.match(r"^\s*\[\s*(\d+)\s*\]", para.text or "")
        if m:
            nums.add(int(m.group(1)))
    return nums


def _inherit_run_format(new_run, para) -> None:
    """让 new_run 复制段内第一个非批注 run 的字符格式（字体/字号/加粗/斜体）。

    找不到源 run 或复制失败都静默忽略 —— 这只是「锦上添花」的格式对齐，
    绝不能因为它把整条参考文献编号流程搞失败。
    """
    try:
        src = None
        for r in para.runs:
            if r is new_run or _is_comment_run(r):
                continue
            if (r.text or "").strip():
                src = r
                break
        if src is None:
            return
        rPr_t = src._r.find(qn("w:rPr"))
        if rPr_t is not None:
            import copy as _copy
            rPr_new = new_run._r.get_or_add_rPr()
            for child in list(rPr_new):
                rPr_new.remove(child)
            for child in rPr_t:
                rPr_new.append(_copy.deepcopy(child))
    except Exception:
        pass


def _insert_prefix_run(para, prefix: str) -> bool:
    """在段落**首位**插入一个纯文本 run，保留段落内原有的一切子元素。

    实现：先 add_run（追加到段末）拿到 w:r 元素，再把它搬到 w:pPr 之后的第一个位置。
    这样段落里原有的 w:drawing / w:footnoteReference / w:fldChar / w:br(type=page) /
    w:hyperlink 等都不受影响（Run.text setter 会 clear_content() 把它们删掉，禁用）。

    新 run 会**继承后继 run 的字体/字号**：编号前缀是在正文档字体流程之后才插入的
    （fix_docx 里 page 刷字体在前、参考文献编号在后），若不继承，前缀 run 就是无字体
    的孤儿 —— 表现为「第一遍跑完前缀还是默认字体、第二遍才变」，等于让用户跑两遍。
    """
    try:
        new_run = para.add_run(prefix)
    except Exception:
        return False
    _inherit_run_format(new_run, para)
    el = new_run._r
    p = para._p
    try:
        p.remove(el)
    except Exception:
        return True       # 搬不动就保持追加位置，至少不丢内容
    try:
        pPr = p.find(qn("w:pPr"))
        if pPr is not None:
            pPr.addnext(el)        # w:pPr 必须是 w:p 的第一个子元素
        else:
            p.insert(0, el)
    except Exception:
        try:
            p.append(el)
        except Exception:
            return False
    return True


def _apply_reference_numbering(entries: list) -> int:
    """给尚无顺序编号的条目补 [n] 前缀，返回实际补了几条。

    - 全部无编号 → 顺序 1..N；
    - 混合编号（部分条目自带 [n]）→ 只补缺号的条目，使用未被占用的序号，避免重号。
    """
    used = _existing_numbers(entries)
    added = 0
    if not used:
        n = 1
        for para in entries:
            if not (para.text or "").strip():
                continue
            if _has_leading_number(para.text):
                continue
            if _insert_prefix_run(para, f"[{n}] "):
                added += 1
            n += 1
        return added

    candidate = 1
    for para in entries:
        if not (para.text or "").strip() or _has_leading_number(para.text):
            continue
        while candidate in used:
            candidate += 1
        if _insert_prefix_run(para, f"[{candidate}] "):
            added += 1
        used.add(candidate)
        candidate += 1
    return added


def _apply_reference_format(doc, target: TargetProfile, ref_indices: list,
                            fix_hang: bool = True, fix_number: bool = True,
                            fix_font: bool = False, fix_size: bool = False,
                            fix_spacing: bool = False, fix_new_page: bool = False,
                            fix_heading: bool = False, fix_rewrite: bool = False) -> dict:
    """参考文献：格式层 + （IEEE 时）best-effort 引文格式重写。

    - 悬挂缩进 + （编号制时）补顺序前缀；
    - v2.3.65：条目字体 / 字号 / 行距、参考文献标题另起一页、标题加粗 / 居中
      —— 这五项都是学校模板里常见的硬要求，以前 target 承载不了、也无从修正；
    - 当且仅当 reference_style == "IEEE" 且条目明显为 APA 作者-年格式、且为单 run 纯文本
      （无超链接/域，避免破坏内联 URL）时，best-effort 改写为 IEEE 编号引文
      （作者首字母缩写 + 题名引号 + 期刊 vol./no./pp. + 年）。幂等、保守、可回归；
      非 APA 模式或无法安全解析的条目一律不动其文本。
    """
    info = {"hanging_applied": False, "numbering_added": 0, "rewritten": 0,
            "entry_reformatted": 0, "heading_formatted": 0, "new_page_applied": False}
    if not ref_indices:
        return info
    entries = [doc.paragraphs[i] for i in ref_indices]

    # 引文内容重写必须**显式开启**（target.reference_rewrite），且**由检查结论门控**
    # （v2.3.70 新增 fix_rewrite）——不能只看 reference_style == "IEEE" 就改：
    # 产品铁律是「只改格式、不改内容」，而把 APA 条目改写成 IEEE 引文是内容层改动。
    # 检查器没报 reference 问题（条目本已合规）时，这里一个字都不许动。
    # 默认关闭 = 参考文献逐字不动（只加 [n] 前缀 + 悬挂缩进），这正是
    # tests/test_review_fixes.py P0-1 断言的不变量；只有走 target_from_spec("IEEE")
    # 的真实产品路径、且检查真的报了问题，才会打开。
    if target.reference_style == "IEEE" and getattr(target, "reference_rewrite", False) \
            and fix_rewrite:
        info["rewritten"] = _reformat_ieee_reference_entries(doc, ref_indices)

    hang = target.reference_hanging_indent_in
    if hang is not None and fix_hang:
        h = float(hang)
        for para in entries:
            pf = para.paragraph_format
            pf.left_indent = Inches(h)
            pf.first_line_indent = Inches(-h)
        info["hanging_applied"] = True

    if target.reference_numbered and fix_number:
        info["numbering_added"] = _apply_reference_numbering(entries)

    # —— 条目字体 / 字号 / 行距（学校模板常与正文不同口径：正文小四双倍、条目五号单倍）——
    want_font = target.reference_font_family if fix_font else None
    want_size = target.reference_size_pt if fix_size else None
    want_sp = target.reference_line_spacing if fix_spacing else None
    if want_font or want_size or want_sp:
        for para in entries:
            if _format_reference_entry(para, want_font, want_size, want_sp):
                info["entry_reformatted"] += 1

    # —— 参考文献标题段：另起一页 / 加粗 / 居中 ——
    if fix_new_page or fix_heading:
        ref_start, _ = find_reference_section(list(doc.paragraphs))
        if ref_start is not None:
            head = doc.paragraphs[ref_start]
            if fix_new_page and target.reference_new_page is not None:
                want_np = bool(target.reference_new_page)
                pf_head = head.paragraph_format
                if want_np and not pf_head.page_break_before:
                    pf_head.page_break_before = True
                    info["new_page_applied"] = True
                elif not want_np:
                    if pf_head.page_break_before:
                        pf_head.page_break_before = False
                        info["new_page_applied"] = True
                    # 手工插的显式分页符也要撤掉，否则复检仍判「另起一页」
                    for run in list(head.runs):
                        if _is_page_break_only(run):
                            head._p.remove(run._r)
                            info["new_page_applied"] = True
            if fix_heading:
                changed = False
                if target.reference_heading_bold is not None:
                    for run in head.runs:
                        if _apply_tri_state(head, run, "bold",
                                            target.reference_heading_bold):
                            changed = True
                if target.reference_heading_centered is not None:
                    want_al = (WD_ALIGN_PARAGRAPH.CENTER
                               if target.reference_heading_centered
                               else WD_ALIGN_PARAGRAPH.LEFT)
                    if head.alignment != want_al:
                        head.alignment = want_al
                        changed = True
                if changed:
                    info["heading_formatted"] = 1
    return info


def _is_page_break_only(run) -> bool:
    """该 run 是否只含一个分页符（没有可见文字）。"""
    if (run.text or "").strip():
        return False
    try:
        for br in run._r.findall(qn("w:br")):
            if br.get(qn("w:type")) == "page":
                return True
    except Exception:
        return False
    return False


def _format_reference_entry(para, font=None, size_pt=None, line_spacing=None) -> bool:
    """把一参考文献条目的字体 / 字号 / 行距落到直接格式上，返回是否有改动（幂等）。"""
    changed = False
    pf = para.paragraph_format
    if line_spacing is not None:
        cur = pf.line_spacing
        try:
            same = cur is not None and abs(float(cur) - float(line_spacing)) <= 1e-6
        except (TypeError, ValueError):
            same = False
        if not same:
            try:
                pf.line_spacing = float(line_spacing)
                changed = True
            except Exception:
                pass
    for run in para.runs:
        if _is_comment_run(run):
            continue
        if font:
            if _apply_run_font(run, font):
                changed = True
        if size_pt:
            cur_sz = run.font.size
            if cur_sz is None or abs(float(cur_sz.pt) - float(size_pt)) > 1e-6:
                try:
                    run.font.size = Pt(size_pt)
                    changed = True
                except Exception:
                    pass
    return changed


def _apply_run_font(run, font: str) -> bool:
    """把字体写到 run 的 w:rFonts 四个属性上（ascii/hAnsi/cs/eastAsia），返回是否有改动。"""
    changed = False
    try:
        if run.font.name != font:
            run.font.name = font
            changed = True
    except Exception:
        pass
    try:
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            if rFonts.get(qn(attr)) != font:
                rFonts.set(qn(attr), font)
                changed = True
    except Exception:
        pass
    return changed


# ---------------------------------------------------------------------------
# IEEE 参考文献 best-effort 引文格式重写（仅 IEEE；保守、幂等、不碰含域/超链接条目）
# ---------------------------------------------------------------------------
# 重要边界：本工具核心原则是「只改格式、不改内容」。参考文献**内容**重写属内容层改动，
# 与这一原则相悖、且任意 APA→IEEE 启发式都无法覆盖 IEEE 全部文献类型（期刊/会议/书/在线/标准…），
# 因此这里只做「明显 APA 作者-年 + 单 run 纯文本」条目的 best-effort 转换，并显式标注「请核对」。
# 用户若要求严格「格式-only」，可关闭此分支（仅保留编号 + 悬挂缩进）。

_APA_REF_RE = re.compile(
    r"^\s*(?P<authors>.+?)\s*\((?P<year>\d{4}(?:[-/]\d{1,4})?)\)\.\s*"
    r"(?P<title>.+?)\.\s*(?P<rest>.+?)\s*\.?$"
)


def _looks_like_ieee(text: str) -> bool:
    """幂等闸门：已呈 IEEE 形态（含 vol./no./pp. 或首字母缩写作者开头）则跳过。"""
    t = re.sub(r"^\[\s*\d+\s*\]\s*", "", (text or "").strip())
    if re.search(r"\bvol\.\s*\d", t) or re.search(r"\bno\.\s*\d", t) \
            or re.search(r"\bpp\.\s", t):
        return True
    if re.match(r"^[A-Z]\.\s+\S+", t):
        return True
    return False


def _apa_authors_to_ieee(s: str):
    """把 APA 作者串 best-effort 转成 IEEE 'F. Last and S. Lee' 形态。无法安全解析返回 None。

    用 re.findall 抓全部 'Last, Initials' 配对（Initials 可为 'T.' / 'T. M.'），
    天然兼容「姓, 名首字母, 姓, 名首字母, & 姓, 名首字母」这种一个 & 段里塞多位作者的
    APA 写法（例如 'Becker, T., Muller, R., & Schmidt, F.' → 三位作者都抓到）。
    """
    # 去掉可能残留的前导编号（"1. Anderson, K."）
    s = re.sub(r"^(?:\[\s*\d+\s*\]|\d+\s*[.)])\s*", "", (s or "").strip())
    pairs = re.findall(r"([A-Z][\w'-]*),\s*([A-Z](?:\.\s*[A-Z])*)", s)
    if not pairs:
        return None
    out = []
    for last, ini in pairs:
        last = last.strip()
        if not last:
            return None
        # ini 形如 'T' / 'T. M' / 'T.M' → 规整为 'T. M.'
        letters = re.findall(r"[A-Z]", ini)
        ini_ieee = ". ".join(letters) + "." if letters else ""
        out.append((ini_ieee + " " + last).strip())
    if len(out) == 1:
        return out[0]
    if len(out) == 2:
        return " and ".join(out)
    return ", ".join(out[:-1]) + ", and " + out[-1]   # IEEE 三位及以上用牛津逗号


def _apa_to_ieee(text: str):
    """把一条明显的 APA 条目 best-effort 转成 IEEE 编号引文串（不含 [n]）。失败返回 None。"""
    t = text.strip()
    # 去掉可能已存在的前导编号：IEEE 的 [n] 或 APA 编号列表的 "1." / "2)"（否则 "1."
    # 会被误吞进姓氏，变成 "K. 1. Anderson" 这类垃圾）
    t = re.sub(r"^(?:\[\s*\d+\s*\]|\d+\s*[.)])\s*", "", t)
    m = _APA_REF_RE.match(t)
    if not m:
        return None
    authors = _apa_authors_to_ieee(m.group("authors"))
    if authors is None:
        return None
    year = m.group("year")
    title = m.group("title").strip().rstrip(".")
    rest = m.group("rest").strip()
    jm = re.match(
        r"^(?P<journal>.+?),\s*(?P<vol>\d+)\s*\(\s*(?P<iss>\d+)\s*\)\s*,"
        r"\s*(?P<pages>[\d\-–—]+)\s*$",
        rest,
    )
    if jm:                                       # 期刊型
        journal = jm.group("journal").strip().rstrip(".")
        return (f'{authors}, "{title}," {journal}, vol. {jm.group("vol")}, '
                f'no. {jm.group("iss")}, pp. {jm.group("pages")}, {year}.')
    publisher = rest.rstrip(".").strip()          # 书型兜底
    if publisher:
        return f"{authors}, {title}. {publisher}, {year}."
    return None


def _reformat_ieee_reference_entries(doc, ref_indices) -> int:
    """对明显 APA 的 IEEE 参考文献条目做 best-effort 内容重写。返回重写条数。"""
    changed = 0
    for idx in ref_indices:
        para = doc.paragraphs[idx]
        if _para_has_field(para):                 # 含域/超链接 → 不碰（保留内联 URL）
            continue
        runs = [r for r in para.runs if not _is_comment_run(r)]
        if len(runs) != 1:                        # 多 run 重写风险高 → 跳过
            continue
        run = runs[0]
        text = run.text or ""
        if not text.strip() or _looks_like_ieee(text):
            continue
        new = _apa_to_ieee(text)
        if new and new != text:
            run.text = new
            changed += 1
    return changed



# ---------------------------------------------------------------------------
# 标题样式：直接格式化（绝不写 styleId），并加英文批注（规避国内版 Pit 1）
# ---------------------------------------------------------------------------

def _apply_tri_state(para, run, attr: str, want) -> bool:
    """三态布尔属性（bold / italic）落到 run 上，返回是否真的发生了改动。

    ``want`` 为 None = 本级不规定该属性 → 一个字节都不写（学校没说的不臆造）。
    为 True / False 时落**显式**值：显式值才是「我说了算」，不会随 Heading 样式
    被改动而漂移；比较用 ``is``，第二遍即 no-op，天然幂等。
    """
    if want is not True and want is not False:
        return False
    try:
        cur = getattr(run.font, attr)
    except Exception:
        return False
    if cur is want:
        return False
    try:
        setattr(run.font, attr, want)
        return True
    except Exception:
        return False


def _format_heading_one(para, f, body_font, body_size):
    """对单个标题段直接套格式（值落在段落 / run 属性，不引用任何样式 id）。

    返回是否实际发生了改动（用于幂等计数，避免「每遍都计 5 处改动」的假非零）。

    v2.3.65：支持「每级标题自己的」字体 / 字号 / 缩进量（学校模板里普遍是
    「一级黑体三号居中、二级楷体四号左对齐、三级宋体小四首行缩进 0.5 英寸」）。
    - ``f.font_family`` 优先于正文字体；留空则沿用正文字体（旧行为，不臆造）。
    - ``f.indent_in`` 给出精确缩进量；只给 ``f.indent=True`` 时仍回退旧的 0.5 英寸。
    - ``f.size_pt`` 在 run 没有显式字号时**也**要设：旧写法要求 run 本来就有字号，
      于是「同字号靠样式继承」的标题永远改不到，客户看到「字号没修」。
    """
    changed = False
    pf = para.paragraph_format
    if f.alignment:
        target_al = _ALIGN_ENUM.get(f.alignment)
        if pf.alignment != target_al:
            pf.alignment = target_al
            changed = True
    # 缩进量：显式英寸数优先；否则按旧开关回退 0.5 英寸；都没有则顶格（清缩进）
    indent_in = getattr(f, "indent_in", None)
    if indent_in is not None:
        target_indent = Inches(float(indent_in))
        cur = pf.first_line_indent
        if cur is None or abs(float(cur.inches) - float(indent_in)) > 1e-3:
            try:
                pf.first_line_indent = target_indent
                changed = True
            except Exception:
                pass
    else:
        # 显示型标题顶格（清首行缩进）；接排型(L4/L5)设首行缩进
        target_indent = Inches(0.5) if f.indent else None
        if (pf.first_line_indent is None) != (target_indent is None):
            try:
                pf.first_line_indent = target_indent
                changed = True
            except Exception:
                pass
    # 本级标题字体：显式给出则覆盖正文字体（很多学校「一级黑体、二级楷体」）
    head_font = getattr(f, "font_family", None) or body_font
    for run in para.runs:
        if _is_comment_run(run):
            continue          # 批注标记不是标题正文，跳过（否则第二遍恒显"已改"）
        # 三态加粗 / 斜体：None=本级不规定（不碰）；True/False=必须/必须不是。
        # 幂等靠 `is` 比较（run.font.bold 三态 True/False/None），不会每遍都算改动。
        if _apply_tri_state(para, run, "bold", getattr(f, "bold", None)):
            changed = True
        if _apply_tri_state(para, run, "italic", getattr(f, "italic", None)):
            changed = True
        if head_font:
            if run.font.name != head_font:
                try:
                    run.font.name = head_font
                    changed = True
                except Exception:
                    pass
            try:
                rPr = run._r.get_or_add_rPr()
                rFonts = rPr.find(qn("w:rFonts"))
                if rFonts is None:
                    rFonts = OxmlElement("w:rFonts")
                    rPr.insert(0, rFonts)
                for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                    if rFonts.get(qn(a)) != head_font:
                        rFonts.set(qn(a), head_font)
                        changed = True
            except Exception:
                pass
        if getattr(f, "size_pt", None):
            cur_sz = run.font.size
            if cur_sz is None or abs(float(cur_sz.pt) - float(f.size_pt)) > 1e-6:
                try:
                    run.font.size = Pt(f.size_pt)
                    changed = True
                except Exception:
                    pass
    if f.run_in and _ensure_trailing_period(para):
        changed = True
    # 段前分页（v2.3.67）：三态。None=本级不规定（不碰）；True/False 直接写 pPr/pageBreakBefore。
    # 幂等靠 is 比较（paragraph_format.page_break_before 三态 True/False/None）。
    _pbb = getattr(f, "page_break_before", None)
    if _pbb is not None:
        try:
            if bool(pf.page_break_before) != bool(_pbb):
                pf.page_break_before = bool(_pbb)
                changed = True
        except Exception:
            pass
    return changed


def _ensure_trailing_period(para) -> bool:
    """接排标题(L4/L5)段尾补句点（若缺失）。只动最后一个纯文本 run 的 text。返回是否改动。"""
    runs = para.runs
    if not runs:
        return False
    last = runs[-1]
    cur = last.text or ""
    if cur.rstrip().endswith("."):
        return False
    last.text = (cur.rstrip() + ". ") if cur.endswith(" ") else (cur + ". ")
    return True


# ---------------------------------------------------------------------------
# 章节编号（仅 IEEE 等顺序编号制标题风格使用）：罗马 / 字母 / 数字前缀 + 幂等维护
# ---------------------------------------------------------------------------

def _to_roman(n: int) -> str:
    """整数 → 罗马数字（n>=1；<=0 回退 '0'，编号场景不会出现）。"""
    if n <= 0:
        return "0"
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"),
            (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"),
            (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out


def _section_number_prefix(kind: str, n: int):
    """按编号样式生成前缀串（含尾随分隔空格）。不支持的 kind 返回 None。"""
    if kind == "roman":
        return _to_roman(n) + ". "
    if kind == "upper_letter":
        return chr(64 + n) + ". "
    if kind == "number":
        return str(n) + ". "
    if kind == "number_paren":
        return str(n) + ") "
    if kind == "lower_letter":
        return chr(96 + n) + ") "
    if kind == "lower_roman":
        return _to_roman(n).lower() + ") "
    return None


def _already_has_prefix(text: str, prefix: str) -> bool:
    """幂等判定：段落首 run 文本是否已（不区分大小写）以该前缀开头。"""
    if not text or not prefix:
        return False
    return text.lower().startswith(prefix.lower())


def _apply_heading_text_transform(para, f, prefix) -> bool:
    """对标题段做【全大写】与【章节编号前缀】（均幂等）。返回是否发生改动。

    - 全大写：逐 run 文本 upper()，天然幂等（upper(upper)=upper）。
    - 编号前缀：仅当首 run 文本尚未以该前缀开头时才在首 run 前插入；已带正确前缀
      则跳过，绝不重复叠加（多次运行 / 重跑均安全）。编号属 IEEE 等风格的显式约定，
      不视为正文改写（与参考文献 [n] 前缀同一立场）。
    """
    changed = False
    runs = [r for r in para.runs if not _is_comment_run(r)]
    if not runs:
        return False
    if f.uppercase:
        for r in runs:
            t = r.text or ""
            u = t.upper()
            if u != t:
                r.text = u
                changed = True
    if prefix:
        first = runs[0]
        cur = first.text or ""
        if not _already_has_prefix(cur, prefix):
            new = prefix + cur
            if new != cur:
                first.text = new
                changed = True
    return changed


def _heading_comment_text(spec, lvl, f, prefix=None) -> str:
    name = _spec_label(spec)
    txt = f"{name} Level {lvl} heading: {f.desc}. Applied."
    if prefix:
        txt += (" Uppercased." if f.uppercase else "") + f" Numbered '{prefix.strip()}'."
    return txt


def _assign_heading_levels(doc, target: TargetProfile, exclude_ps: set | None = None) -> dict:
    """给每个标题段分配 1-based 层级，返回 {w:p 元素: (level, is_pseudo)}。

    ⚠️ 键必须是段落的 **w:p 元素**（para._p）而不是 id(para)：python-docx 每次迭代
    doc.paragraphs 都会新建包装对象，id() 在两次迭代间不稳定，用它当键会全盘错位
    （实测导致正文被当成标题格式化）。

    识别两类标题：
    1) 带标题样式 / outline 的段：层级由 heading_level() 直接给出（最可靠，is_pseudo=False）；
    2) 无样式但整段加粗、短、无句读结尾的「伪标题」（IEEE/APA 常见 11–12pt 加粗章节
       标题）：当目标规范显式声明 heading_levels>0 时才识别（用户就是要给标题套格式/
       编号），is_pseudo=True。层级用*字号档*启发式推断：
       - 出现 ≥2 种不同字号 → 最大档=L1、次大=L2…（真实「大节/小节」场景可正确分层）；
       - 仅 1 种字号（均匀加粗、无样式）→ 全部当作 L1（best-effort，无法可靠推断层级，
         由调用方在汇总里显式提示用户核对）。

    ``exclude_ps``：需要**排除**的 w:p 集合（标题页标题/作者/单位）。标题页标题被
    标题页流程加粗后，会反过来满足「整段加粗 + 短」的伪标题条件 → 第二遍被当成
    章节标题编号（实测把论文大标题编成 "I."，并把所有章节编号整体推一位，
    产出 "II. I. INTRODUCTION" 这种脏结果）。故标题页段落必须由标题页流程独占。
    """
    out: dict = {}
    heading_levels = getattr(target, "heading_levels", None)
    # 兜底：只要目标声明了「每级标题格式」，就必须去识别标题 —— 不能因为客户/模板
    # 没单独填「层级数」就整条链空转（学校模板给每级格式却不填层级数是常态）。
    if not heading_levels and not (getattr(target, "heading_formats", None) or {}):
        return out
    exclude_ps = set(exclude_ps or ())
    # v2.3.73：论文大标题（首页居中加粗、全篇最大字号）不是章节标题。
    # 不摘出来 → 它占掉 L1 字号档、章节标题被推成 L2，检查与修正会一起按错误的
    # 分层跑（实测把论文标题改成模板的 14pt 左对齐）。与 read_docx 采样同一函数。
    _dti = doc_title_index(doc)
    if _dti is not None:
        exclude_ps.add(doc.paragraphs[_dti]._p)
    pseudo_paras = []
    for para in doc.paragraphs:
        if _para_has_field(para) or para._p in exclude_ps:
            continue
        hl = heading_level(para)
        if hl is not None:
            out[para._p] = (hl, False)
        else:
            pseudo_paras.append(para)
    if pseudo_paras:
        # 伪标题按字号档分层——与 docx_reader 的采样共用同一函数（单一真相源），
        # 否则「检查所依据的分层」与「修正所施加的分层」会各算一套。
        lvl_map = pseudo_heading_level_map(pseudo_paras, exclude_ps=exclude_ps)
        for p in pseudo_paras:
            lvl = lvl_map.get(p._p)
            if lvl is not None:
                out[p._p] = (lvl, True)
    return out


# 这些标题通常是论文后置部分，IEEE 不对其编号（但仍套格式：加粗/居中/全大写）
_NON_NUMBERED_HEADINGS = {
    "references", "reference", "bibliography", "works cited",
    "acknowledgments", "acknowledgements", "acknowledgment", "acknowledgement",
    "appendix", "appendices", "abstract", "keywords",
}


def _apply_heading_format(doc, target: TargetProfile, spec, level_map: dict | None = None,
                          only_levels=None) -> dict:
    """按规范每级规则直接格式化标题段（不写 styleId）。

    - ``only_levels``：v2.3.70 —— 只修这些层级（来自检查结论的 FixAction.levels）。
      传 None = 不限级（旧行为：整维度全级重刷）。编号计数器**始终全量推进**，
      否则跨级重置会断（多章节文档的二级编号会一路累加），只是不在范围内的层级
      不写格式、不加批注。

    - 仅处理已分配层级的标题段（样式/outline 或识别出的伪标题），且非域段落；
    - 层级由 _assign_heading_levels() 给出（1-based）；规范未定义的更深层级回退到最深一级；
      调用方应把同一份 level_map 传给 _select_paragraphs 做正文保护（见该函数说明）；
    - 若规范定义了 number_prefix（IEEE 等），按文档顺序维护计数（同层 +1、遇更浅层
      重置更深层）生成章节编号并幂等插入；uppercase 则全大写；
      后置部分（References/Appendix 等）按 IEEE 惯例不编号，仅套格式；
    - 每处理一段加一条英文批注说明规则；返回 {formatted, comments, pseudo_formatted,
      hierarchy_inferred}。
    """
    info = {"formatted": 0, "comments": 0, "pseudo_formatted": 0,
            "hierarchy_inferred": True}
    # v2.3.65：模板 / 问卷 / AI 显式给出的每级标题参数**优先于**内置规范；
    # 旧写法 `spec.heading_formats or generic_...()` 完全无视 target.heading_formats，
    # 学校模板写着「一级左对齐黑体」也会被写死的「L1 居中加粗」覆盖。
    explicit_hf = getattr(target, "heading_formats", None) or {}
    if not getattr(target, "heading_levels", None) and not explicit_hf:
        return info
    # 🔴 v2.3.73：规范的每级标题格式**只允许在「规范基线」模式下兜底**。
    # 非规范基线（学校模板 / 手动输入 / AI 导入）时，客户顶多在下拉里选了「参考文献按
    # IEEE」—— resolve_spec() 会因此返回整套 IEEE（含五级标题格式 + 编号 + 全大写）。
    # 若在这里兜底，客户的手排标题会被就地重写成 "A. Introduction"、"I. 大标题"，
    # 而预览（_heading_change_lines）拿的是 explicit，一行都不会提 —— 背着客户改文字。
    # 老板口径：选规范 = 全套按规范；不选规范 = 规范只借给参考文献，不许越界。
    spec_hf = (getattr(spec, "heading_formats", None)
               if spec_is_baseline(target) else None)
    fmt_map = resolve_heading_formats(
        explicit_hf,
        spec_hf,
        getattr(target, "heading_levels", None),
    )
    if not fmt_map:
        return info
    deepest = max(fmt_map.keys())
    page = target.page or {}
    body_font = page.get("font_family")
    body_size = page.get("font_size_pt")
    if level_map is None:
        level_map = _assign_heading_levels(doc, target)
    if any(is_pseudo for (_lvl, is_pseudo) in level_map.values()):
        # 存在靠伪标题识别（非样式）赋层的段 → 层级属 best-effort 推断
        info["hierarchy_inferred"] = False
    counters = {}                       # 层级 → 当前编号（IEEE 章节编号用）
    for para in doc.paragraphs:
        item = level_map.get(para._p)
        if item is None:
            continue
        lvl, is_pseudo = item
        f = fmt_map.get(lvl)
        if f is None:
            # 目标没有给这一级格式。两种情形必须区别对待（v2.3.71 客户视角实测）：
            #   ① 目标**只声明了部分层级**（学校模板常常只示范章标题，问卷也只填了 1 级）
            #      → 更深层级**没有任何依据**，绝不能拿一级格式去套：实测把二级标题也改成
            #        14pt 加粗居中，整篇层级关系当场崩掉，而且预览里一个字都没提。
            #   ② 规范声明了完整多级（APA/IEEE 五级），文档却更深（第 6 级）
            #      → 回退到已声明的最深一级，保持旧行为（编号体系不断档）。
            if len(fmt_map) <= 1:
                continue
            f = fmt_map.get(deepest)
        if f is None:
            continue
        # 后置部分（References/Appendix…）IEEE 不编号，仅套格式
        text_lc = (para.text or "").strip().lower()
        numbered = f.number_prefix and text_lc not in _NON_NUMBERED_HEADINGS
        _in_scope = only_levels is None or lvl in only_levels
        prefix = None
        if numbered:
            # 计数器**始终推进**（含不在本次修正范围内的层级）：编号的跨级重置依赖
            # 每一级都被计数，跳级不计数会让多章节文档的二级编号一路累加（A→B→C…）。
            counters[lvl] = counters.get(lvl, 0) + 1
            for deeper in [k for k in counters if k > lvl]:
                counters[deeper] = 0     # 回到更浅层 → 更深层编号归零
            if _in_scope:
                prefix = _section_number_prefix(f.number_prefix, counters[lvl])
        if not _in_scope:
            # 该级检查没报问题 → 一字节都不动（不写格式、不改大小写、不加批注）
            continue
        changed = _format_heading_one(para, f, body_font, body_size)
        if _apply_heading_text_transform(para, f, prefix):
            changed = True
        if changed:
            info["formatted"] += 1
            if is_pseudo:
                info["pseudo_formatted"] += 1
            add_comment(doc, para, _heading_comment_text(spec, lvl, f, prefix))
            info["comments"] += 1
    return info


# ---------------------------------------------------------------------------
# 标题页（title page）格式化：APA/IEEE 等含标题页论文的标题 / 作者 / 单位
# ---------------------------------------------------------------------------
# 背景：引擎的历史判定只把「带标题样式 / outline / 伪标题(整段加粗≥13pt)」的段落当标题。
# 而 APA/IEEE 标题页标题是 12pt、纯文本、无样式 —— 永远落不到标题判定，于是只被当正文刷字体，
# 不会居中也不会加粗（用户实测：跑完软件标题仍是左对齐、未加粗）。这是「只修能识别的标题」
# 这一架构立场的最显眼症状。
#
# 做法（保守、规则驱动、可回归）：仅当**存在标题页正面信号**（文档含 Abstract/Keywords 标签）
# 且首段不是已识别标题时，才把首段当作标题页标题、紧随其后的若干短行当作作者/单位，
# 直接套格式。具体「居中 / 加粗」由各风格在 specs.py 的 TitlePageFormat 显式声明
# （APA/IEEE=标题居中加粗、作者居中；MLA=标题居中不加粗、作者不居中；Other=留空整体跳过，
# 由模板驱动接管）。绝不对无信号文档盲猜。
# 检测实现已抽到 docx_reader.detect_title_page（checker 与 fixer 共用，单一真相源）：
# 若各写一套，会出现「检查按 A 段判、修正改了 B 段」的漂移。
_detect_title_page = detect_title_page


def _apply_title_page_format(doc, target: TargetProfile, spec) -> dict:
    """对 APA/IEEE 等含标题页的论文：标题页标题 / 作者 / 单位做直接格式化。

    **规则优先级**：客户显式给出（学校模板实测 / 问卷 / AI 的三态字段）> 内置规范的
    ``spec.title_page``。None 一律表示「不规定」→ 该属性一个字节都不动。
    显式来源一个都没给、且规范也没声明（如 Other / 自定义模板）→ 整体跳过。

    ⚠️ 调用方必须以 ``"title_page" in fix_dims`` 为前置条件（见 fix_docx）：
    这条链过去是「spec 有 title_page 就悄悄改」，客户看不到、报告里也没有，
    违反「检查→修正单一真相源」。现在由 checker 出结论 + 绑 FixAction 后才允许施加。
    复用 _format_heading_one 原语（绝不写 styleId、幂等）。
    返回 {title_formatted, comments}。
    """
    info = {"title_formatted": 0, "comments": 0}
    # 🔴 v2.3.73：标题页的「规范兜底」同样只在规范基线模式下成立。
    # 非规范基线时客户可能只是把「参考文献」选成了 APA —— APA 的 title_page 声明
    # （标题居中加粗 / 作者居中）不许因此飘进学校模板、手填、AI 导入的文档。
    tp_spec = spec.title_page if spec_is_baseline(target) else None
    # 客户显式三态优先；缺省才回落到规范声明
    t_center = getattr(target, "title_page_title_centered", None)
    t_bold = getattr(target, "title_page_title_bold", None)
    a_center = getattr(target, "title_page_author_centered", None)
    if tp_spec is not None:
        if t_center is None:
            t_center = tp_spec.title_centered
        if t_bold is None:
            t_bold = tp_spec.title_bold
        if a_center is None:
            a_center = tp_spec.author_centered
    if t_center is None and t_bold is None and a_center is None:
        return info
    tp = _detect_title_page(doc)
    if not tp:
        return info
    page = target.page or {}
    body_font = page.get("font_family")
    body_size = page.get("font_size_pt")

    title_align = "center" if t_center else (None if t_center is None else "left")
    title_bold = t_bold          # 三态：None = 不动加粗
    # 关键字传参：HeadingLevelFormat 第 7 个位置参数是 number_prefix（会真的往正文里
    # 插前缀），把描述串按位置传进去会把它当成编号前缀（实测曾埋雷）。
    title_fmt = HeadingLevelFormat(
        alignment=title_align,
        bold=title_bold,
        desc="%s title-page title%s" % (
            "centered" if t_center else "left-aligned",
            "" if t_bold is None else (" + bold" if t_bold else " + not bold")),
    )
    p = doc.paragraphs[tp["title"]]
    changed = _format_heading_one(p, title_fmt, body_font, body_size)

    author_align = "center" if a_center else (None if a_center is None else "left")
    for ai in tp["authors"]:
        pa = doc.paragraphs[ai]
        afmt = HeadingLevelFormat(
            alignment=author_align,
            desc="centered author/affiliation" if a_center else "left-aligned author/affiliation",
        )
        if _format_heading_one(pa, afmt, body_font, body_size):
            changed = True

    if changed:
        info["title_formatted"] = 1
        info["comments"] = 1
        bits = []
        if t_center:
            bits.append("title centered")
        if t_bold:
            bits.append("title bold")
        if a_center:
            bits.append("author/affiliation centered")
        add_comment(
            doc, p,
            "Title page formatted per %s: %s." % (_spec_label(spec), "; ".join(bits) or "no change"),
        )
    return info


def _spec_label(spec) -> str:
    """规范显示名（纯英文）：用 spec.key 而非 spec.name，避免规范名里的本地化
    后缀（如「美国心理学会」或「自定义…」）泄漏进用户可见的批注。"""
    key = getattr(spec, "key", None) or "Other"
    return "Custom" if key == "Other" else key


def _line_spacing_label(v) -> str:
    """行距数值 → 人类可读标签（single / 1.5 lines / double / Nx）。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return ""
    if abs(f - 1.0) <= 1e-6:
        return "single"
    if abs(f - 1.5) <= 1e-6:
        return "1.5 lines"
    if abs(f - 2.0) <= 1e-6:
        return "double"
    return f"{f:g}x"


def _align_label(v) -> str:
    return ALIGN_LABEL.get(str(v), str(v) if v is not None else "unknown")


def _before(prof, attr, default="unknown"):
    """安全取修前画像值；prof 为 None 或缺失时返回 default。"""
    if prof is None:
        return default
    val = getattr(prof, attr, None)
    return val if val is not None else default


def _margins_before(prof) -> str:
    m = (prof.margins if prof is not None else None) or {}
    def _one(k):
        v = m.get(k)
        return f"{float(v):.2f}\"" if v is not None else "n/a"
    return (f"top {_one('top')} / bottom {_one('bottom')} / "
            f"left {_one('left')} / right {_one('right')}")


def _build_summary_text(target, applied, spec, ref_info, prof, head_info=None) -> str:
    """生成本遍全部改动的英文「改动清单」批注（before → after）。

    仅报**真改了**的项；格式为带 from→to 的要点列表 + 作用范围说明，挂在文档
    首段与末段，让客户在任意一头都能一眼看到「改了什么、从什么改成什么、是否动了正文」。
    """
    has_ref = bool(ref_info and (ref_info.get("hanging_applied")
                                 or ref_info.get("numbering_added")))
    if not (applied or has_ref):
        return ""
    name = _spec_label(spec)
    lines: list = []

    # —— 页面 / 正文级（统一改动）——
    if applied.get("margins") is not None:
        after = f"{float(applied['margins']):.2f}\" on all four sides"
        lines.append(f"Page margins: {_margins_before(prof)}  →  {after}")
    if applied.get("font_family"):
        lines.append(f"Body font: {_before(prof, 'font_family')}  →  {applied['font_family']}")
    if applied.get("font_size_pt") is not None:
        b = _before(prof, "font_size_pt")
        if b == "unknown":
            lines.append(f"Font size: set to {applied['font_size_pt']} pt")
        else:
            lines.append(f"Font size: {b} pt  →  {applied['font_size_pt']} pt")
    if applied.get("line_spacing") is not None:
        raw = _before(prof, "line_spacing")
        b = _line_spacing_label(raw)
        a = _line_spacing_label(applied["line_spacing"])
        lines.append(f"Line spacing: set to {a}" if not b else f"Line spacing: {b}  →  {a}")
    if applied.get("body_alignment"):
        b = _align_label(_before(prof, "body_alignment"))
        a = _align_label(applied["body_alignment"])
        lines.append(f"Body alignment: {b}  →  {a}")
    if applied.get("first_line_indent_in") is not None:
        b = _before(prof, "first_line_indent_in")
        btxt = f"{float(b):.2f}\"" if b != "unknown" else "none"
        lines.append(f"First-line indent: {btxt}  →  {float(applied['first_line_indent_in']):.2f}\"")
    if applied.get("space_after_pt") is not None:
        b = _before(prof, "space_after_pt")
        btxt = f"{b} pt" if b != "unknown" else "0 pt"
        lines.append(f"Space after paragraph: {btxt}  →  {applied['space_after_pt']} pt")

    # —— 参考文献 ——
    if has_ref:
        bits = []
        if ref_info.get("hanging_applied"):
            bits.append("hanging indent applied")
        if ref_info.get("numbering_added"):
            bits.append(f"{ref_info['numbering_added']} entries numbered [1]..")
        lines.append(f"Reference list ({target.reference_style}): " + "; ".join(bits))
        if ref_info.get("rewritten"):
            lines.append(
                f"Reference entries best-effort rewritten to IEEE citation format: "
                f"{ref_info['rewritten']} (PLEASE VERIFY each entry — heuristic, "
                f"covers common journal/book APA patterns only)")

    # —— 章节标题 ——
    if head_info and head_info.get("formatted"):
        name = _spec_label(spec)
        lines.append(
            f"Section headings formatted per {name} (bold / alignment / numbering as "
            f"applicable): {head_info['formatted']} heading(s)")
        if not head_info.get("hierarchy_inferred"):
            lines.append(
                "Headings were detected as bold paragraphs WITHOUT heading styles, so "
                "section hierarchy was inferred best-effort (uniform-size bold titles "
                "were all treated as top-level). For correct I./A./1) sub-section "
                "numbering, apply Heading 1/2/3 styles or supply a template — "
                "PLEASE VERIFY the section levels.")

    if not lines:
        return ""
    body = "\n".join("• " + ln for ln in lines)
    ieee_ref_note = (
        "\n\nNote: for IEEE, reference entries were best-effort reformatted to IEEE "
        "citation style (author initials, quoted title, vol./no./pp., year). This is a "
        "heuristic content change — please verify each entry, as it covers only common "
        "journal/book APA patterns and may not match every source type."
    ) if (target.reference_style == "IEEE" and ref_info and ref_info.get("rewritten")) else ""
    return (
        f"PaperFormat Pro ({name}) \u2014 Format adjustments applied to this document:\n"
        f"{body}\n\n"
        f"Scope: these changes apply to all body paragraphs. Headings, the reference "
        f"list, and tables are handled separately (see their own comments). No body "
        f"text content was added, deleted, or rewritten"
        f"{ieee_ref_note}."
    )


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------

def _last_meaningful_paragraph(doc):
    """返回文档最后一个非空段落（用于把汇总批注也挂在文档末尾，方便客户在尾部查看）。

    若全篇皆空则退回最后一个段落；用底层 w:p 元素身份判断，避免 python-docx
    多次访问 doc.paragraphs 产生不同包装对象导致「同一段被加两条汇总」。
    """
    last = None
    for para in doc.paragraphs:
        if (para.text or "").strip():
            last = para
    if last is None:
        return doc.paragraphs[-1] if doc.paragraphs else None
    return last


def fix_docx(input_path: str, output_path: str, target: TargetProfile,
             findings: Optional[list] = None, spec=None) -> dict:
    """对 input 做纯格式修正，写入 output（不覆盖原件）。返回变更摘要 dict。

    raises:
        ValueError: 输出路径与输入指向同一文件（禁止覆盖原件）。
        DocxReadError: 文档无法读取。
        OSError: 输出无法写入（目录不存在/无权限）。
    """
    if not output_path or same_file(output_path, input_path):
        raise ValueError("The output path must differ from the input path (the original is not overwritten)")
    try:
        doc = Document(input_path)
    except Exception as e:  # not a docx / corrupted / encrypted
        raise DocxReadError(
            f"Cannot read the Word document: {e} (make sure it is a .docx file, not corrupted, and not encrypted)"
        ) from e

    strip_our_comments(doc)  # 清掉本工具上一遍批注（幂等，防 Pit 4 旧气泡带下）

    # 修前画像：用作批注的 before 值（profiling 失败不阻断修正，仅退化为无 before 值）
    try:
        prof = read_docx(input_path)
    except Exception:
        prof = None

    ref_indices = reference_paragraph_indices(doc)
    ref_set = set(ref_indices)

    # 规范解析：优先使用 caller 传入的 spec（来自 checker 的同一次解析，单一真相源），
    # 未传入时自行用 resolve_spec 解析——与 checker 走同一表达式，根除「规范还不一样」的漂移。
    if spec is None:
        spec = resolve_spec(target)
    # 修正项来源：优先消费 caller 传入的 findings（即报告里「检查出来的问题」），
    # 未传入时自算一次（向后兼容旧调用方）。fixer 只修正带 fix 的 Finding，绝不独立重算。
    if findings is None:
        _, findings = run_check(input_path, target)
    fix_dims = {f.fix.dimension for f in findings if getattr(f, "fix", None) is not None}

    # v2.3.70：标题维度再**细化到级**（detail 见 _heading_scope 的说明），
    # 并与预览 _compute_changes 共用同一函数 —— 单一真相源，杜绝「预览说改五级、实际改一级」。
    heading_only_levels = _heading_scope(findings)
    # v2.3.71：标题也必须有**真实结论**才动。检查器给每条 heading Finding 都绑了修正
    # （含 pass），若不门控，客户拿一份全合规的论文点「修正」→ 预览空、文档却被重排 +
    # 插入批注（实测 8 个标题 + 8 条批注）。只改检查报了的，是这条链的底线。
    _heading_reported = _heading_problems_reported(findings)

    # 参考文献「内容改写」（APA→IEEE）属于改内容，必须由检查结论驱动：
    # 检查器没报 reference 问题就一个字都不许改（旧版只看 reference_style 就改写）。
    _ref_reported = _reference_problems_reported(findings)

    # 标题页段落（标题/作者/单位）：其排版由 spec.title_page 独占管辖，必须在标题
    # 判定**之前**先摘出来 —— 否则标题被标题页流程加粗后会反过来满足伪标题条件，
    # 第二遍被当成章节标题编号（实测把论文大标题编成 "I."，章节编号整体推一位）。
    #
    # v2.3.71（客户视角实测的真 bug）：**学校模板模式**（spec_key=None → 解析到 Other，
    # 而 Other 没有 title_page 声明）本来压根不做这步，后果是一连串：
    #   ① 论文大标题（16pt 居中加粗）被当成伪标题 → 占掉 L1 字号档；
    #   ② 学生手排的章节标题（14pt）被推成 **L2**；
    #   ③ 模板只声明 1 级标题格式 → 检查器对 L2 无从比较（不报）、修正器按「只修报错的
    #      层级」跳过 L2 —— 客户点完「修正」，6 个章节标题一个字都没变，报告却写着已修正。
    # 所以只要目标声明了标题格式（模板/问卷/AI 任一来源），就要探测标题页并把它们摘出去。
    # v2.3.73：再加一条 —— 客户显式给了标题页三态（模板/问卷/AI）时同样要探测，否则那三条
    # 通道里「标题页是否居中加粗」永远没有探查对象。
    _need_title_page_probe = (
        (spec.title_page is not None and spec_is_baseline(target))
        or bool(getattr(target, "heading_formats", None))
        or any(getattr(target, _k, None) is not None
               for _k in ("title_page_title_centered", "title_page_title_bold",
                          "title_page_author_centered")))
    tp = _detect_title_page(doc) if _need_title_page_probe else None
    tp_ps: set = set()
    if tp:
        for idx in [tp["title"], *tp["authors"]]:
            if 0 <= idx < len(doc.paragraphs):
                tp_ps.add(doc.paragraphs[idx]._p)

    # 参考文献段不参与章节标题层级判定 —— 但**只排除条目**，标题行是否排除要看
    # 目标有没有显式管辖它：
    #   ① 目标给了 reference_heading_bold / _centered（学校模板、问卷、AI 都常给）
    #      → 参考文献标题的格式由 reference_heading 维度独占，必须排除，否则两套流程
    #        轮流改同一段，且 "References" 会因字号档被认成二级标题；
    #   ② 目标没管（如 IEEE 规范本身不规定标题格式）→ 保持旧行为，References 照常
    #      当后置标题套格式（大写/加粗），只是不编号（_NON_NUMBERED_HEADINGS）。
    # ⚠️ 必须与 tp_ps 分开：tp_ps 还兼作「不吃正文对齐/首行缩进」的豁免集合，
    # 若把参考文献段塞进去，参考文献标题会连正文对齐都不再修（实测打红 test_polish）。
    heading_exclude_ps = set(tp_ps)
    if ref_indices:
        heading_exclude_ps |= {doc.paragraphs[i]._p for i in ref_indices
                               if 0 <= i < len(doc.paragraphs)}
        _ref_head_owned = (target.reference_heading_bold is not None
                           or target.reference_heading_centered is not None)
        if _ref_head_owned:
            _rh_idx = ref_indices[0] - 1
            if 0 <= _rh_idx < len(doc.paragraphs):
                heading_exclude_ps.add(doc.paragraphs[_rh_idx]._p)

    # 标题判定只做一次，正文档保护与标题格式化共用同一份结果（见 _select_paragraphs
    # 文档字符串：两处判定不一致会导致标题被正文格式污染 + 「恒报已改」假非零）。
    # 仅当真的判出标题才覆盖：空集合会让 _select_paragraphs 跳过 is_heading 兜底，
    # 反而使真 Heading 样式段失去保护（被压成正文字号）。
    level_map = _assign_heading_levels(doc, target, exclude_ps=heading_exclude_ps)
    heading_ps = set(level_map.keys()) or None

    # 标题保护：no_heading 含参考文献；body_only 排除参考文献
    no_heading = _select_paragraphs(doc, ref_set, include_reference=True,
                                    heading_ps=heading_ps)
    body_only = _select_paragraphs(doc, ref_set, include_reference=False,
                                   heading_ps=heading_ps)

    # 标题页段落（标题/作者/单位）的**排版维度**由 spec.title_page 独占管辖：正文对齐
    # (body_alignment=justify) 与正文首行缩进都不能再刷它们 —— 否则两个流程每遍互相改回
    # 对方（对齐 CENTER↔JUSTIFY、首行缩进 None↔Inches(0)），造成「明明没变却恒报
    # title_page_formatted=1 / first_line_indent」的假非零，且结果取决于谁最后跑（顺序脆弱）。
    # 字体/字号/行距仍照常刷（标题页标题也该用正文字体），故只挡对齐与缩进。
    # （spec / tp / tp_ps 已在上面先行算好，供标题判定与正文保护共用）
    align_targets = [p for p in no_heading if p._p not in tp_ps]

    # 首行缩进豁免集合：标题页段落 + 摘要正文段 + 块引用段。
    # 摘要正文必须顶格（不缩进）、块引用应左缩进而非首行缩进——这三类都不能吃正文 0.5″ 首行缩进。
    abstract_ps = {doc.paragraphs[i]._p for i in _abstract_body_indices(doc)}
    # 块引用集合 = 「Quote 样式段」∪「checker 认定的长引文候选」。
    # 后者必须算进来：学生常常把长引文直接写成 Normal 段，检查器会判「应缩进 0.5″」，
    # 但旧版修正只认 Quote 样式 → 客户看到问题、点完修正一处没动（检查/修正不同源）。
    quote_ps = {p._p for p in body_only if _is_block_quote(p)}
    for _c in (getattr(prof, "block_quote_candidates", None) or []):
        if not _c.get("has_quote"):
            continue
        _li = _c.get("left_indent_in")
        if _li is not None and abs((_li or 0) - 0.5) <= 0.15:
            continue                       # 已经是 0.5″ 块引用，不用动
        _ci = _c.get("idx")
        if isinstance(_ci, int) and 0 <= _ci < len(doc.paragraphs):
            quote_ps.add(doc.paragraphs[_ci]._p)
    indent_exempt_ps = tp_ps | abstract_ps | quote_ps
    indent_targets = [p for p in body_only if p._p not in indent_exempt_ps]

    summary: dict = {}
    applied: dict = {}

    page = target.page or {}
    if page:
        # 仅当 checker 在该维度发现需修正（带 fix 的 Finding）时才施加——fixer 消费
        # check 的结论，绝不独立重算「要改什么」，从而与报告严格一致。
        if "margins" in fix_dims and _apply_margins(doc, page):
            applied["margins"] = page.get("margin_top_in")
        # 🔴 字体与字号**分开门控**：两者共用 _apply_font_and_size，旧写法只要「字号要改」
        # 就把 page 里整个 font_family 一起传进去 —— 结果客户用 Calibri（APA 允许的字体）
        # 只因字号差 1pt，字体也被换成 Times New Roman。一个维度只能有一个 owner。
        _want_font = page.get("font_family") if "font" in fix_dims else None
        _want_size = page.get("font_size_pt") if "font_size" in fix_dims else None
        if (_want_font or _want_size is not None) and _apply_font_and_size(
                {"font_family": _want_font, "font_size_pt": _want_size}, no_heading):
            if _want_font:
                applied["font_family"] = _want_font
            if _want_size is not None:
                applied["font_size_pt"] = _want_size
        if "line_spacing" in fix_dims and _apply_line_spacing(page, no_heading):
            applied["line_spacing"] = page.get("line_spacing")
        if "alignment" in fix_dims and _apply_alignment(page, align_targets):
            applied["body_alignment"] = page.get("body_alignment")
        # 只把**真发生了改动**的项记进 applied（原来这两处不看返回值、无条件记录，
        # 汇总批注于是会声称"已设缩进/段后间距"，而原文可能本来就合规）
        if "first_line_indent" in fix_dims and _apply_first_line_indent(page, indent_targets):
            applied["first_line_indent_in"] = page.get("first_line_indent_in")
        if "space_after" in fix_dims and _apply_space_after(page, indent_targets):
            applied["space_after_pt"] = page.get("space_after_pt")
        if "header_distance" in fix_dims and _apply_header_distance(doc, page):
            applied["header_distance_in"] = page.get("header_distance_in")
        # 块引用：左缩进 0.5″、首行不缩进（已从上方的 indent_targets 豁免，这里单独正向施加）。
        # 属正文格式归一化，不依赖某条 finding，保持原有「检测到块引用即施加」行为。
        # 块引用：同样改成由检查结论驱动（checker 绑 block_quote FixAction 才动手）。
        if "block_quote" in fix_dims and quote_ps \
                and _apply_block_quote_indent([p for p in body_only if p._p in quote_ps]):
            applied["block_quote_indent_in"] = 0.5
        summary["page"] = dict(page)

    # 栏数（结构性维度）：仅当 checker 绑定了 columns 修正时才施加（单一真相源）。
    # 它不属于 page 字典，故独立于上面的 if page: 块。
    if "columns" in fix_dims and _apply_columns(doc, target):
        applied["columns"] = target.columns

    # 纸张大小（通用维度：Letter/A4）：同样独立于 page 字典。
    if "paper_size" in fix_dims and _apply_paper_size(doc, target):
        applied["paper_size"] = target.paper_size

    # 页码位置（v2.3.66）：fixer 第一次操作页眉/页脚 part，必须由检查结论驱动。
    # 只有客户/学校/AI 显式要求时 checker 才会绑这个 FixAction（spec 来源只提示）。
    if "page_number" in fix_dims and target.page_number_position:
        _pn_info = _apply_page_number(doc, target.page_number_position)
        if _pn_info:
            applied["page_number"] = _pn_info

    # 摘要页特化（v2.3.67）：只有客户/学校/AI 显式要求时 checker 才绑 FixAction。
    if "abstract_indent" in fix_dims and target.abstract_first_line_indent_in is not None:
        if _apply_abstract_indent(doc, target.abstract_first_line_indent_in):
            applied["abstract_first_line_indent_in"] = target.abstract_first_line_indent_in
    if "abstract_new_page" in fix_dims and target.abstract_new_page is not None:
        if _apply_abstract_new_page(doc, target.abstract_new_page):
            applied["abstract_new_page"] = target.abstract_new_page

    # 图表 caption 字号/行距（v2.3.67）：只有客户/学校/AI 显式要求时 checker 才绑 FixAction。
    if "caption_size" in fix_dims and target.caption_font_size_pt is not None:
        if _apply_caption_size(doc, target.caption_font_size_pt):
            applied["caption_font_size_pt"] = target.caption_font_size_pt
    if "caption_spacing" in fix_dims and target.caption_line_spacing is not None:
        if _apply_caption_spacing(doc, target.caption_line_spacing):
            applied["caption_line_spacing"] = target.caption_line_spacing

    # 字体白名单的「字体 → 字号」配对：逐 run 按各自字体改字号，不动 w:rFonts。
    if "font_size_by_font" in fix_dims:
        _pairs = getattr(spec, "font_allowlist_sizes", None) if spec is not None else None
        if not _pairs:
            try:
                _pairs = getattr(resolve_spec(target), "font_allowlist_sizes", None)
            except Exception:
                _pairs = None
        if _pairs:
            _n_size, _size_detail = _apply_size_by_font_family(no_heading, _pairs)
            if _n_size:
                applied["font_size_by_font"] = _size_detail

    # v2.3.65：不再用 reference_style 做总闸 —— 学校模板可能只给「条目五号宋体、
    # 单倍行距」而没选任何引用风格，旧写法会让这些修正整段跳过（填了也不生效）。
    _ref_dims = {"reference_hanging_indent", "reference_numbered", "reference_font",
                 "reference_size", "reference_spacing", "reference_new_page",
                 "reference_heading"}
    if target.reference_style or (_ref_dims & fix_dims):
        ref_info = _apply_reference_format(
            doc, target, ref_indices,
            fix_hang="reference_hanging_indent" in fix_dims,
            fix_number="reference_numbered" in fix_dims,
            fix_font="reference_font" in fix_dims,
            fix_size="reference_size" in fix_dims,
            fix_spacing="reference_spacing" in fix_dims,
            fix_new_page="reference_new_page" in fix_dims,
            fix_heading="reference_heading" in fix_dims,
            fix_rewrite=_ref_reported,
        )
        summary["reference_style"] = target.reference_style
        summary["reference_hanging_indent_in"] = target.reference_hanging_indent_in
        summary["reference_numbered"] = target.reference_numbered
        summary["reference_numbering_added"] = ref_info["numbering_added"]
        summary["reference_entries_reformatted"] = ref_info["entry_reformatted"]
        summary["reference_heading_formatted"] = ref_info["heading_formatted"]
        summary["reference_new_page_applied"] = ref_info["new_page_applied"]
        summary["reference_entry_count"] = len(ref_indices)

    # 标题样式：直接格式化（不写 styleId）+ 逐段英文批注。
    # 仅当 checker 绑定了 heading 修正动作时才施加（与「检查→修正」单一真相源一致）。
    head_info = {"formatted": 0, "pseudo_formatted": 0, "hierarchy_inferred": 0, "comments": 0}
    if "heading" in fix_dims and _heading_reported:
        head_info = _apply_heading_format(doc, target, spec, level_map=level_map,
                                          only_levels=heading_only_levels)
    summary["headings_formatted"] = head_info["formatted"]
    summary["headings_pseudo_formatted"] = head_info["pseudo_formatted"]
    summary["headings_hierarchy_inferred"] = head_info["hierarchy_inferred"]

    # 标题页（title page）：APA/MLA 等标题页论文的标题/作者/单位直接格式化
    # 标题页：必须由检查结论驱动（checker 的 title_page 维度绑了 FixAction 才动手）。
    # 旧版这里无条件调用 → 只要 spec.title_page 有值就悄悄改标题页，客户看不到。
    tp_info = {"title_formatted": 0, "comments": 0}
    if "title_page" in fix_dims:
        tp_info = _apply_title_page_format(doc, target, spec)
    summary["title_page_formatted"] = tp_info["title_formatted"]

    # 批注：文档首段 + 末段各一条「改动清单」汇总本遍全部改动；参考文献段加批注
    summary_text = _build_summary_text(
        target, applied, spec, ref_info if target.reference_style else None, prof, head_info
    )
    comments_added = head_info["comments"]
    comments_added += tp_info["comments"]
    if summary_text and doc.paragraphs:
        first_para = doc.paragraphs[0]
        add_comment(doc, first_para, summary_text)
        comments_added += 1
        last_para = _last_meaningful_paragraph(doc)
        if last_para is not None and last_para._p is not first_para._p:
            add_comment(doc, last_para, summary_text)
            comments_added += 1
    if target.reference_style and ref_indices and ref_info and (
            ref_info.get("hanging_applied") or ref_info.get("numbering_added")):
        ref_heading_idx = ref_indices[0] - 1
        if 0 <= ref_heading_idx < len(doc.paragraphs):
            rc = (f"References formatted per {target.reference_style}: hanging indent applied"
                  + (f"; {ref_info['numbering_added']} entries numbered [1].."
                     if ref_info.get("numbering_added") else ""))
            add_comment(doc, doc.paragraphs[ref_heading_idx], rc)
            comments_added += 1
    summary["comments_added"] = comments_added

    # 记录细分应用项，便于 CLI/GUI 展示与测试断言
    summary["applied"] = applied
    summary["headings_preserved"] = True   # 标题段恒被跳过（铁律）

    try:
        doc.save(output_path)
    except Exception as e:
        raise OSError(f"Could not write output file: {e}") from e
    return summary


# ---------------------------------------------------------------------------
# 预览（现状 → 目标 diff，不落盘）
# ---------------------------------------------------------------------------
# i18n：修正预览是「用户可见的 GUI 文案」，必须随界面语言走。海外版默认英文，
# 不能因为引擎是中国团队写的就硬编码中文（老板 2026-10-04 硬要求：英文环境零中文）。
# GUI 把当前 self.lang 透传进来；引擎不反向依赖 ui 包，对齐标签本地内置一份。

_DIM_LABELS = {
    "en": {
        "margin_top": "Top margin", "margin_bottom": "Bottom margin",
        "margin_left": "Left margin", "margin_right": "Right margin",
        "font_family": "Body font", "font_size": "Font size",
        "line_spacing": "Line spacing", "body_alignment": "Body alignment",
        "first_line_indent": "First-line indent", "space_after": "Space after paragraph",
        "header_distance": "Header distance from top",
        "columns": "Column layout",
        "paper_size": "Paper size",
        "page_number": "Page number",
        "font_size_by_font": "Body size for this font",
        "abstract_indent": "Abstract first-line indent",
        "abstract_new_page": "Abstract on new page",
        "caption_size": "Caption font size",
        "caption_spacing": "Caption line spacing",
        "title_page": "Title page", "block_quote": "Block quotation",
        "ref_hanging": "Reference hanging indent", "ref_numbering": "Reference numbering",
        "heading_lvl": "Level %s heading",
        "ref_font": "Reference entry font", "ref_size": "Reference entry font size",
        "ref_spacing": "Reference entry line spacing",
        "ref_new_page": "References on new page",
        "ref_heading": "Reference heading format",
        "ref_entry": "Reference entry format",
    },
    "zh": {
        "margin_top": "上页边距", "margin_bottom": "下页边距",
        "margin_left": "左页边距", "margin_right": "右页边距",
        "font_family": "正文字体", "font_size": "正文字号",
        "line_spacing": "行距", "body_alignment": "正文对齐",
        "first_line_indent": "正文首行缩进", "space_after": "段后间距",
        "header_distance": "页眉距页顶",
        "columns": "栏数",
        "paper_size": "纸张大小",
        "page_number": "页码位置",
        "font_size_by_font": "正文字号（按字体）",
        "abstract_indent": "摘要首行缩进",
        "abstract_new_page": "摘要另起一页",
        "caption_size": "图表标题字号",
        "caption_spacing": "图表标题行距",
        "title_page": "标题页", "block_quote": "块引用",
        "ref_hanging": "参考文献悬挂缩进", "ref_numbering": "参考文献编号",
        "heading_lvl": "第 %s 级标题",
        "ref_font": "参考文献条目字体", "ref_size": "参考文献条目字号",
        "ref_spacing": "参考文献条目行距",
        "ref_new_page": "参考文献另起一页",
        "ref_heading": "参考文献标题格式",
        "ref_entry": "参考文献条目格式",
    },
}
_ALIGN_ZH = {"left": "左对齐", "justify": "两端对齐", "center": "居中", "right": "右对齐"}
# 页码位置的人类可读文案（预览/批注共用）。语义按「页眉 / 页脚」写清楚 ——
# 客户最容易搞错的就是把 APA 的 top-right 理解成「页脚右上」。
_PN_LABELS = {
    "en": {
        "top-right": "top-right of the header",
        "bottom-center": "bottom-center of the footer",
        "bottom-right": "bottom-right of the footer",
        "header": "header", "footer": "footer",
        "auto": "automatic field",
        "typed": "typed number",
    },
    "zh": {
        "top-right": "页眉右上",
        "bottom-center": "页脚居中",
        "bottom-right": "页脚右下",
        "header": "页眉", "footer": "页脚",
        "auto": "自动域",
        "typed": "手打数字",
    },
}
_UNREC = {"en": "unrecognized", "zh": "未识别"}
_UNKNOWN = {"en": "unknown", "zh": "未知"}
_NO_CHANGE = {
    "en": "This document already matches the target format — no changes needed.",
    "zh": "当前文档已符合目标格式，无需修正。",
}
_NOTES_TMPL = {
    "en": "Note: text inside tables ({n} paragraph(s)) is not modified; only body "
          "paragraph formatting is fixed this run.",
    "zh": "说明：表格内的文字（共 {n} 段）不会被修改；本遍仅修正正文段落格式。",
}


def _lang_of(lang):
    return lang if lang in ("en", "zh") else "en"


def _heading_scope(findings) -> frozenset | None:
    """从检查结论里读出「到底报了哪几级标题」，修正时只动这几级。

    v2.3.70。老板定的执行逻辑：先检查、再按检查结果改，不是自己瞎改。
    检查器现在把 heading 结论细化到级（``FixAction.levels``，例如 ``(3,)``
    表示「Level 3 heading format」不合格），修正侧就必须读得到这个明细：

      · 有带 levels 的 heading 动作 → 只修这些级（计数器仍全量推进，保证 II./III. 连续）；
      · **只有不带 levels 的整体结论**（如「没有 Heading 样式」「层级数超出上限」）→
        返回**空集合**：这类结论说明的是「层级识别/层级数」问题，不给任何一级
        授权改格式。旧行为把它当「不限级」→ 全部层级重刷，实测把论文大标题编成
        "I."、把章节标题改成 "A. Introduction"，而预览一行都没提（v2.3.73 修）；
      · 完全没有 heading 动作 → 返回 None（fix_dims 本就不含 heading，不会修）。

    ⚠️ **pass 的结论一律不算证据**：检查器每条 heading Finding 都会绑 fix（含
    「Heading levels: Detected 3 levels」这种已经通过的整维度总览）。若不跳过它，
    只要文档里有标题，这条 pass 总览就会把范围撑成「不限级」→ 每级标题全量重刷，
    刚做好的「只改报错层级」当场作废。

    返回 ``None`` = 不限级（仅旧调用方不传 findings 时出现）。
    ``fix_docx`` 与 ``_compute_changes`` 共用本函数，
    保证「预览说要改的」与「实际会改的」逐级一致。
    """
    if findings is None:
        return None
    scoped: set = set()
    for f in findings:
        fx = getattr(f, "fix", None)
        if fx is None or fx.dimension != "heading":
            continue
        if (getattr(f, "severity", "") or "") == "pass":
            continue
        lvs = getattr(fx, "levels", ()) or ()
        scoped.update(int(x) for x in lvs if int(x) > 0)
    return frozenset(scoped)


def _reference_problems_reported(findings) -> bool:
    """检查器是否**真的报了**参考文献问题 —— 决定能不能改写条目内容（APA→IEEE）。

    内容层改写必须由检查结论驱动：条目本身合规时一个字都不许动。
    ``findings is None`` 属旧调用方（fix_docx 内部自会跑一次检查，不会走到这里），
    保守返回 False（宁可不改，绝不乱改）。
    """
    if findings is None:
        return False
    return any(f.dimension == "reference" and f.severity != "pass" for f in findings)


def _heading_problems_reported(findings) -> bool:
    """检查器是否**真的报了**标题格式问题 —— 决定要不要动标题。

    v2.3.71 客户视角实测的真 bug：检查器为了「客户填了标题参数就得生效」，会给
    **每一条** heading Finding 绑修正动作（含已通过的）。于是 ``fix_dims`` 里永远有
    heading，客户拿一份本来就合规的论文点「修正」：
        · 预览一行都没列（没东西要改），
        · 文档里 8 个标题被重新套格式、还插进 8 条批注（客户：谁让你动的？）。
    这不是「多改了点」，是「预览说没改动、实际改了文档」——信任问题。

    现在：只有检查真的报了标题问题（非 pass）才动标题。``findings is None`` 时
    保守返回 False（宁可不改）。
    """
    if findings is None:
        return False
    return any(f.dimension == "heading" and f.severity != "pass" for f in findings)


def _dim_label(lang, key):
    return _DIM_LABELS[_lang_of(lang)].get(key, key)


def _align_text(lang, v):
    """对齐值 → 显示名（按语言）。"""
    if _lang_of(lang) == "zh":
        return _ALIGN_ZH.get(str(v), str(v) if v is not None else _UNKNOWN["zh"])
    return ALIGN_LABEL.get(str(v), str(v) if v is not None else _UNKNOWN["en"])


def _diff_line(label: str, cur, exp, unit: str, lang="en") -> str | None:
    """生成一行 diff；无变化返回 None。文案随语言走。"""
    L = _lang_of(lang)
    if exp is None:
        return None
    if cur is None:
        return f"{label}：将设为 {exp}{unit}" if L == "zh" else f"{label}: set to {exp}{unit}"
    try:
        same = abs(float(cur) - float(exp)) <= 1e-6
    except (TypeError, ValueError):
        same = str(cur) == str(exp)
    if same:
        return None
    if L == "zh":
        return f"{label}：{cur}{unit} → {exp}{unit}"
    return f"{label}: {cur}{unit} → {exp}{unit}"


def _compute_changes(prof, target: TargetProfile, lang="en", findings=None) -> list:
    """根据实测画像算出「真正会发生变化」的格式修正项（可能为空列表）。

    单一真相源：若传入 findings（来自 ``checker.run_check`` 的同一份比对结论），则只列出
    checker 判定「需要且可自动修正」的维度（fix_dims），从而与 ``fix_docx`` 实际施加的项
    严格一致——预览里写的就是 fixer 真会改的，绝不出现「预览说要改、修了却没改」或反之。
    未传入 findings 时（向后兼容旧调用方）不门控，退回原「全维度 prof-vs-target diff」。

    ``lang`` 决定每行 diff 的文案语言（"en" / "zh"），由 GUI 透传当前界面语言，
    保证英文界面不泄漏中文（见模块顶部标签表）。
    """
    L = _lang_of(lang)
    lines: list = []
    page = target.page or {}
    # fix_dims：checker 绑定了 FixAction 的维度集合；None 表示不门控（兼容旧调用）。
    fix_dims = (
        {f.fix.dimension for f in findings if getattr(f, "fix", None) is not None}
        if findings is not None else None
    )

    def _gate(*dims) -> bool:
        # findings 为 None（兼容旧调用）时不设门控，退回全维度 diff
        return fix_dims is None or any(d in fix_dims for d in dims)

    # —— 页边距 ——
    if _gate("margins"):
        for key, side, _label in _MARGIN_KEYS:
            if page.get(key) is not None:
                dim = key[:-3] if key.endswith("_in") else key
                line = _diff_line(_dim_label(L, dim), prof.margins.get(side),
                                  float(page[key]), '"', lang)
                if line:
                    lines.append(line)

    # —— 字体 / 字号 / 行距 ——
    if _gate("font", "font_size"):
        if page.get("font_family"):
            cur_f = prof.font_family
            if (cur_f or "").lower() != str(page["font_family"]).lower():
                cur_disp = _UNKNOWN[L] if not cur_f else cur_f
                lines.append(f"{_dim_label(L, 'font_family')}: {cur_disp} → {page['font_family']}")
        line = _diff_line(_dim_label(L, "font_size"), prof.font_size_pt,
                          page.get("font_size_pt"), "pt", lang)
        if line:
            lines.append(line)
    if _gate("line_spacing"):
        line = _diff_line(_dim_label(L, "line_spacing"), prof.line_spacing,
                          page.get("line_spacing"), "x", lang)
        if line:
            lines.append(line)

    # —— 段落级 ——
    if _gate("alignment") and page.get("body_alignment"):
        cur_a = prof.body_alignment
        if cur_a != page["body_alignment"]:
            lines.append("{}: {} → {}".format(
                _dim_label(L, "body_alignment"),
                _align_text(L, cur_a), _align_text(L, page["body_alignment"])))
    if _gate("first_line_indent"):
        line = _diff_line(_dim_label(L, "first_line_indent"), prof.first_line_indent_in,
                          page.get("first_line_indent_in"), '"', lang)
        if line:
            lines.append(line)
    if _gate("space_after"):
        line = _diff_line(_dim_label(L, "space_after"), prof.space_after_pt,
                          page.get("space_after_pt"), "pt", lang)
        if line:
            lines.append(line)
    if _gate("header_distance"):
        line = _diff_line(_dim_label(L, "header_distance"), prof.header_distance_in,
                          page.get("header_distance_in"), '"', lang)
        if line:
            lines.append(line)

    # —— 栏数（结构性维度）：单列一行，保证预览与 fixer 实际施加一致 ——
    if _gate("columns") and target.columns is not None:
        cur_c = prof.columns if prof.columns is not None else 1
        want_c = int(target.columns)
        if cur_c != want_c:
            lines.append("%s: %d → %d" % (_dim_label(L, "columns"), cur_c, want_c))

    # —— 纸张大小（通用维度：Letter / A4）：北美 vs 欧洲，与引用规范无关 ——
    if _gate("paper_size") and target.paper_size:
        want_p = str(target.paper_size)
        cur_p = getattr(prof, "paper_size", None)
        cur_desc = describe_paper_size(getattr(prof, "paper_width_in", None),
                                       getattr(prof, "paper_height_in", None))
        if cur_p != want_p:
            sep = " → " if L != "zh" else " → "
            lines.append("%s: %s%s%s" % (_dim_label(L, "paper_size"),
                                         cur_desc if cur_desc != "unknown" else _UNKNOWN[L],
                                         sep, want_p))

    # —— 标题页 / 块引用：这两项以前会改但预览里看不见，现在补上 ——
    # —— 页码位置（v2.3.66）——
    # 三种情形都要能让客户在改之前看见：
    #   ① 原来没有页码 → 「— → 页眉右上」
    #   ② 有但位置/对齐不对 → 「页脚居中 → 页眉右上」
    #   ③ 位置对但是手打的 → 「手打数字 → 自动域」（客户必须知道自己那串数字被换成域了）
    if _gate("page_number") and target.page_number_position:
        _pn_want = str(target.page_number_position)
        _pn_kind, _pn_align = PAGE_NUMBER_LAYOUT.get(_pn_want, ("header", "right"))
        _pn_lab = _PN_LABELS[L]
        _want_txt = _pn_lab.get(_pn_want, _pn_want)
        _cur_in = getattr(prof, "page_number_in", None)
        _cur_al = getattr(prof, "page_number_alignment", None)
        _line = None
        if not getattr(prof, "page_number_present", False):
            _line = "%s: %s → %s" % (_dim_label(L, "page_number"),
                                     _UNKNOWN[L] if L == "en" else "无页码",
                                     _want_txt)
        elif _cur_in != _pn_kind or _cur_al != _pn_align:
            _cur_txt = "%s (%s)" % (_pn_lab.get(str(_cur_in), str(_cur_in)),
                                    _pn_lab.get(str(_cur_al), str(_cur_al))
                                    if str(_cur_al) in ("left", "center", "right") else str(_cur_al))
            _line = "%s: %s → %s" % (_dim_label(L, "page_number"), _cur_txt, _want_txt)
        elif getattr(prof, "page_number_is_static", False):
            _line = "%s: %s → %s" % (_dim_label(L, "page_number"),
                                     _pn_lab["typed"], _pn_lab["auto"])
        if _line:
            lines.append(_line)

    # —— 摘要页特化（v2.3.67）——
    # 摘要首行缩进 + 是否另起一页，两行独立预览；与 fixer 实际施加一致（fix_dims 门控）。
    if _gate("abstract_indent") and target.abstract_first_line_indent_in is not None:
        _cur_abs_fli = getattr(prof, "abstract_first_line_indent_in", None)
        if _cur_abs_fli is not None and abs(float(_cur_abs_fli) - float(target.abstract_first_line_indent_in)) > 0.01:
            lines.append("%s: %g″ → %g″" % (
                _dim_label(L, "abstract_indent"),
                float(_cur_abs_fli), float(target.abstract_first_line_indent_in)))
    if _gate("abstract_new_page") and target.abstract_new_page is not None:
        _cur_abs_np = getattr(prof, "abstract_new_page", None)
        if _cur_abs_np is not None and bool(_cur_abs_np) != bool(target.abstract_new_page):
            if L == "zh":
                lines.append("%s：%s" % (_dim_label(L, "abstract_new_page"),
                                         "接排 → 另起一页" if target.abstract_new_page else "另起一页 → 接排"))
            else:
                lines.append("%s: %s" % (_dim_label(L, "abstract_new_page"),
                                         "same page → new page" if target.abstract_new_page
                                         else "new page → same page"))

    # —— 图表 caption 字号/行距（v2.3.67）——
    if _gate("caption_size") and target.caption_font_size_pt is not None:
        _cur_cap_size = getattr(prof, "caption_font_size_pt", None)
        if _cur_cap_size is not None and abs(float(_cur_cap_size) - float(target.caption_font_size_pt)) > 0.01:
            lines.append("%s: %g → %g pt" % (
                _dim_label(L, "caption_size"),
                float(_cur_cap_size), float(target.caption_font_size_pt)))
    if _gate("caption_spacing") and target.caption_line_spacing is not None:
        _cur_cap_sp = getattr(prof, "caption_line_spacing", None)
        if _cur_cap_sp is not None and abs(float(_cur_cap_sp) - float(target.caption_line_spacing)) > 0.01:
            lines.append("%s: %g → %g x" % (
                _dim_label(L, "caption_spacing"),
                float(_cur_cap_sp), float(target.caption_line_spacing)))
    # 只改字号、不换字体：用 Calibri 就按 Calibri 的规范字号（11pt），用 Times New Roman
    # 就按 12pt，正文混用两种字体时**各改各的**（所以逐字体列，不只列主字体）。
    if _gate("font_size_by_font"):
        try:
            _pairs = getattr(resolve_spec(target), "font_allowlist_sizes", None) or {}
        except Exception:
            _pairs = {}
        _fam_sizes = dict(getattr(prof, "font_size_by_family", None) or {})
        if not _fam_sizes and getattr(prof, "font_family", None):
            _fam_sizes = {prof.font_family: getattr(prof, "font_size_pt", None)}
        for _f, _cur in _fam_sizes.items():
            _want_pt = size_for_font(_f, _pairs)
            if not _want_pt or not _cur:
                continue
            if abs(float(_cur) - float(_want_pt)) > 0.01:
                lines.append("%s: %s %g → %g pt" % (
                    _dim_label(L, "font_size_by_font"), _f,
                    float(_cur), float(_want_pt)))

    if _gate("title_page") and getattr(prof, "title_page_detected", False):
        _tp_bits = []
        if target.title_page_title_centered is not None:
            _w = "center" if target.title_page_title_centered else "left"
            if prof.tp_title_alignment is not None and prof.tp_title_alignment != _w:
                _tp_bits.append("%s → %s" % (_align_text(L, prof.tp_title_alignment),
                                             _align_text(L, _w)))
        if target.title_page_title_bold is not None and prof.tp_title_bold is not None \
                and bool(prof.tp_title_bold) != bool(target.title_page_title_bold):
            _tp_bits.append(("加粗" if target.title_page_title_bold else "取消加粗")
                            if L == "zh" else
                            ("bold" if target.title_page_title_bold else "not bold"))
        if target.title_page_author_centered is not None \
                and prof.tp_author_alignment is not None:
            _w2 = "center" if target.title_page_author_centered else "left"
            if prof.tp_author_alignment != _w2:
                _tp_bits.append(("作者行 " if L == "zh" else "author lines ")
                                + "%s → %s" % (_align_text(L, prof.tp_author_alignment),
                                                _align_text(L, _w2)))
        if _tp_bits:
            sep = "：" if L == "zh" else ": "
            lines.append("%s%s%s" % (_dim_label(L, "title_page"), sep,
                                     "，".join(_tp_bits) if L == "zh" else ", ".join(_tp_bits)))
    if _gate("block_quote"):
        # 与 checker 的判定同一口径（含引号 + 长 + 左缩进不是 0.5″）
        _bq = [c for c in (getattr(prof, "block_quote_candidates", None) or [])
               if c.get("has_quote") and (c.get("left_indent_in") is None
                                          or abs((c.get("left_indent_in") or 0) - 0.5) > 0.15)]
        if _bq:
            lines.append("%s%s%s" % (_dim_label(L, "block_quote"),
                                     "：" if L == "zh" else ": ",
                                     "%d 处缩进 0.5″" % len(_bq) if L == "zh"
                                     else 'indent 0.5" (%d paragraph(s))' % len(_bq)))

    # —— 参考文献：只列实际会变的项 ——
    # v2.3.65：不再以 reference_style 为总闸（学校模板可不选风格，只给条目字体/行距等）
    if _gate("reference_hanging_indent"):
        want_hang = target.reference_hanging_indent_in
        cur_hang = prof.reference_hanging_indent_in
        if want_hang is not None:
            need_hang = (cur_hang is None
                         or abs(float(cur_hang) - float(want_hang)) > 1e-6)
            if need_hang:
                cur_txt = _UNREC[L] if cur_hang is None else f"{cur_hang}\""
                lines.append(f"{_dim_label(L, 'ref_hanging')}: {cur_txt} → {want_hang}\"")
    if _gate("reference_numbered"):
        missing = [t for t in prof.reference_entries if t.strip()
                   and not _has_leading_number(t)]
        if missing:
            n = len(missing)
            ent = "entry" if n == 1 else "entries"
            prefix = _dim_label(L, "ref_numbering")
            if L == "zh":
                lines.append(
                    f"{prefix}：为 {n} 条未编号条目补 [n] 顺序编码"
                    f"（{target.reference_style}）")
            else:
                lines.append(
                    f"{prefix}: add [n] sequential code to {n} "
                    f"unnumbered {ent} ({target.reference_style})")

    # —— 参考文献条目字体 / 字号 / 行距 + 标题另起一页 / 加粗 / 居中（v2.3.65）——
    if _gate("reference_font") and target.reference_font_family:
        cur_f = prof.reference_font_family
        if (cur_f or "").lower() != str(target.reference_font_family).lower():
            cur_disp = _UNKNOWN[L] if not cur_f else cur_f
            lines.append("%s: %s → %s" % (_dim_label(L, "ref_font"), cur_disp,
                                          target.reference_font_family))
    if _gate("reference_size"):
        line = _diff_line(_dim_label(L, "ref_size"), prof.reference_size_pt,
                          target.reference_size_pt, "pt", lang)
        if line:
            lines.append(line)
    if _gate("reference_spacing"):
        line = _diff_line(_dim_label(L, "ref_spacing"), prof.reference_line_spacing,
                          target.reference_line_spacing, "x", lang)
        if line:
            lines.append(line)
    if _gate("reference_new_page") and target.reference_new_page is not None:
        cur_np = prof.reference_new_page
        if cur_np is None:
            lines.append("%s: %s" % (_dim_label(L, "ref_new_page"),
                                     "将设为另起一页" if L == "zh" else "set to start on a new page"))
        elif bool(cur_np) != bool(target.reference_new_page):
            if L == "zh":
                lines.append("%s：%s" % (_dim_label(L, "ref_new_page"),
                                         "上一页接排 → 另起一页" if target.reference_new_page
                                         else "另起一页 → 接排"))
            else:
                lines.append("%s: %s" % (_dim_label(L, "ref_new_page"),
                                         "on the same page → new page" if target.reference_new_page
                                         else "new page → same page"))
    if _gate("reference_heading"):
        _h_bits: list = []
        if target.reference_heading_bold is not None \
                and bool(prof.reference_heading_bold) != bool(target.reference_heading_bold):
            if L == "zh":
                _h_bits.append("加粗" if target.reference_heading_bold else "不加粗")
            else:
                _h_bits.append("bold" if target.reference_heading_bold else "not bold")
        if target.reference_heading_centered is not None \
                and bool(prof.reference_heading_centered) != bool(target.reference_heading_centered):
            if L == "zh":
                _h_bits.append("居中" if target.reference_heading_centered else "左对齐")
            else:
                _h_bits.append("centered" if target.reference_heading_centered
                               else "left-aligned")
        if _h_bits:
            lines.append("%s%s%s" % (_dim_label(L, "ref_heading"),
                                     "：" if L == "zh" else ": ",
                                     "、".join(_h_bits) if L == "zh" else ", ".join(_h_bits)))

    # —— 每级标题的字体 / 字号 / 对齐 / 加粗 / 斜体 / 缩进（v2.3.65）——
    # 为什么必须列进预览：``compute_changes`` 是「要不要出副本、要不要扣次」的唯一依据。
    # 只改每级标题参数时若这里不列，App 会判「文档已符合」→ 压根不跑修正，
    # 客户填了一堆参数却什么也没发生（等于假功能）。
    # v2.3.70：与 fix_docx 共用 _heading_scope —— 检查只报了三级，就只列三级的改动行。
    # v2.3.71：再与 fix_docx 共用 _heading_problems_reported —— 检查没报标题问题就一行都不列
    # （否则会变成「预览说要改、修了却没改」的反向不一致）。
    if _gate("heading") and _heading_problems_reported(findings):
        lines.extend(_heading_change_lines(prof, target, L,
                                           only_levels=_heading_scope(findings)))

    # —— 参考文献条目内容改写（APA→IEEE）：这是**改内容**，必须在预览里写出来 ——
    # 客户改前有权知道「哪几条被重写了」，否则等于背着人改字（违反只改格式铁律的知情前提）。
    if _reference_problems_reported(findings) \
            and getattr(target, "reference_style", None) == "IEEE" \
            and getattr(target, "reference_rewrite", False):
        _rw = [e for e in (getattr(prof, "reference_entries", None) or [])
               if str(e).strip() and not _looks_like_ieee(str(e)) and _apa_to_ieee(str(e))]
        if _rw:
            _n = len(_rw)
            if L == "zh":
                lines.append("%s：按 IEEE 规范改写 %d 条条目"
                             "（作者缩写 / 题名加引号 / vol. no. pp. / 年份）"
                             % (_dim_label(L, "ref_entry"), _n))
            else:
                lines.append("%s: rewrite %d %s into IEEE style "
                             "(initials, quoted title, vol./no./pp., year)"
                             % (_dim_label(L, "ref_entry"), _n,
                                "entry" if _n == 1 else "entries"))

    return lines


def _heading_change_lines(prof, target: TargetProfile, L: str,
                          only_levels=None) -> list:
    """把「每级标题规则 vs 文档实测」的差异列成预览行（不再重复比一遍逻辑：

    与 checker 的新维度同源 —— 都走 ``resolve_heading_formats`` 取目标规则、
    都读 ``prof.heading_formats`` 的采样值。

    v2.3.70：``only_levels`` 为检查结论里报错的层级集合（None = 不限级）。
    与 ``fix_docx(only_levels=...)`` 同口径 —— 只给会被真正修改的层级出预览行，
    兑现「预览里写的就是 fixer 真会改的」这一承诺。
    """
    want_map = resolve_heading_formats(
        getattr(target, "heading_formats", None) or {},
        None,
        getattr(target, "heading_levels", None),
    )
    if not want_map:
        return []
    have_map = getattr(prof, "heading_formats", None) or {}
    out: list = []
    for lvl in sorted(want_map):
        if only_levels is not None and lvl not in only_levels:
            continue
        want = want_map[lvl]
        got = have_map.get(lvl)
        if got is None:
            continue
        bits: list = []

        def _num(field, cur, exp, unit=""):
            if exp is None or cur is None:
                return
            if abs(float(cur) - float(exp)) <= 1e-6:
                return
            bits.append("%g%s → %g%s" % (float(cur), unit, float(exp), unit))

        if want.font_family and got.font_family and \
                str(got.font_family).strip().lower() != str(want.font_family).strip().lower():
            bits.append("%s → %s" % (got.font_family, want.font_family))
        _num("size", got.size_pt, want.size_pt, "pt")
        if want.alignment:
            # 与 docx_reader 同口径：没写 w:jc 就是左对齐（None 不许当成「没差异」，
            # 否则 fixer 会改、预览却不列 → 客户被背着改）。
            _got_al = got.alignment or "left"
            if _got_al != want.alignment:
                bits.append("%s → %s" % (_align_text(L, _got_al),
                                         _align_text(L, want.alignment)))
        ind_want = want.indent_in if want.indent_in is not None else (
            0.5 if want.indent else None)
        _num("indent", got.indent_in, ind_want, '"')
        for want_b, got_b, on, off in (
                (want.bold, got.bold,
                 "加粗" if L == "zh" else "bold",
                 "取消加粗" if L == "zh" else "not bold"),
                (want.italic, got.italic,
                 "斜体" if L == "zh" else "italic",
                 "取消斜体" if L == "zh" else "not italic")):
            if want_b is None or got_b is None or bool(want_b) == bool(got_b):
                continue
            bits.append(on if want_b else off)
        # 段前分页（v2.3.67）：三态，want None=不规定，got None=读不到（跳过）
        if want.page_break_before is not None and got.page_break_before is not None \
                and bool(want.page_break_before) != bool(got.page_break_before):
            bits.append(("段前分页" if want.page_break_before else "取消段前分页")
                        if L == "zh" else
                        ("page break before" if want.page_break_before
                         else "no page break before"))
        # 🔴 v2.3.73：**编号与全大写是改文字**，预览必须写出来。
        # 旧版只比对齐/字体/字号，于是「加编号 + 全大写」这类最刺眼的改动预览里一片空白
        # （实测：客户只选参考文献 IEEE，点修正后标题直接变成 "A. INTRODUCTION"，预览 0 行）。
        # 检查器已按级报出编号缺失（Finding.message = "Level N heading format"），
        # 这里必须同一口径给出一行，兑现「预览 = 实际」。
        if want.number_prefix and got.number_prefix is None:
            bits.append(("加编号 " if L == "zh" else "add numbering ")
                        + (_NUMBER_PREFIX_HINT.get(want.number_prefix)
                           or str(want.number_prefix)))
        if want.uppercase and got.uppercase is False:
            bits.append("全大写" if L == "zh" else "UPPERCASE")
        if not bits:
            continue
        label = _dim_label(L, "heading_lvl") % lvl
        sep = "：" if L == "zh" else ": "
        out.append("%s%s%s" % (label, sep,
                               "，".join(bits) if L == "zh" else ", ".join(bits)))
    return out


def _notes(prof, lang="en") -> list:
    """非修正类提示（不算「会变的项」，不影响是否消耗试用）。文案随语言走。"""
    notes: list = []
    n = getattr(prof, "table_paragraph_count", 0) or 0
    if n:
        notes.append(_NOTES_TMPL[_lang_of(lang)].format(n=n))
    return notes


def compute_changes(input_path: str, target: TargetProfile, lang="en") -> list:
    """返回真正会发生的格式修正项（不含提示语、不含「无需修正」占位）。

    供调用方判断「文档是否已合规」——已合规时不应生成副本、不应消耗试用额度。
    单一真相源：内部用 ``run_check`` 取得 (prof, findings)，预览项完全由 findings 的
    fix_dims 派生，与 ``fix_docx`` 实际施加项严格一致。
    ``lang`` 决定返回文案的语言（"en" / "zh"）。

    raises: DocxReadError
    """
    prof, findings = run_check(input_path, target)
    return _compute_changes(prof, target, lang, findings=findings)


def preview_fix(input_path: str, target: TargetProfile, lang="en") -> list:
    """返回将进行的格式修正项（现状 → 目标），不修改任何文件。

    只列出「实际会发生变化」的项（由同一份 findings 派生，与 ``fix_docx`` 一致）；
    已经符合的项不再列出。文案随语言走。

    raises: DocxReadError（文档无法读取）
    """
    prof, findings = run_check(input_path, target)
    lines = _compute_changes(prof, target, lang, findings=findings)
    tail = _notes(prof, lang)
    if not lines:
        return [_NO_CHANGE[_lang_of(lang)]] + tail
    return lines + tail
