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

from .checker import ALIGN_LABEL, run_check, Finding, FixAction  # 对齐标签单一来源（避免两处硬编码）
from .docx_reader import (
    DocxReadError, heading_level, is_heading, is_pseudo_heading, read_docx,
    reference_paragraph_indices, _max_run_size_pt, _PSEUDO_HEADING_FORMAT_MIN_PT,
)
from .questionnaire import TargetProfile
from .specs import MARGIN_DIMS, get_spec, resolve_spec, generic_heading_formats, HeadingLevelFormat, TitlePageFormat
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
                            fix_hang: bool = True, fix_number: bool = True) -> dict:
    """参考文献：格式层 + （IEEE 时）best-effort 引文格式重写。

    - 悬挂缩进 + （编号制时）补顺序前缀；
    - 当且仅当 reference_style == "IEEE" 且条目明显为 APA 作者-年格式、且为单 run 纯文本
      （无超链接/域，避免破坏内联 URL）时，best-effort 改写为 IEEE 编号引文
      （作者首字母缩写 + 题名引号 + 期刊 vol./no./pp. + 年）。幂等、保守、可回归；
      非 APA 模式或无法安全解析的条目一律不动其文本。
    """
    info = {"hanging_applied": False, "numbering_added": 0, "rewritten": 0}
    if not ref_indices:
        return info
    entries = [doc.paragraphs[i] for i in ref_indices]

    # 引文内容重写必须**显式开启**（target.reference_rewrite），不能只看
    # reference_style == "IEEE" 就改：产品铁律是「只改格式、不改内容」，而把
    # APA 条目改写成 IEEE 引文是内容层改动。默认关闭 = 参考文献逐字不动（只加 [n]
    # 前缀 + 悬挂缩进），这正是 tests/test_review_fixes.py P0-1 断言的不变量；
    # 只有走 target_from_spec("IEEE") 的真实产品路径才会打开（见 questionnaire.py）。
    if target.reference_style == "IEEE" and getattr(target, "reference_rewrite", False):
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
    return info


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

def _format_heading_one(para, f, body_font, body_size):
    """对单个标题段直接套格式（值落在段落 / run 属性，不引用任何样式 id）。

    返回是否实际发生了改动（用于幂等计数，避免「每遍都计 5 处改动」的假非零）。
    """
    changed = False
    pf = para.paragraph_format
    if f.alignment:
        target_al = _ALIGN_ENUM.get(f.alignment)
        if pf.alignment != target_al:
            pf.alignment = target_al
            changed = True
    # 显示型标题顶格（清首行缩进）；接排型(L4/L5)设首行缩进
    target_indent = Inches(0.5) if f.indent else None
    if (pf.first_line_indent is None) != (target_indent is None):
        try:
            pf.first_line_indent = target_indent
            changed = True
        except Exception:
            pass
    for run in para.runs:
        if _is_comment_run(run):
            continue          # 批注标记不是标题正文，跳过（否则第二遍恒显"已改"）
        if f.bold and run.font.bold is not True:
            run.font.bold = True
            changed = True
        if f.italic and run.font.italic is not True:
            run.font.italic = True
            changed = True
        if body_font:
            if run.font.name != body_font:
                try:
                    run.font.name = body_font
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
                    if rFonts.get(qn(a)) != body_font:
                        rFonts.set(qn(a), body_font)
                        changed = True
            except Exception:
                pass
        if f.size_pt and run.font.size is not None \
                and abs(float(run.font.size.pt) - float(f.size_pt)) > 1e-6:
            try:
                run.font.size = Pt(f.size_pt)
                changed = True
            except Exception:
                pass
    if f.run_in and _ensure_trailing_period(para):
        changed = True
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
    if not heading_levels:
        return out
    exclude_ps = exclude_ps or set()
    pseudo = []
    for para in doc.paragraphs:
        if _para_has_field(para) or para._p in exclude_ps:
            continue
        hl = heading_level(para)
        if hl is not None:
            out[para._p] = (hl, False)
        elif is_pseudo_heading(para, min_pt=_PSEUDO_HEADING_FORMAT_MIN_PT):
            pseudo.append(para)
    if pseudo:
        sizes = [round(s, 1) for p in pseudo
                 if (s := _max_run_size_pt(p)) is not None]
        tiers = sorted(set(sizes)) if sizes else []
        if len(tiers) >= 2:
            # 最大字号 → L1，次大 → L2 …
            tier_to_lvl = {sz: i + 1 for i, sz in enumerate(reversed(tiers))}
            for p in pseudo:
                s = _max_run_size_pt(p)
                out[p._p] = (tier_to_lvl.get(round(s, 1) if s is not None else tiers[0], 1), True)
        else:
            # 均匀字号：无法可靠分层，全部视为顶层（best-effort）
            for p in pseudo:
                out[p._p] = (1, True)
    return out


# 这些标题通常是论文后置部分，IEEE 不对其编号（但仍套格式：加粗/居中/全大写）
_NON_NUMBERED_HEADINGS = {
    "references", "reference", "bibliography", "works cited",
    "acknowledgments", "acknowledgements", "acknowledgment", "acknowledgement",
    "appendix", "appendices", "abstract", "keywords",
}


def _apply_heading_format(doc, target: TargetProfile, spec, level_map: dict | None = None) -> dict:
    """按规范每级规则直接格式化标题段（不写 styleId）。

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
    if not getattr(target, "heading_levels", None):
        return info
    fmt_map = spec.heading_formats or generic_heading_formats(target.heading_levels)
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
        f = fmt_map.get(lvl) or fmt_map.get(deepest)
        if f is None:
            continue
        # 后置部分（References/Appendix…）IEEE 不编号，仅套格式
        text_lc = (para.text or "").strip().lower()
        numbered = f.number_prefix and text_lc not in _NON_NUMBERED_HEADINGS
        prefix = None
        if numbered:
            counters[lvl] = counters.get(lvl, 0) + 1
            for deeper in [k for k in counters if k > lvl]:
                counters[deeper] = 0     # 回到更浅层 → 更深层编号归零
            prefix = _section_number_prefix(f.number_prefix, counters[lvl])
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
def _detect_title_page(doc):
    """返回 {'title': int, 'authors': list[int]} 或 None（未识别到标题页）。"""
    paras = doc.paragraphs
    idx0 = None
    for i, p in enumerate(paras):
        if (p.text or "").strip():
            idx0 = i
            break
    if idx0 is None:
        return None
    if is_heading(paras[idx0]):
        return None                      # 首段已是标题 → 交给 heading 流程，不重复处理
    # 正面信号：前 40 段内出现 Abstract / Keywords 标签（标题页论文的强特征）
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


def _apply_title_page_format(doc, target: TargetProfile, spec) -> dict:
    """对 APA/IEEE 等含标题页的论文：标题页标题 / 作者 / 单位做直接格式化。

    规则完全由 spec.title_page 驱动（各风格在 specs.py 内显式声明，不再靠
    「Level 1 是否加粗」这种脆弱启发）。spec.title_page 为 None（如 Other / 自定义模板）
    时整体跳过，交给模板驱动逻辑。复用 _format_heading_one 原语（绝不写 styleId、幂等）。
    返回 {title_formatted, comments}。
    """
    info = {"title_formatted": 0, "comments": 0}
    tp_spec = spec.title_page
    if tp_spec is None:
        return info
    tp = _detect_title_page(doc)
    if not tp:
        return info
    page = target.page or {}
    body_font = page.get("font_family")
    body_size = page.get("font_size_pt")

    title_align = "center" if tp_spec.title_centered else None
    title_bold = bool(tp_spec.title_bold)
    # 关键字传参：HeadingLevelFormat 第 7 个位置参数是 number_prefix（会真的往正文里
    # 插前缀），把描述串按位置传进去会把它当成编号前缀（实测曾埋雷）。
    title_fmt = HeadingLevelFormat(
        alignment=title_align,
        bold=title_bold,
        desc="centered%s title-page title" % (" + bold" if title_bold else ""),
    )
    p = doc.paragraphs[tp["title"]]
    changed = _format_heading_one(p, title_fmt, body_font, body_size)

    author_align = "center" if tp_spec.author_centered else None
    for ai in tp["authors"]:
        pa = doc.paragraphs[ai]
        afmt = HeadingLevelFormat(
            alignment=author_align,
            desc="centered author/affiliation",
        )
        if _format_heading_one(pa, afmt, body_font, body_size):
            changed = True

    if changed:
        info["title_formatted"] = 1
        info["comments"] = 1
        bits = []
        if tp_spec.title_centered:
            bits.append("title centered")
        if tp_spec.title_bold:
            bits.append("title bold")
        if tp_spec.author_centered:
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

    # 标题页段落（标题/作者/单位）：其排版由 spec.title_page 独占管辖，必须在标题
    # 判定**之前**先摘出来 —— 否则标题被标题页流程加粗后会反过来满足伪标题条件，
    # 第二遍被当成章节标题编号（实测把论文大标题编成 "I."，章节编号整体推一位）。
    tp = _detect_title_page(doc) if spec.title_page is not None else None
    tp_ps: set = set()
    if tp:
        for idx in [tp["title"], *tp["authors"]]:
            if 0 <= idx < len(doc.paragraphs):
                tp_ps.add(doc.paragraphs[idx]._p)

    # 标题判定只做一次，正文档保护与标题格式化共用同一份结果（见 _select_paragraphs
    # 文档字符串：两处判定不一致会导致标题被正文格式污染 + 「恒报已改」假非零）。
    # 仅当真的判出标题才覆盖：空集合会让 _select_paragraphs 跳过 is_heading 兜底，
    # 反而使真 Heading 样式段失去保护（被压成正文字号）。
    level_map = _assign_heading_levels(doc, target, exclude_ps=tp_ps)
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
    quote_ps = {p._p for p in body_only if _is_block_quote(p)}
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
        if ("font" in fix_dims or "font_size" in fix_dims) and _apply_font_and_size(page, no_heading):
            if page.get("font_family"):
                applied["font_family"] = page.get("font_family")
            if page.get("font_size_pt") is not None:
                applied["font_size_pt"] = page.get("font_size_pt")
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
        if quote_ps and _apply_block_quote_indent([p for p in body_only if p._p in quote_ps]):
            applied["block_quote_indent_in"] = 0.5
        summary["page"] = dict(page)

    # 栏数（结构性维度）：仅当 checker 绑定了 columns 修正时才施加（单一真相源）。
    # 它不属于 page 字典，故独立于上面的 if page: 块。
    if "columns" in fix_dims and _apply_columns(doc, target):
        applied["columns"] = target.columns

    if target.reference_style:
        ref_info = _apply_reference_format(
            doc, target, ref_indices,
            fix_hang="reference_hanging_indent" in fix_dims,
            fix_number="reference_numbered" in fix_dims,
        )
        summary["reference_style"] = target.reference_style
        summary["reference_hanging_indent_in"] = target.reference_hanging_indent_in
        summary["reference_numbered"] = target.reference_numbered
        summary["reference_numbering_added"] = ref_info["numbering_added"]
        summary["reference_entry_count"] = len(ref_indices)

    # 标题样式：直接格式化（不写 styleId）+ 逐段英文批注。
    # 仅当 checker 绑定了 heading 修正动作时才施加（与「检查→修正」单一真相源一致）。
    head_info = {"formatted": 0, "pseudo_formatted": 0, "hierarchy_inferred": 0, "comments": 0}
    if "heading" in fix_dims:
        head_info = _apply_heading_format(doc, target, spec, level_map=level_map)
    summary["headings_formatted"] = head_info["formatted"]
    summary["headings_pseudo_formatted"] = head_info["pseudo_formatted"]
    summary["headings_hierarchy_inferred"] = head_info["hierarchy_inferred"]

    # 标题页（title page）：APA/MLA 等标题页论文的标题/作者/单位直接格式化
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
        "ref_hanging": "Reference hanging indent", "ref_numbering": "Reference numbering",
    },
    "zh": {
        "margin_top": "上页边距", "margin_bottom": "下页边距",
        "margin_left": "左页边距", "margin_right": "右页边距",
        "font_family": "正文字体", "font_size": "正文字号",
        "line_spacing": "行距", "body_alignment": "正文对齐",
        "first_line_indent": "正文首行缩进", "space_after": "段后间距",
        "header_distance": "页眉距页顶",
        "columns": "栏数",
        "ref_hanging": "参考文献悬挂缩进", "ref_numbering": "参考文献编号",
    },
}
_ALIGN_ZH = {"left": "左对齐", "justify": "两端对齐", "center": "居中", "right": "右对齐"}
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

    # —— 参考文献：只列实际会变的项 ——
    if target.reference_style:
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

    return lines


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
