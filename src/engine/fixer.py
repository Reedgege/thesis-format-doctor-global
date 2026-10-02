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

from .checker import ALIGN_LABEL          # 对齐标签单一来源（避免两处硬编码）
from .docx_reader import (
    DocxReadError, heading_level, is_heading, read_docx, reference_paragraph_indices,
)
from .questionnaire import TargetProfile
from .specs import MARGIN_DIMS, get_spec, generic_heading_formats
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


def _select_paragraphs(doc, ref_set: set, include_reference: bool = True) -> list:
    """选取可被格式修正的段落：一律跳过标题段与域段落；可选跳过参考文献段。

    域段落（TOC/书签/交叉引用）通常不是标题样式，标题守卫拦不住，故单独排除；
    但参考文献条目本身可能含 URL 超链接（w:hyperlink），仍要参与格式修正，
    因此仅当该段不在 ref_set 时才跳过域段落。
    """
    out: list = []
    for i, para in enumerate(doc.paragraphs):
        if is_heading(para):
            continue
        if _para_has_field(para) and i not in ref_set:
            continue
        if not include_reference and i in ref_set:
            continue
        out.append(para)
    return out


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


def _insert_prefix_run(para, prefix: str) -> bool:
    """在段落**首位**插入一个纯文本 run，保留段落内原有的一切子元素。

    实现：先 add_run（追加到段末）拿到 w:r 元素，再把它搬到 w:pPr 之后的第一个位置。
    这样段落里原有的 w:drawing / w:footnoteReference / w:fldChar / w:br(type=page) /
    w:hyperlink 等都不受影响（Run.text setter 会 clear_content() 把它们删掉，禁用）。
    """
    try:
        new_run = para.add_run(prefix)
    except Exception:
        return False
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


def _apply_reference_format(doc, target: TargetProfile, ref_indices: list) -> dict:
    """参考文献：仅格式层。悬挂缩进 + （编号制时）补顺序前缀。

    不重写作者/年份/题名文本；条目数不变；补编号只插在段落首位，保留原格式与内联内容。
    """
    info = {"hanging_applied": False, "numbering_added": 0}
    if not ref_indices:
        return info
    entries = [doc.paragraphs[i] for i in ref_indices]

    hang = target.reference_hanging_indent_in
    if hang is not None:
        h = float(hang)
        for para in entries:
            pf = para.paragraph_format
            pf.left_indent = Inches(h)
            pf.first_line_indent = Inches(-h)
        info["hanging_applied"] = True

    if target.reference_numbered:
        info["numbering_added"] = _apply_reference_numbering(entries)
    return info


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


def _heading_comment_text(spec, lvl, f) -> str:
    name = _spec_label(spec)
    return f"{name} Level {lvl} heading: {f.desc}. Applied."


def _apply_heading_format(doc, target: TargetProfile, spec) -> dict:
    """按规范每级规则直接格式化标题段（不写 styleId）。

    - 仅处理 is_heading 且非域段落的段；
    - 层级由 heading_level() 判定（1-based）；规范未定义的更深层级回退到最深一级；
    - 每处理一段加一条英文批注说明规则；返回 {formatted, comments}。
    """
    info = {"formatted": 0, "comments": 0}
    if not getattr(target, "heading_levels", None):
        return info
    fmt_map = spec.heading_formats or generic_heading_formats(target.heading_levels)
    if not fmt_map:
        return info
    deepest = max(fmt_map.keys())
    page = target.page or {}
    body_font = page.get("font_family")
    body_size = page.get("font_size_pt")
    for para in doc.paragraphs:
        if not (is_heading(para) and not _para_has_field(para)):
            continue
        lvl = heading_level(para)
        if lvl is None:
            continue
        f = fmt_map.get(lvl) or fmt_map.get(deepest)
        if f is None:
            continue
        # 只对**真的被改过**的标题段加批注。原来 add_comment 写在 if 外面 —— 不管这段
        # 有没有改动都加一条写着 "Applied" 的批注：全合规的论文也会被塞满假批注，
        # 客户按批注去找改动必然扑空，正是「让客户知道改了哪儿」要的反效果。
        if _format_heading_one(para, f, body_font, body_size):
            info["formatted"] += 1
            add_comment(doc, para, _heading_comment_text(spec, lvl, f))
            info["comments"] += 1
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


def _build_summary_text(target, applied, spec, ref_info, prof) -> str:
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

    if not lines:
        return ""
    body = "\n".join("• " + ln for ln in lines)
    return (
        f"PaperFormat Pro ({name}) \u2014 Format adjustments applied to this document:\n"
        f"{body}\n\n"
        f"Scope: these changes apply to all body paragraphs. Headings, the reference "
        f"list, and tables are handled separately (see their own comments). No text "
        f"content was added, deleted, or rewritten."
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


def fix_docx(input_path: str, output_path: str, target: TargetProfile) -> dict:
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

    # 标题保护：no_heading 含参考文献；body_only 排除参考文献
    no_heading = _select_paragraphs(doc, ref_set, include_reference=True)
    body_only = _select_paragraphs(doc, ref_set, include_reference=False)

    summary: dict = {}
    applied: dict = {}

    page = target.page or {}
    if page:
        if _apply_margins(doc, page):
            applied["margins"] = page.get("margin_top_in")
        if _apply_font_and_size(page, no_heading):
            if page.get("font_family"):
                applied["font_family"] = page.get("font_family")
            if page.get("font_size_pt") is not None:
                applied["font_size_pt"] = page.get("font_size_pt")
        if _apply_line_spacing(page, no_heading):
            applied["line_spacing"] = page.get("line_spacing")
        if _apply_alignment(page, no_heading):
            applied["body_alignment"] = page.get("body_alignment")
        # 只把**真发生了改动**的项记进 applied（原来这两处不看返回值、无条件记录，
        # 汇总批注于是会声称"已设缩进/段后间距"，而原文可能本来就合规）
        if _apply_first_line_indent(page, body_only):
            applied["first_line_indent_in"] = page.get("first_line_indent_in")
        if _apply_space_after(page, body_only):
            applied["space_after_pt"] = page.get("space_after_pt")
        summary["page"] = dict(page)

    if target.reference_style:
        ref_info = _apply_reference_format(doc, target, ref_indices)
        summary["reference_style"] = target.reference_style
        summary["reference_hanging_indent_in"] = target.reference_hanging_indent_in
        summary["reference_numbered"] = target.reference_numbered
        summary["reference_numbering_added"] = ref_info["numbering_added"]
        summary["reference_entry_count"] = len(ref_indices)

    # 标题样式：直接格式化（不写 styleId）+ 逐段英文批注
    spec_key = getattr(target, "spec_key", None)
    spec = get_spec(spec_key) if spec_key else get_spec("Other")
    head_info = _apply_heading_format(doc, target, spec)
    summary["headings_formatted"] = head_info["formatted"]

    # 批注：文档首段 + 末段各一条「改动清单」汇总本遍全部改动；参考文献段加批注
    summary_text = _build_summary_text(
        target, applied, spec, ref_info if target.reference_style else None, prof
    )
    comments_added = head_info["comments"]
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
        raise OSError(f"无法写入输出文件：{e}") from e
    return summary


# ---------------------------------------------------------------------------
# 预览（现状 → 目标 diff，不落盘）
# ---------------------------------------------------------------------------

def _diff_line(label: str, cur, exp, unit: str) -> str | None:
    """生成一行 diff；无变化返回 None。"""
    if exp is None:
        return None
    if cur is None:
        return f"{label}：将设为 {exp}{unit}"
    try:
        same = abs(float(cur) - float(exp)) <= 1e-6
    except (TypeError, ValueError):
        same = str(cur) == str(exp)
    if same:
        return None
    return f"{label}：{cur}{unit} → {exp}{unit}"


def _compute_changes(prof, target: TargetProfile) -> list:
    """根据实测画像算出「真正会发生变化」的格式修正项（可能为空列表）。"""
    lines: list = []
    page = target.page or {}

    # —— 页边距 ——
    for key, side, label in _MARGIN_KEYS:
        if page.get(key) is not None:
            line = _diff_line(label, prof.margins.get(side), float(page[key]), "″")
            if line:
                lines.append(line)

    # —— 字体 / 字号 / 行距 ——
    if page.get("font_family"):
        cur_f = prof.font_family
        if (cur_f or "").lower() != str(page["font_family"]).lower():
            lines.append(f"正文字体：{cur_f or '未知'} → {page['font_family']}")
    line = _diff_line("正文字号", prof.font_size_pt, page.get("font_size_pt"), "pt")
    if line:
        lines.append(line)
    line = _diff_line("行距", prof.line_spacing, page.get("line_spacing"), "x")
    if line:
        lines.append(line)

    # —— 段落级 ——
    if page.get("body_alignment"):
        cur_a = prof.body_alignment
        if cur_a != page["body_alignment"]:
            lines.append("正文对齐：{} → {}".format(
                ALIGN_LABEL.get(cur_a, cur_a or "未知"),
                ALIGN_LABEL.get(str(page["body_alignment"]), page["body_alignment"])))
    line = _diff_line("正文首行缩进", prof.first_line_indent_in,
                      page.get("first_line_indent_in"), "″")
    if line:
        lines.append(line)
    line = _diff_line("段后间距", prof.space_after_pt, page.get("space_after_pt"), "pt")
    if line:
        lines.append(line)

    # —— 参考文献：只列实际会变的项 ——
    if target.reference_style:
        want_hang = target.reference_hanging_indent_in
        cur_hang = prof.reference_hanging_indent_in
        if want_hang is not None:
            need_hang = (cur_hang is None
                         or abs(float(cur_hang) - float(want_hang)) > 1e-6)
            if need_hang:
                cur_txt = "未识别" if cur_hang is None else f"{cur_hang}″"
                lines.append(f"参考文献悬挂缩进：{cur_txt} → {want_hang}″")
        if target.reference_numbered:
            missing = [t for t in prof.reference_entries if t.strip()
                       and not _has_leading_number(t)]
            if missing:
                lines.append(
                    f"参考文献编号：为 {len(missing)} 条未编号条目补 [n] 顺序编码"
                    f"（{target.reference_style}）")

    return lines


def _notes(prof) -> list:
    """非修正类提示（不算「会变的项」，不影响是否消耗试用）。"""
    notes: list = []
    n = getattr(prof, "table_paragraph_count", 0) or 0
    if n:
        notes.append(f"Note: text inside tables ({n} paragraph(s)) is not modified; only body paragraph formatting is fixed this run.")
    return notes


def compute_changes(input_path: str, target: TargetProfile) -> list:
    """返回真正会发生的格式修正项（不含提示语、不含「无需修正」占位）。

    供调用方判断「文档是否已合规」——已合规时不应生成副本、不应消耗试用额度。

    raises: DocxReadError
    """
    prof = read_docx(input_path)
    return _compute_changes(prof, target)


def preview_fix(input_path: str, target: TargetProfile) -> list:
    """返回将进行的格式修正项（现状 → 目标），不修改任何文件。

    只列出「实际会发生变化」的项；已经符合的项不再列出，避免误导。
    读取不到实测值时标注「将设为」。

    raises: DocxReadError（文档无法读取）
    """
    prof = read_docx(input_path)
    lines = _compute_changes(prof, target)
    tail = _notes(prof)
    if not lines:
        return ["当前文档已符合目标格式，无需修正。"] + tail
    return lines + tail
