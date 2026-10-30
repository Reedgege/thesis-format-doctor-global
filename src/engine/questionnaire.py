"""格式问卷（Format Questionnaire）—— 四输入统一的「中间层」。

PaperFormat Pro 的三种目标来源，全部收敛到同一个 TargetProfile：
  A) 选中的内置规范（spec）
  B) 用户手填的格式问卷（questionnaire / user）
  C) 用户把学校模板丢给任意 AI，让 AI 按本 schema 填出的 JSON（ai_import）
  D) 解析学校 Word 模板得到的结构（school / template）

设计原则（产品铁律）：
- 页面排版（页边距/字体/字号/行距/标题层级）**以学校/问卷为准**；
  规范里的 page 值仅作「兜底参考」，检查时不强制 fail（checker 对 spec 来源的页面维度降级为 info）。
- 参考文献格式**以规范为准**，除非学校/问卷显式给出引用字段。
- 多通道可组合：build_target() 统一合并并记录冲突，避免 GUI/CLI 各写一套。
"""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field, is_dataclass
from typing import Any, Optional

from .specs import (get_spec, REFERENCE_STYLE_KEYS, REFERENCE_STYLE_CHOICES,
                    REFERENCE_STYLE_CUSTOM, SPEC_ORDER)
from .docx_reader import read_docx, PAPER_SIZES, identify_paper_size
from .comment_spec_en import parse_comments

# 页码位置可选值（**唯一真相源**）：与 specs.StyleSpec.page_number_position 同口径。
# 语义严格按「页眉 / 页脚」分开 —— 这是修好「页脚右对齐被误判符合 APA top-right」的关键：
#   top-right      = 页眉右上（APA 7th §2.1 / MLA 9th 都是这一条）
#   bottom-center  = 页脚居中（中文论文最常见的「第 X 页」位置）
#   bottom-right   = 页脚右侧
PAGE_NUMBER_POSITIONS = ("top-right", "bottom-center", "bottom-right")

# 位置 → (写进页眉还是页脚, 对齐)。fixer 与 checker 共用，避免两处各写一份判断。
PAGE_NUMBER_LAYOUT = {
    "top-right": ("header", "right"),
    "bottom-center": ("footer", "center"),
    "bottom-right": ("footer", "right"),
}

# 参考文献条目的「已带顺序编号」形态：[1] / (1) / 1. / 1)
# 与 checker._HAS_LEADING_NUMBER_RE、fixer._LEADING_NUMBER_RE 保持同一口径。
_LEADING_NUM_RE = re.compile(r"^\s*(?:\[\s*\d+\s*\]|\(\s*\d+\s*\)|\d+\s*[.)])\s")

# 模板实测与所选风格的悬挂缩进差异超过这个值才提示冲突（1/16 英寸内视为排版误差）
REF_HANG_CONFLICT_TOL = 0.06

# 来源标签映射（resolution 里缺省时按 TargetProfile.source 兜底；单一真相）
_SOURCE_ORIGIN = {
    "template": "school",
    "school": "school",
    # 模板批注：比「实测」更细的来源标签，但**归一后同样是「学校显式要求」**。
    # 🔴 必须映射：checker/fixer 判「这是不是客户明确要求」用的是归一后的
    # user/school/ai_import/spec 四个值；少了这一条，批注要求会被当成外人，
    # _fix_value 门控不通过 → 报告说可修、点修正却不动（同一类坑踩过）。
    "school-comment": "school",
    # 模板既没写批注、也量不出时的「通用兜底」值（目前只有表后间隙用到）：
    # 归一后同样是 school —— 兜底也要能修，但报告里要说清「模板没写，按通用处理」。
    "school-default": "school",    "ai_import": "ai_import",
    "ai": "ai_import",
    "user": "user",
    "questionnaire": "user",
    "spec": "spec",
}


def origin_of_source(tag: Optional[str]) -> Optional[str]:
    """把任意来源标签归一成 ``user`` / ``school`` / ``ai_import`` / ``spec``。

    🔴 为什么需要：``target_from_dict`` 的 resolution 记的是原始来源（"questionnaire"、
    "template"、"ai"），而 checker / fixer 判「这是不是客户明确要求」时用的是归一后的
    四个值。不做这一步，「客户手填」会被当成「不是显式」—— 填了等于没填。
    """
    if tag is None:
        return None
    return _SOURCE_ORIGIN.get(str(tag), str(tag))


# 模板批注的独立来源标签：比 "school"（实测）更细，供报告区分「批注说的」vs「量出来的」。
_COMMENT_ORIGIN = "school-comment"
# 模板沉默时的通用兜底来源（当前用于「表后间隙」）：报告要标明不是模板要求的。
_DEFAULT_ORIGIN = "school-default"

# TargetProfile 的合法字段名缓存（懒加载）：批注解析器将来多出字段时，
# 合并侧直接丢弃而不是让 ``TargetProfile(**extra)`` 抛 TypeError 把整个流程搞崩。
_TP_FIELD_NAMES_CACHE: set = set()


def _merge_comment_values(dst: dict, src: dict, resolution: dict, origin: str,
                          allowed: Optional[set] = None) -> list:
    """把模板批注解析值覆盖到「实测」字典上（**批注 > 实测**），返回被覆盖的键。

    只覆盖批注**真的提到**的字段；没提到的键保持实测值 —— 与国内版
    「批注 override、样式兜底」同一口径，绝不因为「模板有批注」就丢掉实测。

    ``allowed``: 允许的键集合。``page`` 字典的键（margin_* / font_* …）**不是**
    dataclass 字段，必须传 ``TargetProfile.PAGE_KEYS``，否则会被字段名过滤全筛掉
    （实测踩过：批注写 12pt，目标里仍是实测 11pt）。
    """
    if not src:
        return []
    global _TP_FIELD_NAMES_CACHE
    if allowed is None:
        if not _TP_FIELD_NAMES_CACHE:
            from dataclasses import fields as _dc_fields
            _TP_FIELD_NAMES_CACHE = {f.name for f in _dc_fields(TargetProfile)}
        allowed = _TP_FIELD_NAMES_CACHE
    applied = []
    for key, val in src.items():
        if val is None:
            continue
        if key not in allowed:
            continue                      # 解析器新增了 TargetProfile 不认识的键 → 丢弃
        dst[key] = val
        resolution[key] = origin
        applied.append(key)
    return applied

# 叠加顺序（FR2 明文优先级 template > ai > questionnaire）：
# _overlay_one 是「后写覆盖先写」，故按优先级**从低到高**依次叠加。
_CHANNEL_ORDER = ("questionnaire", "ai", "template")


# ---------------------------------------------------------------------------
# 问卷 schema：驱动「手填 UI」与「AI 填表 JSON」的字段定义。
# 每个字段：key / label / group / type / options / default / unit / help
#          （可选）range=(min,max) —— 供 UI 做越界提示，避免 UI 侧再抄一份范围表
# group 用于 UI 分组与报告维度归类。本列表为字段唯一来源。
# ---------------------------------------------------------------------------

QUESTIONNAIRE_SCHEMA: list[dict] = [
    # —— 页面排版 ——
    {"key": "margin_top_in", "label": "上页边距", "group": "page", "type": "float",
     "default": 1.0, "unit": "英寸", "range": (0.2, 3.0),
     "help": "顶部页边距，单位英寸（1 英寸≈2.54 厘米）"},
    {"key": "margin_bottom_in", "label": "下页边距", "group": "page", "type": "float",
     "default": 1.0, "unit": "英寸", "range": (0.2, 3.0), "help": "底部页边距"},
    {"key": "margin_left_in", "label": "左页边距", "group": "page", "type": "float",
     "default": 1.0, "unit": "英寸", "range": (0.2, 3.0),
     "help": "左侧页边距（装订侧常更大）"},
    {"key": "margin_right_in", "label": "右页边距", "group": "page", "type": "float",
     "default": 1.0, "unit": "英寸", "range": (0.2, 3.0), "help": "右侧页边距"},
    {"key": "font_family", "label": "正文字体", "group": "page", "type": "str",
     "default": "Times New Roman", "unit": "", "help": "如 Times New Roman / Arial / Calibri"},
    {"key": "font_size_pt", "label": "正文字号", "group": "page", "type": "float",
     "default": 12.0, "unit": "磅", "range": (6.0, 72.0), "help": "正文推荐 10–12 磅"},
    {"key": "line_spacing", "label": "行距", "group": "page", "type": "float",
     "default": 2.0, "unit": "倍", "range": (0.8, 4.0),
     "help": "2.0=双倍，1.5=一倍半，1.0=单倍"},
    {"key": "body_alignment", "label": "正文对齐", "group": "page", "type": "choice",
     "options": ["left", "justify", "center", "right"], "default": "left", "unit": "",
     "help": "left=左对齐（APA/MLA 等），justify=两端对齐（IEEE 等）"},
    {"key": "first_line_indent_in", "label": "正文首行缩进", "group": "page", "type": "float",
     "default": 0.5, "unit": "英寸", "range": (0.0, 2.0),
     "help": "常见 0.5 英寸；0 = 不缩进（IEEE）"},
    {"key": "space_before_pt", "label": "段前间距", "group": "page", "type": "float",
     "default": 0.0, "unit": "磅", "range": (0.0, 72.0),
     "help": "常见 0 = 段间不留空行；学校模板若规定「段前 6 磅」按模板填"},
    {"key": "space_after_pt", "label": "段后间距", "group": "page", "type": "float",
     "default": 0.0, "unit": "磅", "range": (0.0, 72.0),
     "help": "常见 0 = 段间不留空行，用首行缩进区分段落"},
    # 页眉距页顶：多数格式不规定，故默认留空（填写后才参与检查/修正，不臆造）。
    # MLA 9th 要求 ½″（0.5）；学校模板能读到时以模板为准。
    {"key": "header_distance_in", "label": "页眉距页顶", "group": "page", "type": "float",
     "default": "", "unit": "英寸", "range": (0.0, 3.0),
     "help": "页眉上边缘到页面顶部的距离；多数格式不规定，留空即不检查（MLA 9 = 0.5）"},
    # 纸张大小：**通用项，不属于任何引用规范**。北美（美/加）用 Letter、欧洲/中国/澳洲
    # 用 A4，同一篇 APA 论文换个国家交就不一样 → 必须让客户自己选/由学校模板实测。
    # 结构性页面维度，不进 PAGE_KEYS；复核弹窗用单选下拉。
    {"key": "paper_size", "label": "纸张大小", "group": "page", "type": "choice",
     "options": list(PAPER_SIZES), "default": "", "unit": "",
     "help": "Letter=8.5×11 英寸（北美）；A4=210×297 毫米（欧洲/中国）。留空即不检查"},

    # 页码位置（v2.3.66 起客户可填）：旧版只能由内置规范带，客户无处可填、也改不了。
    # 注意 top-right 指的是**页眉**右上，不是页脚右上 —— 中文论文的页脚居中要选 bottom-center。
    {"key": "page_number_position", "label": "页码位置", "group": "page", "type": "choice",
     "options": list(PAGE_NUMBER_POSITIONS), "default": "", "unit": "",
     "help": "top-right=页眉右上（APA/MLA）；bottom-center=页脚居中（中文论文常见）。留空即不检查"},

    # 标题页（论文大标题 / 作者 / 单位）：学校模板里写得最死的一条，以前客户无处可填。
    # 三态：是 / 否 / —（不规定）。学校模板能读到时自动填，读不到就留「—」。
    {"key": "title_page_title_centered", "label": "标题页标题居中", "group": "page",
     "type": "tri", "default": "", "unit": "",
     "help": "论文大标题是否居中；APA/IEEE 要求居中，MLA 亦居中。留「—」= 不规定"},
    {"key": "title_page_title_bold", "label": "标题页标题加粗", "group": "page",
     "type": "tri", "default": "", "unit": "",
     "help": "APA/IEEE 要求标题加粗；MLA 明确不加粗。留「—」= 不规定"},
    {"key": "title_page_author_centered", "label": "标题页作者/单位居中", "group": "page",
     "type": "tri", "default": "", "unit": "",
     "help": "姓名、院系、单位那几行是否居中（APA/IEEE 居中，MLA 不居中）"},
    # v2.3.76：标题页标题段的段前/段后。APA 要求「标题在页顶下 3–4 行」、中文封面
    # 常要求「标题上方留 N 磅 / 标题与作者之间空 M 磅」——此前只有居中/加粗，间距无处填。
    {"key": "title_page_space_before_pt", "label": "标题页标题段前", "group": "page",
     "type": "float", "default": "", "unit": "磅", "range": (0.0, 200.0),
     "help": "论文大标题上方的间距（磅）；留空=不规定。作者块与标题之间的距离用「标题段后」表达"},
    {"key": "title_page_space_after_pt", "label": "标题页标题段后", "group": "page",
     "type": "float", "default": "", "unit": "磅", "range": (0.0, 200.0),
     "help": "论文大标题下方的间距（磅，通常就是标题与作者块之间）；留空=不规定"},
    # v2.3.93：首页结构两项。v2.3.91/92 只让**内置规范**能表达（APA 学生版要求
    # 题名页独立成页、其余元素之间不留额外段距），学校模板 / 手填 / AI 三个通道
    # 根本无处填 —— 客户学校模板明确要求「封面单独一页」也表达不出来，
    # 传进 target_from_dict 还会被 _coerce 静默丢掉。现在三路齐平。
    {"key": "title_page_new_page", "label": "题名页独立成页", "group": "page",
     "type": "tri", "default": "", "unit": "",
     "help": "是=题名页之后强制分页（APA 学生版要求）；否=明确允许接排；—=不规定"},
    {"key": "title_page_zero_extra_spacing", "label": "题名页元素间不留额外段距",
     "group": "page", "type": "tri", "default": "", "unit": "",
     "help": "是=作者/单位/课程/教师/日期各行之间段前段后清 0（APA 学生版要求）；"
             "否=明确不要求；—=不规定"},
    # v2.3.114：标题字号 / 字体（与「学校模板批注」同一组 KEY，三路齐平）。
    # 此前只有模板批注能读，AI 导入 / 手填都无处填 → 被 _coerce 静默丢掉。
    {"key": "title_page_title_size_pt", "label": "标题页标题字号", "group": "page",
     "type": "float", "default": "", "unit": "磅", "range": (6.0, 120.0),
     "help": "论文大标题的字号（磅）；留空=不规定，跟随模板批注 / 规范"},
    {"key": "title_page_title_font", "label": "标题页标题字体", "group": "page",
     "type": "str", "default": "Times New Roman", "unit": "",
     "help": "论文大标题字体，如 Times New Roman / Arial / Calibri；留空=不规定"},

    # 栏数：单栏=1 / 双栏=2。结构性页面维度，不进 PAGE_KEYS；复核弹窗用单选下拉。
    # 标签双语：zh 走此处的 label，en 走 i18n.Q_LABELS_EN（英文界面绝不泄漏中文）。
    {"key": "columns", "label": "栏数", "group": "page", "type": "int",
     "default": 1, "unit": "", "range": (1, 4),
     "help": "1=单栏（多数格式）；2=双栏（IEEE 会议/期刊）"},

    # —— 表格（v2.3.112 起独立成组）——
    # 🔴 v2.3.112：这 6 个字段原先散落在「页面排版」和「参考文献」两块里 —— 客户想在
    # 手动表单里找「三线表」得先点开「参考文献」，找「表格间距」得翻「页面排版」，
    # 全都找不到。表格类单独立组，客户找得到才是好参数。group 只驱动 UI 分组，
    # 引擎的检查/修正/合并一律不看 group，改分组对逻辑零影响。
    #
    # 表格线（三线表）放在组内第一位 —— 它是最常被学校点名的一条，先进眼帘。
    # 学校模板里「表格用三线表，不要竖线」是高频硬要求。留空=不检查不改动（模板驱动）。
    {"key": "table_lines", "label": "表格线", "group": "table", "type": "choice",
     "options": ["three_line"], "default": "", "unit": "",
     "help": "three_line=学术三线表（顶线/表头下线/底线，去掉四边竖线与内部横线）；留空=不检查"},
    # 表格里「不要多余空行」是学校硬要求里很常见的一条；表格段落的段前段后与正文
    # 是两套口径（表格通常 0、正文常有 6–12 磅），必须能分开填。
    {"key": "table_space_before_pt", "label": "表格内段落段前", "group": "table",
     "type": "float", "default": "", "unit": "磅", "range": (0.0, 72.0),
     "help": "表格单元格里每个段落的段前间距；常见 0（表格内不留空行）；留空=不规定"},
    {"key": "table_space_after_pt", "label": "表格内段落段后", "group": "table",
     "type": "float", "default": "", "unit": "磅", "range": (0.0, 72.0),
     "help": "表格单元格里每个段落的段后间距；常见 0；留空=不规定"},
    # v2.3.98：表格与**前后文字**之间的空隙（不是表格内部）。
    # APA 7 原文：「嵌入正文的表格要与任何文字之间额外空一条双倍行距」→ 12pt 正文 = 24 磅。
    {"key": "table_gap_pt", "label": "表格与正文的间距", "group": "table",
     "type": "float", "default": "", "unit": "磅", "range": (0.0, 200.0),
     "help": "表格上方/下方与文字之间应留的空隙；APA=24 磅（12pt 正文的一条双倍行距）；留空=不规定"},
    # v2.3.100（裁定 C-2）：table_gap 的硬下限与推荐比例（三档阈值的另两个参数）。
    {"key": "table_gap_min_pt", "label": "表格间距硬下限", "group": "table",
     "type": "float", "default": "", "unit": "磅", "range": (0.0, 200.0),
     "help": "低于此值才判定为不合格；APA=12 磅（官方样例实测 12 磅按此放行）；留空=只按推荐值判"},
    {"key": "table_gap_ratio", "label": "推荐值达成比例", "group": "table",
     "type": "float", "default": "", "range": (0.0, 1.0),
     "help": "实测达到推荐值的多少比例即算合格；APA=0.8（24 磅 × 0.8 = 19.2 磅）；留空=不判比例"},

    # —— 目录（TOC，v2.3.77）——
    # 🔴 目录条目是**域段落**：Word 更新目录(F9)时会用 TOC 1..9 样式整体重排，
    # 逐条改直接格式会被冲掉。所以只有「目录标题段」可修；条目只提示。
    {"key": "toc_heading_space_before_pt", "label": "目录标题段前", "group": "page",
     "type": "float", "default": "", "unit": "磅", "range": (0.0, 200.0),
     "help": "「目录 / Contents」这一段的段前间距；留空=不规定"},
    {"key": "toc_heading_space_after_pt", "label": "目录标题段后", "group": "page",
     "type": "float", "default": "", "unit": "磅", "range": (0.0, 200.0),
     "help": "「目录 / Contents」这一段的段后间距；留空=不规定"},
    {"key": "toc_entry_space_before_pt", "label": "目录条目段前（仅提示）", "group": "page",
     "type": "float", "default": "", "unit": "磅", "range": (0.0, 72.0),
     "help": "目录条目之间的段前间距。⚠️ 只核对、不自动修改：目录条目由 Word 生成，"
             "要改请改「TOC 1…9」样式的段前段后"},
    {"key": "toc_entry_space_after_pt", "label": "目录条目段后（仅提示）", "group": "page",
     "type": "float", "default": "", "unit": "磅", "range": (0.0, 72.0),
     "help": "目录条目之间的段后间距。⚠️ 只核对、不自动修改（同上，请改 TOC 样式）"},

    # —— 标题层级 ——
    {"key": "heading_levels", "label": "标题层级数", "group": "heading", "type": "int",
     "default": 0, "unit": "级", "range": (0, 9), "help": "0=不强制；APA 常 5 级，其余多为学校定"},

    # —— 参考文献 ——
    # 🔴 v2.3.73（老板 2026-10-20 定的口径）：**参考文献不许被定死成某个规范**。
    # 客户手动输入 / AI 导入，正是因为有学校自己的要求或自己的个性化值；旧版默认 APA、
    # 只有三选一，等于替他拍板。现在默认「按我校要求（自定义）」：
    #   · 自定义 → 只执行他填/AI 读出的条目细则（字体/字号/行距/悬挂/是否编号/标题加粗居中），
    #     不拿 APA/MLA/IEEE 的引用体例去核条目内容；
    #   · 选了 APA/MLA/IEEE → **只对参考文献这一块**按该规范核对，
    #     页面与标题仍按他给的（规范基线才全篇按规范）。
    {"key": "reference_style", "label": "参考文献格式", "group": "reference", "type": "choice",
     "options": list(REFERENCE_STYLE_CHOICES), "default": REFERENCE_STYLE_CUSTOM,
     "unit": "",
     "help": "「按我校要求」=只按你填的条目细则执行；选 APA/MLA/IEEE 只核参考文献，不动页面与标题"},
    {"key": "reference_hanging_indent_in", "label": "参考文献悬挂缩进", "group": "reference",
     "type": "float", "default": 0.5, "unit": "英寸", "range": (0.0, 2.0),
     "help": "悬挂缩进，常见 0.5 英寸；IEEE 为 0"},
    {"key": "reference_numbered", "label": "是否顺序编码", "group": "reference", "type": "bool",
     "default": False, "unit": "", "help": "True=如 IEEE [1]；False=著者-出版年制"},

    # —— 参考文献：条目本身的排版（v2.3.65）——
    # 学校模板里这几项常与正文不同口径（正文小四双倍、条目五号单倍），必须单独可填。
    {"key": "reference_font_family", "label": "条目字体", "group": "reference", "type": "str",
     "default": "", "unit": "",
     "help": "留空=跟随正文；学校常写「参考文献用宋体/Times New Roman」"},
    {"key": "reference_size_pt", "label": "条目字号", "group": "reference", "type": "float",
     "default": "", "unit": "磅", "range": (6.0, 72.0),
     "help": "留空=跟随正文；五号≈10.5 磅，小四≈12 磅"},
    {"key": "reference_line_spacing", "label": "条目行距", "group": "reference", "type": "float",
     "default": "", "unit": "倍", "range": (0.8, 4.0),
     "help": "留空=跟随正文；1.0=单倍（多数学校条目要求单倍）"},
    # v2.3.75：条目之间的段前/段后。「条目之间不要空行」是学校硬要求里很常见的一条，
    # 此前只有行距，管不住「每条之间多出 6 磅」——检查也看不见。
    {"key": "reference_space_before_pt", "label": "条目段前", "group": "reference", "type": "float",
     "default": "", "unit": "磅", "range": (0.0, 72.0),
     "help": "常见 0=条目之间不空行；留空=不规定"},
    {"key": "reference_space_after_pt", "label": "条目段后", "group": "reference", "type": "float",
     "default": "", "unit": "磅", "range": (0.0, 72.0),
     "help": "常见 0=条目之间不空行；留空=不规定"},
    {"key": "reference_new_page", "label": "参考文献另起一页", "group": "reference", "type": "tri",
     "default": "", "unit": "", "help": "选「是」=标题前插分页；选「否」=明确不要求；留「—」=不规定"},
    {"key": "reference_heading_bold", "label": "参考文献标题加粗", "group": "reference", "type": "tri",
     "default": "", "unit": "", "help": "是=粗体；否=明确不加粗（MLA 要求不加粗）；—=不规定"},
    {"key": "reference_heading_centered", "label": "参考文献标题居中", "group": "reference", "type": "tri",
     "default": "", "unit": "", "help": "是=居中；否=明确左对齐（IEEE 左对齐）；—=不规定"},

    # —— 摘要页特化（v2.3.67）——
    # 摘要正文口径与正文相反（APA 摘要顶格不缩进、专业版另起一页），独立可填。
    {"key": "abstract_first_line_indent_in", "label": "摘要首行缩进", "group": "page", "type": "float",
     "default": "", "unit": "英寸", "range": (0.0, 2.0),
     "help": "摘要正文首行缩进；APA 摘要顶格=0（留空=不规定，跟随正文不检查）"},
    {"key": "abstract_new_page", "label": "摘要另起一页", "group": "page", "type": "tri",
     "default": "", "unit": "",
     "help": "是=摘要标题前插分页（APA 专业版要求）；否=明确不另起；—=不规定"},
    # v2.3.93：摘要之后正文是否另起一页。同样从「只有内置规范能表达」放开到三路齐平
    # （APA 学生版/专业版都要求正文另起一页；国内学校模板常写「摘要页后接正文」）。
    {"key": "text_new_page", "label": "正文另起一页", "group": "page", "type": "tri",
     "default": "", "unit": "",
     "help": "是=摘要之后正文另起一页（APA 要求）；否=明确允许摘要后接排；—=不规定"},
    # v2.3.97：摘要**标题**自身的格式。以前只有「摘要正文」的维度，标题这一行没人管
    # （APA 产物里它被当成二级章节标题套了左对齐）。三家口径差别大，故独立三问。
    {"key": "abstract_heading_alignment", "label": "摘要标题对齐", "group": "page",
     "type": "choice", "options": ["left", "center"], "default": "", "unit": "",
     "help": "center=居中（APA §2.9）；left=左对齐（多数工科模板）；—=不规定"},
    {"key": "abstract_heading_bold", "label": "摘要标题加粗", "group": "page", "type": "tri",
     "default": "", "unit": "",
     "help": "是=「Abstract」这一行加粗（APA 要求）；否=明确不加粗；—=不规定"},
    {"key": "abstract_heading_italic", "label": "摘要标题斜体", "group": "page", "type": "tri",
     "default": "", "unit": "",
     "help": "是=「Abstract」这一行斜体（部分 IEEE 模板）；否=明确不斜体；—=不规定"},
    # v2.3.97：关键词行（"Keywords: ..."）。APA §2.10 要标签斜体 + 整行缩进 0.5″；
    # 学校模板同样会写「关键词：小四、空两格」，所以三路输入都必须能表达这两项。
    {"key": "keywords_label_italic", "label": "关键词标签斜体", "group": "page", "type": "tri",
     "default": "", "unit": "",
     "help": "是=只把「Keywords:」这个标签设成斜体，词条不斜体（APA 要求）；否=明确不斜体；—=不规定"},
    {"key": "keywords_indent_in", "label": "关键词行缩进", "group": "page", "type": "float",
     "default": "", "unit": "英寸", "range": (0.0, 2.0),
     "help": "关键词行整行的首行缩进；APA=0.5（留空=不规定）"},
    # v2.3.76：摘要正文段的段前/段后（摘要上方与下方各留多少，学校模板常写死）。
    {"key": "abstract_space_before_pt", "label": "摘要正文段前", "group": "page", "type": "float",
     "default": "", "unit": "磅", "range": (0.0, 200.0),
     "help": "摘要正文段上方的间距（磅）；留空=不规定"},
    {"key": "abstract_space_after_pt", "label": "摘要正文段后", "group": "page", "type": "float",
     "default": "", "unit": "磅", "range": (0.0, 200.0),
     "help": "摘要正文段下方的间距（磅）；留空=不规定"},
    # v2.3.114：摘要正文字号 / 加粗 / 对齐（与「学校模板批注」同一组 KEY，三路齐平）。
    # 此前只有模板批注能读，AI 导入 / 手填都无处填 → 被 _coerce 静默丢掉。
    {"key": "abstract_font_size_pt", "label": "摘要正文字号", "group": "page",
     "type": "float", "default": "", "unit": "磅", "range": (6.0, 120.0),
     "help": "摘要正文字号（磅）；留空=不规定，跟随模板批注 / 规范"},
    {"key": "abstract_bold", "label": "摘要正文加粗", "group": "page", "type": "tri",
     "default": "", "unit": "",
     "help": "是=摘要正文加粗；否=明确不加粗；—=不规定（模板批注未提 bold 时按不加粗理解）"},
    {"key": "abstract_alignment", "label": "摘要正文对齐", "group": "page", "type": "choice",
     "options": ["left", "justify", "center", "right"], "default": "", "unit": "",
     "help": "摘要正文对齐方式；留空=不规定（跟随模板批注 / 规范）"},

    # —— Running head（v2.3.67）——
    {"key": "running_head_required", "label": "要求页眉 Running Head", "group": "page", "type": "tri",
     "default": "", "unit": "",
     "help": "是=要求页眉有 running head（短标题/姓氏）；APA 专业版要求、MLA 要求姓氏+页码；—=不规定"},
    {"key": "page_number_include_author_name", "label": "页码含作者姓氏", "group": "page", "type": "tri",
     "default": "", "unit": "",
     "help": "是=页码旁须带作者姓氏（MLA 的「Smith 1」形态）；—=不规定"},

    # —— 图表 caption（v2.3.67）——
    {"key": "caption_font_size_pt", "label": "图表标题字号", "group": "reference", "type": "float",
     "default": "", "unit": "磅", "range": (6.0, 72.0),
     "help": "图表标题（Caption）字号；留空=跟随正文。APA 表题常与正文同号或小一号"},
    {"key": "caption_line_spacing", "label": "图表标题行距", "group": "reference", "type": "float",
     "default": "", "unit": "倍", "range": (0.8, 4.0),
     "help": "图表标题行距；留空=跟随正文"},
    # v2.3.76：图表标题段的段前/段后。表题在表格上方、图题在图片下方，
    # 「题与图表之间留多少、题与正文之间留多少」是排版硬要求，此前完全管不了。
    {"key": "caption_space_before_pt", "label": "图表标题段前", "group": "reference",
     "type": "float", "default": "", "unit": "磅", "range": (0.0, 200.0),
     "help": "图表标题段上方的间距（磅）；留空=不规定"},
    {"key": "caption_space_after_pt", "label": "图表标题段后", "group": "reference",
     "type": "float", "default": "", "unit": "磅", "range": (0.0, 200.0),
     "help": "图表标题段下方的间距（磅）；留空=不规定"},
    {"key": "caption_position_table", "label": "表题位置", "group": "reference", "type": "choice",
     "options": ["top", "bottom"], "default": "", "unit": "",
     "help": "表题在表格上方(top)还是下方(bottom)；APA/IEEE 表题在上。留空=不检查"},
    {"key": "caption_position_figure", "label": "图题位置", "group": "reference", "type": "choice",
     "options": ["top", "bottom"], "default": "", "unit": "",
     "help": "图题在图上方(top)还是下方(bottom)；APA/IEEE 图题在下。留空=不检查"},
    # （v2.3.112：table_lines 已移到「表格」组，见本文件前面的 table 分组。）
    # v2.3.112：公式（Word 原生公式）。学校模板里「公式居中、编号加括号靠右、与正文同
    # 字号、段前段后 0」是高频硬要求，此前客户无处可填、引擎也**读不到** Word 公式。
    # 单独成组（equation），不与正文/表格混在一张表里 —— 公式的口径常与正文完全不同。
    {"key": "equation_alignment", "label": "公式对齐", "group": "equation", "type": "choice",
     "options": ["center", "left"], "default": "", "unit": "",
     "help": "center=公式在本栏内居中（多数学校要求）；left=左对齐；留空=不检查"},
    {"key": "equation_number_position", "label": "公式编号", "group": "equation",
     "type": "choice", "options": ["right", "none"], "default": "", "unit": "",
     "help": "right=编号加括号、与右边界齐平；none=不编号；留空=不检查"},
    {"key": "equation_size_pt", "label": "公式字号", "group": "equation", "type": "float",
     "range": (4.0, 72.0), "default": "", "unit": "pt",
     "help": "公式里的数学符号字号（常与正文同字号）；留空=不检查"},
    {"key": "equation_line_spacing", "label": "公式行距", "group": "equation", "type": "float",
     "range": (0.8, 3.0), "default": "", "unit": "倍",
     "help": "公式所在段落的行距倍数；留空=不检查"},
    {"key": "equation_space_before_pt", "label": "公式段前", "group": "equation", "type": "float",
     "range": (0.0, 200.0), "default": "", "unit": "pt",
     "help": "公式段段前间距（多数要求 0）；留空=不检查"},
    {"key": "equation_space_after_pt", "label": "公式段后", "group": "equation", "type": "float",
     "range": (0.0, 200.0), "default": "", "unit": "pt",
     "help": "公式段段后间距（多数要求 0）；留空=不检查"},
    # v2.3.112：公式字体（**只提示、绝不自动替换**）。在 Word 里换数学字体风险极高——
    # 字形映射一变符号可能渲染成空框，比不对齐严重得多，故引擎只如实告知、不代客户改。
    # 与另外两路同口径：模板批注(school-comment)本就能设，这里补齐手动问卷/AI 导入，
    # 让「学校要求公式用 Cambria Math」在三条输入通道都能表达（此前只有批注路径能设，
    # 手填表单不渲染、AI 导入经 _coerce 因不在 schema 被静默丢弃——与「模板驱动」铁律相悖）。
    {"key": "equation_font_family", "label": "公式字体", "group": "equation", "type": "str",
     "default": "", "unit": "",
     "help": "公式内数学符号的字体（如 Cambria Math）。⚠️ 只提示、绝不自动替换：在 Word 里"
             "换数学字体可能让符号渲染成空框。留空=不检查；学校要求时填字体名，"
             "修正后请在 Word 里手动设置（公式工具→转换）"},
]

# 数值字段的合理范围（由 schema 派生，避免 UI 侧再抄一份）
FIELD_RANGES: dict = {f["key"]: f["range"] for f in QUESTIONNAIRE_SCHEMA if f.get("range")}


@dataclass
class TargetProfile:
    """检查器消费的唯一目标格式画像。

    resolution: 维度 -> 来源标签，供对比报告标注「以规范为准 / 以学校模板为准」。
    conflicts: 显式来源之间互相冲突的维度记录（由 build_target 填充）。
    """

    source: str = "spec"                 # spec / school / user / ai_import 的组合
    spec_key: Optional[str] = None
    page: dict = field(default_factory=dict)        # 仅含已设置的页面值
    heading_levels: Optional[int] = None
    reference_style: Optional[str] = None
    reference_hanging_indent_in: Optional[float] = None
    reference_numbered: Optional[bool] = None
    # —— 参考文献细项（v2.3.65）——
    # 为什么单独加这一组：学校模板里「参考文献五号宋体、单倍行距、另起一页、标题加粗」
    # 是硬要求，而且口径常与正文不同。以前 target 只承载 style/悬挂/编号三项，
    # 模板读到了也无处可放 → 检查器按内置规范判、修正器按正文格式改，模板白读。
    # None = 不规定（不检查、不改动），与 heading_formats 的三态口径一致。
    reference_font_family: Optional[str] = None
    reference_size_pt: Optional[float] = None
    reference_line_spacing: Optional[float] = None
    reference_space_before_pt: Optional[float] = None   # 条目段前（磅，v2.3.75）
    reference_space_after_pt: Optional[float] = None    # 条目段后（磅，v2.3.75）
    reference_new_page: Optional[bool] = None
    reference_heading_bold: Optional[bool] = None
    reference_heading_centered: Optional[bool] = None
    # 分节栏数（单栏=1 / 双栏=2）。None = 不检查（Other/自定义且用户未声明）。
    # 属于**结构性页面维度**，与 page 里的标量分开：它不进 PAGE_KEYS（复核弹窗里用
    # 单选下拉而非文本输入），checker 也不对它做「spec→info」软参考降级。
    columns: Optional[int] = None
    # 纸张大小（Letter / A4 / Legal / A5）。None = 不检查（不规定）。
    # 与 columns 同属**结构性页面维度**：不进 PAGE_KEYS（复核弹窗用单选下拉而非文本框），
    # 也不吃 spec→info 降级判断里的「软参考」语气。它是**通用项**，不属于任何引用规范：
    # 北美 Letter、欧洲/中国 A4，同一份 APA 论文换个国家交就不一样。
    paper_size: Optional[str] = None
    # 页码位置（v2.3.66）：**客户可填** + 内置规范能带（APA/MLA = top-right）。
    # 取值见 ``PAGE_NUMBER_POSITIONS``；None = 不规定（不检查、不改动）。
    # 语义按「页眉/页脚」区分清楚：top-right = **页眉**右上（APA 7th / MLA 9th），
    # bottom-center / bottom-right = **页脚**居中/右侧（中文论文常规）。
    # 🔴 旧版 checker 只判「含页码的段落是否右对齐」，页脚右对齐也被判符合 top-right。
    page_number_position: Optional[str] = None
    # —— 标题页（title page）格式（v2.3.65）——
    # 为什么单独加：标题页（论文大标题 / 作者 / 单位）的「居中 / 加粗」是学校模板里
    # 写得最死的一条，以前客户**无处可填**，只能全靠内置规范；而且修正器里有一条
    # 不受检查驱动的流程会悄悄改它（客户看不到、报告里也没有）。
    # 三态：None = 不规定（不检查、不改动）/ True = 必须 / False = 必须不是。
    title_page_title_centered: Optional[bool] = None
    title_page_title_bold: Optional[bool] = None
    title_page_author_centered: Optional[bool] = None
    # v2.3.76：标题页**标题段**的段前 / 段后（磅）。None = 不规定。
    # 「标题离页顶多远」用段前表达，「标题与作者块之间」用段后表达（作者块本身
    # 不单独建字段：它跟随标题、且引擎只识别「标题 + 作者块」两段，多建会臆造）。
    title_page_space_before_pt: Optional[float] = None
    title_page_space_after_pt: Optional[float] = None
    # v2.3.114：题目字号 / 字体（模板驱动）。此前漏了这两个维度，模板批注写的
    # 「Times New Roman 24 pt」永远进不来画像，修正器只能套正文 12pt → 题目改不对。
    # None = 不规定（不检查、不改动）。
    title_page_title_size_pt: Optional[float] = None
    title_page_title_font: Optional[str] = None
    # v2.3.93：首页结构（题名页独立成页 / 其余元素零额外段距）。三态 None = 不规定。
    # 这两项 v2.3.91/92 只存在于内置规范（``StyleSpec``）里，三个显式通道无处表达，
    # 客户学校模板明确写了也带不出来 —— 与「模板驱动」铁律相悖，故补进画像。
    title_page_new_page: Optional[bool] = None
    title_page_zero_extra_spacing: Optional[bool] = None
    # —— 摘要页特化（v2.3.67）——
    # 摘要正文的口径与正文相反：APA 摘要不首行缩进（顶格）、专业版摘要另起一页。
    # None = 不规定（不检查、不改动）。这两个字段独立于正文 first_line_indent_in，
    # 因为「正文 0.5″ 首行缩进」绝不能刷到摘要段上。
    abstract_first_line_indent_in: Optional[float] = None
    abstract_new_page: Optional[bool] = None
    # v2.3.93：摘要之后正文是否另起一页（同上，从「只有内置规范能表达」放开）。
    text_new_page: Optional[bool] = None
    # v2.3.97：摘要**标题**（"Abstract" 这一行）自身的格式。老板 2026-10-23 打开 APA 产物
    # 一眼看出「摘要标题没居中」—— 此前完全没有这个维度，`Abstract` 还被伪标题兜底采样
    # 认成二级章节标题、套上 Level 2 的左对齐。三家口径差别大，故规范驱动、绝不一刀切：
    # APA 7 §2.9 = 居中 + 加粗；MLA 9 默认不要求摘要；IEEE 各模板差异过大。
    abstract_heading_alignment: Optional[str] = None
    abstract_heading_bold: Optional[bool] = None
    abstract_heading_italic: Optional[bool] = None
    # v2.3.97：关键词行（"Keywords: ..."）的标签斜体 + 整行首行缩进。
    # APA §2.10：标签斜体、整行缩进 0.5″；关键词小写/逗号分隔/末尾无句号属**内容**，只提示。
    keywords_label_italic: Optional[bool] = None
    keywords_indent_in: Optional[float] = None
    # v2.3.76：摘要正文段的段前/段后（磅）。None = 不规定（不检查、不改动）。
    abstract_space_before_pt: Optional[float] = None
    abstract_space_after_pt: Optional[float] = None
    # v2.3.114：摘要**正文**的字号 / 加粗 / 对齐（模板驱动）。此前漏了这三个维度，
    # 模板批注的「12 pt、不加粗」进不来，检查器转而套内置规范硬编码的 9pt 加粗。
    # None = 不规定（不检查、不改动）。
    abstract_font_size_pt: Optional[float] = None
    abstract_bold: Optional[bool] = None
    abstract_alignment: Optional[str] = None
    # —— Running head（v2.3.67）——
    # 页眉 running head 的两个显式开关。以前只能由内置规范（APA 专业版 / MLA）隐式带出，
    # 客户/学校想显式声明「要 running head」或「页码要含作者姓氏」无处可填。
    # running_head_required=True = 要求页眉有 running head（短标题或姓氏）；
    # page_number_include_author_name=True = 页码旁须带作者姓氏（MLA 的「Smith 1」）。
    # 文本内容（短标题 / 姓氏）因涉及「改内容」且无法可靠自动生成，仍只提示客户手动补；
    # 页码位置本身已由 page_number_position 维度可查可修。
    running_head_required: Optional[bool] = None
    page_number_include_author_name: Optional[bool] = None
    # —— 图表 caption（v2.3.67）——
    # 图表标题的字号/行距可修；表题/图题位置只能提示。None = 不规定（不检查、不改动）。
    caption_font_size_pt: Optional[float] = None
    caption_line_spacing: Optional[float] = None
    # v2.3.76：图表标题段的段前/段后（磅）。None = 不规定。
    caption_space_before_pt: Optional[float] = None
    caption_space_after_pt: Optional[float] = None
    caption_position_table: Optional[str] = None   # "top" / "bottom"
    caption_position_figure: Optional[str] = None  # "top" / "bottom"
    # —— 表格内段落（v2.3.77）——
    # 表格单元格段落的段前/段后。与正文分开承载：表格里通常 0、正文常有 6–12 磅，
    # 混成一个字段就会「改表格顺手把正文也改了」。None = 不规定。
    table_space_before_pt: Optional[float] = None
    table_space_after_pt: Optional[float] = None
    # —— 表格与**前后文字**之间的空隙（v2.3.98）——
    # 与上面两个（表格**单元格内部**的段距）语义完全不同：APA 7 要求嵌入正文的表格
    # 「与任何文字之间额外空一条双倍行距」，12pt 正文即 24pt。此前没有任何维度管它，
    # 而「正文段前段后应为 0」还会把原文的 14pt 清掉 → 表格紧贴标题（老板截图）。
    table_gap_pt: Optional[float] = None
    # v2.3.100（F-2 / 裁定 C-2）：table_gap 的三档阈值另两个参数（硬下限 / 推荐比例）。
    table_gap_min_pt: Optional[float] = None
    table_gap_ratio: Optional[float] = None
    # v2.3.100（F-5 / 裁定 C-6）：byline（作者/单位/城市/邮箱）字号，IEEE = 11pt。
    byline_size_pt: Optional[float] = None
    # —— 表格线（v2.3.85）——
    # "three_line" = 学术三线表（只留顶线 / 表头下线 / 底线，无竖线无内部横线）；
    # None = 不规定（不检查、不改动）。与 columns / paper_size 同属**结构维度**：
    # 不进 PAGE_KEYS（复核弹窗用单选下拉），checker 也不做 spec→info 降级。
    table_lines: Optional[str] = None
    # —— 公式（Word 原生 OMML，v2.3.112）——
    # OMML 的数学 run 是 ``m:r``，而 python-docx 的 ``para.runs`` 只返回 ``w:r`` ——
    # 公式段此前在引擎眼里是「除了编号一个字都没有的空段」，任何维度都判不到。
    # 与 table_lines 同属**元素级版式**：来源显式给出（模板批注 / 问卷 / AI）才硬判。
    # None 一律 = 不规定（不检查、不改动），绝不替客户拍板。
    equation_alignment: Optional[str] = None        # "center" / "left"
    equation_number_position: Optional[str] = None  # "right" / "none"
    equation_size_pt: Optional[float] = None
    equation_font_family: Optional[str] = None      # 只提示，不替换数学字体
    equation_line_spacing: Optional[float] = None
    equation_space_before_pt: Optional[float] = None
    equation_space_after_pt: Optional[float] = None
    # —— 目录（TOC，v2.3.77）——
    # 目录标题段可查可修；条目只核对、不自动改（条目是域段落，改直接格式会被
    # 「更新目录」冲掉，正确做法是改 TOC 1..9 样式）。None = 不规定。
    toc_heading_space_before_pt: Optional[float] = None
    toc_heading_space_after_pt: Optional[float] = None
    toc_entry_space_before_pt: Optional[float] = None
    toc_entry_space_after_pt: Optional[float] = None
    # 每级标题格式样本（学校模板驱动 / 问卷 / AI 显式给出）：level(int) -> dict
    # {alignment, bold, italic, indent, size_pt, number_prefix, uppercase}。
    # 仅当来源显式提供时才填入；其余入口留空、不臆造。
    heading_formats: dict = field(default_factory=dict)
    # 是否允许把参考文献**内容**改写为规范引文格式（best-effort，如 APA→IEEE）。
    # 默认 False：产品铁律是「只改格式、不改内容」，参考文献默认逐字不动（仅加 [n]
    # 前缀 + 悬挂缩进）。只有走 target_from_spec() 的真实产品路径、且该规范显式
    # 声明支持时才为 True —— 保证纯格式承诺在其它入口（含单测/单参调用）恒成立。
    reference_rewrite: bool = False
    resolution: dict = field(default_factory=dict)  # 维度 -> 来源
    conflicts: list = field(default_factory=list)   # 显式来源间冲突：{"dim","entries":[(label,val)]}
    notes: list = field(default_factory=list)       # 提示（如「学校模板未识别到字段，已回落规范」）

    PAGE_KEYS = ("margin_top_in", "margin_bottom_in", "margin_left_in",
                 "margin_right_in", "font_family", "font_size_pt", "line_spacing",
                 "body_alignment", "first_line_indent_in",
                 "space_before_pt", "space_after_pt",
                 "header_distance_in")

    # 标题页格式（不进 PAGE_KEYS：只作用于标题页那几行，与正文口径分开）
    TITLE_PAGE_KEYS = ("title_page_title_centered", "title_page_title_bold",
                       "title_page_author_centered",
                       # v2.3.76：标题页标题段的段前 / 段后
                       "title_page_space_before_pt", "title_page_space_after_pt",
                       # v2.3.114：题目字号 / 字体
                       "title_page_title_size_pt", "title_page_title_font",
                       # v2.3.93：首页结构两项（三路通道齐平，此前只有内置规范能表达）
                       "title_page_new_page", "title_page_zero_extra_spacing")

    # 参考文献细项（不进 PAGE_KEYS：它们只作用于参考文献段，与正文口径分开）
    REFERENCE_KEYS = ("reference_font_family", "reference_size_pt",
                      "reference_line_spacing",
                      # v2.3.75：条目之间的段前/段后（「条目之间不空行」是硬要求，
                      # 以前只有行距，管不住「每条之间多出 6 磅」）
                      "reference_space_before_pt", "reference_space_after_pt",
                      "reference_new_page",
                      "reference_heading_bold", "reference_heading_centered")

    # 摘要页特化（不进 PAGE_KEYS：它们只作用于摘要正文段，与正文口径分开）
    ABSTRACT_KEYS = ("abstract_first_line_indent_in", "abstract_new_page",
                     # v2.3.76：摘要正文段的段前 / 段后
                     "abstract_space_before_pt", "abstract_space_after_pt",
                     # v2.3.114：摘要正文字号 / 加粗 / 对齐（模板驱动）
                     "abstract_font_size_pt", "abstract_bold", "abstract_alignment",
                     # v2.3.93：摘要之后正文另起一页
                     "text_new_page",
                     # v2.3.97：摘要标题本身的对齐 / 加粗 / 斜体
                     "abstract_heading_alignment", "abstract_heading_bold",
                     "abstract_heading_italic",
                     # v2.3.97：关键词行标签斜体 / 整行缩进
                     "keywords_label_italic", "keywords_indent_in")

    # 表格内段落（v2.3.77；不进 PAGE_KEYS：只作用于表格单元格段落）
    # v2.3.98 起把 table_gap_pt（表格与前后**文字**之间的空隙）也放进来 ——
    # 三者都是「表格相关、数值型、只作用于表格」，合并/落盘/来源登记可共用一条路径。
    TABLE_KEYS = ("table_space_before_pt", "table_space_after_pt", "table_gap_pt")

    # 公式（v2.3.112；不进 PAGE_KEYS：它们只作用于含公式的段落，与正文口径分开）。
    # 与 TABLE_KEYS 同理 —— 都是「数值型、只作用于某类元素」，合并/落盘/来源登记共用一条路径。
    EQUATION_KEYS = ("equation_size_pt", "equation_line_spacing",
                     "equation_space_before_pt", "equation_space_after_pt")

    # 元素级**单值**维度（v2.3.112）：一个字段一个值、只作用于某类元素（表格 / 公式）。
    # 既不进 PAGE_KEYS（段落级数值），也不进 TABLE_KEYS（数值型间距）。
    # 🔴 新增这类字段时**必须**同时在 `_overlay_one` 里登记，否则四条通道一合并就被
    # 静默丢弃 —— `table_lines` 就是这么丢的（详见 _overlay_one 里的备注）。
    ELEMENT_KEYS = ("table_lines", "equation_alignment", "equation_number_position",
                    "equation_font_family") + EQUATION_KEYS

    # 目录（v2.3.77；不进 PAGE_KEYS：只作用于目录标题段 / 目录条目）
    TOC_KEYS = ("toc_heading_space_before_pt", "toc_heading_space_after_pt",
                "toc_entry_space_before_pt", "toc_entry_space_after_pt")

    # Running head（不进 PAGE_KEYS：页眉维度，文本内容只提示、页码位置由 page_number 管）
    RUNNING_HEAD_KEYS = ("running_head_required", "page_number_include_author_name")

    # 图表 caption（不进 PAGE_KEYS：图表标题维度，字号/行距可修、位置只提示）
    CAPTION_KEYS = ("caption_font_size_pt", "caption_line_spacing",
                    # v2.3.76：图表标题段的段前 / 段后
                    "caption_space_before_pt", "caption_space_after_pt",
                    "caption_position_table", "caption_position_figure")

    def to_dict(self) -> dict:
        return {
            "source": self.source,
            "spec_key": self.spec_key,
            "page": self.page,
            "heading_levels": self.heading_levels,
            "reference_style": self.reference_style,
            "reference_hanging_indent_in": self.reference_hanging_indent_in,
            "reference_numbered": self.reference_numbered,
            **{k: getattr(self, k) for k in self.REFERENCE_KEYS},
            "columns": self.columns,
            "paper_size": self.paper_size,
            "page_number_position": self.page_number_position,
            **{k: getattr(self, k) for k in self.TITLE_PAGE_KEYS},
            **{k: getattr(self, k) for k in self.ABSTRACT_KEYS},
            **{k: getattr(self, k) for k in self.RUNNING_HEAD_KEYS},
            **{k: getattr(self, k) for k in self.CAPTION_KEYS},
            "table_lines": self.table_lines,
            # v2.3.112：公式（对齐 / 编号 / 字号 / 行距 / 段距）
            "equation_alignment": self.equation_alignment,
            "equation_number_position": self.equation_number_position,
            "equation_font_family": self.equation_font_family,
            **{k: getattr(self, k) for k in self.EQUATION_KEYS},
            "heading_formats": {str(k): v for k, v in self.heading_formats.items()},
            "reference_rewrite": self.reference_rewrite,
            "resolution": self.resolution,
            "conflicts": self.conflicts,
            "notes": self.notes,
        }


def target_from_spec(key: str) -> TargetProfile:
    """仅用内置规范生成目标画像。

    页面值标 resolution="spec"（checker 会将其降级为 info，不强制 fail）；
    引用维度标 resolution="spec"（引用归规范，硬判）。
    """
    spec = get_spec(key)
    page = {k: v for k, v in spec.page.__dict__.items() if v is not None}
    ref = spec.reference
    resolution = {k: "spec" for k in page}
    if spec.heading_levels:
        resolution["heading_levels"] = "spec"
    resolution["reference_style"] = "spec"
    resolution["reference_hanging_indent_in"] = "spec"
    resolution["reference_numbered"] = "spec"
    if spec.columns is not None:
        resolution["columns"] = "spec"
    if spec.paper_size:
        resolution["paper_size"] = "spec"
    if spec.page_number_position:
        resolution["page_number_position"] = "spec"
    if spec.abstract_first_line_indent_in is not None:
        resolution["abstract_first_line_indent_in"] = "spec"
    if spec.abstract_new_page is not None:
        resolution["abstract_new_page"] = "spec"
    for _k in ("abstract_heading_alignment", "abstract_heading_bold",
               "abstract_heading_italic", "keywords_label_italic",
               "keywords_indent_in"):
        if getattr(spec, _k, None) is not None:
            resolution[_k] = "spec"
    if spec.caption_font_size_pt is not None:
        resolution["caption_font_size_pt"] = "spec"
    if spec.caption_line_spacing is not None:
        resolution["caption_line_spacing"] = "spec"
    if spec.caption_position_table is not None:
        resolution["caption_position_table"] = "spec"
    if spec.caption_position_figure is not None:
        resolution["caption_position_figure"] = "spec"
    # v2.3.85：表格线（三线表）。与栏数同为版式结构维度，spec 来源即显式意图。
    if getattr(spec, "table_lines", None):
        resolution["table_lines"] = "spec"
    if getattr(spec, "table_gap_pt", None) is not None:
        resolution["table_gap_pt"] = "spec"
    for _k in ("table_gap_min_pt", "table_gap_ratio", "byline_size_pt",
               "reference_size_pt"):
        if getattr(spec, _k, None) is not None:
            resolution[_k] = "spec"
    # v2.3.112：公式。内置规范当前**全部留 None**（理由见 specs.StyleSpec 注释：
    # APA/MLA 手册无此条款，IEEE 两档分别受「不碰版式」冻结裁定与「随各模板」约束）。
    # 这里照样把 resolution 与取值一起接上 —— 将来 spec 里填了值就直接生效，不必回来
    # 再补一遍（v2.3.100 就吃过「spec 明明写了、target 没带 → 维度整段不判」的亏）。
    _eq_spec_keys = ("equation_alignment", "equation_number_position", "equation_size_pt",
                     "equation_font_family", "equation_line_spacing",
                     "equation_space_before_pt", "equation_space_after_pt")
    for _k in _eq_spec_keys:
        if getattr(spec, _k, None) is not None:
            resolution[_k] = "spec"
    # v2.3.69：把规范的**每级标题格式**一并带进目标画像。
    # 此前只带了 heading_levels（层级数），heading_formats 漏传，造成两处脱节：
    #   ① 手填问卷选 APA/IEEE 后「每级标题明细」表全空 —— 客户以为规范没要求；
    #   ② 检查器的每级标题格式比对整段被跳过（checker 只在 heading_formats
    #      非空时才比），而修正器自己另读 spec.heading_formats 照改 —— 正是
    #      「检查说没事、修正却改一批」的脱节。
    # 转换无损：HeadingLevelFormat ↔ dict 由 asdict / heading_level_format_from_dict
    # 对称覆盖全部字段。MLA / Other 本就无多级标题体系（heading_levels=0），保持空。
    hf: dict = {}
    for _lv, _f in (spec.heading_formats or {}).items():
        try:
            hf[int(_lv)] = asdict(_f) if is_dataclass(_f) else dict(_f)
        except Exception:
            continue
    if hf:
        resolution["heading_formats"] = "spec"
    return TargetProfile(
        source="spec",
        spec_key=key,
        page=page,
        heading_levels=spec.heading_levels or None,
        heading_formats=hf,
        columns=spec.columns,
        paper_size=spec.paper_size,
        page_number_position=spec.page_number_position,
        abstract_first_line_indent_in=spec.abstract_first_line_indent_in,
        abstract_new_page=spec.abstract_new_page,
        abstract_heading_alignment=spec.abstract_heading_alignment,
        abstract_heading_bold=spec.abstract_heading_bold,
        abstract_heading_italic=spec.abstract_heading_italic,
        keywords_label_italic=getattr(spec, "keywords_label_italic", None),
        keywords_indent_in=getattr(spec, "keywords_indent_in", None),
        caption_font_size_pt=spec.caption_font_size_pt,
        caption_line_spacing=spec.caption_line_spacing,
        caption_position_table=spec.caption_position_table,
        caption_position_figure=spec.caption_position_figure,
        table_lines=getattr(spec, "table_lines", None),
        table_gap_pt=getattr(spec, "table_gap_pt", None),
        # v2.3.100（F-2 / F-5 / F-7）：这三个字段此前**没从 spec 带进 target**，
        # 于是 spec 里明明写了值、checker 读 target 却拿到 None → 维度整段不判
        # （"看起来改了、实际没生效"的又一例，四关之第③关）。
        table_gap_min_pt=getattr(spec, "table_gap_min_pt", None),
        table_gap_ratio=getattr(spec, "table_gap_ratio", None),
        # v2.3.112：公式（spec 侧当前恒 None，接线先备好，见上方说明）
        **{_k: getattr(spec, _k, None) for _k in _eq_spec_keys},
        byline_size_pt=getattr(spec, "byline_size_pt", None),
        # ⚠️ StyleSpec 的字段名是 ``reference_font_size_pt``（不是 reference_size_pt）——
        # 我第一次读错了名字，拿到 None，维度整段不判（"看起来改了、实际没生效"）。
        reference_size_pt=getattr(spec, "reference_font_size_pt", None),
        reference_style=ref.style_key,
        reference_hanging_indent_in=ref.hanging_indent_in,
        reference_numbered=ref.numbered,
        # 🔴 v2.3.84：**默认关闭**参考文献正文改写（老板 2026-10-21 更新清单第 5 条：
        # 「文字内容 100% 锁死，只改 Style 与 Layout」）。
        # 旧行为：规范选 IEEE 且顺序编码 → 自动把 APA 作者-年条目改写成 IEEE 引文
        #        （实测把 "Ivanov, D. (2020). Predicting…" 改成 "D. Ivanov, Predicting…"，
        #        等于改了客户的文字与标点，与第 5 条承诺直接冲突）。
        # 新行为：条目文字一个字都不动，只补 [n] 顺序前缀（列表标记）+ 悬挂缩进等格式；
        #        引用格式本身不符合规范时由检查器如实报告，交客户自行决定。
        # 需要旧行为的话可显式把 TargetProfile.reference_rewrite 置 True（内部/测试用）。
        reference_rewrite=False,
        resolution=resolution,
        # 把所选规范的「期望说明」（含学生/专业分层的前置材料差异）带入画像，
        # 供报告展示「正确说明」。空说明不塞，避免报告出现空行。
        notes=[spec.notes] if spec.notes else [],
    )


def _coerce(schema: dict, raw: dict) -> dict:
    """按 schema 把任意来源（手填/AI-JSON）的原始值做类型校正与裁剪。"""
    out: dict = {}
    valid = {f["key"]: f for f in schema}
    for k, f in valid.items():
        if k not in raw:
            continue
        v = raw[k]
        if v in (None, ""):
            continue
        t = f["type"]
        try:
            if t == "tri":
                # 三态开关：None/"" = 不规定（丢弃），True/False = 明确要求
                if isinstance(v, bool):
                    out[k] = v
                else:
                    s = str(v).strip().lower()
                    if s in ("1", "true", "yes", "y"):
                        out[k] = True
                    elif s in ("0", "false", "no", "n"):
                        out[k] = False
                    # 其它（含 "none"/"—"）= 不规定，直接丢弃
            elif t == "float":
                out[k] = float(v)
            elif t == "int":
                out[k] = int(v)
            elif t == "bool":
                out[k] = bool(v) if not isinstance(v, str) else v.strip().lower() in ("1", "true", "yes", "y")
            elif t == "choice":
                s = str(v).strip()
                opts = f.get("options") or []
                if opts:
                    # 大小写不敏感匹配，归一化到白名单原始写法；不在白名单则丢弃（不臆造值）
                    hit = next((o for o in opts if str(o).lower() == s.lower()), None)
                    if hit is None:
                        continue
                    out[k] = hit
                else:
                    out[k] = s
            else:
                out[k] = str(v).strip()
        except (ValueError, TypeError):
            continue
    return out


# 每级标题格式的字段白名单与取值归一（手填问卷 / AI 填表 / 学校模板三路共用一套）
HEADING_LEVEL_KEYS = ("alignment", "bold", "italic", "run_in", "indent", "indent_in",
                      "font_family", "size_pt", "number_prefix", "uppercase",
                      "page_break_before",
                      # v2.3.75：段前 / 段后（磅）与行距倍数（三态，空=不规定）
                      "space_before_pt", "space_after_pt", "line_spacing",
                      # v2.3.100（C 批 D-11 / 裁定 Q-1）：**小型大写**（三态）与
                      # **接排收尾符号**（"." 句点 / ":" 冒号）。
                      # ⚠️ 这两个字段不进白名单就会被**静默丢弃** —— 规范里写了、spec 里存了，
                      # 传到 target 时没了 → 修正器拿不到 → 行为退回旧默认（已踩过）。
                      "small_caps", "run_in_sep")

_HEADING_ALIGN = ("left", "center", "right", "justify")
_HEADING_PREFIX = ("number", "roman", "upper_letter", "lower_letter",
                   "lower_roman", "number_paren")


def _coerce_heading_level(raw: dict) -> dict:
    """把单级标题的原始值（手填文本框 / AI 的 YAML）归一成引擎认识的类型。

    只保留白名单字段；类型不对或不在允许值内的一律丢弃（不臆造、不猜）。
    """
    out: dict = {}
    for k, v in (raw or {}).items():
        if k not in HEADING_LEVEL_KEYS:
            continue
        if v is None or v == "":
            continue
        try:
            if k in ("bold", "italic", "run_in", "indent", "uppercase",
                     "page_break_before", "small_caps"):
                if isinstance(v, str):
                    out[k] = v.strip().lower() in ("1", "true", "yes", "y")
                else:
                    out[k] = bool(v)
            elif k == "run_in_sep":
                _sep = str(v).strip()
                if _sep in (".", ":"):
                    out[k] = _sep
            elif k in ("indent_in", "size_pt", "space_before_pt", "space_after_pt",
                       "line_spacing"):
                out[k] = float(v)
            elif k == "alignment":
                s = str(v).strip().lower()
                if s not in _HEADING_ALIGN:
                    continue
                out[k] = s
            elif k == "number_prefix":
                s = str(v).strip()
                if s.lower() in ("none", "no", "false"):
                    continue
                if s not in _HEADING_PREFIX:
                    continue
                out[k] = s
            else:  # font_family
                s = str(v).strip()
                if s:
                    out[k] = s
        except (ValueError, TypeError):
            continue
    return out


def target_from_dict(data: dict, source: str = "questionnaire") -> TargetProfile:
    """手填问卷 / AI 填表 JSON -> TargetProfile。

    修复点：
    - 兼容 AI 产出的 {"filled": {...}} / {"data": {...}} 等包装，不再静默丢弃；
    - 若解析结果为空（无任何有效字段），显式抛出 ValueError，而非返回空画像。
    - 仅给 reference_style 时，自动反填规范的悬挂缩进/编码方式，保持四通道结构一致。
    """
    raw = dict(data)
    # 解包常见 AI 包装
    for wrapper in ("filled", "data", "answers", "responses"):
        if wrapper in raw and isinstance(raw[wrapper], dict):
            raw = raw[wrapper]
            break
    # 兼容嵌套 page 结构
    if "page" in raw and isinstance(raw["page"], dict):
        for k, v in raw["page"].items():
            raw.setdefault(k, v)

    clean = _coerce(QUESTIONNAIRE_SCHEMA, raw)
    page = {k: clean[k] for k in TargetProfile.PAGE_KEYS if k in clean}
    # 参考文献细项（字体/字号/行距/另起一页/标题加粗居中）：与正文分开承载
    ref_extra = {k: clean[k] for k in TargetProfile.REFERENCE_KEYS if k in clean}
    heading = clean.get("heading_levels")
    cols = clean.get("columns")            # 栏数（int；不在 PAGE_KEYS 里，单独承载）
    paper = clean.get("paper_size")        # 纸型（Letter/A4/...；通用项，单独承载）
    page_num = clean.get("page_number_position")   # 页码位置（top-right/...；单独承载）
    tp_extra = {k: clean[k] for k in TargetProfile.TITLE_PAGE_KEYS if k in clean}
    abs_extra = {k: clean[k] for k in TargetProfile.ABSTRACT_KEYS if k in clean}
    rh_extra = {k: clean[k] for k in TargetProfile.RUNNING_HEAD_KEYS if k in clean}
    cap_extra = {k: clean[k] for k in TargetProfile.CAPTION_KEYS if k in clean}
    # v2.3.77：表格内段落 / 目录（都只在各自范围内生效，与正文口径分开）
    tbl_extra = {k: clean[k] for k in TargetProfile.TABLE_KEYS if k in clean}
    toc_extra = {k: clean[k] for k in TargetProfile.TOC_KEYS if k in clean}
    # v2.3.85：表格线（三线表）—— choice，不进 TABLE_KEYS（那里是数值型间距），单独承载。
    tbl_lines = clean.get("table_lines")
    if tbl_lines not in ("three_line", "boxed"):
        tbl_lines = None
    # v2.3.112：公式 —— 两个 choice + 一组数值（EQUATION_KEYS）+ 一个只提示的字体名。
    # 大小写归一与「不在白名单就丢弃」由 ``_coerce`` 负责（"Center" → "center"）；
    # 下面两道判断只是**兜底**：万一将来有调用方绕过 _coerce 直接构造，也不让非法值
    # 静默透传 —— 传进去的话 checker 判等永远不成立，客户填了既不报错也不生效。
    eq_extra = {k: clean[k] for k in TargetProfile.EQUATION_KEYS if k in clean}
    eq_align = clean.get("equation_alignment")
    if eq_align not in ("center", "left"):
        eq_align = None
    eq_num = clean.get("equation_number_position")
    if eq_num not in ("right", "none"):
        eq_num = None
    eq_font = clean.get("equation_font_family") or None
    ref_style = clean.get("reference_style")
    # 🔴 v2.3.73：「按我校要求（自定义）」= 没有选任何内置引用规范。
    # 归一成 None，后面所有「选了规范才做」的分支自然跳过（不核条目体例、不反填规范值）。
    if ref_style == REFERENCE_STYLE_CUSTOM:
        ref_style = None
    ref_hang = clean.get("reference_hanging_indent_in")
    ref_num = clean.get("reference_numbered")

    # 每级标题格式（问卷 / AI 显式给出时纳入；否则留空，不臆造）
    hf = raw.get("heading_formats") or raw.get("headingLevels")
    heading_formats: dict = {}
    if isinstance(hf, dict):
        for _k, _v in hf.items():
            try:
                _lvl = int(_k)
            except (ValueError, TypeError):
                continue
            if isinstance(_v, dict):
                clean_lvl = _coerce_heading_level(_v)
                if clean_lvl:
                    heading_formats[_lvl] = clean_lvl

    resolution = {k: source for k in page}
    if heading is not None:
        resolution["heading_levels"] = source
    if cols is not None:
        resolution["columns"] = source
    if paper is not None:
        resolution["paper_size"] = source
    if page_num is not None:
        resolution["page_number_position"] = source
    for _k in tp_extra:
        resolution[_k] = source
    for _k in abs_extra:
        resolution[_k] = source
    for _k in rh_extra:
        resolution[_k] = source
    for _k in cap_extra:
        resolution[_k] = source
    for _k in tbl_extra:
        resolution[_k] = source
    if tbl_lines is not None:
        resolution["table_lines"] = source
    # v2.3.112：公式
    for _k in eq_extra:
        resolution[_k] = source
    for _k, _v in (("equation_alignment", eq_align),
                   ("equation_number_position", eq_num),
                   ("equation_font_family", eq_font)):
        if _v is not None:
            resolution[_k] = source
    for _k in toc_extra:
        resolution[_k] = source
    if ref_style is not None:
        resolution["reference_style"] = source
    if ref_hang is not None:
        resolution["reference_hanging_indent_in"] = source
    if ref_num is not None:
        resolution["reference_numbered"] = source
    for _k in ref_extra:
        resolution[_k] = source

    # 仅给 reference_style 时，反填规范引用字段，保证结构完整
    if ref_style in REFERENCE_STYLE_KEYS:
        sp = get_spec(ref_style)
        if ref_hang is None:
            ref_hang = sp.reference.hanging_indent_in
            resolution["reference_hanging_indent_in"] = "spec"
        if ref_num is None:
            ref_num = sp.reference.numbered
            resolution["reference_numbered"] = "spec"

    # 每级标题明细（heading_formats）本身就是有效内容：只填了它也算「解析成功」，
    # 层级数由最深那一级反推（不要求客户重复填一遍）。
    if heading_formats:
        resolution["heading_formats"] = source
        if heading is None:
            heading = max(heading_formats)
            resolution["heading_levels"] = source

    # 空校验：避免静默得到空画像（只填了参考文献细项也算有效内容）
    if not page and ref_style is None and heading is None and cols is None \
            and paper is None and page_num is None and tbl_lines is None \
            and not ref_extra and not tp_extra and not abs_extra and not rh_extra \
            and not cap_extra and not tbl_extra and not toc_extra and not eq_extra \
            and eq_align is None and eq_num is None and eq_font is None:
        raise ValueError(
            "No valid format fields could be parsed from the input "
            "(margins, font, reference style, and heading levels are all empty). "
            "Check that the YAML/JSON is well-formed, or that it contains wrapper fields such as 'filled'."
        )

    # 🔴 v2.3.73：``spec_key`` 恒为 None —— 手动问卷/AI 导入里的「参考文献格式」
    # 只说明**参考文献这一块**按哪个规范核，不等于「全篇按该规范」。
    # 旧写法 `spec_key=ref_style` 会把客户手填的画像整体升级成规范基线，
    # 于是 resolve_spec() 返回整套规范规则（五级标题格式 / 标题页 / 页面），
    # 修正时把他的手排标题重写成 "A. INTRODUCTION"（老板 2026-10-20 明确否掉）。
    # 参考文献仍走 resolve_spec()（它按 reference_style 兜底取同一套规范），不受影响；
    # 页面/标题/标题页的「规范兜底」由 specs.spec_is_baseline() 一律关掉。
    return TargetProfile(
        source=source,
        spec_key=None,
        page=page,
        heading_levels=heading,
        columns=cols,
        paper_size=paper,
        page_number_position=page_num,
        table_lines=tbl_lines,
        # v2.3.112：公式（对齐 / 编号 / 字体只提示 / 字号·行距·段距走 EQUATION_KEYS）
        equation_alignment=eq_align,
        equation_number_position=eq_num,
        equation_font_family=eq_font,
        **eq_extra,
        **tp_extra,
        **abs_extra,
        **rh_extra,
        **cap_extra,
        **tbl_extra,
        **toc_extra,
        heading_formats=heading_formats,
        reference_style=ref_style,
        reference_hanging_indent_in=ref_hang,
        reference_numbered=ref_num,
        **ref_extra,
        resolution=resolution,
    )


def target_from_ai(path: str) -> TargetProfile:
    """导入 AI 填好的配置文件（YAML 或旧 JSON）。

    两种结构都支持：① 顶层扁平键值；② 包在 filled/data 等对象里。
    文件可以是 .yaml/.yml（推荐）或旧版 .json；按内容自动识别，用户无需关心扩展名。
    """
    import json
    try:
        from yaml import safe_load
    except ImportError:  # pragma: no cover - pyyaml 应随包分发
        safe_load = None
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()
    if not text.strip():
        raise ValueError("The AI config file is empty.")
    # JSON 是 YAML 的子集；先按 JSON 试（报错信息更精准），失败再回落 YAML。
    try:
        data = json.loads(text)
    except (ValueError, json.JSONDecodeError):
        if safe_load is None:
            raise ValueError("PyYAML is not installed and the YAML template cannot be parsed; use JSON instead, or install pyyaml.")
        data = safe_load(text)
    if not isinstance(data, dict):
        raise ValueError("The top level of the AI config file must be an object (key-value pairs); got: %s" % type(data).__name__)
    return target_from_dict(data, source="ai_import")


def target_from_template_docx(doc_profile, reference_style: Optional[str] = None) -> TargetProfile:
    """把学校模板 docx 的实测画像转成 school 目标（页面排版以学校为准）。

    无规范基线模式（产品铁律）：spec_key 恒为 None，不套任何内置规范的页面值，
    避免把规范值错标成「学校模板」、或抢占用户在高级面板里显式选的引用风格。
    引用风格由「学校模板」子面板的下拉框驱动（reference_style 参数）：
      - 留空（None）→ **按学校模板**：模板批注里写了 APA/MLA/IEEE 就按它核条目体例；
        模板也没写才不查流派。注意这**不等于「不检查参考文献」** —— 悬挂缩进 / 条目
        字体字号行距 / 编号 / 段距这些版式维度，空值时照样按模板实测值判定；
      - 选 APA/MLA/IEEE → 自动带入该风格的引用开关（悬挂缩进 / 编码方式，以及
        running head / 页码位置 / 参考文献另起页 / 字体白名单 / 标题页 / 标题加粗规则，
        这些开关由 checker 经 resolve_spec(target) 按 reference_style 取出，无需在此落盘）。
    """
    page: dict = {}
    notes: list = []          # 提示收集（纸型非标 / 模板与规范冲突…），全函数随时可用

    # —— 模板批注（Word comments）：海外「学校模板」模式的一手规范来源 ——
    # 优先级：**批注 > 实测**（与国内版「批注 > 样式定义」同一口径）。
    # 解析失败 / 无批注一律退化为 {} —— 读模板是锦上添花，绝不能让它把主流程搞崩。
    try:
        _cm = parse_comments(getattr(doc_profile, "template_comments", None) or [])
    except Exception:
        _cm = {}
    _CM = _COMMENT_ORIGIN
    if doc_profile.margins:
        for k_src, k_dst in (("top", "margin_top_in"), ("bottom", "margin_bottom_in"),
                             ("left", "margin_left_in"), ("right", "margin_right_in")):
            v = doc_profile.margins.get(k_src)
            if v is not None:
                page[k_dst] = round(float(v), 3)
    if doc_profile.font_family:
        page["font_family"] = doc_profile.font_family
    if doc_profile.font_size_pt:
        page["font_size_pt"] = round(float(doc_profile.font_size_pt), 1)
    if doc_profile.line_spacing:
        page["line_spacing"] = round(float(doc_profile.line_spacing), 2)
    if doc_profile.body_alignment:
        page["body_alignment"] = doc_profile.body_alignment
    if doc_profile.first_line_indent_in is not None:
        page["first_line_indent_in"] = round(float(doc_profile.first_line_indent_in), 3)
    if doc_profile.space_before_pt is not None:
        page["space_before_pt"] = round(float(doc_profile.space_before_pt), 1)
    if doc_profile.space_after_pt is not None:
        page["space_after_pt"] = round(float(doc_profile.space_after_pt), 1)
    # 页眉距页顶（MLA 9th 要求 ½″）：学校模板读出则硬判，读不到则留空（不臆造）
    if doc_profile.header_distance_in is not None:
        page["header_distance_in"] = round(float(doc_profile.header_distance_in), 3)
    if doc_profile.margins_inconsistent:
        # 多分节页边距不一致：页面维度不强行以模板为准，留给用户注意
        pass

    resolution = {k: "school" for k in page}

    # 栏数：学校模板实测值（多分节取众数；读不到则为 None，不臆造）——学校为准，硬判
    tpl_columns = getattr(doc_profile, "columns", None)
    if tpl_columns is not None:
        resolution["columns"] = "school"

    # 纸张大小：学校模板实测值（通用项，学校模板说了算）。读不到纸型但读到了实测尺寸
    # （比如学校用的是 8.5×13 非标纸）时，把实测尺寸写进 notes，绝不硬塞成 Letter/A4。
    tpl_paper = getattr(doc_profile, "paper_size", None)
    tpl_w = getattr(doc_profile, "paper_width_in", None)
    tpl_h = getattr(doc_profile, "paper_height_in", None)
    if tpl_paper:
        resolution["paper_size"] = "school"
    elif tpl_w and tpl_h:
        notes.append(
            "Paper size: the school template measures %.2f x %.2f in, which is not a standard "
            "Letter/A4/Legal/A5 size. Left unchecked — confirm the requirement with your school."
            % (float(tpl_w), float(tpl_h)))

    # 页码位置：模板里能直接看到 → 按实测「在页眉还是页脚 + 对齐」反推。
    # 学校模板是最强证据（比内置规范的典型值更该听），只做精确匹配，不猜：
    # 模板页脚居中就是 bottom-center，页眉右对齐就是 top-right；对齐读不准就不填。
    tpl_page_num: Optional[str] = None
    _pn_present = bool(getattr(doc_profile, "page_number_present", False))
    _pn_in = getattr(doc_profile, "page_number_in", None)
    _pn_al = getattr(doc_profile, "page_number_alignment", None)
    if _pn_present and _pn_in and _pn_al:
        tpl_page_num = {
            ("header", "right"): "top-right",
            ("footer", "center"): "bottom-center",
            ("footer", "right"): "bottom-right",
        }.get((_pn_in, _pn_al))
        if tpl_page_num:
            resolution["page_number_position"] = "school"
        else:
            notes.append(
                "Page number: the school template has a page number in the %s (%s-aligned), "
                "which does not match a standard position. Left unchecked — set it yourself "
                "if your school specifies one." % (_pn_in, _pn_al))

    # 每级标题格式样本（学校模板驱动）：层级 → dict
    heading_formats: dict = {}
    for lvl, sample in (doc_profile.heading_formats or {}).items():
        heading_formats[lvl] = sample.to_dict() if hasattr(sample, "to_dict") else dict(sample)

    # 标题页格式：模板实测（识别到标题页才写，读不到就不臆造）。
    # 学校模板里「大标题居中加粗、作者单位居中」是写得最死的一条，必须能带出来。
    tpl_tp: dict = {}
    if getattr(doc_profile, "title_page_detected", False):
        _tal = getattr(doc_profile, "tp_title_alignment", None)
        if _tal is not None:
            tpl_tp["title_page_title_centered"] = (_tal == "center")
        _tb = getattr(doc_profile, "tp_title_bold", None)
        if _tb is not None:
            tpl_tp["title_page_title_bold"] = bool(_tb)
        _aal = getattr(doc_profile, "tp_author_alignment", None)
        if _aal is not None:
            tpl_tp["title_page_author_centered"] = (_aal == "center")
        # 标题段的段前/段后（v2.3.76）：模板里常写「段前 1 行」，已由 reader 折算成磅值。
        for _key, _attr in (("title_page_space_before_pt", "tp_title_space_before_pt"),
                            ("title_page_space_after_pt", "tp_title_space_after_pt")):
            _val = getattr(doc_profile, _attr, None)
            if _val is not None:
                tpl_tp[_key] = round(float(_val), 1)
        # 首页结构（v2.3.93）：模板里题名页之后的分页是**看得见的事实**，
        # 「模板要求题名页独立成页」是它最强的证据，必须能带出来。
        _tp_np = getattr(doc_profile, "tp_next_break_before", None)
        if _tp_np is not None:
            tpl_tp["title_page_new_page"] = bool(_tp_np)
        # 「其余元素之间不留额外段距」只在模板**确实是 0** 时才声明；
        # 模板自己留着 6/12 磅 → 说明它不禁止，此时留「不规定」，绝不反向要求「必须有段距」。
        _tp_mx = getattr(doc_profile, "tp_author_space_max_pt", None)
        if _tp_mx is not None and float(_tp_mx) <= 0.5:
            tpl_tp["title_page_zero_extra_spacing"] = True
        for _k in tpl_tp:
            resolution[_k] = "school"

    # 摘要页特化：模板实测（识别到摘要段才写，读不到就不臆造）。
    tpl_abs: dict = {}
    _abs_fli = getattr(doc_profile, "abstract_first_line_indent_in", None)
    if _abs_fli is not None:
        tpl_abs["abstract_first_line_indent_in"] = round(float(_abs_fli), 3)
    _abs_np = getattr(doc_profile, "abstract_new_page", None)
    if _abs_np is not None:
        tpl_abs["abstract_new_page"] = bool(_abs_np)
    # 摘要正文段前/段后（v2.3.76）
    for _key in ("abstract_space_before_pt", "abstract_space_after_pt"):
        _val = getattr(doc_profile, _key, None)
        if _val is not None:
            tpl_abs[_key] = round(float(_val), 1)
    # 摘要之后正文是否另起一页（v2.3.93）：只在真的定位到摘要时才成立
    # （读不到摘要 = 无从判断正文首段在哪，绝不臆造）。
    if getattr(doc_profile, "has_abstract", False):
        _tx_np = getattr(doc_profile, "text_break_before", None)
        if _tx_np is not None:
            tpl_abs["text_new_page"] = bool(_tx_np)
    # 摘要标题格式（v2.3.97）：模板里 "Abstract" 这一行长什么样，就是它要求的样子。
    # 三项各自独立 —— 模板只写了居中、没写加粗，就只带出「居中」，绝不替它补加粗。
    _ah_al = getattr(doc_profile, "abstract_heading_alignment", None)
    if _ah_al is not None:
        tpl_abs["abstract_heading_alignment"] = _ah_al
    for _k, _attr in (("abstract_heading_bold", "abstract_heading_bold"),
                      ("abstract_heading_italic", "abstract_heading_italic")):
        _v = getattr(doc_profile, _attr, None)
        if _v is not None:
            tpl_abs[_k] = bool(_v)
    # 关键词行（v2.3.97）：模板里有关键词行才带出，没有就不规定。
    _kw_it = getattr(doc_profile, "keywords_label_italic", None)
    if _kw_it is not None:
        tpl_abs["keywords_label_italic"] = bool(_kw_it)
    _kw_in = getattr(doc_profile, "keywords_indent_in", None)
    if _kw_in is not None:
        tpl_abs["keywords_indent_in"] = round(float(_kw_in), 3)
    for _k in tpl_abs:
        resolution[_k] = "school"

    # 图表标题（Caption）：模板实测（模板里通常只给字号/行距/段前段后，位置另判）
    tpl_cap: dict = {}
    # caption 段前/段后实测为 0.0 仅表示它继承了正文样式的段间距，≠「模板要求必须是 0」。
    # 若模板带表格且触发 table_gap（表格与前后文字至少留 N pt 空隙），修正器会把这条
    # 空隙写进「表题段前」→ 与这里硬判的「0pt」互相打架，点完「全部修正」报告仍挂红项。
    # 故实测 0 视为「未规定」(None)，不写进 TargetProfile；显式批注仍照常覆盖（见下方
    # _merge_comment_values 合并，批注值会重新落进来）。仅放过「恰好=0」这一档，非零
    # 实测（如模板明确写了 6pt）照常带出。
    _CAP_SKIP_ZERO = ("caption_space_before_pt", "caption_space_after_pt")
    for _key, _rnd in (("caption_font_size_pt", 1),
                       ("caption_line_spacing", 2),
                       ("caption_space_before_pt", 1),
                       ("caption_space_after_pt", 1)):
        _val = getattr(doc_profile, _key, None)
        if _val is None:
            continue
        _val = round(float(_val), _rnd) if _rnd else float(_val)
        if _key in _CAP_SKIP_ZERO and _val == 0.0:
            continue
        tpl_cap[_key] = _val
        resolution[_key] = "school"

    # 表格内段落（v2.3.77）：模板里三线表常见「单元格内不空行」= 0 磅，必须带出来，
    # 否则客户拿模板选完，表格里多出来的空行永远指不出来。
    tpl_tbl: dict = {}
    for _key in TargetProfile.TABLE_KEYS:
        _val = getattr(doc_profile, _key, None)
        if _val is None:
            continue
        tpl_tbl[_key] = round(float(_val), 1)
        resolution[_key] = "school"

    # 表格与前后文字的间距（v2.3.98）：模板里表格前后实际留了多少，就照抄多少。
    # 取**最小值** —— 模板里每张表的留白可能被手调得不一样，最小那个才是它的下限；
    # 读到「真空段」的（blank=True）不计入 pt，另按「一行」折算不了就不写（宁可不规定）。
    _gap_vals = [g for _s in (getattr(doc_profile, "table_gap_samples", None) or [])
                 for g in (_s.get("before_pt"), _s.get("after_pt")) if g is not None]
    if _gap_vals:
        tpl_tbl["table_gap_pt"] = round(min(_gap_vals), 1)
        resolution["table_gap_pt"] = "school"

    # 目录（v2.3.77）：只有模板里真的带目录（TOC 条目）才写，读不到就不臆造。
    tpl_toc: dict = {}
    if getattr(doc_profile, "toc_detected", False):
        for _key in TargetProfile.TOC_KEYS:
            _val = getattr(doc_profile, _key, None)
            if _val is None:
                continue
            tpl_toc[_key] = round(float(_val), 1)
            resolution[_key] = "school"

    # 参考文献流派：**显式下拉 > 模板批注推断 > 不查**（老板 2026-10-29 拍板）。
    # 只从批注推断、不从实测条目形态猜流派 —— 实测猜流派不可靠。
    _cm_ref = (_cm.get("reference") or {})
    # 🔴 REFERENCE_STYLE_KEYS 是**大写**（"APA"/"MLA"/"IEEE"），解析器输出小写 → 必须归一，
    #    否则 `"apa" in ("APA",...)` 恒 False，批注里的流派永远被丢掉（实测踩过）。
    _cm_style = str(_cm_ref.get("reference_style") or "").strip().upper() or None
    # 客户是否**显式**在「参考文献格式」下拉里选了既定规范（≠ 由模板批注推断出来）。
    # v2.3.115 起这个区分决定编码方式的优先级，见下方 reference_numbered 分支。
    _explicit_style = reference_style in REFERENCE_STYLE_KEYS
    if reference_style in REFERENCE_STYLE_KEYS:
        ref_style = reference_style
    elif _cm_style in REFERENCE_STYLE_KEYS:
        ref_style = _cm_style
        resolution["reference_style"] = _CM
    else:
        ref_style = None
    ref_hang = None
    ref_num = None
    ref_extra: dict = {}
    spec_ref = get_spec(ref_style).reference if ref_style is not None else None

    # ① 悬挂缩进 / 编码方式：**模板实测优先**，模板读不到才用所选风格的规范值。
    # 旧写法直接用规范值（还标成 "school"）—— 模板写「悬挂 0.25″」、客户选了 APA（0.5″），
    # 检查器就按 0.5″ 判 fail、修正器把学生的 0.25″ 改成 0.5″，模板被自己的"规范"推翻。
    tpl_hang = getattr(doc_profile, "reference_hanging_indent_in", None)
    _cm_hang = _cm_ref.get("reference_hanging_indent_in")
    if _cm_hang is not None:
        # 批注 > 实测：批注写「悬挂 0.25″」就按 0.25″，不再用模板实际条目的缩进
        ref_hang = round(float(_cm_hang), 3)
        resolution["reference_hanging_indent_in"] = _CM
    elif tpl_hang is not None:
        ref_hang = round(float(tpl_hang), 3)
        resolution["reference_hanging_indent_in"] = "school"
        if spec_ref is not None and spec_ref.hanging_indent_in is not None \
                and abs(float(spec_ref.hanging_indent_in) - float(ref_hang)) > REF_HANG_CONFLICT_TOL:
            notes.append(
                "Reference hanging indent: the school template measures %.2f in, "
                "while %s requires %.2f in. The template value is used."
                % (float(ref_hang), ref_style,
                   float(spec_ref.hanging_indent_in)))
    elif spec_ref is not None:
        ref_hang = spec_ref.hanging_indent_in
        resolution["reference_hanging_indent_in"] = "spec"

    # 编码方式：条目实测是否带 [n]/1. 前缀（有条目才判，没条目不下结论）
    ref_entries = [t for t in (getattr(doc_profile, "reference_entries", None) or [])
                   if str(t).strip()]
    _cm_num = _cm_ref.get("reference_numbered")
    if _cm_num is not None:
        # 批注 > 实测：批注明确写了「编号 / 不编号」就照批注，不再拿条目形态反推
        ref_num = bool(_cm_num)
        resolution["reference_numbered"] = _CM
    elif _explicit_style and spec_ref is not None:
        # 🔴 v2.3.115（老板 2026-10-30 拍板）：「编号形态以学校模板为准；**客户显式选了
        # 既定规范**就按规范的」。旧行为让模板实测压过客户选择 —— 模板条目一旦读成
        # 「不编号」，就会吞掉客户选的 IEEE 编号，还会走 _strip_reference_numbering 把
        # 客户手打的 `1.` 编号整列剥掉（实测：老板把模板 [1] 的括号删成 1 之后，
        # 产物条目一条编号都不剩）。没选规范时仍以模板实测为准（见下面的分支）。
        ref_num = bool(spec_ref.numbered)
        resolution["reference_numbered"] = "spec"
    elif ref_entries:
        _numbered = sum(1 for t in ref_entries if _LEADING_NUM_RE.match(str(t).strip()))
        ref_num = _numbered >= max(1, len(ref_entries) // 2)      # 过半即视为编号制
        resolution["reference_numbered"] = "school"
        if spec_ref is not None and bool(spec_ref.numbered) != bool(ref_num):
            notes.append(
                "Reference numbering: the school template's entries are %s, while %s is %s. "
                "The template value is used."
                % ("numbered" if ref_num else "not numbered", ref_style,
                   "numbered" if spec_ref.numbered else "not numbered"))
    elif spec_ref is not None:
        ref_num = spec_ref.numbered
        resolution["reference_numbered"] = "spec"

    # ② 条目字体 / 字号 / 行距 / 段前段后：模板实测（模板里条目常与正文不同口径）
    for _key, _val, _round in (("reference_font_family",
                                getattr(doc_profile, "reference_font_family", None), None),
                               ("reference_size_pt",
                                getattr(doc_profile, "reference_size_pt", None), 1),
                               ("reference_line_spacing",
                                getattr(doc_profile, "reference_line_spacing", None), 2),
                               ("reference_space_before_pt",
                                getattr(doc_profile, "reference_space_before_pt", None), 1),
                               ("reference_space_after_pt",
                                getattr(doc_profile, "reference_space_after_pt", None), 1)):
        if _val is None or _val == "":
            continue
        ref_extra[_key] = round(float(_val), _round) if _round else str(_val)
        resolution[_key] = "school"

    # ③ 另起一页 / 标题加粗 / 标题居中：模板实测（仅在检测到参考文献段时成立）
    if getattr(doc_profile, "reference_new_page", None) is not None:
        ref_extra["reference_new_page"] = bool(doc_profile.reference_new_page)
        resolution["reference_new_page"] = "school"
    if ref_entries or getattr(doc_profile, "reference_heading_is_style", False):
        ref_extra["reference_heading_bold"] = bool(
            getattr(doc_profile, "reference_heading_bold", False))
        ref_extra["reference_heading_centered"] = bool(
            getattr(doc_profile, "reference_heading_centered", False))
        resolution["reference_heading_bold"] = "school"
        resolution["reference_heading_centered"] = "school"
        if getattr(doc_profile, "reference_heading_is_style", False):
            # 标题用 Heading 样式承载时，粗体/居中由样式给出，读不到直接值 → 不硬判
            ref_extra.pop("reference_heading_bold", None)
            ref_extra.pop("reference_heading_centered", None)
            resolution.pop("reference_heading_bold", None)
            resolution.pop("reference_heading_centered", None)

    if ref_style is not None:
        # setdefault：流派若是从**批注**推断出来的，上面已标了更细的 "school-comment"，
        # 这里不能把它覆盖回 "school"（否则报告分不清「批注说的」还是「量出来的」）。
        resolution.setdefault("reference_style", "school")

    # 层级数：模板用 Heading 样式时取样式层级；模板是「手排加粗标题」（无样式）时
    # heading_levels_used 为 0，需由伪标题样本反推 —— 否则 heading_levels=None 会让
    # 修正链在 _assign_heading_levels 提前返回，模板里明明读到的每级标题格式全部空转
    # （v2.3.65 实测：学校模板选完后点修正，标题一点没动）。
    _tpl_levels = doc_profile.heading_levels_used or (
        max(heading_formats) if heading_formats else None)
    if heading_formats:
        resolution["heading_formats"] = "school"
        if _tpl_levels:
            resolution["heading_levels"] = "school"

    # ↑ 以上全部是「实测」。下面按 **批注 > 实测** 覆盖批注真正提到的字段；
    #   没提到的键保持实测，绝不因为「模板有批注」就把实测结果丢掉。
    _cm_applied: list = []
    _cm_page = _cm.get("page") or {}
    # page 字典的键（margin_* / font_* …）**不是** dataclass 字段 → 必须显式传 PAGE_KEYS，
    # 否则会被字段名过滤整片筛掉（实测：批注写 12pt，目标里仍是实测 11pt）。
    _cm_applied += _merge_comment_values(page, _cm_page, resolution, _CM,
                                         allowed=set(TargetProfile.PAGE_KEYS))
    # 页码位置是**独立字段**（不在 page 字典里），单独取
    if _cm_page.get("page_number_position"):
        tpl_page_num = str(_cm_page["page_number_position"])
        resolution["page_number_position"] = _CM
        _cm_applied.append("page_number_position")
    _cm_applied += _merge_comment_values(tpl_tp, _cm.get("title_page"), resolution, _CM)
    _cm_applied += _merge_comment_values(tpl_abs, _cm.get("abstract"), resolution, _CM)
    _cm_applied += _merge_comment_values(tpl_cap, _cm.get("caption"), resolution, _CM)
    _cm_applied += _merge_comment_values(tpl_tbl, _cm.get("table"), resolution, _CM)
    _cm_applied += _merge_comment_values(tpl_toc, _cm.get("toc"), resolution, _CM)
    # 参考文献细项：这三项已由上面的显式分支定过（下拉 > 批注 > 实测 > 规范），
    # 不能再进 ref_extra —— 否则和 TargetProfile 的显式关键字重复，直接 TypeError。
    _cm_ref_rest = {k: v for k, v in _cm_ref.items()
                    if k not in ("reference_style", "reference_hanging_indent_in",
                                 "reference_numbered")}
    _cm_applied += _merge_comment_values(ref_extra, _cm_ref_rest, resolution, _CM)

    # —— 每级标题格式：批注逐字段覆盖实测（只覆盖批注写了的字段）——
    _cm_hf = _cm.get("heading_formats") or {}
    for _lvl, _vals in _cm_hf.items():
        _merged = dict(heading_formats.get(_lvl) or {})
        _hit = False
        for _k, _v in (_vals or {}).items():
            if _v is None:
                continue
            _merged[_k] = _v
            _hit = True
        if _hit:
            heading_formats[_lvl] = _merged
            _cm_applied.append("heading_formats")
    if _cm_hf:
        resolution["heading_formats"] = _CM
        # 批注可能声明了实测里没有的层级（模板只写了规范、没排样例）→ 层级数要跟上，
        # 否则修正链在 _assign_heading_levels 处提前返回，标题一个都不改。
        _max_cm_lvl = max(_cm_hf)
        if _tpl_levels is None or _max_cm_lvl > _tpl_levels:
            _tpl_levels = _max_cm_lvl
            resolution["heading_levels"] = _CM

    # —— 结构维度：栏数 / 纸张（批注 > 实测）——
    _cm_struct = _cm.get("structural") or {}
    if _cm_struct.get("columns") is not None:
        tpl_columns = int(_cm_struct["columns"])
        resolution["columns"] = _CM
        _cm_applied.append("columns")
    if _cm_struct.get("paper_size"):
        tpl_paper = str(_cm_struct["paper_size"])
        resolution["paper_size"] = _CM
        _cm_applied.append("paper_size")

    # —— 表格线（三线表）——
    # 🔴 此前模板模式**从没把 table_lines 传进 TargetProfile**，于是勾了学校模板也
    # 永远不查三线表（客户明确要求「表格要看是不是三线表」根本落不了地）。现补上：
    # 先按模板自己的表推断（过半是三线表即视为模板要求），再让批注覆盖。
    tpl_table_lines = None
    _tls = getattr(doc_profile, "table_line_states", None) or []
    if _tls:
        _n3 = sum(1 for _s in _tls if isinstance(_s, dict) and _s.get("three_line"))
        if _n3 and _n3 * 2 >= len(_tls):
            tpl_table_lines = "three_line"
            resolution["table_lines"] = "school"
    if _cm_struct.get("table_lines"):
        tpl_table_lines = str(_cm_struct["table_lines"])
        resolution["table_lines"] = _CM
        _cm_applied.append("table_lines")

    # —— 公式（v2.3.112）——
    # 🔴 只认**批注**，不从模板实物推断：三线表可以「看模板自己的表长什么样」反推，
    # 公式不行 —— 模板里往往根本没放公式，凭空推断等于替客户拍板。模板没写公式批注
    # 就是不规定（不检查、也不改动）。
    _cm_eq = _cm.get("equation") or {}
    _eq_vals: dict = {}
    if _cm_eq:
        for _k, _allowed in (("equation_alignment", ("center", "left")),
                             ("equation_number_position", ("right", "none"))):
            _v = _cm_eq.get(_k)
            if _v in _allowed:
                _eq_vals[_k] = _v
                resolution[_k] = _CM
                _cm_applied.append(_k)
        if _cm_eq.get("equation_font_family"):
            _eq_vals["equation_font_family"] = str(_cm_eq["equation_font_family"])
            resolution["equation_font_family"] = _CM
            _cm_applied.append("equation_font_family")
        for _k in TargetProfile.EQUATION_KEYS:
            _v = _cm_eq.get(_k)
            if _v is not None:
                _eq_vals[_k] = _v
                resolution[_k] = _CM
                _cm_applied.append(_k)

    # —— 表后间隙兜底（老板 2026-10-29 决策 A）——
    # 「表格与前后文字之间必须留空隙」是硬要求。模板若既不写批注、实测又是 0（表格
    # 紧贴文字 = 模板自己没留），一律按通用兜底：至少一个空行 ≈ 12pt，并在报告里
    # 明说「模板未明确，按通用留白处理」。只在模板**确实有表格**时才兜底。
    _gap_tpl = tpl_tbl.get("table_gap_pt")
    if _tls and (_gap_tpl is None or float(_gap_tpl) <= 0.0):
        tpl_tbl["table_gap_pt"] = 12.0
        resolution["table_gap_pt"] = _DEFAULT_ORIGIN
        notes.append(
            "Table spacing: your template does not state how much space should separate "
            "tables from the text, so the general default is applied — at least 12 pt "
            "(about one blank line) above and below each table.")

    if _cm_applied:
        notes.append(
            "%d formatting rule(s) came from the Word comments in your template. "
            "Where a comment and the template's own formatting disagree, the comment wins."
            % len(_cm_applied))

    return TargetProfile(
        source="template",
        spec_key=None,
        page=page,
        heading_levels=_tpl_levels,
        columns=tpl_columns,
        paper_size=tpl_paper,
        table_lines=tpl_table_lines,
        page_number_position=tpl_page_num,
        # v2.3.112：公式（只来自批注；模板自己没有公式时不推断）
        **_eq_vals,
        **tpl_tp,
        **tpl_abs,
        **tpl_cap,
        **tpl_tbl,
        **tpl_toc,
        heading_formats=heading_formats,
        reference_style=ref_style,
        reference_hanging_indent_in=ref_hang,
        reference_numbered=ref_num,
        **ref_extra,
        resolution=resolution,
        notes=notes,
    )


def _overlay_one(merged: TargetProfile, extra: TargetProfile, conflicts: list) -> TargetProfile:
    """把一个显式来源叠加到 merged 上，并登记来源间冲突。

    来源标签优先取 extra.resolution 里的显式标注；缺失时按 extra.source 经
    _SOURCE_ORIGIN 映射（不再硬编码 "school"）。
    """
    fallback = _SOURCE_ORIGIN.get(extra.source, extra.source)

    def _origin(key: str) -> str:
        return extra.resolution.get(key) or fallback

    def _register(key, prev_val, prev_origin, new_val, new_origin):
        if (prev_val is not None and prev_origin not in (None, "spec")
                and prev_val != new_val):
            conflicts.append({
                "dim": key,
                "entries": [
                    (prev_origin, prev_val),
                    (new_origin, new_val),
                ],
            })

    for k, v in extra.page.items():
        _register(k, merged.page.get(k), merged.resolution.get(k), v, _origin(k))
        merged.page[k] = v
        merged.resolution[k] = _origin(k)

    if extra.heading_levels is not None:
        _register("heading_levels", merged.heading_levels,
                  merged.resolution.get("heading_levels"),
                  extra.heading_levels, _origin("heading_levels"))
        merged.heading_levels = extra.heading_levels
        merged.resolution["heading_levels"] = _origin("heading_levels")

    if extra.columns is not None:
        _register("columns", merged.columns, merged.resolution.get("columns"),
                  extra.columns, _origin("columns"))
        merged.columns = extra.columns
        merged.resolution["columns"] = _origin("columns")

    if extra.paper_size is not None:
        _register("paper_size", merged.paper_size, merged.resolution.get("paper_size"),
                  extra.paper_size, _origin("paper_size"))
        merged.paper_size = extra.paper_size
        merged.resolution["paper_size"] = _origin("paper_size")

    if extra.page_number_position is not None:
        _register("page_number_position", merged.page_number_position,
                  merged.resolution.get("page_number_position"),
                  extra.page_number_position, _origin("page_number_position"))
        merged.page_number_position = extra.page_number_position
        merged.resolution["page_number_position"] = _origin("page_number_position")

    for _k in TargetProfile.TITLE_PAGE_KEYS:
        _v = getattr(extra, _k, None)
        if _v is None:
            continue
        _register(_k, getattr(merged, _k), merged.resolution.get(_k), _v, _origin(_k))
        setattr(merged, _k, _v)
        merged.resolution[_k] = _origin(_k)

    for _k in TargetProfile.ABSTRACT_KEYS:
        _v = getattr(extra, _k, None)
        if _v is None:
            continue
        _register(_k, getattr(merged, _k), merged.resolution.get(_k), _v, _origin(_k))
        setattr(merged, _k, _v)
        merged.resolution[_k] = _origin(_k)

    for _k in TargetProfile.RUNNING_HEAD_KEYS:
        _v = getattr(extra, _k, None)
        if _v is None:
            continue
        _register(_k, getattr(merged, _k), merged.resolution.get(_k), _v, _origin(_k))
        setattr(merged, _k, _v)
        merged.resolution[_k] = _origin(_k)

    for _k in TargetProfile.CAPTION_KEYS:
        _v = getattr(extra, _k, None)
        if _v is None:
            continue
        _register(_k, getattr(merged, _k), merged.resolution.get(_k), _v, _origin(_k))
        setattr(merged, _k, _v)
        merged.resolution[_k] = _origin(_k)

    # v2.3.77：表格内段落 / 目录 —— 与 caption 同构，逐键叠加并登记来源
    for _k in TargetProfile.TABLE_KEYS + TargetProfile.TOC_KEYS:
        _v = getattr(extra, _k, None)
        if _v is None:
            continue
        _register(_k, getattr(merged, _k), merged.resolution.get(_k), _v, _origin(_k))
        setattr(merged, _k, _v)
        merged.resolution[_k] = _origin(_k)

    # v2.3.112：元素级单值维度（表格线 + 公式六项）逐键叠加。
    # 🔴 这一段补的是一个**既有缺口**：`target_from_template_docx` 明明算出了
    # `table_lines="three_line"`（来源 school-comment），但经 build_target —— CLI 与 GUI
    # 的统一入口 —— 合并时这里没有搬运，整条被静默丢弃。后果是「学校模板要求三线表」
    # 在真实产品路径上从未生效过，只有直接调 target_from_template_docx 的单测看得见。
    # 教训：新增一个维度，**必须**在这里登记，否则等于没做。
    for _k in TargetProfile.ELEMENT_KEYS:
        _v = getattr(extra, _k, None)
        if _v is None:
            continue
        _register(_k, getattr(merged, _k, None), merged.resolution.get(_k), _v, _origin(_k))
        setattr(merged, _k, _v)
        merged.resolution[_k] = _origin(_k)

    if extra.reference_style is not None:
        _register("reference_style", merged.reference_style,
                  merged.resolution.get("reference_style"),
                  extra.reference_style, _origin("reference_style"))
        merged.reference_style = extra.reference_style
        merged.resolution["reference_style"] = _origin("reference_style")
    if extra.reference_hanging_indent_in is not None:
        _register("reference_hanging_indent_in", merged.reference_hanging_indent_in,
                  merged.resolution.get("reference_hanging_indent_in"),
                  extra.reference_hanging_indent_in, _origin("reference_hanging_indent_in"))
        merged.reference_hanging_indent_in = extra.reference_hanging_indent_in
        merged.resolution["reference_hanging_indent_in"] = _origin("reference_hanging_indent_in")
    if extra.reference_numbered is not None:
        _register("reference_numbered", merged.reference_numbered,
                  merged.resolution.get("reference_numbered"),
                  extra.reference_numbered, _origin("reference_numbered"))
        merged.reference_numbered = extra.reference_numbered
        merged.resolution["reference_numbered"] = _origin("reference_numbered")
    if extra.heading_formats:
        merged.heading_formats = dict(extra.heading_formats)
        merged.resolution["heading_formats"] = _origin("heading_formats")
    # v2.3.65：参考文献细项（条目字体/字号/行距、另起一页、标题加粗/居中）也要叠上来。
    # 不搬的话，学校模板读到的这几项在合并时被丢弃 —— 模板白读（实测踩过）。
    for _k in TargetProfile.REFERENCE_KEYS:
        _v = getattr(extra, _k, None)
        if _v is None:
            continue
        _register(_k, getattr(merged, _k, None), merged.resolution.get(_k), _v, _origin(_k))
        setattr(merged, _k, _v)
        merged.resolution[_k] = _origin(_k)
    for _note in (getattr(extra, "notes", None) or []):
        if _note and _note not in merged.notes:
            merged.notes.append(_note)
    return merged


def build_target(spec_key: Optional[str] = None, template: Optional[str] = None,
                 ai: Optional[str] = None,
                 questionnaire: Optional[dict] = None,
                 reference_style: Optional[str] = None) -> TargetProfile:
    """四通道统一合并入口（FR2：任选其一或组合）。

    两种基线模式：
    - 规范基线：spec_key 为真实规范（APA-Student / APA-Professional / MLA / IEEE，
      含别名 "APA"）时，以 target_from_spec 为基线（页面软参考 / 引用硬判）。
    - 无规范基线：spec_key 为 None 或 "Other"（学校模板 / 自定义来源）时，
      **不套任何内置规范页面值**，页面以学校模板 / 问卷 / AI 为准；引用由
      reference_style 参数（学校模板子下拉）或问卷 / AI 数据驱动，spec_key 恒为 None。

    template: 学校 docx 模板路径
    ai: AI 填表 JSON 路径
    questionnaire: 用户手填问卷 dict
    reference_style: 学校模板子下拉选中的引用风格（仅无规范基线模式生效；
                    留空表示「按学校模板」——模板批注有流派就按它、没有才不查，
                    不是「整块参考文献都不看」）
    """
    # v2.3.65：**空串必须当"无规范"**。UI 里「无规范基线」存的是 ``""``（StringVar
    # 的默认值），旧写法只判 ``is not None`` → 空串被当成合法规范键，于是客户选了
    # 「学校模板 / 自定义」，背后仍然套着一层规范基线（实测会凭空注入
    # reference_style / reference_numbered / reference_hanging_indent_in 三项硬规则），
    # 与"留空即不检查"直接冲突。
    use_spec = bool(spec_key) and spec_key != "Other"
    if use_spec:
        base = target_from_spec(spec_key)
        ref_for_template = None  # 规范已权威，学校模板子下拉不再覆盖引用风格
    else:
        base = TargetProfile(source="")  # 空基线：页面 / 引用全由显式来源提供
        ref_for_template = reference_style if reference_style in REFERENCE_STYLE_KEYS else None

    built: dict = {}
    if template:
        built["template"] = ("school", target_from_template_docx(read_docx(template), ref_for_template))
    if ai:
        built["ai"] = ("ai_import", target_from_ai(ai))
    if questionnaire:
        built["questionnaire"] = ("user", target_from_dict(questionnaire, "user"))

    merged = base
    conflicts: list = []
    notes: list = list(base.notes)  # 继承规范自带的「期望说明」（学生/专业分层等）
    applied_labels: list = []
    for channel in _CHANNEL_ORDER:              # 从低优先级到高优先级依次叠加
        item = built.get(channel)
        if item is None:
            continue
        t = item[1]
        merged = _overlay_one(merged, t, conflicts)
        applied_labels.append(_SOURCE_ORIGIN.get(t.source, t.source))

    if use_spec:
        merged.source = "+".join(["spec"] + applied_labels)
    else:
        merged.source = "+".join([s for s in applied_labels] or ["custom"])
    merged.conflicts = conflicts
    # 各来源自带的提示（如「模板与所选风格的悬挂缩进冲突，以模板为准」）也要带到结果里
    _channel_notes = [n for n in merged.notes if n not in notes]
    merged.notes = notes + _channel_notes
    return merged


# 页面可数值化的字段（录入弹窗里按 float 解析；其余按字符串原样写回）
PAGE_FLOAT_KEYS = {
    "margin_top_in", "margin_bottom_in", "margin_left_in", "margin_right_in",
    "font_size_pt", "line_spacing", "first_line_indent_in",
    "space_after_pt", "header_distance_in",
}


def apply_page_overrides(target: "TargetProfile", edits: dict) -> None:
    """复核弹窗：仅把用户「真正改动过」的 Page Layout 字段写回 target。

    edits: {page_key: raw_string}. 仅当 raw 与 target.page 当前值**不同**才写回，
    并把 resolution 标 "user"（显式覆盖）。未改动 / 仅回显的字段保持原 resolution，
    避免把规范的软参考（spec 来源 → checker 降级为 info）误升级为硬规则（warn/fail）。

    纯函数、不依赖 GUI，便于 headless 单测。
    """
    for key, raw in edits.items():
        # 栏数是结构性维度（不在 PAGE_KEYS 里），单独处理：int 归一、仅改动才升级。
        if key == "columns":
            raw = (raw or "").strip()
            if raw == "":
                continue
            try:
                iv = int(raw)
            except (ValueError, TypeError):
                continue
            if iv < 1:
                continue
            if iv == target.columns:
                continue  # 与提取/规范值一致 → 不覆盖、不升级
            target.columns = iv
            target.resolution["columns"] = "user"
            continue
        if key == "paper_size":
            # 纸型：通用维度（不在 PAGE_KEYS 里）；「—」= 不规定 → 清空并取消检查。
            raw = (raw or "").strip()
            if raw in ("", "—", "-", "none", "None"):
                if target.paper_size is not None:
                    target.paper_size = None
                    target.resolution.pop("paper_size", None)
                continue
            if raw not in PAPER_SIZES:
                continue
            if raw == target.paper_size:
                continue
            target.paper_size = raw
            target.resolution["paper_size"] = "user"
            continue
        if key == "page_number_position":
            # 页码位置：与纸型同套路 ——「—」= 不规定 → 清空并取消检查；
            # 白名单外的值一律丢弃（绝不把拼错的值当"客户要求"去改文档）。
            raw = (raw or "").strip()
            if raw in ("", "—", "-", "none", "None"):
                if target.page_number_position is not None:
                    target.page_number_position = None
                    target.resolution.pop("page_number_position", None)
                continue
            if raw not in PAGE_NUMBER_POSITIONS:
                continue
            if raw == target.page_number_position:
                continue
            target.page_number_position = raw
            target.resolution["page_number_position"] = "user"
            continue
        # 摘要页特化（v2.3.67）：首行缩进是 float、另起一页是三态，都走「—」=清空逻辑。
        if key == "abstract_first_line_indent_in":
            raw = (raw or "").strip()
            if raw in ("", "—", "-", "none", "None"):
                if target.abstract_first_line_indent_in is not None:
                    target.abstract_first_line_indent_in = None
                    target.resolution.pop("abstract_first_line_indent_in", None)
                continue
            try:
                fv = float(raw)
            except (ValueError, TypeError):
                continue
            if fv == target.abstract_first_line_indent_in:
                continue
            target.abstract_first_line_indent_in = fv
            target.resolution["abstract_first_line_indent_in"] = "user"
            continue
        if key == "abstract_new_page":
            raw = (raw or "").strip()
            if raw in ("", "—", "-", "none", "None"):
                if target.abstract_new_page is not None:
                    target.abstract_new_page = None
                    target.resolution.pop("abstract_new_page", None)
                continue
            _tv = raw.lower() in ("1", "true", "yes", "y")
            _fv2 = raw.lower() in ("0", "false", "no", "n")
            if not (_tv or _fv2):
                continue
            if bool(target.abstract_new_page) == _tv:
                continue
            target.abstract_new_page = _tv
            target.resolution["abstract_new_page"] = "user"
            continue
        # Running head（v2.3.67）：两个三态开关，都走「—」=清空逻辑。
        if key in TargetProfile.RUNNING_HEAD_KEYS:
            raw = (raw or "").strip()
            if raw in ("", "—", "-", "none", "None"):
                if getattr(target, key, None) is not None:
                    setattr(target, key, None)
                    target.resolution.pop(key, None)
                continue
            _tv = raw.lower() in ("1", "true", "yes", "y")
            _fv2 = raw.lower() in ("0", "false", "no", "n")
            if not (_tv or _fv2):
                continue
            if bool(getattr(target, key, None)) == _tv:
                continue
            setattr(target, key, _tv)
            target.resolution[key] = "user"
            continue
        # 摘要标题格式（v2.3.97）：对齐是 choice、加粗/斜体是三态，都走「—」=清空。
        if key == "abstract_heading_alignment":
            raw = (raw or "").strip()
            if raw in ("", "—", "-", "none", "None"):
                if target.abstract_heading_alignment is not None:
                    target.abstract_heading_alignment = None
                    target.resolution.pop("abstract_heading_alignment", None)
                continue
            _al = raw.lower()
            if _al not in ("left", "center"):
                continue
            if target.abstract_heading_alignment == _al:
                continue
            target.abstract_heading_alignment = _al
            target.resolution["abstract_heading_alignment"] = "user"
            continue
        if key in ("abstract_heading_bold", "abstract_heading_italic",
                   "keywords_label_italic"):
            raw = (raw or "").strip()
            if raw in ("", "—", "-", "none", "None"):
                if getattr(target, key, None) is not None:
                    setattr(target, key, None)
                    target.resolution.pop(key, None)
                continue
            _tv = raw.lower() in ("1", "true", "yes", "y")
            _fv2 = raw.lower() in ("0", "false", "no", "n")
            if not (_tv or _fv2):
                continue
            if getattr(target, key, None) is _tv:
                continue
            setattr(target, key, _tv)
            target.resolution[key] = "user"
            continue
        if key == "keywords_indent_in":
            raw = (raw or "").strip()
            if raw in ("", "—", "-", "none", "None"):
                if target.keywords_indent_in is not None:
                    target.keywords_indent_in = None
                    target.resolution.pop("keywords_indent_in", None)
                continue
            try:
                fv = float(raw)
            except (ValueError, TypeError):
                continue
            if fv == target.keywords_indent_in:
                continue
            target.keywords_indent_in = fv
            target.resolution["keywords_indent_in"] = "user"
            continue
        # 图表 caption（v2.3.67）：字号/行距是 float、位置是 choice，都走「—」=清空。
        if key in ("caption_font_size_pt", "caption_line_spacing"):
            raw = (raw or "").strip()
            if raw in ("", "—", "-", "none", "None"):
                if getattr(target, key, None) is not None:
                    setattr(target, key, None)
                    target.resolution.pop(key, None)
                continue
            try:
                fv = float(raw)
            except (ValueError, TypeError):
                continue
            if fv == getattr(target, key, None):
                continue
            setattr(target, key, fv)
            target.resolution[key] = "user"
            continue
        if key in ("caption_position_table", "caption_position_figure"):
            raw = (raw or "").strip()
            if raw in ("", "—", "-", "none", "None"):
                if getattr(target, key, None) is not None:
                    setattr(target, key, None)
                    target.resolution.pop(key, None)
                continue
            _val = raw.lower()
            if _val not in ("top", "bottom"):
                continue
            if _val == getattr(target, key, None):
                continue
            setattr(target, key, _val)
            target.resolution[key] = "user"
            continue
        # v2.3.85：表格线（三线表）。「—」= 清空（不检查、不改动表格边框）。
        if key == "table_lines":
            raw = (raw or "").strip()
            if raw in ("", "—", "-", "none", "None"):
                if getattr(target, key, None) is not None:
                    setattr(target, key, None)
                    target.resolution.pop(key, None)
                continue
            _val = raw.lower().replace("-", "_")
            if _val not in ("three_line", "boxed"):
                continue
            if _val == getattr(target, key, None):
                continue
            setattr(target, key, _val)
            target.resolution[key] = "user"
            continue
        if key not in TargetProfile.PAGE_KEYS:
            continue
        raw = (raw or "").strip()
        if raw == "":
            continue
        orig = target.page.get(key)
        if raw == ("" if orig is None else str(orig)):
            continue  # 与提取/规范值一致 → 不覆盖、不升级
        if key in PAGE_FLOAT_KEYS:
            try:
                fv = float(raw)
            except (ValueError, TypeError):
                continue
            target.page[key] = fv
        else:
            target.page[key] = raw
        target.resolution[key] = "user"


def merge_school(spec_target: TargetProfile, school: TargetProfile) -> TargetProfile:
    """向后兼容的二元合并（页面以学校为准，引用以规范为准）。

    保留 school 自带的来源标签（不再统一覆盖成 "school"）。
    """
    conflicts: list = []
    merged = _overlay_one(spec_target, school, conflicts)
    merged.conflicts = conflicts
    merged.source = "spec+school"
    return merged
