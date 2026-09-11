"""界面文案 / 字体族 / 语言偏好（海外版：中英双语，**默认英文**）。

设计约束
--------
- 纯数据 + 纯函数，**不 import tkinter**：字体族名的平台兜底与词条取用都能离线单测
  （tkinter 只在 ``fonts.py`` 里用于构造真正的 Font 对象）。
- 所有面向用户的字符串集中在本模块，GUI 只按 key 取词 —— 禁止在界面代码里散落硬编码，
  否则加一种语言就要满仓库找字符串。
- **品牌口径（海外版铁律）**：只出现官网 ``reedskill.com`` 与邮箱 ``hi@reedskill.com``。
  **绝不出现微信 / 公众号 / 小程序**（海外用户不用这些，老板 2026-09-11 明确要求）；
  国内版的公众号二维码、客服微信号在搬布局时一并剔除。

语言偏好落盘在 ``~/.thesis-format-doctor-global/settings.json``（与 license.json 同目录），
可用环境变量 ``TFD_SETTINGS_FILE`` 覆盖 —— 与授权状态文件同样的隔离策略，便于测试。
"""

from __future__ import annotations

import json
import os
import platform
from pathlib import Path

LANGS = ("en", "zh")
DEFAULT_LANG = "en"          # 海外版默认英文（老板 2026-09-11 定）

SETTINGS_FILE = Path.home() / ".thesis-format-doctor-global" / "settings.json"

BRAND_EMAIL = "hi@reedskill.com"
BRAND_SITE = "reedskill.com"
BRAND_SITE_URL = "https://reedskill.com"
BRAND_NAME = "ReedSkill"


def settings_file() -> Path:
    """实际设置文件路径（每次读 env，便于运行时切换 / 测试隔离）。"""
    env = os.environ.get("TFD_SETTINGS_FILE")
    if env:
        return Path(env)
    return SETTINGS_FILE


def load_lang() -> str:
    """读语言偏好；任何异常都回退到默认语言（设置文件绝不阻塞启动）。"""
    try:
        f = settings_file()
        if f.is_file():
            raw = json.loads(f.read_text(encoding="utf-8"))
            lang = (raw or {}).get("lang")
            if lang in LANGS:
                return lang
    except Exception:
        pass
    return DEFAULT_LANG


def save_lang(lang: str) -> bool:
    """写语言偏好（原子写）；失败返回 False —— 记不住偏好不影响本次使用。"""
    if lang not in LANGS:
        return False
    try:
        f = settings_file()
        f.parent.mkdir(parents=True, exist_ok=True)
        data = {"lang": lang}
        tmp = f.with_name(f.name + ".tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        os.replace(str(tmp), str(f))
        return True
    except Exception:
        return False


def other_lang(lang: str) -> str:
    """语言切换按钮的目标语言（en <-> zh）。"""
    return "zh" if lang == "en" else "en"


# ---------------------------------------------------------------------------
# 字体族：英文优先「漂亮」的衬线 + 清爽无衬线；中文沿用国内版的楷体 + 雅黑
# ---------------------------------------------------------------------------
# 为什么不用一套字体通吃：
#   英文标题用 Georgia（学术衬线，笔画有粗细对比、字重饱满，打印/屏幕都体面），
#   正文用 Segoe UI（Windows 原生 UI 字体，小字号清晰、字面开阔）；
#   中文字体若排在英文之前，英文字符会用中文字体的西文字形（雅黑的西文偏窄偏硬），
#   所以中英各留一套，由当前语言决定。
#
# 跨平台兜底：macOS / Linux 上没有 Georgia / Segoe UI / 楷体 / 雅黑，
# 映射到系统自带的对等字体，避免 tkinter 静默回退成"方块字"（tofu）。
_FAMILY_FALLBACK = {
    # 英文
    "Georgia": {"Darwin": "Georgia", "Linux": "Noto Serif"},
    "Segoe UI": {"Darwin": "Helvetica Neue", "Linux": "Noto Sans"},
    "Consolas": {"Darwin": "Menlo", "Linux": "DejaVu Sans Mono"},
    # 中文
    "KaiTi": {"Darwin": "STKaiti", "Linux": "Noto Serif CJK SC"},
    "Microsoft YaHei": {"Darwin": "PingFang SC", "Linux": "Noto Sans CJK SC"},
}

# name -> (family, size, weight?)；weight 缺省为 normal
_FONT_SPEC = {
    "en": {
        "F_TITLE":        ("Georgia", 21, "bold"),      # 主标题（学术衬线）
        "F_DIALOG_TITLE": ("Georgia", 15, "bold"),      # 弹窗标题
        "F_HDR":          ("Georgia", 13, "bold"),      # 章节标题
        "F_CARD_HDR":     ("Georgia", 13, "bold"),      # 卡片标题
        "F_STAT":         ("Georgia", 12),              # 状态文字
        "F_ICON":         ("Georgia", 12, "bold"),      # 行首印记
        "F_SUB":          ("Segoe UI", 10),             # 副标题
        "F_SUBTITLE":     ("Segoe UI", 11),             # 小标题 / 元信息
        "F_BODY":         ("Segoe UI", 12),             # 正文
        "F_SMALL":        ("Segoe UI", 10),             # 说明 / 次要
        "F_SMALL_B":      ("Segoe UI", 10, "bold"),     # 提示框标题
        "F_BTN":          ("Segoe UI", 12, "bold"),     # 按钮
        "F_FOOT":         ("Segoe UI", 9),              # 页脚 / 状态栏
        "F_MONO":         ("Consolas", 9),              # 机器码 / 离线码
    },
    "zh": {
        "F_TITLE":        ("KaiTi", 20, "bold"),
        "F_DIALOG_TITLE": ("KaiTi", 14, "bold"),
        "F_HDR":          ("KaiTi", 13, "bold"),
        "F_CARD_HDR":     ("KaiTi", 13, "bold"),
        "F_STAT":         ("KaiTi", 12),
        "F_ICON":         ("KaiTi", 12, "bold"),
        "F_SUB":          ("Microsoft YaHei", 10),
        "F_SUBTITLE":     ("Microsoft YaHei", 11),
        "F_BODY":         ("Microsoft YaHei", 12),
        "F_SMALL":        ("Microsoft YaHei", 10),
        "F_SMALL_B":      ("Microsoft YaHei", 10, "bold"),
        "F_BTN":          ("Microsoft YaHei", 12, "bold"),
        "F_FOOT":         ("Microsoft YaHei", 9),
        "F_MONO":         ("Consolas", 9),
    },
}

FONT_NAMES = tuple(_FONT_SPEC["en"].keys())


def resolve_family(family: str, available=None) -> str:
    """把字体族名解析成当前系统真实可用的族名。

    available: 可迭代的系统字体族名集合（tkinter 的 ``tkfont.families()``）。
    为 None 时跳过"是否存在"检查，直接在非 Windows 平台映射兜底值。
    """
    if family in _FAMILY_FALLBACK:
        system = platform.system()
        cand = _FAMILY_FALLBACK[family].get(system)
        if system == "Windows":
            cand = None                     # Windows 自带该字体，直接用
    else:
        cand = None
    if available is not None:
        avail = set(available)
        if family in avail:
            return family
        if cand and cand in avail:
            return cand
        # 最后挣扎：Linux 上字体名带不带空格版本都可能存在
        for probe in _probe_names(family):
            if probe in avail:
                return probe
        return family if not cand else cand
    return cand or family


def _probe_names(family: str) -> tuple:
    if family == "Noto Serif":
        return ("Noto Serif", "Liberation Serif", "DejaVu Serif", "Times New Roman")
    if family == "Noto Sans":
        return ("Noto Sans", "Liberation Sans", "DejaVu Sans")
    if family == "DejaVu Sans Mono":
        return ("DejaVu Sans Mono", "Liberation Mono", "Courier New")
    if family == "STKaiti":
        return ("STKaiti", "Kaiti SC", "Songti SC", "STSong")
    if family == "PingFang SC":
        return ("PingFang SC", "Heiti SC", "Hiragino Sans GB", "Arial Unicode MS")
    if family == "Noto Serif CJK SC":
        return ("Noto Serif CJK SC", "Source Han Serif SC", "AR PL UMing CN", "Noto Serif")
    if family == "Noto Sans CJK SC":
        return ("Noto Sans CJK SC", "Source Han Sans SC", "WenQuanYi Micro Hei", "Noto Sans")
    return ()


def font_spec(lang: str, available=None) -> dict:
    """返回 ``{F_XXX: (family, size, weight)}``，family 已按平台兜底。

    缺省 weight 的项统一补成 "normal"，调用方无需判断长度。
    """
    lang = lang if lang in LANGS else DEFAULT_LANG
    out: dict = {}
    for name, spec in _FONT_SPEC[lang].items():
        family = resolve_family(spec[0], available)
        weight = spec[2] if len(spec) > 2 else "normal"
        out[name] = (family, spec[1], weight)
    return out


# ---------------------------------------------------------------------------
# 界面文案
# ---------------------------------------------------------------------------
STRINGS: dict = {
    "en": {
        # —— 顶栏 / 标题区 ——
        "link_help": "Help",
        "link_about": "About",
        "lang_button": "中文",                       # 显示"可切换到"的语言
        "app_title": "Thesis Format Doctor",
        "app_subtitle": "GLOBAL EDITION · FORMAT COMPLIANCE FOR ACADEMIC PAPERS",

        # —— 左栏：文件选择 ——
        "files_title": "Files",
        "badge_offline": "Offline",
        "mark_required": "Required",
        "mark_optional": "Optional",
        "btn_choose": "Choose…",
        "ic_spec": "S", "ic_template": "T", "ic_latex": "L",
        "ic_questionnaire": "Q", "ic_ai": "A", "ic_paper": "P",
        "row_spec_title": "Citation Style",
        "row_spec_desc": "APA · MLA · Chicago · IEEE · Harvard · Other",
        "row_template_title": "School Template",
        "row_template_desc": "Word template — page layout follows your school",
        "row_latex_title": "LaTeX Template",
        "row_latex_desc": ".cls / .sty / .tex (parsed offline)",
        "row_questionnaire_title": "Format Questionnaire",
        "row_questionnaire_desc": "Page layout: margins, font, size, spacing",
        "btn_questionnaire": "Fill in",
        "row_ai_title": "AI-filled JSON",
        "row_ai_desc": "Let any AI fill the questionnaire template for you",
        "btn_ai_json": "Import",
        "row_paper_title": "Paper to Check",
        "row_paper_desc": "Word document (.docx)",
        "chip_style_hint": "Other = custom; pair it with a school template or questionnaire.",

        # 选中状态行
        "st_paper_none": "No paper selected yet",
        "st_paper_set": "Paper",
        "st_tpl_none": "No template selected (optional)",
        "st_tpl_set": "Template",
        "st_latex_none": "No LaTeX template (optional)",
        "st_latex_set": "LaTeX",
        "st_q_none": "No questionnaire filled in (optional)",
        "st_q_set": "Questionnaire filled in",
        "st_ai_none": "No AI-filled JSON imported (optional)",
        "st_ai_set": "AI-filled JSON",

        # —— 右栏：处理步骤 ——
        "steps_title": "Workflow",
        "step_counter": "%d / %d",
        "step1_title": "Choose Material",
        "step1_desc": "Citation style and the paper to process",
        "step2_title": "Format Check",
        "step2_desc": "Free and unlimited — read-only",
        "step3_title": "One-click Fix",
        "step3_desc": "First fix free, then activation",
        "btn_check": "▶  Run Format Check",
        "btn_fix": "✎  One-click Fix",
        "btn_save_report": "Save Report",
        "btn_export_ai": "Export AI Template",
        "btn_activate": "Activation",
        "chapter_one": "Ⅰ · SELECT",
        "chapter_two": "Ⅱ · PROCESS",

        # —— 报告区 ——
        "report_title": "Report",
        "report_placeholder": (
            "Pick a citation style and your paper, then run the format check.\n\n"
            "Everything runs on this computer — no upload, no account, no network."
        ),

        # —— 状态栏 ——
        "bar_left": "Thesis Format Doctor Global v%s",
        "bar_idle": "Ready",
        "bar_checking": "Checking format…",
        "bar_fixing": "Applying format fixes…",
        "bar_check_done": "Check complete",
        "bar_fix_done": "Fix complete",
        "bar_error": "Error",
        "status_idle": "Ready — select a style and a paper to begin.",
        "status_checking": "Checking the format of your document…",
        "status_fixing": "Applying format fixes (content stays untouched)…",
        "status_error": "Something went wrong.",

        # —— 页脚 ——
        "footer_line1": "All processing happens on your device — your paper never leaves this computer.",
        "footer_line2": "© 2026 %s · %s" % (BRAND_NAME, BRAND_EMAIL),
        "footer_site": "Website: %s" % BRAND_SITE,

        # —— 通用提示 ——
        "msg_missing_paper_title": "No paper selected",
        "msg_missing_paper": "Please choose the Word document (.docx) you want to process.",
        "msg_read_error_title": "Cannot read the document",
        "msg_check_error_title": "Check failed",
        "msg_check_error_unexpected": "Check failed (unexpected error)",
        "msg_fix_preview_error_title": "Preview failed",
        "msg_fix_error_title": "Fix failed",
        "msg_fix_error_unexpected": "Preview failed (unexpected error)",
        "msg_cannot_overwrite_title": "Cannot overwrite the original",
        "msg_cannot_overwrite": "The output path is the same file as your paper. Please choose another location.",
        "msg_no_change_title": "Nothing to fix",
        "msg_no_change": ("Your document already matches the target format. "
                          "No copy was created and no trial quota was used."),
        "msg_need_activation_title": "Activation required",
        "msg_fix_confirm_title": "Confirm the fix",
        "msg_fix_confirm_body": "\n\nClick Yes to apply the fixes, then choose where to save.\nYour original file will not be modified.",
        "msg_fix_pick_output": "Choose where to save the fixed copy (default: name_fixed.docx)",
        "msg_fix_done_title": "Fix complete",
        "msg_fix_done_before": "Saved to:",
        "msg_fix_done_before_file": "Saved:",
        "preview_header": "Format fixes to apply (format only — the text itself is never changed):",
        "preview_trial": "License status: %s",

        # —— 保存报告 / 导出模板 ——
        "save_no_report_title": "No report yet",
        "save_no_report": "Run the format check first to generate a report.",
        "save_report_dialog": "Save the check report",
        "save_report_done_title": "Saved",
        "save_report_done": "Report saved to:",
        "export_done_title": "Exported",
        "export_done": ("The AI questionnaire template has been exported to:\n%s\n\n"
                        "Hand it to any AI assistant and ask it to fill it in for your school."),

        # —— 问卷弹窗 ——
        "q_title": "Format Questionnaire",
        "q_hint": "Fill in what you know — anything left blank stays unset.",
        "q_save": "Save questionnaire",
        "q_saved_title": "Saved",
        "q_saved": "Questionnaire saved. It will be used the next time you run.",
        "q_bad_number_title": "Not a number",
        "q_bad_number": "'%s' expects a number, but got: %s",
        "q_range_title": "Outside the usual range",
        "q_range_body": "These values are outside the usual range. Save anyway?\n\n%s",
        "unit_in": "in", "unit_pt": "pt", "unit_x": "x", "unit_level": "levels",
        "bool_yes": "Yes", "bool_no": "No",

        # —— 激活弹窗 ——
        "act_title": "Activation",
        "act_subtitle": "Unlock unlimited fixes on this computer. No subscription, no account.",
        "act_online_label": "① Activation code (issued online)",
        "act_online_btn": "Activate",
        "act_online_hint": "Paste the code you received.",
        "act_machine_label": "② Your machine code — send it to us for an offline code",
        "act_machine_hint": "We sign it offline and send the code back. Nothing is uploaded by the app.",
        "act_copy_btn": "Copy machine code",
        "act_copy_ok_title": "Copied",
        "act_copy_ok": "Machine code copied. Email it to %s and we will send your code." % BRAND_EMAIL,
        "act_offline_label": "③ Offline activation code (no internet needed)",
        "act_offline_btn": "Activate offline",
        "act_offline_hint": "Signed for this machine only.",
        "act_need_code": "Please enter the activation code first.",
        "act_need_offline": "Please paste the offline code you received first.",
        "act_result_title": "Activation",

        # —— 关于 / 帮助 ——
        "about_title": "About",
        "about_version": "Version %s",
        "about_tagline": "Hold your paper to your school's standard.",
        "about_body": ("Heading levels, fonts and sizes, margins, line spacing, references — "
                       "checked one by one against the target, and corrected to match."),
        "about_privacy_hdr": "Privacy",
        "about_privacy": "Everything runs on your device. Your paper is never uploaded, never collected.",
        "about_source_hdr": "Where the rules come from",
        "about_source": ("Page layout follows your school's template; citation format follows the "
                         "style you pick (APA, MLA, Chicago, IEEE, Harvard)."),
        "about_smartscreen": ("Unsigned build — Windows SmartScreen may warn on first launch. "
                              "Click More info → Run anyway."),
        "about_contact_hdr": "Contact",
        "about_footer": "© 2026 %s" % BRAND_NAME,

        "help_title": "Help",
        "help_intro": "Tell the app your school's rules, pick your paper, and let it do the rest.",
        "help_sec1": "1 · How to use it",
        "help_sec2": "2 · Frequently asked",
        "help_steps": [
            ("Pick a citation style",
             "APA, MLA, Chicago, IEEE or Harvard. Choose Other if your school uses something else."),
            ("Add your school's rules (optional but recommended)",
             "Upload the Word template your school handed out, fill in the questionnaire, import a "
             "JSON an AI filled for you, or hand over a LaTeX template. Page layout comes from here."),
            ("Choose the paper",
             "Your .docx file. Nothing is uploaded — the file is read and written on this computer only."),
            ("Run the format check",
             "Free and unlimited. You get a report listing every mismatch and which side wins "
             "when two sources disagree."),
            ("One-click fix (optional)",
             "Applies the format fixes and writes a new .docx. Text, punctuation and wording are "
             "never touched. Your original file is left exactly as it was."),
        ],
        "help_faqs": [
            ("Q1 · My school does not give out a template. Can I still use it?",
             "Yes. Without a template the app falls back to the defaults of the citation style you "
             "picked. Page layout is where schools differ most, so a template — or the questionnaire — "
             "gives a much closer match."),
            ("Q2 · Does the fix change my writing?",
             "No. Only formatting: fonts, sizes, margins, line spacing, indentation, alignment and "
             "reference formatting. Headings keep their existing size and weight on purpose, and the "
             "app always writes a new file instead of overwriting yours."),
            ("Q3 · Is my paper uploaded anywhere?",
             "No. There is no server in the loop — checking and fixing both run entirely on your "
             "computer, and the app works with the network unplugged."),
            ("Q4 · It says activation is required. How do I activate?",
             "The first fix is free. After that, open Activation: either paste an online activation "
             "code, or copy your machine code, email it to %s, and paste back the offline code "
             "we send you." % BRAND_EMAIL),
            ("Q5 · Windows warns me when I open the app.",
             "The build is not code-signed yet, so SmartScreen flags it. Click More info → Run anyway. "
             "The executable is built from the public source of this project."),
        ],
    },

    "zh": {
        "link_help": "帮 助",
        "link_about": "关 于",
        "lang_button": "English",
        "app_title": "论 文 格 式 医 生",
        "app_subtitle": "海外版 · 学术论文格式规范引擎",

        "files_title": "文件选择",
        "badge_offline": "本机离线",
        "mark_required": "必选",
        "mark_optional": "可选",
        "btn_choose": "选择…",
        "ic_spec": "规", "ic_template": "模", "ic_latex": "L",
        "ic_questionnaire": "问", "ic_ai": "A", "ic_paper": "论",
        "row_spec_title": "引用规范",
        "row_spec_desc": "APA · MLA · Chicago · IEEE · Harvard · Other",
        "row_template_title": "学校模板",
        "row_template_desc": "Word 模板 —— 页面排版以学校为准",
        "row_latex_title": "LaTeX 模板",
        "row_latex_desc": ".cls / .sty / .tex（本机离线解析）",
        "row_questionnaire_title": "格式问卷",
        "row_questionnaire_desc": "页面排版：页边距、字体、字号、行距",
        "btn_questionnaire": "手填问卷",
        "row_ai_title": "AI 填表 JSON",
        "row_ai_desc": "把问卷模板交给任意 AI，让它按学校要求填好",
        "btn_ai_json": "导入",
        "row_paper_title": "待检论文",
        "row_paper_desc": "Word 文档（.docx）",
        "chip_style_hint": "Other = 自定义；须配学校模板或格式问卷。",

        "st_paper_none": "尚未选择论文",
        "st_paper_set": "论文",
        "st_tpl_none": "模板未选（可选）",
        "st_tpl_set": "模板",
        "st_latex_none": "LaTeX 模板未选（可选）",
        "st_latex_set": "LaTeX 模板",
        "st_q_none": "问卷未填（可选）",
        "st_q_set": "问卷已填",
        "st_ai_none": "未导入 AI 填表 JSON（可选）",
        "st_ai_set": "AI 填表 JSON",

        "steps_title": "处理步骤",
        "step_counter": "%d / %d",
        "step1_title": "选择材料",
        "step1_desc": "引用规范与待处理的论文",
        "step2_title": "格式体检",
        "step2_desc": "免费不限次 —— 只读不改",
        "step3_title": "一键修正",
        "step3_desc": "首次免费，之后需激活码",
        "btn_check": "▶  运行格式体检",
        "btn_fix": "✎  一键修正（仅改格式）",
        "btn_save_report": "保存报告",
        "btn_export_ai": "导出 AI 问卷模板",
        "btn_activate": "激活 / 授权",
        "chapter_one": "壹 · 选择",
        "chapter_two": "贰 · 处理",

        "report_title": "体检报告",
        "report_placeholder": (
            "选好引用规范与论文，点「运行格式体检」即可。\n\n"
            "全程在本机完成 —— 不上传、不需要账号、不联网。"
        ),

        "bar_left": "海外版论文格式医生 v%s",
        "bar_idle": "就绪",
        "bar_checking": "正在体检格式…",
        "bar_fixing": "正在应用格式修正…",
        "bar_check_done": "体检完成",
        "bar_fix_done": "修正完成",
        "bar_error": "出错了",
        "status_idle": "就绪 —— 选择规范与论文即可开始。",
        "status_checking": "正在检查文档格式…",
        "status_fixing": "正在应用格式修正（正文内容不会改动）…",
        "status_error": "出错了。",

        "footer_line1": "全部处理在本机完成 —— 论文不会离开这台电脑",
        "footer_line2": "© 2026 %s · %s" % (BRAND_NAME, BRAND_EMAIL),
        "footer_site": "官网：%s" % BRAND_SITE,

        "msg_missing_paper_title": "缺少文档",
        "msg_missing_paper": "请先选择要处理的 Word 文档（.docx）。",
        "msg_read_error_title": "无法读取文档",
        "msg_check_error_title": "检查失败",
        "msg_check_error_unexpected": "检查失败（未预期错误）",
        "msg_fix_preview_error_title": "预览失败",
        "msg_fix_error_title": "修正失败",
        "msg_fix_error_unexpected": "预览失败（未预期错误）",
        "msg_cannot_overwrite_title": "不能覆盖原件",
        "msg_cannot_overwrite": "输出路径与待修正文档是同一个文件，请另选保存位置。",
        "msg_no_change_title": "无需修正",
        "msg_no_change": "当前文档已符合目标格式，未生成副本、未消耗试用额度。",
        "msg_need_activation_title": "需要激活",
        "msg_fix_confirm_title": "确认修正",
        "msg_fix_confirm_body": "\n\n点「是」开始修正，随后选择保存位置（原件不会被改动）。",
        "msg_fix_pick_output": "选择修正稿保存位置（默认 原名_fixed.docx）",
        "msg_fix_done_title": "修正完成",
        "msg_fix_done_before": "已保存到：",
        "msg_fix_done_before_file": "已生成：",
        "preview_header": "将进行的格式修正（仅改格式、绝不改正文）：",
        "preview_trial": "试用状态：%s",

        "save_no_report_title": "没有报告",
        "save_no_report": "请先运行「格式体检」生成报告。",
        "save_report_dialog": "保存体检报告",
        "save_report_done_title": "已保存",
        "save_report_done": "报告已保存到：",
        "export_done_title": "已导出",
        "export_done": ("AI 问卷模板已导出到：\n%s\n\n把它交给任意 AI，让它按你学校的要求填好即可。"),

        "q_title": "格式问卷",
        "q_hint": "知道多少填多少，留空即视为不设定。",
        "q_save": "保存问卷",
        "q_saved_title": "已保存",
        "q_saved": "问卷已填入，运行体检时生效。",
        "q_bad_number_title": "数值格式不对",
        "q_bad_number": "「%s」需要填数字，当前是：%s",
        "q_range_title": "数值超出常见范围",
        "q_range_body": "以下数值超出常见范围，是否仍要保存？\n\n%s",
        "unit_in": "英寸", "unit_pt": "磅", "unit_x": "倍", "unit_level": "级",
        "bool_yes": "是", "bool_no": "否",

        "act_title": "激活",
        "act_subtitle": "一次授权，本机无限次修正。不需要订阅、不需要账号。",
        "act_online_label": "① 在线激活码（中台签发）",
        "act_online_btn": "激活",
        "act_online_hint": "粘贴你收到的激活码。",
        "act_machine_label": "② 本机机器码 —— 发给我们换取离线激活码",
        "act_machine_hint": "我们在离线状态下签发，再把码发还给你。软件本身不会上传任何东西。",
        "act_copy_btn": "复制机器码",
        "act_copy_ok_title": "已复制",
        "act_copy_ok": "本机机器码已复制，发到 %s 即可。" % BRAND_EMAIL,
        "act_offline_label": "③ 离线激活码（无需联网）",
        "act_offline_btn": "离线激活",
        "act_offline_hint": "仅对本机有效。",
        "act_need_code": "请先输入激活码。",
        "act_need_offline": "请先粘贴卖家给的离线激活码。",
        "act_result_title": "激活结果",

        "about_title": "关于",
        "about_version": "版本 v%s",
        "about_tagline": "以学校模板为准绳，为论文格式把脉。",
        "about_body": "标题层级、字体字号、页边距、行距、参考文献 —— 逐项对照目标规范，改至合范。",
        "about_privacy_hdr": "论文安全",
        "about_privacy": "所有处理均在本机完成，论文不联网、不上传、不被收集。",
        "about_source_hdr": "规则从哪来",
        "about_source": "页面排版以学校模板为准；引用格式以所选规范为准（APA / MLA / Chicago / IEEE / Harvard）。",
        "about_smartscreen": "未签名程序 —— 首次打开时 Windows SmartScreen 可能拦截，点「详细信息」→「仍要运行」即可。",
        "about_contact_hdr": "联系我们",
        "about_footer": "© 2026 %s" % BRAND_NAME,

        "help_title": "使用帮助",
        "help_intro": "把学校的格式要求告诉它，导入论文，剩下的交给它。",
        "help_sec1": "一、怎么用",
        "help_sec2": "二、常见问题",
        "help_steps": [
            ("1. 选择引用规范",
             "APA、MLA、Chicago、IEEE 或 Harvard。学校用别的，就选 Other。"),
            ("2. 告诉它学校的格式要求（可选，但强烈建议）",
             "上传学校下发的 Word 模板、手填问卷、导入 AI 填好的 JSON，或给一份 LaTeX 模板。"
             "页面排版以这里为准。"),
            ("3. 选择论文",
             "你的 .docx 文件。全程不上传 —— 文件的读写都只在这台电脑上完成。"),
            ("4. 运行格式体检",
             "免费且不限次。你会拿到一份报告，逐项列出不符之处；两个来源打架时，也会写明以谁为准。"),
            ("5. 一键修正（可选）",
             "套用格式修正并另存一份新的 .docx。文字、标点、措辞一律不动，原文件也原样保留。"),
        ],
        "help_faqs": [
            ("Q1 · 学校不发模板，还能用吗？",
             "能。没有模板时，软件按你所选规范的默认值处理。各校差异最大的就是页面排版，"
             "所以给一份模板或填一份问卷，结果会贴合得多。"),
            ("Q2 · 修正会改我的文字吗？",
             "不会。只动格式：字体、字号、页边距、行距、缩进、对齐与参考文献格式。"
             "标题段有意保留原有字号字重；且软件一律另存新文件，绝不覆盖你的原稿。"),
            ("Q3 · 我的论文会被上传吗？",
             "不会。整个流程没有服务器参与 —— 体检与修正都在本机完成，拔了网线也能用。"),
            ("Q4 · 提示需要激活，怎么激活？",
             "首次修正免费。之后打开「激活」：粘贴在线激活码，或复制本机机器码发到 %s，"
             "把回给你的离线码贴回来即可。" % BRAND_EMAIL),
            ("Q5 · 打开时 Windows 报警怎么办？",
             "程序尚未做代码签名，所以会被 SmartScreen 标注。点「详细信息」→「仍要运行」即可。"
             "可执行文件由本项目的公开源码构建。"),
        ],
    },
}


def t(lang: str, key: str, *args) -> str:
    """取词；缺失时回退英文、再回退 key 本身（界面绝不因缺词而崩）。

    带 ``%s`` 的词条支持位置参数格式化，格式化失败则返回原始词条。
    """
    lang = lang if lang in LANGS else DEFAULT_LANG
    val = STRINGS.get(lang, {}).get(key)
    if val is None:
        val = STRINGS[DEFAULT_LANG].get(key, key)
    if args:
        try:
            return val % args
        except Exception:
            return val
    return val


# ---------------------------------------------------------------------------
# 问卷字段的英文标签 / 帮助（引擎 schema 的中文 label 是单一真相，这里只做显示层映射）
# ---------------------------------------------------------------------------
Q_LABELS_EN = {
    "margin_top_in": "Top margin", "margin_bottom_in": "Bottom margin",
    "margin_left_in": "Left margin", "margin_right_in": "Right margin",
    "font_family": "Body font", "font_size_pt": "Body font size",
    "line_spacing": "Line spacing", "body_alignment": "Body alignment",
    "first_line_indent_in": "First-line indent", "space_after_pt": "Space after paragraph",
    "heading_levels": "Heading levels",
    "reference_style": "Reference style",
    "reference_hanging_indent_in": "Reference hanging indent",
    "reference_numbered": "Numbered references",
}

Q_HELP_EN = {
    "margin_top_in": "Top margin in inches (1 in ≈ 2.54 cm)",
    "margin_bottom_in": "Bottom margin in inches",
    "margin_left_in": "Left margin in inches (often wider for binding)",
    "margin_right_in": "Right margin in inches",
    "font_family": "e.g. Times New Roman / Arial / Calibri",
    "font_size_pt": "Body text is usually 10–12 pt",
    "line_spacing": "2.0 = double, 1.5 = one-and-a-half, 1.0 = single",
    "body_alignment": "left for APA/MLA, justify for IEEE",
    "first_line_indent_in": "0.5 in is common; 0 = no indent (IEEE)",
    "space_after_pt": "0 is common — paragraphs are told apart by the first-line indent",
    "heading_levels": "0 = not enforced; APA usually has 5 levels, others follow the school",
    "reference_style": "Keep it consistent with the citation style you picked above",
    "reference_hanging_indent_in": "0.5 in is common; 0 for IEEE",
    "reference_numbered": "Yes for numbered styles like IEEE [1]; No for author–date styles",
}

# 对齐取值的显示名（引擎给的原始值是 left/justify/center/right）
ALIGN_LABELS = {
    "en": {"left": "left", "justify": "justify", "center": "center", "right": "right"},
    "zh": {"left": "左对齐", "justify": "两端对齐", "center": "居中", "right": "右对齐"},
}


def questionnaire_label(lang: str, field: dict) -> str:
    """问卷字段标签：英文模式下用映射表，缺失则回退引擎自带的中文标签。"""
    if lang == "en":
        return Q_LABELS_EN.get(field["key"], field.get("label", field["key"]))
    return field.get("label", field["key"])


def questionnaire_help(lang: str, field: dict) -> str:
    """问卷字段的帮助文字（含单位）。

    引擎 schema 里的 unit 是中文（英寸 / 磅 / 倍 / 级），英文模式必须过一遍
    ``unit_label`` 换算，否则英文界面上会混出中文单位。
    """
    if lang == "en":
        unit = unit_label(lang, field.get("unit") or "")
        help_txt = Q_HELP_EN.get(field["key"], "")
        return " · ".join(p for p in (unit, help_txt) if p)
    hint = field.get("unit", "") or ""
    if field.get("help"):
        hint = (hint + " " + field["help"]).strip()
    return hint


def unit_label(lang: str, unit: str) -> str:
    """单位显示名（引擎 schema 里的单位是中文，英文模式需换算）。"""
    mapping = {"英寸": "unit_in", "磅": "unit_pt", "倍": "unit_x", "级": "unit_level"}
    key = mapping.get(unit)
    if key and lang == "en":
        return t(lang, key)
    return unit
