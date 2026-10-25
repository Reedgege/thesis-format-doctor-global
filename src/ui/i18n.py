"""界面文案 / 字体族 / 语言偏好（海外版：中英双语，**默认英文**）。

设计约束
--------
- 纯数据 + 纯函数，**不 import tkinter**：字体族名的平台兜底与词条取用都能离线单测
  （tkinter 只在 ``fonts.py`` 里用于构造真正的 Font 对象）。
- 所有面向用户的字符串集中在本模块，GUI 只按 key 取词 —— 禁止在界面代码里散落硬编码，
  否则加一种语言就要满仓库找字符串。
- **品牌口径（海外版铁律）**：只出现官网 ``paperformatpro.com`` 与邮箱 ``hi@paperformatpro.com``。
  **绝不出现微信 / 公众号 / 小程序**（海外用户不用这些，老板 2026-09-11 明确要求）；
   国内版遗留的第三方联系入口（二维码 / 客服号）在搬布局时一并剔除。

语言偏好落盘在 ``~/.paperformatpro/settings.json``（与 license.json 同目录），
可用环境变量 ``TFD_SETTINGS_FILE`` 覆盖 —— 与授权状态文件同样的隔离策略，便于测试。
"""

from __future__ import annotations

import json
import os
import platform
import shutil
from pathlib import Path

LANGS = ("en", "zh")
DEFAULT_LANG = "en"          # 海外版默认英文（老板 2026-09-11 定）

SETTINGS_FILE = Path.home() / ".paperformatpro" / "settings.json"

# v2.3.10 重命名：旧隐藏目录 .thesis-format-doctor-global 一次性迁移回退（见 license.py）。
SETTINGS_FILE_LEGACY = Path.home() / ".thesis-format-doctor-global" / "settings.json"

BRAND_EMAIL = "hi@paperformatpro.com"
BRAND_SITE = "paperformatpro.com"
BRAND_SITE_URL = "https://paperformatpro.com"
BRAND_BUY_URL = "https://paperformatpro.com/pricing"
BRAND_NAME = "ReedSkill"


def settings_file() -> Path:
    """实际设置文件路径（每次读 env，便于运行时切换 / 测试隔离）。

    v2.3.10 重命名：新目录 ``.paperformatpro`` 不存在但旧目录
    ``.thesis-format-doctor-global`` 的 settings.json 还在时，一次性复制到新路径，
    后续读写都走新路径 —— 老用户升级后语言偏好不丢。
    """
    env = os.environ.get("TFD_SETTINGS_FILE")
    if env:
        return Path(env)
    if not SETTINGS_FILE.exists() and SETTINGS_FILE_LEGACY.exists():
        try:
            SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(SETTINGS_FILE_LEGACY, SETTINGS_FILE)
        except Exception:
            return SETTINGS_FILE_LEGACY
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


def detect_lang() -> str:
    """返回当前生效语言（读取已保存的偏好，任何异常回退默认）。

    供 UI 组件在自建界面（如弹窗）时取「当前语言」用 —— 与 :func:`load_lang`
    同义；界面上所谓"detect"即按已存偏好决定显示语言，不读系统区域设置
    （海外版默认英文，用户用语言切换按钮改完即落盘，见 :func:`save_lang`）。
    """
    return load_lang()


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
    # 字号换算：规格书给的是浏览器 px，tkinter 的字号单位是磅（96dpi 下 1pt≈1.33px），
    # 所以 brand 32px→24pt / section 22px→16pt / body 14px→11pt / caption 12px→9pt。
    "en": {
        # 字号纪律（v2.3.0）：主视图要在 1280×720 整屏放下，字号按最小窗口倒推，不再
        # 照抄规格书的浏览器 px（tkinter 用磅，150% 缩放的机器上 1pt≈2px，直接照抄会撑爆）。
        "F_BRAND":        ("Georgia", 22, "bold"),      # 品牌标题
        "F_TITLE":        ("Georgia", 17, "bold"),      # 弹窗主标题
        "F_DIALOG_TITLE": ("Georgia", 14, "bold"),      # 弹窗标题
        "F_HDR":          ("Georgia", 12, "bold"),      # 弹窗章节标题
        "F_CARD_HDR":     ("Segoe UI", 14, "bold"),      # 卡片标题（无衬线，v2.3.4 视觉收敛）
        "F_STAT":         ("Segoe UI", 10),             # 状态行
        "F_ICON":         ("Georgia", 12, "bold"),      # 行首印记
        "F_SUB":          ("Segoe UI", 11),             # 品牌标语
        "F_LABEL":        ("Segoe UI", 10, "bold"),     # 字段标签
        "F_SUBTITLE":     ("Segoe UI", 10),             # 小标题 / 元信息
        "F_BODY":         ("Segoe UI", 10),             # 正文
        "F_HELP":         ("Segoe UI", 9),              # 字段说明 / 次要文字
        "F_SMALL":        ("Segoe UI", 9),              # 说明 / 次要（弹窗复用）
        "F_SMALL_B":      ("Segoe UI", 9, "bold"),      # 提示框标题
        "F_BTN":          ("Segoe UI", 11, "bold"),     # 主按钮
        "F_BTN_S":        ("Segoe UI", 10),             # 次按钮（描边，弱化）
        "F_STEP_N":       ("Segoe UI", 9, "bold"),      # 步骤圆点里的数字/勾
        "F_FOOT":         ("Segoe UI", 9),              # 页脚 / 状态栏
        "F_MONO":         ("Consolas", 9),              # 机器码 / 离线码
    },
    "zh": {
        "F_BRAND":        ("KaiTi", 16, "bold"),
        "F_TITLE":        ("KaiTi", 15, "bold"),
        "F_DIALOG_TITLE": ("KaiTi", 12, "bold"),
        "F_HDR":          ("KaiTi", 11, "bold"),
        "F_CARD_HDR":     ("Microsoft YaHei", 12, "bold"),
        "F_STAT":         ("Microsoft YaHei", 9),
        "F_ICON":         ("KaiTi", 11, "bold"),
        "F_SUB":          ("Microsoft YaHei", 10),
        "F_LABEL":        ("Microsoft YaHei", 9, "bold"),
        "F_SUBTITLE":     ("Microsoft YaHei", 9),
        "F_BODY":         ("Microsoft YaHei", 9),
        "F_HELP":         ("Microsoft YaHei", 8),
        "F_SMALL":        ("Microsoft YaHei", 8),
        "F_SMALL_B":      ("Microsoft YaHei", 8, "bold"),
        "F_BTN":          ("Microsoft YaHei", 10, "bold"),
        "F_BTN_S":        ("Microsoft YaHei", 9),
        "F_STEP_N":       ("Microsoft YaHei", 8, "bold"),
        "F_FOOT":         ("Microsoft YaHei", 8),
        "F_MONO":         ("Consolas", 8),
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


def resolve_font_spec(lang: str) -> dict:
    """按语言返回字体规格字典 ``{F_XXX: (family, size, weight)}``。

    供 GUI 组件直接用 tuple 形式传给 ``font=``（无需先建 ``tkinter.font.Font`` 对象）。
    返回形状与 :func:`font_spec` 一致；此处**不传** ``available``，family 用原始族名，
    由 tkinter 在渲染时按系统兜底（缺字落到系统默认，不会崩）。

    与 :func:`font_spec` 的分工：
    - 主界面用 ``fonts.Fonts`` 持有可缩放的命名 ``Font`` 对象（它走
      ``font_spec(lang, available)`` 做平台级精确兜底）；
    - 弹窗等一次性自建界面图省事，直接拿本函数返回的 tuple 规格喂 ``font=`` 即可。
    """
    return font_spec(lang)


# ---------------------------------------------------------------------------
# 界面文案
# ---------------------------------------------------------------------------
STRINGS: dict = {
    "en": {
        # —— 顶栏 / 标题区 ——
        # 规格书：隐私声明从卡片里**移到顶栏**（卡片右上角不再挂 Offline 徽标）。
        "link_help": "Help",
        "link_about": "About",
        "lang_button": "Chinese",
        "lang_button_tip": "Switches the interface language only",
        "app_title": "PaperFormat Pro",
        "app_tagline": "Academic formatting, made simple.",
        # 注：不用 🔒 之类的非 BMP emoji —— Tk 在 Windows 上常渲染成方框（tofu）。
        "privacy_line": "Your thesis never leaves your computer",
        "privacy_sub": "Private  ·  Secure  ·  Offline processing",

        # —— 左栏：文档选择 ——
        # 规格书第 4 节：Files → "Your Documents"、Choose Material → "Select your documents"、
        # Paper to Check → "Paper"、School Template → "University template"、
        # AI-filled 配置（YAML/JSON）→ "Import configuration"。
        "left_title": "Your Documents",
        "left_desc": "Select your document and formatting requirements. You can change these settings anytime before checking.",
        "mark_required": "Required",
        "mark_optional": "Optional",
        "row_spec_title": "Citation style",
        "row_spec_desc": "Choose the citation style required by your institution.",
        "row_template_title": "University template",
        "row_template_desc": "Use your university's official template (if available).",
        "row_template_placeholder": "Choose a template",
        # v2.3.64：两行标题下不再挂「会检查哪些格式项」的长说明
        # （原 row_questionnaire_desc / row_ai_desc 已删）—— 那串既列不全，
        # 又会把同行的按钮挤到裁字。要看明细请去复核弹窗。
        "row_questionnaire_title": "Questionnaire",
        "btn_questionnaire": "Fill in",
        "row_ai_title": "Import configuration",
        "btn_ai_import": "Import",
        "row_paper_title": "Paper",
        # ⚠️ 对外能力声明必须与实际一致（上架铁律）。引擎只读 .docx：
        #   src/engine/docx_reader.py 用 python-docx，打开文件的对话框也只收 *.docx；
        #   PDF 无任何解析路径（venv 里连 pdf 库都没装）。
        # 规格书整屏稿写的是 "PDF, DOCX (max 50 MB)"，那是**做不到的承诺**——
        # 照着写等于虚假宣传（用户选 PDF 会发现根本选不了），故此处按真实支持面写。
        "row_paper_helper": "Upload or select your document to check.",
        "paper_support": "Word document (.docx)",
        "drop_hint": "or drag and drop your file here",
        "drop_active": "Drop your document here",
        "btn_select_doc": "Select a document",
        "adv_label": "Advanced options",
        "adv_hint": "Questionnaire  ·  Import configuration  ·  AI assistance",
        "chip_style_hint": "“Other” means a custom template — pair it with a university template or questionnaire.",

        # 选中状态行（新布局里这些显示在字段控件内部：未选=占位灰字，已选=蓝色文件名）
        "st_paper_none": "No document selected",
        "st_paper_sample": "Thesis_Final.docx",
        "st_paper_error": "File type not supported",
        "st_tpl_none": "Choose a template",
        # 状态行用词（v2.3.64）：只回答「这行填没填、填了多少」，不罗列格式项。
        # ``*_n`` 词条走 i18n.t() 的复数约定：N=1 时自动改用 ``*_one`` 文案。
        "st_q_none": "Not filled in",
        "st_q_set_n": "Filled in · %s fields",
        "st_q_set_one": "Filled in · 1 field",
        "st_ai_none": "Not imported",
        "st_ai_set": "Imported · %s",

        # —— 右栏：Review & Fix ——
        # 规格书：Workflow → "Review & Fix"；Run Format Check → "Check formatting"；
        # One-click Fix → "Fix formatting"。步骤标签只留 Prepare / Check / Fix（去掉长描述与 1/3 计数）。
        "right_title": "Review & Fix",
        "right_desc": "Follow the steps to ensure your paper meets the required format.",
        "step1_title": "Prepare",
        "step2_title": "Check",
        "step3_title": "Fix",
        "how_title": "How it works",
        "how_steps": [
            ("Prepare", "Select your material and settings."),
            ("Check", "We'll analyze your paper for formatting issues."),
            ("Fix", "Review the results and apply fixes with one click."),
        ],
        "btn_check": "Check formatting ›",
        "btn_check_busy": "Checking…",
        "btn_fix": "Fix issues",
        "btn_fix_n": "Fix %s issues  ›",
        "btn_fix_one": "Fix 1 issue  ›",
        "btn_review": "Review changes",
        "trust_title": "Built for academic work",
        "trust_body": ("Formatting rules are applied locally and reviewed before changes are made."),
        "btn_save_report": "Save report",
        "btn_export_ai": "Export AI template",
        "btn_activate": "Activation",

        # —— 商业区（规格 COMMERCIAL UI：主页只放一条紧凑的试用/升级入口，
        #     点开才展示定价弹窗；定价/权益逻辑本身在 license 层，UI 只做呈现） ——
        "trial_fix_one": "Trial · %s check left",
        "trial_fix_many": "Trial · %s checks left",
        "trial_none": "Your trial is complete",
        "trial_licensed": "Licensed · unlimited checks",
        "trial_licensed_monthly": "Licensed · monthly plan",
        "trial_licensed_monthly_exp": "Licensed · monthly · expires %s",
        "trial_revoked": "License expired or revoked",
        # 四档授权（times/quarterly/annual/lifetime）状态行文案；%s 由 license_summary 注入
        "lic_times_left": "%s · %s checks left",
        "lic_times_left_one": "%s · %s check left",
        "lic_times_exhausted": "%s used up — buy another pack",
        "lic_exp": "%s · expires %s",
        "lic_exp_soon": "%s · expires %s (renew soon)",
        "lic_exp_soon_nd": "%s · renews soon",
        "lic_exp_expired": "%s expired — renew or buy again",
        "lic_exp_nd": "%s",
        # v2.3.16：授权层 gate 的 reason → 本地图文案（严禁直接显示 license 层硬编码中文）
        "gate_trial": "First free trial (up to %s fix)",
        "gate_activated": "Activated",
        "gate_trial_exhausted": "Your free trial is used up. Enter an activation code to unlock unlimited fixes.",
        "gate_needs_reactivation": "We detected an older license file (activated before this upgrade). Re-enter your activation code once to restore it: paste an offline code directly, or sign in online for an online code.",
        "gate_revoked": "Your license was revoked or expired. Please re-enter your activation code — or buy again to get a new one.",
        "gate_state_error": "Your license status could not be read or saved, so we can't start a fix right now. Please retry, or contact support if it keeps happening.",
        # 启动一次性核验（授权 + 版本）的状态机文案
        "launch_offline": "Network unavailable — fixes still work (offline mode).",
        "upd_title": "A new version is available",
        "upd_body": "Version %s is out (you have %s).",
        "upd_btn": "Update",
        "upd_later": "Later",
        # v2.3.16：右卡常驻「下一步」引导条（随状态机更新，治"向导性不够好"）
        "guide_pick": "Next: drop your .docx here, then click “Check formatting”.",
        "guide_check": "Next: click “Check formatting” to scan your document.",
        "guide_checking": "Scanning your document for formatting issues…",
        "guide_fix": "Found %s issues — click “Fix” to correct them automatically.",
        "guide_ok": "No issues found. Your document looks good.",
        "guide_fixed": "Done. Open the fixed file to review the changes.",
        "btn_upgrade": "Upgrade to continue →",
        "up_title": "Upgrade to continue",
        "up_subtitle": "Buy a license, or paste a code you already have.",
        "up_sku_line": "Two-use $5.99 · Quarterly $39 · Annual $79 · Lifetime $149",
        "up_buy": "Buy",
        "up_buy_email": "Email us to buy",
        "up_act_label": "Activation code",
        "up_offline_expand": "Offline activation (no internet)",
        "up_note": "Plans are confirmed by email before purchase — no account, no subscription charges inside the app.",
        "up_close": "Close",
        "btn_continue": "Continue",

        # —— 结果状态行（规格 states 段） ——
        "state_missing_paper": "Select your paper to start — it is the only required input.",
        "state_issues": "%s formatting issues found.",
        "state_issue_one": "1 formatting issue found.",
        "state_clean": "No formatting issues found.",
        "state_fixed": "All fixes applied — your original file is untouched.",
        "state_counts_hint": ("Pass %s · Info %s · Warn %s · Fail %s  —  "
                              "the full report is exported with “Save report”."),

        # —— 报告区 ——
        "report_placeholder": (
            "Your thesis never leaves your computer.\n\n"
            "No upload. No cloud processing. No waiting in a queue.\n\n"
            "Choose a citation style and your thesis, then check the formatting."
        ),
        "report_summary_fix": "Fix preview ready · %s change(s) — confirm in the dialog.",
        "report_summary_generic": "Done. Click “Save report” to export the full report.",
        "report_summary_hint": ("The full report is not shown here — click “Save report” "
                                "to export it as a file."),

        # —— 状态栏 ——
        "bar_left": "PaperFormat Pro · v%s",
        "bar_idle": "Ready · Your thesis never leaves your computer",
        "bar_checking": "Checking formatting…",
        "bar_fixing": "Applying format fixes…",
        "bar_check_done": "Formatting check complete",
        "bar_error": "Error",
        "status_checking": "Checking the format of your document…",
        "status_fixing": "Applying format fixes (content stays untouched)…",

        # —— 页脚（规格 footer 段：左邮箱 / 中官网 / 右标语，**只此三项**） ——
        # 旧版的「隐私声明 + © 2026」不在这里：隐私声明已上移到顶栏，规格要求页脚只有这三项。
        "footer_email": BRAND_EMAIL,
        "footer_site": BRAND_SITE,
        "footer_tagline": "Better formatting. Greater academic success.",

        # —— 通用提示 ——
        "msg_read_error_title": "Cannot read the document",
        "msg_check_error_title": "Check failed",
        "msg_check_error_unexpected": "Check failed (unexpected error)",
        "msg_fix_preview_error_title": "Preview failed",
        "msg_fix_error_title": "Fix failed",
        "msg_fix_error_unexpected": "Preview failed (unexpected error)",
        "msg_cannot_overwrite_title": "Cannot overwrite the original",
        "msg_cannot_overwrite": "The output path is the same file as your paper. Please choose another location.",
        "msg_no_change": ("No changes needed — the items in the report are advisory "
                          "for this style. Nothing was rewritten, no copy was created, "
                          "and no trial quota was used."),
        "msg_need_activation_title": "Activation required",
        "msg_fix_confirm_title": "Confirm the fix",
        "msg_fix_confirm_body": "\n\nClick Yes to apply the fixes, then choose where to save.\nYour original file will not be modified.",
        "msg_fix_pick_output": "Choose where to save the fixed copy (default: name_fixed.docx)",
        "msg_fix_done_title": "Fix complete",
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
        "q_cancel": "Cancel",
        # v2.3.64：问卷弹窗分三块（按 schema 的 group），不再 16 行一字排开
        "q_sec_page": "Page layout",
        "q_sec_heading": "Headings",
        "q_sec_refs": "References",
        "q_saved_title": "Saved",
        "q_saved": "Questionnaire saved. It will be used the next time you run.",
        "q_bad_number_title": "Not a number",
        "q_bad_number": "'%s' expects a number, but got: %s",
        "q_range_title": "Outside the usual range",
        "q_range_body": "These values are outside the usual range. Save anyway?\n\n%s",
        "unit_in": "in", "unit_pt": "pt", "unit_x": "x", "unit_level": "levels",
        "bool_yes": "Yes", "bool_no": "No",
        # v2.3.65：每级标题明细表（手填问卷与 AI 填表共用同一套字段）
        "q_h_table_hint": ("Per-level heading details. Leave a row at “—” and it stays unset "
                           "(nothing is assumed on your behalf)."),
        "q_h_lvl": "Level %s",
        "q_h_font": "Font",
        "q_h_size": "Size (pt)",
        "q_h_bold": "Bold",
        "q_h_italic": "Italic",
        "q_h_pagebreak": "Page break",
        "q_prefilled": ("Blank reference fields were pre-filled from %s — page setup and "
                        "headings are untouched and stay as you set them, and anything you "
                        "filled in yourself is never overwritten."),
        "q_h_upper": "UPPER",
        "q_h_runin": "Run-in",
        "q_h_align": "Align",
        "q_h_indent": "Indent (in)",
        "q_h_num": "Numbered",
        # v2.3.75：段前 / 段后 / 行距（标题级）。表头尽量短，表本身已 13 列。
        "q_h_before": "Sp. before",
        "q_h_after": "Sp. after",
        "q_h_lsp": "Line sp.",
        "q_none": "—",
        "q_align_left": "Left", "q_align_center": "Center",
        "q_align_right": "Right", "q_align_justify": "Justified",
        "q_prefix_number": "1.", "q_prefix_roman": "I.",
        "q_prefix_upper_letter": "A.", "q_prefix_lower_letter": "a)",
        "q_prefix_lower_roman": "i)", "q_prefix_number_paren": "1)",

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
        "about_privacy": "Everything runs on your device. Your thesis is never uploaded, never collected.",
        "about_source_hdr": "Where the rules come from",
        "about_source": ("Page layout follows your school's template; citation format follows the "
                         "style you pick (APA, MLA, IEEE, or Other)."),
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
             "APA, MLA or IEEE. Choose Other if your school uses something else."),
            ("Add your school's rules (optional but recommended)",
             "Upload the Word template your school handed out, fill in the questionnaire, or import a "
             "YAML config an AI filled for you. Page layout comes from here."),
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
            ("Q3 · Is my thesis uploaded anywhere?",
             "No. Your thesis never leaves this computer — checking and fixing run entirely on your "
             "machine (no upload, works with the network unplugged). The app contacts the server only "
             "for activation and for the one-time free-trial registration, and sends the machine code "
             "only — never your document."),
            ("Q4 · It says activation is required. How do I activate?",
             "The first fix is free. After that, open Activation: either paste an online activation "
             "code, or copy your machine code, email it to %s, and paste back the offline code "
             "we send you." % BRAND_EMAIL),
            ("Q5 · Windows warns me when I open the app.",
             "The build is not code-signed yet, so SmartScreen flags it. Click More info → Run anyway. "
             "The executable is built from the public source of this project."),
        ],

        # —— 反馈窗口（规格：顶栏「Feedback」入口；Worker /api/feedback → 邮件） ——
        "link_feedback": "Feedback",
        "fb_title": "Feedback",
        "fb_intro": "Found a bug, or have an idea? Tell us — we read every message.",
        "fb_email": "Your email (so we can reply)",
        "fb_email_hint": "Optional, but we can't reply without it.",
        "fb_message": "Your message",
        "fb_message_hint": "What went wrong, or what you'd like to see.",
        "fb_privacy": "We only use your email to reply to this feedback.",
        "fb_submit": "Send feedback",
        "fb_cancel": "Cancel",
        "fb_invalid_email": "Please enter a valid email address.",
        "fb_empty": "Please write a few words first.",
        "fb_sending": "Sending…",
        "fb_sent": "Thanks! Your feedback is on its way.",
        "fb_sent_title": "Feedback sent",

        # —— 最近使用（v2.3.78）：手填问卷模板 / AI 导入文件，落盘 ~/.paperformatpro/recents.json ——
        "recents_manual_label": "Recent templates",
        "recents_ai_label": "Recent imports",
        "recents_none": "No recent items",
        "recents_removed": "File no longer exists — removed from history",
        "fb_failed_title": "Couldn't send",
        "fb_mailto_fallback": ("Couldn't reach the server. We've opened your email app with a "
                               "pre-filled message — just press send."),

        # —— 格式来源（7 选 1，**互斥**：选中哪一项，界面只显示那一项对应的输入）——
        "source_title": "Format source",
        "source_helper": "Pick one. Only the settings for that choice will show below.",
        "src_apa_student": "APA 7th — Student (Word)",
        "src_apa_pro": "APA 7th — Professional (Word)",
        "src_mla": "MLA 9th — Standard (Word)",
        "src_ieee": "IEEE Conference (two-column)",
        "src_ieee_journal": "IEEE Journal (editorial rules)",
        "src_school": "School Template",
        "src_manual": "Fill in questionnaire (manual)",
        "src_ai": "Import configuration (AI)",
        "adv_school_title": "School Template",
        "adv_school_desc": "Import your university's .docx template; we read its real styles.",
        "school_ref_title": "Reference / citation style",
        "school_ref_none": "Do not check references",
        "review_title": "Confirm your format settings",
        "review_helper": "We read the values below. Review them, correct any gap, then run the check.",
        "review_page": "Page Layout",
        "review_columns": "Columns",
        "review_paper_size": "Paper size",
        "review_page_number": "Page number",
        "review_one_column": "Single column",
        "review_two_column": "Two columns",
        "review_headings": "Headings",
        "review_refs": "References",
        "review_frontmatter": "Abstract / Running Head / Captions",
        "review_style": "Style Rules",
        "review_confirm": "Run check",
        "review_back": "Back",
        "review_missing": "Not detected — confirm or fill in",
        # v2.3.94：客户选了「本身不定义分级编号标题」的内置规范（MLA 9th）。
        # 以前这里复用了 review_missing，等于说「没检测出来，你来填」——
        # 让客户去补一项规范根本没要求的东西。{name} 填规范显示名。
        "review_headings_not_required":
            "{name} does not require numbered heading levels — use one consistent "
            "style of your own. Nothing to fill in here.",
        # v2.3.94：这个弹窗里标题表是**只读**的，所以「confirm or fill in」是在
        # 指点一个它自己做不到的动作。改成说清「去哪儿填」（Advanced options 里的
        # 手动问卷 / 换一个定义了标题的来源）。
        "review_headings_missing":
            "Not detected — this source sets no heading levels. If your school "
            "requires them, fill them in under Advanced options.",
        # v2.3.102（老板 2026-10-24 截图）：规范**要求**分级标题、但**不规定各级字形**
        # （IEEE-Journal 按裁定「不碰版式」，各级格式交给目标期刊模板）。
        # 以前这里复用 review_headings_missing → 说成 "this source sets no heading levels"，
        # 与事实相反（Journal 是 heading_levels=4），客户会以为自己要补标题层级。
        "review_headings_levels_only":
            "{name} requires {n} heading levels, but does not dictate each level's "
            "size / alignment / numbering — your target template does. Nothing to fill in here.",
        # v2.3.103（老板 2026-10-24 追问三条通道）：手填 / AI 已经给了层级数、
        # 但没填各级明细（heading_formats 为空）。以前也落到 review_headings_missing
        # → 说成 "this source sets no heading levels"，与 target.heading_levels=N 矛盾。
        "review_headings_levels_no_formats":
            "{n} heading levels are set, but their per-level formatting is not. "
            "Fill in the per-level details under Advanced options if you need them.",
        "review_lvl": "Level",
        "review_align": "Align",
        "review_bold": "Bold",
        "review_italic": "Italic",
        "review_size": "Size",
        "review_prefix": "Prefix",
        "review_upper": "Upper",
        "review_enabled": "On",
        "review_disabled": "Off",
    },

    "zh": {
        "link_help": "帮 助",
        "link_about": "关 于",
        "lang_button": "英文",
        "lang_button_tip": "仅切换界面语言，不影响排版格式",
        # 中文名 = 文档格式先生（英文 Mr. Document Formatter / 海外 brand PaperFormat Pro）。
        # ⚠️ 命名铁律：「论文格式医生」是**国内论文专用**产品（tfd-desktop / tfd-mentor）的名字，
        # 本海外版是通用文档格式产品，中文界面一律写「文档格式先生」，不得混用。
        "app_title": "文 档 格 式 先 生",
        "app_tagline": "学术格式，本可以更简单。",
        "privacy_line": "论文不离开这台电脑",
        "privacy_sub": "私密  ·  安全  ·  本机离线处理",

        "left_title": "你的文档",
        "left_desc": "选择文档与格式要求。体检前随时可以修改这些设置。",
        "mark_required": "必选",
        "mark_optional": "可选",
        "row_spec_title": "引用规范",
        "row_spec_desc": "选择你所在院校要求的引用规范。",
        "row_template_title": "学校模板",
        "row_template_desc": "如果有学校官方模板，用它更准。",
        "row_template_placeholder": "选择模板",
        "row_questionnaire_title": "格式问卷",
        "btn_questionnaire": "手填问卷",
        "row_ai_title": "导入配置",
        "btn_ai_import": "导入",
        "row_paper_title": "论文",
        "row_paper_helper": "上传或选择要体检的文档。",
        # 同 EN：只写真实支持面（引擎只读 .docx），不照抄整屏稿里做不到的承诺。
        "paper_support": "Word 文档（.docx）",
        "drop_hint": "或将文件拖拽到此处",
        "drop_active": "松开即可上传文档",
        "btn_select_doc": "选择文档",
        "adv_label": "高级选项",
        "adv_hint": "格式问卷  ·  导入配置  ·  AI 辅助",
        "chip_style_hint": "「Other」= 自定义规范；请配学校模板或格式问卷。",

        "st_paper_none": "未选择文档",
        "st_paper_sample": "Thesis_Final.docx",
        "st_paper_error": "文件类型不支持",
        "st_tpl_none": "选择模板",
        "st_q_none": "未填写",
        # 与 EN 同构：``st_q_set_n`` / ``st_q_set_one`` 交给 t() 的复数约定挑一条。
        "st_q_set_n": "已填写 · %s 项",
        "st_q_set_one": "已填写 · 1 项",
        "st_ai_none": "未导入",
        "st_ai_set": "已导入 · %s",

        # —— 格式来源（6 选 1 互斥，海外版零中文：zh 亦用英文值）——
        "source_title": "Format source",
        "source_helper": "Pick one. Only the settings for that choice will show below.",
        "src_apa_student": "APA 7th — Student (Word)",
        "src_apa_pro": "APA 7th — Professional (Word)",
        "src_mla": "MLA 9th — Standard (Word)",
        "src_ieee": "IEEE Conference (two-column)",
        "src_ieee_journal": "IEEE Journal (editorial rules)",
        "src_school": "School Template",
        "src_manual": "Fill in questionnaire (manual)",
        "src_ai": "Import configuration (AI)",
        "adv_school_title": "School Template",
        "adv_school_desc": "Import your university's .docx template; we read its real styles.",
        "school_ref_title": "Reference / citation style",
        "school_ref_none": "Do not check references",
        "review_title": "Confirm your format settings",
        "review_helper": "We read the values below. Review them, correct any gap, then run the check.",
        "review_page": "页面设置",
        "review_columns": "栏数",
        "review_paper_size": "纸张大小",
        "review_page_number": "页码位置",
        "review_one_column": "单栏",
        "review_two_column": "Two columns",
        "review_headings": "Headings",
        "review_refs": "References",
        "review_frontmatter": "Abstract / Running Head / Captions",
        "review_style": "Style Rules",
        "review_confirm": "Run check",
        "review_back": "Back",
        "review_missing": "Not detected — confirm or fill in",
        # v2.3.94：客户选了「本身不定义分级编号标题」的内置规范（MLA 9th）。
        # 以前这里复用了 review_missing，等于说「没检测出来，你来填」——
        # 让客户去补一项规范根本没要求的东西。{name} 填规范显示名。
        "review_headings_not_required":
            "{name} does not require numbered heading levels — use one consistent "
            "style of your own. Nothing to fill in here.",
        # v2.3.94：这个弹窗里标题表是**只读**的，所以「confirm or fill in」是在
        # 指点一个它自己做不到的动作。改成说清「去哪儿填」（Advanced options 里的
        # 手动问卷 / 换一个定义了标题的来源）。
        "review_headings_missing":
            "Not detected — this source sets no heading levels. If your school "
            "requires them, fill them in under Advanced options.",
        # v2.3.102（老板 2026-10-24 截图）：规范**要求**分级标题、但**不规定各级字形**
        # （IEEE-Journal 按裁定「不碰版式」，各级格式交给目标期刊模板）。
        # 以前这里复用 review_headings_missing → 说成 "this source sets no heading levels"，
        # 与事实相反（Journal 是 heading_levels=4），客户会以为自己要补标题层级。
        "review_headings_levels_only":
            "{name} requires {n} heading levels, but does not dictate each level's "
            "size / alignment / numbering — your target template does. Nothing to fill in here.",
        # v2.3.103（老板 2026-10-24 追问三条通道）：手填 / AI 已经给了层级数、
        # 但没填各级明细（heading_formats 为空）。以前也落到 review_headings_missing
        # → 说成 "this source sets no heading levels"，与 target.heading_levels=N 矛盾。
        "review_headings_levels_no_formats":
            "{n} heading levels are set, but their per-level formatting is not. "
            "Fill in the per-level details under Advanced options if you need them.",
        "review_lvl": "Level",
        "review_align": "Align",
        "review_bold": "Bold",
        "review_italic": "Italic",
        "review_size": "Size",
        "review_prefix": "Prefix",
        "review_upper": "Upper",
        "review_enabled": "On",
        "review_disabled": "Off",

        "right_title": "检查与修正",
        "right_desc": "按步骤操作，确保论文符合要求的格式。",
        "step1_title": "准备",
        "step2_title": "体检",
        "step3_title": "修正",
        "how_title": "使用流程",
        "how_steps": [
            ("准备", "选择材料与设置。"),
            ("体检", "分析论文的格式问题。"),
            ("修正", "查看结果，一键应用修正。"),
        ],
        "btn_check": "运行格式体检 ›",
        "btn_check_busy": "体检中…",
        "btn_fix": "修正问题",
        "btn_fix_n": "修正 %s 个问题  ›",
        "btn_fix_one": "修正 1 个问题  ›",
        "btn_review": "查看改动",
        "trust_title": "为学术写作而做",
        "trust_body": "格式规则在本机应用，改动前均经过复核。",
        "btn_save_report": "保存报告",
        "btn_export_ai": "导出 AI 问卷模板",
        "btn_activate": "激活 / 授权",

        "trial_fix_one": "试用 · 还剩 %s 次检查",
        "trial_fix_many": "试用 · 还剩 %s 次检查",
        "trial_none": "试用已结束",
        "trial_licensed": "已授权 · 不限次检查",
        "trial_licensed_monthly": "已授权 · 月卡",
        "trial_licensed_monthly_exp": "已授权 · 月卡 · %s 到期",
        "trial_revoked": "授权已过期或已撤销",
        # 四档授权状态行文案（与英文一一对应）；%s 由 license_summary 注入
        "lic_times_left": "%s · 还剩 %s 次",
        "lic_times_left_one": "%s · 还剩 %s 次",
        "lic_times_exhausted": "%s 已用完 —— 请再购买一包",
        "lic_exp": "%s · %s 到期",
        "lic_exp_soon": "%s · %s 即将到期（请续费）",
        "lic_exp_soon_nd": "%s · 即将续费",
        "lic_exp_expired": "%s 已到期 —— 请续费或再次购买",
        "lic_exp_nd": "%s",
        # v2.3.16：授权层 gate 的 reason → 本地图文案（严禁直接显示 license 层硬编码中文）
        "gate_trial": "首次免费试用（共 %s 次修正）",
        "gate_activated": "已激活",
        "gate_trial_exhausted": "免费试用已用完，请输入激活码解锁无限修正。",
        "gate_needs_reactivation": "检测到旧版授权文件（本次升级前激活的）。请重新输入一次激活码即可恢复：离线码直接粘贴，在线码需联网。",
        "gate_revoked": "授权已被吊销或已过期，请重新输入激活码 —— 或再次购买获取新码。",
        "gate_state_error": "授权状态读取或写入失败，暂时无法进行修正。请重试；若持续出现请联系客服。",
        # 启动一次性核验（授权 + 版本）的状态机文案
        "launch_offline": "未能联网 —— 修正仍可使用（离线模式）。",
        "upd_title": "有新版本可用",
        "upd_body": "已发布 %s（你当前是 %s）。",
        "upd_btn": "更新",
        "upd_later": "以后再说",
        # v2.3.16：右卡常驻「下一步」引导条（随状态机更新，治"向导性不够好"）
        "guide_pick": "下一步：把 Word 文档拖到这里，然后点击「检查格式」。",
        "guide_check": "下一步：点击「检查格式」开始扫描文档。",
        "guide_checking": "正在扫描文档中的格式问题…",
        "guide_fix": "发现 %s 处问题 —— 点击「修正」即可自动校正。",
        "guide_ok": "未发现问题，文档格式看起来不错。",
        "guide_fixed": "已完成。打开修正后的文件查看改动。",
        "btn_upgrade": "升级以继续使用 →",
        "up_title": "升级以继续使用",
        "up_subtitle": "购买授权，或粘贴你已有的激活码。",
        "up_sku_line": "次卡 $5.99 · 季卡 $39 · 年卡 $79 · 终身 $149",
        "up_buy": "购买",
        "up_buy_email": "邮件购买",
        "up_act_label": "激活码",
        "up_offline_expand": "离线激活（断网时用）",
        "up_note": "方案价格由邮件确认后开通 —— 软件内不需要账号、不自动扣费。",
        "up_close": "关闭",
        "btn_continue": "继续",

        "state_missing_paper": "请先选择论文 —— 这是唯一必填项。",
        "state_issues": "发现 %s 个格式问题。",
        "state_issue_one": "发现 1 个格式问题。",
        "state_clean": "没有发现格式问题。",
        "state_fixed": "修正已全部应用 —— 原文件未被改动。",
        "state_counts_hint": ("通过 %s · 提示 %s · 警告 %s · 不达标 %s  ——  "
                              "完整报告用「保存报告」导出。"),

        "report_placeholder": (
            "选好引用规范与论文，点「运行格式体检」即可。\n\n"
            "全程在本机完成 —— 不上传、不需要账号；论文不离开这台电脑。"
        ),
        "report_summary_fix": "修正预览已就绪 · 共 %s 处改动，请在弹窗中确认。",
        "report_summary_generic": "已完成。点「保存报告」导出完整报告。",
        "report_summary_hint": "完整报告不在界面展开 —— 点「保存报告」即可导出为文件。",

        "bar_left": "PaperFormat Pro · v%s",
        "bar_idle": "就绪 · 论文始终留在你的电脑上",
        "bar_checking": "正在体检格式…",
        "bar_fixing": "正在应用格式修正…",
        "bar_check_done": "体检完成",
        "bar_error": "出错了",
        "status_checking": "正在检查文档格式…",
        "status_fixing": "正在应用格式修正（正文内容不会改动）…",

        "footer_email": BRAND_EMAIL,
        "footer_site": BRAND_SITE,
        "footer_tagline": "更好的格式，更好的学术成绩。",

        "msg_read_error_title": "无法读取文档",
        "msg_check_error_title": "检查失败",
        "msg_check_error_unexpected": "检查失败（未预期错误）",
        "msg_fix_preview_error_title": "预览失败",
        "msg_fix_error_title": "修正失败",
        "msg_fix_error_unexpected": "预览失败（未预期错误）",
        "msg_cannot_overwrite_title": "不能覆盖原件",
        "msg_cannot_overwrite": "输出路径与待修正文档是同一个文件，请另选保存位置。",
        "msg_no_change": "无需修改 —— 报告里的条目对该规范只是提示性质；未改写文档、未生成副本、也未消耗试用次数。",
        "msg_need_activation_title": "需要激活",
        "msg_fix_confirm_title": "确认修正",
        "msg_fix_confirm_body": "\n\n点「是」开始修正，随后选择保存位置（原件不会被改动）。",
        "msg_fix_pick_output": "选择修正稿保存位置（默认 原名_fixed.docx）",
        "msg_fix_done_title": "修正完成",
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
        "q_cancel": "取消",
        "q_sec_page": "页面排版",
        "q_sec_heading": "标题层级",
        "q_sec_refs": "参考文献",
        "q_saved_title": "已保存",
        "q_saved": "问卷已填入，运行体检时生效。",
        "q_bad_number_title": "数值格式不对",
        "q_bad_number": "「%s」需要填数字，当前是：%s",
        "q_range_title": "数值超出常见范围",
        "q_range_body": "以下数值超出常见范围，是否仍要保存？\n\n%s",
        "unit_in": "英寸", "unit_pt": "磅", "unit_x": "倍", "unit_level": "级",
        "bool_yes": "是", "bool_no": "否",
        "q_h_table_hint": "每级标题的明细。留「—」即不设定（不会替您猜一个值）。",
        "q_h_lvl": "第 %s 级",
        "q_h_font": "字体",
        "q_h_size": "字号（磅）",
        "q_h_bold": "加粗",
        "q_h_italic": "斜体",
        "q_h_pagebreak": "分页",
        "q_prefilled": "参考文献相关空白项已按 %s 预填；页面与标题不受影响，仍按您填的走（您自己填过的不会被覆盖）。",
        "q_h_upper": "全大写",
        "q_h_runin": "接排",
        "q_h_align": "对齐",
        "q_h_indent": "缩进（英寸）",
        "q_h_num": "编号",
        "q_h_before": "段前(磅)",
        "q_h_after": "段后(磅)",
        "q_h_lsp": "行距(倍)",
        "q_none": "—",
        "q_align_left": "左对齐", "q_align_center": "居中",
        "q_align_right": "右对齐", "q_align_justify": "两端对齐",
        "q_prefix_number": "1.", "q_prefix_roman": "I.",
        "q_prefix_upper_letter": "A.", "q_prefix_lower_letter": "a)",
        "q_prefix_lower_roman": "i)", "q_prefix_number_paren": "1)",

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
        "about_source": "页面排版以学校模板为准；引用格式以所选规范为准（APA / MLA / IEEE / Other）。",
        "about_smartscreen": "未签名程序 —— 首次打开时 Windows SmartScreen 可能拦截，点「详细信息」→「仍要运行」即可。",
        "about_contact_hdr": "联系我们",
        "about_footer": "© 2026 %s" % BRAND_NAME,

        "help_title": "使用帮助",
        "help_intro": "把学校的格式要求告诉它，导入论文，剩下的交给它。",
        "help_sec1": "一、怎么用",
        "help_sec2": "二、常见问题",
        "help_steps": [
            ("1. 选择引用规范",
             "APA、MLA、IEEE 或 Other。学校用别的，就选 Other。"),
            ("2. 告诉它学校的格式要求（可选，但强烈建议）",
             "上传学校下发的 Word 模板、手填问卷、导入 AI 填好的配置（YAML）。"
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

        # —— 反馈窗口（规格：顶栏「Feedback」入口；Worker /api/feedback → 邮件） ——
        "link_feedback": "反 馈",
        "fb_title": "意见反馈",
        "fb_intro": "遇到 bug，或有什么好点子？告诉我们 —— 每一条留言我们都会看。",
        "fb_email": "你的邮箱（方便我们回复）",
        "fb_email_hint": "选填，但没它我们没法回复你。",
        "fb_message": "反馈内容",
        "fb_message_hint": "哪里不对、或你希望加什么功能。",
        "fb_privacy": "我们仅用你的邮箱回复这条反馈。",
        "fb_submit": "提交反馈",
        "fb_cancel": "取消",
        "fb_invalid_email": "请填写有效的邮箱地址。",
        "fb_empty": "请先写几句反馈。",
        "fb_sending": "提交中…",
        "fb_sent": "已收到，谢谢你的反馈！",
        "fb_sent_title": "反馈已发送",

        # —— 最近使用（v2.3.78）——
        "recents_manual_label": "最近模板",
        "recents_ai_label": "最近导入",
        "recents_none": "暂无记录",
        "recents_removed": "文件已不存在，已从历史移除",
        "fb_failed_title": "发送失败",
        "fb_mailto_fallback": "没能连上服务器，已为你打开邮件客户端并填好内容 —— 直接点发送即可。",
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
    # 复数约定：``*_n`` 词条在第一个参数为 1 时优先用同族 ``*_one`` 文案。
    # 这样 "Fix %s issues" 在 N=1 时自动变 "Fix 1 issue"，而调用方（与回归测试）
    # 仍然只需认 ``*_n`` 一个 key，不必在业务代码里分散判断单复数。
    if args and key.endswith("_n") and args[0] == 1:
        table = STRINGS.get(lang, {})
        val = table.get(key[:-2] + "_one",
                        STRINGS[DEFAULT_LANG].get(key[:-2] + "_one", val))
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
    "first_line_indent_in": "First-line indent",
    "space_before_pt": "Space before paragraph",
    "space_after_pt": "Space after paragraph",
    "header_distance_in": "Header distance from top",
    "columns": "Columns",
    "table_space_before_pt": "Table cell space before",
    "table_space_after_pt": "Table cell space after",
    "table_gap_pt": "Spacing between tables and the text",
    # v2.3.100（F-2 / 裁定 C-2）：table_gap 的三档阈值另两个参数
    "table_gap_min_pt": "Minimum spacing between tables and the text",
    "table_gap_ratio": "Fraction of the recommended spacing that still passes",
    "toc_heading_space_before_pt": "Contents heading space before",
    "toc_heading_space_after_pt": "Contents heading space after",
    "toc_entry_space_before_pt": "Contents entry space before (check only)",
    "toc_entry_space_after_pt": "Contents entry space after (check only)",
    "paper_size": "Paper size",
    "page_number_position": "Page number position",
    "title_page_title_centered": "Title page: title centered",
    "title_page_title_bold": "Title page: title bold",
    "title_page_author_centered": "Title page: author lines centered",
    "title_page_space_before_pt": "Title page: space before title",
    "title_page_space_after_pt": "Title page: space after title",
    "title_page_new_page": "Title page on its own page",
    "title_page_zero_extra_spacing": "Title page: no extra paragraph spacing",
    "text_new_page": "Body starts on a new page",
    "heading_levels": "Heading levels",
    "reference_style": "Reference style",
    "reference_hanging_indent_in": "Reference hanging indent",
    "reference_numbered": "Numbered references",
    "reference_font_family": "Reference entry font",
    "reference_size_pt": "Reference entry font size",
    "reference_line_spacing": "Reference entry line spacing",
    "reference_space_before_pt": "Space before entry",
    "reference_space_after_pt": "Space after entry",
    "reference_new_page": "References on a new page",
    "reference_heading_bold": "Reference heading bold",
    "reference_heading_centered": "Reference heading centered",
    "abstract_first_line_indent_in": "Abstract first-line indent",
    "abstract_new_page": "Abstract on a new page",
    "abstract_space_before_pt": "Abstract: space before paragraph",
    "abstract_space_after_pt": "Abstract: space after paragraph",
    "abstract_heading_alignment": "Abstract heading alignment",
    "abstract_heading_bold": "Abstract heading bold",
    "abstract_heading_italic": "Abstract heading italic",
    "keywords_label_italic": "Keywords label italic",
    "keywords_indent_in": "Keywords line indent",
    "running_head_required": "Running head required",
    "page_number_include_author_name": "Page number includes author name",
    "caption_font_size_pt": "Caption font size",
    "caption_line_spacing": "Caption line spacing",
    "caption_space_before_pt": "Caption: space before",
    "caption_space_after_pt": "Caption: space after",
    "caption_position_table": "Table caption position",
    "caption_position_figure": "Figure caption position",
    "table_lines": "Table rules (three-line)",
}
Q_HELP_EN = {
    "table_lines": "three_line = academic three-line table (top rule / rule under the header "
                   "/ bottom rule, no vertical or inner rules). Leave blank = not checked.",
    "margin_top_in": "Top margin in inches (1 in ≈ 2.54 cm)",
    "margin_bottom_in": "Bottom margin in inches",
    "margin_left_in": "Left margin in inches (often wider for binding)",
    "margin_right_in": "Right margin in inches",
    "font_family": "e.g. Times New Roman / Arial / Calibri",
    "font_size_pt": "Body text is usually 10–12 pt",
    "line_spacing": "2.0 = double, 1.5 = one-and-a-half, 1.0 = single",
    "body_alignment": "left for APA/MLA, justify for IEEE",
    "first_line_indent_in": "0.5 in is common; 0 = no indent (IEEE)",
    "space_before_pt": "0 is common; fill in the template value if the school specifies one",
    "space_after_pt": "0 is common — paragraphs are told apart by the first-line indent",
    "header_distance_in": "Distance from the page top to the header; leave blank if unspecified (MLA 9 = 0.5)",
    "columns": "1 = single column (most formats); 2 = two-column (IEEE conference/journal)",
    "table_space_before_pt": "Space above each paragraph inside table cells; 0 is common "
                             "(no blank line inside a table). Leave blank to skip",
    "table_space_after_pt": "Space below each paragraph inside table cells; 0 is common. "
                            "Leave blank to skip",
    "table_gap_pt": "Space kept between a table and the text above/below it. APA wants an "
                    "extra double-spaced line (24 pt for 12 pt text) so tables never touch "
                    "the surrounding paragraphs. Leave blank to skip",
    # v2.3.100（F-2 / 裁定 C-2）：table_gap 的三档阈值另两个参数
    "table_gap_min_pt": "Hard floor: only a gap BELOW this value is reported as a problem. "
                        "APA's official samples use 12 pt, so 12 is the usual value. "
                        "Leave blank to skip",
    "table_gap_ratio": "A gap counts as acceptable once it reaches this fraction of the "
                       "recommended value (APA uses 0.8, i.e. 24 x 0.8 = 19.2 pt). "
                       "Leave blank to skip",
    "toc_heading_space_before_pt": "Space above the 'Contents' heading paragraph, in points. "
                                   "Leave blank to skip",
    "toc_heading_space_after_pt": "Space below the 'Contents' heading paragraph, in points. "
                                  "Leave blank to skip",
    "toc_entry_space_before_pt": "Space between TOC entries. CHECK ONLY — the entries are "
                                 "generated by Word, so change the TOC 1-9 styles in Word instead",
    "toc_entry_space_after_pt": "Space between TOC entries. CHECK ONLY — same reason as above",
    "paper_size": "Letter = 8.5 x 11 in (North America); A4 = 210 x 297 mm "
                  "(Europe, China, Australia). Leave it as — to skip this check",
    "page_number_position": "top-right = top-right corner of the HEADER (APA/MLA); "
                            "bottom-center = centered in the footer (common in Chinese "
                            "theses). Leave as — to skip",
    "title_page_title_centered": "Should the thesis title on the title page be centered "
                                  "(APA/IEEE yes). Leave as — to skip",
    "title_page_title_bold": "Should the title be bold (APA/IEEE yes, MLA explicitly not). "
                             "Leave as — to skip",
    "title_page_author_centered": "Should your name / department / affiliation lines be "
                                  "centered (APA/IEEE yes, MLA no). Leave as — to skip",
    "title_page_space_before_pt": "Space above the title on the title page, in points; "
                                  "leave blank to skip",
    "title_page_space_after_pt": "Space below the title on the title page, in points; "
                                 "leave blank to skip",
    "title_page_new_page": "Yes = the title page occupies its own page (APA student papers "
                           "require this); No = it may run on; — = unspecified",
    "title_page_zero_extra_spacing": "Yes = no extra paragraph spacing between the author, "
                                     "affiliation, course, instructor and date lines (APA "
                                     "student papers require this); — = unspecified",
    "text_new_page": "Yes = the body starts on a new page after the abstract (APA requires "
                     "this); No = it may continue right after the abstract; — = unspecified",
    "heading_levels": "0 = not enforced; APA usually has 5 levels, others follow the school",
    "reference_style": ("'My school's own' = apply only the entry rules you set here; "
                        "picking APA/MLA/IEEE checks the reference list only and never "
                        "touches page setup or headings"),
    "reference_hanging_indent_in": "0.5 in is common; 0 for IEEE",
    "reference_numbered": "Yes for numbered styles like IEEE [1]; No for author–date styles",
    "reference_font_family": "Leave blank to follow the body font; schools often require "
                             "the entries to use one specific font",
    "reference_size_pt": "Leave blank to follow the body text; entries are often 10.5 pt or 12 pt",
    "reference_line_spacing": "Leave blank to follow the body text; 1.0 = single (common for entries)",
    "reference_space_before_pt": "0 is common (no blank line between entries); leave blank if unspecified",
    "reference_space_after_pt": "0 is common (no blank line between entries); leave blank if unspecified",
    "reference_new_page": "Yes = insert a page break before the reference heading; No = not required; — = unspecified",
    "reference_heading_bold": "Yes = bold; No = explicitly not bold (MLA); — = unspecified",
    "reference_heading_centered": "Yes = centered; No = explicitly left-aligned (IEEE); — = unspecified",
    "abstract_first_line_indent_in": "Abstract body first-line indent; APA abstracts are flush left (0). Leave blank to skip",
    "abstract_new_page": "Yes = page break before the Abstract heading (APA Professional); No = not required; — = unspecified",
    "abstract_space_before_pt": "Space above each Abstract paragraph, in points; APA adds none (0). Leave blank to skip",
    "abstract_space_after_pt": "Space below each Abstract paragraph, in points; APA adds none (0). Leave blank to skip",
    "abstract_heading_alignment": "Alignment of the word 'Abstract' itself: center = APA style; left = most engineering templates. — = unspecified",
    "abstract_heading_bold": "Yes = the 'Abstract' line is bold (APA requires this); No = explicitly not bold; — = unspecified",
    "abstract_heading_italic": "Yes = the 'Abstract' line is italic (some IEEE templates); No = explicitly not italic; — = unspecified",
    "keywords_label_italic": "Yes = only the word 'Keywords:' is italic and the terms stay upright (APA requires this); No = explicitly not italic; — = unspecified",
    "keywords_indent_in": "First-line indent of the whole Keywords line; APA uses 0.5 in. Leave blank to skip",
    "running_head_required": "Yes = a running head in the header (short title or last name). Leave as — to skip",
    "page_number_include_author_name": "Yes = the page number is preceded by the author's last name (MLA's 'Smith 1'). — = unspecified",
    "caption_font_size_pt": "Figure/table caption font size; leave blank to follow the body text",
    "caption_line_spacing": "Figure/table caption line spacing; leave blank to follow the body text",
    "caption_space_before_pt": "Space above the caption, in points (often 6–12 pt to separate "
                               "it from the body); leave blank to skip",
    "caption_space_after_pt": "Space below the caption, in points; leave blank to skip",
    "caption_position_table": "Table caption above (top) or below (bottom) the table; APA/IEEE use top. Leave as — to skip",
    "caption_position_figure": "Figure caption above (top) or below (bottom) the figure; APA/IEEE use bottom. Leave as — to skip",
}

# 对齐取值的显示名（引擎给的原始值是 left/justify/center/right）
ALIGN_LABELS = {
    "en": {"left": "left", "justify": "justify", "center": "center", "right": "right"},
    "zh": {"left": "左对齐", "justify": "两端对齐", "center": "居中", "right": "右对齐"},
}


def questionnaire_label(lang: str, field: dict) -> str:
    """Questionnaire field label.

    In English mode the English map (Q_LABELS_EN) is the single source of truth;
    if a key is somehow missing it falls back to the English key (never the raw
    Chinese label baked into the schema), so no Chinese can surface in the UI.
    """
    if lang == "en":
        return Q_LABELS_EN.get(field["key"], field["key"])
    return field.get("label", field["key"])


# 问卷 choice 类字段的**选项展示名**（引擎值 → 界面显示名，双语）。
# 值必须保持引擎认得的形态（"custom" / "APA" / "A4" …）；显示名只在 UI 层替换，
# 保存时由 app._opt_val 反查回原值 —— 绝不把显示文字反解析成值。
# v2.3.73 新增 "reference_style.custom"：参考文献可以「按我校要求」，不再被定死成 APA。
Q_OPTION_LABELS: dict = {
    "reference_style.custom": {"zh": "按我校要求（自定义）",
                               "en": "My school's own (custom)"},
    # v2.3.85：表格线的选项值（引擎值是 three_line，界面上说人话）
    "table_lines.three_line": {"zh": "学术三线表", "en": "Three-line (academic)"},
    # v2.3.97：摘要标题对齐（引擎值是 left/center）
    "abstract_heading_alignment.left": {"zh": "左对齐", "en": "Left-aligned"},
    "abstract_heading_alignment.center": {"zh": "居中", "en": "Centered"},
}


def questionnaire_option(lang: str, field_key: str, value):
    """choice 选项的显示名（未登记的选项原样显示，保证旧数据不受影响）。"""
    m = Q_OPTION_LABELS.get("%s.%s" % (field_key, value))
    if not m:
        return value
    return m.get(lang) or m.get("en") or value


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
