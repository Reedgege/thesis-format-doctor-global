"""界面语言包 / 字体的断言式测试（纯数据层，不需要 tkinter，可离线跑）。

把三件"老板明确要求过"的事钉成回归：
1. 海外版**默认英文**，语言偏好可持久化、可切换，非法值回退英文；
2. 中英词条必须**一一对应**（少一条就会在界面上露出 key 或英文串）；
3. **品牌口径铁律**：界面文案里不得出现微信 / 公众号 / 小程序（海外用户不用），
   只允许官网 reedskill.com 与邮箱 hi@reedskill.com。
"""

from __future__ import annotations

import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.ui import i18n          # noqa: E402


# ------------------------------------------------------------------ 语言偏好
@pytest.fixture
def tmp_settings(tmp_path, monkeypatch):
    """把设置文件指到临时路径（隔离，避免污染真实用户偏好）。"""
    p = tmp_path / "settings.json"
    monkeypatch.setenv("TFD_SETTINGS_FILE", str(p))
    return p


def test_default_language_is_english(tmp_path, monkeypatch):
    """没有设置文件时，海外版必须默认英文。"""
    monkeypatch.setenv("TFD_SETTINGS_FILE", str(tmp_path / "nope.json"))
    assert i18n.DEFAULT_LANG == "en"
    assert i18n.load_lang() == "en"


def test_language_roundtrip(tmp_settings):
    assert i18n.save_lang("zh") is True
    assert i18n.load_lang() == "zh"
    assert i18n.save_lang("en") is True
    assert i18n.load_lang() == "en"


def test_invalid_language_falls_back_to_english(tmp_settings):
    tmp_settings.write_text('{"lang": "fr"}', encoding="utf-8")
    assert i18n.load_lang() == "en"
    assert i18n.save_lang("fr") is False


def test_corrupt_settings_do_not_break_startup(tmp_settings):
    """设置文件损坏绝不能阻塞启动（回退默认语言即可）。"""
    tmp_settings.write_text("{ this is not json", encoding="utf-8")
    assert i18n.load_lang() == "en"


def test_other_lang_toggles():
    assert i18n.other_lang("en") == "zh"
    assert i18n.other_lang("zh") == "en"


# ------------------------------------------------------------------ 词条完整
def test_string_keys_match_across_languages():
    """中英词条 key 必须完全一致，否则某种语言下会露出 key 本身。"""
    en = set(i18n.STRINGS["en"])
    zh = set(i18n.STRINGS["zh"])
    assert en == zh, "词条缺失：en 独有 %s / zh 独有 %s" % (sorted(en - zh), sorted(zh - en))


def test_no_empty_strings():
    for lang in i18n.LANGS:
        for key, val in i18n.STRINGS[lang].items():
            if isinstance(val, str):
                assert val.strip(), "%s.%s 是空串" % (lang, key)


def test_t_returns_text_and_formats():
    assert i18n.t("en", "left_title") == "Your Documents"
    assert i18n.t("zh", "left_title") == "你的文档"
    assert i18n.t("en", "about_version", "2.0.2") == "Version 2.0.2"
    assert i18n.t("zh", "about_version", "2.0.2") == "版本 v2.0.2"


def test_t_falls_back_gracefully():
    """未知 key / 未知语言 / 格式化参数不匹配，都不能让界面崩。"""
    assert i18n.t("en", "no_such_key") == "no_such_key"
    assert i18n.t("fr", "left_title") == "Your Documents"   # 非法语言 -> 英文
    assert i18n.t("en", "left_title", "extra") == "Your Documents"  # 无占位符时忽略多余参数


def test_help_lists_have_pairs():
    for lang in i18n.LANGS:
        for key in ("help_steps", "help_faqs", "how_steps"):
            items = i18n.t(lang, key)
            assert isinstance(items, list) and items, "%s.%s 应为非空列表" % (lang, key)
            for entry in items:
                assert isinstance(entry, tuple) and len(entry) == 2
                assert entry[0].strip() and entry[1].strip()


# ------------------------------------------------------------------ 品牌口径
_BANNED = ("微信", "公众号", "小程序", "WeChat", "wechat", "reedgege",
           "二维码", "扫一扫", "QR code", "QRCode")


def _walk_strings(node):
    if isinstance(node, str):
        yield node
    elif isinstance(node, dict):
        for v in node.values():
            yield from _walk_strings(v)
    elif isinstance(node, (list, tuple)):
        for v in node:
            yield from _walk_strings(v)


def test_no_wechat_or_miniapp_anywhere():
    """海外版界面文案不得出现微信 / 公众号 / 小程序（老板 2026-09-11 明确要求）。"""
    offenders = []
    for lang in i18n.LANGS:
        for s in _walk_strings(i18n.STRINGS[lang]):
            for bad in _BANNED:
                if bad in s:
                    offenders.append("%s: %s" % (lang, s))
    for s in _walk_strings(i18n.Q_LABELS_EN):
        for bad in _BANNED:
            if bad in s:
                offenders.append("Q_LABELS_EN: %s" % s)
    assert not offenders, "出现微信/小程序字样：%s" % offenders


def test_brand_contact_is_site_and_email_only():
    assert i18n.BRAND_EMAIL == "hi@reedskill.com"
    assert i18n.BRAND_SITE == "reedskill.com"
    # v2.2.0：页脚按设计规格只留三项（邮箱 / 官网 / 标语），原来的 footer_line1/2 已移除
    for lang in i18n.LANGS:
        footer = (i18n.t(lang, "footer_email") + i18n.t(lang, "footer_site")
                  + i18n.t(lang, "footer_tagline"))
        assert "hi@reedskill.com" in footer
        assert "reedskill.com" in footer


# ------------------------------------------------------------------ 问卷显示层
def test_questionnaire_labels_are_localized():
    from src.engine.questionnaire import QUESTIONNAIRE_SCHEMA

    missing = [f["key"] for f in QUESTIONNAIRE_SCHEMA if f["key"] not in i18n.Q_LABELS_EN]
    assert not missing, "以下问卷字段缺英文标签：%s" % missing
    for field in QUESTIONNAIRE_SCHEMA:
        en = i18n.questionnaire_label("en", field)
        zh = i18n.questionnaire_label("zh", field)
        assert en and zh
        assert en != zh, "字段 %s 的中英标签相同（多半是漏翻）" % field["key"]
        # 英文模式下不能残留中文字符
        assert not any("\u4e00" <= ch <= "\u9fff" for ch in en), \
            "字段 %s 的英文标签里混入中文：%s" % (field["key"], en)


def test_questionnaire_help_has_no_chinese_in_english_mode():
    from src.engine.questionnaire import QUESTIONNAIRE_SCHEMA

    for field in QUESTIONNAIRE_SCHEMA:
        help_txt = i18n.questionnaire_help("en", field)
        assert not any("\u4e00" <= ch <= "\u9fff" for ch in help_txt), \
            "字段 %s 的英文帮助里混入中文：%s" % (field["key"], help_txt)


# ------------------------------------------------------------------ 字体
def test_font_names_are_complete():
    for lang in i18n.LANGS:
        spec = i18n.font_spec(lang, available=None)
        assert set(spec) == set(i18n.FONT_NAMES)
        for name, val in spec.items():
            assert len(val) == 3 and isinstance(val[1], int)
            assert val[2] in ("normal", "bold")


def test_english_prefers_serif_for_title_and_sans_for_body():
    """英文界面：标题走衬线（Georgia），正文走无衬线（Segoe UI），别套中文字体。"""
    spec = i18n.font_spec("en", available=None)
    assert spec["F_TITLE"][2] == "bold"
    assert i18n.resolve_family("Georgia") == "Georgia"
    assert i18n.resolve_family("Segoe UI") == "Segoe UI"
    # 英文模式下不该出现中文字体族
    families = {v[0] for v in spec.values()}
    assert not {"KaiTi", "Microsoft YaHei"} & families


def test_chinese_keeps_kaiti_and_yahei():
    spec = i18n.font_spec("zh", available=None)
    assert spec["F_TITLE"][0] == "KaiTi"
    assert spec["F_BODY"][0] == "Microsoft YaHei"


def test_platform_fallbacks(monkeypatch):
    """macOS / Linux 没有 Georgia / Segoe UI / 楷体 / 雅黑时映射到系统对等字体。"""
    monkeypatch.setattr(i18n.platform, "system", lambda: "Linux")
    assert i18n.resolve_family("Georgia") == "Noto Serif"
    assert i18n.resolve_family("Segoe UI") == "Noto Sans"
    assert i18n.resolve_family("KaiTi") == "Noto Serif CJK SC"

    monkeypatch.setattr(i18n.platform, "system", lambda: "Darwin")
    assert i18n.resolve_family("Segoe UI") == "Helvetica Neue"
    assert i18n.resolve_family("Microsoft YaHei") == "PingFang SC"

    monkeypatch.setattr(i18n.platform, "system", lambda: "Windows")
    assert i18n.resolve_family("Georgia") == "Georgia"
    assert i18n.resolve_family("Microsoft YaHei") == "Microsoft YaHei"


def test_resolve_family_prefers_installed_one(monkeypatch):
    """系统里真有首选字体就用它；没有则用平台兜底值。"""
    monkeypatch.setattr(i18n.platform, "system", lambda: "Linux")
    assert i18n.resolve_family("Segoe UI", available={"Segoe UI", "Arial"}) == "Segoe UI"
    assert i18n.resolve_family("Segoe UI", available={"Arial"}) == "Noto Sans"


def test_unknown_family_passes_through():
    assert i18n.resolve_family("Some Font") == "Some Font"
