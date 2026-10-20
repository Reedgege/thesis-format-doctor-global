"""recents.py 纯数据层断言式测试（不需要 tkinter，可离线跑）。

把 v2.3.78「最近使用」的本地持久化逻辑钉成回归：
- 手填问卷按内容去重、按时间倒序、最多 5 条；
- AI 导入按路径去重、最多 5 条，且只认真实存在的文件；
- last_source 记录与读取；
- 启动自动恢复：manual 取最新 dict、ai 取最新且存在的路径；
- 摘要标签双语、英文模式零中文泄漏；
- 任何损坏/缺失都不抛、不阻塞（回退空结构）。
"""

from __future__ import annotations

import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from src.ui import recents                 # noqa: E402


@pytest.fixture
def tmp_recents(tmp_path, monkeypatch):
    """把 recents 文件指到临时路径（隔离，避免污染真实用户数据）。"""
    p = tmp_path / "recents.json"
    monkeypatch.setenv("TFD_RECENTS_FILE", str(p))
    return p


def test_default_is_empty(tmp_recents):
    assert recents.load_recents() == {"manual": [], "ai": [], "last_source": ""}


def test_add_manual_dedup_and_cap(tmp_recents):
    rec = recents.load_recents()
    # 加 6 条不同内容 → 应只留 5 条，且最新的在最前
    for i in range(6):
        rec = recents.add_manual(rec, {"font_size_pt": float(i)})
    assert len(rec["manual"]) == 5
    assert rec["manual"][0]["data"] == {"font_size_pt": 5.0}
    # 再次加相同内容 → 去重（仍 5 条），且被顶到最前
    rec = recents.add_manual(rec, {"font_size_pt": 5.0})
    assert len(rec["manual"]) == 5
    assert rec["manual"][0]["data"] == {"font_size_pt": 5.0}
    # 不同键序但同内容 → 视为相同（canonical 去重）
    rec = recents.add_manual(rec, {"font_size_pt": 5.0, "x": 1})
    rec = recents.add_manual(rec, {"x": 1, "font_size_pt": 5.0})
    assert len(rec["manual"]) == 5


def test_add_ai_dedup_and_cap(tmp_recents):
    rec = recents.load_recents()
    for i in range(6):
        rec = recents.add_ai(rec, "/tmp/cfg_%d.yaml" % i)
    assert len(rec["ai"]) == 5
    assert rec["ai"][0]["path"] == "/tmp/cfg_5.yaml"
    # 同路径（不同大小写混合）→ 去重
    rec = recents.add_ai(rec, "/tmp/CFG_5.yaml")
    assert len(rec["ai"]) == 5
    assert rec["ai"][0]["path"] == "/tmp/CFG_5.yaml"


def test_last_source_roundtrip(tmp_recents):
    rec = recents.set_last_source(recents.load_recents(), "src_manual")
    assert recents.save_recents(rec) is True
    assert recents.load_recents()["last_source"] == "src_manual"


def test_remove_ai(tmp_recents):
    rec = recents.add_ai(recents.load_recents(), "/tmp/a.yaml")
    rec = recents.add_ai(rec, "/tmp/b.yaml")
    rec = recents.remove_ai(rec, "/tmp/a.yaml")
    assert [x["path"] for x in rec["ai"]] == ["/tmp/b.yaml"]


def test_first_manual_returns_latest(tmp_recents):
    rec = recents.add_manual(recents.load_recents(), {"a": 1})
    rec = recents.add_manual(rec, {"a": 2})
    assert recents.first_manual(rec) == {"a": 2}
    assert recents.first_manual({"manual": []}) is None


def test_first_existing_ai_skips_missing(tmp_path, tmp_recents):
    good = tmp_path / "good.yaml"
    good.write_text("x: 1", encoding="utf-8")
    rec = recents.add_ai(recents.load_recents(), "/tmp/gone.yaml")
    rec = recents.add_ai(rec, str(good))
    # 最新的存在文件优先
    assert recents.first_existing_ai(rec) == str(good)


def test_manual_summary_bilingual():
    data = {"font_family": "Times New Roman", "font_size_pt": 12,
            "margin_top_in": 1, "columns": "2",
            "heading_formats": {"1": {}, "2": {}}}
    en = recents.manual_summary(data, "en")
    zh = recents.manual_summary(data, "zh")
    assert "Times New Roman 12pt" in en
    assert "2 columns" in en
    assert "2 levels" in en
    assert "·" in en
    # 英文模式绝不夹中文
    assert not any(ord(c) > 0x2E7F for c in en)  # 基本拉丁 + 常见标点
    assert "双栏" in zh
    assert "2级标题" in zh
    # 空 dict
    assert recents.manual_summary({}, "en") == "Empty template"
    assert recents.manual_summary({}, "zh") == "空白模板"


def test_corrupt_file_does_not_raise(tmp_path, monkeypatch):
    p = tmp_path / "recents.json"
    p.write_text("{ this is not json", encoding="utf-8")
    monkeypatch.setenv("TFD_RECENTS_FILE", str(p))
    # 读损坏文件回退空结构，不抛
    assert recents.load_recents() == {"manual": [], "ai": [], "last_source": ""}


def test_save_and_load_roundtrip(tmp_path, monkeypatch):
    p = tmp_path / "recents.json"
    monkeypatch.setenv("TFD_RECENTS_FILE", str(p))
    rec = recents.add_manual(recents.load_recents(), {"k": "v"})
    rec = recents.add_ai(rec, "/tmp/c.yaml")
    rec = recents.set_last_source(rec, "src_ai")
    assert recents.save_recents(rec) is True
    back = recents.load_recents()
    assert back["manual"][0]["data"] == {"k": "v"}
    assert back["ai"][0]["path"] == "/tmp/c.yaml"
    assert back["last_source"] == "src_ai"
    # 原子写：临时文件不应残留
    assert not p.with_name(p.name + ".tmp").exists()
