"""全仓测试共享夹具。

v2.3.78 起软件会把「最近使用」（手填问卷模板 / AI 导入文件）落盘到用户目录
``~/.paperformatpro/recents.json``，并在启动时自动恢复上一次来源。任何构建
App 的测试若没隔离该路径，会出两个真实事故（2026-10-20 已踩）：

1. 测试数据被写进**真实用户配置目录**（test_ui_redesign 的问卷保存用例把
   真实 recents.json 写成了 src_manual + SimHei/KaiTi 假模板）；
2. 同机后续所有测试建 App 时被「启动自动恢复」读回脏数据 ——
   ``test_collapse_toggles_visibility`` 断言默认内置规范下 ``_adv_enabled is False``
   直接翻红（跟当时正在改的图标毫无关系，纯背锅）。

故用 autouse 夹具给**每个用例**一个独立的 recents 路径。个别用例如需自定义，
在自己的夹具里再 ``monkeypatch.setenv`` 覆盖即可（后 set 者优先）。
"""

from __future__ import annotations

import pytest


@pytest.fixture
def legacy_text_transform(monkeypatch):
    """临时放开「改写文字」能力（v2.3.84 产品默认已按老板清单第 5 条锁死）。

    这批用例覆盖的是**底层实现**（章节编号/全大写/引文重写）是否还正确、是否幂等；
    产品默认行为由 tests/test_text_lock.py 钉死。检查器与修正器两个开关必须一起翻，
    否则会造出"报了改不了"的假环境。
    """
    from src.engine import checker, fixer
    monkeypatch.setattr(checker, "TEXT_IS_LOCKED", False)
    monkeypatch.setattr(fixer, "TEXT_IS_LOCKED", False)
    yield


@pytest.fixture(autouse=True)
def _isolated_recents(tmp_path, monkeypatch):
    """每个用例独立 recents.json，绝不读写真实用户配置。"""
    monkeypatch.setenv("TFD_RECENTS_FILE", str(tmp_path / "recents.json"))
    yield
