# -*- coding: utf-8 -*-
"""打包后「相对 __file__」的回归测试（v2.3.123 起的硬约束）。

真实事故（2.3.122 线上）
------------------------
macOS / Linux 发行版里**页码修正完全失效且不报错**：fix 返回成功、报告说页码
已改，实际文档里连页眉都没创建；Windows 是好的，同一份源码在 macOS 上跑也是好的。

根因：python-docx 用 ``os.path.split(__file__)[0]`` 拼自带模板的路径。
Nuitka 打包后在 macOS/Linux 上 ``__file__`` 是**相对路径**（``docx/parts/hdrftr.py``），
相对路径按进程当前目录解析 —— 客户从任何目录启动都读不到
``docx/templates/default-header.xml``，``open()`` 抛 FileNotFoundError，
而 ``fixer._ensure_page_number_in_part`` 用 ``except Exception`` 兜底，异常被吞掉。

``src/docx_boot.py`` 在 ``src`` 包初始化时把这些相对 ``__file__`` 修成绝对路径。
本文件锁住这条行为：**只要有人把这个修复去掉或改坏，这里立刻红。**

为什么不用「编译一份产物再验」来测：那要几分钟且依赖 CI；把 ``__file__``
改成相对是等价的触发条件，本地毫秒级就能复现同一故障。
"""

from __future__ import annotations

import os

import pytest


_DOCX_MODULES = (
    "docx",
    "docx.parts.hdrftr",
    "docx.parts.comments",
    "docx.parts.document",
    "docx.oxml.settings",
)


@pytest.fixture
def relative_docx_files():
    """把 docx 各模块的 __file__ 改成相对路径（模拟 macOS/Linux 打包后的形态）。

    退出时无条件还原，并重置 docx_boot 的幂等标志 —— 否则会污染同机其它用例。
    """
    import importlib

    import src.docx_boot as boot

    saved: dict[str, str] = {}
    site_dir = ""
    for name in _DOCX_MODULES:
        mod = importlib.import_module(name)
        cur = getattr(mod, "__file__", None)
        if not cur:
            continue
        saved[name] = cur
        cur_posix = cur.replace("\\", "/")
        idx = cur_posix.find("docx/")
        if idx >= 0:
            mod.__file__ = cur_posix[idx:]
        if name == "docx":
            # site-packages 目录：docx 包的上一层
            site_dir = os.path.dirname(os.path.dirname(os.path.abspath(cur)))

    old_applied = boot._applied
    boot._applied = False
    try:
        yield site_dir
    finally:
        for name, cur in saved.items():
            importlib.import_module(name).__file__ = cur
        boot._applied = old_applied


def test_relative_file_breaks_and_boot_fixes_it(relative_docx_files):
    """相对 __file__ 必须能复现故障，修完之后必须恢复。"""
    from docx.parts.hdrftr import HeaderPart

    site_dir = relative_docx_files
    assert site_dir, "拿不到 site-packages 目录，夹具失效"

    # ① 复现：此时读不到模板（这正是线上 macOS 的症状）
    with pytest.raises(Exception):
        HeaderPart._default_header_xml()

    # ② 修复
    import src.docx_boot as boot

    fixed = boot.apply(base_dir=site_dir)
    assert fixed >= 1, "docx_boot.apply() 一个模块都没修，等于没生效"

    # ③ 恢复：模板读得到，且内容是合法的页眉 XML
    raw = HeaderPart._default_header_xml()
    assert raw and b"hdr" in raw, "读到的不是页眉模板"


def test_page_number_can_be_written_under_relative_files(relative_docx_files, tmp_path):
    """真正的验收：相对 __file__ 下，页码要能真的写进文档。

    这一条直接对应客户看到的现象（「修完还是红的」），比只看模板能否读到更硬。
    """
    import docx

    from src.engine import fixer
    from src.engine.questionnaire import build_target
    import src.docx_boot as boot

    site_dir = relative_docx_files
    assert boot.apply(base_dir=site_dir) >= 1

    fixture = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "tools", "crossplat", "fixtures", "bad_apa.docx")
    if not os.path.isfile(fixture):
        pytest.skip("缺跨平台测试文档 tools/crossplat/fixtures/bad_apa.docx")

    d = docx.Document(fixture)
    target = build_target("APA-Student")
    info = fixer._apply_page_number(d, "top-right")

    assert info, "页码步骤什么都没做（就是线上那个静默失败）"
    assert info.get("added") == 1, "页码应当被新增，实际 %r" % (info,)
    assert info.get("location") == "header"
    assert info.get("alignment") == "right"

    out = os.path.join(str(tmp_path), "fixed.docx")
    d.save(out)
    assert os.path.isfile(out)

    import zipfile

    with zipfile.ZipFile(out) as z:
        names = z.namelist()
    assert "word/header1.xml" in names, "修正稿里没有页眉部件，页码实际上没落地"
