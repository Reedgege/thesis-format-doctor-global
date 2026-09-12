"""启动入口 & 图标路径回归测试。

背景：v2.0.1 的 Windows 包双击「毫无反应」——exe 以**窗口子系统**编译（没有控制台），
而无参数启动时原来只 ``print_help()`` 就退出，帮助文本没有任何地方能显示，
用户看到的就是"程序打不开"。本文件把「双击必须进 GUI」钉成回归测试。

同时覆盖：显式子命令语义不变、致命错误必须被"看见"、窗口图标路径解析。
"""

from __future__ import annotations

import importlib.util
import os
import sys
import types

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from src import main as srcmain          # noqa: E402
from src.ui import iconpath              # noqa: E402


def _load_root_launcher():
    """把仓库根的 main.py 当普通模块加载（它是打包入口，不在 src 包内）。"""
    path = os.path.join(ROOT, "main.py")
    spec = importlib.util.spec_from_file_location("_tfdg_root_launcher", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


# --------------------------------------------------------------- 根启动器
def test_root_launcher_no_args_injects_gui(monkeypatch):
    """双击/快捷方式（无参数）必须自动补上 gui 子命令。"""
    launcher = _load_root_launcher()
    seen = {}

    def fake_cli_main(argv):
        seen["argv"] = argv
        return 0

    monkeypatch.setattr(srcmain, "main", fake_cli_main)
    monkeypatch.setattr(sys, "argv", ["thesis-format-doctor-global.exe"])

    assert launcher.main([]) == 0
    assert seen["argv"] == ["gui"], "无参数启动没有进入 GUI —— 这就是「双击打不开」的根因"


def test_root_launcher_no_args_also_fixes_sys_argv(monkeypatch):
    """sys.argv 也要被补全（子命令解析可能回读 sys.argv）。"""
    launcher = _load_root_launcher()
    monkeypatch.setattr(srcmain, "main", lambda argv: 0)
    monkeypatch.setattr(sys, "argv", ["app.exe"])
    launcher.main([])
    assert sys.argv[1:] == ["gui"]


def test_root_launcher_passes_explicit_args_through(monkeypatch):
    """显式子命令（check/fix/activate/status/export-schema）原样透传。"""
    launcher = _load_root_launcher()
    seen = {}

    def fake_cli_main(argv):
        seen["argv"] = argv
        return 0

    monkeypatch.setattr(srcmain, "main", fake_cli_main)
    monkeypatch.setattr(sys, "argv", ["app.exe"])

    assert launcher.main(["status"]) == 0
    assert seen["argv"] == ["status"]


def test_root_launcher_reports_fatal_load_error(monkeypatch):
    """模块加载失败要弹给用户看，而不是静默 return。"""
    launcher = _load_root_launcher()
    msgs = []
    monkeypatch.setattr(launcher, "_fatal", lambda msg: msgs.append(msg))
    # sys.modules 里放 None -> `from src.main import main` 抛 ImportError
    monkeypatch.setitem(sys.modules, "src.main", None)

    assert launcher.main(["check", "x.docx"]) == 1
    assert msgs and "Failed to load" in msgs[0]


def test_fatal_never_raises(monkeypatch):
    """弹窗通道本身出问题时也不能把异常抛出去（否则又变成静默崩溃）。"""
    launcher = _load_root_launcher()
    import ctypes

    fake_user32 = types.SimpleNamespace(MessageBoxW=lambda *a, **k: 0)
    monkeypatch.setattr(ctypes, "windll", types.SimpleNamespace(user32=fake_user32),
                        raising=False)
    monkeypatch.setattr(sys, "stderr", None)
    launcher._fatal("boom")   # 不应抛异常


# --------------------------------------------------------------- src.main
def test_no_subcommand_with_console_prints_help(capsys):
    """源码/终端里跑 `python -m src.main` 仍应显示帮助（不弹 GUI）。"""
    rc = srcmain.main([])
    out = capsys.readouterr().out
    assert rc == 0
    assert "check" in out and "fix" in out


def test_no_subcommand_without_console_goes_to_gui(monkeypatch, capsys):
    """窗口子系统（stdout 为 None）下无子命令 -> 必须进 GUI。"""
    monkeypatch.setattr(srcmain, "_has_console", lambda: False)
    hits = []
    monkeypatch.setattr(srcmain, "cmd_gui", lambda args: hits.append(args) or 0)

    assert srcmain.main([]) == 0
    assert hits, "无控制台时没有调用 cmd_gui —— 双击会静默退出"


def test_explicit_subcommand_still_dispatches(monkeypatch):
    """有子命令时不得被 GUI 兜底逻辑截胡。"""
    monkeypatch.setattr(srcmain, "_has_console", lambda: False)
    hits = []
    monkeypatch.setattr(srcmain, "cmd_status", lambda args: hits.append("status") or 0)

    assert srcmain.main(["status"]) == 0
    assert hits == ["status"]


def test_cmd_gui_reports_when_tk_missing(monkeypatch):
    """Tk 不可用时 cmd_gui 必须返回 1 并弹窗，绝不能静默 return。"""
    msgs = []
    monkeypatch.setattr(srcmain, "_fatal_dialog", lambda m: msgs.append(m))
    monkeypatch.setitem(sys.modules, "src.ui.app", None)   # 令 import 失败

    assert srcmain.cmd_gui(None) == 1
    assert msgs and "Tk" in msgs[0]


def test_cmd_gui_reports_when_ui_crashes(monkeypatch):
    """GUI 起得来但跑崩了，也要让用户看得见。"""
    msgs = []
    monkeypatch.setattr(srcmain, "_fatal_dialog", lambda m: msgs.append(m))

    fake = types.ModuleType("src.ui.app")

    def _boom():
        raise RuntimeError("display gone")

    fake.main = _boom
    monkeypatch.setitem(sys.modules, "src.ui.app", fake)

    assert srcmain.cmd_gui(None) == 1
    assert msgs and "display gone" in msgs[0]


# --------------------------------------------------------------- 图标路径
def test_icon_found_in_source_layout():
    """源码运行时：src/data/icon.png 必须被找到（`--include-package-data=src` 同理）。"""
    assert iconpath.find_icon() is not None


def test_icon_candidates_cover_frozen_layout(monkeypatch, tmp_path):
    """打包后：包内 data/ 不可用时，必须能 fallback 到 <exe 目录>/assets/icon.png。"""
    exe_dir = tmp_path / "app"
    (exe_dir / "assets").mkdir(parents=True)
    ico = exe_dir / "assets" / "icon.png"
    ico.write_bytes(b"\x89PNG\r\n\x1a\n")   # 内容不解析，这里只验证路径解析

    # 模拟冻结布局：__file__ 指向一个不存在的包路径（包内 data/ 探不到）
    monkeypatch.setattr(iconpath, "__file__", str(tmp_path / "pkg" / "ui" / "iconpath.py"))
    monkeypatch.setattr(sys, "argv", [str(exe_dir / "thesis-format-doctor-global.exe")])
    monkeypatch.setattr(sys, "executable", str(exe_dir / "thesis-format-doctor-global.exe"))
    monkeypatch.chdir(tmp_path)

    cands = [os.path.normpath(c) for c in iconpath.icon_candidates()]
    assert str(ico) in cands
    assert iconpath.find_icon() == str(ico)


def test_icon_returns_none_when_absent(monkeypatch, tmp_path):
    """一个都找不到 -> None（调用方静默跳过，绝不能因为图标影响启动）。"""
    monkeypatch.setattr(sys, "argv", [str(tmp_path / "nope" / "app.exe")])
    monkeypatch.setattr(sys, "executable", str(tmp_path / "nope" / "app.exe"))
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(iconpath, "__file__", str(tmp_path / "pkg" / "iconpath.py"))

    assert iconpath.find_icon() is None


def test_icon_candidates_are_deduped_and_absolute():
    cands = iconpath.icon_candidates()
    assert cands, "候选列表不应为空"
    assert all(os.path.isabs(c) for c in cands)
    assert len(cands) == len(set(cands)), "候选路径不应重复"


@pytest.fixture(scope="session")
def _tk_session():
    """整轮测试只创建一个 Tk 解释器。

    本机（系统 Python 3.13）里，同一进程**第二次** ``tk.Tk()`` 会因为 Tcl 库
    ``ttk/scale.tcl`` 加载失败而抛 TclError —— 属环境问题（首次创建正常），
    不是产品缺陷；改用共享 root 就绕开了，也让 GUI 回归能在 pytest 里跑起来。
    """
    if importlib.util.find_spec("tkinter") is None:
        pytest.skip("本机 Python 无 tkinter（用系统 Python 跑 GUI 冒烟）")
    import tkinter as tk

    try:
        root = tk.Tk()
    except Exception as e:                     # noqa: BLE001
        pytest.skip("Tk 不可用：%s" % e)
    root.withdraw()                            # 不闪窗
    yield root
    try:
        root.destroy()
    except Exception:
        pass


@pytest.fixture
def tk_root(_tk_session):
    """每个用例一个干净的 root（复用同一 Tk 实例，先清掉上个用例的控件）。"""
    for w in list(_tk_session.winfo_children()):
        try:
            w.destroy()
        except Exception:
            pass
    return _tk_session


def test_window_icon_applies_without_error(tk_root):
    """真机 GUI 冒烟：起一个 Tk 窗口、套图标，构造过程**绝不能抛异常**。

    图标本身是「尽力而为」：个别 runner（Tcl 没编进 PNG 解码器、或没有桌面会话）
    挂不上图标，属**环境限制**而非产品缺陷 —— 所以这里不断言图标一定生效，免得
    偶发红灯掩盖真实回归。「引用必须保留（防 GC）」这个真回归点由下面那条
    **确定性**单测钉住，不依赖本机 Tcl 的解码能力。
    """
    from src.ui import app as ui_app

    ui = ui_app.App(tk_root)
    assert ui is not None


def test_window_icon_reference_is_retained(tk_root, monkeypatch):
    """确定性回归：图标的 PhotoImage 必须保留 Python 引用。

    ``iconphoto(True, img)`` 里只传临时对象时，Image 一旦被 GC 回收，部分 Tk 版本
    会把窗口图标还原成默认 —— v2.0.2 踩过的坑。这里把 PhotoImage 换成哨兵对象，
    断言引用真的挂在实例属性上、且 ``iconphoto`` 被调用过。**不依赖本机 Tcl 能否
    解码 PNG**，所以在任何 runner 上结论都一致。
    """
    import tkinter as tk

    from src.ui import app as ui_app
    from src.ui import iconpath as ui_iconpath

    sentinel = object()
    monkeypatch.setattr(ui_iconpath, "find_icon", lambda: os.path.abspath(__file__))
    monkeypatch.setattr(ui_app.tk, "PhotoImage", lambda *a, **kw: sentinel)
    applied = []
    monkeypatch.setattr(tk_root, "iconphoto", lambda *a, **kw: applied.append(a))

    ui = ui_app.App(tk_root)
    assert getattr(ui, "_window_icon", None) is sentinel, "图标引用未保留（会被 GC）"
    assert applied, "iconphoto 未被调用"


def test_window_icon_never_breaks_startup(tk_root, monkeypatch):
    """确定性回归：图标解析/挂载整个炸掉，也绝不能把启动带崩（P0 级保护）。

    「双击打不开」是本产品最重的 P0 历史事故；图标只是锦上添花，必须保证
    ``_apply_window_icon`` 在任何异常下静默收场。这里让图标路径解析直接抛异常。
    """
    from src.ui import app as ui_app
    from src.ui import iconpath as ui_iconpath

    def boom(*a, **kw):
        raise RuntimeError("icon path resolution exploded")

    monkeypatch.setattr(ui_iconpath, "find_icon", boom)
    monkeypatch.setattr(ui_iconpath, "find_icon_ico", boom)

    ui = ui_app.App(tk_root)                       # 构造过程绝不能抛异常
    assert getattr(ui, "_window_icon", None) is None


# --------------------------------------------------------------- 语言（GUI）
def _all_widgets(parent):
    """递归收集所有子控件（用于断言界面上到底显示了什么字）。"""
    out = []
    for child in parent.winfo_children():
        out.append(child)
        out.extend(_all_widgets(child))
    return out


def _label_texts(root) -> list:
    import tkinter as tk

    texts = []
    for w in _all_widgets(root):
        if isinstance(w, tk.Label):
            try:
                texts.append(str(w.cget("text")))
            except Exception:
                pass
    return texts


def test_gui_defaults_to_english(tmp_path, monkeypatch, tk_root):
    """海外版 GUI 必须**默认英文**（老板 2026-09-11 定：国外以英文为主）。"""
    monkeypatch.setenv("TFD_SETTINGS_FILE", str(tmp_path / "settings.json"))

    from src.ui import app as ui_app

    ui = ui_app.App(tk_root)
    assert ui.lang == "en"
    texts = _label_texts(tk_root)
    assert "Thesis Format Doctor" in texts
    assert "Files" in texts and "Workflow" in texts
    # 除语言切换按钮外，英文界面不应出现中文（那个按钮按设计显示目标语言名"中文"）
    lang_btn = ui.tr("lang_button")
    rest = [t for t in texts if t != lang_btn]
    assert not [t for t in rest if any("\u4e00" <= c <= "\u9fff" for c in t)]


def test_gui_language_switch_rebuilds_and_persists(tmp_path, monkeypatch, tk_root):
    """切到中文：界面整体变中文、偏好落盘、已选文件不丢。"""
    monkeypatch.setenv("TFD_SETTINGS_FILE", str(tmp_path / "settings.json"))

    from src.ui import app as ui_app, i18n

    ui = ui_app.App(tk_root)
    ui.docx_path.set(r"C:\tmp\my thesis.docx")   # 假装用户已选论文
    ui._refresh_rows()

    ui._set_lang("zh")
    assert ui.lang == "zh"
    assert i18n.load_lang() == "zh"              # 偏好已落盘

    texts = _label_texts(tk_root)
    assert "论 文 格 式 医 生" in texts
    assert "文件选择" in texts and "处理步骤" in texts
    # 状态保持：已选论文不能被语言切换清掉
    assert ui.docx_path.get() == r"C:\tmp\my thesis.docx"
    assert any("论文" in t for t in texts)

    # 再切回英文
    ui._set_lang("en")
    assert ui.lang == "en"
    assert "Files" in _label_texts(tk_root)
    assert i18n.load_lang() == "en"


def test_gui_shows_placeholder_then_survives_language_switch(tmp_path, monkeypatch,
                                                             tk_root):
    """报告区先显示占位说明；切换语言后要能重建，且已跑出的报告不丢。"""
    monkeypatch.setenv("TFD_SETTINGS_FILE", str(tmp_path / "settings.json"))

    from src.ui import app as ui_app

    ui = ui_app.App(tk_root)
    assert ui._report_text.strip()

    ui._set_report("SOME REPORT BODY")
    ui._set_lang("zh")
    # 用户已经跑过体检 —— 报告正文不该被语言切换清空
    assert ui._report_text == "SOME REPORT BODY"

