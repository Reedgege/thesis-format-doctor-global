"""复核弹窗动作栏吸底回归测试（tkinter，真起窗口）。

背景（老板 2026-10-22 截图圈出）：`Confirm your format settings` 弹窗里
Back / Run check 两个按钮排在可滚动内容的**末尾**，客户必须一路滚到底才看得见。
老板原话：「把下边这个地方固定住，不要滚动到最下面才看到」。

病根不在按钮样式，而在 `ModalShell` 的 sticky 模式没打开：这个弹窗是四个弹窗里
内容最高的（4 个分区 + 页面网格 + 栏数/纸型/页码三行 + 标题表 + 风格规则），
内容必然溢出 720 的可视区，按钮栏于是被挤到折线以下。

本文件钉住两件事：
1. 按钮栏是**窗口底部**的真控件（父链在 `_button_bar` 下，不在 `_inner` 里）；
2. 🔴 把画布滚到底之后，按钮的屏幕坐标**一个像素都不动** —— 这才是客户视角的判据；
3. 吸底模式下 `_fit_modal` 必须把按钮栏高度算进窗口高度，否则画布会比内容矮
   一条按钮栏、明明装得下也永远挂着滚动条（顺手一起钉住）。

无 tkinter / 无法建窗的环境自动 skip（见 ``_tk_session``）。
"""
from __future__ import annotations

import importlib.util
import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)


@pytest.fixture(scope="session")
def _tk_session():
    """整个会话共用一个 Tk root（与 test_clear_buttons / test_ui_redesign 同款）。"""
    if importlib.util.find_spec("tkinter") is None:
        pytest.skip("本机 Python 无 tkinter（GUI 回归需带 tkinter 的解释器）")
    import tkinter as tk

    try:
        root = tk.Tk()
    except Exception as e:                     # noqa: BLE001
        pytest.skip("Tk 不可用：%s" % e)
    yield root
    try:
        root.destroy()
    except Exception:
        pass


@pytest.fixture
def ui(tmp_path, monkeypatch, _tk_session):
    monkeypatch.setenv("TFD_SETTINGS_FILE", str(tmp_path / "settings.json"))
    for w in list(_tk_session.winfo_children()):
        try:
            w.destroy()
        except Exception:
            pass
    from src.ui import app as ui_app
    # 掐掉启动时的授权心跳线程：它在**非主线程**里调 ``root.after()``（Tk 不是线程安全的），
    # pytest 同一进程里连续建十几个 App，残留线程会在主循环之外注册命令，抛
    # ``RuntimeError: main thread is not in main loop``，**偶发**把后续用例的
    # setup 打成 ERROR（2026-10-22 实测：本文件单跑 11 passed，全量连跑出现 5 errors，
    # 再跑一遍又全绿）。修 CI 的确定性，不改产品行为。
    monkeypatch.setattr(ui_app.App, "_refresh_license_at_launch",
                        lambda self: None, raising=False)
    a = ui_app.App(_tk_session)
    _tk_session.update_idletasks()
    yield a
    try:
        a._close_popups()
    except Exception:
        pass


def _mapped(root) -> bool:
    """窗口是否真的映射到屏幕（无桌面会话时为 False）。

    ⚠️ 这是本仓库的既有约定（见 test_ui_redesign._mapped 的注释）：CI 的
    windows-latest runner 没有交互式桌面，窗口建得出但**量不到确定像素** ——
    ``winfo_rooty/winfo_height`` 会给出垃圾值。断言实际像素的用例必须先用本函数
    守卫，否则本地全绿、CI 全红（v2.3.88 就是这么把发版卡住的）。
    """
    try:
        return bool(root.winfo_ismapped())
    except Exception:
        return False


def _require_desktop(ui):
    """无桌面会话 → 跳过；跳过原因写清楚，免得以后有人以为是「测试没用」。"""
    if not _mapped(ui.root):
        pytest.skip("无桌面会话（headless runner），量不到真实像素，本用例跳过")


def _settle(root, win, rounds: int = 4):
    """让 Tk 真正把几何算完再断言。

    只 ``update_idletasks()`` 是不够的：``_fit_modal`` 改完 geometry 后，窗口管理器
    可能再钳一次尺寸，pack 的重排要等一轮完整的 ``update()`` 才落地。不泵事件就断言
    会读到「画布按旧高度排好、按钮栏还在窗外」的中间态（我第一版测试就是这么假报的）。
    """
    for _ in range(rounds):
        win.update_idletasks()
        root.update()


def _open_confirm(ui):
    """打开「确认格式设置」弹窗（学校模板/问卷/AI 三来源共用）。"""
    from src.engine.questionnaire import target_from_spec

    target = target_from_spec("APA-Student")
    ui._confirm_dialog(target)
    # 弹窗是 _track 过的最后一个 Toplevel
    wins = [w for w in ui.root.winfo_children()
            if w.winfo_class() == "Toplevel"]
    assert wins, "没有找到确认弹窗"
    win = wins[-1]
    _settle(ui.root, win)
    return win


def _open_upgrade(ui):
    """打开「升级 / 购买 / 激活」弹窗（含默认折叠的离线激活区）。"""
    ui._upgrade_dialog()
    wins = [w for w in ui.root.winfo_children()
            if w.winfo_class() == "Toplevel"]
    assert wins, "没有找到升级/激活弹窗"
    win = wins[-1]
    _settle(ui.root, win)
    return win


def _find_offline_head(win):
    """找「离线激活」折叠区的标题行（cursor=hand2 的 Frame，绑着 <Button-1>）。"""
    def walk(w):
        for ch in w.winfo_children():
            try:
                if ch.winfo_class() == "Frame" and str(ch.cget("cursor")) == "hand2":
                    return ch
            except Exception:
                pass
            got = walk(ch)
            if got is not None:
                return got
        return None
    return walk(win)


def test_upgrade_dialog_uses_sticky_action_bar(ui):
    win = _open_upgrade(ui)
    assert getattr(win, "_sticky", False) is True


def test_upgrade_close_button_lives_in_window_bottom_bar(ui):
    win = _open_upgrade(ui)
    btn = win._primary_btn or (win._secondary_btns[0] if win._secondary_btns else None)
    assert btn is not None, "升级弹窗没有找到动作按钮"
    chain, w = [], btn
    while w is not None:
        chain.append(str(w))
        parent_name = w.winfo_parent()
        if not parent_name:
            break
        w = w.nametowidget(parent_name)
    assert str(win._button_bar) in chain, chain
    assert str(win._inner) not in chain, "关闭按钮还在滚动内容里：%s" % chain


def test_upgrade_close_button_stays_glued_to_window_bottom_on_expand(ui):
    """🔴 展开「离线激活」把内容撑高后，Close 必须仍贴在**窗口底边**、且完整可见。

    注意口径：这里不是「屏幕坐标一动不动」—— 窗口会跟着长高，钉在窗底的按钮
    自然会一起下移，那是正确行为。要钉的是「按钮底边 ≈ 窗口底边」这个关系，
    以及「窗口长高后按钮还在窗内」（旧版展开后按钮被推到折线以下、要滚才见）。
    """
    _require_desktop(ui)
    win = _open_upgrade(ui)
    btn = win._primary_btn or (win._secondary_btns[0] if win._secondary_btns else None)
    assert btn is not None

    def gap_to_bottom():
        return (win.winfo_rooty() + win.winfo_height()) - (btn.winfo_rooty() + btn.winfo_height())

    _settle(ui.root, win)
    gap_before = gap_to_bottom()
    inner_before = int(win._inner.winfo_reqheight())

    head = _find_offline_head(win)
    assert head is not None, "没找到离线激活折叠区的标题行"
    head.event_generate("<Button-1>", x=4, y=4)
    _settle(ui.root, win)

    inner_after = int(win._inner.winfo_reqheight())
    assert inner_after > inner_before, \
        "折叠区没有真的展开（内容高 %s → %s）" % (inner_before, inner_after)

    # ① 仍然贴底（允许窗口边框/底留白带来的小差）
    assert abs(gap_to_bottom() - gap_before) <= 6, \
        "展开后按钮不再贴窗口底边：%s → %s" % (gap_before, gap_to_bottom())
    # ② 完整落在窗内
    assert btn.winfo_rooty() + btn.winfo_height() <= win.winfo_rooty() + win.winfo_height() + 2


def test_upgrade_button_visible_after_expand_and_scroll(ui):
    """展开 + 滚到最底，按钮仍完整落在窗口内。"""
    _require_desktop(ui)
    win = _open_upgrade(ui)
    btn = win._primary_btn or (win._secondary_btns[0] if win._secondary_btns else None)
    head = _find_offline_head(win)
    if head is not None:
        head.event_generate("<Button-1>", x=4, y=4)
        _settle(ui.root, win)
    win._cv.yview_moveto(1.0)
    _settle(ui.root, win)

    win_h = win.winfo_height()
    bar_h = int(win._button_bar.winfo_reqheight())
    assert abs((win._cv.winfo_height() + bar_h) - win_h) <= 4, \
        (win._cv.winfo_height(), bar_h, win_h)
    btn_bottom = btn.winfo_rooty() + btn.winfo_height()
    win_bottom = win.winfo_rooty() + win_h
    assert btn_bottom <= win_bottom + 2, (btn_bottom, win_bottom)


def test_upgrade_window_grows_when_offline_fold_expands(ui):
    """点开「离线激活」时窗口必须跟着长高 —— 新内容不能落在折线以下。

    病灶（v2.3.88 及以前，本次连带修掉）：`_toggle` 在 ``WM_SETREDRAW(0)`` 冻结
    期间调 ``_fit_modal``，而**冻结期间的窗口尺寸变更会被整条丢掉**（实测：冻结中
    请求 560，``wm geometry`` 仍报 400，解冻后也不补）。结果就是「点开离线激活 →
    内容长高了、窗口原地不动」，机器码/离线码全落在折线以下，客户以为点了没反应。
    修法：先解冻（不重绘）→ 定高 → 补一次重绘。
    """
    _require_desktop(ui)
    win = _open_upgrade(ui)
    h0 = int(win.winfo_height())
    head = _find_offline_head(win)
    assert head is not None, "没找到离线激活折叠区的标题行"

    head.event_generate("<Button-1>", x=4, y=4)
    _settle(ui.root, win)
    h1 = int(win.winfo_height())
    assert h1 > h0 + 20, "展开折叠区后窗口没有长高：%s → %s" % (h0, h1)

    head.event_generate("<Button-1>", x=4, y=4)     # 再点一次收起
    _settle(ui.root, win)
    assert abs(int(win.winfo_height()) - h0) <= 4, \
        "收起折叠区后窗口没有回到原高度：%s → %s" % (h0, win.winfo_height())


def test_confirm_dialog_uses_sticky_action_bar(ui):
    win = _open_confirm(ui)
    assert getattr(win, "_sticky", False) is True


def test_primary_button_lives_in_window_bottom_bar(ui):
    """按钮的父容器必须挂在 _button_bar 下，而不是可滚动内容 _inner 里。"""
    win = _open_confirm(ui)
    btn = win._primary_btn
    assert btn is not None
    # 父链一路往上，必须经过 _button_bar，且**不能**经过 _inner
    chain, w = [], btn
    while w is not None:
        chain.append(str(w))
        parent_name = w.winfo_parent()
        if not parent_name:
            break
        w = w.nametowidget(parent_name)
    assert str(win._button_bar) in chain, chain
    assert str(win._inner) not in chain, "按钮还在滚动内容里：%s" % chain


def test_button_does_not_move_when_content_scrolls_to_bottom(ui):
    """🔴 客户视角的判据：滚到底，按钮屏幕坐标一动不动。"""
    _require_desktop(ui)
    win = _open_confirm(ui)
    btn = win._primary_btn
    _settle(ui.root, win)

    before_y = btn.winfo_rooty()
    before_x = btn.winfo_rootx()

    # 内容必须真的溢出（否则这条用例是空转）
    first, last = win._cv.yview()
    assert last - first < 0.999, "内容没有溢出可视区，用例不成立"

    win._cv.yview_moveto(1.0)          # 滚到最底
    _settle(ui.root, win)

    assert btn.winfo_rooty() == before_y, \
        "滚到底后按钮跑位了：%s → %s" % (before_y, btn.winfo_rooty())
    assert btn.winfo_rootx() == before_x


def test_action_bar_is_visible_without_scrolling(ui):
    """弹窗一打开（还没滚过），按钮就必须已经在可视区里。"""
    _require_desktop(ui)
    win = _open_confirm(ui)
    _settle(ui.root, win)
    assert win._cv.yview()[0] == pytest.approx(0.0, abs=1e-6), "刚打开时不该是滚过的状态"
    btn = win._primary_btn
    assert btn.winfo_ismapped()

    win_h = win.winfo_height()
    bar_top = win._button_bar.winfo_y()
    # 按钮栏必须落在窗口内（相对窗口的 y 小于窗口高）
    assert 0 <= bar_top < win_h, (bar_top, win_h)
    # 按钮底边不能超出窗口
    btn_bottom = btn.winfo_rooty() + btn.winfo_height()
    win_bottom = win.winfo_rooty() + win_h
    assert btn_bottom <= win_bottom + 2, (btn_bottom, win_bottom)


def test_scroll_region_covers_whole_content(ui):
    """吸底后滚动区仍须覆盖全部内容 —— 不能因为按钮栏搬走就把末段裁掉。"""
    win = _open_confirm(ui)
    _settle(ui.root, win)
    x1, y1, x2, y2 = win._cv.cget("scrollregion").split()
    inner_h = int(win._inner.winfo_reqheight())
    assert float(y2) >= inner_h - 2, (y2, inner_h)


def test_window_height_accounts_for_action_bar(ui):
    """吸底弹窗的窗口高度必须容得下「内容 + 按钮栏」。

    反例（旧 _fit_modal）：窗口高 = 内容高，按钮栏再从里面切走一条 →
    画布比内容矮一条按钮栏，明明装得下也永远挂滚动条、末行被裁在按钮后面。
    """
    _require_desktop(ui)
    win = _open_confirm(ui)
    _settle(ui.root, win)
    inner_h = int(win._inner.winfo_reqheight())
    bar_h = int(win._button_bar.winfo_reqheight())
    assert bar_h > 0
    # 画布可视高 + 按钮栏高 必须正好等于窗口客户区高（不多不少）
    cv_h = win._cv.winfo_height()
    assert abs((cv_h + bar_h) - win.winfo_height()) <= 4, \
        (cv_h, bar_h, win.winfo_height())
    # 只有在「屏幕本来装得下」时才苛求装得下；内容比屏幕还高时必然被夹，
    # 此时唯一要求是「画布 + 按钮栏 = 客户区高」（上一行已验），别拿屏幕高度当判据
    # —— 我第一版就是拿 screenheight-64 当「是否被夹」的判据，720 的小屏上必然假报。
    screen_h = int(ui.root.winfo_screenheight())
    if inner_h + bar_h <= screen_h - 96:
        assert win.winfo_height() >= inner_h + bar_h - 8, \
            (win.winfo_height(), inner_h, bar_h)
    else:
        # 装不下 → 必须真的挂出滚动条，否则末段内容被裁又滚不到
        assert win._sb.winfo_ismapped(), "内容溢出但滚动条没挂出来"
