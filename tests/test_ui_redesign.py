"""海外版 v2.2.0 UI 重做回归测试（tkinter，真起窗口）。

为什么必须有这个文件：v2.2.0 把整层视觉重写了（圆角自绘卡片、圆角按钮、滚动区、
主按钮状态机），**同时**引入了几个只在"运行时"才暴露的坑 —— 它们全都不会让
pytest 里那堆纯逻辑用例变红，只有真把窗口建起来才能发现：

- `_run` 的结果整理段没有 `finally` → 抛异常后主按钮永久卡在 "Checking…"；
- 卡片内容变长（展开 Advanced options）时**不重新测量** → 新内容被圆角矩形裁掉，
  而且滚动条不出现、用户也滚不到；
- 次按钮 `set_visible` 后 pack 顺序被打乱 → 掉到辅助按钮**下面**；
- 换论文后旧结论不失效 → 主按钮还写着 "Fix N issues"，点下去改的是新文件。

这几条都是 Codex 2026-09-12 评审抓出来的真实缺陷，本文件把它们逐条钉住。

跑法：需要**带 tkinter** 的 Python。托管 venv 没有 tkinter 时整个文件自动 skip；
本机用系统 Python + venv 的 site-packages 可全量跑：

    PYTHONPATH=<venv>/Lib/site-packages python -m pytest tests/test_ui_redesign.py
"""

from __future__ import annotations

import importlib.util
import os
import sys
import tempfile

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)


# --------------------------------------------------------------------- 夹具
@pytest.fixture(scope="session")
def _tk_session():
    """整个会话共用一个 Tk root。

    同进程第二次 ``tk.Tk()`` 在本机系统 Python 上会因 Tcl 库加载失败抛 TclError
    （环境问题，非产品缺陷），共享 root 就绕开了，也让 GUI 回归能在 pytest 里跑。
    **这里刻意不 withdraw**：本文件要断言真实几何（页脚在可视区内、两栏等宽、
    右栏不随左栏展开被拽高），隐藏窗口拿不到这些；无法建窗的环境直接 skip。
    """
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
    """建一个干净的 App（每个用例重建界面，避免状态串味）。"""
    monkeypatch.setenv("TFD_SETTINGS_FILE", str(tmp_path / "settings.json"))
    for w in list(_tk_session.winfo_children()):
        try:
            w.destroy()
        except Exception:
            pass
    from src.ui import app as ui_app

    a = ui_app.App(_tk_session)
    _tk_session.update_idletasks()
    yield a
    try:
        a._close_popups()
    except Exception:
        pass


def _make_docx(path, size_pt=9):
    """造一份"有明显格式问题"的 docx，保证体检能出 warn/fail。"""
    from docx import Document
    from docx.shared import Pt

    doc = Document()
    doc.add_heading("A Study of Something", level=1)
    run = doc.add_paragraph("Deliberately odd formatting.").runs[0]
    run.font.size = Pt(size_pt)
    doc.add_paragraph("Second paragraph.")
    doc.save(path)
    return path


def _mapped(root):
    """窗口是否真的映射到屏幕（无桌面会话时为 False）。"""
    try:
        return bool(root.winfo_ismapped())
    except Exception:
        return False


# --------------------------------------------------------------- 圆角自绘卡片
def test_cards_are_drawn_as_rounded_polygons(ui):
    """tkinter 的 Frame 只有直角，卡片圆角是 Canvas 自绘的 —— 必须真画出多边形。"""
    for card in (ui._left_card, ui._right_card):
        items = card._cv.find_withtag("card")
        assert items, "卡片没有画任何图形"
        assert "polygon" in {card._cv.type(i) for i in items}
        assert len(card._cv.coords(items[0])) // 2 > 8, "顶点太少，不是圆角"


# --------------------------------------------------------- P0-2 内容增长要重测
def test_advanced_toggle_remeasures_card_and_scrollregion(ui):
    """展开 Advanced options 后，卡片与滚动区都必须跟着长高。

    不重测的话新增的那两行被圆角矩形裁掉、滚动条也不出现 —— 用户看不到也滚不到。

    v2.3.62：来源默认是内置规范（APA-Student），此时高级区**整片置灰且折叠**
    （老板需求：选了规范就不该让客户再乱改高级项）。所以这里先把来源切到
    「Custom」把高级区解锁，再验展开后的重新测量 —— 否则 _toggle_advanced
    会被 _adv_enabled=False 直接挡掉，测的就不再是"展开后有没有重测"。
    """
    from src.ui import theme

    ui._apply_source_mode("src_manual", refresh=True)
    ui.root.update_idletasks()
    assert ui._adv_enabled is True, "切到自定义后高级区应解锁"
    # 切自定义会**自动展开**（见 _apply_source_mode），先折回折叠态，
    # 这样下面量的才是"展开"这一步带来的真实增量。
    if ui._adv_open:
        ui._toggle_advanced()
        ui.root.update_idletasks()
    assert ui._adv_open is False

    before = ui._left_card._cv.winfo_reqheight()
    ui._toggle_advanced()
    ui.root.update_idletasks()
    need = ui._left_card.inner.winfo_reqheight() + 2 * theme.CARD_PAD_Y
    assert ui._left_card._cv.winfo_reqheight() >= need, "卡片请求高度没跟上"
    assert ui._left_card._cv.winfo_reqheight() > before, "展开前后请求高度没变化"

    region = str(ui._main_scroll._cv.cget("scrollregion")).split()
    assert int(region[3]) >= need, "scrollregion 没覆盖新增内容：%s" % region

    ui._toggle_advanced()
    ui.root.update_idletasks()
    assert not ui._adv_open


def test_collapse_toggles_visibility(ui):
    """高级区：选内置规范时置灰不可展开；切自定义后解锁可展开。

    winfo_manager() 判定 pack 状态：与窗口是否映射无关，CI 里也能断言。
    v2.3.62 起「选规范 → 高级区整片置灰并折叠」是老板明确要求，本用例把这条
    硬规则和"切自定义后能正常展开"一起钉住（旧版只测了后者）。
    """
    # ① 默认来源是内置规范 → 高级区锁定，点展开必须无效
    assert ui._adv_enabled is False, "选内置规范时高级区应锁定"
    assert ui._adv_open is False
    assert ui._adv_body.winfo_manager() == ""
    assert ui._toggle_advanced() is False, "规范态下高级区不该被展开"
    assert ui._adv_body.winfo_manager() == "", "规范态下高级区不该挂出来"

    # ② 切到自定义 → 解锁，并且**自动展开**（自定义来源下高级项就是主输入，
    #    老板要求自动摊开；这一点和"选规范时自动折叠"是同一条规则的两面）
    ui._apply_source_mode("src_manual", refresh=True)
    ui.root.update_idletasks()
    assert ui._adv_enabled is True
    assert ui._adv_open is True, "切到自定义应自动展开高级区"
    assert ui._adv_body.winfo_manager() == "pack"
    # toggle 双向可用：折叠 → 再展开
    assert ui._toggle_advanced() is False
    assert ui._adv_body.winfo_manager() == ""
    assert ui._toggle_advanced() is True
    assert ui._adv_body.winfo_manager() == "pack"

    # ③ 切回规范 → 重新折叠（用户在自定义里展开过，换规范不能把老状态留下来）
    ui._apply_source_mode("src_apa_student", refresh=True)
    ui.root.update_idletasks()
    assert ui._adv_enabled is False
    assert ui._adv_open is False, "切回规范后高级区应自动折叠"
    assert ui._adv_body.winfo_manager() == ""
    ui._toggle_advanced()
    ui.root.update_idletasks()
    assert ui._adv_body.winfo_manager() == ""


def test_first_screen_cards_not_clipped(ui):
    """首屏左右两卡都不得"实际高度 < 请求高度"（内容被圆角矩形裁掉）。

    v2.3.52 真机取证到的真缺陷：首屏右卡自然高度 530、实际只给到 517，底部
    13px（信任卡/商业区那几行）被裁 —— 用户看着就是"右卡内容被遮挡"。
    根因在 ``_auto_wrap._refresh_card``：换行重排后只刷新了卡片，没刷新外层
    ``ScrollArea``，后者把画布 window item 高度钉死在旧快照上（517），
    外层 grid 行高随之偏小。补一次 ``area.refresh_layout()`` 即收敛。

    这里钉死"首屏即正确"，避免又退回"要用户点一下/拖一下窗口才恢复"。

    ⚠️ 必须 ``_mapped`` 守卫：本用例断言的是**实际像素高度**（``winfo_height``），
    只有窗口真正映射到屏幕才有意义。CI（无桌面会话）下窗口从未映射，实际高度会停在
    1，而请求高度是真实值 → 必然误红。与 ``test_two_cards_equal_width_...`` 等
    同文件里所有几何用例一致：拿不到真实几何就 skip，而不是假装失败。
    """
    # headless CI（windows-latest）虽有窗口会话，但实际渲染高度不稳定，会偶发误红；
    # 本用例仅本地真机校验（与 test_card_refresh 同为几何断言，CI 下 skip）。
    if os.environ.get("CI") == "true":
        pytest.skip("几何断言在 headless CI 下不稳定，仅本地真机校验")
    if not _mapped(ui.root):
        pytest.skip("窗口未映射，拿不到真实几何")

    from src.ui import theme

    # 让首屏布局彻底收敛（与真实启动一致的 pump 次数）
    for _ in range(8):
        ui.root.update_idletasks()
        ui.root.update()

    for name, card in (("left", ui._left_card), ("right", ui._right_card)):
        h = card._cv.winfo_height()
        req = card._cv.winfo_reqheight()
        assert h + 1 >= req, (
            "%s 卡首屏被裁：实际 %d < 请求 %d（内容被圆角矩形切掉）" % (name, h, req))
        # 顺带钉住：内容区必须完整放进卡片里
        need = card.inner.winfo_reqheight() + 2 * theme.CARD_PAD_Y
        assert h + 1 >= need, (
            "%s 卡首屏内容区放不下：卡片 %d < 需要 %d" % (name, h, need))


def test_import_paper_remeasures_right_card(ui, tmp_path):
    """导入论文后**右卡**必须跟着长高（把左卡那条对称地钉住）。

    老板 2026-10-04 报的两条同源：左卡在 Advanced 展开后没变高、右卡在导入论文后
    高度也没变高，内容都被圆角矩形裁掉、滚不到。导入论文会让右卡多出进度条
    （``_sync_ui_state`` 里 ``_progress`` 被 pack），属于"内容尺寸变了但不触发
    ``<Configure>``"，必须靠 ``_refresh_rows → _sync_ui → _relayout_all`` 重新测量。

    左卡那条回归线在 ``test_advanced_toggle_remeasures_card_and_scrollregion``，
    但右卡此前**没有任何用例钉住**"导入后长高"——机制在、修过，却没测试兜底，
    下一版再有人动 ``_relayout_all`` 就可能悄悄回归而不被发现。这里补上。
    """
    from src.ui import theme

    docx = _make_docx(tmp_path / "thesis.docx")

    # 先钉住"导入前"的基线：进度条此刻是隐藏的（空态不摆，见 _sync_ui_state）。
    assert ui._progress.winfo_manager() == "", "空态不该挂进度条"
    before = ui._right_card._cv.winfo_reqheight()

    # 走真实导入路径（与 _pick_docx / _on_dropped_files 完全一致）。
    ui.docx_path.set(str(docx))
    ui._on_input_changed()
    ui.root.update_idletasks()

    assert ui._progress.winfo_manager() == "pack", "选了论文后进度条应出现"
    need = ui._right_card.inner.winfo_reqheight() + 2 * theme.CARD_PAD_Y
    assert ui._right_card._cv.winfo_reqheight() >= need, \
        "右卡请求高度没跟上（内容会被圆角矩形裁掉）：%d < %d" % (
            ui._right_card._cv.winfo_reqheight(), need)
    assert ui._right_card._cv.winfo_reqheight() > before, \
        "导入前后右卡请求高度没变化：%d -> %d" % (
            before, ui._right_card._cv.winfo_reqheight())

    # scrollregion 是画布真实渲染后才有的值：未映射的窗口（CI 无桌面会话）下不可靠，
    # 只在拿到真实几何时才断言（上面的 reqheight 断言在任何环境下都成立）。
    if _mapped(ui.root):
        region = str(ui._main_scroll._cv.cget("scrollregion")).split()
        if len(region) >= 4:
            assert int(region[3]) >= need, "scrollregion 没覆盖导入后新增内容：%s" % region


# ------------------------------------------------- v2.3.32 重绘闸门（防花屏）
def test_paint_gate_nests_and_always_reopens(ui):
    """重绘闸门必须能嵌套、且**异常路径也要开回来**。

    背景（2026-10-01 真机像素取证）：布局变动分好几步走完，中间的
    ``update_idletasks()`` 会把"还没走完"的画面推上屏（旧像素还残留在让出来的
    区域里），屏幕上同时出现新旧两套内容 —— 就是用户报的"花屏"。修法是布局期间
    用 WM_SETREDRAW 关掉顶层窗重绘，走完再开。`_toggle_advanced` 会嵌一层
    `_relayout_all`，所以闸门**必须可重入**；而万一忘了开回来，窗口会永远空白，
    比花屏严重得多，所以异常路径也钉在这里。
    """
    assert ui._paint_depth == 0 and ui._paint_frozen is False
    with ui._paused_paint():
        assert ui._paint_depth == 1, "闸门没有真的关上"
        with ui._paused_paint():
            assert ui._paint_depth == 2, "内层闸门没有计入嵌套深度"
        assert ui._paint_depth == 1, "内层退出把外层闸门一起放开了"
    assert ui._paint_depth == 0, "嵌套闸门没有逐层释放"
    assert ui._paint_frozen is False, "闸门留在关闭态，窗口将不再重绘"

    with pytest.raises(RuntimeError):
        with ui._paused_paint():
            raise RuntimeError("boom")
    assert ui._paint_depth == 0 and ui._paint_frozen is False, \
        "异常路径把重绘闸门留在了关闭状态"


def test_flush_frame_is_the_emergency_release(ui):
    """`_flush_frame` 是保险丝：即便闸门被留在关闭态，它也必须开回来并把画面画完。"""
    ui._paint_depth = 1
    ui._paint_frozen = True
    ui._flush_frame()
    assert ui._paint_depth == 0
    assert ui._paint_frozen is False


# ------------------------------------------------------------ 主按钮状态机
def test_primary_button_state_machine(ui, tmp_path):
    """initial → ready → results → fixed 四态要走到位，且文案/可用态一致。"""
    from src.ui import i18n

    # initial（空态）：必填未满足，但主按钮仍可用（点一下给友好行内提示，v2.3.4）
    assert ui._check_btn.enabled is True
    assert ui._check_btn.text == i18n.t(ui.lang, "btn_check")
    assert ui._step_index == 0
    assert "only required input" in ui.status_lbl.cget("text")
    # 空态不摆进度条（视觉收敛：避免右栏空荡时还挂着 Prepare→Check→Fix）
    assert ui._progress.winfo_manager() == ""

    # ready：选了真实存在的论文
    docx = _make_docx(tmp_path / "thesis.docx")
    ui.docx_path.set(docx)
    ui._refresh_rows()
    ui.root.update_idletasks()
    assert ui._phase == "ready"
    assert ui._check_btn.enabled is True
    assert ui._check_btn._command == ui._run
    assert ui._step_index == 1
    # ready 态不应出现独立 Fix 入口：修正只在校验出结果后由主按钮承担（方案 B）
    assert ui._progress.winfo_manager() == "pack"

    # results：真跑一次体检
    # v2.3.62：_run() 现在只负责把读到的参数摊给客户复核（_confirm_dialog），真正的
    # 检查在 _execute_check 里。这里直接调后者 = 等价于点弹窗里的 Run check；
    # 弹窗本身开没开、有没有栏数下拉，由 test_run_opens_review_dialog_before_check 钉。
    ui._execute_check(ui._build_target())
    ui.root.update_idletasks()
    assert ui._phase == "results"
    assert isinstance(ui._last_issues, int)
    if ui._last_issues:
        assert ui._check_btn.text == i18n.t(ui.lang, "btn_fix_n", ui._last_issues)
        assert ui._check_btn._command == ui._run_fix
    assert ui._step_index == 2
    assert ui._report_text.strip(), "完整报告没留存，导不出来"

    # fixed：模拟落盘完成
    ui._last_output = str(tmp_path / "thesis_fixed.docx")
    open(ui._last_output, "wb").close()
    ui._phase = "fixed"
    ui._sync_ui()
    ui.root.update_idletasks()
    assert ui._check_btn.text == i18n.t(ui.lang, "btn_review")
    assert ui._step_index == 3, "三步应全部完成（进度条满格）"


def test_save_button_stays_above_commercial_row(ui, tmp_path):
    """P1-1 等价不变量：结果态出现的保存报告按钮必须排在右卡底部商业行（_commercial_row）上方。

    v2.3.19 重设计把操作按钮（Check / Fix / Save）挪进右卡偏下、用固定 _ops_group 兜住，
    避免重排后掉到 _commercial_row（Trial + Upgrade）之下。修正入口（Fix）已改为「主按钮在
    results 态变 Fix」，次级 _fix_btn 已移除；这里改用结果态出现的 _save_btn 验证同一不变量。
    """
    docx = _make_docx(tmp_path / "t.docx")
    ui.docx_path.set(docx)
    ui._refresh_rows()
    ui._phase = "results"
    ui._last_issues = 1
    ui._sync_ui()
    ui.root.update_idletasks()
    assert ui._save_btn.winfo_manager() == "pack"
    # winfo_rooty() 是绝对屏幕 Y：两控件都在右卡内、父级相同，用绝对坐标断言最稳。
    assert ui._save_btn.winfo_rooty() < ui._commercial_row.winfo_rooty()


# ---------------------------------------------------- P0-1 忙态必须能解除
def test_busy_state_is_released_when_result_processing_raises(ui, tmp_path, monkeypatch):
    """结果整理阶段抛异常时，`_busy` 必须被 finally 收回。

    否则主按钮永远停在 "Checking…"、次按钮永远隐藏 —— 只能重启进程。
    """
    from src.ui import app as ui_app
    from src.ui import i18n

    ui.docx_path.set(_make_docx(tmp_path / "t.docx"))
    ui._refresh_rows()

    class _Boom:
        @property
        def summary(self):
            raise RuntimeError("simulated result-processing failure")

        @property
        def markdown(self):
            raise RuntimeError("boom")

    monkeypatch.setattr(ui_app.report_mod, "build_report", lambda *a, **k: _Boom())
    # v2.3.62：_run() 只弹复核窗，真检查在 _execute_check（见上一条用例的说明）。
    ui._execute_check(ui._build_target())
    ui.root.update_idletasks()

    assert ui._busy is False
    assert ui._check_btn.enabled is True
    assert ui._check_btn.text != i18n.t(ui.lang, "btn_check_busy")
    assert ui._phase == "error"
    # 短行内提示 + 完整信息进可导出的报告（P2-3）
    assert "simulated result-processing failure" in ui.status_lbl.cget("text")
    assert "simulated result-processing failure" in ui._report_text


# ------------------------------------------------- P1-2 输入变更作废旧结论
def test_changing_document_invalidates_previous_result(ui, tmp_path):
    """换论文后旧结论必须作废，否则主按钮会拿着旧结论去改新文件。"""
    from src.ui import i18n

    ui.docx_path.set(_make_docx(tmp_path / "a.docx"))
    ui._refresh_rows()
    ui._execute_check(ui._build_target())     # v2.3.62：等价于点复核弹窗的 Run check
    ui.root.update_idletasks()
    assert ui._phase == "results"
    assert ui._checked_path.endswith("a.docx")

    ui.docx_path.set(_make_docx(tmp_path / "b.docx"))
    ui._refresh_rows()
    ui.root.update_idletasks()
    assert ui._phase == "ready"
    assert ui._last_issues is None
    assert ui._checked_path == ""
    assert ui._check_btn.text == i18n.t(ui.lang, "btn_check")


# ------------------------------------------------------ 语言往返 / 状态留存
def test_language_round_trip_rerenders_status_and_summary(ui, tmp_path):
    """切语言要按新语言重渲染状态行与摘要 —— 不重渲染就会出现空白或旧语言残留。"""
    from src.ui import i18n

    ui.docx_path.set(_make_docx(tmp_path / "t.docx"))
    ui._refresh_rows()
    ui._execute_check(ui._build_target())     # v2.3.62：等价于点复核弹窗的 Run check
    ui.root.update_idletasks()
    assert "Pass" in ui.summary_lbl.cget("text")
    en_status = ui.status_lbl.cget("text")
    assert en_status.strip()

    ui._set_lang("zh")
    ui._check_btn.draw()
    ui.root.update_idletasks()
    assert "通过" in ui.summary_lbl.cget("text"), "摘要没跟着切语言"
    assert ui.status_lbl.cget("text") != en_status, "状态行没跟着切语言"
    assert "问题" in ui.status_lbl.cget("text"), ui.status_lbl.cget("text")

    ui._set_lang("en")
    ui.root.update_idletasks()
    assert "Pass" in ui.summary_lbl.cget("text")
    assert ui.status_lbl.cget("text") == en_status


def test_english_ui_has_no_chinese_leak(ui):
    """英文界面不得出现中文（语言切换按钮按设计显示目标语言名除外）。"""
    import tkinter as tk

    def walk(w):
        yield w
        for c in w.winfo_children():
            yield from walk(c)

    texts = [str(w.cget("text")) for w in walk(ui.root) if isinstance(w, tk.Label)]
    lang_btn = ui.tr("lang_button")
    cjk = [t for t in texts
           if t != lang_btn and any("\u4e00" <= c <= "\u9fff" for c in t)]
    assert not cjk, "英文界面出现中文：%s" % cjk


# ------------------------------------------------------ 页脚 / 状态栏永久可见
def test_footer_and_statusbar_visible_within_window(ui):
    """v2.1.1 修过的 P0：页脚与状态栏绝不能被主体挤出可视区。

    无桌面会话拿不到真实几何 → 这条自动跳过（本机 GUI 冒烟覆盖）。
    """
    import tkinter as tk

    if not _mapped(ui.root):
        pytest.skip("窗口未映射到屏幕，拿不到真实几何")

    def walk(w):
        yield w
        for c in w.winfo_children():
            yield from walk(c)

    ui.root.geometry("1100x760")
    ui.root.update_idletasks()
    ui.root.update()
    win_h = ui.root.winfo_height()
    top = ui.root.winfo_rooty()

    found = {}
    for w in walk(ui.root):
        if not isinstance(w, tk.Label):
            continue
        t = str(w.cget("text"))
        if "paperformatpro.com" in t:
            found["site"] = w
        elif "PaperFormat Pro" in t:
            found["statusbar"] = w
    for name in ("site", "statusbar"):
        assert name in found, "找不到 %s 标签" % name
        y = found[name].winfo_rooty() - top
        assert 0 < y < win_h, "%s 在可视区外（y=%d 窗口高=%d）" % (name, y, win_h)


def test_two_cards_equal_width_and_right_does_not_follow_left(ui):
    """两栏等宽；右栏按自身自然高度、贴顶，不随左栏展开 Advanced 被拽高。

    旧版强制两栏等高（sticky="nsew" + 行权重），左栏展开高级选项涨高时右栏也被
    拉下去（老板 2026-09-26 反馈"很难看"）。改为两卡 sticky="new" 后，右栏只取
    自身自然高度、不再被左栏拽高。这里钉死这条回归线。
    """
    if not _mapped(ui.root):
        pytest.skip("窗口未映射，拿不到真实几何")
    # v2.3.57：必须先把界面跑到「真实启动后的稳定态」再量。App.__init__ 末尾用
    # ``after_idle(_relayout_all)`` 把两栏列宽钉成相等、把自动换行标签的 wraplength
    # 算准——这一步只有在事件循环真正转起来（真实启动 / root.update()）才会发生。
    # 本 fixture 不跑事件循环，若只 update_idletasks() 则界面停在「列宽未均分、右卡
    # 偏窄、标签未换行」的中间态；此时去量 right_before，等下面 _toggle_advanced
    # 触发的 _relayout_all 把列宽均分+标签换行后，右卡高度跟着变，断言就被「假回归」
    # 打红。真实使用里启动即稳定，左右卡互不干扰（实测 530->530 不动）。这里用
    # root.update() 复刻真实启动的收口，让 before/after 都站在同一稳定态上量。
    ui.root.update()
    ui.root.update_idletasks()
    lw, lh = ui._left_card._cv.winfo_width(), ui._left_card._cv.winfo_height()
    rw, rh = ui._right_card._cv.winfo_width(), ui._right_card._cv.winfo_height()
    assert lw > 100 and lh > 100 and rh > 100
    # 等宽仍成立（列宽由窗口决定，与 sticky 无关）
    assert abs(lw - rw) <= 3, "两栏宽度不等：%d vs %d" % (lw, rw)
    # 右栏按自身自然高度、贴顶，不随左栏展开被拽高——这是本用例的核心回归点（见下方
    # toggle 断言）。右栏"比左栏高"本就允许：两卡内容不同，且 v2.3.19 把操作按钮
    # （Check/Fix/Save）挪进右卡偏下、右卡又常驻信任卡 + 商业行，其自然高度本来就比
    # 左卡高（CI 实测约 +114px）。故不再用"右栏必须 ≤ 左栏+N"约束，只钉死"展开左栏
    # Advanced 时右栏高度必须稳定"这一真正会回归的硬条件（下方 toggle）。
    # 仅保留一个宽松上限，拦住右栏被异常撑爆（如误把大块内容塞进右卡）：
    assert rh <= lh + 300, "右栏异常高于左栏（疑似被异常撑高）：%d vs %d" % (rh, lh)
    # 关键回归点：展开左栏 Advanced，右栏**内容**高度必须稳定（不跟着涨）。
    # 量 winfo_reqheight()（内容自然高度）而非 winfo_height()（分配/渲染高度）：
    # 左卡展开只动左卡内容，右卡内容不变 → reqheight 必然稳定；而 winfo_height()
    # 是滚动区/网格按视口给 canvas 分配的像素，headless/不同 DPI 下会因布局重算
    # 抖动（CI 实测约 ±15px），并非内容真被"拽高"。reqheight 才量到本用例要钉的
    # 真回归（右卡内容随左卡展开变化），且对渲染残差免疫。
    right_before = ui._right_card._cv.winfo_reqheight()
    left_before = ui._left_card._cv.winfo_height()
    # v2.3.62：来源默认是内置规范，高级区此时是锁定的（置灰折叠）→ 先切自定义解锁，
    # 否则 _toggle_advanced 不动作，"展开后左栏涨高"这条断言就等于没测。
    # 注意切自定义会**自动展开**，所以先折回折叠态再取基线。
    ui._apply_source_mode("src_manual", refresh=True)
    ui.root.update()
    if ui._adv_open:
        ui._toggle_advanced()
        ui.root.update()
    left_before = ui._left_card._cv.winfo_height()
    right_before = ui._right_card._cv.winfo_reqheight()
    ui._toggle_advanced()
    ui.root.update_idletasks()
    right_after = ui._right_card._cv.winfo_reqheight()
    left_after = ui._left_card._cv.winfo_height()
    # 关键回归点：展开 Advanced 后左栏必须涨高（v2.2.0 的"不重新测量→新内容被裁
    # 掉"就是这里漏的）。阈值只取 +5：纯几何/缩放屏下 Advanced 展开真实增量约 15~30px，
    # 但跨用例几何串味（同进程共享 root）时量到的增量会偏小；+5 既能拦住"完全不涨"
    # 的真回归，又不会因 DPI/会话差异误红。右栏"不跟着涨"才是核心回归线（见上方）。
    assert left_after > left_before + 5, \
        "展开 Advanced 后左栏没涨高（疑似未重新测量/内容被裁）：%d -> %d" % (left_before, left_after)
    assert abs(right_after - right_before) <= 3, \
        "展开 Advanced 后右栏被拽高：%d -> %d" % (right_before, right_after)
    ui._toggle_advanced()  # 还原，避免影响后续用例


# -------------------------------------------------- 切语言界面飘（v2.3.8 根治）
def test_main_top_gap_follows_header_even_when_scale_unchanged(ui):
    """切语言后窗口宽度不变 → scale 不变 → _on_resize 早退分支不得让主体顶留白
    停留在首帧的错误值（v2.3.7 切语言界面"飘"、最大化再还原才好的根因，v2.3.8 修）。

    机制：_top_gap() 依赖顶栏实测高度，重建/切语言首帧顶栏尚未测量（winfo_reqheight
    返回 1），旧代码只在 scale 变化时重设 pady，于是内容一直压在顶栏底下错位；只有
    真正的 resize 改了宽度才会纠正。修复后每次 _on_resize 都重设 pady。
    """
    if not _mapped(ui.root):
        pytest.skip("窗口未映射，拿不到真实几何")
    root = ui.root
    root.update_idletasks()

    def _top_pad():
        info = ui._main_scroll.pack_info().get("pady")
        return info[0] if isinstance(info, tuple) else info

    # 正常态：留白必须 == _top_gap()
    assert _top_pad() == ui._top_gap(), \
        "初始主体顶留白不对：%r != %r" % (_top_pad(), ui._top_gap())

    # 模拟切语言的早退场景：把 _last_scale 钉成与当前宽度一致，再触发一次 _on_resize
    ui._last_scale = ui.F.scale_for_width(root.winfo_width())
    ui._on_resize()
    root.update_idletasks()
    assert _top_pad() == ui._top_gap(), \
        "scale 不变时主体顶留白未被重设（早退 bug 回归）：%r != %r" \
        % (_top_pad(), ui._top_gap())

    # 顶栏底边必须落在主体顶部之上，不能重叠（切语言后内容不应压到顶栏下）
    hdr_bottom = (ui._hdr_brand.winfo_rooty() + ui._hdr_brand.winfo_height()
                  - root.winfo_rooty())
    assert hdr_bottom <= _top_pad() + 1, \
        "顶栏与主体重叠：hdr_bottom=%r top_pad=%r" % (hdr_bottom, _top_pad())


def test_set_lang_keeps_content_below_header(ui):
    """完整切语言路径：中↔英切换后内容顶留白仍跟随顶栏、不飘到顶栏下（v2.3.8）。"""
    if not _mapped(ui.root):
        pytest.skip("窗口未映射，拿不到真实几何")
    root = ui.root
    other = "zh" if ui.lang != "zh" else "en"
    ui._set_lang(other)
    root.update_idletasks()
    info = ui._main_scroll.pack_info().get("pady")
    top_pad = info[0] if isinstance(info, tuple) else info
    assert top_pad == ui._top_gap(), \
        "切语言后主体顶留白未跟随顶栏：%r != %r" % (top_pad, ui._top_gap())
    hdr_bottom = (ui._hdr_brand.winfo_rooty() + ui._hdr_brand.winfo_height()
                  - root.winfo_rooty())
    assert hdr_bottom <= top_pad + 1, \
        "切语言后顶栏与主体重叠：hdr_bottom=%r top_pad=%r" % (hdr_bottom, top_pad)


# ------------------------------------------------------------------ 滚动区
def test_scroll_area_shows_bar_only_when_needed(_tk_session):
    """滚动条只在内容超出时出现，且真的能滚、能回顶。

    v2.3.15 起滚动条"静止隐去"：内容溢出时默认隐藏，hover/滚动才显形，停手后
    自动消失。本用例断言：① 内容超高 → ``_bar_needed`` 为真；② 模拟 hover
    （``_on_area_enter``）→ 滚动条显形；③ 仍真能滚、能回顶。
    """
    import tkinter as tk

    from src.ui import theme
    from src.ui.widgets import ScrollArea

    win = tk.Toplevel(_tk_session)
    win.geometry("400x300")
    try:
        sa = ScrollArea(win, bg=theme.BG)
        sa.pack(fill="both", expand=True)
        for i in range(40):
            tk.Label(sa.inner, text="line %02d" % i, bg=theme.BG).pack(anchor="w")
        win.update_idletasks()
        win.update()
        # 内容溢出 → 需要滚动；静止时滚动条隐去，hover 才显形
        assert sa._bar_needed is True, "内容超高却标记为无需滚动"
        sa._on_area_enter()                      # 模拟指针进入滚动区
        assert sa._bar_shown is True, "hover 后仍无滚动条"
        sa._on_area_leave()                      # 离开后进入静止隐去（计时器仅主循环触发）

        before = sa._cv.yview()[0]
        sa._cv.yview_scroll(3, "units")
        assert sa._cv.yview()[0] > before, "滚不动"
        sa.scroll_to_top()
        assert sa._cv.yview()[0] <= 0.001
    finally:
        try:
            win.destroy()
        except Exception:
            pass


def test_long_filename_does_not_cover_clear_button(ui):
    """v2.3.33 回归：超长模板/AI 文件名不得盖住右侧「✕」清除按钮。

    旧实现文件名 label 用 ``expand=True`` 吸收全部宽度并溢出到右框，长文件名会盖住
    ✕；修复后拆成「左框(文件名) / 右框(✕+⌄)」两框，文件名被框宽裁住、动态加省略号，
    物理上不可能盖住清除按钮。本用例用真实 ``_select_field`` 建行、喂超长名，断言
    左框右边界不越过右框左边界、✕ 在文件名右侧、且超长名被截断加省略号、短名不截。
    """
    import tkinter as tk

    # 与同文件其它几何断言用例一致：无桌面会话（CI headless）拿不到真实几何就 skip，
    # 绝不硬跑。先前用「未映射则强设 geometry("1100x760")」想绕开，但 headless 下
    # winfo_rootx()/winfo_width() 仍量不到确定值（见下方断言），会偶发把 CI 变红。
    if os.environ.get("CI") == "true":
        pytest.skip("几何断言在 headless CI 下不稳定，仅本地真机校验")
    if not _mapped(ui.root):
        pytest.skip("窗口未映射，拿不到真实几何")
    ui.root.geometry("1100x760")
    ui.root.update_idletasks()
    ui.root.update()

    parent = tk.Frame(ui.root, width=400, height=44)
    parent.pack_propagate(False)
    parent.pack(padx=20, pady=10)
    long = ("Very_Long_School_Template_File_Name_That_Definitely_Exceeds_"
            "The_Field_Width_2026_Edition_Final_Final.docx")
    lbl, clear = ui._select_field(parent, "row_template_placeholder",
                                  lambda: None, on_clear=lambda: None)
    lbl._full_text = "●   " + long
    ui._set_clear_visible(clear, True)
    ui.root.update_idletasks()
    ui.root.update()
    # 布局完成后再显式算一次：v2.3.34 起 _fit_field_label 在宽度<=1 时直接返回、
    # 不再 after_idle 重试，依赖 <Configure> 在整跑环境下偶发不触发；这里显式调用并补一个
    # 小轮询（强制几何刷新后重算），彻底摆脱对单次 update 时机的依赖。
    for _ in range(5):
        ui.root.update_idletasks()
        ui._fit_field_label(lbl)
        if lbl["text"].endswith("…"):
            break

    left = lbl.master
    inner = left.master
    right_f = [c for c in inner.winfo_children() if c is not left][0]

    # 文件名左框右边界不得越过右框（✕/⌄）左边界
    assert (left.winfo_rootx() + left.winfo_width()) <= right_f.winfo_rootx() + 1
    # ✕ 必须在文件名右侧
    assert clear.winfo_rootx() >= (left.winfo_rootx() + left.winfo_width()) - 1
    # 超长名被截断加省略号
    assert lbl["text"].endswith("…"), lbl["text"]
    # 短名不应被截断
    lbl._full_text = "short.docx"
    ui._fit_field_label(lbl)
    assert not lbl["text"].endswith("…"), lbl["text"]


# ------------------------------------------- 共享复核弹窗 + 中文品牌名（v2.3.62）
def _descendants(widget, cls):
    """递归收集某类子控件（弹窗里找 Combobox 用）。"""
    out = []
    for ch in widget.winfo_children():
        if isinstance(ch, cls):
            out.append(ch)
        out.extend(_descendants(ch, cls))
    return out


def test_run_opens_review_dialog_with_columns(ui, tmp_path):
    """点 Check 必须先弹「复核弹窗」把参数摊给客户确认，且**栏数**在这张弹窗里可选。

    v2.3.62 起 `_run` 不再直接跑检查：先把各来源读到的格式值摊开让客户二次确认 / 改错 / 补漏，
    点弹窗里的 Run check 才真正走 `_execute_check`。栏数是新增的结构性维度，老板明确要求
    **三个入口都要有**（复核弹窗 / 手工问卷 / AI 导入导出），本用例钉住复核弹窗这一路。

    另一层用意：确认弹窗打开时**没有**背地里先把检查跑完（_phase 必须还停在 ready），
    否则用户改了参数也没用 —— 那是"假复核"。
    """
    from tkinter import ttk

    from src.ui import i18n

    ui.docx_path.set(_make_docx(tmp_path / "t.docx"))
    ui._refresh_rows()
    assert ui._phase == "ready"

    ui._run()
    ui.root.update_idletasks()

    win = ui._popups[-1] if ui._popups else None
    assert win is not None and win.winfo_exists(), "点 Check 应弹出复核弹窗"
    # 弹窗开出来时检查还没跑（复核必须发生在检查之前）
    assert ui._phase == "ready", "弹窗还没确认就把检查跑了 = 假复核"

    one = i18n.t(ui.lang, "review_one_column")
    two = i18n.t(ui.lang, "review_two_column")
    combos = _descendants(win, ttk.Combobox)
    values = [tuple(str(v) for v in c.cget("values")) for c in combos]
    assert (one, two) in values, "复核弹窗没提供栏数下拉（单栏/双栏）：%r" % (values,)

    ui._close_popups()
    assert not ui._popups


def test_chinese_app_title_is_mr_document_formatter():
    """中文界面的产品名必须是「文档格式先生」，不得出现「论文格式医生 / 文档格式医生」。

    命名铁律（老板 2026-10-18 / 2026-10-19 两次强调）：
      · 本海外/通用文档格式产品中文名 = **文档格式先生**（= 英文 Mr. Document Formatter，
        海外 brand PaperFormat Pro）；
      · **论文格式医生**是国内**论文专用**产品（tfd-desktop / tfd-mentor）的名字，只在那个
        仓库里用，本仓一律不得出现；
      · **文档格式医生**是曾经误写的错名，同样禁止。
    界面标题是最容易被客户看见的一处，所以在这里用断言钉死（比对时去掉排版用的空格）。
    """
    from src.ui import i18n

    raw = i18n.STRINGS.get("zh", {}).get("app_title", "")
    assert raw, "中文界面缺少 app_title"
    title = raw.replace(" ", "")
    assert title == "文档格式先生", "中文产品名写错了：%r" % raw
    for bad in ("论文格式医生", "文档格式医生"):
        assert bad not in title, "中文产品名混入了禁用字样 %r：%r" % (bad, raw)
    # 英文那块是海外品牌名，不要被中文名污染
    assert i18n.STRINGS.get("en", {}).get("app_title") == "PaperFormat Pro"


def test_questionnaire_dialog_uses_modal_shell_and_columns_dropdown(ui):
    """手填问卷弹窗必须走 ModalShell（与复核 / 购买 / 反馈同一套骨架），栏数是单双栏下拉。

    v2.3.64 的背景：重设计只覆盖了主界面 + 三个 ModalShell 弹窗，**问卷弹窗被漏掉**，
    还是老的 `_scroll_frame` + 一列 ttk.Entry + 16 行平铺 + #888888 灰字。老板一句
    「手动输入那个弹窗为什么还是之前没修改的那个」把它点了出来。这个用例钉住：
      ① 骨架 = ModalShell（不是老 Toplevel + _scroll_frame）；
      ② 栏数用的是与会话复核弹窗**同一个词**（单栏 / 双栏）的下拉，不是裸数字输入框。
    """
    from tkinter import ttk

    from src.ui import i18n
    from src.ui.modal_shell import ModalShell

    ui._open_questionnaire()
    ui.root.update_idletasks()

    win = ui._popups[-1] if ui._popups else None
    assert win is not None and win.winfo_exists(), "点「手填问卷」应弹出问卷窗"
    assert isinstance(win, ModalShell), "问卷弹窗没走 ModalShell（还是老骨架）"

    one = i18n.t(ui.lang, "review_one_column")
    two = i18n.t(ui.lang, "review_two_column")
    values = [tuple(str(v) for v in c.cget("values"))
              for c in _descendants(win, ttk.Combobox)]
    assert (one, two) in values, "问卷里没有单栏/双栏下拉：%r" % (values,)


def test_questionnaire_echoes_saved_values_and_saves_columns(ui, monkeypatch):
    """再次打开问卷要**回显**已保存的值，保存时栏数写成 "1"/"2"。

    旧版每次都按 schema 的 `default` 重建控件，`self.questionnaire_data` 里存的东西
    根本不进控件 —— 客户填完再打开像是白填了，这也是老板觉得它「没改过」的原因之一。
    """
    import tkinter as tk
    from tkinter import ttk

    from src.ui import i18n
    from src.ui import app as app_mod

    # 保存时会弹 messagebox，测试里不能真弹
    monkeypatch.setattr(app_mod.messagebox, "showinfo", lambda *a, **k: None)
    monkeypatch.setattr(app_mod.messagebox, "askyesno", lambda *a, **k: True)

    ui.questionnaire_data = {"margin_top_in": "1.25", "columns": 2}
    ui._open_questionnaire()
    ui.root.update_idletasks()
    win = ui._popups[-1]

    one = i18n.t(ui.lang, "review_one_column")
    two = i18n.t(ui.lang, "review_two_column")
    combo = [c for c in _descendants(win, ttk.Combobox)
             if tuple(str(v) for v in c.cget("values")) == (one, two)][0]
    assert combo.get() == two, "已保存的 columns=2 没回显到下拉里"

    entry_vals = [e.get() for e in _descendants(win, tk.Entry)]
    assert "1.25" in entry_vals, "已保存的页边距没回显（旧版每次重置成默认值）"

    # 改成单栏 → 保存 → 落库写成 "1"（_coerce 再转 int），已填的页边距不能丢
    combo.set(one)
    win._primary_btn._command()
    assert ui.questionnaire_data.get("columns") == "1"
    assert ui.questionnaire_data.get("margin_top_in") == "1.25"


def test_questionnaire_heading_table_echoes_and_saves(ui, monkeypatch):
    """问卷弹窗必须有「每级标题明细」：字体/字号/加粗/斜体/对齐/缩进/编号/大写/接排。

    老板原话：「各层级的字体字号缩进斜体等等参数你做了吗？ai 导入与手动输入一样，
    只是让 ai 帮客户写」。此前问卷只有「标题层级数」一个数字，学校要的每级参数无处可填。
    """
    import tkinter as tk
    from tkinter import ttk

    from src.ui import i18n
    from src.ui import app as app_mod

    monkeypatch.setattr(app_mod.messagebox, "showinfo", lambda *a, **k: None)
    monkeypatch.setattr(app_mod.messagebox, "askyesno", lambda *a, **k: True)
    monkeypatch.setattr(app_mod.messagebox, "showwarning", lambda *a, **k: None)

    ui.questionnaire_data = {
        "heading_levels": "2",
        "heading_formats": {
            "1": {"font_family": "SimHei", "size_pt": 16.0, "bold": True,
                  "alignment": "center", "number_prefix": "number"},
            "2": {"font_family": "KaiTi", "size_pt": 14.0, "italic": True,
                  "indent_in": 0.25},
        },
    }
    ui._open_questionnaire()
    ui.root.update_idletasks()
    win = ui._popups[-1]

    # 回显：字体/字号必须按级回填（不填就等于白填）
    combos = [c.get() for c in _descendants(win, ttk.Combobox)]
    entries = [e.get() for e in _descendants(win, tk.Entry)]
    assert "SimHei" in combos and "KaiTi" in combos
    assert "16.0" in entries and "14.0" in entries and "0.25" in entries

    none_txt = i18n.t(ui.lang, "q_none")
    assert none_txt in combos, "三态下拉必须有「不规定」这一档（—）"
    assert i18n.t(ui.lang, "q_align_center") in combos
    assert i18n.t(ui.lang, "q_prefix_number") in combos

    # 层级数 = 2 → 只露两行（第 3 行的"Level 3"标签不该是可见的）
    shown = [w.cget("text") for w in _descendants(win, tk.Label)
             if w.cget("text").startswith(i18n.t(ui.lang, "q_h_lvl", "").rstrip())
             and w.winfo_ismapped()]
    assert len(shown) == 2, shown

    # 保存：明细必须原样落进 questionnaire_data（手填与 AI 同一套字段名）
    win._primary_btn._command()
    hf = (ui.questionnaire_data or {}).get("heading_formats") or {}
    assert set(hf) == {"1", "2"}, hf
    assert hf["1"]["font_family"] == "SimHei"
    assert hf["1"]["size_pt"] == 16.0
    assert hf["1"]["bold"] is True
    assert hf["1"]["alignment"] == "center"
    assert hf["1"]["number_prefix"] == "number"
    assert hf["2"]["indent_in"] == 0.25
    assert hf["2"]["italic"] is True


def test_questionnaire_spec_linkage_prefills_and_keeps_user_values(ui):
    """选参考文献规范 → 空字段按该规范典型值预填；客户填过/保存过的一律不动。

    老板 2026-10-19：客户选中哪个规范，默认值要调整为哪个规范的。
    预填值的单一真相源是引擎 ``target_from_spec``，UI 不抄一份默认值。
    """
    import tkinter as tk
    from tkinter import ttk

    from src.engine.questionnaire import target_from_spec
    from src.ui import i18n

    # 客户上次保存过页边距 1.25 —— 联动绝不能覆盖它
    ui.questionnaire_data = {"margin_top_in": "1.25"}
    ui._open_questionnaire()
    ui.root.update_idletasks()
    win = ui._popups[-1]

    ieee = target_from_spec("IEEE")
    style_combo = [c for c in _descendants(win, ttk.Combobox)
                   if tuple(str(v) for v in c.cget("values")) == ("APA", "MLA", "IEEE")][0]
    style_combo.set("IEEE")
    ui.root.update_idletasks()

    combos = [c.get() for c in _descendants(win, ttk.Combobox)]
    entries = [e.get() for e in _descendants(win, tk.Entry)]
    # 客户保存过的值原样保留
    assert "1.25" in entries, "联动覆盖了客户已保存的页边距"
    # 空字段按 IEEE 典型值预填（单一真相源 = target_from_spec("IEEE")）
    assert str(ieee.page["font_size_pt"]) in entries          # "10.0"
    assert str(ieee.heading_levels) in entries                # "5"
    assert str(ieee.page["font_family"]) in entries           # Times New Roman（正文字体是文本框）
    assert str(ieee.page["body_alignment"]) in combos         # justify（下拉）
    # 栏数联动为 IEEE 双栏（与复核弹窗同一个词）
    one = i18n.t(ui.lang, "review_one_column")
    two = i18n.t(ui.lang, "review_two_column")
    col_combo = [c for c in _descendants(win, ttk.Combobox)
                 if tuple(str(v) for v in c.cget("values")) == (one, two)][0]
    assert col_combo.get() == two, "IEEE 应联动为双栏"
    # 参考文献细项：IEEE 顺序编码 = 是
    yes = i18n.t(ui.lang, "bool_yes")
    numbered = [c for c in _descendants(win, ttk.Combobox)
                if tuple(str(v) for v in c.cget("values")) == (yes, i18n.t(ui.lang, "bool_no"))]
    assert any(c.get() == yes for c in numbered), "IEEE 顺序编码应预填为「是」"
    # 表题位置预填 top、图题 bottom（IEEE/APA 口径；choice 前面带「—」档）
    none_txt = i18n.t(ui.lang, "q_none")
    pos = [c.get() for c in _descendants(win, ttk.Combobox)
           if tuple(str(v) for v in c.cget("values")) == (none_txt, "top", "bottom")]
    assert sorted(pos) == ["bottom", "top"], pos


def test_questionnaire_sections_are_collapsible(ui):
    """三大区块要能折叠：点标题收起（内容不可见），再点展开恢复。

    老板 2026-10-19：这些大的几个方面要能折叠，填完后可以折叠。
    """
    import tkinter as tk

    from src.ui import i18n

    ui._open_questionnaire()
    ui.root.update_idletasks()
    win = ui._popups[-1]

    def mapped_entries():
        return [e for e in _descendants(win, tk.Entry) if e.winfo_ismapped()]

    before = len(mapped_entries())
    assert before > 0, "展开状态下应有可见的输入框"

    page_title = i18n.t(ui.lang, "q_sec_page")
    # 折叠头：标题在左、chevron 在右。用户真实点击的是标题文字 Label（子控件），
    # 之前只绑在 Frame 上——点子 Label 不触发，测试却直接对 Frame 发事件导致假绿。
    # 这里就模拟真实点击子控件，作为回归护栏。
    title_lbls = [w for w in _descendants(win, tk.Label)
                   if w.cget("text").endswith(page_title)]
    assert title_lbls, "找不到可折叠区块标题（Page layout）"
    title_lbl = title_lbls[0]   # 用户真实点的就是标题文字控件

    title_lbl.event_generate("<Button-1>")
    ui.root.update_idletasks()
    after = len(mapped_entries())
    assert after < before, "点击区块标题后没有收起（可见输入框数量未减少）"

    title_lbl.event_generate("<Button-1>")
    ui.root.update_idletasks()
    assert len(mapped_entries()) == before, "再次点击后没有重新展开"

    # 专项：直接点右侧小箭头（mark_lbl）也必须折叠——老板原话就是"小箭头点了没反应"。
    head_frame = title_lbl.master
    chevrons = [w for w in head_frame.winfo_children()
                if isinstance(w, tk.Label) and w.cget("text") in ("▾", "▸")]
    assert chevrons, "找不到折叠用的 chevron 标记"
    chevrons[0].event_generate("<Button-1>")
    ui.root.update_idletasks()
    assert len(mapped_entries()) < before, "点击右侧小箭头后没有收起"


def test_questionnaire_heading_table_feeds_target(ui, monkeypatch):
    """手填的每级明细要真的进 TargetProfile（不能只存在界面里）。"""
    from src.ui import app as app_mod

    monkeypatch.setattr(app_mod.messagebox, "showinfo", lambda *a, **k: None)
    monkeypatch.setattr(app_mod.messagebox, "askyesno", lambda *a, **k: True)

    ui.source_mode.set("src_manual")
    ui.questionnaire_data = {"heading_levels": "1",
                             "heading_formats": {"1": {"font_family": "SimHei",
                                                       "size_pt": 16.0}}}
    t = ui._build_target()
    assert t.heading_formats.get(1, {}).get("font_family") == "SimHei"
    assert t.heading_levels == 1


def test_source_mode_shows_only_its_own_panel(ui):
    """来源 7 选 1、**互斥** → 高级区里只显示被选中那一项对应的面板。

    v2.3.65：老板指出「这七个选项是互斥的，用户只能七选一，只是学校模板那个地方
    可以选参考文献格式」。根因有两层：
      ① `_apply_source_mode` 只负责解锁/折叠整片高级区，从没区分过子面板；
      ② 「手填问卷」与「AI 导入」被塞进同一个 Custom 选项里。
    现在三块面板（学校模板 / 手填问卷 / AI 导入）一一对应三个来源，互斥显示。
    """
    def _vis(mode):
        ui._apply_source_mode(mode, refresh=True)
        ui.root.update_idletasks()
        return (bool(ui._school_frame.winfo_ismapped()),
                bool(ui._manual_frame.winfo_ismapped()),
                bool(ui._ai_frame.winfo_ismapped()))

    assert _vis("src_school") == (True, False, False), "选 School Template 露出了别的面板"
    assert _vis("src_manual") == (False, True, False), "选手填问卷露出了别的面板"
    assert _vis("src_ai") == (False, False, True), "选 AI 导入露出了别的面板"
    # 内置规范：高级区整片折叠，三块都不露
    for mode in ("src_apa_student", "src_apa_pro", "src_mla", "src_ieee"):
        assert _vis(mode) == (False, False, False), "%s 露出了子面板" % mode


def test_export_ai_button_only_in_ai_mode(ui):
    """「Export AI template」只属于 AI 导入这条路（老板：只在 AI 导入时显示）。"""
    def _shown(mode):
        ui._apply_source_mode(mode, refresh=True)
        ui.root.update_idletasks()
        return ui._export_btn.winfo_manager() == "pack"

    assert _shown("src_ai") is True
    assert _shown("src_manual") is False
    assert _shown("src_school") is False
    assert _shown("src_apa_student") is False
    assert _shown("src_ai") is True


def test_source_dropdown_has_seven_exclusive_options(ui):
    """来源下拉必须是 7 项，且每项唯一（互斥）。"""
    import copy
    import io
    import os
    import re

    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    src = io.open(os.path.join(root, "src", "ui", "app.py"), encoding="utf-8").read()
    m = re.search(r"self\._source_map = \[(.*?)\]", src, re.S)
    assert m, "没找到 _source_map"
    keys = re.findall(r'"(\w+)"', m.group(1))
    src_keys = [k for k in keys if k.startswith("src_")]
    assert len(src_keys) == 7, "来源下拉应为 7 项，实际 %d 项：%r" % (len(src_keys), src_keys)
    assert len(set(src_keys)) == 7, "来源下拉有重复项：%r" % (src_keys,)
    assert "src_manual" in src_keys and "src_ai" in src_keys

    shown = [ui.tr(k) for k in src_keys]
    assert len(set(shown)) == 7, "两项显示同一个文案（客户分不清）：%r" % (shown,)


def test_build_target_only_uses_inputs_of_current_source(ui):
    """目标画像只能取**当前来源**的输入，别的通道不许在背后生效。

    两个隐形生效的坑：
      ① 选 Custom 时，之前点过的学校模板还在参与合并（界面看不见、实际生效）；
      ② 选内置规范时，之前填过的问卷会以"高优先级叠加"盖掉规范本身。
    """
    ui.school_template.set(r"D:/no_such_template.docx")   # 故意无效：Custom 模式下不该被读
    ui.questionnaire_data = {"margin_top_in": "1.75", "font_family": "Arial"}

    # 真实交互里 source_mode 与 _apply_source_mode 是同一动作（trace 兜底），
    # 这里两步都做：只调 _apply_source_mode 的话 _build_target 读到的还是旧来源。
    ui.source_mode.set("src_manual")
    ui._apply_source_mode("src_manual", refresh=True)
    t = ui._build_target()
    assert "school" not in t.source and "spec" not in t.source, \
        "Custom 模式还带着别的来源：source=%r" % (t.source,)
    assert t.page.get("margin_top_in") == 1.75

    ui.source_mode.set("src_apa_student")
    ui._apply_source_mode("src_apa_student", refresh=True)
    t2 = ui._build_target()
    assert t2.spec_key == "APA-Student"
    assert t2.source == "spec", "规范模式被问卷叠加上去了：source=%r" % (t2.source,)
    assert t2.page.get("margin_top_in") == 1.0, "规范模式被问卷改了页边距"
