"""进程 DPI 感知锁定 + 原生对话框之后的显示状态自愈。

背景（老板 2026-10-21 报，已定性）：
    软件「保存修正后的论文」时会弹 Windows 原生「另存为」对话框。保存完成后，
    软件窗口突然**变小一圈、文字变清楚**，并且**一直保持到关掉软件**。

根因（取证结论）：
    ① 打包产物是 **DPI 不感知**的 —— 我读过本地 Nuitka 4.2.1 的默认 manifest 模板
       （``nuitka/utils/WindowsResources.py::getDefaultWindowsExecutableManifest``），
       全文只有 ``requestedExecutionLevel`` 与 ``longPathAware``，**没有 ``<dpiAware>``**。
    ② DPI 不感知的窗口在 150% 缩放的机器上由 Windows（DWM）整窗位图拉伸显示：
       窗口 1126×640 的逻辑尺寸被放大到约 1689×960 物理像素，字看起来有点糊。
    ③ Windows 原生 shell 对话框（另存为 / 消息框）会把调用方的 DPI 感知提升为
       「感知」模式。提升之后 Windows **停止位图拉伸**，窗口改按物理像素呈现 →
       在 150% 屏上瞬间缩到约 2/3、且因为不再插值而变清晰。
    ④ DPI 感知是**进程级一次性设定**，提升了就退不回去 —— 所以只有重启软件才恢复。
       （老板原话："一直保持到关软件"、"变小了"、"变清楚了"，与③④完全对应。）

两手应对：
    ① :func:`pin_process_dpi_awareness` —— 在创建任何窗口**之前**，把进程 DPI 感知
       **显式**锁成 UNAWARE。显式设置之后同进程再改会失败（Windows 只允许设一次），
       于是原生对话框也无法再把我们提升；显示方式与现状**完全一致**（Windows 继续
       拉伸），只是再也跳不了。这是主治手段。
    ② :func:`snapshot` / :func:`restore_if_drifted` —— 万一仍有东西把缩放改了
       （别的 Windows 版本行为不同），原生对话框返回后立刻把窗口的**物理**尺寸与
       tk 缩放恢复成对话框弹出前的样子。这是保险，只在**实测到变化**时才动手。

不改全局布局：本模块不碰任何控件尺寸与字体磅值，因此不存在「一改就崩版」的风险。
"""

from __future__ import annotations

import sys

_IS_WINDOWS = sys.platform.startswith("win")

# DPI_AWARENESS_CONTEXT_* 的句柄值（winuser.h）
DPI_AWARENESS_CONTEXT_UNAWARE = -1
DPI_AWARENESS_CONTEXT_SYSTEM_AWARE = -2
DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE = -3
DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
DPI_AWARENESS_CONTEXT_UNAWARE_GDISCALED = -5

_AWARE_NAMES = {
    0: "UNAWARE",
    1: "SYSTEM_AWARE",
    2: "PER_MONITOR_AWARE",
}

# 本模块最近一次锁定结果（供调试 / About 展示，客户看不到）
LAST_PIN = "N/A"


def _user32():
    import ctypes

    u = ctypes.windll.user32
    # 必须声明签名：DPI_AWARENESS_CONTEXT 是句柄（指针宽度），不声明会被截断成 int32，
    # 64 位下 -1 会被当成 0xFFFFFFFF 传进去 → 调用失败且没有任何提示。
    u.GetThreadDpiAwarenessContext.restype = ctypes.c_void_p
    u.GetAwarenessFromDpiAwarenessContext.argtypes = [ctypes.c_void_p]
    u.GetAwarenessFromDpiAwarenessContext.restype = ctypes.c_int
    u.SetProcessDpiAwarenessContext.argtypes = [ctypes.c_void_p]
    u.SetProcessDpiAwarenessContext.restype = ctypes.c_bool
    return u


def awareness_code() -> int:
    """当前进程的 DPI 感知模式（0=UNAWARE / 1=SYSTEM_AWARE / 2=PER_MONITOR_AWARE）。"""
    if not _IS_WINDOWS:
        return -1
    try:
        import ctypes

        u = _user32()
        ctx = u.GetThreadDpiAwarenessContext()
        return int(u.GetAwarenessFromDpiAwarenessContext(ctypes.c_void_p(ctx)))
    except Exception:
        return -1


def awareness_name() -> str:
    return _AWARE_NAMES.get(awareness_code(), "N/A")


def pin_process_dpi_awareness() -> str:
    """把进程 DPI 感知锁成 UNAWARE（幂等，**必须在创建任何窗口之前调用**）。

    返回实际生效的模式名。任何异常都吞掉：这一步失败最多是"跟旧版一样会跳"，
    绝不能因为锁不住就不让人用软件。
    """
    global LAST_PIN
    if not _IS_WINDOWS:
        LAST_PIN = "N/A"
        return LAST_PIN
    try:
        import ctypes

        u = _user32()
        # 显式设为 UNAWARE：与当前实际表现一致（Nuitka manifest 没声明 → 默认 unaware），
        # 但"显式"之后就锁死了，后续（含 shell 对话框）再设会失败。
        u.SetProcessDpiAwarenessContext(ctypes.c_void_p(DPI_AWARENESS_CONTEXT_UNAWARE))
    except Exception:
        # Windows 8.1 及更早：走 shcore（PROCESS_DPI_UNAWARE = 0）
        try:
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(0)
        except Exception:
            pass
    LAST_PIN = awareness_name()
    return LAST_PIN


def snapshot(root) -> dict:
    """记录窗口当前的显示状态（原生对话框弹出**之前**抓一次）。"""
    snap = {"geo": None, "scaling": None, "fpixels": None, "awareness": None}
    try:
        snap["geo"] = (int(root.winfo_width()), int(root.winfo_height()),
                       int(root.winfo_x()), int(root.winfo_y()))
    except Exception:
        pass
    try:
        snap["scaling"] = float(root.tk.call("tk", "scaling"))
    except Exception:
        pass
    try:
        snap["fpixels"] = float(root.winfo_fpixels("1i"))
    except Exception:
        pass
    snap["awareness"] = awareness_code()
    return snap


def restore_if_drifted(root, snap) -> list:
    """原生对话框返回后调用：如果显示状态被动过，恢复成 ``snap`` 的样子。

    只在**实测到变化**时才动手，正常路径下是空操作（返回空列表）。
    返回实际做过的动作名列表，便于测试与排查。
    """
    acted = []
    if not snap or not _IS_WINDOWS:
        return acted
    try:
        now_fp = float(root.winfo_fpixels("1i"))
    except Exception:
        now_fp = None
    try:
        now_sc = float(root.tk.call("tk", "scaling"))
    except Exception:
        now_sc = None

    old_fp = snap.get("fpixels")
    old_sc = snap.get("scaling")
    geo = snap.get("geo")

    # ① DPI 比例变了（进程感知被提升 / 窗口被系统按新 DPI 重排）：
    #    物理尺寸 = 逻辑尺寸 × (当前 DPI / 96)。要"看起来跟原来一样大"，
    #    逻辑尺寸就得按比例放大，让物理尺寸回到原值。
    #    字体是 point 制，会随新 DPI 一起变大，所以比例保持一致、不会错版。
    if old_fp and now_fp and abs(now_fp - old_fp) > 0.5 and geo:
        ratio = now_fp / old_fp
        try:
            sw = int(root.winfo_screenwidth())
            sh = int(root.winfo_screenheight())
            nw = min(int(round(geo[0] * ratio)), max(320, sw - 16))
            nh = min(int(round(geo[1] * ratio)), max(240, sh - 64))
            nx = min(max(0, geo[2]), max(0, sw - nw))
            ny = min(max(0, geo[3]), max(0, sh - nh - 40))
            root.geometry("%dx%d+%d+%d" % (nw, nh, nx, ny))
            root.update_idletasks()
            acted.append("geometry x%.3f" % ratio)
        except Exception:
            pass
    else:
        # ② 没有 DPI 变化，但窗口框体被挪动/缩放过（例如系统把 WM_DPICHANGED 的
        #    建议矩形套到了我们身上）→ 按原逻辑尺寸摆回去。
        if geo:
            try:
                cur = (int(root.winfo_width()), int(root.winfo_height()),
                       int(root.winfo_x()), int(root.winfo_y()))
                if any(abs(a - b) > 1 for a, b in zip(cur, geo)):
                    root.geometry("%dx%d+%d+%d" % geo)
                    root.update_idletasks()
                    acted.append("geometry")
            except Exception:
                pass
        # ③ 字体缩放被改过（Tk 的 tk scaling）→ 恢复，否则全部文字会一起变大变小。
        if old_sc and now_sc and abs(now_sc - old_sc) > 1e-6:
            try:
                root.tk.call("tk", "scaling", old_sc)
                acted.append("tk scaling")
            except Exception:
                pass
    return acted
