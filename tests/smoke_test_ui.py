#!/usr/bin/env python3
"""冒烟测试：启动海外版 GUI 并做基本交互，10 秒后自动退出。"""

import sys
import time
import tkinter as tk
import traceback

# 添加 src 到路径
sys.path.insert(0, "D:/AgentSpace/thesis-format-doctor-global/src")

from ui import app as app_module


def smoke_test():
    """启动 GUI，等待 10 秒后自动退出。"""
    print("=" * 60)
    print("海外版 UI 冒烟测试")
    print("=" * 60)
    print()
    print("预期：")
    print("  1. 窗口正常弹出（标题：Thesis Format Doctor）")
    print("  2. 背景正确（#F6F3EC 象牙纸）")
    print("  3. 两栏卡片显示（Your Documents / Review & Fix）")
    print("  4. 顶栏显示品牌和导航（语言·About·Help）")
    print("  5. 页脚显示联系信息")
    print("  6. 状态栏显示就绪状态")
    print()
    print("测试将运行 10 秒，请手动检查上述要点...")
    print()

    # 启动应用
    root = tk.Tk()
    root.withdraw()  # 先隐藏根窗口

    try:
        app = app_module.App(root, lang="en")
    except Exception as e:
        print(f"❌ 启动失败：{e}")
        traceback.print_exc()
        return

    # 10 秒后自动退出
    root.after(10000, root.destroy)

    # 启动主循环
    root.mainloop()

    print()
    print("✓ 冒烟测试完成")


if __name__ == "__main__":
    smoke_test()