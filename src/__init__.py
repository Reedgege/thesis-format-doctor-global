"""PaperFormat Pro — offline, privacy-first desktop edition."""

# 🔴 必须在任何引擎操作之前执行两件事（见 docx_boot 模块注释）：
#   1. 让 python-docx 完成导入；
#   2. 把它各模块的相对 __file__ 修正成绝对路径 —— 否则 macOS / Linux 打包版
#      读不到自带的模板文件（页眉页脚 / 批注 / 样式），建部件时静默失败。
# 源码运行时 __file__ 本来就是绝对路径，这一步什么都不做。
try:
    import docx as _docx  # noqa: F401
except Exception:
    pass

from . import docx_boot as _docx_boot

_docx_boot.apply()

# 版本号单一来源：永远经 versioninfo.app_version() 从仓库根 VERSION 文件读取，
# 不再在此硬编码，避免发版漏改导致展示版本与产物版本不一致（F3）。
from .versioninfo import app_version as _app_version

__version__ = _app_version()
