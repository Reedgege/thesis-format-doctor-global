"""PaperFormat Pro — offline, privacy-first desktop edition."""

# 版本号单一来源：永远经 versioninfo.app_version() 从仓库根 VERSION 文件读取，
# 不再在此硬编码，避免发版漏改导致展示版本与产物版本不一致（F3）。
from .versioninfo import app_version as _app_version

__version__ = _app_version()
