"""Thesis Format Doctor Global - packaging entry point.

Nuitka compiles this file into the executable stub; the real logic lives in
`src.main`. Keeping a thin root launcher lets `src` stay a normal package so
its intra-package relative imports (`from .engine import ...`) keep working
after compilation.

Two things are critical for the shipped Windows build:

* **Double-clicking must open the window.** The exe is built with the console
  subsystem disabled, so a "print help and exit" path is completely invisible
  to the user - the program looks like it simply does not start. Therefore:
  when launched with no arguments (double-click, Start-menu / desktop
  shortcut) we default to the `gui` subcommand. Explicit subcommands
  (`check` / `fix` / `activate` / `status` / `export-schema`) keep the plain
  CLI behaviour untouched.
* **Never die silently.** Any startup failure is surfaced through a native
  message box on Windows (stderr elsewhere), so a problem is reportable
  instead of looking like "nothing happened".
"""

from __future__ import annotations

import sys


def _fatal(msg: str) -> None:
    """Report a fatal startup error through whatever channel is available."""
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(0, msg, "Thesis Format Doctor Global", 0x10)
        return
    except Exception:
        pass
    try:
        print(msg, file=sys.stderr)
    except Exception:
        pass


def main(argv=None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if not args:
        # 双击 / 快捷方式启动：没有参数 -> 直接进 GUI。
        args = ["gui"]
        sys.argv = [sys.argv[0]] + args

    try:
        from src.main import main as cli_main
    except Exception as exc:  # pragma: no cover - only reachable in a broken build
        _fatal("Failed to load application modules.\n\n%s: %s"
               % (type(exc).__name__, exc))
        return 1

    try:
        return cli_main(args)
    except SystemExit:
        raise
    except Exception as exc:  # pragma: no cover - last-resort safety net
        _fatal("Unexpected error while running.\n\n%s: %s"
               % (type(exc).__name__, exc))
        return 1


if __name__ == "__main__":
    sys.exit(main())
