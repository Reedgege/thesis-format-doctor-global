# -*- coding: utf-8 -*-
"""Build the offline code generator (tools/reedcode_gui.py) into a standalone exe via Nuitka.

This is the *seller-side* tool (对应国内版 reedcode.exe): double-click, paste the
customer's machine code, click generate, send the offline code back. It is a pure
tkinter + stdlib script (no src/ dependency), so it compiles to a small standalone
folder. On Windows we then zip it for distribution.

Mirrors build.py's low-false-positive strategy (standalone folder, not onefile
self-extracting dropper; tk-inter plugin bundles Tcl/Tk).

Usage:
  python build_reedcode.py            # build for the current platform into dist/
"""
import os
import sys
import shutil
import subprocess

# Windows CI console is cp1252 by default; force utf-8 so any print never exits non-zero.
try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
WIN = sys.platform.startswith("win")

APP_NAME = "reedcode-global"                     # ASCII on purpose (cp1252 CI safety)
ENTRY = os.path.join(HERE, "tools", "reedcode_gui.py")
ICON_ICO = os.path.join(HERE, "assets", "icon.ico")
ASSETS_DIR = os.path.join(HERE, "assets")

COMPANY_NAME = "ReedSkill"
PRODUCT_NAME = "Thesis Format Doctor Global - Offline Code Generator"


def _read_version():
    try:
        with open(os.path.join(HERE, "VERSION"), encoding="utf-8") as fh:
            return fh.read().strip().lstrip("vV")
    except Exception:
        return ""


def _nuitka_options():
    opt = [
        sys.executable, "-m", "nuitka", ENTRY,
        "--standalone",                 # folder mode: no self-extract -> no dropper fingerprint
        "--enable-plugins=tk-inter",    # tkinter support (Tcl/Tk bundled in)
        "--assume-yes-for-downloads",
        "--output-dir=" + os.path.join(HERE, "dist"),
        "--output-filename=" + APP_NAME + (".exe" if WIN else ""),
        "--remove-output",
        "--show-progress",
        "--include-package=rsa",
    ]
    if WIN and os.path.isfile(ICON_ICO):
        opt += ["--windows-icon-from-ico=" + ICON_ICO]
    if WIN:
        # GUI tool: no black console box on launch.
        opt += ["--windows-console-mode=disable"]
        _ver = _read_version()
        if _ver:
            opt += ["--company-name=" + COMPANY_NAME,
                    "--product-name=" + PRODUCT_NAME,
                    "--product-version=" + _ver,
                    "--file-version=" + _ver,
                    "--file-description=" + "Offline activation code generator (seller tool)",
                    "--copyright=Copyright (c) " + COMPANY_NAME]
    return opt


def _normalize_and_bundle():
    """Rename Nuitka's <entry>.dist to APP_NAME, then copy assets/ next to the exe
    so the GUI can load assets/icon.ico at runtime (HERE == exe directory)."""
    dist = os.path.join(HERE, "dist")
    src = os.path.join(dist, "reedcode_gui.dist")
    dst = os.path.join(dist, APP_NAME)
    if os.path.isdir(src) and not os.path.exists(dst):
        shutil.move(src, dst)
    if not os.path.isdir(dst):
        return src
    # bundle assets (icon.ico) next to the exe
    target_assets = os.path.join(dst, "assets")
    if os.path.isdir(ASSETS_DIR):
        os.makedirs(target_assets, exist_ok=True)
        for f in os.listdir(ASSETS_DIR):
            shutil.copy2(os.path.join(ASSETS_DIR, f), os.path.join(target_assets, f))
    return dst


def build():
    cmd = _nuitka_options()
    print(">>> " + " ".join(cmd))
    rc = subprocess.call(cmd, cwd=HERE)
    if rc != 0:
        print("Build failed with return code", rc)
        sys.exit(rc)
    out = _normalize_and_bundle()
    print("\nBuild complete -> " + out)
    print("executable -> " + os.path.join(out, APP_NAME + (".exe" if WIN else "")))


if __name__ == "__main__":
    build()
