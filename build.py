# -*- coding: utf-8 -*-
"""Cross-platform build script (Nuitka) - Thesis Format Doctor Global.

Compiles the overseas edition into a native "folder" build; on Windows we then
wrap it with NSIS into a single installer exe. This avoids the PyInstaller
onefile self-extracting-dropper fingerprint that trips antivirus heuristics.

  Windows -> dist/thesis-format-doctor-global/  + installer.nsi -> thesis-format-doctor-global-setup.exe
  macOS   -> dist/thesis-format-doctor-global.app
  Linux   -> dist/thesis-format-doctor-global/

Nuitka turns Python into C -> native machine code. The product in Windows
Defender's eyes is just a normal C++ program (very low false-positive rate),
and the source is compiled away. Customers need no Python installed - Nuitka
standalone bundles the Python runtime + Tcl/Tk into the folder.

Usage:
  python build.py            # build for the current platform into dist/
  python build.py --clean   # clean dist/ first, then build
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
DARWIN = sys.platform == "darwin"

APP_NAME = "thesis-format-doctor-global"
# All ASCII on purpose: the Windows CI runner console is cp1252, and non-ASCII
# file/dir names blow up in Compress-Archive / NSIS / log prints. English product
# name is also the natural choice for an overseas audience.
ENTRY = os.path.join(HERE, "main.py")  # root launcher -> imports src.main internally
ASSETS_DIR = os.path.join(HERE, "assets")
ICON_ICO = os.path.join(HERE, "assets", "icon.ico")
ICON_ICNS = os.path.join(HERE, "assets", "icon.icns")
ICON_PNG = os.path.join(HERE, "assets", "icon.png")

PRODUCT_NAME = "Thesis Format Doctor Global"
COMPANY_NAME = "ReedSkill"


def _read_version():
    """Read VERSION, strip a leading 'v', return the bare numeric version (e.g. 2.0.0)."""
    try:
        with open(os.path.join(HERE, "VERSION"), encoding="utf-8") as fh:
            return fh.read().strip().lstrip("vV")
    except Exception:
        return ""


def _nuitka_options():
    opt = [
        sys.executable, "-m", "nuitka", ENTRY,
        "--standalone",                 # folder mode: no self-extract at runtime -> no dropper fingerprint
        "--enable-plugins=tk-inter",    # tkinter support (Tcl/Tk bundled in)
        "--assume-yes-for-downloads",   # auto-agree when Nuitka needs to fetch ccache/patches
        "--output-dir=" + os.path.join(HERE, "dist"),
        "--output-filename=" + APP_NAME + (".exe" if WIN else ""),
        "--remove-output",              # drop the .build workdir after use; keep only .dist/.app
        "--show-progress",
    ]
    # The whole src package (engine/license/ui + data/*.json) is compiled into the build.
    # src/data/ also carries icon.png, so the GUI can set its window icon at runtime.
    opt += ["--include-package=src",
            "--include-package-data=src"]
    # Belt and braces: drop assets/ next to the executable as well. The GUI looks for
    # <exe dir>/assets/icon.png, which keeps the window icon working even if the
    # package-data layout differs between Nuitka versions.
    if os.path.isdir(ASSETS_DIR):
        opt += ["--include-data-dir=" + ASSETS_DIR + "=assets"]
    # VERSION travels with the build: the About dialog and the status bar read it at
    # runtime (in the source tree app.py walks up to the repo root; once compiled that
    # path no longer exists, so the file itself must ship next to the executable).
    version_file = os.path.join(HERE, "VERSION")
    if os.path.isfile(version_file):
        opt += ["--include-data-file=" + version_file + "=VERSION"]
    if WIN and os.path.isfile(ICON_ICO):
        opt += ["--windows-icon-from-ico=" + ICON_ICO]
    # Windows GUI app: disable the console subsystem so no black box flashes on launch.
    # Also inject version resources (file properties -> company/product/version) to make
    # the unsigned exe look more legitimate and lower heuristic scores.
    if WIN:
        opt += ["--windows-console-mode=disable"]
        _ver = _read_version()
        if _ver:
            opt += ["--company-name=" + COMPANY_NAME,
                    "--product-name=" + PRODUCT_NAME,
                    "--product-version=" + _ver,
                    "--file-description=" + PRODUCT_NAME + " (offline thesis format checker & fixer)",
                    "--file-version=" + _ver,
                    "--copyright=Copyright (c) " + COMPANY_NAME]
    if DARWIN:
        if os.path.isfile(ICON_ICNS):
            opt += ["--macos-app-icon=" + ICON_ICNS]
        opt += ["--macos-create-app-bundle",
                "--macos-app-mode=gui",
                "--macos-app-name=" + PRODUCT_NAME]
    return opt


def _normalize_output():
    """Nuitka's standalone dir is always <entry-basename>.dist / .app; rename to the
    publish dir name. macOS/Linux keep the ASCII APP_NAME."""
    dist = os.path.join(HERE, "dist")
    if DARWIN:
        src, dst = os.path.join(dist, "main.app"), os.path.join(dist, APP_NAME + ".app")
    else:
        src, dst = os.path.join(dist, "main.dist"), os.path.join(dist, APP_NAME)
    if os.path.isdir(src) and not os.path.exists(dst):
        shutil.move(src, dst)
        return dst
    if os.path.isdir(dst):
        return dst
    return src


def _report():
    out = _normalize_output()
    print("\nBuild complete -> " + out)
    if WIN:
        print("executable -> " + os.path.join(out, APP_NAME + ".exe"))
        print("installer  -> run `makensis installer.nsi` at repo root to produce setup.exe")
    elif DARWIN:
        print("app bundle -> " + out)
    else:
        print("executable -> " + os.path.join(out, APP_NAME))


def build(clean=False):
    if clean:
        d = os.path.join(HERE, "dist")
        if os.path.isdir(d):
            shutil.rmtree(d)
    cmd = _nuitka_options()
    print(">>> " + " ".join(cmd))
    rc = subprocess.call(cmd, cwd=HERE)
    if rc != 0:
        print("Build failed with return code", rc)
        sys.exit(rc)
    _report()


if __name__ == "__main__":
    build(clean="--clean" in sys.argv)
