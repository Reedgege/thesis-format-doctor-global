#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""发版打包助手：CI 产物 → 改名 → 打包安装版 → 算体积/SHA256 →（可选）传 R2。

为什么需要这个（2026-10-25 立项）：
    每次发版都要手工做同一串事——给 4 个产物改带版本号的名字、把安装版 exe 再压成
    zip、算体积与 SHA256、把 4 个 zip 传 R2 的 downloads/<ver>/、再把体积与哈希
    抄进站点 download.html。抄错一个数字客户就会校验失败。

用法
----
    # 1) CI 产物下载到本地某目录后：
    python tools/package_release.py --artifacts <目录> --version 2.3.108
    #    → 生成改名后的 4 个文件 + 打印体积/SHA256 表格 + 打印可直接粘进
    #      download.html 的 <tr> 行

    # 2) 顺带传 R2（需 CLOUDFLARE_API_TOKEN 环境变量）：
    python tools/package_release.py --artifacts <目录> --version 2.3.108 --upload-r2
    #    → 用 wrangler r2 object put 传到 paperformatpro 桶的 downloads/<ver>/

    # 3) 校验站点上的 SHA256 表与本地产物是否一致：
    python tools/package_release.py --artifacts <目录> --version 2.3.108 --check-site <站点目录>

只用标准库；调用 wrangler 走 subprocess。
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile

# CI 产出的原始名 → 发到 R2 的名字（带 v 版本号）
MAPPING = {
    "paperformatpro-windows.zip": "paperformatpro-windows-v{ver}.zip",
    "paperformatpro-macos.zip":   "paperformatpro-macos-v{ver}.zip",
    "paperformatpro-linux.zip":   "paperformatpro-linux-v{ver}.zip",
    "paperformatpro-setup.exe":   "paperformatpro-setup-v{ver}.exe",
}
# 安装版另外压一个 zip（站点下载页给的是 zip，不是裸 exe）
SETUP_ZIP = "paperformatpro-setup-v{ver}.zip"
BUCKET = "paperformatpro"
R2_PREFIX = "downloads"


def sha256_of(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


# 认产物：原始名（paperformatpro-windows.zip）或已带版本号的改名件都认
_KIND_RE = {
    "paperformatpro-windows.zip": re.compile(r"^paperformatpro-windows(-v\d+\.\d+\.\d+)?\.zip$"),
    "paperformatpro-macos.zip":   re.compile(r"^paperformatpro-macos(-v\d+\.\d+\.\d+)?\.zip$"),
    "paperformatpro-linux.zip":   re.compile(r"^paperformatpro-linux(-v\d+\.\d+\.\d+)?\.zip$"),
    "paperformatpro-setup.exe":   re.compile(r"^paperformatpro-setup(-v\d+\.\d+\.\d+)?\.exe$"),
}


def find_artifacts(src: str) -> dict:
    """在目录里找出 4 个 CI 产物。

    🔴 保险：**优先取原始名**（paperformatpro-windows.zip）。
    目录里若同时躺着上一版改过名的件（paperformatpro-windows-v2.3.104.zip），
    绝不能误取旧的——发出去的包版本对不上，客户装的是旧版。
    原始名不存在时才退而接受带版本号的件。
    """
    files = [fn for fn in os.listdir(src)
             if os.path.isfile(os.path.join(src, fn))]
    found = {}
    for orig, pat in _KIND_RE.items():
        exact = [fn for fn in files if fn == orig]
        if exact:
            found[orig] = os.path.join(src, exact[0])
            continue
        variants = [fn for fn in files if pat.match(fn)]
        if len(variants) == 1:
            found[orig] = os.path.join(src, variants[0])
        elif len(variants) > 1:
            raise SystemExit(
                f"目录里有多个 {orig} 的候选：{variants}\n"
                f"为避免发错版本，请只保留本次构建的那一个。")
    return found


def package(src: str, out: str, ver: str) -> dict:
    found = find_artifacts(src)
    missing = [k for k in MAPPING if k not in found]
    if missing:
        raise SystemExit(f"缺少 CI 产物：{missing}\n目录里现有：{os.listdir(src)}")

    os.makedirs(out, exist_ok=True)
    result = {}
    for orig, tmpl in MAPPING.items():
        name = tmpl.format(ver=ver)
        dst = os.path.join(out, name)
        if os.path.abspath(found[orig]) != os.path.abspath(dst):
            shutil.copy2(found[orig], dst)
        result[name] = dst

    # 安装版压 zip（站点给的是 zip）
    setup_zip = os.path.join(out, SETUP_ZIP.format(ver=ver))
    if not os.path.exists(setup_zip):
        with zipfile.ZipFile(setup_zip, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(result[SETUP_ZIP.format(ver=ver).replace(".zip", ".exe")],
                    SETUP_ZIP.format(ver=ver).replace(".zip", ".exe"))
    result[SETUP_ZIP.format(ver=ver)] = setup_zip
    return result


def describe(files: dict) -> list:
    rows = []
    for name, path in sorted(files.items()):
        rows.append({
            "name": name,
            "size": os.path.getsize(path),
            "size_mb": round(os.path.getsize(path) / 1048576, 1),
            "sha256": sha256_of(path),
            "path": path,
        })
    return rows


def print_table(rows: list) -> None:
    print("\n=== 体积与 SHA256 ===")
    for r in rows:
        print(f"  {r['name']}")
        print(f"     {r['size']} B ({r['size_mb']} MB)   sha256={r['sha256']}")


def print_html_rows(rows: list) -> None:
    print("\n=== 可直接粘进 download.html 的 SHA256 表行 ===")
    for r in rows:
        if "setup" in r["name"] and r["name"].endswith(".zip"):
            label = f"{r['name']} (installer)"
        else:
            label = r["name"]
        print(f'        <tr><td style="padding:4px 8px">{label}</td>'
              f'<td style="padding:4px 8px;word-break:break-all;'
              f'font-family:monospace">{r["sha256"]}</td></tr>')


def upload_r2(rows: list, ver: str) -> None:
    wr = shutil.which("wrangler")
    if not wr:
        for cand in (r"C:\Users\yunwu\.workbuddy\binaries\node\versions\22.22.2-6\wrangler",):
            if os.path.exists(cand):
                wr = cand
                break
    if not wr:
        raise SystemExit("找不到 wrangler，请先装/或加进 PATH")
    print(f"\n=== 传 R2 → {BUCKET}/{R2_PREFIX}/{ver}/ ===")
    for r in rows:
        if r["name"].endswith(".exe"):
            continue  # 裸 exe 不对外发，只发 zip（与上次发版一致）
        key = f"{BUCKET}/{R2_PREFIX}/{ver}/{r['name']}"
        print(f"  put {key}")
        p = subprocess.run([wr, "r2", "object", "put", key, "--file", r["path"]],
                           capture_output=True, text=True)
        if p.returncode != 0:
            print(f"    ❌ 失败：{(p.stderr or p.stdout)[:300]}")
        else:
            print(f"    ✅ 完成（{(p.stdout or '').strip()[:120]}）")


def check_site(site_dir: str, rows: list) -> int:
    """站点 download.html 里的 SHA256 表 vs 本地产物。"""
    path = os.path.join(site_dir, "download.html")
    if not os.path.exists(path):
        print(f"找不到 {path}")
        return 1
    body = io.open(path, encoding="utf-8").read()
    bad = 0
    for r in rows:
        m = re.search(re.escape(r["name"]) + r"[^<]*</td><td[^>]*>([0-9a-f]{64})<", body)
        if not m:
            print(f"  ⚠️ 站点表里没有 {r['name']}")
            continue
        if m.group(1) != r["sha256"]:
            print(f"  ❌ {r['name']} 哈希不符：站点 {m.group(1)[:12]}… vs 实际 {r['sha256'][:12]}…")
            bad += 1
        else:
            print(f"  ✅ {r['name']} 哈希相符")
    return 1 if bad else 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="发版打包助手")
    ap.add_argument("--artifacts", required=True, help="CI 产物所在目录")
    ap.add_argument("--version", required=True, help="版本号，如 2.3.108")
    ap.add_argument("--out", help="改名后文件的输出目录（默认 <artifacts>/_staged_<ver>）")
    ap.add_argument("--upload-r2", action="store_true", help="传 R2（需 CLOUDFLARE_API_TOKEN）")
    ap.add_argument("--check-site", help="顺带校验站点 download.html 的 SHA256 表")
    args = ap.parse_args(argv)

    out = args.out or os.path.join(args.artifacts, f"_staged_{args.version}")
    files = package(args.artifacts, out, args.version)
    rows = describe(files)
    print_table(rows)
    print_html_rows(rows)
    print(f"\n输出目录：{out}")

    if args.upload_r2:
        upload_r2(rows, args.version)
    if args.check_site:
        return check_site(args.check_site, rows)
    return 0


if __name__ == "__main__":
    sys.exit(main())
