#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""下载 Release 产物 → 解出可执行文件 → 交给 verify_cli 做功能等价性验证。

为什么单独写这个：Windows 用 PowerShell、macOS 用 bash 解压找 .app，
两条命令写法完全不同，容易一边对一边错。这个脚本用标准库统一掉，
workflow 里三平台就一行命令，减少"测试脚本自己出错"的面积
（上一轮 macOS 冒烟就吃过这个亏：$GITHUB_ENV 传相对路径 → 文件找不到）。

用法：
  python3 tools/crossplat/fetch_and_verify.py \
      --repo Reedgege/thesis-format-doctor-global \
      --label "Windows x64" \
      [--asset paperformatpro-windows.zip] [--tag v2.3.122] [--workdir /tmp/xp]
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.error
import urllib.request
import zipfile

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import verify_cli  # noqa: E402


def _esc(s: str) -> str:
    return s.replace("%", "%25").replace("\r", "").replace("\n", "%0A")


def fail(msg: str) -> None:
    print("::error::" + _esc(msg), flush=True)


def notice(msg: str) -> None:
    print("::notice::" + _esc(msg), flush=True)


def api_get(url: str, token: str) -> dict:
    req = urllib.request.Request(url, headers={
        "Accept": "application/vnd.github+json",
        "User-Agent": "crossplat-verify",
    })
    if token:
        req.add_header("Authorization", "Bearer " + token)
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read().decode("utf-8"))


def download(url: str, dst: str, token: str) -> None:
    """下载 Release 资产。

    优先用 curl：GitHub 的下载链接会 302 到 objects.githubusercontent.com，
    urllib 在本机实测出现过 "Remote end closed connection without response"，
    而 curl -L 稳定。三平台 runner 都自带 curl，不引额外依赖。
    """
    import shutil
    import subprocess

    curl = shutil.which("curl")
    if curl:
        cmd = [curl, "-fL", "--retry", "3", "--retry-delay", "2"]
        if token:
            cmd += ["-H", "Authorization: Bearer " + token]
        cmd += ["-o", dst, url]
        p = subprocess.run(cmd, capture_output=True)
        if p.returncode == 0 and os.path.isfile(dst) and os.path.getsize(dst) > 0:
            return
        tail = (p.stderr or b"").decode("utf-8", "replace")[-400:]
        print("curl 下载失败（rc=%s）：%s" % (p.returncode, tail.strip()))

    req = urllib.request.Request(url, headers={"User-Agent": "crossplat-verify"})
    if token:
        req.add_header("Authorization", "Bearer " + token)
    with urllib.request.urlopen(req, timeout=300) as r, open(dst, "wb") as f:
        while True:
            chunk = r.read(1024 * 256)
            if not chunk:
                break
            f.write(chunk)


def pick_asset(assets: list, want: str) -> dict | None:
    """挑资产。want 支持通配（Linux 的包名带版本号，Windows/macOS 不带）。"""
    import fnmatch

    names = [a.get("name") or "" for a in assets]
    if want:
        for a in assets:
            if a.get("name") == want:
                return a
        hits = [a for a in assets if fnmatch.fnmatch(a.get("name") or "", want)]
        if not hits:
            print("可用资产：%s" % ", ".join(names))
            return None
        # 同一模式命中多个时取最短名字（避开 setup / 便携版之类的变体）
        hits.sort(key=lambda a: len(a.get("name") or ""))
        return hits[0]
    for a in assets:
        if (a.get("name") or "").lower().endswith(".zip"):
            return a
    print("可用资产：%s" % ", ".join(names))
    return None


def find_executable(root: str, prefer: str | None) -> str | None:
    """在解压目录里找主程序。

    Windows：优先 prefer 指定的名字，其次任意 *.exe（排除 uninstall 之类）。
    macOS：沿 *.app/Contents/MacOS/ 找第一个可执行位文件。
    Linux：优先同名的无扩展名文件。
    """
    cands = []
    for dirpath, dirnames, filenames in os.walk(root):
        # 跳过 macOS 的代码签名目录，里面全是无关文件
        if "_CodeSignature" in dirpath:
            continue
        if dirpath.replace("\\", "/").endswith(".app/Contents/MacOS"):
            for fn in filenames:
                p = os.path.join(dirpath, fn)
                if os.access(p, os.X_OK):
                    cands.append(p)
            continue
        for fn in filenames:
            p = os.path.join(dirpath, fn)
            low = fn.lower()
            if low.endswith(".exe"):
                cands.append(p)
            elif sys.platform.startswith("linux") and not os.path.splitext(fn)[1] \
                    and os.access(p, os.X_OK):
                cands.append(p)
    if not cands:
        return None
    if prefer:
        for p in cands:
            if os.path.basename(p) == prefer:
                return p
    # 排除明显不是主程序的
    bad = ("uninstall", "vc_redist", "vcredist", "python", "tclsh", "wish")
    good = [p for p in cands if not any(b in os.path.basename(p).lower() for b in bad)]
    pool = good or cands
    # 主程序通常最大
    pool.sort(key=lambda p: os.path.getsize(p), reverse=True)
    return pool[0]


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", required=True)
    ap.add_argument("--label", required=True)
    ap.add_argument("--asset", default="")
    ap.add_argument("--exe", default="", help="主程序文件名（可选，帮忙定位）")
    ap.add_argument("--tag", default="", help="留空取 latest")
    ap.add_argument("--workdir", default="")
    ap.add_argument("--spec", default="")
    args = ap.parse_args(argv)

    token = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN") or ""
    work = args.workdir or os.path.join(
        os.environ.get("RUNNER_TEMP") or "/tmp", "crossplat")
    os.makedirs(work, exist_ok=True)

    if args.tag:
        api = "https://api.github.com/repos/%s/releases/tags/%s" % (args.repo, args.tag)
    else:
        api = "https://api.github.com/repos/%s/releases/latest" % args.repo
    print("查询 " + api)
    try:
        rel = api_get(api, token)
    except urllib.error.HTTPError as e:
        fail("取 release 失败（HTTP %s）：%s" % (e.code, api))
        return 2
    tag = rel.get("tag_name") or "?"
    print("tag = " + tag)

    asset = pick_asset(rel.get("assets") or [], args.asset)
    if not asset:
        names = ", ".join(a.get("name") or "" for a in rel.get("assets") or [])
        fail("%s 的 release %s 里找不到资产 %r（现有：%s）"
             % (args.repo, tag, args.asset, names))
        return 2
    name = asset["name"]
    zpath = os.path.join(work, name)
    print("下载 %s（%.1f MiB）" % (name, (asset.get("size") or 0) / 1048576.0))
    try:
        download(asset["browser_download_url"], zpath, token)
    except Exception as e:
        fail("下载失败：%s（%s）" % (name, e))
        return 2

    out = os.path.join(work, "unpacked")
    try:
        with zipfile.ZipFile(zpath) as z:
            z.extractall(out)
    except Exception as e:
        fail("解压失败：%s（%s）" % (name, e))
        return 2

    exe = find_executable(out, args.exe or None)
    if not exe:
        fail("解压后找不到可执行文件（%s）" % out)
        return 2
    print("可执行文件 = " + exe)
    notice("[%s] 受测产物 %s（tag %s）" % (args.label, name, tag))

    # 交给 verify_cli 做真正的比对
    vargs = ["--exe", exe, "--label", "%s %s" % (args.label, tag),
             "--workdir", os.path.join(work, "run")]
    if args.spec:
        vargs += ["--spec", args.spec]
    return verify_cli.main(vargs)


if __name__ == "__main__":
    sys.exit(main())
