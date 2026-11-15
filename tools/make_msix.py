# -*- coding: utf-8 -*-
"""把现有的 Windows 产物目录打成 .msix（微软商店要的包格式）。

设计原则：
  · 不重建、不改动产物 —— 直接拿 dist/<产品>/ 或解压出来的发行包，
    原样塞进 MSIX，程序运行行为与 zip 版完全一致；
  · makeappx 是 Windows SDK 自带的命令行工具，本机没装 SDK 也能用
    （GitHub Actions 的 windows-latest 自带），找不到会给出明确提示；
  · Identity 三行从 tools/msix/products.json 读，绝不散落在代码里。

用法：
    python tools/make_msix.py --product paperformatpro --payload dist/paperformatpro
    python tools/make_msix.py --product mdf-stu --payload build/MDF-Stu --version 1.3.162 \
        --out dist/MDF-Stu.msix

打完包后本机可用 PowerShell 自签安装验证（Store 上传不需要自己签）：
    New-SelfSignedCertificate -Type Custom -Subject "<Publisher>" -KeyUsage DigitalSignature \
        -CertStoreLocation "Cert:\\CurrentUser\\My" -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3","2.5.29.19={text}")
    signtool sign /fd SHA256 /a /f <导出.pfx> /p <密码> dist/xxx.msix
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys

# 🔴 Windows CI 上 stdout 走管道时是 cp1252，打印任何中文都会 UnicodeEncodeError
# （报错点还落在无关的 print 行上，极难猜）。这里统一改成 UTF-8 兜底。
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
MSIX_DIR = os.path.join(HERE, "msix")


def find_makeappx(explicit=None):
    """按 显式参数 → PATH → Windows Kits 目录 的顺序找 makeappx.exe。"""
    if explicit:
        if os.path.isfile(explicit):
            return explicit
        raise SystemExit("指定的 makeappx 不存在：%s" % explicit)
    onpath = shutil.which("makeappx") or shutil.which("makeappx.exe")
    if onpath:
        return onpath
    kits = [
        r"C:\Program Files (x86)\Windows Kits\10\bin",
        r"C:\Program Files\Windows Kits\10\bin",
    ]
    cands = []
    for base in kits:
        if not os.path.isdir(base):
            continue
        for ver in os.listdir(base):
            p = os.path.join(base, ver, "x64", "makeappx.exe")
            if os.path.isfile(p):
                cands.append((ver, p))
    if cands:
        cands.sort(reverse=True)
        return cands[0][1]
    raise SystemExit(
        "找不到 makeappx.exe。两条路：\n"
        "  1) 在 CI 里跑（GitHub Actions 的 windows-latest 自带 Windows SDK）；\n"
        "  2) 本机装 Windows SDK（含「MSIX 打包工具」）或用 --makeappx 指定路径。"
    )


def norm_version(v):
    """MSIX 版本必须是四段，我们的 VERSION 是三段，补 .0。"""
    v = str(v).strip().lstrip("vV")
    parts = [p for p in re.split(r"[.\-+]", v) if p.isdigit()]
    if not parts:
        raise SystemExit("版本号解析失败：%r" % v)
    while len(parts) < 4:
        parts.append("0")
    parts = parts[:4]
    for p in parts:
        if int(p) > 65535:
            raise SystemExit("MSIX 版本段不能超过 65535：%s" % v)
    return ".".join(str(int(p)) for p in parts)


def read_version(payload):
    """产物目录里带 VERSION 文件，优先用它的值。"""
    for cand in (os.path.join(payload, "VERSION"), os.path.join(ROOT, "VERSION")):
        if os.path.isfile(cand):
            with open(cand, encoding="utf-8") as fh:
                val = fh.read().strip()
            if val:
                return val
    return "0.0.0"


def resolve_payload(payload):
    """容忍两种形态：目录里直接是 exe 与 .dll；或套了一层产品文件夹。"""
    if not os.path.isdir(payload):
        raise SystemExit("产物目录不存在：%s" % payload)
    exes = []
    for name in os.listdir(payload):
        p = os.path.join(payload, name)
        if os.path.isfile(p) and name.lower().endswith(".exe"):
            exes.append(name)
    if exes:
        return payload, exes
    subs = [d for d in os.listdir(payload) if os.path.isdir(os.path.join(payload, d))]
    if len(subs) == 1:
        inner = os.path.join(payload, subs[0])
        exes = [n for n in os.listdir(inner)
                if n.lower().endswith(".exe") and os.path.isfile(os.path.join(inner, n))]
        if exes:
            print("  产物在子目录里，自动下钻：%s" % subs[0])
            return inner, exes
    raise SystemExit("在 %s 里找不到 .exe 主程序" % payload)


def build_manifest(prod, version, executable):
    tmpl_path = os.path.join(MSIX_DIR, "AppxManifest.xml.tmpl")
    with open(tmpl_path, encoding="utf-8") as fh:
        tmpl = fh.read()
    mapping = {
        "@IDENTITY_NAME@": prod["identity_name"],
        "@PUBLISHER@": prod["publisher"],
        "@VERSION@": version,
        "@DISPLAY_NAME@": prod["display_name"],
        "@PUBLISHER_DISPLAY_NAME@": prod["publisher_display_name"],
        "@SHORT_NAME@": prod["short_name"],
        "@DESCRIPTION@": prod["description"],
        "@LANG@": prod["lang"],
        "@EXECUTABLE@": executable,
    }
    for k, v in mapping.items():
        if "@" in str(v):
            raise SystemExit("产品配置里 %s 还没填（当前值 %r）" % (k, v))
        tmpl = tmpl.replace(k, str(v))
    left = re.findall(r"@[A-Z_]+@", tmpl)
    if left:
        raise SystemExit("清单里还有没替换的占位符：%s" % sorted(set(left)))
    return tmpl


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--product", required=True, help="products.json 里的键")
    ap.add_argument("--payload", required=True, help="Windows 产物目录")
    ap.add_argument("--version", default=None, help="缺省读产物里的 VERSION")
    ap.add_argument("--out", default=None)
    ap.add_argument("--makeappx", default=None)
    ap.add_argument("--keep-stage", action="store_true", help="保留中间目录便于排查")
    args = ap.parse_args()

    with open(os.path.join(MSIX_DIR, "products.json"), encoding="utf-8") as fh:
        products = json.load(fh)
    if args.product not in products or args.product.startswith("_"):
        raise SystemExit("products.json 里没有产品 %r（可选：%s）"
                         % (args.product,
                            [k for k in products if not k.startswith("_")]))
    prod = products[args.product]
    if prod.get("identity_name") in (None, "", "TBD"):
        raise SystemExit("产品 %s 的 identity_name 还没填，先去 Partner Center 取" % args.product)

    payload, exes = resolve_payload(args.payload)
    exe = prod.get("executable") or exes[0]
    if exe not in exes:
        raise SystemExit("产物里没有 %s（实际有：%s）" % (exe, ", ".join(sorted(exes))))

    ver_raw = args.version or read_version(payload)
    version = norm_version(ver_raw)
    out = args.out or os.path.join(ROOT, "dist", "%s-%s-x64.msix" % (args.product, ver_raw.lstrip("vV")))
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)

    makeappx = find_makeappx(args.makeappx)
    print("makeappx  : %s" % makeappx)
    print("产品      : %s（%s）" % (args.product, prod["label"]))
    print("Identity  : %s" % prod["identity_name"])
    print("Publisher : %s" % prod["publisher"])
    print("版本      : %s -> %s" % (ver_raw, version))
    print("产物目录  : %s" % payload)

    stage = os.path.join(ROOT, "build", "msix-stage", args.product)
    if os.path.isdir(stage):
        shutil.rmtree(stage)
    os.makedirs(stage)

    # ① 素材
    src_assets = os.path.join(MSIX_DIR, "assets")
    if not os.path.isdir(src_assets):
        raise SystemExit("缺 MSIX 素材目录，先跑 python tools/msix/make_assets.py")
    shutil.copytree(src_assets, os.path.join(stage, "Assets"))

    # ② 程序本体：整体放进 app\ 子目录，清单里的 Executable 相应写 app\<exe>
    app_dir = os.path.join(stage, "app")
    print("  复制程序文件 ...")
    shutil.copytree(payload, app_dir)

    # ③ 清单
    manifest = build_manifest(prod, version, "app\\" + exe)
    with open(os.path.join(stage, "AppxManifest.xml"), "w", encoding="utf-8") as fh:
        fh.write(manifest)

    total = 0
    count = 0
    for base, _dirs, files in os.walk(stage):
        for f in files:
            total += os.path.getsize(os.path.join(base, f))
            count += 1
    print("  待打包：%d 个文件 / %.1f MiB" % (count, total / 1048576.0))

    if os.path.exists(out):
        os.remove(out)
    cmd = [makeappx, "pack", "/d", stage, "/p", out, "/o"]
    print("STEP %s" % " ".join(cmd))
    r = subprocess.run(cmd, capture_output=True, text=True, errors="replace")
    # makeappx 的输出动辄十几万字符，CI 里 stdout 只能看到一截，
    # 所以**全量落盘**，失败排查以这个文件为准。
    log_path = os.path.join(os.path.dirname(os.path.abspath(out)) or ".", "makeappx.log")
    try:
        with open(log_path, "w", encoding="utf-8") as fh:
            fh.write("CMD: %s\n\n--- stdout ---\n%s\n\n--- stderr ---\n%s\n"
                     % (" ".join(cmd), r.stdout or "", r.stderr or ""))
        print("完整 makeappx 日志：%s" % log_path)
    except OSError as exc:
        print("写 makeappx 日志失败：%s" % exc)
    print(r.stdout.strip()[-3000:] if r.stdout else "")
    if r.returncode != 0:
        print(r.stderr.strip()[-3000:] if r.stderr else "")
        raise SystemExit("makeappx 打包失败，退出码 %d" % r.returncode)
    if not os.path.isfile(out):
        raise SystemExit("makeappx 说成功但没产物")

    size = os.path.getsize(out)
    print("\nOK  %s  %.2f MiB" % (out, size / 1048576.0))
    if size < 5 * 1024 * 1024:
        raise SystemExit("产物只有 %.2f MiB，明显不完整" % (size / 1048576.0))

    if not args.keep_stage:
        shutil.rmtree(stage, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
