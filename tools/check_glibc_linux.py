# -*- coding: utf-8 -*-
"""扫描 Linux 产物的 glibc 需求，超过基线就红灯。

为什么要有这一步：2026-11-14 发现线上 Linux 包要求 GLIBC_2.42（GitHub 的
ubuntu-latest 已经升到 Ubuntu 26.04），Ubuntu 24.04 LTS 都装不上 —— 而 CI 当时
一路绿灯，因为它只验证「编译成功」，从不验证「客户机器能不能跑」。
本脚本把「能不能跑」变成可执行的断言。

原理：动态链接的 ELF 在 .dynstr 里留版本化符号（GLIBC_2.28 / GLIBC_2.42 …），
取全包最高版本即该包的最低系统要求。glibc 只能向下兼容。

用法：
    python tools/check_glibc_linux.py dist --max 2.28
    python tools/check_glibc_linux.py release/xxx-linux.zip --max 2.28
退出码：0 通过；1 超基线或缺产物；2 用法错误
"""
import os
import re
import sys
import zipfile

RE_GLIBC = re.compile(rb"GLIBC_(\d+)\.(\d+)(?:\.(\d+))?")

# 各发行版 glibc（用于把「要 2.42」翻译成人话，方便判断影响面）
DISTRO = [
    ((2, 17), "CentOS 7 / RHEL 7"),
    ((2, 28), "Debian 10 / RHEL 8 / 统信 UOS / 麒麟"),
    ((2, 31), "Ubuntu 20.04 / Debian 11"),
    ((2, 34), "RHEL 9"),
    ((2, 35), "Ubuntu 22.04"),
    ((2, 36), "Debian 12"),
    ((2, 38), "Ubuntu 23.10 / 24.04 之前"),
    ((2, 39), "Ubuntu 24.04 LTS"),
    ((2, 42), "Ubuntu 26.04 LTS"),
]


def parse_ver(s):
    parts = s.split(".")
    if len(parts) != 2:
        raise ValueError("版本要写成 2.28 这种形式：%r" % s)
    return (int(parts[0]), int(parts[1]))


def scan_bytes(data, name, found):
    if data[:4] != b"\x7fELF":
        return
    vers = [(int(a), int(b), int(c or 0)) for a, b, c in RE_GLIBC.findall(data)]
    if vers:
        found.append((max(vers), name))


def scan_path(path):
    found = []
    if os.path.isdir(path):
        for base, _dirs, files in os.walk(path):
            for fn in files:
                p = os.path.join(base, fn)
                try:
                    with open(p, "rb") as f:
                        head = f.read(4)
                        if head != b"\x7fELF":
                            continue
                        f.seek(0)
                        scan_bytes(head + f.read(), os.path.relpath(p, path), found)
                except OSError:
                    continue
    elif zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                with z.open(info) as f:
                    head = f.read(4)
                    if head != b"\x7fELF":
                        continue
                    scan_bytes(head + f.read(), info.filename, found)
    else:
        with open(path, "rb") as f:
            scan_bytes(f.read(), os.path.basename(path), found)
    return found


def main(argv):
    args = [a for a in argv[1:] if not a.startswith("--")]
    max_s = None
    for i, a in enumerate(argv):
        if a == "--max" and i + 1 < len(argv):
            max_s = argv[i + 1]
        elif a.startswith("--max="):
            max_s = a.split("=", 1)[1]
    if not args or not max_s:
        print("用法：python tools/check_glibc_linux.py <目录或zip> --max 2.28")
        return 2
    limit = parse_ver(max_s)
    target = args[0]

    if not os.path.exists(target):
        print("!! 产物不存在：%s" % target)
        return 1

    found = scan_path(target)
    if not found:
        print("!! 在 %s 里没扫到任何带 GLIBC 符号的 ELF —— 产物异常" % target)
        return 1

    found.sort(reverse=True)
    top = found[0][0]
    print("扫描目标：%s" % target)
    print("ELF 数：%d    最高需求：GLIBC_%d.%d" % (len(found), top[0], top[1]))

    worst = [n for v, n in found if v == top]
    if len(worst) <= 5:
        print("要这个版本的文件：%s" % "、".join(worst))
    else:
        print("要这个版本的文件（前 5 个，共 %d 个）：%s" % (len(worst), "、".join(worst[:5])))

    survivors = ["%s(%d.%d)" % (name, v[0], v[1]) for v, name in DISTRO if v >= limit]
    losers = ["%s(%d.%d)" % (name, v[0], v[1]) for v, name in DISTRO if v < limit]
    print("基线 glibc %d.%d 覆盖：%s" % (limit[0], limit[1], "、".join(survivors)))
    if losers:
        print("基线以下（本就不打算支持）：%s" % "、".join(losers))

    if (top[0], top[1]) > limit:
        print("\n::error::产物要求 GLIBC_%d.%d，高于基线 %d.%d —— 客户机器会报 "
              "'GLIBC_%d.%d not found'" % (top[0], top[1], limit[0], limit[1], top[0], top[1]))
        print("多半是构建环境变了（runner 升到更高版本的 glibc），请检查是否仍在 manylinux 容器内构建。")
        return 1

    print("\n✅ 通过：最高需求不高于基线，基线以上的系统都能直接跑")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
