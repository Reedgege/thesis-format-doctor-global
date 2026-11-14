#!/usr/bin/env bash
# 在 manylinux_2_28 容器里构建 Linux 版（glibc 2.28 基线）。
#
# 为什么必须进容器：
#   GitHub 的 ubuntu-latest 已经跑到 Ubuntu 26.04（glibc 2.42），在这里编出来的产物
#   会硬要求 GLIBC_2.42 —— Ubuntu 24.04 LTS（2.39）都装不上，等于没人能用。
#   glibc 只能向下兼容，所以只能在**老**环境里编，编完才能在**新**系统上跑。
#
# manylinux_2_28 = AlmaLinux 8 = glibc 2.28，覆盖：
#   Ubuntu 20.04/22.04/24.04/26.04、Debian 10/11/12、RHEL 8/9、
#   以及国内信创的统信 UOS、麒麟（都是 glibc 2.28 起）
#
# 用法（在仓库根目录）：
#   docker run --rm -v "$PWD:/src" -w /src quay.io/pypa/manylinux_2_28_x86_64 \
#       bash /src/tools/build_linux_manylinux.sh
set -euo pipefail

cd "$(dirname "$0")/.." 2>/dev/null || true
ROOT="$(pwd)"

echo "==================== 容器环境 ===================="
# 🔴 这里**绝不能**用 `| head -1`：set -o pipefail 下 head 提前关闭管道会让上游命令
# 收到 SIGPIPE（退出码 141），整个脚本当场中止。2026-11-14 第一次探针就栽在这。
grep -E "^(PRETTY_NAME|VERSION_ID)=" /etc/os-release || true
ldd --version 2>&1 | sed -n '1p' || true
uname -m

echo
echo "==================== 磁盘与内存 ===================="
df -h / /tmp 2>/dev/null || true
free -m 2>/dev/null || true

echo
echo "==================== 安装构建依赖 ===================="
# tkinter 是硬需求：GUI 全靠它。AlmaLinux 8 的 python3.12/3.11 都有独立的 -tkinter 子包，
# 但不同小版本镜像里可用的 Python 版本不一样，所以候选全装一遍、装不上的跳过。
# install_weak_deps=False：不加这条 dnf 会顺着弱依赖把一堆 X 桌面件拖进来，慢且容易撑爆磁盘。
DNFFLAGS="-y -q --setopt=install_weak_deps=False --nogpgcheck"
for p in gcc gcc-c++ make zip unzip which patchelf xz tar \
         tk-devel tcl-devel \
         python3.12 python3.12-devel python3.12-tkinter python3.12-pip \
         python3.11 python3.11-devel python3.11-tkinter python3.11-pip ; do
  echo "STEP install $p"
  if dnf $DNFFLAGS install "$p" >/dev/null 2>&1; then
    echo "  ok   $p"
  else
    echo "  skip $p（仓库里没有，或已装）"
  fi
done
echo "STEP install done"

echo
echo "==================== 挑一个带 tkinter 的解释器 ===================="
# 注意：/opt/python/cp312-* 是 pypa 自带的构建，**通常没有 _tkinter**（编 Python 时没链 tcl/tk），
# 所以它排在最后，只作为兜底。
PY=""
for cand in python3.12 python3.11 python3 /opt/python/cp312-cp312/bin/python3.12; do
  if ! command -v "$cand" >/dev/null 2>&1 && [ ! -x "$cand" ]; then
    echo "  --   $cand 不存在"
    continue
  fi
  if "$cand" -c "import tkinter, sys; print('tkinter ok', sys.version.split()[0])" 2>/dev/null; then
    echo "  ✅   $cand  ===> 用它"
    PY="$cand"
    break
  else
    echo "  ❌   $cand 没有 tkinter（或缺 tcl/tk 运行库）"
  fi
done

if [ -z "$PY" ]; then
  echo "::error::manylinux 容器里找不到任何带 tkinter 的 Python —— 无法构建 GUI 版"
  echo "排查方向：dnf 是否装上了 python3.12-tkinter / tk-devel；或改用 /opt/python 并自行编译 _tkinter"
  exit 1
fi
echo "选定解释器：$($PY -c 'import sys; print(sys.executable, sys.version.split()[0])')"

pinstall() {
  "$PY" -m pip install "$@" 2>/dev/null || "$PY" -m pip install --break-system-packages "$@"
}

echo
echo "==================== 安装 Nuitka 与依赖 ===================="
"$PY" -m ensurepip --upgrade >/dev/null 2>&1 || true
pinstall --upgrade pip setuptools wheel
pinstall --upgrade nuitka ordered-set zstandard
if [ -f requirements.txt ]; then
  echo "  发现 requirements.txt，安装"
  pinstall -r requirements.txt
else
  echo "  无 requirements.txt（本项目仅用标准库）"
fi

echo
echo "==================== 准备静态 libpython ===================="
# 🔴 pypa 的 manylinux 镜像里 /opt/python/cp3xx 是「为 embed 准备」的构建，只带共享库、
# 不带静态 libpython3.12.a，Nuitka 做 --standalone 静态链接时会直接 FATAL：
#   "Automatic detection of static libpython failed. Nuitka on 'manylinux' has no shared libraries."
# 官方让解出镜像自带的 static-libs-for-embedding-only.tar.xz 即可。
# 2026-11-14 第 3 次探针就是死在这里。
SLIB=""
for cand in /opt/_internal/static-libs-for-embedding-only.tar.xz \
            /opt/python/static-libs-for-embedding-only.tar.xz; do
  [ -f "$cand" ] && SLIB="$cand" && break
done
if [ -z "$SLIB" ]; then
  SLIB=$(find /opt -maxdepth 3 -name 'static-libs-for-embedding-only.tar.xz' 2>/dev/null | sed -n '1p')
fi
if [ -n "$SLIB" ]; then
  echo "STEP 解出 $SLIB"
  ( cd "$(dirname "$SLIB")" && tar xf "$(basename "$SLIB")" ) && echo "  解出完成" || echo "  !! 解出失败"
else
  echo "  -- 镜像里没有 static-libs 包（若用的是系统 python 则不需要）"
fi
echo "--- 现有的 libpython ---"
ls -la /opt/python/cp312-cp312/lib/libpython3.12* 2>/dev/null || true
ls -la /opt/python/cp312-cp312/lib/python3.12/config-*/libpython3.12* 2>/dev/null || true

echo
echo "==================== 开始构建 ===================="
"$PY" build.py

echo
echo "==================== 构建产物 ===================="
ls -la dist/ || true

# 产物属主是容器里的 root，交还给调用方方便后续打包/上传（宿主上 root 也无所谓）。
if [ -n "${HOST_UID:-}" ] && [ -n "${HOST_GID:-}" ]; then
  chown -R "${HOST_UID}:${HOST_GID}" dist release .build 2>/dev/null || true
fi

echo "构建完成"
