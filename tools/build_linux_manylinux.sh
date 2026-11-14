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
echo "==================== 打包 AppImage ===================="
# 客户要的是「双击就能用」。Linux 上唯一能做到这件事的形态就是 .AppImage：
#   · 单文件，chmod +x 之后双击 / ./xxx.AppImage 直接跑，不需要先解压；
#   · 不需要客户装 libfuse2 —— 官方 type2 runtime 把 FUSE 静态链进去了
#     （2026-11-14 实测：922 KB 的 runtime 里有 APPIMAGE_EXTRACT_AND_RUN 与
#      appimage_extract，没有 libfuse.so.2 这个依赖项）；
#   · 万一客户环境连 FUSE 都没有（容器/精简系统），
#     APPIMAGE_EXTRACT_AND_RUN=1 ./xxx.AppImage 可以自解压运行。
# 之前交付的 linux.zip 是「解压出 60 多个 .so 的文件夹，再进目录双击」，
# 对不熟悉命令行的客户门槛过高，故本轮起 Linux 只发 AppImage。

VER="$(tr -d '[:space:]' < VERSION 2>/dev/null | sed 's/^[vV]//' || true)"
[ -n "$VER" ] || VER="0.0.0"

PAYLOAD=""
for d in dist/*/; do
  [ -d "$d" ] || continue
  PAYLOAD="${d%/}"
  break
done
if [ -z "$PAYLOAD" ]; then
  echo "::error::dist/ 下找不到构建产物目录，无法打 AppImage"
  ls -la dist/ 2>/dev/null || true
  exit 1
fi

# 主程序 = 产物目录里唯一一个可执行的普通文件（Nuitka standalone 的顶层只放它，
# 其余是 .so 与子目录）。不能写死名字：三个仓库分别叫 MDF-Stu / MDF-Advisor /
# paperformatpro，本脚本在三仓共用同一份。
BIN_NAME=""
for f in "$PAYLOAD"/*; do
  [ -f "$f" ] || continue
  case "$f" in *.so|*.so.*) continue ;; esac
  [ -x "$f" ] || continue
  BIN_NAME="${f##*/}"
  break
done
if [ -z "$BIN_NAME" ]; then
  echo "::error::$PAYLOAD 里找不到可执行主程序"
  ls -la "$PAYLOAD" | sed -n '1,20p'
  exit 1
fi

# AppImage 规范要求小写、无空格的 ID（用于 .desktop 的 Icon= 与图标文件名）。
APPID="$(printf '%s' "$BIN_NAME" | tr 'A-Z' 'a-z' | sed 's/[^a-z0-9]/-/g')"
# 桌面显示名取 build.py 里的 PRODUCT_NAME（学生版=文档格式先生 / 海外版=PaperFormat Pro）。
# .desktop 是 UTF-8，中文完全合法；取不到就退回二进制名。
DISPLAY_NAME="$(grep -m1 -E '^PRODUCT_NAME[[:space:]]*=' build.py 2>/dev/null \
  | sed -E 's/^[^"]*"([^"]*)".*/\1/' || true)"
[ -n "$DISPLAY_NAME" ] || DISPLAY_NAME="$BIN_NAME"
echo "  主程序      : $BIN_NAME"
echo "  应用 ID     : $APPID"
echo "  显示名      : $DISPLAY_NAME"
echo "  版本        : $VER"

# 图标：优先用源码树里的 icon.png（三仓路径不同：tfd_app/assets/ 或 assets/）。
ICON_SRC="$(find . -maxdepth 3 -name icon.png -not -path './dist/*' -not -path './.git/*' \
  -print 2>/dev/null | sed -n '1p')"
if [ -z "$ICON_SRC" ]; then
  echo "::error::找不到 icon.png，AppImage 必须有图标"
  exit 1
fi
echo "  图标        : $ICON_SRC"

# 线上取官方工具链。appimagetool 自己也是 AppImage，这里用 --appimage-extract 展开后
# 再跑 —— 容器里没有 /dev/fuse，只能走这条路（这也正是客户侧 FUSE 缺失时的兜底路径）。
TOOLS="/tmp/appimage-tools"
rm -rf "$TOOLS"
mkdir -p "$TOOLS"
AT_URL="https://github.com/AppImage/appimagetool/releases/download/continuous/appimagetool-x86_64.AppImage"
RT_URL="https://github.com/AppImage/type2-runtime/releases/download/continuous/runtime-x86_64"
echo "STEP 下载 appimagetool"
curl -fsSL --retry 5 --retry-delay 5 -o "$TOOLS/appimagetool" "$AT_URL" || {
  echo "::error::下载 appimagetool 失败：$AT_URL"; exit 1; }
echo "STEP 下载 type2 runtime（静态 FUSE）"
curl -fsSL --retry 5 --retry-delay 5 -o "$TOOLS/runtime-x86_64" "$RT_URL" || {
  echo "::error::下载 type2 runtime 失败：$RT_URL"; exit 1; }
chmod +x "$TOOLS/appimagetool"
ls -la "$TOOLS"
( cd "$TOOLS" && ./appimagetool --appimage-extract >/dev/null 2>&1 ) || {
  echo "::error::appimagetool 自解压失败"; exit 1; }

# ---- 组装 AppDir ----
APPDIR="/tmp/appimage-work/AppDir"
rm -rf "/tmp/appimage-work"
mkdir -p "$APPDIR/usr/bin"
cp -a "$PAYLOAD" "$APPDIR/usr/bin/$APPID"

# 精简：Linux 包里不该带 macOS 的 .icns（2.90 MB）和 Windows 的 .ico（0.36 MB），
# 共 3.26 MB 纯浪费。GUI 读取这两者时都有 isfile 守卫，删掉不会影响运行。
echo "--- 移除 Linux 用不到的图标 ---"
find "$APPDIR/usr/bin" \( -name 'icon.icns' -o -name 'icon.ico' \) -print -delete 2>/dev/null || true

# AppRun：必须 exec 真实二进制路径。用符号链接会让 $ORIGIN 落错地方，而 Nuitka
# standalone 的 RPATH=$ORIGIN 正是靠「可执行文件所在目录」去找同目录的 .so。
cat > "$APPDIR/AppRun" <<'APPRUN'
#!/bin/sh
# AppImage 入口：转交真正的可执行文件。
HERE="$(dirname "$(readlink -f "$0")")"
exec "$HERE/usr/bin/__APPID__/__BIN__" "$@"
APPRUN
sed -i "s|__APPID__|$APPID|g; s|__BIN__|$BIN_NAME|g" "$APPDIR/AppRun"
chmod +x "$APPDIR/AppRun"

# .desktop：Exec 用二进制名，并在 usr/bin 下放同名软链（桌面集成工具会把 Exec
# 当命令找；AppRun 走的是真实路径，两者互不干扰）。
ln -sf "$APPID/$BIN_NAME" "$APPDIR/usr/bin/$BIN_NAME"
cat > "$APPDIR/$APPID.desktop" <<DESKTOP
[Desktop Entry]
Type=Application
Name=$DISPLAY_NAME
Comment=$DISPLAY_NAME
Exec=$BIN_NAME
Icon=$APPID
Terminal=false
Categories=Office;Utility;
StartupNotify=true
DESKTOP

cp "$ICON_SRC" "$APPDIR/$APPID.png"
ln -sf "$APPID.png" "$APPDIR/.DirIcon"
mkdir -p "$APPDIR/usr/share/icons/hicolor/256x256/apps"
cp "$ICON_SRC" "$APPDIR/usr/share/icons/hicolor/256x256/apps/$APPID.png"

echo "--- AppDir 结构 ---"
ls -la "$APPDIR"
du -sh "$APPDIR" || true

# ---- 出包 ----
OUT="$ROOT/dist/$APPID-$VER-x86_64.AppImage"
rm -f "$OUT"
echo "STEP appimagetool 生成 $OUT"
# --runtime-file：强制用上面那份静态 FUSE 的官方 runtime。
# --no-appstream：跳过 appstreamcli 校验（容器里没有，且我们不提供 metainfo）。
ARCH=x86_64 VERSION="$VER" \
  "$TOOLS/squashfs-root/AppRun" --no-appstream \
  --runtime-file "$TOOLS/runtime-x86_64" "$APPDIR" "$OUT" || {
  echo "::error::appimagetool 失败"; exit 1; }

if [ ! -f "$OUT" ]; then
  echo "::error::AppImage 未生成"; exit 1
fi

# 自检：AppImage type2 = ELF 头 + 偏移 8 处的 0x41 0x49 0x02（"AI\x02"）。
MAGIC="$(od -An -tx1 -j8 -N3 "$OUT" | tr -d ' \n')"
SIZE="$(stat -c %s "$OUT" 2>/dev/null || echo 0)"
echo "  magic=$MAGIC  size=$SIZE"
if [ "$MAGIC" != "414902" ]; then
  echo "::error::产出的不是 AppImage type2（magic=$MAGIC）"; exit 1
fi
if [ "$SIZE" -lt 5000000 ]; then
  echo "::error::AppImage 只有 $SIZE 字节，明显不完整"; exit 1
fi
chmod +x "$OUT"
echo "✅ AppImage 生成成功：$OUT"

echo
echo "==================== 构建产物 ===================="
ls -la dist/ || true

# 产物属主是容器里的 root，交还给调用方方便后续打包/上传（宿主上 root 也无所谓）。
if [ -n "${HOST_UID:-}" ] && [ -n "${HOST_GID:-}" ]; then
  chown -R "${HOST_UID}:${HOST_GID}" dist release .build 2>/dev/null || true
fi

echo "构建完成"
