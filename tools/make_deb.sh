#!/usr/bin/env bash
# 把 Linux 产物打成 .deb（Debian / Ubuntu / Mint / 统信 UOS / 麒麟通用）。
#
# 为什么要 deb：AppImage 这条路客户侧有三个坑 ——
#   ① 下载/拷贝过来的文件没有执行权限（0666），双击毫无反应；
#   ② 运行时挂载需要系统的 fusermount，精简发行版可能没有（报错只进 stderr，
#      GUI 双击时客户什么也看不到）；
#   ③ /tmp 若被挂成 noexec，运行时也起不来。
# deb 由包管理器装到 /opt、权限自己设好、菜单自动出现，三个坑一次性消失：
# 客户双击 .deb → 软件管理器弹出 → 输密码 → 开始菜单里点图标，全程不碰终端。
#
# 本脚本要在 Debian 系（有 dpkg-deb）上跑，CI 里就是 ubuntu runner 这一步。
# 三仓共用同一份，产品名/二进制名/版本都从各自仓库里现读，不写死。
#
# 用法（仓库根目录）：bash tools/make_deb.sh
set -euo pipefail

cd "$(dirname "$0")/.." 2>/dev/null || true
ROOT="$(pwd)"

if ! command -v dpkg-deb >/dev/null 2>&1; then
  echo "::error::找不到 dpkg-deb —— 本脚本要在 Debian 系上跑（CI 里是 ubuntu runner）"
  exit 1
fi

VER="$(tr -d '[:space:]' < VERSION 2>/dev/null | sed 's/^[vV]//' || true)"
[ -n "$VER" ] || { echo "::error::读不到 VERSION"; exit 1; }

PAYLOAD=""
for d in dist/*/; do
  [ -d "$d" ] || continue
  PAYLOAD="${d%/}"
  break
done
[ -n "$PAYLOAD" ] || { echo "::error::dist/ 下没有产物目录"; ls -la dist/ || true; exit 1; }

# 主程序：产物目录里唯一一个可执行的普通文件（三仓分别是 MDF-Stu / MDF-Advisor /
# paperformatpro，不能写死）。
BIN_NAME=""
for f in "$PAYLOAD"/*; do
  [ -f "$f" ] || continue
  case "$f" in *.so|*.so.*) continue ;; esac
  [ -x "$f" ] || continue
  BIN_NAME="${f##*/}"
  break
done
[ -n "$BIN_NAME" ] || { echo "::error::$PAYLOAD 里找不到主程序"; ls -la "$PAYLOAD" | head -20; exit 1; }

APPID="$(printf '%s' "$BIN_NAME" | tr 'A-Z' 'a-z' | sed 's/[^a-z0-9]/-/g')"
DISPLAY_NAME="$(grep -m1 -E '^PRODUCT_NAME[[:space:]]*=' build.py 2>/dev/null \
  | sed -E 's/^[^"]*"([^"]*)".*/\1/' || true)"
[ -n "$DISPLAY_NAME" ] || DISPLAY_NAME="$BIN_NAME"
HOMEPAGE="$(grep -m1 -oE 'https?://[a-z0-9.-]+' README.md 2>/dev/null | sed -n '1p' || true)"
[ -n "$HOMEPAGE" ] || HOMEPAGE="https://reedskill.com"

ICON_SRC="$(find . -maxdepth 3 -name icon.png -not -path './dist/*' -not -path './.git/*' \
  -print 2>/dev/null | sed -n '1p')"

# ---- deb 元数据（逐产品）----
# 🔴 这份脚本三仓共用同一份，所以维护者邮箱、简介文字、首页域名都不能写死：
# 海外版必须全英文 + hi@paperformatpro.com，国内版是中文 + hi@reedskill.com。
# 唯一事实源 = tools/msix/products.json 里该产品的 "deb" 段（与 MSIX 打包同源）。
# 该文件不存在（国内两仓）→ 退回下面的国内默认文案。
META="tools/msix/products.json"
meta() {
  [ -f "$META" ] || return 0
  python3 - "$META" "$APPID" "$1" <<'PY' 2>/dev/null || true
import json, sys
path, appid, key = sys.argv[1:4]
try:
    data = json.load(open(path, encoding="utf-8"))
except Exception:
    raise SystemExit(0)
val = (data.get(appid) or {}).get("deb", {}).get(key)
if val:
    print(val)
PY
}

DEB_SYNOPSIS="$(meta synopsis)"
DEB_EXTENDED="$(meta extended)"
DEB_MAINTAINER="$(meta maintainer)"
DEB_HOMEPAGE="$(meta homepage)"
DEB_COPYRIGHT="$(meta copyright)"
DEB_SUPPORT="$(meta support)"

if [ -f "$META" ]; then
  MISSING=""
  [ -n "$DEB_SYNOPSIS" ]   || MISSING="$MISSING synopsis"
  [ -n "$DEB_MAINTAINER" ] || MISSING="$MISSING maintainer"
  [ -n "$DEB_HOMEPAGE" ]   || MISSING="$MISSING homepage"
  # 有 products.json 却取不到 → 直接红灯。绝不静默退回中文默认值，
  # 否则英文包里又会冒出中文简介（2026-11-17 客户截图就是这么来的）。
  [ -z "$MISSING" ] || { echo "::error::$META 中 $APPID 的 deb 段缺字段:$MISSING"; exit 1; }
fi

[ -n "$DEB_SYNOPSIS" ]   || DEB_SYNOPSIS="$DISPLAY_NAME"
[ -n "$DEB_EXTENDED" ]   || DEB_EXTENDED="模板驱动的论文格式检查与一键修正工具。全部在本机离线运行，不上传论文内容。"
[ -n "$DEB_MAINTAINER" ] || DEB_MAINTAINER="ReedSkill <hi@reedskill.com>"
[ -n "$DEB_HOMEPAGE" ]   || DEB_HOMEPAGE="$HOMEPAGE"
[ -n "$DEB_COPYRIGHT" ]  || DEB_COPYRIGHT="Copyright (c) ReedSkill · 芦苇不熬夜"
[ -n "$DEB_SUPPORT" ]    || DEB_SUPPORT="hi@reedskill.com"

echo "==================== 打 .deb ===================="
echo "  产品显示名 : $DISPLAY_NAME"
echo "  包名       : $APPID"
echo "  主程序     : $BIN_NAME"
echo "  版本       : $VER"
echo "  图标       : ${ICON_SRC:-（无，用图标名占位）}"
echo "  维护者     : $DEB_MAINTAINER"
echo "  首页       : $DEB_HOMEPAGE"
echo "  简介       : $DEB_SYNOPSIS"

WORK="/tmp/deb-work"
rm -rf "$WORK"
DEB="$WORK/pkg"
mkdir -p "$DEB/DEBIAN" \
         "$DEB/opt/$APPID" \
         "$DEB/usr/bin" \
         "$DEB/usr/share/applications" \
         "$DEB/usr/share/icons/hicolor/256x256/apps" \
         "$DEB/usr/share/doc/$APPID"

# ---- 程序本体：整个 standalone 目录装到 /opt/<appid> ----
# 为什么不塞 AppImage：这样就彻底不依赖 FUSE，也不需要客户 chmod。
cp -a "$PAYLOAD"/. "$DEB/opt/$APPID"/
# 双击入口（终端里敲命令也能用）
ln -sf "/opt/$APPID/$BIN_NAME" "$DEB/usr/bin/$BIN_NAME"

# ---- 菜单项 ----
cat > "$DEB/usr/share/applications/$APPID.desktop" <<DESKTOP
[Desktop Entry]
Type=Application
Name=$DISPLAY_NAME
Comment=$DISPLAY_NAME
Exec=/opt/$APPID/$BIN_NAME
Icon=$APPID
Terminal=false
Categories=Office;Utility;
StartupNotify=true
DESKTOP
chmod 644 "$DEB/usr/share/applications/$APPID.desktop"

if [ -n "$ICON_SRC" ]; then
  cp "$ICON_SRC" "$DEB/usr/share/icons/hicolor/256x256/apps/$APPID.png"
  chmod 644 "$DEB/usr/share/icons/hicolor/256x256/apps/$APPID.png"
fi

SIZE_KB="$(du -sk "$DEB" | cut -f1)"

# ---- control ----
# Depends 刻意压到最小：libc6/libstdc++6/zlib1g 是所有桌面发行版默认就有的，
# 写多了会让 apt 因缺包拒装（客户看到「依赖不满足」比装不上更糟）。
cat > "$DEB/DEBIAN/control" <<CONTROL
Package: $APPID
Version: $VER
Section: utils
Priority: optional
Architecture: amd64
Installed-Size: $SIZE_KB
Depends: libc6 (>= 2.28), libstdc++6, zlib1g
Maintainer: $DEB_MAINTAINER
Homepage: $DEB_HOMEPAGE
Description: $DEB_SYNOPSIS
 $DEB_EXTENDED
CONTROL
chmod 644 "$DEB/DEBIAN/control"

# ---- 安装后刷新菜单与图标缓存（命令不存在就跳过，不阻断安装）----
cat > "$DEB/DEBIAN/postinst" <<'POSTINST'
#!/bin/sh
set -e
if [ "$1" = "configure" ]; then
  update-desktop-database -q /usr/share/applications >/dev/null 2>&1 || true
  gtk-update-icon-cache -q -t -f /usr/share/icons/hicolor >/dev/null 2>&1 || true
fi
exit 0
POSTINST
cat > "$DEB/DEBIAN/postrm" <<'POSTRM'
#!/bin/sh
set -e
if [ "$1" = "remove" ] || [ "$1" = "purge" ]; then
  update-desktop-database -q /usr/share/applications >/dev/null 2>&1 || true
fi
exit 0
POSTRM
chmod 755 "$DEB/DEBIAN/postinst" "$DEB/DEBIAN/postrm"

cat > "$DEB/usr/share/doc/$APPID/copyright" <<COPYRIGHT
$DEB_COPYRIGHT
Homepage: $DEB_HOMEPAGE
Support: $DEB_SUPPORT
COPYRIGHT
chmod 644 "$DEB/usr/share/doc/$APPID/copyright"

OUT="$ROOT/dist/$APPID-$VER-amd64.deb"
rm -f "$OUT"
# --root-owner-group：强制包内属主为 root:root，否则 dpkg-deb 会因「属主诡异」报错
echo "STEP dpkg-deb --build -> $OUT"
dpkg-deb --root-owner-group -Zxz --build "$DEB" "$OUT" >/dev/null

# ==================== 自检（不过就红灯，绝不发出坏包）====================
# 🔴 这里所有管道都不能用 `grep -q`：本脚本开着 `set -o pipefail`，`grep -q` 一匹配
# 就提前关闭管道，上游 dpkg-deb 收到 SIGPIPE（141），整条管道被判失败 —— 检查会
# 间歇性误报「缺文件」。改成先整段取出来（awk/grep 读全量），再在变量上判断。
echo "STEP 自检"
dpkg-deb --info "$OUT" | sed -n '1,20p'
CONT="$(dpkg-deb --contents "$OUT")"

echo "--- 包内关键文件 ---"
printf '%s\n' "$CONT" | grep -E " \./opt/$APPID/$BIN_NAME$|/applications/| \./usr/bin/" || true

fail=0
# ① 主程序必须在，且带用户执行位（双击/菜单都靠它）
read -r MODE _ <<<"$(printf '%s\n' "$CONT" | awk -v t="./opt/$APPID/$BIN_NAME" '$NF==t{print $1; exit}')"
if [ -n "${MODE:-}" ] && [ "${MODE:3:1}" = "x" ]; then
  echo "  ✅ 主程序就位，权限 $MODE: $BIN_NAME"
else
  echo "  ❌ 主程序缺失或没有执行位（mode=${MODE:-无}）"
  fail=1
fi
# ② 菜单项必须在（装了没图标 = 客户找不到程序）
DESKLINE="$(printf '%s\n' "$CONT" | awk -v t="./usr/share/applications/$APPID.desktop" '$NF==t{print $1; exit}')"
if [ -n "${DESKLINE:-}" ]; then
  echo "  ✅ 菜单项已收录"
else
  echo "  ❌ 缺 .desktop，装了也没有图标"; fail=1
fi
# ③ Linux 用不到的图标不该被打进去
BADICON="$(printf '%s\n' "$CONT" | grep -E "icon\.(icns|ico)$" || true)"
if [ -n "$BADICON" ]; then
  echo "  ❌ 包里还有 icon.icns / icon.ico（白占 3.26 MB）:"
  printf '%s\n' "$BADICON" | sed -n '1,3p'
  fail=1
else
  echo "  ✅ 无多余图标"
fi
# ④ 体积下限
BYTES="$(stat -c %s "$OUT" 2>/dev/null || echo 0)"
echo "  体积: $BYTES 字节"
if [ "$BYTES" -lt 5000000 ]; then
  echo "  ❌ deb 小于 5 MB，明显不完整"; fail=1
fi

if [ "$fail" -ne 0 ]; then
  echo "::error::.deb 自检未通过"
  exit 1
fi
echo "✅ .deb 生成成功：$OUT"
