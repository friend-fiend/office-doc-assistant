#!/bin/bash
# =====================================================================
#  办公文档一键处理助手 —— macOS 打包脚本（必须在 Mac 上运行）
#
#  用法
#  ----
#     ./build/build_macos_dmg.sh                 # 本机架构
#     ./build/build_macos_dmg.sh universal2      # Intel + Apple Silicon 合体包
#     ./build/build_macos_dmg.sh arm64           # 仅 Apple Silicon
#     ./build/build_macos_dmg.sh x86_64          # 仅 Intel
#
#  产物
#  ----
#     dist/OfficeDocAssistant.app
#     dist/installer/OfficeDocAssistant-<版本>-macOS-<架构>.dmg
#
#  架构怎么选
#  ----------
#     * 2020 年末之后的 Mac 全是 arm64（Apple Silicon）；
#     * arm64 版在 Intel Mac 上【不能】运行；反之亦然；
#     * 不确定目标用户机器时用 universal2（体积约 +70%，但一份通吃）；
#     * 注意：universal2 需要安装 python.org 的 universal2 版 Python。
# =====================================================================
set -euo pipefail

ARCH="${1:-native}"
cd "$(dirname "$0")/.."
PROJ="$(pwd)"
VERSION="$(sed -n 's/^VERSION = "\(.*\)"/\1/p' build/office_doc_assistant.spec | head -1)"
# 用 if 而不是 `[ -z ... ] && ...`：后者在"版本号读到了"时返回 1，
# 配合 set -e 会让脚本在正常路径上直接退出。
if [ -z "$VERSION" ]; then VERSION="1.0.3"; fi
APP_NAME="OfficeDocAssistant"

echo "==== [1/7] 环境体检 ===="
python3 main.py --check || {
  echo "体检未通过。最常见的坑："
  echo "  · 用了 Homebrew 的 python：优先用 python.org 官方安装包"
  echo "    （brew 版升级后路径会失效，编译出来的 app 也容易跟着坏）"
  exit 1
}

echo "==== [2/7] 干净的构建虚拟环境 .venv-build ===="
if [ ! -x ".venv-build/bin/python" ]; then
  python3 -m venv .venv-build
fi
PY=".venv-build/bin/python"

echo "==== [3/7] 安装依赖 ===="
"$PY" -m pip install --upgrade pip -q
"$PY" -m pip install -r requirements.txt "pyinstaller>=6.0" -q

echo "==== [4/7] 核心自检 ===="
PYTHONIOENCODING=utf-8 "$PY" tests/selftest.py || echo "[警告] 自检未全部通过，仍继续。"

echo "==== [5/7] PyInstaller 打包 ===="
rm -rf "dist/$APP_NAME.app" "build/$APP_NAME"
if [ "$ARCH" != "native" ]; then
  export ODA_TARGET_ARCH="$ARCH"   # spec 通过该变量设置 target_arch
fi
"$PY" -m PyInstaller build/office_doc_assistant.spec --noconfirm --clean

echo "==== [6/7] 清理隔离属性 + ad-hoc 签名 ===="
# 从网络下载/跨机器拷贝来的文件带 quarantine 标记，会让 Gatekeeper 拦下 app
xattr -cr "dist/$APP_NAME.app" || true
# 没有 Developer ID 时先做 ad-hoc 签名，保证本机与同版本 macOS 可直接打开
codesign --force --deep --sign - "dist/$APP_NAME.app" 2>/dev/null || true

echo "==== [7/7] 制作 DMG ===="
mkdir -p dist/installer
if [ "$ARCH" = "native" ]; then
  detected="$("$PY" -c 'import platform; print(platform.machine())')"
else
  detected="$ARCH"
fi
DMG="dist/installer/${APP_NAME}-${VERSION}-macOS-${detected}.dmg"
rm -f "$DMG"
hdiutil create -volname "$APP_NAME $VERSION" \
  -srcfolder "dist/$APP_NAME.app" -ov -format UDZO "$DMG"
# UDZO (压缩) 比 UDRW 小得多；如需"拖到 Applications 文件夹"的美化布局，
# 可改用 dmgbuild 或 create-dmg，这里保持零第三方依赖。

echo
echo "完成：$DMG"
echo
echo "==== 对外分发前的两步（可选但强烈建议）===="
echo "1) 用 Developer ID 签名（否则别人打开会被 Gatekeeper 拦）："
echo "   codesign --force --options runtime --deep --sign 'Developer ID Application: 你的名字(TeamID)' dist/$APP_NAME.app"
echo "2) 公证（否则用户第一次打开要右键->打开 或 xattr -cr）："
echo "   xcrun notarytool submit $DMG --apple-id you@example.com --team-id TEAMID --password '@keychain:AC_PASSWORD' --wait"
echo "   xcrun stapler staple $DMG"
