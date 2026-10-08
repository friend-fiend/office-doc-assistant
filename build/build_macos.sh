#!/usr/bin/env bash
# ============================================================
#  办公文档一键处理助手 —— macOS 打包脚本
#  产物：dist/OfficeDocAssistant.app
#
#  重要：macOS 上必须用 python.org 官方安装包提供的 Python
#  （自带 Tk 8.6.13）。Homebrew 的 python 有时缺 _tkinter，
#  系统自带的 /usr/bin/python3 版本又过旧，都不要用。
# ============================================================
set -euo pipefail

cd "$(dirname "$0")/.."
echo "[1/7] 工作目录：$(pwd)"

PY=${PY:-python3.11}
if ! command -v "$PY" >/dev/null 2>&1; then
  echo "[错误] 未找到 $PY。请安装 python.org 的 Python 3.11。"
  echo "       可用 PY=python3.12 ./build/build_macos.sh 指定其他版本。"
  exit 1
fi

"$PY" - <<'EOF'
import sys, tkinter
if sys.platform != "darwin":
    sys.exit("请在 macOS 上运行此脚本")
print(f"Python {sys.version.split()[0]} | Tk {tkinter.TkVersion}")
if tkinter.TkVersion < 8.6:
    sys.exit("Tk 版本过低（需要 8.6+），请改用 python.org 官方 Python")
EOF

echo
echo "[2/7] 检查/创建虚拟环境 .venv-build"
if [ ! -x ".venv-build/bin/python" ]; then
  "$PY" -m venv .venv-build
fi
VENV_PY=".venv-build/bin/python"

echo
echo "[3/7] 安装运行依赖"
"$VENV_PY" -m pip install --upgrade pip -q
"$VENV_PY" -m pip install -r requirements.txt -q

echo
echo "[4/7] 安装 PyInstaller"
"$VENV_PY" -m pip install "pyinstaller>=6.0" -q

echo
echo "[5/7] 环境体检 + 外部依赖"
# 体检用打包环境自己的解释器跑：真正会被打进 app 的就是它。
# 缺 tkinter 时打包出来的 .app 是打不开的，所以这里必须硬拦截。
export PYTHONIOENCODING=utf-8
if ! "$VENV_PY" main.py --check; then
  echo
  echo "[错误] 环境体检未通过（上面标 [失败] 的项）。"
  echo "       macOS 上最常见原因是 Tk 缺失或版本过低。"
  echo "       请改用 python.org 官方 Python，或执行 brew install python-tk 后重试。"
  exit 1
fi
# 上面这次体检已经涵盖 Poppler / LibreOffice / Ghostscript 的探测与安装指引，
# 不再重复检查一遍。
echo "  （外部程序为可选项，缺失只会禁用对应功能，不影响打包）"

echo
echo "[6/7] 打包前自检"
"$VENV_PY" tests/selftest.py || echo "[警告] 自检未全部通过，仍继续打包。"

echo
echo "[7/7] 开始打包"
rm -rf build/OfficeDocAssistant dist/OfficeDocAssistant dist/OfficeDocAssistant.app
"$VENV_PY" -m PyInstaller build/office_doc_assistant.spec --noconfirm --clean

APP="dist/OfficeDocAssistant.app"

echo
echo "清理隔离属性（从网盘/聊天工具下载的目录常带 com.apple.quarantine，会导致打不开）"
xattr -cr "$APP" 2>/dev/null || true

# 临时签名：让本机能直接打开。正式对外分发需要 Developer ID + 公证，见 README。
echo "临时 ad-hoc 签名"
codesign --force --deep --sign - "$APP" 2>/dev/null || echo "  （跳过签名，不影响本机运行）"

echo
echo "============================================================"
echo "打包完成：$APP"
echo
echo "本机验证：open $APP"
echo "分发注意：未做公证的 app 在别人机器上首次打开需右键 → 打开。"
echo "          正式分发请执行："
echo "            codesign --deep --force --options runtime \\"
echo "                     --sign 'Developer ID Application: 你的名字 (TEAMID)' $APP"
echo "            xcrun notarytool submit ... --wait"
echo "            xcrun stapler staple $APP"
echo "============================================================"
