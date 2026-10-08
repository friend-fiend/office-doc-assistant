#!/bin/bash
# =====================================================================
#  Windows 便携版（绿色版）打包：把 onedir 产物压成 zip，免安装直接用
#  适用场景：没有管理员权限的电脑 / U 盘携带 / 不想跑安装包的同事
#  依赖：PowerShell（Windows 10/11 自带）
# =====================================================================
set -euo pipefail
cd "$(dirname "$0")/.."

# 版本号单一来源：优先读 spec 里的 VERSION，读不到再退回默认值。
# 注意用 if 而不是 `[ -z ... ] && ...`：后者在条件不成立时返回 1，
# 配合 set -e 会让脚本在"版本号读到了"这个正常路径上直接退出。
VERSION="$(sed -n 's/^VERSION = "\(.*\)"/\1/p' build/office_doc_assistant.spec | head -1)"
if [ -z "$VERSION" ]; then VERSION="1.0.2"; fi
OUT="dist/installer/OfficeDocAssistant-${VERSION}-portable-win64.zip"

if [ ! -f "dist/OfficeDocAssistant/OfficeDocAssistant.exe" ]; then
  echo "错误：先运行 build\\build_win.bat 产出 dist\\OfficeDocAssistant\\"
  exit 1
fi

mkdir -p dist/installer

# 便携版是"解压即用"，用户拿到的是一个文件夹，所以说明文档必须自己带齐——
# 安装包那边是 iss 用 Source: 单独挂进去的，这边得手动拷。
# 注意顺序：build_win.bat 先跑 ISCC 再跑本脚本，所以这些文件不会污染安装包。
mkdir -p dist/OfficeDocAssistant/docs
cp -f docs/USER_GUIDE.html dist/OfficeDocAssistant/docs/ 2>/dev/null || true
cp -f README.md LICENSE-NOTICES.txt dist/OfficeDocAssistant/ 2>/dev/null || true

# 写一份便携版说明进包里（版本号从 spec 读，和包名保持一致）
cat > dist/OfficeDocAssistant/使用说明.txt <<EOF
办公文档一键处理助手（便携版）  v${VERSION}

* 双击 OfficeDocAssistant.exe 即可使用，无需安装。
* 整个文件夹放在 U 盘/网盘里都可以直接运行。
* 设置与日志保存在 C:\Users\<你>\AppData\Local\OfficeDocAssistant\。
* PDF 转图片 / Office 转 PDF / 深度压缩需要外部工具（Poppler、
  LibreOffice、Ghostscript），缺哪个程序启动后会在"设置"里提示。

更完整的功能说明见 docs\USER_GUIDE.html（双击用浏览器打开）。
EOF

# 直接用 -Force 覆盖同名 zip，不要先 rm：某些受控环境会把 rm 当成危险操作拦下来
powershell.exe -NoProfile -Command \
  "Compress-Archive -Path 'dist\\OfficeDocAssistant\\*' -DestinationPath '$(cygpath -w "$PWD/$OUT")' -Force" \
  || { echo "压缩失败"; exit 1; }

ls -la "$OUT"
echo "便携版完成：$OUT"
