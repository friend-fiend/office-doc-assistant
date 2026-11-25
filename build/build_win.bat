@echo off
chcp 65001 >nul
setlocal

REM ============================================================
REM  办公文档一键处理助手 —— Windows 打包脚本
REM  运行前请确认使用 python.org 的 64 位 Python 3.11（自带 tkinter）
REM  产物：dist\OfficeDocAssistant\OfficeDocAssistant.exe
REM ============================================================

cd /d "%~dp0.."
echo [1/6] 当前目录：%CD%

REM 版本号单一来源：从 spec 里读 VERSION，避免 spec/iss/bat 三处硬编码走偏。
REM findstr 按 "VERSION = " 开头定位那一行，用 = 和空格做分隔取第 2 列，
REM %%~v 去掉引号。读不到就退回下面的默认值，不会让脚本中断。
set "APPVER=1.0.5"
for /f "tokens=2 delims== " %%v in ('findstr /r /c:"^VERSION = " build\office_doc_assistant.spec') do set "APPVER=%%~v"
echo     版本号：%APPVER%

where python >nul 2>nul
if errorlevel 1 (
    echo [错误] 未找到 python，请先安装 Python 3.11 并加入 PATH。
    exit /b 1
)

REM 环境体检：缺 tkinter 时打包出来的是个打不开的 exe，所以必须前置拦住。
REM 用 --check 而不是直接 import tkinter，是为了同时给出其它缺失项的完整清单。
echo.
echo [1/6] 环境体检
set PYTHONIOENCODING=utf-8
python main.py --check
if errorlevel 1 (
    echo.
    echo [错误] 环境体检未通过（上面标 [失败] 的项）。
    echo        最常见原因是当前 Python 缺少 tkinter。
    echo        请改用 python.org 官方安装包（安装时勾选 tcl/tk and IDLE）后重试。
    exit /b 1
)

echo.
echo [2/6] 创建干净的打包虚拟环境 .venv-build
if not exist ".venv-build\Scripts\python.exe" (
    python -m venv .venv-build || exit /b 1
)

echo.
echo [3/6] 安装运行依赖
.venv-build\Scripts\python.exe -m pip install --upgrade pip -q || exit /b 1
.venv-build\Scripts\python.exe -m pip install -r requirements.txt -q || exit /b 1

echo.
echo [4/6] 安装 PyInstaller
.venv-build\Scripts\python.exe -m pip install "pyinstaller>=6.0" -q || exit /b 1

echo.
echo [5/6] 打包前自检（确认核心功能可用）
set PYTHONIOENCODING=utf-8
.venv-build\Scripts\python.exe tests\selftest.py || (
    echo [警告] 自检未全部通过，仍继续打包。请查看上面的失败项。
)

echo.
echo [6/6] 开始打包
rmdir /s /q build\OfficeDocAssistant 2>nul
rmdir /s /q dist\OfficeDocAssistant 2>nul

.venv-build\Scripts\python.exe -m PyInstaller build\office_doc_assistant.spec --noconfirm --clean || exit /b 1

echo.
echo 验证冻结产物...
dist\OfficeDocAssistant\OfficeDocAssistant.exe --check
if errorlevel 1 (
    echo [错误] 打包产物自检失败！安装包将不会生成，请先排查上面的输出。
    exit /b 1
)

REM ------------------------------------------------------------
REM 生成安装包（找到 ISCC 就自动执行，找不到则只出便携版并给出提示）
REM ------------------------------------------------------------
set "ISCC="
if exist "%LOCALAPPDATA%\Temp\innosetup6\ISCC.exe" set "ISCC=%LOCALAPPDATA%\Temp\innosetup6\ISCC.exe"
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"

if defined ISCC (
    echo.
    echo 生成安装包：dist\installer\OfficeDocAssistant-Setup-%APPVER%-win64.exe
    "%ISCC%" /DAppVersion=%APPVER% build\installer_win.iss || exit /b 1
    echo 生成便携版 zip...
    bash build\make_portable_win.sh 2>nul || powershell -NoProfile -Command "Compress-Archive -Path 'dist\OfficeDocAssistant\*' -DestinationPath 'dist\installer\OfficeDocAssistant-%APPVER%-portable-win64.zip' -Force"
) else (
    echo.
    echo [提示] 未找到 Inno Setup，跳过安装包，仅生成便携版 zip。
    echo        安装 Inno Setup 后重跑本脚本即可同时产出 Setup.exe。
    powershell -NoProfile -Command "Compress-Archive -Path 'dist\OfficeDocAssistant\*' -DestinationPath 'dist\installer\OfficeDocAssistant-%APPVER%-portable-win64.zip' -Force"
)

echo.
echo ============================================================
echo 全部完成，产物在 dist\installer\：
echo   OfficeDocAssistant-Setup-%APPVER%-win64.exe   （安装包，发给普通用户）
echo   OfficeDocAssistant-%APPVER%-portable-win64.zip（便携版，免安装）
echo.
echo 提示：
echo   · 若需要把 Poppler / Ghostscript 一起分发，把它们的绿色版
echo     解压到项目根目录的 vendor\ 后再运行本脚本。
echo   · 对外分发的签名/公证/升级注意事项见 DISTRIBUTION.md。
echo ============================================================
endlocal
