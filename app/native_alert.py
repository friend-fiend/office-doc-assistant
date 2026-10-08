"""不依赖 Tkinter 的"最后一道"错误提示通道。

为什么需要这个模块
------------------
最常见的启动失败场景恰恰是 **Tkinter 本身不可用**（Windows embeddable 包、
精简版 Python、Linux 未装 python3-tk、managed/uv 托管的解释器等）。

此时：
  · ``import tkinter`` 会失败  →  ``messagebox`` 也用不了
  · 用户看到的就是一个一闪而过的黑窗口 + 一串 traceback
  · 或者双击 exe 直接"什么都没发生"

所以必须有一条 **完全不依赖任何第三方库、也不依赖 Tk** 的提示路径。
本模块用各平台自带的原生机制实现：

  Windows  →  user32.dll 的 MessageBoxW（ctypes，标准库）
  macOS    →  osascript（系统自带）
  Linux    →  zenity / xmessage（桌面环境通常自带），否则退化为 stderr

所有函数都保证 **不抛异常**：提示本身失败绝不能反过来把程序搞崩。
"""
from __future__ import annotations

import json
import os
import subprocess
import sys

__all__ = ["alert", "warn", "error", "info"]

_ICON_ERROR = 0x10
_ICON_WARNING = 0x30
_ICON_INFO = 0x40

# Windows MessageBox 对超长文本不友好，统一截断
_MAX_TEXT = 2000


def _truncate(text: str) -> str:
    if len(text) <= _MAX_TEXT:
        return text
    head = text[: _MAX_TEXT - 200]
    return f"{head}\n\n…（内容过长，已截断，完整信息见日志文件）"


# ------------------------------------------------------------------ Windows
def _windows_message_box(title: str, text: str, icon: int) -> None:
    import ctypes  # 标准库，无需额外安装

    ctypes.windll.user32.MessageBoxW(None, text, title, icon | 0x0)


# ------------------------------------------------------------------ macOS
def _applescript_quote(s: str) -> str:
    """把 Python 字符串转成合法的 AppleScript 字符串表达式。

    AppleScript 的字符串字面量不支持 ``\\n`` 转义，换行必须写成
    ``" & return & "`` 这种拼接形式，所以这里手工构造。
    """
    escaped = s.replace("\\", "\\\\").replace('"', '\\"')
    return '"' + escaped.replace("\n", '" & return & "') + '"'


def _macos_dialog(title: str, text: str, icon: str) -> None:
    script = (
        f"display dialog {_applescript_quote(text)} "
        f'buttons {{"好"}} default button 1 with title {_applescript_quote(title)} '
        f"with icon {icon}"
    )
    subprocess.run(
        ["osascript", "-e", script],
        timeout=120,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )


# ------------------------------------------------------------------ Linux
def _linux_dialog(title: str, text: str, kind: str) -> bool:
    for cmd in (
        ["zenity", "--no-wrap", f"--{kind}", "--title", title, "--text", text],
        ["xmessage", "-center", "-title", title, f"{title}\n\n{text}"],
    ):
        try:
            subprocess.run(cmd, timeout=120, check=False)
            return True
        except FileNotFoundError:
            continue
        except Exception:
            continue
    return False


# ------------------------------------------------------------------ 统一入口
def alert(title: str, text: str, kind: str = "error") -> None:
    """弹出原生提示框；任何情况下都不抛异常。

    :param kind: ``error`` / ``warning`` / ``info``

    环境变量 ``ODA_SILENT_ALERT=1`` 可强制只写 stderr 不弹窗，
    供自动化测试 / CI / 无人值守脚本使用（否则原生弹窗会阻塞等人点确定）。
    """
    text = _truncate(text)

    if os.environ.get("ODA_SILENT_ALERT") == "1":
        _to_stderr(title, text)
        return

    try:
        if sys.platform.startswith("win"):
            icon = {"error": _ICON_ERROR, "warning": _ICON_WARNING}.get(kind, _ICON_INFO)
            _windows_message_box(title, text, icon)
            return
        if sys.platform == "darwin":
            _macos_dialog(title, text, "stop" if kind == "error" else "caution")
            return
        if _linux_dialog(title, text, "error" if kind == "error" else "warning"):
            return
    except Exception:
        pass

    # 所有图形通道都失败：至少写到 stderr，别让错误彻底消失
    _to_stderr(title, text)


def _to_stderr(title: str, text: str) -> None:
    try:
        print(f"\n[{title}]\n{text}\n", file=sys.stderr, flush=True)
    except Exception:
        pass


def error(title: str, text: str) -> None:
    alert(title, text, "error")


def warn(title: str, text: str) -> None:
    alert(title, text, "warning")


def info(title: str, text: str) -> None:
    alert(title, text, "info")


# ================================================================== tkinter 检查
_TK_HELP = {
    "win32": (
        "Windows 下最常见的原因是用了不含 Tkinter 的 Python：\n"
        "  · 「embeddable package」压缩包（python-3.x.x-embed-amd64.zip）不含 Tkinter\n"
        "  · Microsoft Store 版 Python 部分版本不含\n"
        "  · 某些精简/绿色版发行版不含\n\n"
        "解决办法：到 python.org 下载标准安装包（Windows installer 64-bit），\n"
        "安装时务必勾选「tcl/tk and IDLE」。\n"
        "装好后用新解释器运行：python -c \"import tkinter; print(tkinter.TkVersion)\""
    ),
    "darwin": (
        "macOS 下请使用 python.org 官方安装包或 Homebrew 的 python-tk：\n"
        "  brew install python-tk\n\n"
        "验证：python3 -c \"import tkinter; print(tkinter.TkVersion)\""
    ),
    "linux": (
        "Debian/Ubuntu： sudo apt install python3-tk\n"
        "Fedora/RHEL：  sudo dnf install python3-tkinter\n"
        "Arch：         sudo pacman -S tk\n\n"
        "验证：python3 -c \"import tkinter; print(tkinter.TkVersion)\""
    ),
}


def check_tkinter() -> tuple[bool, str]:
    """检查 Tkinter 是否可用。

    返回 ``(可用, 诊断信息)``。注意：本函数**只尝试 import 与创建解释器**，
    不创建窗口，可以在任何环境下安全调用。
    """
    try:
        import tkinter  # noqa: F401
    except Exception as exc:
        return False, f"import tkinter 失败：{type(exc).__name__}: {exc}"

    try:
        import tkinter

        # 只创建解释器，不建窗口 —— 无显示环境下也不会卡住
        interp = tkinter.Tcl()
        ver = interp.eval("info patchlevel")
        return True, f"Tk {tkinter.TkVersion}（Tcl/Tk {ver}）"
    except Exception as exc:
        return False, f"Tk 初始化失败：{type(exc).__name__}: {exc}"


def tkinter_help_text() -> str:
    key = "darwin" if sys.platform == "darwin" else ("win32" if sys.platform.startswith("win") else "linux")
    return _TK_HELP[key]


def python_where_help() -> str:
    """告诉用户在哪个解释器上出了问题 —— 否则用户根本不知道该换哪个。"""
    exe = sys.executable or "（未知）"
    lines = [
        f"当前解释器：{exe}",
        f"Python 版本：{sys.version.split()[0]}",
        f"平台：{sys.platform}",
    ]
    # venv 会暴露真实解释器，方便定位
    base = getattr(sys, "base_prefix", "")
    if base and base != sys.prefix:
        lines.append(f"虚拟环境基准：{base}")
    return "\n".join(lines) + f"\n工作目录：{os.getcwd()}"
