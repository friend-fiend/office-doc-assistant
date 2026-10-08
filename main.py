#!/usr/bin/env python3
"""办公文档一键处理助手 —— 程序入口。

直接运行：``python main.py``
打包后运行：双击可执行文件 / .app

启动顺序（顺序很重要）
----------------------
1. 修正 sys.path
2. Windows 高 DPI 感知   —— 必须在创建 Tk 根窗口之前
3. **Tkinter 可用性守卫** —— 失败时用原生弹窗报错（此时绝对不能用 tkinter 报错）
4. 初始化日志
5. 导入并启动界面

命令行长这样时只做体检、不启动界面：
    python main.py --check      # 环境体检（人类可读）
    python main.py --doctor     # 同上，附带外部依赖路径
"""
from __future__ import annotations

import os
import sys
import traceback

# 让 `python main.py` 与打包后的可执行文件都能正确 import app 包
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)


def _enable_windows_dpi_awareness() -> None:
    """Windows 高 DPI 屏：不做这个设置界面会被系统拉伸成模糊的。

    必须在创建 Tk 根窗口之前调用。
    """
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes

        # PROCESS_PER_MONITOR_DPI_AWARE = 2
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            import ctypes

            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _enable_macos_app_menu() -> None:
    """macOS 下让程序出现在 Dock 并以窗口形式激活。"""
    if sys.platform != "darwin":
        return
    try:
        from AppKit import NSApplication  # type: ignore

        NSApplication.sharedApplication().setActivationPolicy_(0)
    except Exception:
        pass


# ==================================================================== 启动守卫
def _guard_tkinter() -> bool:
    """确认 Tkinter 可用。

    这是整个程序唯一的硬性前提：没有 Tkinter 就没有界面。
    因为此时 ``tkinter`` 一定不可用，所以报错必须走原生弹窗通道。
    """
    from app.native_alert import check_tkinter, tkinter_help_text, python_where_help

    ok, detail = check_tkinter()
    if ok:
        return True

    from app.config import APP_NAME

    body = (
        f"{APP_NAME} 需要 Tkinter 才能显示界面，但当前 Python 解释器中不可用。\n\n"
        f"检测结果：{detail}\n\n"
        f"──────────────────────\n"
        f"{python_where_help()}\n"
        f"──────────────────────\n\n"
        f"{tkinter_help_text()}"
    )
    from app.native_alert import error

    # error() 本身已经保证写到 stderr，无需重复打印
    error("缺少 Tkinter，无法启动", body)
    return False


def _guard_python_version() -> bool:
    """Python 版本硬性检查。

    3.13 起官方已放弃 Windows 8.1；同时部分第三方库对 <3.10 不兼容。
    这里只做"能否继续运行"的判断，不阻断新版本（新版本自行承担风险）。
    """
    from app.native_alert import error, python_where_help

    vi = sys.version_info
    if vi >= (3, 10):
        return True

    body = (
        f"当前 Python 版本为 {vi.major}.{vi.minor}.{vi.micro}，过低。\n\n"
        "本程序需要 Python 3.10 或更高版本。\n\n"
        f"{python_where_help()}"
    )
    error("Python 版本过低", body)
    return False


def _run_preflight(verbose: bool) -> int:
    """``--check`` / ``--doctor``：只体检不启动界面。

    即使 Tkinter 缺失也能跑完（体检脚本本身不依赖 Tkinter），
    这正是它存在的意义。
    """
    try:
        from tools.preflight import main as preflight_main
    except Exception as exc:
        print(f"无法加载体检脚本：{type(exc).__name__}: {exc}", file=sys.stderr)
        print(f"（预期位置：{os.path.join(_HERE, 'tools', 'preflight.py')}）", file=sys.stderr)
        return 2
    return preflight_main(verbose=verbose)


# ======================================================================== 主流程
def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)

    if "--check" in argv or "--doctor" in argv:
        return _run_preflight(verbose="--doctor" in argv)

    _enable_windows_dpi_awareness()

    # 先 import 不依赖 Tkinter 的基础模块，这样出错时还能落日志
    from app.logging_setup import setup_logging
    from app.settings import SETTINGS

    setup_logging(verbose=bool(SETTINGS.get("verbose_log")))
    logger = setup_logging()

    if not _guard_python_version():
        return 1
    if not _guard_tkinter():
        return 2

    try:
        from app.ui.app import MainWindow, create_root
    except Exception:
        _fatal("加载界面模块失败", traceback.format_exc())
        return 3

    try:
        root, dnd = create_root()
    except Exception:
        _fatal("无法创建窗口", traceback.format_exc())
        return 4

    _enable_macos_app_menu()

    # 未捕获异常统一落日志并弹窗，避免"程序突然消失"
    def report_callback_exception(exc_type, exc_value, exc_tb):
        detail = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        logger.error("界面异常\n%s", detail)
        try:
            from tkinter import messagebox

            messagebox.showerror(
                "程序出错",
                f"{exc_type.__name__}: {exc_value}\n\n"
                "详细信息已写入日志，可在「设置 → 打开日志目录」中查看。",
            )
        except Exception:
            pass

    root.report_callback_exception = report_callback_exception

    MainWindow(root, dnd_available=dnd)
    root.mainloop()
    return 0


def _fatal(title: str, detail: str) -> None:
    """启动阶段出错：优先用原生弹窗（不依赖 Tkinter），再退到控制台。"""
    try:
        from app.native_alert import error

        error(title, detail)
    except Exception:
        pass
    print(f"[{title}]\n{detail}", file=sys.stderr)


if __name__ == "__main__":
    sys.exit(main())
