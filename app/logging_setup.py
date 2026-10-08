"""日志：同时写入滚动文件 + 推送到界面日志区。

界面日志通过一个线程安全的回调总线转发，不直接操作 Tkinter 控件
（Tkinter 只能在主线程更新，见 ui/app.py 的队列泵）。
"""
from __future__ import annotations

import logging
import logging.handlers
import os
import sys
from pathlib import Path


def log_dir() -> Path:
    if sys.platform.startswith("win"):
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    elif sys.platform == "darwin":
        base = os.path.join(os.path.expanduser("~"), "Library", "Logs")
    else:
        base = os.environ.get("XDG_STATE_HOME") or os.path.join(os.path.expanduser("~"), ".local", "state")
    d = Path(base) / "OfficeDocAssistant"
    d.mkdir(parents=True, exist_ok=True)
    return d


class CallbackHandler(logging.Handler):
    """把日志记录转发给一个可调用对象（供 GUI 使用）。"""

    def __init__(self, callback, level=logging.INFO):
        super().__init__(level)
        self.callback = callback

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.callback(self.format(record), record.levelno)
        except Exception:
            pass


_LOGGER: logging.Logger | None = None


def setup_logging(verbose: bool = False, ui_callback=None) -> logging.Logger:
    global _LOGGER
    if _LOGGER is not None:
        return _LOGGER

    logger = logging.getLogger("oda")
    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")

    # 第三方库的告警噪音降级：
    # pypdf 对损坏文件会打出 "EOF marker not found" 等 WARNING，我们已经在
    # BrokenFile 分支给出更友好的中文提示，无需重复输出。
    logging.getLogger("pypdf").setLevel(logging.ERROR)
    logging.getLogger("PIL").setLevel(logging.ERROR)

    try:
        path = log_dir() / "app.log"
        fh = logging.handlers.RotatingFileHandler(
            path, maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8"
        )
        fh.setLevel(logging.DEBUG if verbose else logging.INFO)
        fh.setFormatter(fmt)
        logger.addHandler(fh)
    except Exception:
        pass

    if ui_callback is not None:
        uh = CallbackHandler(ui_callback)
        uh.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(uh)

    _LOGGER = logger
    return logger


def get_logger(name: str = "oda") -> logging.Logger:
    if name == "oda":
        return _LOGGER or setup_logging()
    return logging.getLogger(f"oda.{name}")


def attach_ui_handler(callback) -> CallbackHandler:
    """运行期挂上/更换界面日志接收器。"""
    logger = get_logger()
    for h in list(logger.handlers):
        if isinstance(h, CallbackHandler):
            logger.removeHandler(h)
    h = CallbackHandler(callback)
    h.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(h)
    return h
