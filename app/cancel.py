"""可取消令牌。

工作线程在每个可能耗时较长的位置调用 ``token.check()``，
一旦用户在界面上点了"取消"，这里抛出 Cancelled 中断任务。

设计为"协作式取消"：不强行 kill 线程（Python 无法安全终止线程），
而是让任务在自己的检查点主动退出，保证临时文件能正常清理。
"""
from __future__ import annotations

import threading

from .errors import Cancelled


class CancelToken:
    def __init__(self) -> None:
        self._event = threading.Event()

    def cancel(self) -> None:
        self._event.set()

    @property
    def cancelled(self) -> bool:
        return self._event.is_set()

    def check(self) -> None:
        if self._event.is_set():
            raise Cancelled("任务已被用户取消")

    def reset(self) -> None:
        self._event.clear()
