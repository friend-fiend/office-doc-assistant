"""任务引擎：线程 + 队列的异步执行模型。

线程模型
--------
    主线程(Tkinter)                 工作线程(每个任务一个)
    ────────────────                ────────────────────────
    poll() 消费消息  ←── queue ──   执行 handler
    更新进度条/日志                  在检查点调用 token.check()
    处理密码请求  ──── event ───→    阻塞等待密码
    点击取消      ──── token ───→    抛 Cancelled 主动退出

为什么不用线程池：文件处理任务是"一个接一个"的串行语义（合并顺序、
重命名序号、输出目录），并发反而会破坏语义；真正需要并行的地方
（LibreOffice 分批、pdf2image 内部渲染）已在各自模块里处理。
"""
from __future__ import annotations

import os
import queue
import threading
import time
import traceback
from dataclasses import dataclass, field
from typing import Any

from .cancel import CancelToken
from .config import APP_NAME, ext_set
from .core import OPS, OpContext, get_op
from .deps import get as dep_get
from .errors import (
    BrokenFile,
    Cancelled,
    DependencyMissing,
    ODAError,
    ParamError,
    PasswordRequired,
    TaskReport,
    WrongPassword,
)
from .logging_setup import get_logger
from .paths import cleanup_dir, temp_dir

log = get_logger("engine")

# 消息类型
MSG_LOG = "log"
MSG_PROGRESS = "progress"
MSG_DONE = "done"
MSG_PASSWORD = "password"
MSG_STATE = "state"


@dataclass
class TaskRequest:
    op_key: str
    files: list[str] = field(default_factory=list)
    params: dict = field(default_factory=dict)
    # 输出根目录；为空则输出到第一个文件所在目录
    output_parent: str = ""
    # 物理动作闸门（详见 core.registry.OpContext 的说明）。
    # 默认 False：任何非界面来源（测试、脚本、定时任务）都无法真实打印。
    # 只有 ui/app.py 里用户点击「开始执行」并确认后的那一次请求会置 True。
    allow_real_print: bool = False
    dry_run: bool = False


class MessageBus:
    """工作线程 -> 主线程 的单向消息通道（含密码请求的双向握手）。

    密码握手有一个容易忽略的竞态：主线程弹出对话框后，用户可能一直不操作，
    而工作线程的 ``event.wait`` 已经超时并继续往下跑了。等用户终于点了确定，
    这个结果就落到了一个"已经没人等"的请求上 —— 界面上表现为
    "输了密码但毫无反应"。

    解决办法是给每个请求一个带 ``expired`` 标记的 box，
    超时后作废，``deliver_password`` 校验身份后再唤醒。
    """

    # 密码输入等待上限（秒）。超过则视为用户放弃，按"无密码"继续。
    password_timeout = 300

    def __init__(self) -> None:
        self.queue: "queue.Queue[tuple]" = queue.Queue()
        self._pending_password: dict | None = None
        self._lock = threading.Lock()

    # -------------------------------------------------- 发送（工作线程调用）
    def log(self, text: str, level: str = "info") -> None:
        self.queue.put((MSG_LOG, text, level))

    def progress(self, done: int, total: int, label: str = "") -> None:
        self.queue.put((MSG_PROGRESS, done, max(total, 1), label))

    def state(self, text: str) -> None:
        self.queue.put((MSG_STATE, text))

    def done(self, report: TaskReport) -> None:
        self.queue.put((MSG_DONE, report))

    def ask_password(self, path: str, confirm: bool = False) -> str | None:
        """向主线程请求密码并阻塞等待。

        主线程收到 MSG_PASSWORD 后弹对话框，再调用 ``deliver_password`` 唤醒。
        超时或用户取消都返回 None。
        """
        event = threading.Event()
        box: dict[str, Any] = {
            "value": None,
            "event": event,
            "path": path,
            "confirm": confirm,
            "expired": False,
        }
        with self._lock:
            self._pending_password = box
        self.queue.put((MSG_PASSWORD, path, confirm, box))
        if not event.wait(timeout=self.password_timeout):
            # 作废：主线程后续若弹出对话框也不会再影响任何人
            with self._lock:
                box["expired"] = True
                if self._pending_password is box:
                    self._pending_password = None
            return None
        return box.get("value")

    # -------------------------------------------------- 接收（主线程调用）
    def poll(self, limit: int = 200) -> list[tuple]:
        out: list[tuple] = []
        for _ in range(limit):
            try:
                out.append(self.queue.get_nowait())
            except queue.Empty:
                break
        return out

    def deliver_password(self, value: str | None, box: dict | None = None) -> None:
        """把用户输入的密码交给等待中的工作线程。

        :param box: 收到 MSG_PASSWORD 时拿到的那个 box。传入后本方法会校验
            它是否仍是当前等待中的请求 —— 若期间已超时作废（或已经有别的
            请求取代了它），则**静默忽略**，避免唤醒错误的等待者。
        """
        with self._lock:
            current = self._pending_password
            if box is not None and current is not box:
                return  # 已超时作废，或已被新请求取代
            self._pending_password = None
        if current:
            current["value"] = value
            current["event"].set()


class TaskRunner(threading.Thread):
    """执行单个任务的工作线程。"""

    def __init__(self, request: TaskRequest, bus: MessageBus, token: CancelToken):
        super().__init__(daemon=True, name=f"oda-task-{request.op_key}")
        self.request = request
        self.bus = bus
        self.token = token

    # -------------------------------------------------- 前置校验
    def _validate(self) -> tuple[list[str], dict]:
        op = get_op(self.request.op_key)
        files = [os.path.abspath(f) for f in self.request.files]

        missing = [f for f in files if not os.path.exists(f)]
        if missing:
            raise ParamError("以下文件不存在：" + "、".join(os.path.basename(m) for m in missing[:5]))

        allowed = op.accepts_ext()
        if allowed:
            bad = [f for f in files if os.path.splitext(f)[1].lower() not in allowed]
            if bad:
                raise ParamError(
                    f"「{op.title}」只接受 {op.accepts_label()}，"
                    f"但列表中有 {len(bad)} 个其他类型文件："
                    + "、".join(os.path.basename(b) for b in bad[:3])
                )

        if op.mode == "collection" and len(files) < 2:
            raise ParamError(f"「{op.title}」至少需要 2 个文件")
        if op.mode == "pair" and len(files) != 2:
            raise ParamError(f"「{op.title}」需要恰好 2 个文件，当前 {len(files)} 个")
        if not files:
            raise ParamError("文件列表为空，请先添加文件")

        for key in op.requires:
            st = dep_get(key)
            if not st.available:
                raise DependencyMissing(key, f"「{op.title}」需要 {st.spec.title}，当前未检测到")

        return files, dict(self.request.params)

    # -------------------------------------------------- 线程主体
    def run(self) -> None:
        start = time.time()
        op = get_op(self.request.op_key)
        report = TaskReport(op_title=op.title)
        tmp = ""
        try:
            files, params = self._validate()
            report.input_count = len(files)

            parent = self.request.output_parent or os.path.dirname(files[0])
            if not os.path.isdir(parent):
                parent = os.path.dirname(files[0])

            from .paths import unique_dir

            base_name = op.title
            report.output_dir = unique_dir(parent, base_name)

            tmp = temp_dir()
            ctx = OpContext(self.bus, self.token, op, report.output_dir, report, temp_dir=tmp,
                            allow_real_print=self.request.allow_real_print,
                            dry_run=self.request.dry_run)
            ctx.set_total(len(files))
            ctx.log(f"开始执行：{op.title}")
            ctx.log(f"输入 {len(files)} 个文件，输出目录：{report.output_dir}")

            op.handler(ctx, files, params)

            report.duration = time.time() - start
            ctx.log(f"任务结束：{report.summary_text()}")
            self.bus.done(report)

        except Cancelled:
            report.cancelled = True
            report.duration = time.time() - start
            self.bus.log("任务已取消。", "warn")
            self.bus.done(report)

        except (DependencyMissing, ParamError, BrokenFile, WrongPassword, PasswordRequired) as exc:
            report.duration = time.time() - start
            self.bus.log(f"{exc}", "error")
            report.add_failure("(任务)", str(exc), "error")
            self.bus.done(report)

        except ODAError as exc:
            report.duration = time.time() - start
            self.bus.log(f"{exc}", "error")
            report.add_failure("(任务)", str(exc), "error")
            self.bus.done(report)

        except Exception as exc:  # 兜底：任何未预期异常都要落日志，不能让线程静默死掉
            report.duration = time.time() - start
            detail = traceback.format_exc()
            log.error("任务异常终止\n%s", detail)
            self.bus.log(f"发生未预期的错误：{exc}", "error")
            report.add_failure("(任务)", f"{type(exc).__name__}: {exc}", "error")
            self.bus.done(report)

        finally:
            if tmp:
                cleanup_dir(tmp)


class TaskManager:
    """界面持有它来启动 / 取消任务。"""

    def __init__(self) -> None:
        self.bus = MessageBus()
        self.token = CancelToken()
        self._runner: TaskRunner | None = None
        self.current_op_key: str = ""

    @property
    def running(self) -> bool:
        return self._runner is not None and self._runner.is_alive()

    def start(self, request: TaskRequest) -> None:
        if self.running:
            raise RuntimeError("已有任务正在执行")
        self.token = CancelToken()
        self.current_op_key = request.op_key
        self._runner = TaskRunner(request, self.bus, self.token)
        self._runner.start()

    def cancel(self) -> None:
        self.token.cancel()
        self.bus.log("已发出取消请求，正在等待当前步骤结束 …", "warn")

    def join(self, timeout: float | None = None) -> None:
        if self._runner:
            self._runner.join(timeout)

    # 便捷校验（主线程调用，用于按钮可用性提示）
    @staticmethod
    def precheck(op_key: str, files: list[str]) -> str:
        """返回空字符串表示可执行，否则返回不可执行的原因。"""
        op = OPS.get(op_key)
        if op is None:
            return "未知操作"
        if not files:
            return "请先添加文件"
        allowed = op.accepts_ext()
        if allowed:
            bad = [f for f in files if os.path.splitext(f)[1].lower() not in allowed]
            if bad:
                return f"有 {len(bad)} 个文件类型不符（需要 {op.accepts_label()}）"
        if op.mode == "collection" and len(files) < 2:
            return "至少需要 2 个文件"
        if op.mode == "pair" and len(files) != 2:
            return f"需要恰好 2 个文件（当前 {len(files)} 个）"
        for key in op.requires:
            st = dep_get(key)
            if not st.available:
                return f"缺少 {st.spec.title}"
        return ""


def describe_extensions() -> str:
    groups = ("pdf", "office", "image", "archive")
    return "、".join(f"{g}:{'/'.join(sorted(ext_set(g)))}" for g in groups)


_ = APP_NAME  # 保持导入一致性
