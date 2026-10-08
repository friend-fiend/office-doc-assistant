"""统一异常与任务结果载体。"""
from __future__ import annotations

from dataclasses import dataclass, field


class ODAError(Exception):
    """应用内所有可预期异常的基类。"""


class Cancelled(ODAError):
    """用户主动取消任务。"""


class DependencyMissing(ODAError):
    """缺少外部依赖（Poppler / LibreOffice / Ghostscript）。"""

    def __init__(self, tool_key: str, message: str = ""):
        self.tool_key = tool_key
        super().__init__(message or f"缺少外部依赖：{tool_key}")


class PasswordRequired(ODAError):
    """PDF 已加密，需要密码。"""

    def __init__(self, path: str, message: str = ""):
        self.path = path
        super().__init__(message or f"文件已加密，需要密码：{path}")


class WrongPassword(ODAError):
    """密码错误。"""


class BrokenFile(ODAError):
    """文件损坏或格式不受支持。"""


class ParamError(ODAError):
    """参数校验失败。"""


class PrintNotAuthorized(ODAError):
    """未经人工确认就试图真实打印。

    打印是**物理世界里不可逆**的操作（出纸、耗材、可能被同事看到内容），
    所以它和"写文件"不是一个量级：写错了可以删，纸打出来收不回去。

    因此真实打印必须由人在界面上明确确认，授权通过
    ``OpContext.allow_real_print`` 显式传递。测试脚本、体检脚本、
    定时任务等无人值守路径拿不到这个授权，会被这里拦住。
    """


@dataclass
class Failure:
    """单条失败记录。"""

    path: str
    reason: str
    category: str = "error"  # error | skipped | warning


@dataclass
class TaskReport:
    """一次任务的汇总结果。"""

    op_title: str = ""
    output_dir: str = ""
    ok: int = 0
    failed: int = 0
    skipped: int = 0
    cancelled: bool = False
    duration: float = 0.0
    input_count: int = 0
    outputs: list[str] = field(default_factory=list)
    failures: list[Failure] = field(default_factory=list)
    metrics: dict = field(default_factory=dict)

    @property
    def total(self) -> int:
        return self.ok + self.failed + self.skipped

    def add_failure(self, path: str, reason: str, category: str = "error") -> None:
        self.failures.append(Failure(path=path, reason=reason, category=category))
        if category == "error":
            self.failed += 1
        else:
            self.skipped += 1

    def summary_text(self) -> str:
        parts = [f"成功 {self.ok}", f"失败 {self.failed}"]
        if self.skipped:
            parts.append(f"跳过 {self.skipped}")
        tail = "（已取消）" if self.cancelled else ""
        return "，".join(parts) + tail

    def detail_text(self) -> str:
        lines = [f"任务：{self.op_title}", f"输出目录：{self.output_dir}",
                 f"结果：{self.summary_text()}", f"耗时：{self.duration:.2f} 秒"]
        for k, v in self.metrics.items():
            lines.append(f"{k}：{v}")
        if self.failures:
            lines.append("—— 明细 ——")
            for f in self.failures:
                tag = {"error": "失败", "skipped": "跳过", "warning": "警告"}.get(f.category, f.category)
                lines.append(f"[{tag}] {f.path} :: {f.reason}")
        return "\n".join(lines)
