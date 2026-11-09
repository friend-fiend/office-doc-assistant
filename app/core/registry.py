"""操作注册表：把每个功能声明成一条 OpDef。

界面完全由这张表驱动：
  - 左侧分类按钮      <- categories()
  - 中间文件类型过滤  <- OpDef.accepts
  - 右侧参数控件      <- OpDef.params（按 Param.kind 动态生成）
  - 按钮禁用与提示    <- OpDef.requires（外部依赖）
这样新增功能只需写一个 handler + 一条 OpDef，不必改界面代码。
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Any, Callable, Iterable

from ..config import EXT_GROUP_LABEL

# 界面左侧的分类
CATEGORIES: tuple[tuple[str, str], ...] = (
    ("pdf_pages", "PDF 页面操作"),
    ("convert", "格式转换"),
    ("batch", "批量文件"),
    ("invoice", "发票整理"),
    ("compress", "压缩加密"),
    ("advanced", "高级工具"),
)


@dataclass
class Param:
    """一个参数控件声明。"""

    key: str
    label: str
    kind: str = "entry"  # entry|int|float|choice|bool|text|password|range|csv|position|per_file_config
    default: Any = ""
    choices: tuple[tuple[str, str], ...] = ()  # (值, 显示文本)
    hint: str = ""
    width: int = 14
    # 仅当 other 参数等于该值时才显示（用于联动，例如压缩档位）
    visible_if: tuple[str, tuple[Any, ...]] | None = None
    min_value: float | None = None
    max_value: float | None = None
    # per_file_config 专用：轴定义（每个轴是一个 dict: key/label/choices/(widget=int)/default/min/max）
    axes: tuple[dict, ...] = ()


@dataclass
class OpDef:
    key: str
    title: str
    category: str
    handler: Callable | None = None
    # 接受的扩展名分组，见 config.ext_set
    accepts: tuple[str, ...] = ("any",)
    # 处理模式：
    #   collection —— 所有文件作为一组（PDF 合并、图片转 PDF、批量压缩）
    #   per_file   —— 每个文件独立处理（拆分、加页码、重命名 …）
    #   pair       —— 恰好两个文件（文档对比）
    mode: str = "per_file"
    params: tuple[Param, ...] = ()
    # 依赖的外部工具 key（deps.TOOL_SPECS）
    requires: tuple[str, ...] = ()
    # 执行前是否需要二次确认
    confirm: bool = False
    confirm_text: str = ""
    # 会向物理设备（打印机）提交作业的功能。这类操作不可撤销，
    # 界面必须把真实参数摆给用户看，且只有人工确认过的调用才允许执行。
    hardware: bool = False
    # 动态构造确认文案：接收 (files, params)，返回提示文本。
    # 用于"把即将发生的真实动作讲清楚"（打印机名、份数、预计张数）。
    confirm_builder: Callable[[list, dict], str] | None = None
    description: str = ""
    # 输出形式提示，用于结果弹窗
    output_kind: str = "folder"

    def __post_init__(self) -> None:
        # 会动物理设备的功能必须经过二次确认。
        # 否则"用户在界面上点一下"和"无人值守脚本调一次"就没有区别了，
        # ui 层那条授权闸门也就形同虚设。
        if self.hardware and not self.confirm:
            raise ValueError(
                f"OpDef {self.key!r}：hardware=True 的功能必须同时 confirm=True"
            )

    def accepts_ext(self) -> set[str]:
        from ..config import ext_set

        out: set[str] = set()
        for g in self.accepts:
            out |= ext_set(g)
        return out

    def accepts_label(self) -> str:
        return "、".join(EXT_GROUP_LABEL.get(g, g) for g in self.accepts)


OPS: dict[str, OpDef] = {}


def register(op: OpDef):
    """把 OpDef 注册进 OPS，并返回一个用于绑定 handler 的装饰器。

    用法::

        MY_OP = OpDef(key="x", ...)

        @register(MY_OP)
        def run_x(ctx, files, params):
            ...
    """

    def deco(fn: Callable) -> Callable:
        op.handler = fn
        if op.key in OPS:
            raise ValueError(f"操作 key 重复：{op.key}")
        OPS[op.key] = op
        return fn

    return deco


def get_op(key: str) -> OpDef:
    return OPS[key]


def categories() -> tuple[tuple[str, str], ...]:
    return CATEGORIES


def ops_of(category: str) -> list[OpDef]:
    return [op for op in OPS.values() if op.category == category]


# ------------------------------------------------------------------ 运行上下文
class OpContext:
    """交给 handler 的运行上下文：日志、进度、取消、结果统计。

    全部方法都是线程安全的（内部只往 queue 里放消息，界面在主线程消费）。

    安全阀：``allow_real_print`` / ``dry_run``
    ------------------------------------------
    这两个字段是"物理世界动作"的闸门，默认最保守：

    * ``allow_real_print=False``（默认）—— 任何 handler 想向真实打印机提交
      作业都会被 ``PrintNotAuthorized`` 拦住。只有界面里用户点击「开始执行」
      并确认过的那条路径才会把它置为 True。
    * ``dry_run=False`` —— 置 True 时打印类功能只渲染预览、不碰打印机。

    这样设计的理由：打印不可撤销。测试脚本、体检、定时任务一旦能打印，
    就会变成"用户没做任何操作，打印机自己开始出纸"，且无从追责。
    """

    def __init__(self, bus, token, op: OpDef, out_dir: str, report, temp_dir: str = "",
                 allow_real_print: bool = False, dry_run: bool = False):
        self._bus = bus
        self.token = token
        self.op = op
        self.out_dir = out_dir
        self.report = report
        self.temp_dir = temp_dir
        self._total = 0
        self._done = 0
        # 物理动作闸门 —— 默认关闭，必须由调用方显式打开
        self.allow_real_print = bool(allow_real_print)
        self.dry_run = bool(dry_run)

    # -------------------------------------------------- 基础
    def log(self, message: str, level: str = "info") -> None:
        self._bus.log(message, level)

    def progress(self, done: int, total: int, label: str = "") -> None:
        self._total = max(total, 1)
        self._done = done
        self._bus.progress(done, self._total, label)

    def step(self, label: str = "") -> None:
        """进度 +1。"""
        self._done += 1
        self._bus.progress(self._done, max(self._total, 1), label)

    def set_total(self, total: int) -> None:
        self._total = max(total, 1)
        self._bus.progress(self._done, self._total, "")

    def check(self) -> None:
        """取消检查点。耗时循环里务必调用。"""
        self.token.check()

    def cancelled(self) -> bool:
        return self.token.cancelled

    # -------------------------------------------------- 结果
    def ok(self, out_path: str = "", src: str = "") -> None:
        self.report.ok += 1
        if out_path:
            self.report.outputs.append(out_path)
        if src:
            self.log(f"✓ {os.path.basename(src)}" + (f" → {os.path.basename(out_path)}" if out_path else ""))

    def fail(self, src: str, reason: str) -> None:
        self.report.add_failure(src, reason, "error")
        self.log(f"✗ {os.path.basename(src)}：{reason}", "error")

    def skip(self, src: str, reason: str) -> None:
        self.report.add_failure(src, reason, "skipped")
        self.log(f"⊘ 跳过 {os.path.basename(src)}：{reason}", "warn")

    def warn(self, src: str, reason: str) -> None:
        self.report.add_failure(src, reason, "warning")
        self.log(f"! {os.path.basename(src)}：{reason}", "warn")

    def metric(self, key: str, value: Any) -> None:
        self.report.metrics[key] = value

    # -------------------------------------------------- 密码回调
    def ask_password(self, path: str, confirm: bool = False) -> str | None:
        """向界面请求密码；返回 None 表示用户取消。"""
        return self._bus.ask_password(path, confirm=confirm)


def params_to_dict(params: Iterable[Param], values: dict) -> dict:
    """补齐缺省值。"""
    out = dict(values)
    for p in params:
        if p.key not in out:
            out[p.key] = p.default
    return out
