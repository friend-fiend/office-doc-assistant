"""主窗口。

布局
----
    ┌──────────────────────────────────────────────────────────────┐
    │ 顶栏：标题 · 依赖状态 · 设置 · 关于                            │
    ├──────────┬───────────────────────────────────────────────────┤
    │ 左侧导航  │ 拖拽区                                            │
    │ 分类+功能 │ 文件列表（可拖动排序 / 删除 / 清空）                 │
    │          │ 参数区（由 OpDef.params 动态生成）                  │
    │          │ 操作栏：输出目录 · 开始执行 · 取消 · 进度            │
    │          │ 日志区                                            │
    └──────────┴───────────────────────────────────────────────────┘

线程约定（本文件最重要的一条规则）
------------------------------------
**只有两个方法可以碰控件：``_pump`` 和 ``_startup``。**

Tkinter 不是线程安全的。工作线程直接调 ``root.after()`` 会抛
``RuntimeError: main thread is not in main loop``；直接改控件则可能
**随机崩溃或静默卡死** —— 后者尤其难查。

因此所有跨线程来源都统一走队列，由主线程 ``after(80ms)`` 常驻消息泵消费：

    工作线程 / 依赖探测线程 / logging 回调
              │
              ├─ TaskManager.bus  ──┐
              └─ self._ui_queue   ──┤
                                    ▼
                        _pump()（主线程）
                          · 每 tick 处理量有上限 → 界面不会因日志洪水假死
                          · 进度消息只取最后一条 → 不做无意义的重复重绘
                          · 关窗后不再排程     → 回调不会打到已销毁的控件
                                    │
                                    ▼
                              更新控件

``_append`` 里还有一条"绊线"：一旦检测到非主线程直接写日志控件，
会自动改走队列并记录一条 error —— 把潜在崩溃变成可见的缺陷报告。
"""
from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox, ttk

from ..config import APP_NAME, VERSION
from ..core import categories, get_op, ops_of
from ..deps import detect_all
from ..engine import MSG_DONE, MSG_LOG, MSG_PASSWORD, MSG_PROGRESS, MSG_STATE, TaskManager, TaskRequest
from ..errors import TaskReport
from ..logging_setup import attach_ui_handler, get_logger
from ..paths import open_path
from ..settings import SETTINGS
from . import theme
from .filelist import FileList
from .parampanel import ParamPanel
from .settings_win import SettingsWindow

log = get_logger("ui")

LEVEL_COLOR = {
    "debug": theme.TEXT_SUB,
    "info": theme.TEXT,
    "warn": theme.WARN,
    "warning": theme.WARN,
    "error": theme.ERROR,
    "success": theme.SUCCESS,
}

# 消息泵节奏。80ms 是"看不出来延迟"与"不空转 CPU"之间的经验值。
PUMP_INTERVAL_MS = 80
# 单个 tick 最多处理多少条消息。必须设上限：一个日志量巨大的任务
# （比如 500 个文件各失败一次）会瞬间塞满队列，如果一次全画完，
# 界面会直接假死几十秒。剩下的留到下一个 tick，用户完全无感。
MAX_MSGS_PER_TICK = 400
# 积压超过这个数就提醒一次用户：可能是任务产出日志的速度超过了渲染速度
BACKLOG_WARN = 3000


def create_root() -> tuple[tk.Tk, bool]:
    """创建根窗口；若装了 tkinterdnd2 就用它，从而支持从资源管理器拖入文件。

    Tkinter 本身没有拖放能力，必须靠 tkdnd 这个 Tcl 扩展（tkinterdnd2 是它的
    Python 封装）。它是可选依赖：装不上时程序照常运行，只是拖拽区退化为
    「点击选择文件」，并在界面上明确告知用户。
    """
    try:
        from tkinterdnd2 import TkinterDnD

        root = TkinterDnD.Tk()
        return root, True
    except Exception:
        return tk.Tk(), False


class MainWindow:
    def __init__(self, root: tk.Tk, dnd_available: bool = False):
        self.root = root
        self.dnd_available = dnd_available
        self.manager = TaskManager()
        self.last_output_dir = ""
        self.current_op_key = ""
        self._busy = False
        self._last_report: TaskReport | None = None
        self._dep_status = {}
        # 关窗标志：置位后消息泵不再重新排程。
        # 不设这个标志的话，root.destroy() 之后那个已经排好的 after 回调
        # 仍会被 Tk 调用，然后在 __init__ 之外的控件上炸出
        # "TclError: invalid command name ..."。
        self._closing = False
        # 积压提醒只打一次，否则会刷屏把真正有用的日志冲掉
        self._backlog_warned = False
        # 非 Tk 线程 -> 主线程 的唯一通道。
        # 为什么不用 root.after()：Tkinter 不是线程安全的，从工作线程直接调用
        # after() 会抛 "RuntimeError: main thread is not in main loop"，
        # 或在某些平台上直接崩掉。所有跨线程更新都必须走这条队列。
        self._ui_queue: "queue.Queue[tuple]" = queue.Queue()

        root.title(f"{APP_NAME} v{VERSION}")
        root.minsize(1020, 700)
        geom = SETTINGS.get("window_geometry") or "1180x800"
        try:
            root.geometry(geom)
        except Exception:
            root.geometry("1180x800")

        theme.apply_theme(root)
        self.style = ttk.Style(root)

        attach_ui_handler(self._log_from_library)
        self._build()
        self._install_shortcuts()

        root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._startup)

    # ================================================================ 构建
    def _build(self) -> None:
        self._build_header()

        body = ttk.Frame(self.root)
        body.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.sidebar = self._build_sidebar(body)
        self.sidebar.pack(side="left", fill="y", padx=(0, 12))

        right = ttk.Frame(body)
        right.pack(side="left", fill="both", expand=True)

        self._build_dropzone(right)
        self.filelist = FileList(right, on_change=self._on_files_changed)
        self.filelist.pack(fill="both", expand=True, pady=(8, 8))

        self.param_panel = ParamPanel(right, on_change=self._on_param_changed,
                                      get_files=self.filelist.get_files)
        self.param_panel.configure(height=210)
        self.param_panel.pack(fill="x", pady=(0, 8))
        self.param_panel.pack_propagate(False)

        self._build_actionbar(right)
        self._build_log(right)

    # ---------------------------------------------------------------- 顶栏
    def _build_header(self) -> None:
        bar = ttk.Frame(self.root, padding=(14, 10, 14, 6))
        bar.pack(fill="x")

        ttk.Label(bar, text=APP_NAME, style="Title.TLabel").pack(side="left")
        ttk.Label(bar, text=f"v{VERSION}", style="Sub.TLabel").pack(side="left", padx=(8, 0))

        self.dep_frame = ttk.Frame(bar)
        self.dep_frame.pack(side="left", padx=(22, 0))
        self.dep_labels: dict[str, ttk.Label] = {}
        for name in ("poppler", "soffice", "ghostscript"):
            lbl = ttk.Label(self.dep_frame, text="检测中…", style="Sub.TLabel",
                            cursor="hand2")
            lbl.pack(side="left", padx=(0, 14))
            lbl.bind("<Button-1>", lambda e: self._open_settings())
            self.dep_labels[name] = lbl

        ttk.Button(bar, text="关于", style="Ghost.TButton",
                   command=self._show_about).pack(side="right")
        ttk.Button(bar, text="设置", style="Ghost.TButton",
                   command=self._open_settings).pack(side="right", padx=(0, 4))

        ttk.Separator(self.root).pack(fill="x")

    # ---------------------------------------------------------------- 左导航
    def _build_sidebar(self, parent) -> ttk.Frame:
        outer = ttk.Frame(parent, style="Card.TFrame", width=196)
        outer.pack_propagate(False)

        canvas = tk.Canvas(outer, width=180, highlightthickness=0, bg=theme.PANEL)
        vs = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vs.set)
        canvas.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")

        inner = ttk.Frame(canvas, style="Card.TFrame")
        canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        self.op_buttons: dict[str, ttk.Button] = {}

        ttk.Label(inner, text="功能导航", style="Card.TLabel",
                  font=theme.ui_font(11, True)).pack(anchor="w", padx=12, pady=(12, 6))

        for cat_key, cat_title in categories():
            ttk.Label(inner, text=cat_title, style="CardSub.TLabel",
                      font=theme.ui_font(9, True)).pack(anchor="w", padx=12, pady=(10, 2))
            for op in ops_of(cat_key):
                btn = ttk.Button(inner, text="  " + op.title, style="Op.TButton",
                                 command=lambda k=op.key: self._select_op(k))
                btn.pack(fill="x", padx=6, pady=1)
                self.op_buttons[op.key] = btn

        ttk.Frame(inner, style="Card.TFrame", height=10).pack()
        return outer

    # ---------------------------------------------------------------- 拖拽区
    def _build_dropzone(self, parent) -> None:
        self.drop_frame = tk.Frame(parent, bg=theme.PANEL, highlightthickness=1,
                                   highlightbackground=theme.BORDER,
                                   highlightcolor=theme.PRIMARY, height=64)
        self.drop_frame.pack(fill="x")
        self.drop_frame.pack_propagate(False)

        text = ("把文件或文件夹拖到这里，或点击选择"
                if self.dnd_available
                else "点击这里选择文件（安装 tkinterdnd2 可支持拖拽）")
        self.drop_label = tk.Label(
            self.drop_frame, text="⬇  " + text,
            bg=theme.PANEL, fg=theme.TEXT_SUB, font=theme.ui_font(10),
            cursor="hand2",
        )
        self.drop_label.pack(expand=True)
        self.drop_label.bind("<Button-1>", lambda e: self.filelist.add_files())

        if self.dnd_available:
            try:
                from tkinterdnd2 import DND_FILES

                for widget in (self.drop_frame, self.drop_label):
                    widget.drop_target_register(DND_FILES)
                    widget.dnd_bind("<<Drop>>", self._on_drop)
                    widget.dnd_bind("<<DropEnter>>", lambda e: self._drop_hover(True))
                    widget.dnd_bind("<<DropLeave>>", lambda e: self._drop_hover(False))
            except Exception as exc:
                log.warning("拖放注册失败：%s", exc)
                self.dnd_available = False

    def _drop_hover(self, on: bool) -> None:
        bg = theme.DROPI_BG if on else theme.PANEL
        self.drop_frame.configure(bg=bg, highlightbackground=theme.PRIMARY if on else theme.BORDER)
        self.drop_label.configure(bg=bg)

    def _on_drop(self, event) -> None:
        self._drop_hover(False)
        try:
            paths = list(self.root.tk.splitlist(event.data))
        except Exception:
            paths = str(event.data).split()
        files: list[str] = []
        for p in paths:
            if os.path.isdir(p):
                for root_, _dirs, names in os.walk(p):
                    for n in sorted(names):
                        if n.startswith(".") or n.startswith("~$"):
                            continue
                        files.append(os.path.join(root_, n))
            elif os.path.isfile(p):
                files.append(p)
        if files:
            self.filelist.add_files(files)
            self._log(f"已从拖拽添加 {len(files)} 个文件", "info")

    # ---------------------------------------------------------------- 操作栏
    def _build_actionbar(self, parent) -> None:
        bar = ttk.Frame(parent)
        bar.pack(fill="x", pady=(0, 6))

        row1 = ttk.Frame(bar)
        row1.pack(fill="x")

        ttk.Label(row1, text="输出到：", style="Sub.TLabel").pack(side="left")
        self.out_var = tk.StringVar(value=SETTINGS.get("last_output_dir") or "")
        self.out_entry = ttk.Entry(row1, textvariable=self.out_var, width=44)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(4, 6))
        ttk.Button(row1, text="选择…", command=self._choose_output).pack(side="left")
        ttk.Button(row1, text="重置为源目录", style="Ghost.TButton",
                   command=lambda: self.out_var.set("")).pack(side="left", padx=(4, 0))

        row2 = ttk.Frame(bar)
        row2.pack(fill="x", pady=(8, 4))

        self.run_btn = ttk.Button(row2, text="开始执行", style="Accent.TButton",
                                  command=self._start_task, state="disabled")
        self.run_btn.pack(side="left")
        self.cancel_btn = ttk.Button(row2, text="取消", command=self._cancel_task,
                                     state="disabled")
        self.cancel_btn.pack(side="left", padx=(8, 0))

        self.extra_btn = ttk.Button(row2, text="预览重命名", command=self._preview_rename)
        # 仅在重命名功能下 pack（见 _select_op）

        self.open_out_btn = ttk.Button(row2, text="打开输出文件夹",
                                       command=self._open_last_output)
        self.open_out_btn.pack(side="left", padx=(8, 0))
        self.open_out_btn.state(["disabled"])

        self.hint_label = ttk.Label(row2, text="请先选择一个功能", style="Sub.TLabel")
        self.hint_label.pack(side="left", padx=(12, 0))

        self.progress = ttk.Progressbar(bar, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(6, 0))
        self.progress_label = ttk.Label(bar, text="就绪", style="Sub.TLabel")
        self.progress_label.pack(anchor="w", pady=(2, 0))

    # ---------------------------------------------------------------- 日志区
    def _build_log(self, parent) -> None:
        head = ttk.Frame(parent)
        head.pack(fill="x", pady=(6, 2))
        ttk.Label(head, text="执行日志", style="Section.TLabel").pack(side="left")
        ttk.Button(head, text="保存日志", style="Ghost.TButton",
                   command=self._save_log).pack(side="right")
        ttk.Button(head, text="清空", style="Ghost.TButton",
                   command=self._clear_log).pack(side="right", padx=(0, 4))

        wrap = ttk.Frame(parent, style="Card.TFrame")
        wrap.pack(fill="both", expand=True)
        self.log_text = tk.Text(wrap, height=8, wrap="word", relief="flat",
                                bg=theme.PANEL, fg=theme.TEXT,
                                font=theme.mono_font(9), padx=10, pady=8,
                                insertbackground=theme.TEXT, state="disabled")
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=vs.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")

        for level, color in LEVEL_COLOR.items():
            self.log_text.tag_configure(level, foreground=color)
        self.log_text.tag_configure("tag_bold", font=theme.mono_font(9))

    def _install_shortcuts(self) -> None:
        self.root.bind("<Control-o>", lambda e: self.filelist.add_files())
        self.root.bind("<F5>", lambda e: self._start_task())
        self.root.bind("<Escape>", lambda e: self._cancel_task() if self._busy else None)
        self.root.bind("<Control-comma>", lambda e: self._open_settings())

    # ================================================================ 启动
    def _startup(self) -> None:
        self._log(f"{APP_NAME} 已启动。选择左侧功能 → 添加文件 → 设置参数 → 开始执行。", "info")
        if not self.dnd_available:
            self._log("提示：当前未安装 tkinterdnd2，拖拽功能不可用（可用「添加文件」按钮代替）。"
                      "如需拖拽：pip install tkinterdnd2", "warn")

        # 外部依赖探测会启动子进程，放到后台线程，避免拖慢启动
        def worker():
            try:
                status = detect_all(force=True)
            except Exception as exc:
                log.error("依赖检测失败：%s", exc)
                status = {}
            # 只往队列放数据，由主线程的 _pump 消费并更新控件
            self._ui_queue.put(("deps", status, True))

        threading.Thread(target=worker, daemon=True, name="oda-dep-detect").start()

        # 默认选中第一个功能
        first = categories()[0][0]
        ops = ops_of(first)
        if ops:
            self._select_op(ops[0].key)

        # 消息泵常驻运行（不随任务结束而停止），负责消费两个线程间队列
        self.root.after(80, self._pump)

    def _apply_dep_status(self, status: dict, announce: bool = False) -> None:
        self._dep_status = status or {}
        for key, lbl in self.dep_labels.items():
            st = self._dep_status.get(key)
            if st is None:
                lbl.configure(text=f"{key}：未知", foreground=theme.TEXT_SUB)
                continue
            if st.available:
                ver = f" v{st.version}" if st.version else ""
                lbl.configure(text=f"✓ {st.spec.title.split('（')[0]}{ver}",
                              foreground=theme.SUCCESS)
            else:
                lbl.configure(text=f"✗ {st.spec.title.split('（')[0]} 缺失",
                              foreground=theme.ERROR)
        self._refresh_buttons()
        if announce:
            missing = [st for st in self._dep_status.values() if not st.available]
            if missing:
                self._log("以下外部程序未检测到，相关功能已禁用（点击顶栏名称或「设置」可指定路径）：",
                          "warn")
                for st in missing:
                    self._log(f"  · {st.spec.title} —— 影响：{'、'.join(st.spec.affects)}", "warn")
                    self._log(f"    {st.spec.help_text.splitlines()[0]}", "warn")
            else:
                self._log("外部依赖检测完成，全部就绪。", "success")
            self._style_dep_hint()

    def _style_dep_hint(self) -> None:
        for key, lbl in self.dep_labels.items():
            st = self._dep_status.get(key)
            if st and not st.available:
                lbl.configure(cursor="hand2")

    # ================================================================ 功能选择
    def _select_op(self, op_key: str) -> None:
        if self._busy:
            return
        op = get_op(op_key)
        self.current_op_key = op_key

        for k, btn in self.op_buttons.items():
            btn.configure(style="OpSel.TButton" if k == op_key else "Op.TButton")

        self.filelist.set_filter(op.accepts, op.accepts_label())
        self.param_panel.build(op)

        # 重命名功能特有的预览按钮
        self.extra_btn.pack_forget()
        if op.key == "batch_rename":
            self.extra_btn.pack(side="left", padx=(8, 0))

        self._refresh_buttons()
        self._log(f"已选择功能：{op.title}（接受 {op.accepts_label()}）", "debug")

    def _on_files_changed(self) -> None:
        self.param_panel.refresh_files()
        self._refresh_buttons()

    @staticmethod
    def _is_dry_run(params: dict) -> bool:
        """本次是否演练（不动物理设备）。

        取值来源与 core.print_ops._is_dry_run 保持一致；界面层提前知道是为了
        在启动任务时就把 dry_run 标到请求上，便于日志与调试。
        """
        if str(os.environ.get("ODA_PRINT_DRYRUN", "")).strip().lower() in ("1", "true", "yes", "on"):
            return True
        return bool(params.get("dry_run"))

    def _on_param_changed(self) -> None:
        self._refresh_buttons()

    def _refresh_buttons(self) -> None:
        if not self.current_op_key:
            self.run_btn.state(["disabled"])
            self.hint_label.configure(text="请先选择一个功能")
            return
        if self._busy:
            self.run_btn.state(["disabled"])
            self.cancel_btn.state(["!disabled"])
            return
        self.cancel_btn.state(["disabled"])

        files = self.filelist.get_files()
        reason = TaskManager.precheck(self.current_op_key, files)
        if reason:
            self.run_btn.state(["disabled"])
            self.hint_label.configure(text=reason, foreground=theme.WARN)
        else:
            self.run_btn.state(["!disabled"])
            op = get_op(self.current_op_key)
            self.hint_label.configure(
                text=f"将对 {len(files)} 个文件执行「{op.title}」", foreground=theme.TEXT_SUB)

    # ================================================================ 执行
    def _start_task(self) -> None:
        if self._busy or not self.current_op_key:
            return
        op = get_op(self.current_op_key)
        files = self.filelist.get_files()

        reason = TaskManager.precheck(op.key, files)
        if reason:
            messagebox.showwarning("无法执行", reason, parent=self.root)
            return

        err = self.param_panel.validate()
        if err:
            messagebox.showwarning("参数有误", err, parent=self.root)
            return

        params = self.param_panel.get_values()

        # 写操作二次确认
        if op.confirm:
            # 有 confirm_builder 的功能（目前是打印）会把"即将发生的真实动作"
            # 摊开给用户看：哪台打印机、多少份、大概多少张纸。
            # 打印不可撤销，值得多说两句。
            text = op.confirm_text or f"将对 {len(files)} 个文件执行「{op.title}」，确认继续吗？"
            if op.confirm_builder is not None:
                try:
                    text = op.confirm_builder(files, params)
                except Exception as exc:
                    log.error("生成确认文案失败，退回静态文案：%s", exc)
            if not messagebox.askyesno("请确认", text, parent=self.root):
                return
        if op.key == "batch_rename" and params.get("apply_mode") == "inplace":
            if not messagebox.askyesno(
                "危险操作确认",
                f"「原地重命名」会直接修改这 {len(files)} 个文件在磁盘上的名称。\n\n"
                "程序会在输出文件夹生成回滚清单，但重命名本身不可撤销。\n\n确认继续吗？",
                icon="warning", parent=self.root,
            ):
                return

        output_parent = self.out_var.get().strip()
        if output_parent and not os.path.isdir(output_parent):
            messagebox.showwarning("输出目录无效", f"目录不存在：\n{output_parent}", parent=self.root)
            return

        SETTINGS.set("last_output_dir", output_parent)
        SETTINGS.save()

        self._busy = True
        self._last_report = None
        self.progress.configure(value=0, maximum=100)
        self.progress_label.configure(text=f"正在执行：{op.title}")
        self._clear_log()
        self._log(f"=== 开始执行：{op.title} ===", "info")
        self._log(f"文件数：{len(files)}；输出根目录：{output_parent or '（各文件所在目录）'}", "info")
        self._refresh_buttons()

        try:
            self.manager.start(TaskRequest(
                op_key=op.key, files=files, params=params, output_parent=output_parent,
                # ↓ 物理动作授权。执行到这一行意味着：用户在界面上点了
                # 「开始执行」，而且上面每一个确认框都点了"是"。
                # 这是全程序**唯一**能开出打印授权的路径 —— 测试脚本、
                # 批处理、定时任务走的都是 TaskRequest 默认值 False。
                allow_real_print=True,
                dry_run=self._is_dry_run(params),
            ))
        except Exception as exc:
            self._busy = False
            messagebox.showerror("启动失败", str(exc), parent=self.root)
            self._refresh_buttons()
            return

    def _cancel_task(self) -> None:
        if not self._busy:
            return
        if not messagebox.askyesno("确认取消", "确定要中断当前任务吗？\n已完成的文件会保留。",
                                   parent=self.root):
            return
        self.manager.cancel()
        self.cancel_btn.state(["disabled"])
        self.progress_label.configure(text="正在取消…")

    # ================================================================ 消息泵
    def _pump(self) -> None:
        """主线程唯一的控件更新入口。

        消费两个来源：
          · manager.bus  —— 任务线程的日志/进度/密码请求/完成通知
          · _ui_queue    —— 依赖探测线程与日志库的跨线程消息
        结束后重新排程（关窗后不再排程）。

        三条硬性约束，缺一个都会在真实使用中出问题：
          1. 只有本方法能碰控件 —— 其它任何跨线程路径都必须先入队；
          2. 每 tick 处理量必须有上限 —— 否则日志洪水会把界面拖死；
          3. 关窗后必须停止排程 —— 否则回调落在已销毁的控件上。
        """
        if self._closing:
            return
        try:
            self._consume_bus()
            self._consume_ui_queue()
        except Exception:
            log.error("消息泵异常\n%s", traceback.format_exc())
        finally:
            if not self._closing:
                try:
                    self.root.after(PUMP_INTERVAL_MS, self._pump)
                except Exception:
                    # 窗口已在销毁过程中，正常现象
                    self._closing = True

    def _warn_backlog(self, depth: int) -> None:
        if self._backlog_warned or depth <= BACKLOG_WARN:
            return
        self._backlog_warned = True
        self._log(f"消息积压较多（约 {depth} 条），日志显示会略有滞后，任务本身不受影响。", "warn")

    def _consume_bus(self) -> None:
        """消费任务线程的消息。"""
        msgs = self.manager.bus.poll(limit=MAX_MSGS_PER_TICK)
        backlog = self.manager.bus.queue.qsize()
        if backlog:
            self._warn_backlog(backlog)

        # 进度消息只保留最后一条：一个 tick 内反复重画进度条毫无意义，
        # 纯粹浪费主线程时间，而进度条本来就只反映"最新状态"。
        last_progress = None

        for msg in msgs:
            kind = msg[0]
            if kind == MSG_LOG:
                self._log(msg[1], msg[2])
            elif kind == MSG_PROGRESS:
                last_progress = msg
            elif kind == MSG_STATE:
                self.progress_label.configure(text=msg[1])
            elif kind == MSG_PASSWORD:
                self._handle_password_request(msg)
            elif kind == MSG_DONE:
                self._on_done(msg[1])

        if last_progress is not None:
            _k, done, total, label = last_progress
            pct = (done / total * 100) if total else 0
            self.progress.configure(value=pct, maximum=100)
            text = f"{done}/{total}"
            if label:
                text += f"  {label}"
            self.progress_label.configure(text=text)

    def _consume_ui_queue(self) -> None:
        """消费依赖探测线程与 logging 回调的消息。"""
        for msg in _drain(self._ui_queue, limit=MAX_MSGS_PER_TICK):
            kind = msg[0]
            if kind == "log":
                self._append(msg[1], msg[2])
            elif kind == "deps":
                self._apply_dep_status(msg[1], announce=bool(msg[2]))

    def _handle_password_request(self, msg) -> None:
        _kind, path, confirm, box = msg
        # 工作线程的等待有超时。如果已经超时作废，再弹框就成了"孤儿对话框"：
        # 用户输完密码却没有任何反应，非常困惑。这里直接跳过。
        if box.get("expired"):
            return
        title = "密码错误，请重新输入" if confirm else "该 PDF 已加密"
        prompt = (f"文件：{os.path.basename(path)}\n\n请输入打开密码："
                  if not confirm else
                  f"文件：{os.path.basename(path)}\n\n密码不正确，请重新输入：")
        value = _ask_string(self.root, title, prompt)
        # 带上 box 身份，若期间已超时则不会被误唤醒
        self.manager.bus.deliver_password(value, box=box)
        if value is None:
            self._log(f"用户取消了密码输入：{os.path.basename(path)}", "warn")

    def _on_done(self, report: TaskReport) -> None:
        self._busy = False
        self._last_report = report
        self.progress.configure(value=100 if not report.cancelled else self.progress["value"])
        self.progress_label.configure(text="已完成" if not report.cancelled else "已取消")
        self._refresh_buttons()

        if report.output_dir:
            self.last_output_dir = report.output_dir
            self.open_out_btn.state(["!disabled"])

        level = "error" if report.failed else "success"
        if report.cancelled:
            level = "warn"
        self._log(f"=== {report.summary_text()}（耗时 {report.duration:.1f} 秒）===", level)
        for k, v in report.metrics.items():
            self._log(f"  · {k}：{v}", "info")

        if SETTINGS.get("open_output_when_done") and report.output_dir and not report.cancelled:
            try:
                open_path(report.output_dir)
            except Exception:
                pass

        _ResultDialog(self.root, report, self._open_last_output)

    # ================================================================ 重命名预览
    def _preview_rename(self) -> None:
        from ..core.rename_ops import build_preview

        files = self.filelist.get_files()
        if not files:
            messagebox.showinfo("提示", "请先添加文件。", parent=self.root)
            return
        err = self.param_panel.validate()
        if err:
            messagebox.showwarning("参数有误", err, parent=self.root)
            return
        rows = build_preview(files, self.param_panel.get_values())
        _RenamePreviewDialog(self.root, rows)

    # ================================================================ 日志
    def _log(self, text: str, level: str = "info") -> None:
        self._append(text, level)

    def _log_from_library(self, text: str, levelno: int) -> None:
        """标准 logging 的回调 —— 可能来自工作线程，只能入队，不能碰控件。"""
        level = {10: "debug", 20: "info", 30: "warn", 40: "error"}.get(levelno, "info")
        try:
            self._ui_queue.put(("log", text, level))
        except Exception:
            pass

    def _append(self, text: str, level: str) -> None:
        # --- 线程安全绊线 ---------------------------------------------------
        # 这是唯一真正写入日志控件的函数，所以把检查放在这里。
        # 如果哪天有人新增了一条"从工作线程直接调 _log()"的路径，
        # Tkinter 的表现是**随机崩溃或随机卡死** —— 最难排查的那类问题。
        # 与其等着线上炸，不如在这里拦住并自动改走队列。
        if threading.current_thread() is not threading.main_thread():
            log.error(
                "检测到非主线程直接更新控件，已自动改走消息队列。"
                "这是代码缺陷，请修复调用方（不要绕过 _pump 消息泵）。"
            )
            try:
                self._ui_queue.put(("log", text, level))
            except Exception:
                pass
            return
        # -------------------------------------------------------------------
        try:
            self.log_text.configure(state="normal")
            self.log_text.insert("end", text.rstrip() + "\n", LEVEL_COLOR.get(level, "info"))
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        except Exception:
            pass

    def _clear_log(self) -> None:
        try:
            self.log_text.configure(state="normal")
            self.log_text.delete("1.0", "end")
            self.log_text.configure(state="disabled")
        except Exception:
            pass

    def _save_log(self) -> None:
        content = self.log_text.get("1.0", "end").strip()
        if not content:
            messagebox.showinfo("提示", "日志为空。", parent=self.root)
            return
        path = filedialog.asksaveasfilename(
            title="保存日志", defaultextension=".txt",
            initialfile="执行日志.txt", parent=self.root,
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(content)
            messagebox.showinfo("已保存", f"日志已保存到：\n{path}", parent=self.root)
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc), parent=self.root)

    # ================================================================ 其它
    def _choose_output(self) -> None:
        init = self.out_var.get() or SETTINGS.get("last_output_dir") or os.path.expanduser("~")
        d = filedialog.askdirectory(title="选择输出根目录（程序会在其中新建输出子文件夹）",
                                    initialdir=init)
        if d:
            self.out_var.set(d)

    def _open_last_output(self) -> None:
        target = self.last_output_dir
        if not target:
            report = self._last_report
            target = report.output_dir if report else ""
        if not target or not os.path.isdir(target):
            messagebox.showinfo("提示", "还没有可打开的输出目录。", parent=self.root)
            return
        try:
            open_path(target)
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc), parent=self.root)

    def _open_settings(self) -> None:
        SettingsWindow(self.root, on_saved=self._apply_dep_status)

    def _show_about(self) -> None:
        from ..core import OPS

        lines = [f"{APP_NAME}  v{VERSION}", "",
                 "纯 Python 桌面工具，用于批量处理 PDF 与常见办公文档。", "",
                 f"已注册 {len(OPS)} 项功能。",
                 "不含 PDF 转 Word / PDF 转 Excel / OCR（按需求明确排除）。", "",
                 "所有输出写入新文件夹，不覆盖原文件。"]
        messagebox.showinfo("关于", "\n".join(lines), parent=self.root)

    def _on_close(self) -> None:
        """关窗。

        不能直接 destroy()：工作线程可能正在写输出文件，
        进程一退就会留下**残缺的半成品文件**，而用户会以为处理成功了。
        所以流程是「请求取消 → 给它一点时间收尾 → 真超时了再问用户」。
        """
        if self._busy:
            if not messagebox.askyesno("任务进行中", "仍有任务在执行，确定退出吗？",
                                       parent=self.root):
                return
            self.manager.cancel()
            self._log("正在等待当前步骤结束（需要清理临时文件）…", "warn")
            self.progress_label.configure(text="正在取消…")
            # 让"正在取消"这句提示先真正画出来，否则用户看不到任何反馈
            try:
                self.root.update_idletasks()
            except Exception:
                pass
            # 协作式取消：引擎在检查点抛 Cancelled。正常情况下秒级完成。
            self.manager.join(timeout=5.0)
            if self.manager.running:
                if not messagebox.askyesno(
                    "任务尚未停止",
                    "任务还未响应取消请求（可能正在处理一个很大的文件或等待外部程序）。\n\n"
                    "现在强制退出，可能会留下不完整的输出文件。\n"
                    "建议选「否」，再稍等片刻。\n\n仍要强制退出吗？",
                    parent=self.root,
                ):
                    self.progress_label.configure(text="仍在取消…（可再次关闭窗口）")
                    return

        # 唤醒可能仍在等密码的工作线程，否则它会一直阻塞到 5 分钟超时才释放
        try:
            self.manager.bus.deliver_password(None)
        except Exception:
            pass

        # 先把标志置位，再 destroy —— 顺序反了的话，
        # 已经排队的那次 after 回调仍会打到已销毁的控件上。
        self._closing = True
        try:
            SETTINGS.set("window_geometry", self.root.winfo_geometry())
            SETTINGS.save()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass


# ==================================================================== 辅助对话框
def _drain(q: "queue.Queue", limit: int = MAX_MSGS_PER_TICK) -> list[tuple]:
    """非阻塞取队列，最多取 ``limit`` 条。

    设上限是为了保护界面响应性：剩下的消息会在下一个 tick 被取走，
    而一次性处理上万条会让主线程长时间无法重绘 —— 表现就是"程序卡死了"。
    """
    out: list[tuple] = []
    for _ in range(limit):
        try:
            out.append(q.get_nowait())
        except queue.Empty:
            break
    return out


def _ask_string(parent, title: str, prompt: str) -> str | None:
    """简单的文本输入对话框（标准库没有 askstring 以外的好选择，
    这里自己实现以便控制多行提示的排版）。"""
    win = tk.Toplevel(parent)
    win.title(title)
    win.transient(parent)
    win.resizable(False, False)
    win.configure(bg=theme.BG)

    ttk.Label(win, text=prompt, style="Sub.TLabel", justify="left",
              wraplength=360).pack(anchor="w", padx=18, pady=(16, 8))
    var = tk.StringVar()
    entry = ttk.Entry(win, textvariable=var, show="●", width=36)
    entry.pack(fill="x", padx=18)
    entry.focus_set()

    result: dict = {"value": None}

    def ok(_e=None):
        result["value"] = var.get()
        win.destroy()

    def cancel(_e=None):
        result["value"] = None
        win.destroy()

    btns = ttk.Frame(win)
    btns.pack(fill="x", padx=18, pady=(14, 16))
    ttk.Button(btns, text="取消", command=cancel).pack(side="right")
    ttk.Button(btns, text="确定", style="Accent.TButton", command=ok).pack(side="right", padx=(0, 8))

    win.bind("<Return>", ok)
    win.bind("<Escape>", cancel)
    win.protocol("WM_DELETE_WINDOW", cancel)

    win.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - win.winfo_width()) // 2
    y = parent.winfo_rooty() + 160
    win.geometry(f"+{max(0, x)}+{max(0, y)}")

    win.grab_set()
    win.wait_window()
    return result["value"]


class _ResultDialog(tk.Toplevel):
    """任务完成后的结果面板：成功/失败数量 + 明细 + 打开输出文件夹。"""

    def __init__(self, parent, report: TaskReport, on_open):
        super().__init__(parent)
        self.title("执行结果")
        self.transient(parent)
        self.configure(bg=theme.BG)
        self.resizable(True, True)

        pad = ttk.Frame(self, padding=18)
        pad.pack(fill="both", expand=True)

        icon = "✓" if (report.ok and not report.failed and not report.cancelled) else (
            "■" if report.cancelled else "!")
        color = theme.SUCCESS if icon == "✓" else (theme.WARN if report.cancelled else theme.ERROR)
        ttk.Label(pad, text=f"{icon}  {report.summary_text()}",
                  style="Section.TLabel", foreground=color,
                  font=theme.ui_font(13, True)).pack(anchor="w")

        info = ttk.Frame(pad, style="Card.TFrame", padding=12)
        info.pack(fill="x", pady=(10, 0))
        kv = [("任务", report.op_title), ("耗时", f"{report.duration:.1f} 秒"),
              ("输出目录", report.output_dir or "—")]
        kv += [(k, str(v)) for k, v in report.metrics.items()]
        for i, (k, v) in enumerate(kv):
            ttk.Label(info, text=f"{k}：", style="CardSub.TLabel").grid(
                row=i, column=0, sticky="nw", padx=(0, 8), pady=1)
            ttk.Label(info, text=v, style="Card.TLabel",
                      wraplength=520, justify="left").grid(row=i, column=1, sticky="w", pady=1)

        if report.failures:
            ttk.Label(pad, text="明细", style="Section.TLabel").pack(anchor="w", pady=(14, 4))
            wrap = ttk.Frame(pad, style="Card.TFrame")
            wrap.pack(fill="both", expand=True)
            cols = ("kind", "file", "reason")
            tree = ttk.Treeview(wrap, columns=cols, show="headings", height=7)
            tree.heading("kind", text="状态")
            tree.heading("file", text="文件")
            tree.heading("reason", text="原因")
            tree.column("kind", width=60, anchor="center", stretch=False)
            tree.column("file", width=220, anchor="w")
            tree.column("reason", width=360, anchor="w")
            vs = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=vs.set)
            tree.pack(side="left", fill="both", expand=True)
            vs.pack(side="right", fill="y")
            tree.tag_configure("error", foreground=theme.ERROR)
            tree.tag_configure("skipped", foreground=theme.WARN)
            tree.tag_configure("warning", foreground=theme.WARN)
            for f in report.failures:
                label = {"error": "失败", "skipped": "跳过", "warning": "警告"}.get(f.category, f.category)
                tree.insert("", "end", values=(label, os.path.basename(f.path), f.reason),
                            tags=(f.category,))

        btns = ttk.Frame(pad)
        btns.pack(fill="x", pady=(14, 0))
        ttk.Button(btns, text="关闭", command=self.destroy).pack(side="right")
        if report.output_dir and os.path.isdir(report.output_dir):
            ttk.Button(btns, text="打开输出文件夹", style="Accent.TButton",
                       command=lambda: (on_open(), self.destroy())).pack(side="right", padx=(0, 8))
        try:
            ttk.Button(btns, text="复制明细",
                       command=lambda: self._copy(report)).pack(side="left")
        except Exception:
            pass

        self.update_idletasks()
        w = max(640, self.winfo_width())
        h = max(320, min(self.winfo_height(), 620))
        self.geometry(f"{w}x{h}")
        x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
        y = parent.winfo_rooty() + 80
        self.geometry(f"+{max(0, x)}+{max(0, y)}")
        self.bind("<Escape>", lambda e: self.destroy())
        self.grab_set()

    def _copy(self, report: TaskReport) -> None:
        try:
            self.clipboard_clear()
            self.clipboard_append(report.detail_text())
        except Exception:
            pass


class _RenamePreviewDialog(tk.Toplevel):
    """重命名预览：执行前把每个文件的新名字摆出来，避免"手一抖全改错"。"""

    def __init__(self, parent, rows):
        super().__init__(parent)
        self.title("重命名预览")
        self.transient(parent)
        self.configure(bg=theme.BG)
        self.resizable(True, True)

        pad = ttk.Frame(self, padding=16)
        pad.pack(fill="both", expand=True)

        ok = sum(1 for r in rows if r.status == "ok")
        conflict = sum(1 for r in rows if r.status == "conflict")
        other = len(rows) - ok - conflict

        summary = f"将重命名 {ok} 个文件"
        if conflict:
            summary += f"，{conflict} 个存在命名冲突（将被跳过）"
        if other:
            summary += f"，{other} 个无需处理"
        ttk.Label(pad, text=summary, style="Section.TLabel").pack(anchor="w")
        ttk.Label(pad, text="确认无误后关闭本窗口，再点击「开始执行」应用修改。",
                  style="Sub.TLabel").pack(anchor="w", pady=(4, 10))

        wrap = ttk.Frame(pad, style="Card.TFrame")
        wrap.pack(fill="both", expand=True)
        cols = ("old", "new", "status")
        tree = ttk.Treeview(wrap, columns=cols, show="headings", height=16)
        tree.heading("old", text="原文件名")
        tree.heading("new", text="新文件名")
        tree.heading("status", text="状态")
        tree.column("old", width=280, anchor="w")
        tree.column("new", width=280, anchor="w")
        tree.column("status", width=150, anchor="w")
        vs = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vs.set)
        tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")

        tree.tag_configure("ok", foreground=theme.SUCCESS)
        tree.tag_configure("conflict", foreground=theme.ERROR)
        tree.tag_configure("unchanged", foreground=theme.TEXT_SUB)
        tree.tag_configure("invalid", foreground=theme.ERROR)
        tree.tag_configure("empty", foreground=theme.ERROR)

        for r in rows:
            status = {"ok": "✓ 可重命名", "conflict": "✗ 命名冲突",
                      "unchanged": "— 名称未变化", "invalid": "✗ 参数错误",
                      "empty": "✗ 名称为空"}.get(r.status, r.status)
            if r.note and r.status != "ok":
                status += f"（{r.note}）"
            tree.insert("", "end", values=(r.old_name, r.new_name or "—", status),
                        tags=(r.status,))

        ttk.Button(pad, text="关闭", command=self.destroy).pack(anchor="e", pady=(12, 0))
        self.update_idletasks()
        self.geometry("760x520")
        x = parent.winfo_rootx() + (parent.winfo_width() - 760) // 2
        y = parent.winfo_rooty() + 60
        self.geometry(f"+{max(0, x)}+{max(0, y)}")
        self.bind("<Escape>", lambda e: self.destroy())
        self.grab_set()
