"""参数区：根据 OpDef.params 的声明动态生成控件。

这样做的价值：新增一个功能只需要写 OpDef + handler，界面自动出现对应参数，
不会出现"加了参数忘了在界面加控件"的问题。

支持的控件种类（Param.kind）：
  entry / int / float / range / text / password / choice / bool / position / csv
  per_file_config — 特殊：每文件独立配置表（默认 + 行内覆盖下拉）
"""
from __future__ import annotations

import os
import tkinter as tk
from tkinter import ttk
from typing import Any, Callable

from ..core.registry import OpDef, Param
from . import theme


def _per_file_panel(parent, param: Param, get_files, on_change=None):
    """渲染「每文件独立配置」控件（默认配置 + 文件列表 + 行内覆盖）。

    返回 ``(container, get_values)``，其中 ``get_values()`` 返回：
    ``{"default": {...}, "overrides": {file_idx: {...}}}``
    """
    import tkinter as tk
    from tkinter import ttk

    axes = param.axes or ()
    if not axes:
        ttk.Label(parent, text="（未配置每文件维度）").pack()
        return parent, lambda: param.default

    container = ttk.Frame(parent, style="Card.TFrame")
    container.pack(fill="x", pady=2)

    # ---- 默认配置区
    default_frame = ttk.LabelFrame(container, text="默认配置（应用于未单独设置的文件）",
                                   style="Card.TLabelframe", padding=8)
    default_frame.pack(fill="x", padx=4, pady=(0, 6))
    default_frame._ax_vars = {}
    default_frame._ax_widgets = {}  # key -> Combobox/ Spinbox 对象（用于 label→value 还原）

    for i, ax in enumerate(axes):
        ttk.Label(default_frame, text=f"{ax['label']}：", style="Card.TLabel").grid(
            row=0, column=i * 2, sticky="w", padx=(0, 4))
        if ax.get("widget") == "int":
            var = tk.IntVar(value=int(ax.get("default", 1)))
            widget = ttk.Spinbox(default_frame, textvariable=var, from_=ax.get("min", 1),
                                 to=ax.get("max", 999), width=6,
                                 command=on_change)
            widget.grid(row=0, column=i * 2 + 1, sticky="w", padx=(0, 12))
        else:
            var = tk.StringVar(value=str(ax.get("default", ax["choices"][0][0])))
            cb = ttk.Combobox(default_frame, textvariable=var,
                              values=[label for _v, label in ax["choices"]],
                              state="readonly", width=10)
            cb._value_to_label = {v: label for v, label in ax["choices"]}
            cb._label_to_value = {label: v for v, label in ax["choices"]}
            for v, label in ax["choices"]:
                if v == ax.get("default", ax["choices"][0][0]):
                    cb.set(label)
                    break
            cb.bind("<<ComboboxSelected>>", lambda e: on_change())
            cb.grid(row=0, column=i * 2 + 1, sticky="w", padx=(0, 12))
            widget = cb
        default_frame._ax_vars[ax["key"]] = var
        default_frame._ax_widgets[ax["key"]] = widget

    # ---- 文件表格
    table_frame = ttk.Frame(container, style="Card.TFrame")
    table_frame.pack(fill="x", padx=4)

    cols = ["idx", "name"] + [ax["key"] for ax in axes] + ["_actions"]
    headings = ["#", "文件名"] + [ax["label"] for ax in axes] + ["操作"]
    tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=8)
    for col, head in zip(cols, headings):
        tree.heading(col, text=head)
        if col == "idx":
            tree.column(col, width=44, anchor="center", stretch=False)
        elif col == "name":
            tree.column(col, width=240, anchor="w")
        elif col == "_actions":
            tree.column(col, width=64, anchor="center", stretch=False)
        else:
            tree.column(col, width=110, anchor="center")
    vs = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vs.set)
    tree.pack(side="left", fill="both", expand=True)
    vs.pack(side="right", fill="y")

    # overrides: dict[file_idx_str -> dict[axis_key -> value]]
    overrides_state: dict[str, dict[str, str]] = {"data": {}}
    default_frame._overrides = overrides_state
    default_frame._tree = tree
    default_frame._axes = axes

    def get_default_values() -> dict:
        out = {}
        for ax in axes:
            widget = default_frame._ax_widgets.get(ax["key"])
            if ax.get("widget") == "int":
                try:
                    out[ax["key"]] = int(default_frame._ax_vars[ax["key"]].get())
                except Exception:
                    out[ax["key"]] = int(ax.get("default", 1))
            else:
                # Combobox：把显示 label 还原成内部 value
                if isinstance(widget, ttk.Combobox):
                    label = widget.get()
                    out[ax["key"]] = widget._label_to_value.get(label, label)
                else:
                    out[ax["key"]] = default_frame._ax_vars[ax["key"]].get()
        return out

    def _refresh_rows():
        files = get_files()
        # 保留已有覆盖
        existing = dict(overrides_state["data"])
        overrides_state["data"].clear()
        tree.delete(*tree.get_children())
        for i, fp in enumerate(files):
            name = os.path.basename(fp) if fp else f"文件 {i + 1}"
            values = [str(i + 1), name]
            for ax in axes:
                ov = existing.get(str(i), {}).get(ax["key"], "")
                values.append(_display_value(ax, ov))
            values.append("重置")
            tree.insert("", "end", iid=str(i), values=values)
            if str(i) in existing:
                overrides_state["data"][str(i)] = existing[str(i)]

    # ---- 行内编辑：双击单元格进入编辑模式（弹下拉框）
    def _on_double_click(event):
        region = tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col_id = tree.identify_column(event.x)
        row_id = tree.identify_row(event.y)
        if not row_id or not col_id:
            return
        # 列索引从 #1 开始
        col_idx = int(col_id.replace("#", "")) - 1
        if col_idx < 2 or col_idx >= 2 + len(axes):
            return  # 不是 axes 列
        ax = axes[col_idx - 2]

        # 创建临时 Combobox
        x, y, w, h = tree.bbox(row_id, col_id)
        if not x:
            return

        if ax.get("widget") == "int":
            current = overrides_state["data"].get(row_id, {}).get(ax["key"], "")
            var = tk.StringVar(value=str(current) if current != "" else "")
            editor = ttk.Spinbox(tree, textvariable=var, from_=ax.get("min", 1),
                                 to=ax.get("max", 999), width=8)
        else:
            choices = [("", "继承")] + ax["choices"]
            current_val = overrides_state["data"].get(row_id, {}).get(ax["key"], "")
            var = tk.StringVar()
            editor = ttk.Combobox(tree, textvariable=var, state="readonly",
                                  values=[label for _v, label in choices], width=10)
            # 设置当前显示
            disp = _display_value(ax, current_val)
            editor.set(disp)
            editor._label_to_value = {label: v for v, label in choices}
            editor._value_to_label = {v: label for v, label in choices}

        editor.place(x=x, y=y, width=w, height=h)
        editor.focus_set()

        def _commit(_e=None):
            new = var.get().strip()
            if ax.get("widget") == "int":
                if not new:
                    # 清空 = 继承
                    overrides_state["data"].get(row_id, {}).pop(ax["key"], None)
                    if not overrides_state["data"].get(row_id):
                        overrides_state["data"].pop(row_id, None)
                else:
                    try:
                        iv = int(new)
                    except Exception:
                        iv = ax.get("default", 1)
                    overrides_state["data"].setdefault(row_id, {})[ax["key"]] = iv
            else:
                if new in ("", "继承"):
                    overrides_state["data"].get(row_id, {}).pop(ax["key"], None)
                    if not overrides_state["data"].get(row_id):
                        overrides_state["data"].pop(row_id, None)
                else:
                    actual = editor._label_to_value.get(new, new)
                    overrides_state["data"].setdefault(row_id, {})[ax["key"]] = actual
            editor.destroy()
            _refresh_rows()
            if on_change:
                on_change()

        editor.bind("<Return>", _commit)
        editor.bind("<FocusOut>", _commit)
        editor.bind("<Escape>", lambda e: editor.destroy())

    tree.bind("<Double-1>", _on_double_click)

    # ---- 批量操作按钮
    btns = ttk.Frame(container, style="Card.TFrame")
    btns.pack(fill="x", padx=4, pady=(4, 6))
    ttk.Button(btns, text="套用默认配置到选中行",
               command=lambda: _apply_to_selected()).pack(side="left", padx=(0, 6))
    ttk.Button(btns, text="清除选中行的覆盖",
               command=lambda: _clear_selected()).pack(side="left", padx=(0, 6))
    ttk.Button(btns, text="清除全部覆盖",
               command=lambda: _clear_all()).pack(side="left")
    ttk.Label(btns, text="提示：双击单元格修改；选「继承」= 用默认配置",
              style="CardSub.TLabel").pack(side="right")

    def _apply_to_selected():
        sel = tree.selection()
        if not sel:
            return
        defaults = get_default_values()
        for row_id in sel:
            overrides_state["data"].setdefault(row_id, {}).update(defaults)
        _refresh_rows()
        if on_change:
            on_change()

    def _clear_selected():
        sel = tree.selection()
        for row_id in sel:
            overrides_state["data"].pop(row_id, None)
        _refresh_rows()
        if on_change:
            on_change()

    def _clear_all():
        overrides_state["data"].clear()
        _refresh_rows()
        if on_change:
            on_change()

    def _display_value(ax, val):
        """把内部值转成表格显示文字。空值显示「继承」。"""
        if val is None or val == "" or val == "inherit":
            return "继承"
        if ax.get("widget") == "int":
            return str(val)
        choices = ax.get("choices", [])
        for v, label in choices:
            if v == val:
                return label
        return str(val)

    def get_values() -> dict:
        return {
            "default": get_default_values(),
            "overrides": dict(overrides_state["data"]),
        }

    def refresh_files() -> None:
        _refresh_rows()

    # 暴露刷新接口，外部（FileList 变更时）可调用
    container.refresh_files = refresh_files  # type: ignore[attr-defined]
    container.get_values = get_values  # type: ignore[attr-defined]

    # 初始化一次
    _refresh_rows()

    if param.hint:
        ttk.Label(container, text=param.hint, style="CardSub.TLabel",
                  wraplength=560, justify="left").pack(anchor="w", padx=8, pady=(0, 6))

    return container, get_values


class ScrollableFrame(ttk.Frame):
    """可滚动容器：内容超出高度时出现纵向滚动条。

    Tkinter 没有现成的滚动 Frame，标准做法是 Canvas + inner Frame +
    bind <Configure> 同步 scrollregion，并在 inner 宽度变化时同步 canvas 窗口宽度。
    """

    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.canvas = tk.Canvas(self, highlightthickness=0, bg=theme.BG)
        self.vs = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vs.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vs.grid(row=0, column=1, sticky="ns")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.inner = ttk.Frame(self.canvas, style="Card.TFrame")
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.inner.bind("<Configure>", self._on_inner)
        self.canvas.bind("<Configure>", self._on_canvas)
        # 鼠标进入时才接管滚轮，避免抢走文件列表/日志的滚动
        self.canvas.bind("<Enter>", lambda e: self._bind_wheel(True))
        self.canvas.bind("<Leave>", lambda e: self._bind_wheel(False))

    def _on_inner(self, _e=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas(self, event) -> None:
        self.canvas.itemconfigure(self._win, width=event.width)

    def _bind_wheel(self, on: bool) -> None:
        if on:
            self.canvas.bind_all("<MouseWheel>", self._on_wheel)
            self.canvas.bind_all("<Button-4>", self._on_wheel)
            self.canvas.bind_all("<Button-5>", self._on_wheel)
        else:
            self.canvas.unbind_all("<MouseWheel>")
            self.canvas.unbind_all("<Button-4>")
            self.canvas.unbind_all("<Button-5>")

    def _on_wheel(self, event) -> None:
        delta = 0
        if getattr(event, "num", None) == 4:
            delta = -3
        elif getattr(event, "num", None) == 5:
            delta = 3
        elif getattr(event, "delta", 0):
            # Windows 是 ±120 的倍数，macOS 是小整数
            step = -1 * (event.delta // 120) if abs(event.delta) >= 120 else -1 * event.delta
            delta = step * 3
        if delta:
            self.canvas.yview_scroll(delta, "units")

    def clear(self) -> None:
        """销毁内部全部子控件。

        注意：ParamPanel.build() 里的标题/网格/说明标签都是 inner 的子控件，
        所以每次 build 必须重新创建它们 —— 否则第二次 build 会操作已被销毁的
        控件并抛 TclError: invalid command name（这个坑很隐蔽，只有连续切换
        两个以上功能才会暴露）。
        """
        for child in self.inner.winfo_children():
            child.destroy()

    def refresh(self) -> None:
        """子控件增删后刷新滚动区域。"""
        self._on_inner()


class ParamPanel(ttk.Frame):
    """参数面板。"""

    def __init__(self, master, on_change: Callable[[], None] | None = None,
                 get_files: Callable[[], list[str]] | None = None, **kw):
        super().__init__(master, **kw)
        self.on_change = on_change or (lambda: None)
        self.get_files = get_files or (lambda: [])
        self._vars: dict[str, tk.Variable] = {}
        self._widgets: dict[str, tk.Widget] = {}
        self._rows: dict[str, tk.Widget] = {}
        self._params: tuple[Param, ...] = ()
        self._pw_visible: set[str] = set()
        # 特殊控件：每文件配置表（key -> container, get_values）
        self._per_file_containers: dict[str, ttk.Frame] = {}

        head = ttk.Frame(self)
        head.pack(fill="x", pady=(0, 2))
        ttk.Label(head, text="参数设置", style="Section.TLabel").pack(side="left")
        self.op_title = ttk.Label(head, text="", style="Sub.TLabel")
        self.op_title.pack(side="left", padx=(10, 0))

        self.scroll = ScrollableFrame(self)
        self.scroll.pack(fill="both", expand=True)

        # desc / grid_frame 在每次 build() 时重建，这里不预先创建（见 ScrollableFrame.clear 注释）
        self.desc: ttk.Label | None = None
        self.grid_frame: ttk.Frame | None = None

    def refresh_files(self) -> None:
        """文件列表变更时调用：刷新所有 per_file_config 子控件的行。"""
        for c in self._per_file_containers.values():
            refresh = getattr(c, "refresh_files", None)
            if refresh:
                try:
                    refresh()
                except Exception:
                    pass

    # ================================================================ 构建
    def build(self, op: OpDef) -> None:
        self.scroll.clear()
        self._vars.clear()
        self._widgets.clear()
        self._rows.clear()
        self._params = op.params
        self._per_file_containers.clear()
        self.op_title.configure(text=f"— {op.title}")

        self.desc = ttk.Label(self.scroll.inner, text=op.description or "",
                              style="CardSub.TLabel", justify="left", wraplength=520)
        self.desc.pack(fill="x", padx=14, pady=(12, 0))

        self.grid_frame = ttk.Frame(self.scroll.inner, style="Card.TFrame")
        self.grid_frame.pack(fill="both", expand=True, padx=14, pady=(8, 14))
        self.grid_frame.columnconfigure(0, weight=0, minsize=132)
        self.grid_frame.columnconfigure(1, weight=1)
        self.grid_frame.columnconfigure(2, weight=0, minsize=28)
        self.grid_frame.columnconfigure(3, weight=0, minsize=132)
        self.grid_frame.columnconfigure(4, weight=1)
        self.grid_frame.columnconfigure(5, weight=0)

        if not op.params:
            ttk.Label(self.grid_frame, text="该功能无需设置参数，直接点击「开始执行」即可。",
                      style="CardSub.TLabel").grid(row=0, column=0, columnspan=6,
                                                   sticky="w", pady=8)
            self.scroll.refresh()
            return

        # 两列布局：偶数列放标签，奇数列放控件；太宽的控件独占一行
        row = 0
        col = 0
        for p in op.params:
            wide = p.kind in ("text", "range") or p.width >= 20
            if wide and col != 0:
                row += 1
                col = 0
            self._make_row(p, row, col)
            if wide:
                row += 1
                col = 0
            else:
                col += 2
                if col > 4:
                    col = 0
                    row += 1

        self._apply_visibility()
        self.scroll.refresh()

    def _make_row(self, p: Param, row: int, col: int) -> None:
        container = ttk.Frame(self.grid_frame, style="Card.TFrame")
        wide = p.kind in ("text", "range", "per_file_config") or p.width >= 20
        container.grid(row=row, column=col, columnspan=(6 if col == 0 and wide else 2),
                       sticky="ew", padx=(0, 18), pady=4)
        self._rows[p.key] = container

        if p.kind == "per_file_config":
            # 这种控件自带标题（LabelFrame），不需要外层 label
            widget, _ = _per_file_panel(container, p, self.get_files, self.on_change)
            self._widgets[p.key] = widget
            self._per_file_containers[p.key] = widget
            if p.hint:
                ttk.Label(container, text=p.hint, style="CardSub.TLabel",
                          wraplength=560, justify="left").pack(anchor="w", padx=(0, 2))
            return

        inner = ttk.Frame(container, style="Card.TFrame")
        inner.pack(fill="x")
        ttk.Label(inner, text=p.label + "：", style="Card.TLabel",
                  width=0, anchor="w").pack(side="left")

        widget = self._make_widget(inner, p)
        widget.pack(side="left", fill="x", expand=True, padx=(2, 0))
        self._widgets[p.key] = widget

        if p.hint:
            ttk.Label(container, text=p.hint, style="CardSub.TLabel",
                      wraplength=560, justify="left").pack(anchor="w", padx=(0, 2))

        if p.kind == "bool":
            # 勾选框在左、文字在右更自然
            widget.pack_forget()
            widget.pack(side="left")

    def _make_widget(self, parent, p: Param) -> tk.Widget:
        default = p.default

        if p.kind == "bool":
            var = tk.BooleanVar(value=bool(default))
            self._vars[p.key] = var
            return ttk.Checkbutton(parent, text="启用" if p.key == "add_bookmarks" else "",
                                   variable=var, style="Card.TCheckbutton",
                                   command=self.on_change)

        if p.kind == "choice":
            var = tk.StringVar(value=str(default))
            self._vars[p.key] = var
            values = [label for _v, label in p.choices] or [str(default)]
            cb = ttk.Combobox(parent, textvariable=var, values=values,
                              state="readonly", width=max(10, p.width))
            cb._label_to_value = {label: v for v, label in p.choices}  # type: ignore[attr-defined]
            cb._value_to_label = {v: label for v, label in p.choices}  # type: ignore[attr-defined]
            # 显示默认值对应的中文标签
            for v, label in p.choices:
                if str(v) == str(default):
                    cb.set(label)
                    break
            # 切换选项后必须重算 visible_if，否则按"页码范围"等联动参数永远不显示
            cb.bind("<<ComboboxSelected>>", lambda e: (self.on_change(), self._apply_visibility()))
            return cb

        if p.kind == "position":
            var = tk.StringVar(value=str(default))
            self._vars[p.key] = var
            labels = [label for _v, label in p.choices]
            cb = ttk.Combobox(parent, textvariable=var, values=labels,
                              state="readonly", width=max(10, p.width))
            cb._value_to_label = {v: label for v, label in p.choices}  # type: ignore[attr-defined]
            for v, label in p.choices:
                if str(v) == str(default):
                    cb.set(label)
                    break
            cb.bind("<<ComboboxSelected>>", lambda e: self.on_change())
            return cb

        if p.kind == "password":
            var = tk.StringVar(value=str(default))
            self._vars[p.key] = var
            box = ttk.Frame(parent, style="Card.TFrame")
            entry = ttk.Entry(box, textvariable=var, show="●", width=max(10, p.width))
            entry.pack(side="left", fill="x", expand=True)

            def toggle(_e=None, v=var, en=entry):
                if en.cget("show"):
                    en.configure(show="")
                else:
                    en.configure(show="●")

            eye = ttk.Label(box, text="👁", style="CardSub.TLabel", cursor="hand2")
            eye.pack(side="left", padx=(4, 0))
            eye.bind("<Button-1>", toggle)
            return box

        if p.kind == "text":
            var = tk.StringVar(value=str(default))
            self._vars[p.key] = var
            return ttk.Entry(parent, textvariable=var, width=max(20, p.width))

        # entry / int / float / range / csv 都是单行输入，校验在取值时做
        var = tk.StringVar(value=_fmt(default))
        self._vars[p.key] = var
        e = ttk.Entry(parent, textvariable=var, width=max(8, p.width))
        if p.kind in ("int", "float"):
            e.bind("<KeyRelease>", lambda ev: self.on_change())
        return e

    # ================================================================ 联动
    def _apply_visibility(self) -> None:
        for p in self._params:
            row = self._rows.get(p.key)
            if row is None:
                continue
            if p.visible_if is None:
                row.grid()
                continue
            dep_key, allowed = p.visible_if
            cur = self._raw_value(dep_key)
            if cur in allowed:
                row.grid()
            else:
                row.grid_remove()

    def _raw_value(self, key: str) -> Any:
        widget = self._widgets.get(key)
        if widget is None:
            return None
        if isinstance(widget, ttk.Combobox):
            label = widget.get()
            return getattr(widget, "_label_to_value", {}).get(label, label)
        return self._vars.get(key).get() if key in self._vars else None

    # ================================================================ 取值
    def get_values(self) -> dict:
        """返回参数字典；int/float 已转换，choice 已还原成内部值。"""
        out: dict = {}
        for p in self._params:
            widget = self._widgets.get(p.key)
            if widget is None:
                out[p.key] = p.default
                continue

            if p.kind == "bool":
                out[p.key] = bool(self._vars[p.key].get())
            elif p.kind in ("choice", "position"):
                label = widget.get()  # type: ignore[union-attr]
                out[p.key] = getattr(widget, "_label_to_value", {}).get(label, label)
            elif p.kind == "int":
                out[p.key] = _to_int(self._vars[p.key].get(), p)
            elif p.kind == "float":
                out[p.key] = _to_float(self._vars[p.key].get(), p)
            elif p.kind == "per_file_config":
                # 容器对象上挂的 get_values() 返回 {"default":..., "overrides":...}
                get_v = getattr(widget, "get_values", None)
                out[p.key] = get_v() if get_v else p.default
            else:
                out[p.key] = str(self._vars[p.key].get())
        return out

    def set_values(self, values: dict) -> None:
        for key, val in values.items():
            widget = self._widgets.get(key)
            if widget is None:
                continue
            p = next((x for x in self._params if x.key == key), None)
            if p and p.kind in ("choice", "position") and isinstance(widget, ttk.Combobox):
                label = getattr(widget, "_value_to_label", {}).get(val, str(val))
                widget.set(label)
            elif isinstance(self._vars.get(key), tk.BooleanVar):
                self._vars[key].set(bool(val))
            elif key in self._vars:
                self._vars[key].set(_fmt(val))
        self._apply_visibility()

    # ================================================================ 校验
    def validate(self) -> str:
        """返回空字符串表示通过。"""
        for p in self._params:
            row = self._rows.get(p.key)
            if row is not None and not row.winfo_ismapped():
                continue  # 隐藏的参数不参与校验
            if p.kind == "int":
                raw = self._vars[p.key].get().strip()
                if raw == "":
                    continue
                try:
                    v = int(float(raw))
                except ValueError:
                    return f"「{p.label}」需要填写整数，当前是「{raw}」"
                if p.min_value is not None and v < p.min_value:
                    return f"「{p.label}」不能小于 {int(p.min_value)}"
                if p.max_value is not None and v > p.max_value:
                    return f"「{p.label}」不能大于 {int(p.max_value)}"
            elif p.kind == "float":
                raw = self._vars[p.key].get().strip()
                if raw == "":
                    continue
                try:
                    v = float(raw)
                except ValueError:
                    return f"「{p.label}」需要填写数字，当前是「{raw}」"
                if p.min_value is not None and v < p.min_value:
                    return f"「{p.label}」不能小于 {p.min_value}"
                if p.max_value is not None and v > p.max_value:
                    return f"「{p.label}」不能大于 {p.max_value}"
            elif p.kind == "range":
                raw = self._vars[p.key].get().strip()
                if raw:
                    from ..core.pdfutil import parse_ranges

                    try:
                        parse_ranges(raw, 10_000)
                    except Exception as exc:
                        return f"「{p.label}」格式不正确：{exc}"
        return ""


# ------------------------------------------------------------------ 工具
def _fmt(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        return ("%g" % v)
    return str(v)


def _to_int(raw: str, p: Param) -> int:
    try:
        v = int(float(str(raw).strip()))
    except Exception:
        v = int(p.default or 0)
    if p.min_value is not None:
        v = max(int(p.min_value), v)
    if p.max_value is not None:
        v = min(int(p.max_value), v)
    return v


def _to_float(raw: str, p: Param) -> float:
    try:
        v = float(str(raw).strip())
    except Exception:
        v = float(p.default or 0.0)
    if p.min_value is not None:
        v = max(float(p.min_value), v)
    if p.max_value is not None:
        v = min(float(p.max_value), v)
    return v
