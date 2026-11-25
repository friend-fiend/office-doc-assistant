"""文件列表控件：拖动排序、删除、清空、去重。

拖动排序用 Treeview 原生的 ``move()`` 实现（按住左键拖动行），
比引入第三方排序控件更稳，也不增加打包体积。

关于 Windows/macOS 差异：
  · macOS 上 ``identify_row`` 的 y 坐标受 Retina 缩放影响，统一用事件坐标换算；
  · 行高已在 theme 里固定为 26，避免两个平台行高不同导致拖拽判定错位。
"""
from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Callable

from ..config import EXT_GROUP_LABEL, ext_set
from ..paths import human_size, long_path, open_path, reveal_path
from . import theme


class FileList(ttk.Frame):
    """一个"文件清单"面板，供各个操作共用。"""

    def __init__(self, master, on_change: Callable[[], None] | None = None, **kw):
        super().__init__(master, **kw)
        self.on_change = on_change or (lambda: None)
        self._paths: list[str] = []
        self._drag_item: str | None = None
        self._allowed_ext: set[str] = set()
        self._allowed_label = "任意文件"

        self._build()

    # ================================================================ 构建
    def _build(self) -> None:
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(0, 8))

        # 按语义分级，而不是六个一模一样的白按钮：
        #   添加（主路径，蓝描边）｜ 排序（低权重文字）｜ 移除 / 清空（破坏性红字）
        # 一屏里同等重量的按钮越少，眼睛越容易找到下一步该点哪儿。
        ttk.Button(bar, text="＋ 添加文件", style="Outline.TButton",
                   command=self.add_files).pack(side="left", padx=(0, 6))
        # 「添加文件夹」保持白底细边（次操作），与纯文字按钮再差一档 ——
        # 它也是「添加」，不该和「上移/下移」看起来一样重。
        ttk.Button(bar, text="添加文件夹",
                   command=self.add_folder).pack(side="left", padx=(0, 10))

        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y",
                                                   padx=(0, 8), pady=5)

        ttk.Button(bar, text="↑ 上移", style="Ghost.TButton",
                   command=lambda: self.move_selected(-1)).pack(side="left", padx=(0, 2))
        ttk.Button(bar, text="↓ 下移", style="Ghost.TButton",
                   command=lambda: self.move_selected(1)).pack(side="left", padx=(0, 10))

        ttk.Button(bar, text="移除所选", style="Ghost.TButton",
                   command=self.remove_selected).pack(side="left", padx=(0, 2))
        ttk.Button(bar, text="清空", style="Danger.TButton",
                   command=self.clear).pack(side="left")

        self.count_label = ttk.Label(bar, text="共 0 个文件", style="Sub.TLabel")
        self.count_label.pack(side="right")

        wrap = ttk.Frame(self, style="Card.TFrame")
        wrap.pack(fill="both", expand=True)

        cols = ("idx", "name", "type", "size", "dir")
        # height 是「最小请求高度」，不是显示行数上限：这个表格是 expand 的，
        # 窗口高时自动长到十几行；窗口矮时它先把高度让出来，免得把底部操作栏
        # 和日志行挤出窗口（pack 会裁剪最后装的控件，主按钮会被切掉一半）。
        self.tree = ttk.Treeview(wrap, columns=cols, show="headings",
                                 selectmode="extended", height=4)
        self.tree.heading("idx", text="#")
        self.tree.heading("name", text="文件名")
        self.tree.heading("type", text="类型")
        self.tree.heading("size", text="大小")
        self.tree.heading("dir", text="所在目录")
        self.tree.column("idx", width=44, anchor="center", stretch=False)
        self.tree.column("name", width=280, anchor="w")
        self.tree.column("type", width=70, anchor="center", stretch=False)
        self.tree.column("size", width=90, anchor="e", stretch=False)
        self.tree.column("dir", width=280, anchor="w")

        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

        # 空列表引导：Treeview 本身没有空态，空表就是一大片白，看着像没加载出来。
        # 用 place 叠一句提示在正中，有文件时立刻撤掉。
        self.empty_hint = ttk.Label(
            wrap, text="还没有文件 · 把文件拖到上方，或点「＋ 添加文件」",
            style="CardSub.TLabel")
        self.empty_hint.place(relx=0.5, rely=0.5, anchor="center")

        # 拖动排序
        self.tree.bind("<ButtonPress-1>", self._on_press)
        self.tree.bind("<B1-Motion>", self._on_motion)
        self.tree.bind("<ButtonRelease-1>", self._on_release)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-3>" if not theme.IS_MACOS else "<Button-2>", self._on_right_click)

        self._menu = tk.Menu(self, tearoff=0)
        self._menu.add_command(label="打开文件", command=self._open_selected)
        self._menu.add_command(label="在文件夹中显示", command=self._reveal_selected)
        self._menu.add_separator()
        self._menu.add_command(label="移除所选", command=self.remove_selected)
        self._menu.add_command(label="仅保留所选", command=self._keep_only_selected)
        self._menu.add_separator()
        self._menu.add_command(label="清空列表", command=self.clear)

    # ================================================================ 过滤设置
    def set_filter(self, groups: tuple[str, ...], label: str) -> None:
        """设置当前操作允许的扩展名（空集合=全部允许）。"""
        allowed: set[str] = set()
        for g in groups:
            allowed |= ext_set(g)
        self._allowed_ext = allowed
        self._allowed_label = label

    # ================================================================ 添加 / 删除
    def add_files(self, paths: list[str] | None = None) -> None:
        if paths is None:
            initial = os.path.expanduser("~")
            paths = filedialog.askopenfilenames(
                title="选择文件",
                initialdir=initial,
                filetypes=_filetypes_for(self._allowed_ext),
            )
            paths = list(paths)
        self._add(list(paths))

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择文件夹（会递归扫描其中的文件）")
        if not folder:
            return
        found: list[str] = []
        for root, _dirs, names in os.walk(folder):
            for n in sorted(names):
                if n.startswith(".") or n.startswith("~$"):
                    continue
                p = os.path.join(root, n)
                if self._accepts(p):
                    found.append(p)
        if not found:
            messagebox.showinfo(
                "没有匹配的文件",
                f"该文件夹下没有符合当前功能的文件（需要 {self._allowed_label}）。",
                parent=self,
            )
            return
        self._add(found)

    def _accepts(self, path: str) -> bool:
        if not self._allowed_ext:
            return True
        return os.path.splitext(path)[1].lower() in self._allowed_ext

    def _add(self, paths: list[str]) -> None:
        existing = {os.path.normcase(os.path.abspath(p)) for p in self._paths}
        added = 0
        skipped_type = 0
        skipped_dup = 0
        for p in paths:
            if not os.path.isfile(p):
                continue
            ap = os.path.abspath(p)
            key = os.path.normcase(ap)
            if key in existing:
                skipped_dup += 1
                continue
            if self._allowed_ext and not self._accepts(ap):
                skipped_type += 1
                continue
            self._paths.append(ap)
            existing.add(key)
            added += 1
        self._refresh()
        if added:
            self.on_change()
        notes = []
        if skipped_dup:
            notes.append(f"跳过重复 {skipped_dup} 个")
        if skipped_type:
            notes.append(f"跳过类型不符 {skipped_type} 个（需要 {self._allowed_label}）")
        if notes:
            self.count_label.configure(text=f"共 {len(self._paths)} 个文件（" + "，".join(notes) + "）")

    def remove_selected(self) -> None:
        sel = set(self.tree.selection())
        if not sel:
            return
        idxs = sorted((int(self.tree.item(i, "values")[0]) - 1 for i in sel), reverse=True)
        for i in idxs:
            if 0 <= i < len(self._paths):
                self._paths.pop(i)
        self._refresh()
        self.on_change()

    def _keep_only_selected(self) -> None:
        sel = set(self.tree.selection())
        if not sel:
            return
        keep = {int(self.tree.item(i, "values")[0]) - 1 for i in sel}
        self._paths = [p for i, p in enumerate(self._paths) if i in keep]
        self._refresh()
        self.on_change()

    def clear(self) -> None:
        if not self._paths:
            return
        if not messagebox.askyesno("确认清空", f"将从列表中移除全部 {len(self._paths)} 个文件。\n"
                                              "（只是移出列表，不会删除磁盘上的文件）",
                                  parent=self):
            return
        self._paths.clear()
        self._refresh()
        self.on_change()

    # ================================================================ 排序
    def move_selected(self, delta: int) -> None:
        sel = sorted((int(self.tree.item(i, "values")[0]) - 1 for i in self.tree.selection()))
        if not sel:
            return
        indices = sel if delta < 0 else list(reversed(sel))
        for i in indices:
            j = i + delta
            if 0 <= j < len(self._paths):
                self._paths[i], self._paths[j] = self._paths[j], self._paths[i]
        self._refresh()
        # 保持选中
        moved = {i + delta for i in sel if 0 <= i + delta < len(self._paths)}
        for item in self.tree.get_children():
            if int(self.tree.item(item, "values")[0]) - 1 in moved:
                self.tree.selection_add(item)
        self.on_change()

    def _on_press(self, event) -> None:
        if self.tree.identify_region(event.x, event.y) != "cell":
            self._drag_item = None
            return
        self._drag_item = self.tree.identify_row(event.y)

    def _on_motion(self, event) -> None:
        if not self._drag_item:
            return
        target = self.tree.identify_row(event.y)
        if not target or target == self._drag_item:
            return
        # 在数据层交换（而不是只移动显示行），保证 _paths 顺序与界面一致
        try:
            a = int(self.tree.item(self._drag_item, "values")[0]) - 1
            b = int(self.tree.item(target, "values")[0]) - 1
        except (ValueError, IndexError):
            return
        if not (0 <= a < len(self._paths) and 0 <= b < len(self._paths)):
            return
        item = self._paths.pop(a)
        self._paths.insert(b, item)
        self._drag_item = self.tree.get_children()[b] if b < len(self.tree.get_children()) else None
        self._refresh(keep_selection=False)
        # 重新定位到拖动行
        children = self.tree.get_children()
        if 0 <= b < len(children):
            self._drag_item = children[b]
            self.tree.selection_set(self._drag_item)
        self.on_change()

    def _on_release(self, _event) -> None:
        self._drag_item = None

    # ================================================================ 展示
    def _refresh(self, keep_selection: bool = True) -> None:
        self.tree.delete(*self.tree.get_children())
        for i, p in enumerate(self._paths):
            ext = os.path.splitext(p)[1].lower()
            self.tree.insert("", "end", values=(
                i + 1,
                os.path.basename(p),
                ext.lstrip(".") or "?",
                human_size(_safe_size(p)),
                os.path.dirname(p),
            ))
        self.count_label.configure(text=f"共 {len(self._paths)} 个文件")
        if self._paths:
            self.empty_hint.place_forget()
        else:
            self.empty_hint.place(relx=0.5, rely=0.5, anchor="center")
        if not keep_selection:
            self.tree.selection_remove(*self.tree.selection())

    def _on_double_click(self, event) -> None:
        if self.tree.identify_region(event.x, event.y) == "cell":
            self._open_selected()

    def _on_right_click(self, event) -> None:
        row = self.tree.identify_row(event.y)
        if row and row not in self.tree.selection():
            self.tree.selection_set(row)
        if self.tree.selection():
            self._menu.tk_popup(event.x_root, event.y_root)

    def _open_selected(self) -> None:
        for p in self.selected_paths()[:1]:
            try:
                open_path(p)
            except Exception as exc:
                messagebox.showerror("打开失败", str(exc), parent=self)

    def _reveal_selected(self) -> None:
        for p in self.selected_paths()[:1]:
            try:
                reveal_path(p)
            except Exception as exc:
                messagebox.showerror("定位失败", str(exc), parent=self)

    # ================================================================ 接口
    def get_files(self) -> list[str]:
        return list(self._paths)

    def set_files(self, paths: list[str]) -> None:
        self._paths = [os.path.abspath(p) for p in paths if os.path.isfile(p)]
        self._refresh()
        self.on_change()

    def selected_paths(self) -> list[str]:
        out = []
        for item in self.tree.selection():
            try:
                i = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= i < len(self._paths):
                    out.append(self._paths[i])
            except (ValueError, IndexError):
                continue
        return out

    def count(self) -> int:
        return len(self._paths)


def _safe_size(path: str) -> int:
    try:
        return os.path.getsize(long_path(path))
    except Exception:
        return 0


def _filetypes_for(allowed: set[str]) -> list[tuple[str, str]]:
    if not allowed:
        return [("所有文件", "*.*")]
    pattern = " ".join(f"*{e}" for e in sorted(allowed))
    return [(f"支持的文件 ({len(allowed)} 种)", pattern), ("所有文件", "*.*")]


_ = EXT_GROUP_LABEL
