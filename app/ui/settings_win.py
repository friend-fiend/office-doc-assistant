"""设置窗口：外部依赖状态、手动指定路径、下载指引。"""
from __future__ import annotations

import os
import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox, ttk

from ..deps import TOOL_SPECS, detect_all
from ..paths import open_path
from ..settings import SETTINGS
from . import theme

_KEY_MAP = {
    "poppler": "poppler_dir",
    "soffice": "soffice_path",
    "ghostscript": "ghostscript_path",
}


class SettingsWindow(tk.Toplevel):
    def __init__(self, master, on_saved=None):
        super().__init__(master)
        self.title("设置")
        self.transient(master)
        self.resizable(True, False)
        self.configure(bg=theme.BG)
        self.on_saved = on_saved

        self._vars: dict[str, tk.StringVar] = {}
        self._status_labels: dict[str, ttk.Label] = {}
        self._build()
        self.refresh_status()
        self._center(master)
        self.grab_set()

    # ================================================================ 构建
    def _build(self) -> None:
        pad = ttk.Frame(self)
        pad.pack(fill="both", expand=True, padx=18, pady=16)

        ttk.Label(pad, text="外部程序检测", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            pad,
            text="本程序有三项功能依赖本机已安装的免费外部程序。未安装时对应功能会被禁用，"
                 "其余功能不受影响。你也可以把绿色便携版解压到程序目录下的 vendor 文件夹，"
                 "程序会自动识别。",
            style="Sub.TLabel", wraplength=660, justify="left",
        ).pack(anchor="w", pady=(4, 12))

        for spec in TOOL_SPECS:
            card = ttk.Frame(pad, style="Card.TFrame", padding=12)
            card.pack(fill="x", pady=6)

            top = ttk.Frame(card, style="Card.TFrame")
            top.pack(fill="x")
            ttk.Label(top, text=spec.title, style="Card.TLabel",
                      font=theme.ui_font(10, True)).pack(side="left")
            st = ttk.Label(top, text="", style="CardSub.TLabel")
            st.pack(side="right")
            self._status_labels[spec.key] = st

            row = ttk.Frame(card, style="Card.TFrame")
            row.pack(fill="x", pady=(8, 0))
            key = _KEY_MAP[spec.key]
            var = tk.StringVar(value=str(SETTINGS.get(key) or ""))
            self._vars[key] = var

            label = "目录" if spec.is_dir else "可执行文件"
            ttk.Label(row, text=f"{label}：", style="Card.TLabel").pack(side="left")
            ttk.Entry(row, textvariable=var, width=56).pack(side="left", fill="x",
                                                          expand=True, padx=(4, 6))
            ttk.Button(row, text="浏览…", width=8,
                       command=lambda s=spec, v=var: self._browse(s, v)).pack(side="left")

            hint = ttk.Frame(card, style="Card.TFrame")
            hint.pack(fill="x", pady=(8, 0))
            ttk.Label(hint, text="影响功能：" + "、".join(spec.affects),
                      style="CardSub.TLabel").pack(anchor="w")
            ttk.Label(hint, text=spec.help_text, style="CardSub.TLabel",
                      wraplength=640, justify="left").pack(anchor="w", pady=(2, 4))
            if spec.help_url:
                link = ttk.Label(hint, text="→ 打开下载页面", style="CardSub.TLabel",
                                 foreground=theme.PRIMARY, cursor="hand2")
                link.pack(anchor="w")
                link.bind("<Button-1>", lambda e, u=spec.help_url: webbrowser.open(u))

        ttk.Separator(pad).pack(fill="x", pady=14)

        opt = ttk.Frame(pad)
        opt.pack(fill="x")
        self.var_open_when_done = tk.BooleanVar(
            value=bool(SETTINGS.get("open_output_when_done", False)))
        ttk.Checkbutton(opt, text="任务完成后自动打开输出文件夹",
                        variable=self.var_open_when_done).pack(anchor="w")
        self.var_verbose = tk.BooleanVar(value=bool(SETTINGS.get("verbose_log", False)))
        ttk.Checkbutton(opt, text="记录详细日志（排查问题时开启）",
                        variable=self.var_verbose).pack(anchor="w", pady=(4, 0))

        # ---- 底部按钮
        btns = ttk.Frame(pad)
        btns.pack(fill="x", pady=(18, 0))
        ttk.Button(btns, text="重新检测", command=self.refresh_status).pack(side="left")
        ttk.Button(btns, text="打开日志目录",
                   command=self._open_logs).pack(side="left", padx=(8, 0))
        ttk.Button(btns, text="恢复自动探测",
                   command=self._reset).pack(side="left", padx=(8, 0))
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="保存", style="Accent.TButton",
                   command=self._save).pack(side="right", padx=(0, 8))

    def _center(self, master) -> None:
        self.update_idletasks()
        try:
            w, h = self.winfo_width(), self.winfo_height()
            x = master.winfo_rootx() + (master.winfo_width() - w) // 2
            y = master.winfo_rooty() + 60
            self.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

    # ================================================================ 交互
    def _browse(self, spec, var: tk.StringVar) -> None:
        if spec.is_dir:
            p = filedialog.askdirectory(title=f"选择 {spec.title} 所在目录", parent=self)
            if p:
                var.set(p)
        else:
            p = filedialog.askopenfilename(
                title=f"选择 {spec.title} 的可执行文件", parent=self,
                filetypes=[("可执行文件", "*.exe" if theme.IS_WINDOWS else "*"),
                           ("所有文件", "*.*")])
            if p:
                var.set(p)
        self.refresh_status()

    def _save(self) -> None:
        for k, v in self._vars.items():
            SETTINGS.set(k, v.get().strip())
        SETTINGS.set("open_output_when_done", bool(self.var_open_when_done.get()))
        SETTINGS.set("verbose_log", bool(self.var_verbose.get()))
        SETTINGS.save()
        st = detect_all(force=True)
        if self.on_saved:
            self.on_saved(st)
        missing = [s.spec.title for s in st.values() if not s.available]
        if missing:
            messagebox.showwarning(
                "已保存",
                "设置已保存，但仍有外部程序未检测到：\n· " + "\n· ".join(missing)
                + "\n\n对应的功能按钮会保持禁用状态。",
                parent=self)
        else:
            messagebox.showinfo("已保存", "设置已保存，全部外部程序检测通过。", parent=self)
        self.destroy()

    def _reset(self) -> None:
        for k in self._vars:
            self._vars[k].set("")
        SETTINGS.reset_tools()
        SETTINGS.save()
        self.refresh_status()
        if self.on_saved:
            self.on_saved(detect_all(force=True))

    def _open_logs(self) -> None:
        from ..logging_setup import log_dir

        try:
            open_path(log_dir())
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc), parent=self)

    # ================================================================ 状态
    def refresh_status(self) -> None:
        # 把界面上填的路径先临时生效，这样"填完点检测"就能看到效果
        for k, v in self._vars.items():
            SETTINGS.set(k, v.get().strip())
        st = detect_all(force=True)
        for key, status in st.items():
            lbl = self._status_labels.get(key)
            if lbl is None:
                continue
            if status.available:
                lbl.configure(text=f"✓ 已就绪{(' v' + status.version) if status.version else ''}",
                              foreground=theme.SUCCESS)
            else:
                lbl.configure(text="✗ 未检测到（相关功能已禁用）", foreground=theme.ERROR)
        if self.on_saved:
            self.on_saved(st)


_ = os
