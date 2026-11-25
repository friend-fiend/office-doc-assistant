"""ttk 主题与样式。

跨平台一致性的关键点：
  1. 显式设置字体。Windows 用 "Microsoft YaHei UI"，macOS 用 "PingFang SC"，
     否则 Tk 默认字体会让中文在 macOS 上偏小、在 Windows 上偏模糊。
  2. 行高统一。ttk.Treeview 默认行高在不同平台差异很大（macOS 明显更矮），
     这里统一设成 26，保证控件不错位。
  3. 不依赖第三方 ttkbootstrap 等主题库 —— 只用标准库 ttk，
     通过 ``Style.configure`` 自定义配色，避免多一个打包依赖。
"""
from __future__ import annotations

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

from ..config import IS_MACOS, IS_WINDOWS

# ---------------------------------------------------------------- 调色板（浅色）
BG = "#f7f8fa"
PANEL = "#ffffff"
BORDER = "#e8eaee"
TEXT = "#1f2329"
TEXT_SUB = "#6b7280"
TEXT_HINT = "#9ca3af"
PRIMARY = "#2563eb"
PRIMARY_DARK = "#1d4ed8"
PRIMARY_LIGHT = "#eaf0fe"
SUCCESS = "#0f9d58"
SUCCESS_BG = "#e6f4ec"
WARN = "#d97706"
WARN_BG = "#fdf3e3"
ERROR = "#d93026"
SELECT_BG = "#eaf0fe"
PILL_BG = "#eef0f3"          # 顶栏状态药丸底色（中性）
HOVER = "#eceef1"

MONO = ("Consolas", "Menlo", "monospace")


def ui_font(size: int = 10, bold: bool = False) -> tuple:
    family = "Microsoft YaHei UI" if IS_WINDOWS else ("PingFang SC" if IS_MACOS else "DejaVu Sans")
    return (family, size, "bold" if bold else "normal")


def mono_font(size: int = 9) -> tuple:
    return (MONO[0] if IS_WINDOWS else MONO[1], size, "normal")


def pick_default_font(root: tk.Misc) -> tuple:
    """从系统已有字体中挑一个能显示中文的，尽量避免方块字。"""
    try:
        families = set(tkfont.families(root))
    except Exception:
        return ui_font()
    for name in ("Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC",
                 "Hiragino Sans GB", "Noto Sans CJK SC", "Source Han Sans SC",
                 "WenQuanYi Micro Hei", "SimHei"):
        if name in families:
            return (name, 10)
    return ui_font()


def apply_theme(root: tk.Tk) -> ttk.Style:
    style = ttk.Style(root)
    # clam 是唯一在三大平台上都支持完整配色改写的内置主题（aqua/vista 会锁死部分颜色）
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    base = pick_default_font(root)
    base_small = (base[0], 9)
    base_bold = (base[0], 10, "bold")
    base_title = (base[0], 13, "bold")

    root.configure(bg=BG)
    root.option_add("*Font", base)
    root.option_add("*TCombobox*Listbox.font", base)

    # ---- 基础容器
    style.configure(".", background=BG, foreground=TEXT, font=base,
                    borderwidth=0, focuscolor=BG)
    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=PANEL, relief="flat")
    style.configure("TLabel", background=BG, foreground=TEXT, font=base)
    style.configure("Card.TLabel", background=PANEL, foreground=TEXT)
    style.configure("Sub.TLabel", background=BG, foreground=TEXT_SUB, font=base_small)
    style.configure("CardSub.TLabel", background=PANEL, foreground=TEXT_SUB, font=base_small)
    style.configure("Title.TLabel", background=BG, foreground=TEXT, font=base_title)
    style.configure("Section.TLabel", background=BG, foreground=TEXT, font=base_bold)
    style.configure("TSeparator", background=BORDER)

    # ---- 按钮
    style.configure("TButton", background=PANEL, foreground=TEXT, font=base,
                    borderwidth=1, relief="solid", padding=(10, 5))
    style.map("TButton",
              background=[("pressed", "#eceef1"), ("active", "#f0f2f5"),
                          ("disabled", "#f2f3f5")],
              foreground=[("disabled", "#a8adb5")])
    style.configure("Accent.TButton", background=PRIMARY, foreground="#ffffff",
                    font=ui_font(10, True), borderwidth=0, padding=(22, 8))
    style.map("Accent.TButton",
              background=[("pressed", PRIMARY_DARK), ("active", PRIMARY_DARK),
                          ("disabled", "#b9c6e8")],
              foreground=[("disabled", "#f0f0f0")])
    style.configure("Ghost.TButton", background=BG, foreground=TEXT_SUB,
                    borderwidth=0, padding=(8, 4))
    style.map("Ghost.TButton", background=[("active", HOVER)])

    # 左侧分类按钮（手风琴标题）
    style.configure("Cat.TButton", background=PANEL, foreground=TEXT,
                    font=ui_font(10, True), anchor="w", padding=(12, 8), borderwidth=0)
    style.map("Cat.TButton", background=[("active", HOVER)])
    style.configure("CatSel.TButton", background=PRIMARY_LIGHT, foreground=PRIMARY_DARK,
                    font=ui_font(10, True), anchor="w", padding=(12, 8), borderwidth=0)
    style.map("CatSel.TButton", background=[("active", PRIMARY_LIGHT)])
    # 左侧功能按钮（选中态用 Primary 配色）
    style.configure("Op.TButton", background=PANEL, foreground=TEXT,
                    anchor="w", padding=(24, 6), borderwidth=0)
    style.map("Op.TButton", background=[("active", HOVER)])
    style.configure("OpSel.TButton", background=PRIMARY_LIGHT, foreground=PRIMARY_DARK,
                    anchor="w", padding=(24, 6), borderwidth=0, font=ui_font(9, True))
    style.map("OpSel.TButton", background=[("active", PRIMARY_LIGHT)])

    # ---- 输入控件
    style.configure("TEntry", fieldbackground=PANEL, foreground=TEXT,
                    bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                    insertcolor=TEXT, padding=4)
    style.map("TEntry", bordercolor=[("focus", PRIMARY)])
    style.configure("TCombobox", fieldbackground=PANEL, background=PANEL,
                    foreground=TEXT, arrowcolor=TEXT_SUB, padding=3,
                    bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
    style.map("TCombobox",
              fieldbackground=[("readonly", PANEL), ("disabled", "#f2f3f5")],
              bordercolor=[("focus", PRIMARY)])
    style.configure("TCheckbutton", background=BG, foreground=TEXT, font=base)
    style.map("TCheckbutton", background=[("active", BG)])
    style.configure("Card.TCheckbutton", background=PANEL, foreground=TEXT)
    style.map("Card.TCheckbutton", background=[("active", PANEL)])

    # ---- Notebook
    style.configure("TNotebook", background=BG, borderwidth=0)
    style.configure("TNotebook.Tab", background=BG, foreground=TEXT_SUB,
                    padding=(14, 7), font=base)
    style.map("TNotebook.Tab",
              background=[("selected", PANEL)],
              foreground=[("selected", TEXT)])

    # ---- 进度条
    style.configure("TProgressbar", background=PRIMARY, troughcolor="#e6e8eb",
                    bordercolor="#e6e8eb", lightcolor=PRIMARY, darkcolor=PRIMARY,
                    thickness=8)

    # ---- Treeview（文件列表）
    style.configure("Treeview", background=PANEL, fieldbackground=PANEL,
                    foreground=TEXT, rowheight=26, font=base, borderwidth=0)
    style.map("Treeview",
              background=[("selected", SELECT_BG)],
              foreground=[("selected", TEXT)])
    style.configure("Treeview.Heading", background="#f0f1f3", foreground=TEXT_SUB,
                    font=base_small, relief="flat", padding=(6, 5), borderwidth=0)
    style.map("Treeview.Heading", background=[("active", "#e6e8eb")])

    # ---- 滚动条（细一点，视觉上更现代）
    style.configure("Vertical.TScrollbar", background="#d4d6da", troughcolor=BG,
                    bordercolor=BG, arrowcolor=TEXT_SUB, width=10, relief="flat")
    style.map("Vertical.TScrollbar", background=[("active", "#b9bcc2")])
    style.configure("Horizontal.TScrollbar", background="#d4d6da", troughcolor=BG,
                    bordercolor=BG, arrowcolor=TEXT_SUB, relief="flat")

    # ---- 布局辅助
    style.configure("TLabelframe", background=BG, bordercolor=BORDER, relief="solid",
                    borderwidth=1)
    style.configure("TLabelframe.Label", background=BG, foreground=TEXT_SUB, font=base_small)

    # ---- 拖拽区专用配色（用 tk.Frame 因为要绑定事件）
    style.configure("Drop.TLabel", background=PANEL, foreground=TEXT_SUB, font=base)

    return style


DROPI_IDLE = PANEL
DROPI_BG = "#eef3ff"
DROPI_BORDER = PRIMARY
