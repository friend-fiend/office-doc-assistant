"""ttk 主题与样式。

跨平台一致性的关键点：
  1. 显式设置字体。Windows 用 "Microsoft YaHei UI"，macOS 用 "PingFang SC"，
     否则 Tk 默认字体会让中文在 macOS 上偏小、在 Windows 上偏模糊。
  2. 行高统一。ttk.Treeview 默认行高在不同平台差异很大（macOS 明显更矮），
     这里统一设成 26，保证控件不错位。
  3. 不依赖第三方 ttkbootstrap 等主题库 —— 只用标准库 ttk，
     通过 ``Style.configure`` 自定义配色，避免多一个打包依赖。

配色思路（「有层次但不花」）：
  tkinter 画不了渐变、也没有圆角，所以层次只能靠**明度差**和**细边**来做。
  全局只有三档底色 —— 顶栏纯白 / 侧栏极浅蓝 / 内容区浅灰 —— 卡片纯白加一条
  1px 边浮在灰底上。品牌色（蓝）只出现在「当前步骤、主按钮、选中项」三处，
  其余全部中性。颜色一旦铺满，办公软件就会显廉价、也会重新变「眼花」。
"""
from __future__ import annotations

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

from ..config import IS_MACOS, IS_WINDOWS

# ---------------------------------------------------------------- 调色板（浅色 · 蓝调）
BG = "#f2f4f8"            # 内容区底色（浅灰）
PANEL = "#ffffff"         # 卡片
HEADER_BG = "#ffffff"     # 顶栏
SIDEBAR_BG = "#ebf1fb"    # 左侧导航（极浅蓝，暗示「这是导航，不是内容」）
BORDER = "#e4e7ee"
CARD_BORDER = "#e6e9f2"
TEXT = "#1b1f27"
TEXT_SUB = "#667085"
TEXT_HINT = "#98a2b3"
PRIMARY = "#2563eb"
PRIMARY_DARK = "#1d4ed8"
PRIMARY_LIGHT = "#eaf0fe"
PRIMARY_EDGE = "#c7d7fb"   # 拖拽区常态描边
PRIMARY_SOFT = "#f7faff"   # 拖拽区常态底色
PRIMARY_HOVER = "#e8f0ff"  # 拖拽区悬停底色
SUCCESS = "#0f9d58"
SUCCESS_DARK = "#0b7a44"
SUCCESS_BG = "#e6f4ec"
SUCCESS_EDGE = "#a8d8bf"
WARN = "#c2740a"
WARN_BG = "#fdf3e3"
ERROR = "#d93026"
ERROR_TEXT = "#b42318"
ERROR_BG = "#fdeceb"
SELECT_BG = "#eaf0fe"
PILL_BG = "#eef0f3"
HOVER = "#e9edf3"          # 内容区（灰底）上的悬停
HOVER_WHITE = "#f3f6fb"    # 白底上的悬停
SIDEBAR_HOVER = "#dfe8f7"

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


def _edge_style(name: str, fg: str, edge: str, hover_bg: str,
                base_bold: tuple, padding: tuple = (12, 5)) -> dict:
    """描边按钮的统一配置 —— 白底 + 1px 彩边 + 彩字，用于「次操作」。

    clam 主题下 ttk.Button 的边框由 border 元素绘制，颜色取
    ``bordercolor`` / ``lightcolor`` / ``darkcolor`` 三个参数；三个都设成同一色
    才能得到干净的 1px 纯色边（只设 bordercolor 会出现深浅不一的双色边）。
    """
    return {
        "configure": dict(background=PANEL, foreground=fg, font=base_bold,
                          borderwidth=1, relief="solid",
                          bordercolor=edge, lightcolor=edge, darkcolor=edge,
                          padding=padding),
        "map": dict(
            background=[("active", hover_bg), ("pressed", hover_bg), ("disabled", PANEL)],
            foreground=[("disabled", TEXT_HINT)],
            bordercolor=[("disabled", BORDER)],
            lightcolor=[("disabled", BORDER)],
            darkcolor=[("disabled", BORDER)],
        ),
    }


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
    # 卡片：白底 + 1px 细边（clam 的 TFrame 也认 bordercolor）
    style.configure("Card.TFrame", background=PANEL, relief="solid",
                    borderwidth=1, bordercolor=CARD_BORDER)
    style.configure("Header.TFrame", background=HEADER_BG)
    style.configure("Sidebar.TFrame", background=SIDEBAR_BG)
    style.configure("Flat.TFrame", background=PANEL)

    style.configure("TLabel", background=BG, foreground=TEXT, font=base)
    style.configure("Card.TLabel", background=PANEL, foreground=TEXT)
    style.configure("Sub.TLabel", background=BG, foreground=TEXT_SUB, font=base_small)
    style.configure("CardSub.TLabel", background=PANEL, foreground=TEXT_SUB, font=base_small)
    style.configure("Title.TLabel", background=HEADER_BG, foreground=TEXT, font=base_title)
    style.configure("HeaderSub.TLabel", background=HEADER_BG, foreground=TEXT_HINT,
                    font=base_small)
    style.configure("Section.TLabel", background=BG, foreground=TEXT, font=base_bold)
    style.configure("Section.Card.TLabel", background=PANEL, foreground=TEXT, font=base_bold)
    style.configure("Sidebar.TLabel", background=SIDEBAR_BG, foreground=TEXT)
    style.configure("SidebarTitle.TLabel", background=SIDEBAR_BG, foreground=TEXT_HINT,
                    font=(base[0], 9, "bold"))
    style.configure("TSeparator", background=BORDER)

    # ---- 按钮：按语义分四档，保证「主操作」全屏只有一处
    # 1) 主操作 —— 唯一的实心蓝
    style.configure("Accent.TButton", background=PRIMARY, foreground="#ffffff",
                    font=ui_font(11, True), borderwidth=0, padding=(26, 10))
    style.map("Accent.TButton",
              background=[("pressed", PRIMARY_DARK), ("active", PRIMARY_DARK),
                          ("disabled", "#d5dae4")],
              foreground=[("disabled", "#8f97a6")])

    # 2) 次操作 —— 白底描边（蓝 / 绿两色分别表示「主路径」与「已完成的产出」）
    cfg = _edge_style("Outline.TButton", PRIMARY, PRIMARY, PRIMARY_LIGHT, base_bold)
    style.configure("Outline.TButton", **cfg["configure"])
    style.map("Outline.TButton", **cfg["map"])

    cfg = _edge_style("Success.TButton", SUCCESS_DARK, SUCCESS_EDGE, SUCCESS_BG, base_bold)
    style.configure("Success.TButton", **cfg["configure"])
    style.map("Success.TButton", **cfg["map"])

    # 3) 辅助 —— 无边框文字按钮，hover 才浮出底色（默认不占视觉重量）
    style.configure("Ghost.TButton", background=BG, foreground=TEXT_SUB,
                    borderwidth=0, padding=(10, 5))
    style.map("Ghost.TButton",
              background=[("active", HOVER), ("pressed", HOVER)],
              foreground=[("active", TEXT), ("disabled", TEXT_HINT)])
    style.configure("GhostHeader.TButton", background=HEADER_BG, foreground=TEXT_SUB,
                    borderwidth=0, padding=(10, 5))
    style.map("GhostHeader.TButton",
              background=[("active", HOVER_WHITE), ("pressed", HOVER_WHITE)],
              foreground=[("active", TEXT)])

    # 4) 破坏性 —— 红字，hover 才出现浅红底
    style.configure("Danger.TButton", background=BG, foreground=ERROR_TEXT,
                    borderwidth=0, padding=(10, 5))
    style.map("Danger.TButton",
              background=[("active", ERROR_BG), ("pressed", ERROR_BG)],
              foreground=[("disabled", TEXT_HINT)])

    # ---------------- 步骤条：当前 = 蓝底白字，完成 = 绿底绿字，未到 = 灰字
    style.configure("Step.TLabel", background=BG, foreground=TEXT_HINT,
                    font=base_small, padding=(12, 6))
    style.configure("StepCur.TLabel", background=PRIMARY, foreground="#ffffff",
                    font=ui_font(10, True), padding=(12, 6))
    style.configure("StepDone.TLabel", background=SUCCESS_BG, foreground=SUCCESS_DARK,
                    font=ui_font(10, True), padding=(12, 6))

    # ---------------- 左侧导航（浅蓝底上的一套样式）
    style.configure("SidebarCat.TButton", background=SIDEBAR_BG, foreground=TEXT,
                    font=ui_font(10, True), anchor="w", padding=(12, 9), borderwidth=0)
    style.map("SidebarCat.TButton", background=[("active", SIDEBAR_HOVER)])
    style.configure("SidebarCatSel.TButton", background=PRIMARY_LIGHT, foreground=PRIMARY_DARK,
                    font=ui_font(10, True), anchor="w", padding=(12, 9), borderwidth=0)
    style.map("SidebarCatSel.TButton", background=[("active", PRIMARY_LIGHT)])

    style.configure("SidebarOp.TButton", background=SIDEBAR_BG, foreground=TEXT_SUB,
                    anchor="w", padding=(22, 6), borderwidth=0)
    style.map("SidebarOp.TButton",
              background=[("active", SIDEBAR_HOVER)],
              foreground=[("active", TEXT)])
    style.configure("SidebarOpSel.TButton", background=PRIMARY_LIGHT, foreground=PRIMARY_DARK,
                    anchor="w", padding=(22, 6), borderwidth=0, font=ui_font(9, True))
    style.map("SidebarOpSel.TButton", background=[("active", PRIMARY_LIGHT)])

    # ---- 输入控件
    style.configure("TEntry", fieldbackground=PANEL, foreground=TEXT,
                    bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                    insertcolor=TEXT, padding=5)
    style.map("TEntry", bordercolor=[("focus", PRIMARY)],
              lightcolor=[("focus", PRIMARY)], darkcolor=[("focus", PRIMARY)])
    style.configure("TCombobox", fieldbackground=PANEL, background=PANEL,
                    foreground=TEXT, arrowcolor=TEXT_SUB, padding=4,
                    bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
    style.map("TCombobox",
              fieldbackground=[("readonly", PANEL), ("disabled", "#f2f3f5")],
              bordercolor=[("focus", PRIMARY)],
              lightcolor=[("focus", PRIMARY)], darkcolor=[("focus", PRIMARY)])
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
    style.configure("TProgressbar", background=PRIMARY, troughcolor="#e3e6ec",
                    bordercolor="#e3e6ec", lightcolor=PRIMARY, darkcolor=PRIMARY,
                    thickness=7)

    # ---- Treeview（文件列表）
    style.configure("Treeview", background=PANEL, fieldbackground=PANEL,
                    foreground=TEXT, rowheight=26, font=base, borderwidth=0)
    style.map("Treeview",
              background=[("selected", SELECT_BG)],
              foreground=[("selected", PRIMARY_DARK)])
    style.configure("Treeview.Heading", background="#f5f7fb", foreground=TEXT_SUB,
                    font=base_small, relief="flat", padding=(8, 6), borderwidth=0)
    style.map("Treeview.Heading", background=[("active", "#eaeef6")])

    # ---- 滚动条（细一点，视觉上更现代）
    style.configure("Vertical.TScrollbar", background="#d4d8e0", troughcolor=BG,
                    bordercolor=BG, arrowcolor=TEXT_SUB, width=10, relief="flat")
    style.map("Vertical.TScrollbar", background=[("active", "#b9bfca")])
    style.configure("Horizontal.TScrollbar", background="#d4d8e0", troughcolor=BG,
                    bordercolor=BG, arrowcolor=TEXT_SUB, relief="flat")
    style.map("Horizontal.TScrollbar", background=[("active", "#b9bfca")])

    # ---- 布局辅助
    style.configure("TLabelframe", background=BG, bordercolor=BORDER, relief="solid",
                    borderwidth=1)
    style.configure("TLabelframe.Label", background=BG, foreground=TEXT_SUB, font=base_small)

    # ---- 拖拽区专用配色（用 tk.Frame 因为要绑定事件）
    style.configure("Drop.TLabel", background=PRIMARY_SOFT, foreground=TEXT_SUB, font=base)

    return style


DROPI_IDLE = PRIMARY_SOFT
DROPI_BG = PRIMARY_HOVER
DROPI_BORDER = PRIMARY
