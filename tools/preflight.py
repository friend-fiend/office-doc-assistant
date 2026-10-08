#!/usr/bin/env python3
"""启动前环境体检 —— 只读、不弹窗、不启动界面。

为什么单独做成一个脚本
----------------------
程序最常见的"装完跑不起来"就是缺 Tkinter。此时主程序自己都弹不出窗口，
所以体检脚本必须 **零 Tkinter 依赖**（本文件确实没有 import tkinter），
才能在任何环境下给出可读的诊断报告。

用法
----
    python main.py --check          # 简明清单
    python main.py --doctor         # 附带详细路径与版本
    python tools/preflight.py -v    # 直接调用

退出码
------
    0  硬性条件全部满足，可以启动
    1  存在硬性缺失（缺 Python 版本 / Tkinter / 核心库）
    2  脚本自身异常
"""
from __future__ import annotations

import importlib
import os
import sys
from dataclasses import dataclass, field

# 让 `python tools/preflight.py` 也能找到 app 包
_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)


# ====================================================================== 数据结构
OK = "OK"
WARN = "WARN"
FAIL = "FAIL"

_MARK = {OK: "[ OK ]", WARN: "[警告]", FAIL: "[失败]"}


@dataclass
class Check:
    name: str
    status: str = OK
    detail: str = ""
    hint: str = ""
    # 硬件/硬性：FAIL 会导致退出码 1
    hard: bool = False


@dataclass
class Report:
    checks: list[Check] = field(default_factory=list)

    def add(self, name, status=OK, detail="", hint="", hard=False):
        self.checks.append(Check(name, status, detail, hint, hard))

    @property
    def hard_failed(self) -> list[Check]:
        return [c for c in self.checks if c.hard and c.status == FAIL]

    @property
    def warnings(self) -> list[Check]:
        return [c for c in self.checks if c.status == WARN]

    def render(self, verbose: bool = False) -> str:
        lines: list[str] = []
        lines.append("=" * 68)
        lines.append("  办公文档一键处理助手 —— 环境体检")
        lines.append("=" * 68)
        for c in self.checks:
            lines.append(f"{_MARK[c.status]} {c.name}")
            if c.detail and (verbose or c.status != OK):
                lines.append(f"        {c.detail}")
            if c.hint and c.status in (FAIL, WARN):
                for ln in c.hint.splitlines():
                    lines.append(f"        → {ln}")
        lines.append("-" * 68)
        if self.hard_failed:
            # 不说"程序无法启动"：硬性缺失里既有真的起不来的（Tkinter），
            # 也有"能起来但关键功能废掉"的（PyMuPDF 渲染引擎）。
            # 统一说成"必须先修复、修好前不要打包分发"，两种情况都成立。
            names = "、".join(c.name for c in self.hard_failed)
            lines.append(f"结论：存在 {len(self.hard_failed)} 项硬性缺失（{names}），"
                         f"必须修复；修复前请勿打包分发。")
        elif self.warnings:
            lines.append(f"结论：可以启动，但有 {len(self.warnings)} 项提示（部分功能会禁用或降级）。")
        else:
            lines.append("结论：环境完备，全部功能可用。")
        lines.append("=" * 68)
        return "\n".join(lines)


# ====================================================================== 各检查项
def _check_python(r: Report) -> None:
    vi = sys.version_info
    ver = f"{vi.major}.{vi.minor}.{vi.micro}"
    on_win = sys.platform.startswith("win")
    is_win81 = False
    if on_win:
        # 6.3 = Windows 8.1；10.0 = Windows 10/11
        try:
            is_win81 = sys.getwindowsversion().major == 6 and sys.getwindowsversion().minor == 3
        except Exception:
            is_win81 = False

    if vi < (3, 10):
        r.add(
            "Python 版本", FAIL, f"{ver}（要求 >= 3.10）",
            "请安装 Python 3.10 或 3.11：https://www.python.org/downloads/",
            hard=True,
        )
        return

    if is_win81 and vi >= (3, 13):
        r.add(
            "Python 版本", FAIL, f"{ver} 不支持 Windows 8.1（3.13 起官方已放弃）",
            "Windows 8.1 请改装 Python 3.11.x（3.11 是最后一个支持 8.1 的次要版本）",
            hard=True,
        )
        return

    note = ""
    if vi >= (3, 13):
        note = "；注意 3.13+ 已不支持 Windows 8.1"
    elif vi >= (3, 12):
        note = "；推荐 3.11.x 以覆盖 Windows 8.1"
    r.add("Python 版本", OK, f"{ver}（{sys.executable}）{note}")


def _check_tkinter(r: Report) -> None:
    try:
        from app.native_alert import check_tkinter, tkinter_help_text
    except Exception as exc:
        r.add("Tkinter", FAIL, f"无法加载检查模块：{exc}", hard=True)
        return

    ok, detail = check_tkinter()
    if ok:
        r.add("Tkinter（界面必需）", OK, detail)
    else:
        r.add(
            "Tkinter（界面必需）", FAIL, detail,
            tkinter_help_text(),
            hard=True,
        )


# (导入名, 显示名, 是否硬性, 缺了会怎样)
_LIBS: tuple[tuple[str, str, bool, str], ...] = (
    ("pypdf", "pypdf（PDF 读/写/合并/拆分）", True, "PDF 相关功能全部不可用"),
    ("reportlab", "reportlab（中文页码/水印）", True, "页码、水印、页眉页脚不可用"),
    ("PIL", "Pillow（图片处理）", True, "图片转 PDF、PDF 转图片不可用"),
    ("openpyxl", "openpyxl（Excel）", False, "Excel 合并/拆分不可用"),
    ("docx", "python-docx（Word 读取）", False, "文档对比、Word 转 PDF 受限"),
    ("pptx", "python-pptx（PPT 读取）", False, "PPT 转 PDF 受限"),
    ("py7zr", "py7zr（7z 支持）", False, "7z 压缩/解压不可用（ZIP 仍可用）"),
    ("pdf2image", "pdf2image（PDF 转图片）", False, "PDF 转图片不可用"),
    # pymupdf 标记为硬性：它是 PDF 导出 PDF / 批量打印的渲染后端，
    # 且打包时最容易漏收那对 26 MB+13 MB 的二进制（mupdfcpp64.dll / _mupdf.pyd）。
    # 声明为硬性后，冻结产物一旦缺它，`--check` 直接返回非零，
    # build_win.bat / CI 就会当场中止，不会把坏包发出去。
    ("pymupdf", "PyMuPDF（PDF 导出 / 批量打印渲染）", True,
     "PDF 导出 PDF、PDF 批量打印不可用"),
)


def _check_libs(r: Report) -> None:
    for mod, title, hard, impact in _LIBS:
        try:
            m = importlib.import_module(mod)
            ver = getattr(m, "__version__", "") or getattr(m, "Version", "") or ""
            r.add(title, OK, f"v{ver}" if ver else "已安装")
        except Exception as exc:
            r.add(
                title, FAIL if hard else WARN, f"导入失败：{type(exc).__name__}: {exc}",
                f"{impact}。修复：pip install -r requirements.txt",
                hard=hard,
            )

    # 可选：cryptography（决定 PDF 加密强度）
    try:
        import cryptography
        r.add("cryptography（AES 加密）", OK, f"v{cryptography.__version__}")
    except Exception:
        r.add(
            "cryptography（AES 加密）", WARN, "未安装，PDF 加密将降级为 RC4-128",
            "pip install cryptography   （或用 pypdf[crypto]）",
        )

    _check_pymupdf_render(r)


def _check_pymupdf_render(r: Report) -> None:
    """真正渲染一次，确认 mupdfcpp64.dll 能加载。

    为什么不能只 import：pymupdf 的 __init__ 是纯 Python，import 成功不代表
    C 扩展 _mupdf.pyd 和它依赖的 mupdfcpp64.dll 能找到。打包时漏收那个 26 MB 的
    DLL，import 依然返回成功，只有真正开文档/渲染时才抛
    "ImportError: DLL load failed" —— 用户点下按钮才炸。
    所以这里在内存里造一页 PDF 再渲染成 pixmap，几毫秒就能验到底。
    """
    try:
        import pymupdf
    except Exception:
        return  # 上面 _LIBS 已经报过缺失了，不重复刷屏

    try:
        doc = pymupdf.open()
        page = doc.new_page(width=200, height=100)
        page.insert_text((20, 50), "OK")
        pix = page.get_pixmap(dpi=36)
        n = pix.width * pix.height
        doc.close()
        if n <= 0:
            raise RuntimeError("渲染结果为空")
        r.add("PyMuPDF 渲染引擎", OK, "C 扩展与 mupdfcpp64.dll 均已就位")
    except Exception as exc:
        r.add(
            "PyMuPDF 渲染引擎", FAIL,
            f"{type(exc).__name__}: {exc}",
            "PDF 导出 PDF、PDF 批量打印会失败。若是打包版，通常说明 "
            "pymupdf 的 mupdfcpp64.dll / _mupdf.pyd 没被收进 _internal/pymupdf/，"
            "请检查 build/office_doc_assistant.spec 的收集团块。",
            hard=True,
        )


def _check_dnd(r: Report) -> None:
    """拖拽是可选能力：Tkinter 没有原生拖放，必须靠第三方扩展。"""
    try:
        import tkinterdnd2  # noqa: F401

        r.add("tkinterdnd2（拖拽支持）", OK, "已安装，可使用拖拽添加文件")
    except Exception:
        r.add(
            "tkinterdnd2（拖拽支持）", WARN, "未安装，将退化为「点击选择文件」",
            "pip install tkinterdnd2   （可选；标准库本身无法实现拖拽）",
        )


def _check_external(r: Report) -> None:
    try:
        from app.deps import TOOL_SPECS, detect_all
    except Exception as exc:
        r.add("外部程序探测", WARN, f"探测模块加载失败：{exc}")
        return

    try:
        statuses = detect_all()
    except Exception as exc:
        r.add("外部程序探测", WARN, f"探测过程出错：{exc}")
        return

    for spec in TOOL_SPECS:
        st = statuses[spec.key]
        if st.available:
            r.add(spec.title, OK, st.display)
        else:
            # 全部外部程序都是可选：缺失只影响部分功能
            r.add(
                spec.title, WARN, f"未检测到（影响：{'、'.join(spec.affects)}）",
                spec.help_text,
            )


def _check_vendor(r: Report) -> None:
    """绿色便携版：把外部程序放进 vendor/ 就会被自动识别。"""
    try:
        from app.config import app_root
    except Exception:
        return
    for name in ("poppler", "soffice", "ghostscript"):
        p = os.path.join(app_root(), "vendor", name)
        if os.path.isdir(p):
            r.add(f"vendor/{name}", OK, p)


def _check_writable(r: Report) -> None:
    """配置目录必须可写，否则设置存不下来。

    注意：判定标准是"能不能写成功"，删除探测文件只是**尽力而为的善后**。
    某些受限环境（沙盒 / 回收站不可用 / 只读挂载）下删除会失败，
    这不代表目录不可写，因此不能因为删不掉就报 WARN。
    """
    try:
        from app.settings import config_dir

        d = str(config_dir())
        probe = os.path.join(d, ".write_probe")
        with open(probe, "w", encoding="utf-8") as fh:
            fh.write("ok")
        try:
            os.remove(probe)
        except OSError:
            pass  # 清理失败不影响结论
        r.add("配置目录可写", OK, d)
    except Exception as exc:
        r.add(
            "配置目录可写", WARN, f"{type(exc).__name__}: {exc}",
            "设置将无法保存，但功能可正常使用。",
            hard=False,
        )


def _check_windows_extras(r: Report) -> None:
    if not sys.platform.startswith("win"):
        return

    # ----------------------------------------------------------------
    # pywin32：PDF 批量打印必需
    # ----------------------------------------------------------------
    # 这三个模块是 .pyd 扩展，依赖 pywin32_system32/ 下的 pythoncom3xx.dll、
    # pywintypes3xx.dll。打包装箱时最容易漏掉的就是这两个 DLL——
    # 表现是 import 直接抛 "DLL load failed"，所以这里逐个 import 实测，
    # 而不是只看包名在不在。
    try:
        import win32con  # noqa: F401
        import win32print  # noqa: F401
        import win32ui  # noqa: F401

        try:
            printers = win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS,
                1,
            )
            n = len(printers)
        except Exception:
            n = 0
        try:
            from app.core.print_ops import (
                list_printers, list_virtual_pdf_printers, _printer_default_summary,
            )

            real = list_printers()
            virtual = list_virtual_pdf_printers()
            # 只读地把默认 DEVMODE 读一遍：如果这里就报错，说明打印功能
            # 在真机上会有问题，早点暴露比出纸时才发现好。
            if real:
                info = _printer_default_summary(real[0])
                paper = info.get("paper")
                extra = f"，{real[0]} 默认纸张 ID={paper}" if paper else ""
            else:
                extra = ""
            r.add(
                "打印机枚举", OK,
                f"真实打印机 {len(real)} 台，虚拟打印设备 {len(virtual)} 台已隐藏{extra}",
                "真实打印必须在界面上确认后才会执行；测试/脚本路径默认被拒绝。"
                "如需完全关闭出纸，可勾选「仅演练不打印」或设置 ODA_NO_PRINT=1。",
            )
        except Exception as exc:
            r.add("打印机枚举", WARN, f"读取失败：{type(exc).__name__}: {exc}")

        r.add("pywin32（PDF 批量打印）", OK,
              f"已就绪，检测到 {n} 个打印机" if n else "已就绪")
    except Exception as exc:
        r.add(
            "pywin32（PDF 批量打印）", WARN,
            f"导入失败：{type(exc).__name__}: {exc}",
            "PDF 批量打印不可用（其余功能不受影响）。"
            "修复：pip install pywin32  然后运行 .venv\\Scripts\\python.exe "
            "Scripts\\pywin32_postinstall.py -install",
        )

    try:
        from app.paths import is_windows_long_path_enabled

        if is_windows_long_path_enabled():
            r.add("Windows 长路径支持", OK, "已启用（>260 字符路径可用）")
        else:
            r.add(
                "Windows 长路径支持", WARN, "未启用，超过 260 字符的路径可能失败",
                "程序内部会自动加 \\\\?\\ 前缀规避；如需彻底开启，"
                "运行 reg add \"HKLM\\SYSTEM\\CurrentControlSet\\Control\\FileSystem\" "
                "/v LongPathsEnabled /t REG_DWORD /d 1 /f（需管理员）",
            )
    except Exception:
        pass


# ====================================================================== 入口
def _check_version_resource(r: Report) -> None:
    """核对「程序自报版本」与「exe 版本资源」是否一致。

    为什么值得单独一项：这两处曾经真的对不上，而且不报错、不告警。
    version_info.txt 原本是手写文件，spec 里又写了一份 VERSION，于是同一个 exe
    有了两个版本来源。1.0.1 → 1.0.2 那次只顺手改了字符串表里的 "1.0.2.0"，
    ffi 的 filevers / prodvers 元组还是 (1, 0, 1, 0)。结果：

        · 右键「属性 → 详细信息」显示 1.0.2.0（读字符串表）
        · win32api.GetFileVersionInfo 读出 1.0.1.0（读 ffi 元组）

    企业资产盘点、升级决策、签名校验读的都是后者 —— 版本号判错还一声不响。
    现在版本资源由 spec 按 VERSION 现场生成，这里再交叉校验一次兜底，
    并且**在冻结产物上直接判 FAIL**（hard=True），让 build_win.bat / CI 当场中止，
    坏包发不出去。
    """
    try:
        from app.config import VERSION as APP_VERSION
    except Exception:
        return  # 拿不到自报版本就不下结论，不要制造假告警

    if not sys.platform.startswith("win"):
        # macOS 的版本写进 Info.plist（BUNDLE(version=VERSION)），路径不同。
        r.add("版本资源一致性", OK, f"当前版本 {APP_VERSION}（macOS，由 Info.plist 承载）")
        return

    if not bool(getattr(sys, "frozen", False)):
        # 开发态：exe 是 python.exe，它的版本资源和本程序无关，跳过。
        r.add("版本资源一致性", OK, f"开发态，当前版本 {APP_VERSION}")
        return

    nums = [p for p in APP_VERSION.split(".") if p.strip()]
    expected = ".".join((nums + ["0", "0", "0", "0"])[:4])
    try:
        import win32api

        info = win32api.GetFileVersionInfo(sys.executable, "\\")
        ms, ls = info["FileVersionMS"], info["FileVersionLS"]
        actual = f"{ms >> 16}.{ms & 0xFFFF}.{ls >> 16}.{ls & 0xFFFF}"
    except Exception as exc:
        r.add("版本资源一致性", WARN,
              f"读不到 exe 的版本资源：{type(exc).__name__}: {exc}",
              "不影响功能，但无法确认版本号是否可信。")
        return

    if actual == expected:
        r.add("版本资源一致性", OK,
              f"exe 版本资源 {actual} 与程序自报 {APP_VERSION} 一致")
    else:
        r.add(
            "版本资源一致性", FAIL,
            f"exe 版本资源是 {actual}，但程序自报 {APP_VERSION}，两者对不上",
            "版本资源由 build/office_doc_assistant.spec 按 VERSION 现场生成。"
            "修复：删除 build/version_info.txt 后重新打包（构建时会自动重建）。"
            "修复前请勿分发 —— 用户和资产盘点工具读到的版本号是错的。",
            hard=True,
        )


def main(verbose: bool = False) -> int:
    try:
        r = Report()
        _check_python(r)
        _check_tkinter(r)
        _check_libs(r)
        _check_dnd(r)
        _check_external(r)
        _check_vendor(r)
        _check_writable(r)
        _check_version_resource(r)
        _check_windows_extras(r)
    except Exception:
        import traceback

        print("体检脚本自身异常：\n" + traceback.format_exc(), file=sys.stderr)
        return 2

    print(r.render(verbose=verbose))
    return 1 if r.hard_failed else 0


if __name__ == "__main__":
    sys.exit(main(verbose=any(a in ("-v", "--verbose", "--doctor") for a in sys.argv[1:])))
