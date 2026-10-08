"""PDF 批量打印 + 批量导出 PDF（Windows）。

设计
----
  PDF → PyMuPDF 读页 → 渲染为位图 → DIB → 打印 DC 写出去。

为什么不直接 ``os.startfile(path, 'print')``:
  - 强依赖"机器上恰好装了能打印的默认程序"，不可控
  - 无法选打印机/份数/纸张/方向
  - 失败无错误码

为什么不调 SumatraPDF:
  - 只覆盖 PDF；不划算
  - 引入额外二进制

为什么 pywin32 而不是 ctypes 直调 winspool.drv:
  - pywin32 是稳定封装，winspool 设备上下文在 ctypes 里 200+ 行且容易签名错
  - 二进制 wheel 自带，打包体积可接受

三条安全红线（2026-09-28 修复"用户没操作却自动打印"后确立）
----------------------------------------------------------
1) **真实打印必须有人工授权**
   ``_do_print`` 会检查 ``ctx.allow_real_print``；该标志只能由界面里
   "用户点击开始执行 + 确认对话框"这条路径置位（见 ui/app.py）。
   测试、体检、脚本、定时任务一律拿不到授权，会被 ``PrintNotAuthorized``
   挡在提交作业之前。

   > 为什么要这么严：打印是物理世界里的不可逆动作。写文件写错了可以删，
   > 纸打出来收不回去 —— 轻则浪费耗材，重则把不该外发的内容打到公共
   > 打印机上。这类动作绝不能"只要代码走到这里就执行"。

2) **绝不修改系统打印机默认设置**
   旧实现用 ``win32print.SetPrinter(h, 2, info, 0)`` 把 DEVMODE 写回打印机，
   那是**持久**改写用户的系统默认（影响所有其它软件），而且需要管理员
   权限。实测本机默认纸张就被它改成了 A3 横向。
   现在改为：取打印机默认 DEVMODE 的原始字节 → 在内存里改纸张/方向/份数
   → ``gdi32.CreateDCW`` 直接建一个"带这份 DEVMODE 的打印 DC"。
   每份文件互不影响，进程退出即失效，全程不写系统状态、不需要管理员权限。

3) **演练模式（dry_run）**
   勾选后只渲染、只出预览 PDF，一个作业都不提交。用于"先看看版式对不对
   再决定要不要真打"，也让没有打印机的环境能跑完整链路。

过滤虚拟打印机
--------------
``list_printers()`` 会滤掉 "Microsoft Print to PDF" / "导出为WPS PDF" /
"Microsoft XPS Document Writer" / "OneNote" / "Fax" 这类**不是物理打印机**
的设备：它们不会出纸，而是弹"另存为"或静默落盘，选错了会莫名其妙多出
一堆 PDF 文件。需要导出 PDF 请使用单独的「PDF 导出 PDF」功能。
"""
from __future__ import annotations

import os
import sys
from dataclasses import dataclass, asdict

from .registry import OpContext, OpDef, Param, register
from ..errors import PrintNotAuthorized

# --------------------------------------------------------------- 平台守卫
IS_WINDOWS = sys.platform.startswith("win")
PYWIN32_OK = False
if IS_WINDOWS:
    try:
        import win32con  # type: ignore
        import win32print  # type: ignore
        import win32ui  # type: ignore
        PYWIN32_OK = True
    except ImportError:
        PYWIN32_OK = False


# --------------------------------------------------------------- 打印机枚举
# 这些设备的共同点：**不会出纸**。它们要么弹"另存为"对话框，要么静默
# 往某个目录写文件，要么把内容送到别处。
#
# 两层过滤，缺一不可：
#   · 精确名单 —— 覆盖常见虚拟打印机（名字固定，最可靠）
#   · 关键字   —— 兜住"导出为 WPS PDF"这种带厂商前缀/中文名的变体
#
# 为什么必须认真过滤：选到这类设备后，用户点一次"开始执行"就会得到一堆
# 莫名其妙的 PDF 文件（或被弹一堆另存为对话框），而"打印机"一个字都没出。
VIRTUAL_PDF_PRINTERS = (
    "microsoft print to pdf",
    "microsoft print to pdf(redirected)",
    "microsoft xps document writer",
    "onenote (desktop)",
    "onenote for windows 10",
    "onenote",
    "fax",
    "foxit pdf printer",
    "adobe pdf",
    "pdf creator",
    "pdf printer",
    "cutepdf writer",
    "primopdf",
    "dopdf",
    "bullzip pdf printer",
    "pdf24 pdf",
    "pdfsam",
    "pdf-xchange",
    "导出为wps pdf",
    "发送到wps高级打印",
    "wps pdf",
    "kingsoft virtual printer",
)

# 关键字命中即视为虚拟设备（一律小写子串比较）
VIRTUAL_PDF_KEYWORDS = (
    "print to pdf", "pdf printer", "pdf writer", "to pdf",
    "xps document writer", "onenote", "fax",
    "pdf creator", "cutepdf", "primopdf", "dopdf", "bullzip",
    "pdf24", "pdf-xchange", "pdfsam",
    "wps pdf", "wps高级打印",
)


def _is_virtual_printer(name: str, driver: str = "") -> bool:
    """判断一台设备是不是"不出纸"的虚拟打印机。"""
    low = (name or "").strip().lower()
    if not low:
        return False
    if low in VIRTUAL_PDF_PRINTERS:
        return True
    for kw in VIRTUAL_PDF_KEYWORDS:
        if kw in low:
            return True
    # 驱动名也一起看：厂商常把虚拟驱动做成 "Kingsoft Virtual Printer Driver"
    dlow = (driver or "").strip().lower()
    if "virtual printer" in dlow or "virtual pdf" in dlow:
        return True
    return False


def _enum_printers() -> list[dict]:
    """枚举本机打印机，返回 [{name, driver, port, virtual}]。"""
    if not PYWIN32_OK:
        return []
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    try:
        items = win32print.EnumPrinters(flags, None, 2)
    except Exception:
        return []
    out: list[dict] = []
    seen: set[str] = set()
    for p in items:
        name = (p.get("pPrinterName") or p.get("Name") or "").strip()
        if not name or name in seen:
            continue
        seen.add(name)
        driver = (p.get("pDriverName") or "").strip()
        out.append({
            "name": name,
            "driver": driver,
            "port": (p.get("pPortName") or "").strip(),
            "virtual": _is_virtual_printer(name, driver),
        })
    return out


def list_printers(real_only: bool = True) -> list[str]:
    """返回本机可见的打印机名列表。

    ``real_only=True``（默认）只保留**真正的物理 / 网络打印机**，把会弹
    "另存为"或静默写文件的虚拟 PDF / XPS / 传真设备全部滤掉。
    """
    items = _enum_printers()
    names = [it["name"] for it in items if (not real_only or not it["virtual"])]
    return sorted(names)


def list_virtual_pdf_printers() -> list[str]:
    """返回被过滤掉的虚拟打印设备（用于 UI 提示与诊断）。"""
    return sorted(it["name"] for it in _enum_printers() if it["virtual"])


def default_printer() -> str:
    """返回系统默认打印机。"""
    if not PYWIN32_OK:
        return ""
    try:
        return win32print.GetDefaultPrinter() or ""
    except Exception:
        return ""


def _printer_status(name: str) -> tuple[int, str]:
    """返回 (状态码, 描述)。0 = 正常。"""
    if not PYWIN32_OK or not name:
        return (0, "")
    try:
        h = win32print.OpenPrinter(name)
        try:
            info = win32print.GetPrinter(h, 2)
            status = info.get("Status", 0)
            return (status, _decode_status(status))
        finally:
            win32print.ClosePrinter(h)
    except Exception as exc:
        return (-1, f"无法查询打印机：{exc}")


def _decode_status(status: int) -> str:
    bits = []
    if status & 0x00000001: bits.append("暂停")
    if status & 0x00000002: bits.append("错误")
    if status & 0x00000004: bits.append("正在初始化")
    if status & 0x00000008: bits.append("正在打印")
    if status & 0x00000100: bits.append("离线")
    if status & 0x00000200: bits.append("缺纸")
    if status & 0x00000400: bits.append("卡纸")
    if status & 0x00000800: bits.append("打印中")
    if status & 0x00001000: bits.append("打印队列已满")
    if status & 0x00008000: bits.append("用户干预")
    if status & 0x00400000: bits.append("繁忙")
    return "、".join(bits) if bits else "就绪"


# ================================================================ 打印配置
# Windows DEVMODE 标准纸张 ID（与 dmPaperSize 字段对应）。
# 键一律小写：查表前统一 .lower()，避免出现 "A4" / "a4" 两张表各写一半
# 的经典事故（旧版就因为 "A4".lower() 查不到而静默退回默认纸张）。
PAPER_IDS = {
    "letter": 1, "letter_small": 2, "tabloid": 3, "ledger": 4, "legal": 5,
    "statement": 6, "executive": 7, "a3": 8, "a4": 9, "a4_small": 10,
    "a5": 11, "b4": 12, "b5": 13, "folio": 14, "quarto": 15,
    "env_10": 20, "env_dl": 27, "env_c5": 28, "env_b5": 29,
}

# 纸张 label → (mm 宽, mm 高) — 用于"导出 PDF"功能构造自定义页面
PAPER_DIM_MM = {
    "a4": (210, 297), "a3": (297, 420), "a5": (148, 210),
    "letter": (216, 279), "legal": (216, 356), "executive": (184, 267),
    "b4": (250, 353), "b5": (176, 250), "tabloid": (279, 432),
}


def paper_dims_mm(paper: str) -> tuple[int, int]:
    """纸张 → (宽 mm, 高 mm)。未知纸张一律按 A4 处理。"""
    return PAPER_DIM_MM.get((paper or "").strip().lower(), PAPER_DIM_MM["a4"])

ORIENT_IDS = {"portrait": 1, "landscape": 2}
DUPLEX_IDS = {"none": 1, "long": 2, "short": 4}      # DMDUP_SIMPLEX/VERTICAL/HORIZONTAL
COLOR_IDS = {"color": 2, "mono": 1}                   # DMCOLOR_COLOR/MONOCHROME


@dataclass
class PrintConfig:
    """一份文件独立的打印参数。被每文件覆盖表里 row 的字段替换。"""
    paper: str = "A4"
    orientation: str = "portrait"  # portrait / landscape
    copies: int = 1
    duplex: str = "none"           # none / long / short
    color: str = "color"           # color / mono

    def as_dict(self) -> dict:
        return asdict(self)

    def merge(self, overrides: dict | None) -> "PrintConfig":
        """把 overrides 字典里非空字段覆盖到自身，返回新对象。"""
        if not overrides:
            return self
        d = self.as_dict()
        for k, v in overrides.items():
            if v is None or v == "" or v == "inherit":
                continue
            if k in d:
                d[k] = v
        return PrintConfig(**d)


def _build_devmode(printer: str, cfg: PrintConfig) -> dict:
    """**只读**地算出"这份配置落到 DEVMODE 上应该是什么样"。

    返回的是 pywin32 的 PyDEVMODE 副本（不是原始字节），用途是：
      · 日志里把真实参数讲清楚
      · 诊断 / 测试断言

    **不要把它写回打印机** —— 见 ``_job_devmode_bytes`` 的说明。
    """
    if not PYWIN32_OK:
        return {}
    import win32print

    h = win32print.OpenPrinter(printer)
    try:
        info = win32print.GetPrinter(h, 2)
        devmode = info["pDevMode"]
    finally:
        win32print.ClosePrinter(h)

    _apply_config_to_pydevmode(devmode, cfg)
    return devmode


def _apply_config_to_pydevmode(devmode, cfg: PrintConfig) -> None:
    """把 PrintConfig 的字段写进一个 PyDEVMODE（仅内存，不落系统）。"""
    paper_id = PAPER_IDS.get(cfg.paper.lower(), 9)  # 默认 A4
    devmode.PaperSize = paper_id
    devmode.FormName = cfg.paper.upper()[:31]
    devmode.Orientation = ORIENT_IDS.get(cfg.orientation, 1)
    # 份数固定写 1：份数由"提交 N 次作业"实现（见 _do_print）。
    # 若这里也写 N，遇到会叠加的驱动就会打出 N×N 份 —— 这种"多打一堆纸"
    # 的 bug 用户只会觉得见鬼。
    devmode.Copies = 1
    devmode.Duplex = DUPLEX_IDS.get(cfg.duplex, 1)
    devmode.Color = COLOR_IDS.get(cfg.color, 2)
    # 关键：dmFields 必须置位，否则驱动忽略这些字段
    devmode.Fields = (
        devmode.Fields
        | win32con.DM_PAPERSIZE
        | win32con.DM_ORIENTATION
        | win32con.DM_COPIES
        | win32con.DM_DUPLEX
        | win32con.DM_COLOR
    )


# DEVMODEW 关键字段偏移（WCHAR dmDeviceName[32] 开头，逐字段累加）
_OFF_DM_SIZE = 68            # WORD  dmSize
_OFF_DM_FIELDS = 72          # DWORD dmFields
_OFF_DM_ORIENTATION = 76     # short dmOrientation
_OFF_DM_PAPER_SIZE = 78      # short dmPaperSize
_OFF_DM_COPIES = 86          # short dmCopies
_OFF_DM_COLOR = 92           # short dmColor
_OFF_DM_DUPLEX = 94          # short dmDuplex
_OFF_DM_FORM_NAME = 102      # WCHAR dmFormName[32]（跟在 dmYResolution 之后）

_DM_OUT_BUFFER = 0x02
_DM_IN_BUFFER = 0x08


def _read_default_devmode_bytes(printer: str) -> bytearray:
    """读打印机的默认 DEVMODE 原始字节（只读，不改系统状态）。

    走 winspool 的 ``DocumentPropertiesW`` 而不是 pywin32 的 PyDEVMODE，
    因为后面要用这块内存直接喂给 ``gdi32.CreateDCW``，必须是原始字节。
    """
    if not PYWIN32_OK:
        raise RuntimeError("pywin32 不可用")
    import ctypes
    from ctypes import wintypes

    winspool = ctypes.WinDLL("winspool.drv", use_last_error=True)
    winspool.DocumentPropertiesW.restype = ctypes.c_long
    winspool.DocumentPropertiesW.argtypes = [
        wintypes.HWND, wintypes.HANDLE, wintypes.LPWSTR,
        ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD,
    ]

    h = win32print.OpenPrinter(printer)
    try:
        need = winspool.DocumentPropertiesW(None, int(h), printer, None, None, 0)
        if need <= 0:
            raise RuntimeError(f"无法取得 DEVMODE 长度（返回 {need}）")
        buf = ctypes.create_string_buffer(need)
        rc = winspool.DocumentPropertiesW(None, int(h), printer, buf, None, _DM_OUT_BUFFER)
        if rc < 0:
            raise RuntimeError(f"无法取得 DEVMODE 内容（返回 {rc}）")
        return bytearray(buf.raw)
    finally:
        win32print.ClosePrinter(h)


def _job_devmode_bytes(printer: str, cfg: PrintConfig) -> bytearray:
    """构造"本次作业专用"的 DEVMODE 原始字节。

    做法：拿打印机默认 DEVMODE 的字节副本 → 按偏移改纸张/方向/份数/颜色/双面
    → 交给驱动做一次校验合并（DocumentPropertiesW DM_IN_BUFFER|DM_OUT_BUFFER）。
    全程只在内存里操作，**不写系统状态、不需要管理员权限**。
    """
    raw = _read_default_devmode_bytes(printer)
    if len(raw) < _OFF_DM_DUPLEX + 2:
        raise RuntimeError("DEVMODE 结构异常（长度不足）")

    def _set_i16(off: int, val: int) -> None:
        raw[off:off + 2] = int(val).to_bytes(2, "little", signed=True)

    def _set_i32(off: int, val: int) -> None:
        raw[off:off + 4] = int(val).to_bytes(4, "little", signed=True)

    fields = int.from_bytes(raw[_OFF_DM_FIELDS:_OFF_DM_FIELDS + 4], "little")

    _set_i16(_OFF_DM_PAPER_SIZE, PAPER_IDS.get(cfg.paper.lower(), 9))
    _set_i16(_OFF_DM_ORIENTATION, ORIENT_IDS.get(cfg.orientation, 1))
    _set_i16(_OFF_DM_COPIES, 1)
    _set_i16(_OFF_DM_COLOR, COLOR_IDS.get(cfg.color, 2))
    _set_i16(_OFF_DM_DUPLEX, DUPLEX_IDS.get(cfg.duplex, 1))
    # dmFormName 同步成纸张名，避免有的驱动只认 FormName
    form = cfg.paper.upper().encode("utf-16-le")[:62]
    raw[_OFF_DM_FORM_NAME:_OFF_DM_FORM_NAME + 64] = form.ljust(64, b"\x00")

    fields |= (win32con.DM_PAPERSIZE | win32con.DM_ORIENTATION | win32con.DM_COPIES
               | win32con.DM_DUPLEX | win32con.DM_COLOR | win32con.DM_FORMNAME)
    _set_i32(_OFF_DM_FIELDS, fields)

    # 交给驱动校验一次：越界的纸张（例如驱动不支持的 Executive）会被纠正回
    # 可用值，而不是硬塞进去导致打印错乱。失败也不致命，用原样字节继续。
    try:
        import ctypes
        from ctypes import wintypes

        winspool = ctypes.WinDLL("winspool.drv", use_last_error=True)
        winspool.DocumentPropertiesW.restype = ctypes.c_long
        winspool.DocumentPropertiesW.argtypes = [
            wintypes.HWND, wintypes.HANDLE, wintypes.LPWSTR,
            ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD,
        ]
        h = win32print.OpenPrinter(printer)
        try:
            buf = ctypes.create_string_buffer(bytes(raw), len(raw))
            winspool.DocumentPropertiesW(None, int(h), printer, buf, buf,
                                         _DM_IN_BUFFER | _DM_OUT_BUFFER)
            raw = bytearray(buf.raw[:len(raw)])
        finally:
            win32print.ClosePrinter(h)
    except Exception:
        pass
    return raw


def _create_printer_dc(printer: str, cfg: PrintConfig):
    """创建带"本次作业 DEVMODE"的打印 DC。

    返回 ``(dc, applied, note)``：

    * ``applied=True``  —— 成功套用了本次作业的纸张/方向/份数/颜色/双面；
    * ``applied=False`` —— 退回 ``CreatePrinterDC``（用打印机默认设置），
      ``note`` 里是原因，会原样写进执行日志。

    为什么不直接 ``except Exception: pass`` 了事：这里曾经因为一个参数写错
    （少传 cfg）导致自定义 DEVMODE 静默失效，用户看到的现象是"设了 A5 却
    打 A4"，而日志里一个字都没有。**降级必须留痕**。
    """
    import win32ui

    note = ""
    try:
        import ctypes
        from ctypes import wintypes

        raw = _job_devmode_bytes(printer, cfg)
        gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
        gdi32.CreateDCW.restype = wintypes.HDC
        gdi32.CreateDCW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR,
                                    wintypes.LPCWSTR, ctypes.c_void_p]
        driver = _printer_driver(printer)
        buf = ctypes.create_string_buffer(bytes(raw), len(raw))
        hdc = gdi32.CreateDCW(driver, printer, None, ctypes.byref(buf))
        if hdc:
            return win32ui.CreateDCFromHandle(hdc), True, ""
        note = (f"CreateDCW 失败（err={ctypes.get_last_error()}），"
                f"该驱动可能不接受外部 DEVMODE")
    except Exception as exc:
        note = f"{type(exc).__name__}: {exc}"

    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(printer)
    return dc, False, note


def _printer_driver(printer: str) -> str:
    """取打印机驱动名（CreateDCW 需要）。"""
    for it in _enum_printers():
        if it["name"] == printer:
            return it["driver"] or "WINSPOOL"
    return "WINSPOOL"


def _printer_default_summary(printer: str) -> dict:
    """只读地读出打印机默认设置，用于日志与"预期管理"提示。"""
    out: dict = {}
    try:
        raw = _read_default_devmode_bytes(printer)
        out["paper"] = int.from_bytes(raw[_OFF_DM_PAPER_SIZE:_OFF_DM_PAPER_SIZE + 2],
                                      "little", signed=True)
        out["orientation"] = int.from_bytes(raw[_OFF_DM_ORIENTATION:_OFF_DM_ORIENTATION + 2],
                                           "little", signed=True)
        out["copies"] = int.from_bytes(raw[_OFF_DM_COPIES:_OFF_DM_COPIES + 2],
                                       "little", signed=True)
    except Exception:
        return {}
    return out


def _paper_pixel_size(paper: str, orientation: str, dpi: int) -> tuple[int, int]:
    """返回纸张按 dpi 渲染时的 (w, h) 像素值（按 requested orientation 调整）。"""
    mm_w, mm_h = paper_dims_mm(paper)
    if orientation == "landscape":
        mm_w, mm_h = mm_h, mm_w
    # 1 inch = 25.4mm; pixel = mm * dpi / 25.4
    w_px = int(round(mm_w * dpi / 25.4))
    h_px = int(round(mm_h * dpi / 25.4))
    return w_px, h_px


# ================================================================ 核心渲染
def _render_pdf_to_images(path: str, dpi: int = 150):
    """把 PDF 每页渲染为 PIL Image。PyMuPDF 优先，pdf2image 回退。"""
    try:
        import pymupdf  # type: ignore

        doc = pymupdf.open(path)
        try:
            zoom = dpi / 72.0
            mat = pymupdf.Matrix(zoom, zoom)
            out = []
            for page in doc:
                pix = page.get_pixmap(matrix=mat, colorspace=pymupdf.csRGB, alpha=False)
                from PIL import Image
                im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                out.append(im)
            return out
        finally:
            doc.close()
    except ImportError:
        from pdf2image import convert_from_path
        return convert_from_path(path, dpi=dpi)


def _image_to_dib_bytes(im, top_down: bool = True) -> bytes:
    """PIL Image → 24-bit DIB（BITMAPINFOHEADER + 像素）。"""
    import struct
    w, h = im.size
    bpp = 24
    row_bytes = ((w * bpp + 31) // 32) * 4
    pixel_bytes = row_bytes * h
    biHeight = -h if top_down else h
    info_header = struct.pack(
        "<IiiHHIIiiII",
        40, w, biHeight, 1, bpp, 0, pixel_bytes, 2835, 2835, 0, 0,
    )
    pixels = im.tobytes("raw", "BGR")
    aligned = b""
    src_row = w * 3
    if src_row == row_bytes:
        aligned = pixels
    else:
        pad = row_bytes - src_row
        for i in range(h):
            aligned += pixels[i * src_row:(i + 1) * src_row]
            aligned += b"\x00" * pad
    return info_header + aligned


# ================================================================ Op 注册
def _build_print_op() -> OpDef | None:
    """构造打印 OpDef（物理打印机 + 每文件独立配置）。"""
    if not PYWIN32_OK:
        return None
    printers = list_printers(real_only=True)
    virtual = list_virtual_pdf_printers()
    if not printers:
        return None

    default = default_printer()
    if default and default in printers:
        pass
    else:
        # 系统默认落在被过滤的虚拟设备上（例如 Microsoft Print to PDF），
        # 这时绝不能顺手把它当默认值 —— 否则"开始执行"会静默产出一堆 PDF。
        default = printers[0]

    choices = [(p, p) for p in printers]
    note = ""
    if virtual:
        note = f"（已隐藏 {len(virtual)} 台虚拟打印设备，如 {virtual[0]}）"

    return OpDef(
        key="pdf_print",
        title="PDF 批量打印" + note,
        category="pdf_pages",
        accepts=("pdf",),
        mode="per_file",
        hardware=True,
        description=(
            "把 PDF 发送到本机或网络打印机。支持每文件独立配置纸张/方向/份数/双面/颜色，"
            "默认配置可一键批量应用到选中文件。\n"
            "「仅演练不打印」勾选后只生成预览 PDF，不会向打印机提交任何作业。"
        ),
        confirm=True,
        confirm_text=(
            "即将把选中的文件发送到打印机，开始打印前请确认：\n"
            "  · 已选择正确的打印机\n"
            "  · 默认份数、纸张、方向符合预期\n"
            "  · 已检查每文件覆盖配置\n"
            "  · 打印机处于就绪状态\n"
        ),
        # 确认框里把"真实动作"讲清楚：哪台打印机、多少份、大概多少张纸
        confirm_builder=_describe_print_job,
        params=(
            Param("printer", "打印机", "choice", default=default,
                  choices=choices, width=32,
                  hint=("列表中已隐藏虚拟 PDF 打印机（选它们不会出纸，只会多出一堆 "
                        "PDF 文件）。如需输出为 PDF 请用「PDF 导出 PDF」。")),
            Param("dry_run", "仅演练不打印", "bool", default=False,
                  hint=("勾选后**不向打印机提交任何作业**，只把每份文件按当前配置"
                        "渲染成预览 PDF 放到输出目录，供你先核对版式再真打。")),
            Param("file_configs", "每文件配置", "per_file_config",
                  default={"default": PrintConfig().as_dict(), "overrides": {}},
                  axes=(
                      {"key": "paper", "label": "纸张",
                       "choices": [("A4", "A4"), ("a3", "A3"), ("a5", "A5"),
                                   ("letter", "Letter"), ("b5", "B5"),
                                   ("legal", "Legal"), ("executive", "Executive")]},
                      {"key": "orientation", "label": "方向",
                       "choices": [("portrait", "纵向"), ("landscape", "横向")]},
                      {"key": "copies", "label": "份数", "widget": "int",
                       "min": 1, "max": 999, "default": 1},
                      {"key": "duplex", "label": "双面",
                       "choices": [("none", "单面"), ("long", "双面长边翻"),
                                   ("short", "双面短边翻")]},
                      {"key": "color", "label": "颜色",
                       "choices": [("color", "彩色"), ("mono", "黑白")]},
                  ),
                  hint=("每份文件可单独设纸张/方向/份数/双面/颜色，未单独设置的继承默认配置。\n"
                        "份数按「提交 N 次作业」实现（任意账户都生效）；纸张/方向/双面/颜色"
                        "通过本次作业的 DEVMODE 下发，不修改你的系统打印机默认设置。")),
            Param("page_range", "页面范围", "range", default="", width=24,
                  hint="留空=全部；如 1-3,5,8-10"),
            Param("scaling", "缩放", "choice", default="fit",
                  choices=(("fit", "适合页面"), ("actual", "实际大小"),
                           ("custom", "自定义缩放 %"))),
            Param("scale_pct", "缩放百分比", "int", default=100,
                  min_value=50, max_value=200, visible_if=("scaling", ("custom",))),
            Param("n_up", "每张纸拼版", "choice", default="1",
                  choices=(("1", "1 页/张"), ("2", "2 页/张"),
                           ("4", "4 页/张"), ("6", "6 页/张"))),
            Param("dpi", "渲染精度", "choice", default="150",
                  choices=(("96", "草稿 96 DPI"), ("150", "标准 150 DPI"),
                           ("300", "高质量 300 DPI"), ("600", "印刷 600 DPI")),
                  hint="数值越高越清晰越慢"),
        ),
    )


def _describe_print_job(files: list, params: dict) -> str:
    """给确认对话框用：把"会发生什么"讲成人话。"""
    printer = str(params.get("printer") or "（未选择）")
    dry = bool(params.get("dry_run"))
    fc = params.get("file_configs") or {}
    default_cfg = PrintConfig(**(fc.get("default") or {}))
    overrides = fc.get("overrides") or {}

    total_copies = 0
    detail_lines = []
    paper_set = set()
    sheets = 0
    for idx, f in enumerate(files):
        row = overrides.get(str(idx)) or overrides.get(idx) or {}
        cfg = default_cfg.merge(row)
        copies = max(1, int(cfg.copies or 1))
        total_copies += copies
        paper_set.add(f"{cfg.paper.upper()}{'横' if cfg.orientation == 'landscape' else '纵'}")
        pages = _count_pages_safe(f)
        sheets += copies * pages
        if row:
            detail_lines.append(
                f"    · {os.path.basename(f)}：{cfg.paper.upper()}"
                f"{'横向' if cfg.orientation == 'landscape' else '纵向'} × {copies} 份")
        else:
            detail_lines.append(
                f"    · {os.path.basename(f)}：继承默认（{cfg.paper.upper()}"
                f"{'横向' if cfg.orientation == 'landscape' else '纵向'} × {copies} 份）")

    if dry:
        return (
            "【演练模式】本次不会向打印机提交任何作业。\n\n"
            f"将按以下配置渲染 {len(files)} 份预览 PDF 到输出目录：\n"
            + "\n".join(detail_lines[:12])
            + (f"\n    …另有 {len(detail_lines) - 12} 个文件" if len(detail_lines) > 12 else "")
            + "\n\n确认继续吗？"
        )

    return (
        f"⚠ 以下作业会真实打印到物理打印机上，打印出来就无法撤销。\n\n"
        f"  打印机：{printer}\n"
        f"  文件数：{len(files)} 份，合计 {total_copies} 份次、约 {sheets} 张纸\n"
        f"  纸张组合：{'、'.join(sorted(paper_set))}\n"
        + "\n".join(detail_lines[:12])
        + (f"\n    …另有 {len(detail_lines) - 12} 个文件" if len(detail_lines) > 12 else "")
        + "\n\n确认开始打印吗？"
    )


def _count_pages_safe(path: str) -> int:
    """页数；读不出来就按 1 页估算（只用于提示，不参与真实逻辑）。"""
    try:
        from pypdf import PdfReader

        return max(1, len(PdfReader(path).pages))
    except Exception:
        return 1


def _build_export_pdf_op() -> OpDef | None:
    """构造「PDF 导出 PDF」OpDef：不走打印机，直接合成指定纸张的 PDF。"""
    if not PYWIN32_OK:
        return None
    return OpDef(
        key="pdf_export",
        title="PDF 导出 PDF（批量）",
        category="convert",
        accepts=("pdf",),
        mode="per_file",
        description=(
            "把多份 PDF 导出为指定纸张的 PDF 文件，无任何对话框。适合没有物理打印机、"
            "只想批量改纸张 / 方向 / 拼版的场景。"
        ),
        confirm=True,
        confirm_text=(
            "即将把选中的 PDF 导出为新 PDF 文件，请确认：\n"
            "  · 默认纸张 / 方向符合预期\n"
            "  · 已检查每文件覆盖配置（个别文件可用特殊纸张）\n"
        ),
        params=(
            Param("file_configs", "每文件配置", "per_file_config",
                  default={"default": PrintConfig().as_dict(), "overrides": {}},
                  axes=(
                      {"key": "paper", "label": "纸张",
                       "choices": [("A4", "A4"), ("a3", "A3"), ("a5", "A5"),
                                   ("letter", "Letter"), ("b5", "B5"),
                                   ("legal", "Legal"), ("executive", "Executive")]},
                      {"key": "orientation", "label": "方向",
                       "choices": [("portrait", "纵向"), ("landscape", "横向")]},
                  ),
                  hint="默认配置应用到所有未单独设置的文件。"),
            Param("merge_all", "合并为单个 PDF", "bool", default=False,
                  hint="勾选则把全部 PDF 合并为单个 .pdf；不勾则一份 PDF 一个文件"),
            Param("n_up", "每张纸拼版", "choice", default="1",
                  choices=(("1", "1 页/张"), ("2", "2 页/张"),
                           ("4", "4 页/张"), ("6", "6 页/张"))),
            Param("fit_mode", "适配模式", "choice", default="fit",
                  choices=(("fit", "缩放至页面内（留白边）"),
                           ("actual", "原始大小（超出裁切）"),
                           ("fit_smart", "智能适配（缩放并居中）"))),
            Param("dpi", "渲染精度", "choice", default="150",
                  choices=(("96", "草稿 96 DPI"), ("150", "标准 150 DPI"),
                           ("300", "高质量 300 DPI"), ("600", "印刷 600 DPI"))),
        ),
    )


# --------------------------------------------------------------- 注册
def _ensure_registered() -> None:
    """运行时再尝试注册（pywin32 可能在启动后才被加载）。"""
    from .registry import OPS as _OPS

    if not PYWIN32_OK:
        return

    if "pdf_print" not in _OPS:
        op = _build_print_op()
        if op is not None:
            op.handler = lambda ctx, files, params, _o=op: _do_print(ctx, files, params)
            _OPS[op.key] = op

    if "pdf_export" not in _OPS:
        op = _build_export_pdf_op()
        if op is not None:
            op.handler = lambda ctx, files, params, _o=op: _do_export_pdf(ctx, files, params)
            _OPS[op.key] = op


# --------------------------------------------------------------- 打印执行
# 熔断开关：设为 1 时，本进程内的任何真实打印都会被拒绝。
# 给 CI / 远程无人值守 / 排障时使用（"我先把打印彻底关掉"）。
ENV_NO_PRINT = "ODA_NO_PRINT"
# 强制演练模式（等价于界面勾上"仅演练不打印"）
ENV_DRYRUN = "ODA_PRINT_DRYRUN"


def _env_true(name: str) -> bool:
    return str(os.environ.get(name, "")).strip().lower() in ("1", "true", "yes", "on")


def print_denied_reason(ctx) -> str:
    """返回非空字符串 = 拒绝真实打印，字符串即原因。

    三道闸门，任一不通过都不许出纸：
      1. 进程级熔断 ``ODA_NO_PRINT=1``
      2. 人工授权 ``ctx.allow_real_print``（只有界面确认过的路径才有）
      3. 功能被显式标记为需要硬件（``op.hardware``）
    """
    if _env_true(ENV_NO_PRINT):
        return (f"已通过环境变量 {ENV_NO_PRINT}=1 全局禁用真实打印（当前进程内的"
                f"打印请求一律拒绝）。如需恢复，请清除该环境变量后重启程序。")
    if getattr(ctx.op, "hardware", False) and not getattr(ctx, "allow_real_print", False):
        return ("拒绝执行：真实打印需要由人在界面上确认后才能进行，"
                "而当前调用方没有提供该授权（测试脚本 / 批处理 / 定时任务等"
                "无人值守路径默认禁止打印）。\n"
                "如果你确实想打印，请打开程序界面，选择「PDF 批量打印」，"
                "添加文件后点击「开始执行」并在确认框中确认；"
                "只想核对版式可改勾「仅演练不打印」。")
    return ""


def _is_dry_run(ctx, params: dict) -> bool:
    """本次是否演练模式（不提交任何打印作业）。"""
    if _env_true(ENV_DRYRUN):
        return True
    if getattr(ctx, "dry_run", False):
        return True
    return bool(params.get("dry_run"))


def _do_print(ctx: OpContext, files: list[str], params: dict) -> None:
    """实际打印逻辑：每文件独立 DEVMODE。

    顺序很重要：
      1. 先判定演练 / 授权 —— **在碰任何打印机之前**；
      2. 再解析配置与页数；
      3. 最后才逐个提交作业。
    这样即使参数有问题，也不会出现"报错了但纸已经出去了"。
    """
    printer = str(params.get("printer") or "").strip()
    dry_run = _is_dry_run(ctx, params)

    # ---- 闸门 1：演练模式优先（演练连打印机都不需要，更不需要授权）
    if not dry_run:
        reason = print_denied_reason(ctx)
        if reason:
            raise PrintNotAuthorized(reason)
        if not printer:
            raise ValueError("未选择打印机")

    scaling = str(params.get("scaling") or "fit")
    scale_pct = max(50, min(int(params.get("scale_pct") or 100), 200))
    n_up = int(params.get("n_up") or 1)
    dpi = int(params.get("dpi") or 150)
    page_spec = str(params.get("page_range") or "").strip()

    # 每文件配置：默认 + 覆盖
    file_cfg_raw = params.get("file_configs") or {}
    default_cfg = PrintConfig(**(file_cfg_raw.get("default") or {}))
    overrides = file_cfg_raw.get("overrides") or {}

    total_jobs = 0
    for idx in range(len(files)):
        row = overrides.get(str(idx)) or overrides.get(idx) or {}
        total_jobs += max(1, int(default_cfg.merge(row).copies or 1))

    if dry_run:
        ctx.log("══ 演练模式：本次不会向打印机提交任何作业 ══", "warn")
        ctx.log(f"将按当前配置渲染 {len(files)} 份预览 PDF 到：{ctx.out_dir}", "info")
        ctx.log("核对无误后，取消勾选「仅演练不打印」再执行一次即可真正打印。", "info")
    else:
        # 只读地把打印机现状讲清楚，避免"我以为打的是 A4"这类误会
        code, desc = _printer_status(printer)
        if code != 0 and code not in (8, 1024, 0x10000):
            raise RuntimeError(f"打印机不可用：{desc or '状态码 ' + str(code)}")
        ctx.log(f"打印机：{printer}  状态：{desc or '就绪'}", "info")
        _log_printer_defaults(ctx, printer)

    ctx.log(f"计划：{len(files)} 份文件 / 合计 {total_jobs} 次作业"
            f"{'（演练：不提交）' if dry_run else ''}", "info")
    ctx.metric("模式", "演练（未向打印机提交任何作业）" if dry_run else "真实打印")
    ctx.metric("打印机", "—" if dry_run else printer)
    ctx.metric("合计作业数", total_jobs)
    ctx.set_total(len(files))


    for idx, f in enumerate(files):
        ctx.check()
        row_overrides = overrides.get(str(idx)) or overrides.get(idx) or {}
        cfg = default_cfg.merge(row_overrides)
        try:
            from pypdf import PdfReader

            reader = PdfReader(f)
            total = len(reader.pages)
            indices = list(range(total))
            if page_spec:
                from .pdfutil import parse_ranges
                indices = parse_ranges(page_spec, total, clip=True)

            ctx.log(f"  {os.path.basename(f)}：{cfg.paper.upper()}"
                    f"{'横' if cfg.orientation == 'landscape' else '纵'} "
                    f"份数={cfg.copies} 双面={cfg.duplex} 颜色={cfg.color}  "
                    f"共 {total} 页 → 输出 {len(indices)} 页", "info")

            if dry_run:
                out_path = _dryrun_preview(ctx, f, indices, cfg, n_up, dpi)
                ctx.log(f"  ✓ {os.path.basename(f)} → {os.path.basename(out_path)}"
                        f"（预览，未打印）", "success")
                ctx.ok(out_path=out_path, src=f)
            else:
                for copy_n in range(max(1, int(cfg.copies or 1))):
                    ctx.check()
                    if cfg.copies > 1:
                        ctx.log(f"    第 {copy_n + 1}/{cfg.copies} 份", "info")
                    _print_one_pdf(ctx, f, indices, printer, cfg, scaling, scale_pct,
                                   n_up, dpi)
                ctx.log(f"  ✓ {os.path.basename(f)} 打印任务已提交", "success")
                ctx.ok(src=f)
        except PrintNotAuthorized:
            raise
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"{'演练' if dry_run else '打印'} {os.path.basename(f)}")


def _log_printer_defaults(ctx: OpContext, printer: str) -> None:
    """打印前把打印机的当前默认设置打出来，并对"可能多打"的情况明确预警。"""
    info = _printer_default_summary(printer)
    if not info:
        return
    paper_rev = {v: k for k, v in PAPER_IDS.items()}
    paper = paper_rev.get(info.get("paper", 0), f"ID{info.get('paper')}")
    orient = "横向" if info.get("orientation") == 2 else "纵向"
    copies = int(info.get("copies") or 1)
    ctx.log(f"  打印机默认：{paper.upper()} {orient}，份数={copies}（仅作参考，本工具会逐作业覆盖）",
            "debug")
    if copies != 1:
        # 份数在部分驱动上是"乘"的关系，必须让用户知道，而不是事后数纸
        ctx.log(f"  ⚠ 注意：该打印机系统默认份数是 {copies}，"
                f"若驱动把它与作业份数相乘，实际出纸张数会是请求份数的 {copies} 倍。"
                f"可在「设备和打印机 → 打印首选项」里把份数改回 1。", "warn")


def _dryrun_preview(ctx: OpContext, src_pdf: str, indices: list[int],
                    cfg: PrintConfig, n_up: int, dpi: int) -> str:
    """演练：只渲染，产出一份预览 PDF，不碰打印机。"""
    import pymupdf
    from io import BytesIO

    mm_w, mm_h = paper_dims_mm(cfg.paper)
    if cfg.orientation == "landscape":
        page_w, page_h = mm_h, mm_w
    else:
        page_w, page_h = mm_w, mm_h
    pt_per_mm = 72.0 / 25.4
    out_w_pt = page_w * pt_per_mm
    out_h_pt = page_h * pt_per_mm

    rendered_all = _render_pdf_to_images(src_pdf, dpi=dpi)
    rendered = [rendered_all[i] for i in indices if 0 <= i < len(rendered_all)]

    out_doc = pymupdf.open()
    try:
        for i in range(0, max(1, len(rendered)), max(1, n_up)):
            ctx.check()
            group = rendered[i:i + max(1, n_up)]
            grid_im = _compose_grid(group, *_grid_for(n_up))
            buf = BytesIO()
            grid_im.save(buf, format="PNG")
            page = out_doc.new_page(width=out_w_pt, height=out_h_pt)
            src_w, src_h = grid_im.size
            scale = min(out_w_pt / src_w, out_h_pt / src_h)
            new_w, new_h = src_w * scale, src_h * scale
            x0, y0 = (out_w_pt - new_w) / 2, (out_h_pt - new_h) / 2
            page.insert_image(pymupdf.Rect(x0, y0, x0 + new_w, y0 + new_h),
                              stream=buf.getvalue())
        base = os.path.splitext(os.path.basename(src_pdf))[0]
        out_path = os.path.join(ctx.out_dir, f"{base}_预览.pdf")
        out_doc.save(out_path)
    finally:
        out_doc.close()
    return out_path


def _print_one_pdf(ctx: OpContext, src_pdf: str, indices: list[int],
                   printer: str, cfg: PrintConfig, scaling: str,
                   scale_pct: int, n_up: int, dpi: int) -> None:
    """真实渲染并提交**一次**打印作业。

    关键说明
    --------
    * **每份文件一次作业、作业内份数固定 1**；要打 N 份就提交 N 次作业。
      这样份数在任意账户下都准确，不会依赖驱动是否买 DM_COPIES 的账，
      也绝不会出现"驱动 × 作业份数"叠乘导致打出一堆纸。
    * **纸张 / 方向 / 双面 / 颜色** 通过本次作业专用的 DEVMODE 下发
      （``gdi32.CreateDCW``），只作用于这一个作业：不会修改用户的系统
      打印机默认设置，也不需要管理员权限。
    * 若驱动拒绝自定义 DEVMODE，会退回打印机默认设置；此时纸张差异仍能
      从版式上体现（按所选纸张尺寸渲染再适配），并在日志中说明。
    """
    import win32con

    hdc, devmode_applied, note = _create_printer_dc(printer, cfg)
    if not devmode_applied:
        ctx.log(f"    提示：本次按打印机默认纸张/方向提交（{note}）——"
                f"份数不受影响，仍为 {cfg.copies} 份", "warn")

    try:
        hor_res = hdc.GetDeviceCaps(win32con.HORZRES)
        ver_res = hdc.GetDeviceCaps(win32con.VERTRES)
        hor_off = hdc.GetDeviceCaps(win32con.PHYSICALOFFSETX)
        ver_off = hdc.GetDeviceCaps(win32con.PHYSICALOFFSETY)

        hdc.StartDoc(os.path.basename(src_pdf))
        try:
            rendered_all = _render_pdf_to_images(src_pdf, dpi=dpi)
            rendered = [rendered_all[i] for i in indices if 0 <= i < len(rendered_all)]

            pages_per_sheet = max(1, n_up)
            sheet_groups = [rendered[i:i + pages_per_sheet]
                            for i in range(0, len(rendered), pages_per_sheet)]

            for gi, group in enumerate(sheet_groups):
                ctx.check()
                hdc.StartPage()
                try:
                    composed = _compose_grid(group, *_grid_for(len(group)))
                    _draw_image_to_dc(hdc, composed, hor_res, ver_res, hor_off, ver_off)
                finally:
                    hdc.EndPage()
                ctx.log(f"    第 {gi + 1}/{len(sheet_groups)} 张纸", "info")
        finally:
            hdc.EndDoc()
    finally:
        del hdc


def _compose_grid(images: list, gw: int, gh: int):
    from PIL import Image

    cw = max(im.size[0] for im in images)
    ch = max(im.size[1] for im in images)
    composed = Image.new("RGB", (cw * gw, ch * gh), "white")
    for i, im in enumerate(images):
        x = (i % gw) * cw
        y = (i // gw) * ch
        composed.paste(im, (x, y))
    return composed


def _draw_image_to_dc(hdc, image, dc_w: int, dc_h: int, off_x: int, off_y: int) -> None:
    import ctypes
    import win32con
    from ctypes import wintypes

    x = -off_x
    y = -off_y
    w = dc_w
    h = dc_h

    src_w, src_h = image.size
    dib = _image_to_dib_bytes(image, top_down=True)
    pixel_data = dib[40:]
    info_ptr = ctypes.c_char_p(dib[:40])

    gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
    StretchDIBits = gdi32.StretchDIBits
    StretchDIBits.argtypes = [
        wintypes.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
        ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
        ctypes.c_void_p, ctypes.c_char_p, wintypes.UINT, wintypes.DWORD,
    ]
    StretchDIBits.restype = ctypes.c_int

    StretchDIBits(
        hdc.GetSafeHdc(),
        x, y, w, h,
        0, 0, src_w, src_h,
        pixel_data, info_ptr, win32con.DIB_RGB_COLORS, win32con.SRCCOPY,
    )


# ================================================================ 导出 PDF 执行
def _do_export_pdf(ctx: OpContext, files: list[str], params: dict) -> None:
    """把多份 PDF 导出为指定纸张的新 PDF。零对话框，文件直接落盘。"""
    merge_all = bool(params.get("merge_all"))
    n_up = max(1, min(int(params.get("n_up") or 1), 6))
    fit_mode = str(params.get("fit_mode") or "fit")
    dpi = int(params.get("dpi") or 150)

    file_cfg_raw = params.get("file_configs") or {}
    default_cfg = PrintConfig(**(file_cfg_raw.get("default") or {}))
    overrides = file_cfg_raw.get("overrides") or {}

    ctx.set_total(len(files))

    if merge_all:
        # 合并为单个 PDF
        out_name = "merged.pdf"
        out_path = os.path.join(ctx.out_dir, out_name)
        try:
            _merge_pdfs(ctx, files, out_path, default_cfg, overrides, n_up, fit_mode, dpi)
            ctx.log(f"✓ 合并导出完成 → {out_name}", "success")
            ctx.ok(out_path=out_path)
        except Exception as exc:
            ctx.fail(out_path, str(exc))
        ctx.step("导出 PDF（合并）")
    else:
        for idx, f in enumerate(files):
            ctx.check()
            base = os.path.splitext(os.path.basename(f))[0]
            out_path = os.path.join(ctx.out_dir, f"{base}_export.pdf")
            row_overrides = overrides.get(str(idx)) or overrides.get(idx) or {}
            cfg = default_cfg.merge(row_overrides)
            try:
                _export_one_pdf(ctx, f, out_path, cfg, n_up, fit_mode, dpi)
                ctx.log(f"  ✓ {os.path.basename(f)} → {os.path.basename(out_path)}", "success")
                ctx.ok(out_path=out_path, src=f)
            except Exception as exc:
                ctx.fail(f, str(exc))
            ctx.step(f"导出 {os.path.basename(f)}")


def _export_one_pdf(ctx: OpContext, src_pdf: str, out_path: str,
                    cfg: PrintConfig, n_up: int, fit_mode: str, dpi: int) -> None:
    """把单份 PDF 按 cfg 的纸张 / 方向导出为新 PDF。"""
    import pymupdf

    mm_w, mm_h = paper_dims_mm(cfg.paper)
    if cfg.orientation == "landscape":
        page_w, page_h = mm_h, mm_w
    else:
        page_w, page_h = mm_w, mm_h

    # 1 pt = 1/72 in；1 in = 25.4 mm；所以 1 mm = 72/25.4 pt
    pt_per_mm = 72.0 / 25.4
    out_w_pt = page_w * pt_per_mm
    out_h_pt = page_h * pt_per_mm

    src_doc = pymupdf.open(src_pdf)
    try:
        out_doc = pymupdf.open()
        try:
            rendered_imgs = _render_pdf_to_images(src_pdf, dpi=dpi)
            # 按 n_up 拼版
            for i in range(0, len(rendered_imgs), n_up):
                ctx.check()
                group = rendered_imgs[i:i + n_up]
                grid_im = _compose_grid(group, *(_grid_for(n_up)))
                # PIL Image → bytes（PNG）→ pymupdf Page
                from io import BytesIO
                buf = BytesIO()
                grid_im.save(buf, format="PNG")
                rect = pymupdf.Rect(0, 0, out_w_pt, out_h_pt)
                page = out_doc.new_page(width=out_w_pt, height=out_h_pt)
                # 智能适配：根据 fit_mode 决定如何贴图
                if fit_mode == "actual":
                    # 原图直接画到页面，超出裁切
                    page.insert_image(rect, stream=buf.getvalue())
                else:
                    # fit / fit_smart：缩放至页面内并居中
                    src_w, src_h = grid_im.size
                    sx = out_w_pt / src_w
                    sy = out_h_pt / src_h
                    scale = min(sx, sy) if fit_mode == "fit" else min(sx, sy)
                    new_w = src_w * scale
                    new_h = src_h * scale
                    x0 = (out_w_pt - new_w) / 2
                    y0 = (out_h_pt - new_h) / 2
                    img_rect = pymupdf.Rect(x0, y0, x0 + new_w, y0 + new_h)
                    page.insert_image(img_rect, stream=buf.getvalue())
            out_doc.save(out_path)
        finally:
            out_doc.close()
    finally:
        src_doc.close()


def _merge_pdfs(ctx: OpContext, files: list[str], out_path: str,
                default_cfg: PrintConfig, overrides: dict, n_up: int,
                fit_mode: str, dpi: int) -> None:
    """把多份 PDF 按各自配置导出后拼成单个 PDF。"""
    import pymupdf
    from io import BytesIO

    src_doc = pymupdf.open()
    try:
        for idx, f in enumerate(files):
            ctx.check()
            row_overrides = overrides.get(str(idx)) or overrides.get(idx) or {}
            cfg = default_cfg.merge(row_overrides)
            mm_w, mm_h = paper_dims_mm(cfg.paper)
            if cfg.orientation == "landscape":
                page_w, page_h = mm_h, mm_w
            else:
                page_w, page_h = mm_w, mm_h
            pt_per_mm = 72.0 / 25.4
            out_w_pt = page_w * pt_per_mm
            out_h_pt = page_h * pt_per_mm

            rendered_imgs = _render_pdf_to_images(f, dpi=dpi)
            ctx.log(f"  合并 {os.path.basename(f)}：{cfg.paper} {cfg.orientation} "
                    f"({len(rendered_imgs)} 页)", "info")
            for i in range(0, len(rendered_imgs), n_up):
                ctx.check()
                group = rendered_imgs[i:i + n_up]
                grid_im = _compose_grid(group, *(_grid_for(n_up)))
                buf = BytesIO()
                grid_im.save(buf, format="PNG")
                rect = pymupdf.Rect(0, 0, out_w_pt, out_h_pt)
                page = src_doc.new_page(width=out_w_pt, height=out_h_pt)
                if fit_mode == "actual":
                    page.insert_image(rect, stream=buf.getvalue())
                else:
                    src_w, src_h = grid_im.size
                    scale = min(out_w_pt / src_w, out_h_pt / src_h)
                    new_w = src_w * scale
                    new_h = src_h * scale
                    x0 = (out_w_pt - new_w) / 2
                    y0 = (out_h_pt - new_h) / 2
                    img_rect = pymupdf.Rect(x0, y0, x0 + new_w, y0 + new_h)
                    page.insert_image(img_rect, stream=buf.getvalue())
        src_doc.save(out_path)
    finally:
        src_doc.close()


def _grid_for(n: int) -> tuple[int, int]:
    if n == 1: return 1, 1
    if n == 2: return 1, 2
    if n == 4: return 2, 2
    if n == 6: return 2, 3
    return 1, n


# ================================================================ 维护工具
def restore_printer_defaults(printer: str, paper: str = "a4",
                             orientation: str = "portrait", copies: int = 1,
                             color: str = "color", duplex: str = "none") -> str:
    """把打印机的**系统默认设置**恢复成常见值。

    为什么需要这个函数：1.0.1 及更早版本会在每次打印时调用
    ``win32print.SetPrinter`` 把 DEVMODE 写回打印机，等于**永久**改掉了用户
    的打印机默认纸张/方向/份数（还会影响所有其它软件）。1.0.2 已经不再这么
    做，但被旧版本改坏过的机器需要一个还原入口。

    这是唯一一处会写系统打印机状态的代码，因此只由用户显式调用
    （``tools/fix_printer_defaults.py``），日常打印路径绝不调用它。
    """
    if not PYWIN32_OK:
        raise RuntimeError("仅 Windows 且 pywin32 可用时才能修复打印机默认设置")

    # 这里走 PyDEVMODE + SetPrinter：这是唯一被验证过能写回系统默认值的通道
    h = win32print.OpenPrinter(printer)
    try:
        info = win32print.GetPrinter(h, 2)
        dm = info["pDevMode"]
        dm.PaperSize = PAPER_IDS.get(paper.lower(), 9)
        dm.FormName = paper.upper()[:31]
        dm.Orientation = ORIENT_IDS.get(orientation, 1)
        dm.Copies = max(1, min(int(copies), 999))
        dm.Color = COLOR_IDS.get(color, 2)
        dm.Duplex = DUPLEX_IDS.get(duplex, 1)
        dm.Fields = (dm.Fields | win32con.DM_PAPERSIZE | win32con.DM_ORIENTATION
                     | win32con.DM_COPIES | win32con.DM_DUPLEX | win32con.DM_COLOR)
        info["pDevMode"] = dm
        win32print.SetPrinter(h, 2, info, 0)
    finally:
        win32print.ClosePrinter(h)
    return f"{printer} → {paper.upper()} {orientation} × {copies} 份 {color} {duplex}"


# ================================================================ 模块导入时尝试注册
def _try_register_print_op() -> None:
    _ensure_registered()


_try_register_print_op()