"""PDF 公共工具：读取器、页范围解析、加密处理、字体注册。

所有 PDF 相关模块共用这里的函数，保证行为一致（尤其是加密与损坏文件的处理）。
"""
from __future__ import annotations

import io
import re
from typing import Iterable

from pypdf import PdfReader, PdfWriter
from pypdf.errors import PdfReadError

from ..errors import BrokenFile, ParamError, PasswordRequired, WrongPassword

# ---------------------------------------------------------------- 页码范围
_RANGE_RE = re.compile(r"^\s*(\d+)\s*(?:-\s*(\d+)\s*)?$")


def parse_ranges(spec: str, total: int, *, clip: bool = True) -> list[int]:
    """把 "1-3,5,8-10" 解析为 0 基页码列表（去重，保持书写顺序）。

    :param clip: True 时把超出范围的页码裁掉；False 时抛 ParamError。
    """
    if not spec or not spec.strip():
        raise ParamError("页码范围不能为空，例如：1-3,5,8-10")
    spec = spec.strip().replace("，", ",").replace("－", "-").replace("—", "-").replace("~", "-")
    result: list[int] = []
    seen: set[int] = set()
    for chunk in spec.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        m = _RANGE_RE.match(chunk)
        if not m:
            raise ParamError(f"无法识别的页码片段：{chunk!r}")
        start = int(m.group(1))
        end = int(m.group(2)) if m.group(2) else start
        if start < 1 or end < 1:
            raise ParamError("页码从 1 开始，不能小于 1")
        lo, hi = (start, end) if start <= end else (end, start)
        for n in range(lo, hi + 1):
            if n > total:
                if clip:
                    continue
                raise ParamError(f"页码 {n} 超出文档总页数 {total}")
            idx = n - 1
            if idx not in seen:
                seen.add(idx)
                result.append(idx)
    if not result:
        raise ParamError(f"页码范围未匹配到任何页面（文档共 {total} 页）")
    return result


def format_page_spec(indices: Iterable[int]) -> str:
    """把 0 基页码列表压成 "1-3,5" 这样的字符串（仅用于日志/文件名）。"""
    nums = sorted(i + 1 for i in indices)
    if not nums:
        return ""
    parts: list[str] = []
    start = prev = nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
            continue
        parts.append(f"{start}-{prev}" if start != prev else str(start))
        start = prev = n
    parts.append(f"{start}-{prev}" if start != prev else str(start))
    spec = ",".join(parts)
    return spec if len(spec) <= 60 else spec[:57] + "..."


# ---------------------------------------------------------------- 读取器
def open_reader(path: str, password: str | None = None) -> PdfReader:
    """打开 PDF，自动处理加密与损坏。

    - 未加密：直接返回
    - 已加密且密码正确：decrypt 后返回
    - 已加密但未给密码：抛 PasswordRequired（界面据此弹密码框）
    - 密码错误：抛 WrongPassword
    - 损坏：抛 BrokenFile
    """
    try:
        reader = PdfReader(path, strict=False)
    except PdfReadError as exc:
        raise BrokenFile(f"PDF 结构损坏或不是有效的 PDF：{exc}") from exc
    except Exception as exc:
        raise BrokenFile(f"无法打开文件：{exc}") from exc

    if reader.is_encrypted:
        if password is None:
            raise PasswordRequired(path)
        try:
            ok = reader.decrypt(password)
        except Exception as exc:
            raise WrongPassword(f"密码解析失败：{exc}") from exc
        if not ok:
            raise WrongPassword("密码错误")
    return reader


def page_count(path: str, password: str | None = None) -> int:
    return len(open_reader(path, password).pages)


def writer_from(reader: PdfReader) -> PdfWriter:
    """拷贝 reader 的全部页面到一个新 Writer（保留书签与页面属性）。"""
    w = PdfWriter()
    w.append(reader)
    return w


def is_pdf(path: str) -> bool:
    try:
        with open(path, "rb") as fh:
            return fh.read(5).startswith(b"%PDF")
    except Exception:
        return False


# ---------------------------------------------------------------- 输出
def write_pdf(writer: PdfWriter, out_path: str, *, compress: bool = True) -> str:
    """写出 PDF，返回路径。

    注意：pypdf 6.x 起 ``page.compress_content_streams()`` 要求页面已经属于某个
    PdfWriter（否则抛 "Page must be part of a PdfWriter"）。因此这里在 write 前
    统一遍历 ``writer.pages`` 压缩，而不是在调用方对 reader.pages 操作。
    """
    try:
        if compress:
            for page in writer.pages:
                try:
                    page.compress_content_streams(level=9)
                except Exception:
                    pass
        with open(out_path, "wb") as fh:
            writer.write(fh)
    except OSError as exc:
        raise BrokenFile(f"写入失败（路径过长或无权限）：{exc}") from exc
    return out_path


def dedupe_objects(writer: PdfWriter) -> None:
    """对象去重 + 孤儿清理。兼容 pypdf 6.x(旧参数) 与 7.x(新参数)。"""
    try:
        writer.compress_identical_objects(remove_duplicates=True, remove_unreferenced=True)
        return
    except TypeError:
        pass
    except Exception:
        return
    try:
        writer.compress_identical_objects(remove_identicals=True, remove_orphans=True)
    except Exception:
        pass


# ---------------------------------------------------------------- 中文字体
_CJK_FONT = "STSong-Light"
_cjk_ready = False


def ensure_cjk_font() -> str:
    """注册 reportlab 内置的中日韩字体，用于页码/水印/页眉页脚显示中文。

    STSong-Light 随 reportlab 分发（CID 字体），无需外部 ttf，也不受
    系统字体差异影响 —— 这是跨平台显示中文最稳的方案。
    """
    global _cjk_ready
    if _cjk_ready:
        return _CJK_FONT
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont

        pdfmetrics.registerFont(UnicodeCIDFont(_CJK_FONT))
        _cjk_ready = True
    except Exception:
        pass
    return _CJK_FONT


def overlay_bytes(canvas_draw, width: float, height: float) -> bytes:
    """生成一页透明覆盖层，返回 PDF 字节（供 merge_page 使用）。

    :param canvas_draw: 接收 (canvas, width, height) 的绘制回调
    """
    from reportlab.pdfgen import canvas as rl_canvas

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(width, height))
    canvas_draw(c, width, height)
    c.showPage()
    c.save()
    return buf.getvalue()


def overlay_reader(data: bytes) -> PdfReader:
    return PdfReader(io.BytesIO(data))


def normalize_rotation(page) -> None:
    """把页面的 /Rotate 旋转写入内容流并清零。

    否则在已旋转页面上叠加内容会算错坐标系（水印歪 90° 的经典原因）。
    """
    try:
        if getattr(page, "rotation", 0):
            page.transfer_rotation_to_content()
    except Exception:
        pass
