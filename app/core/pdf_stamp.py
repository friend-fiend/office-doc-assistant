"""PDF 叠加类操作：页码、水印、页眉页脚。

原理（与 PDF24 / Adobe 一致）：
  用 reportlab 生成一张与目标页等大的"透明覆盖层" PDF，
  再用 pypdf 的 merge_page 叠加到原页内容之上。
  覆盖层不写入字体子集之外的内容，体积增量通常在几 KB 量级。
"""
from __future__ import annotations

import math
import os

from pypdf import PdfWriter

from ..errors import ParamError
from ..paths import sanitize_name, unique_file
from .pdfutil import (
    ensure_cjk_font,
    normalize_rotation,
    open_reader,
    overlay_bytes,
    overlay_reader,
    parse_ranges,
    write_pdf,
)
from .registry import OpContext, OpDef, Param, register

_POS_ALIGN = {
    "bottom-left": ("left", "bottom"),
    "bottom-center": ("center", "bottom"),
    "bottom-right": ("right", "bottom"),
    "top-left": ("left", "top"),
    "top-center": ("center", "top"),
    "top-right": ("right", "top"),
}

_POS_PARAM = Param("position", "位置", "position", default="bottom-center",
                   choices=tuple((k, v) for k, v in (
                       ("bottom-center", "底部居中"), ("bottom-right", "右下角"),
                       ("bottom-left", "左下角"), ("top-center", "顶部居中"),
                       ("top-right", "右上角"), ("top-left", "左上角"))))


# ================================================================ 1. 页码
PAGENUM_OP = OpDef(
    key="pdf_page_number",
    title="PDF 加页码",
    category="pdf_pages",
    accepts=("pdf",),
    mode="per_file",
    description="在页面指定位置添加页码，可设置格式、起始编号、字体大小，并跳过封面。",
    params=(
        Param("pnum_format", "页码格式", "choice", default="第 {n} 页",
              choices=(
                  ("{n}", "1"),
                  ("第 {n} 页", "第 1 页"),
                  ("- {n} -", "- 1 -"),
                  ("{n} / {total}", "1 / 10"),
                  ("第 {n} 页 / 共 {total} 页", "第 1 页 / 共 10 页"),
                  ("Page {n} of {total}", "Page 1 of 10"),
              ),
              hint="{n} = 当前页号，{total} = 总页数"),
        _POS_PARAM,
        Param("start_at", "起始编号", "int", default=1, min_value=0, max_value=99999),
        Param("skip_first", "跳过前 N 页（封面）", "int", default=1, min_value=0, max_value=999,
              hint="跳过的页不显示页码，但仍参与页数统计"),
        Param("font_size", "字号", "int", default=10, min_value=6, max_value=72),
        Param("margin_mm", "边距 (mm)", "int", default=12, min_value=2, max_value=60),
        Param("font_style", "字体", "choice", default="cjk",
              choices=(("cjk", "宋体（支持中文）"), ("helv", "Helvetica（仅西文）"))),
    ),
)


@register(PAGENUM_OP)
def run_page_number(ctx: OpContext, files: list[str], params: dict) -> None:
    fmt = str(params.get("pnum_format") or "{n}")
    skip = int(params.get("skip_first") or 0)
    start_at = int(params.get("start_at") or 1)
    size = int(params.get("font_size") or 10)
    margin_mm = int(params.get("margin_mm") or 12)
    position = params.get("position") or "bottom-center"
    margin = margin_mm * 72 / 25.4
    font = _font_for(params)

    ctx.set_total(len(files))
    for f in files:
        ctx.check()
        stem = os.path.splitext(os.path.basename(f))[0]
        try:
            reader = _open_with_password(ctx, f)
            total = len(reader.pages)
            w = PdfWriter()
            for i, page in enumerate(reader.pages):
                ctx.check()
                if i >= skip:
                    normalize_rotation(page)
                    pw, ph = float(page.mediabox.width), float(page.mediabox.height)
                    text = _format_pagenum(fmt, start_at + (i - skip), total)
                    data = overlay_bytes(
                        lambda c, W, H, t=text: _draw_at(c, W, H, t, font, size, position, margin),
                        pw, ph,
                    )
                    page.merge_page(overlay_reader(data).pages[0], over=True)
                w.add_page(page)
            out = unique_file(os.path.join(ctx.out_dir, f"{sanitize_name(stem)}_加页码.pdf"))
            write_pdf(w, out, compress=False)
            w.close()
            ctx.ok(src=f)
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"加页码 {os.path.basename(f)}")


def _format_pagenum(fmt: str, n: int, total: int) -> str:
    try:
        return fmt.replace("{total}", str(total)).replace("{n}", str(n))
    except Exception:
        return str(n)


def _font_for(params: dict) -> str:
    if (params.get("font_style") or "cjk") == "helv":
        return "Helvetica"
    return ensure_cjk_font()


def _measure(text: str, font: str, size: float) -> float:
    try:
        from reportlab.pdfbase import pdfmetrics

        return pdfmetrics.stringWidth(text, font, size)
    except Exception:
        return len(text) * size * 0.6


def _draw_at(c, W: float, H: float, text: str, font: str, size: float,
             position: str, margin: float) -> None:
    """在页面上按九宫格位置绘制一行文字。"""
    align, vpos = _POS_ALIGN.get(position, ("center", "bottom"))
    tw = _measure(text, font, size)
    if align == "left":
        x = margin
    elif align == "right":
        x = W - margin - tw
    else:
        x = (W - tw) / 2.0
    y = margin if vpos == "bottom" else H - margin - size
    c.setFont(font, size)
    c.setFillColorRGB(0.15, 0.15, 0.15)
    c.drawString(x, y, text)


# ================================================================ 2. 水印
WATERMARK_OP = OpDef(
    key="pdf_watermark",
    title="PDF 加水印",
    category="advanced",
    accepts=("pdf",),
    mode="per_file",
    description="添加文字水印，可调字号、颜色、透明度、旋转角度，支持平铺。",
    params=(
        Param("wm_text", "水印文字", "entry", default="内部资料 请勿外传", width=24),
        Param("wm_size", "字号", "int", default=40, min_value=6, max_value=200),
        Param("wm_opacity", "透明度", "float", default=0.18, min_value=0.02, max_value=1.0,
              hint="0.02 ~ 1.0，越小越淡"),
        Param("wm_angle", "旋转角度", "choice", default="45",
              choices=(("0", "水平"), ("45", "45°"), ("-45", "-45°"), ("90", "垂直"))),
        Param("wm_color", "颜色", "choice", default="red",
              choices=(("red", "红色"), ("gray", "灰色"), ("blue", "蓝色"),
                       ("black", "黑色"), ("green", "绿色"))),
        Param("wm_tile", "平铺（满页重复）", "bool", default=False),
        Param("wm_pages", "仅这些页", "range", default="", width=20,
              hint="留空 = 全部页；支持 1-3,5"),
        Param("wm_behind", "置于内容下方", "bool", default=False),
    ),
)

_COLORS = {
    "red": (0.85, 0.10, 0.10),
    "gray": (0.45, 0.45, 0.45),
    "blue": (0.10, 0.25, 0.75),
    "black": (0.0, 0.0, 0.0),
    "green": (0.10, 0.55, 0.20),
}


@register(WATERMARK_OP)
def run_watermark(ctx: OpContext, files: list[str], params: dict) -> None:
    text = str(params.get("wm_text") or "").strip()
    if not text:
        raise ParamError("水印文字不能为空")
    size = int(params.get("wm_size") or 40)
    opacity = float(params.get("wm_opacity") or 0.18)
    angle = float(params.get("wm_angle") or 45)
    color = _COLORS.get(str(params.get("wm_color") or "red"), _COLORS["red"])
    tile = bool(params.get("wm_tile"))
    behind = bool(params.get("wm_behind"))
    spec = str(params.get("wm_pages") or "").strip()
    font = ensure_cjk_font()

    ctx.set_total(len(files))
    for f in files:
        ctx.check()
        stem = os.path.splitext(os.path.basename(f))[0]
        try:
            reader = _open_with_password(ctx, f)
            total = len(reader.pages)
            targets = set(parse_ranges(spec, total, clip=True)) if spec else set(range(total))
            w = PdfWriter()
            for i, page in enumerate(reader.pages):
                ctx.check()
                if i in targets:
                    normalize_rotation(page)
                    pw, ph = float(page.mediabox.width), float(page.mediabox.height)
                    data = overlay_bytes(
                        lambda c, W, H: _draw_watermark(c, W, H, text, font, size,
                                                        opacity, angle, color, tile),
                        pw, ph,
                    )
                    page.merge_page(overlay_reader(data).pages[0], over=not behind)
                w.add_page(page)
            suffix = "_水印" if not behind else "_水印底层"
            out = unique_file(os.path.join(ctx.out_dir, f"{sanitize_name(stem)}{suffix}.pdf"))
            write_pdf(w, out, compress=False)
            w.close()
            ctx.ok(src=f)
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"加水印 {os.path.basename(f)}")


def _draw_watermark(c, W: float, H: float, text: str, font: str, size: float,
                    opacity: float, angle: float, color: tuple, tile: bool) -> None:
    try:
        c.setFillAlpha(opacity)
        c.setStrokeAlpha(opacity)
    except Exception:
        pass  # 极老版本的 reportlab 没有 alpha，退化为实色
    c.setFont(font, size)
    c.setFillColorRGB(*color)

    if tile:
        step_x = max(_measure(text, font, size) + size * 2, 120)
        step_y = max(size * 4, 120)
        rows = int(H / step_y) + 2
        cols = int(W / step_x) + 2
        for r in range(rows):
            for col in range(cols):
                c.saveState()
                c.translate(col * step_x - step_x / 3, r * step_y + step_y / 3)
                c.rotate(angle)
                c.drawString(0, 0, text)
                c.restoreState()
        return

    c.saveState()
    c.translate(W / 2.0, H / 2.0)
    c.rotate(angle)
    tw = _measure(text, font, size)
    c.drawString(-tw / 2.0, -size / 3.0, text)
    c.restoreState()


# ================================================================ 3. 页眉页脚
HEADERFOOTER_OP = OpDef(
    key="pdf_header_footer",
    title="PDF 页眉页脚",
    category="advanced",
    accepts=("pdf",),
    mode="per_file",
    description="同时添加页眉与页脚，支持占位符 {n}/{total}/{date}/{filename}。",
    params=(
        Param("header_text", "页眉文字", "entry", default="", width=26,
              hint="可用占位符 {n} {total} {date} {filename}；留空则不加页眉"),
        Param("footer_text", "页脚文字", "entry", default="第 {n} 页 / 共 {total} 页", width=26),
        Param("hf_align", "对齐", "choice", default="center",
              choices=(("left", "左对齐"), ("center", "居中"), ("right", "右对齐"))),
        Param("hf_size", "字号", "int", default=9, min_value=5, max_value=48),
        Param("hf_margin", "边距 (mm)", "int", default=10, min_value=2, max_value=60),
    ),
)


@register(HEADERFOOTER_OP)
def run_header_footer(ctx: OpContext, files: list[str], params: dict) -> None:
    header = str(params.get("header_text") or "")
    footer = str(params.get("footer_text") or "")
    if not header.strip() and not footer.strip():
        raise ParamError("页眉和页脚至少填写一项")
    align = str(params.get("hf_align") or "center")
    size = int(params.get("hf_size") or 9)
    margin = int(params.get("hf_margin") or 10) * 72 / 25.4
    font = ensure_cjk_font()
    import datetime

    today = datetime.date.today().strftime("%Y-%m-%d")

    ctx.set_total(len(files))
    for f in files:
        ctx.check()
        stem = os.path.splitext(os.path.basename(f))[0]
        try:
            reader = _open_with_password(ctx, f)
            total = len(reader.pages)
            w = PdfWriter()
            for i, page in enumerate(reader.pages):
                ctx.check()
                normalize_rotation(page)
                pw, ph = float(page.mediabox.width), float(page.mediabox.height)
                ctx_data = {"n": str(i + 1), "total": str(total), "date": today,
                            "filename": os.path.basename(f)}

                def draw(c, W, H, _d=ctx_data):
                    if header.strip():
                        _draw_line(c, W, H, _render(header, _d), font, size, align, "top", margin)
                    if footer.strip():
                        _draw_line(c, W, H, _render(footer, _d), font, size, align, "bottom", margin)

                page.merge_page(overlay_reader(overlay_bytes(draw, pw, ph)).pages[0], over=True)
                w.add_page(page)
            out = unique_file(os.path.join(ctx.out_dir, f"{sanitize_name(stem)}_页眉页脚.pdf"))
            write_pdf(w, out, compress=False)
            w.close()
            ctx.ok(src=f)
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"页眉页脚 {os.path.basename(f)}")


def _render(template: str, data: dict) -> str:
    out = template
    for k, v in data.items():
        out = out.replace("{" + k + "}", v)
    return out


def _draw_line(c, W: float, H: float, text: str, font: str, size: float,
               align: str, vpos: str, margin: float) -> None:
    tw = _measure(text, font, size)
    if align == "left":
        x = margin
    elif align == "right":
        x = W - margin - tw
    else:
        x = (W - tw) / 2.0
    y = margin if vpos == "bottom" else H - margin - size
    c.setFont(font, size)
    c.setFillColorRGB(0.25, 0.25, 0.25)
    c.drawString(x, y, text)


def _open_with_password(ctx: OpContext, path: str):
    from .pdf_pages import _open_with_password as op

    return op(ctx, path)
