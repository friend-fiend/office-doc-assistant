"""图片转 PDF：多图排序合成一个 PDF。

Pillow 直接把多张图写成多页 PDF。两个容易出问题的细节这里都处理了：
  · 带透明通道（PNG/WebP）的图直接转 PDF 会出现黑块 —— 先合成到白底；
  · 不同尺寸的图直接输出会让页面大小参差 —— 提供"统一页尺寸"选项做等比居中。
"""
from __future__ import annotations

import os
import re

from ..errors import ParamError
from ..paths import sanitize_name, unique_file
from .registry import OpContext, OpDef, Param, register

IMG2PDF_OP = OpDef(
    key="image_to_pdf",
    title="图片转 PDF",
    category="convert",
    accepts=("image",),
    mode="collection",
    description=(
        "把列表中的多张图片按当前顺序合成一个 PDF。\n"
        "支持拖动调整顺序；可统一页面尺寸或按原图尺寸输出。"
    ),
    params=(
        Param("out_name", "输出文件名", "entry", default="图片合成", width=22,
              hint="不需要写 .pdf 后缀"),
        Param("page_size", "页面尺寸", "choice", default="a4",
              choices=(("a4", "A4 纵向（图片等比居中）"), ("a4l", "A4 横向"),
                       ("original", "跟随原图尺寸"), ("letter", "Letter 纵向"))),
        Param("margin_mm", "页边距 (mm)", "int", default=10, min_value=0, max_value=60,
              visible_if=("page_size", ("a4", "a4l", "letter"))),
        Param("quality", "JPEG 质量", "int", default=90, min_value=40, max_value=100),
        Param("sort_natural", "按文件名智能排序", "bool", default=True,
              hint="开启后 图片2 会排在 图片10 前面"),
    ),
)

_PAGES = {
    "a4": (595.28, 841.89),
    "a4l": (841.89, 595.28),
    "letter": (612.0, 792.0),
}


@register(IMG2PDF_OP)
def run_image_to_pdf(ctx: OpContext, files: list[str], params: dict) -> None:
    from PIL import Image

    if params.get("sort_natural", True):
        files = sorted(files, key=natural_key)

    order = [os.path.basename(f) for f in files]
    ctx.log("合成顺序：" + " → ".join(order[:12]) + (" …" if len(order) > 12 else ""))

    page_size = str(params.get("page_size") or "a4")
    margin = int(params.get("margin_mm") or 0) * 72 / 25.4
    quality = int(params.get("quality") or 90)

    pages: list = []
    for f in files:
        ctx.check()
        try:
            img = Image.open(f)
            img.load()
            pages.append(to_rgb(img))
            ctx.ok(src=f)
        except Exception as exc:
            ctx.fail(f, f"无法读取图片：{exc}")

    if not pages:
        raise ParamError("没有可用图片，无法合成 PDF")

    out_name = sanitize_name(str(params.get("out_name") or "图片合成"))
    if not out_name.lower().endswith(".pdf"):
        out_name += ".pdf"
    out_path = unique_file(os.path.join(ctx.out_dir, out_name))

    if page_size == "original":
        final = pages
    else:
        pw, ph = _PAGES.get(page_size, _PAGES["a4"])
        final = [fit_on_page(img, pw, ph, margin) for img in pages]

    ctx.set_total(len(final))
    ctx.log(f"正在写出 {len(final)} 页 …")
    # resolution=72.0 是关键：Pillow 的 PDF 插件按 点数 = 像素 / resolution * 72
    # 换算页面尺寸。取 72 时像素与点数 1:1，上面算好的 595x842 才会真正落成 A4；
    # 若沿用 200 会得到 214x303pt 的"缩水 A4"（这是个很容易踩的坑）。
    final[0].save(
        out_path, "PDF", resolution=72.0,
        save_all=True, append_images=final[1:],
        quality=quality, optimize=True,
    )
    for img in final:
        if img not in pages:
            img.close()
    for img in pages:
        img.close()

    ctx.metric("页数", len(final))
    ctx.metric("输出文件", out_path)
    ctx.report.outputs.append(out_path)
    ctx.log(f"合成完成 → {out_path}")


def to_rgb(img):
    """统一转 RGB；带透明通道的合成到白底，避免 PDF 里变黑块。"""
    from PIL import Image

    if img.mode == "RGB":
        return img
    if img.mode in ("RGBA", "LA", "P"):
        rgba = img.convert("RGBA")
        bg = Image.new("RGB", rgba.size, (255, 255, 255))
        bg.paste(rgba, mask=rgba.split()[-1])
        rgba.close()
        return bg
    return img.convert("RGB")


def fit_on_page(img, pw: float, ph: float, margin: float):
    """把图片等比缩放并居中放到指定尺寸的白色页面上。"""
    from PIL import Image

    box_w = max(1.0, pw - margin * 2)
    box_h = max(1.0, ph - margin * 2)
    scale = min(box_w / img.width, box_h / img.height)
    new_size = (max(1, int(img.width * scale)), max(1, int(img.height * scale)))
    resized = img.resize(new_size, Image.LANCZOS)
    canvas = Image.new("RGB", (int(pw), int(ph)), (255, 255, 255))
    canvas.paste(resized, ((int(pw) - new_size[0]) // 2, (int(ph) - new_size[1]) // 2))
    resized.close()
    return canvas


def natural_key(path: str):
    """自然排序：文件名里的数字段按数值比较。"""
    name = os.path.basename(path).lower()
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", name)]
