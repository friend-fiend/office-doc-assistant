"""PDF 转图片：每页导出 PNG/JPG，可指定分辨率与页码范围。

依赖 Poppler（pdftoppm）。若使用了 pdftocairo 后端还可获得更清晰的矢量渲染。
"""
from __future__ import annotations

import os

from ..deps import get as dep_get
from ..errors import DependencyMissing, ParamError
from ..paths import sanitize_name, unique_dir
from .pdf_pages import _open_with_password
from .pdfutil import open_reader, parse_ranges
from .registry import OpContext, OpDef, Param, register

PDF2IMG_OP = OpDef(
    key="pdf_to_image",
    title="PDF 转图片",
    category="convert",
    accepts=("pdf",),
    mode="per_file",
    requires=("poppler",),
    description="每页导出为一张 PNG 或 JPG。DPI 越高越清晰、文件越大（打印级建议 300）。",
    params=(
        Param("img_format", "图片格式", "choice", default="png",
              choices=(("png", "PNG（无损，体积大）"), ("jpeg", "JPG（有损，体积小）"))),
        Param("dpi", "分辨率 DPI", "choice", default="150",
              choices=(("72", "72（屏幕预览）"), ("150", "150（推荐）"),
                       ("200", "200"), ("300", "300（打印）"), ("600", "600（高清）"))),
        Param("jpeg_quality", "JPG 质量", "int", default=88, min_value=40, max_value=100,
              hint="仅 JPG 生效"),
        Param("pages", "页码范围", "range", default="", width=20,
              hint="留空 = 全部页；支持 1-3,5"),
        Param("each_subfolder", "每个 PDF 单独建子文件夹", "bool", default=True),
    ),
)


@register(PDF2IMG_OP)
def run_pdf_to_image(ctx: OpContext, files: list[str], params: dict) -> None:
    poppler = dep_get("poppler")
    if not poppler.available:
        raise DependencyMissing("poppler")

    try:
        from pdf2image import convert_from_path
        from pdf2image.exceptions import PDFInfoNotInstalledError, PDFPageCountError
    except Exception as exc:  # pragma: no cover
        raise DependencyMissing("poppler", f"pdf2image 不可用：{exc}") from exc

    fmt = str(params.get("img_format") or "png")
    dpi = int(params.get("dpi") or 150)
    quality = int(params.get("jpeg_quality") or 88)
    spec = str(params.get("pages") or "").strip()
    sub = bool(params.get("each_subfolder", True))

    ctx.set_total(len(files))
    total_images = 0
    for f in files:
        ctx.check()
        stem = sanitize_name(os.path.splitext(os.path.basename(f))[0])
        try:
            reader = _open_with_password(ctx, f)
            total_pages = len(reader.pages)
            indices = parse_ranges(spec, total_pages, clip=True) if spec else list(range(total_pages))
            first_page = min(indices) + 1
            last_page = max(indices) + 1

            target = unique_dir(ctx.out_dir, stem) if sub else ctx.out_dir
            kwargs: dict = dict(
                dpi=dpi,
                fmt="jpeg" if fmt == "jpeg" else "png",
                output_folder=None,  # 交给 Pillow 对象在内存中处理，便于精确控制命名
                first_page=first_page,
                last_page=last_page,
                poppler_path=poppler.directory,
                thread_count=max(1, min(4, (os.cpu_count() or 2) // 2)),
            )
            if fmt == "jpeg":
                kwargs["jpegopt"] = {"quality": quality, "optimize": True, "progressive": True}

            ctx.log(f"渲染 {os.path.basename(f)} 第 {first_page}-{last_page} 页 @ {dpi} DPI …")
            images = convert_from_path(f, **kwargs)

            for offset, img in enumerate(images):
                ctx.check()
                page_no = first_page + offset
                ext = "jpg" if fmt == "jpeg" else "png"
                out = os.path.join(target, f"{stem}_第{page_no:04d}页.{ext}")
                if fmt == "jpeg":
                    img.convert("RGB").save(out, "JPEG", quality=quality, optimize=True)
                else:
                    img.save(out, "PNG", optimize=True)
                img.close()
                total_images += 1
            if sub:
                ctx.report.outputs.append(target)
            ctx.ok(src=f)
        except PDFInfoNotInstalledError as exc:
            raise DependencyMissing("poppler", f"Poppler 不可用：{exc}") from exc
        except PDFPageCountError as exc:
            ctx.fail(f, f"PDF 损坏或无法读取页数：{exc}")
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"PDF 转图片 {os.path.basename(f)}")

    ctx.metric("导出图片数", total_images)

