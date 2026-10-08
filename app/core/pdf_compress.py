"""PDF 压缩：三档策略（参考 PDF24 Tools 的组合思路）。

轻度 —— 纯 pypdf，无损：
    · 内容流 zlib 重压（FlateDecode）
    · 相同对象去重、孤儿对象清理
    · 元数据精简、XMP 删除
中度 / 强力 —— Ghostscript 重建：
    · pdfwrite 设备重新写出，图片按 DPI 降采样，JPEG 质量 80 / 60
    · 字体子集化、重复图片合并

Ghostscript 缺失时自动降级为轻度策略，并在日志与结果里明确提示
"压缩效果可能有限"，符合需求中的降级要求。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile

from pypdf import PdfReader, PdfWriter

from ..config import GS_PRESETS, IS_WINDOWS
from ..deps import get as dep_get
from ..paths import human_size, sanitize_name, unique_file
from .pdfutil import dedupe_objects, open_reader, write_pdf
from .registry import OpContext, OpDef, Param, register

_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000) if IS_WINDOWS else 0


COMPRESS_OP = OpDef(
    key="pdf_compress",
    title="PDF 压缩",
    category="compress",
    accepts=("pdf",),
    mode="per_file",
    description=(
        "轻度=无损结构优化；中度/强力=Ghostscript 重建并降采样图片。\n"
        "处理结束后会显示压缩前/后大小与缩减比例。"
    ),
    params=(
        Param("level", "压缩档位", "choice", default="medium",
              choices=(("light", "轻度（无损优先）"),
                       ("medium", "中度（推荐，150 DPI / JPEG 80）"),
                       ("strong", "强力（96 DPI / JPEG 60，体积最小）"))),
        Param("strip_metadata", "清除文档元数据（作者/标题等）", "bool", default=True),
        Param("keep_if_bigger", "若压缩后反而变大则保留原文件副本", "bool", default=True),
    ),
)


@register(COMPRESS_OP)
def run_compress(ctx: OpContext, files: list[str], params: dict) -> None:
    level = str(params.get("level") or "medium")
    strip = bool(params.get("strip_metadata", True))
    keep_if_bigger = bool(params.get("keep_if_bigger", True))

    gs = dep_get("ghostscript")
    use_gs = level in GS_PRESETS and gs.available
    if level in GS_PRESETS and not gs.available:
        ctx.log("未检测到 Ghostscript —— 已自动降级为「pypdf 结构优化」，压缩效果可能有限。", "warn")
        ctx.report.metrics["降级提示"] = "Ghostscript 缺失，使用 pypdf 结构优化，压缩效果可能有限"
    if use_gs:
        ctx.log(f"使用 Ghostscript v{gs.version} 进行深度压缩（档位：{level}）。")

    ctx.set_total(len(files))
    total_before = 0
    total_after = 0

    for f in files:
        ctx.check()
        stem = os.path.splitext(os.path.basename(f))[0]
        try:
            before = os.path.getsize(f)
            out = unique_file(os.path.join(ctx.out_dir, f"{sanitize_name(stem)}_压缩.pdf"))

            if use_gs:
                try:
                    _ghostscript_compress(ctx, gs.path, f, out, level)
                except Exception as exc:
                    ctx.log(f"Ghostscript 压缩失败，回退到结构优化：{exc}", "warn")
                    _light_compress(ctx, f, out, strip)
            else:
                _light_compress(ctx, f, out, strip)

            after = os.path.getsize(out)
            if keep_if_bigger and after >= before:
                # 压缩没效果时，仍然保留输出文件，但明确告知用户
                ctx.warn(f, f"压缩后体积未减小（{human_size(before)} → {human_size(after)}），"
                            f"该文件可能已是优化过的最佳状态")

            total_before += before
            total_after += after
            ratio = (1 - after / before) * 100 if before else 0
            ctx.log(f"  {human_size(before)} → {human_size(after)}（缩减 {ratio:.1f}%）")
            ctx.ok(src=f)
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"压缩 {os.path.basename(f)}")

    if total_before:
        saved = total_before - total_after
        ctx.metric("压缩前总大小", human_size(total_before))
        ctx.metric("压缩后总大小", human_size(total_after))
        ctx.metric("总体缩减", f"{saved / 1024 / 1024:.2f} MB（{(1 - total_after / total_before) * 100:.1f}%）")


# ------------------------------------------------------------------ 轻度（无损）
def _light_compress(ctx: OpContext, src: str, dst: str, strip: bool) -> None:
    """轻度压缩（尽量无损）。

    顺序很重要：pypdf 6.x 要求页面先归属 PdfWriter 才能重压内容流，
    所以这里"先 append 再压缩"。
    """
    reader = open_reader(src)
    writer = PdfWriter()
    for page in reader.pages:
        ctx.check()
        writer.add_page(page)

    # 内容流重压 + 对象去重（在 writer 上操作）
    for page in writer.pages:
        ctx.check()
        try:
            page.compress_content_streams(level=9)
        except Exception:
            pass
    dedupe_objects(writer)

    # 元数据精简：清空 Info 字典 + 尝试摘掉 XMP 流（XMP 常带几十 KB 冗余）
    if strip:
        try:
            writer.add_metadata({})
        except Exception:
            pass
        try:
            root = getattr(writer, "_root_object", None)
            if root is not None and "/Metadata" in root:
                del root["/Metadata"]
        except Exception:
            pass

    write_pdf(writer, dst, compress=False)  # 已在上面压缩过，避免重复遍历
    writer.close()


# ------------------------------------------------------------------ 中/强力（Ghostscript）
def _ghostscript_compress(ctx: OpContext, gs_path: str, src: str, dst: str, level: str) -> None:
    preset = GS_PRESETS.get(level, GS_PRESETS["medium"])

    args: list[str] = [
        gs_path,
        "-q", "-dNOPAUSE", "-dBATCH", "-dSAFER", "-dQUIET",
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.7",
        "-dDetectDuplicateImages=true",
        "-dCompressFonts=true",
        "-dSubsetFonts=true",
        "-dEmbedAllFonts=true",
        "-dAutoRotatePages=/None",
        "-dPreserveMarkedContent=true",
        # ---- 彩色图片
        "-dDownsampleColorImages=true",
        f"-dColorImageResolution={preset['color_dpi']}",
        "-dColorImageDownsampleType=/Bicubic",
        "-dAutoFilterColorImages=false",
        "-dColorImageFilter=/DCTEncode",
        f"-dJPEGQ={preset['jpeg_q']}",
        # ---- 灰度图片
        "-dDownsampleGrayImages=true",
        f"-dGrayImageResolution={preset['gray_dpi']}",
        "-dGrayImageDownsampleType=/Bicubic",
        "-dAutoFilterGrayImages=false",
        "-dGrayImageFilter=/DCTEncode",
        # ---- 单色图片
        "-dDownsampleMonoImages=true",
        f"-dMonoImageResolution={preset['mono_dpi']}",
        "-dMonoImageDownsampleType=/Subsample",
        # ---- 嵌字与透明度
        "-dHaveTransparency=true",
        "-dTransferFunctionInfo=/Remove",
        "-dColorConversionStrategy=/LeaveColorUnchanged",
    ]

    tmp_out = os.path.join(tempfile.mkdtemp(prefix="oda_gs_"), "out.pdf")
    try:
        args += [f"-sOutputFile={tmp_out}", src]
        proc = subprocess.run(
            args,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            creationflags=_NO_WINDOW, timeout=1800,
            encoding="utf-8", errors="replace",
        )
        if proc.returncode != 0:
            raise RuntimeError(f"Ghostscript 退出码 {proc.returncode}：{(proc.stdout or '')[-400:]}")
        if not os.path.exists(tmp_out) or os.path.getsize(tmp_out) == 0:
            raise RuntimeError("Ghostscript 未生成有效输出（可能已损坏或加密）")
        # 校验产物能被正常解析，避免"压缩出一个坏文件"
        PdfReader(tmp_out, strict=False)
        shutil.move(tmp_out, dst)
    finally:
        shutil.rmtree(os.path.dirname(tmp_out), ignore_errors=True)
