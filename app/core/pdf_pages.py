"""PDF 页面操作：合并、拆分、提取指定页、删除页、旋转。

全部基于 pypdf，不依赖任何外部程序，因此在缺少 Poppler/LibreOffice/Ghostscript
的环境下这些功能依然可用。
"""
from __future__ import annotations

import os

from pypdf import PdfWriter

from ..errors import BrokenFile
from ..paths import sanitize_name, unique_file
from .pdfutil import format_page_spec, open_reader, parse_ranges, write_pdf
from .registry import OpContext, OpDef, Param, register

# ================================================================ 1. 合并
MERGE_OP = OpDef(
    key="pdf_merge",
    title="PDF 合并",
    category="pdf_pages",
    accepts=("pdf",),
    mode="collection",
    description="把列表中的多个 PDF 按当前顺序合并成一个文件，支持拖动调整顺序。",
    params=(
        Param("out_name", "输出文件名", "entry", default="合并结果", width=22,
              hint="不需要写 .pdf 后缀"),
        Param("add_bookmarks", "为每个来源文件生成书签", "bool", default=True),
    ),
)


@register(MERGE_OP)
def run_merge(ctx: OpContext, files: list[str], params: dict) -> None:
    total = len(files)
    ctx.set_total(total)
    ctx.log(f"开始合并 {total} 个文件 …")

    writer = PdfWriter()
    merged_from: list[tuple[str, int, int]] = []  # (文件名, 起始页, 页数)
    ok_files = 0
    total_pages = 0

    for f in files:
        ctx.check()
        try:
            reader = _open_with_password(ctx, f)
            start = len(writer.pages)
            for page in reader.pages:
                ctx.check()
                writer.add_page(page)
            merged_from.append((os.path.basename(f), start, len(reader.pages)))
            ok_files += 1
            total_pages += len(reader.pages)
            ctx.ok(src=f)
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"合并 {os.path.basename(f)}")

    if not writer.pages:
        raise BrokenFile("没有任何可用页面，合并中止")

    if params.get("add_bookmarks", True):
        try:
            for name, start, count in merged_from:
                writer.add_outline_item(name, start)
        except Exception as exc:  # 书签失败不影响主流程
            ctx.log(f"书签生成失败（已忽略）：{exc}", "warn")

    out_name = sanitize_name(str(params.get("out_name") or "合并结果"))
    if not out_name.lower().endswith(".pdf"):
        out_name += ".pdf"
    out_path = unique_file(os.path.join(ctx.out_dir, out_name))
    write_pdf(writer, out_path)
    writer.close()

    ctx.metric("合并来源", f"{ok_files} 个文件")
    ctx.metric("总页数", total_pages)
    ctx.metric("输出文件", out_path)
    ctx.log(f"合并完成 → {out_path}", "info")


# ================================================================ 2. 拆分
SPLIT_OP = OpDef(
    key="pdf_split",
    title="PDF 拆分",
    category="pdf_pages",
    accepts=("pdf",),
    mode="per_file",
    description="按「每页一个 / 每 N 页一个 / 按自定义页码范围」把 PDF 拆成多个文件。",
    params=(
        Param("split_mode", "拆分方式", "choice", default="single",
              choices=(("single", "每页一个文件"), ("every", "每 N 页一个文件"),
                       ("ranges", "按页码范围拆分")),
              hint="页码范围写法：1-3,5,8-10"),
        Param("every_n", "每 N 页", "int", default=2, min_value=1, max_value=999,
              visible_if=("split_mode", ("every",))),
        Param("ranges", "页码范围", "range", default="1-3,4-6",
              visible_if=("split_mode", ("ranges",))),
    ),
)


@register(SPLIT_OP)
def run_split(ctx: OpContext, files: list[str], params: dict) -> None:
    ctx.set_total(len(files))
    for f in files:
        ctx.check()
        stem = os.path.splitext(os.path.basename(f))[0]
        try:
            reader = _open_with_password(ctx, f)
            total_pages = len(reader.pages)
            if total_pages == 0:
                ctx.skip(f, "文件没有任何页面")
                continue
            groups = _split_groups(params, total_pages)
            made = 0
            for label, indices in groups:
                ctx.check()
                w = PdfWriter()
                for i in indices:
                    w.add_page(reader.pages[i])
                out = unique_file(os.path.join(ctx.out_dir, f"{sanitize_name(stem)}_{label}.pdf"))
                write_pdf(w, out, compress=False)
                w.close()
                made += 1
            ctx.report.outputs.append(ctx.out_dir)
            ctx.ok(src=f)
            ctx.log(f"  ↳ 共生成 {made} 个文件")
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"拆分 {os.path.basename(f)}")


def _split_groups(params: dict, total: int) -> list[tuple[str, list[int]]]:
    """返回 [(文件名片段, 0基页码列表), ...]"""
    mode = params.get("split_mode", "single")
    if mode == "every":
        n = max(1, int(params.get("every_n") or 1))
        groups = []
        for start in range(0, total, n):
            indices = list(range(start, min(start + n, total)))
            first, last = indices[0] + 1, indices[-1] + 1
            label = f"p{first}" if first == last else f"p{first}-{last}"
            groups.append((label, indices))
        return groups
    if mode == "ranges":
        spec = str(params.get("ranges") or "")
        raw = [c.strip() for c in spec.replace("，", ",").split(",") if c.strip()]
        groups = []
        for idx, chunk in enumerate(raw, start=1):
            try:
                indices = parse_ranges(chunk, total, clip=False)
            except Exception:
                indices = parse_ranges(chunk, total, clip=True)
            if indices:
                groups.append((f"{format_page_spec(indices)}".replace(",", "_"), indices))
        if not groups:
            return [("p1", list(range(total)))]
        return groups
    # single
    return [(f"p{i + 1}", [i]) for i in range(total)]


# ================================================================ 3. 提取指定页
EXTRACT_OP = OpDef(
    key="pdf_extract",
    title="提取指定页",
    category="pdf_pages",
    accepts=("pdf",),
    mode="per_file",
    description="只保留指定页码，生成一个新的 PDF。如 1-3,5,8-10。",
    params=(
        Param("pages", "要提取的页", "range", default="1-3,5", width=24,
              hint="从 1 开始计数，支持 1-3,5,8-10"),
        Param("merge_all", "把所有文件提取结果再合并成一个", "bool", default=False),
    ),
)


@register(EXTRACT_OP)
def run_extract(ctx: OpContext, files: list[str], params: dict) -> None:
    ctx.set_total(len(files))
    merge_all = bool(params.get("merge_all"))
    combined = PdfWriter() if merge_all else None
    kept_files = 0

    for f in files:
        ctx.check()
        stem = os.path.splitext(os.path.basename(f))[0]
        try:
            reader = _open_with_password(ctx, f)
            indices = parse_ranges(str(params.get("pages") or ""), len(reader.pages), clip=True)
            w = PdfWriter()
            for i in indices:
                w.add_page(reader.pages[i])
            if merge_all and combined is not None:
                for page in w.pages:
                    combined.add_page(page)
            else:
                out = unique_file(os.path.join(
                    ctx.out_dir, f"{sanitize_name(stem)}_提取_{format_page_spec(indices)}.pdf"))
                write_pdf(w, out, compress=False)
            w.close()
            kept_files += 1
            ctx.ok(src=f)
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"提取 {os.path.basename(f)}")

    if merge_all and combined is not None and combined.pages:
        out = unique_file(os.path.join(ctx.out_dir, "提取结果_合并.pdf"))
        write_pdf(combined, out, compress=False)
        ctx.metric("合并输出", out)
    ctx.metric("处理文件数", kept_files)


# ================================================================ 4. 删除页
DELETE_OP = OpDef(
    key="pdf_delete_pages",
    title="删除指定页",
    category="pdf_pages",
    accepts=("pdf",),
    mode="per_file",
    confirm=True,
    confirm_text="将从副本中删除指定页面。原文件不会被修改，确认继续吗？",
    description="删除指定页码后另存为新文件（原文件保持不变）。",
    params=(
        Param("pages", "要删除的页", "range", default="", width=24,
              hint="从 1 开始计数，支持 1-3,5,8-10"),
    ),
)


@register(DELETE_OP)
def run_delete_pages(ctx: OpContext, files: list[str], params: dict) -> None:
    spec = str(params.get("pages") or "").strip()
    if not spec:
        from ..errors import ParamError

        raise ParamError("请填写要删除的页码")
    ctx.set_total(len(files))
    for f in files:
        ctx.check()
        stem = os.path.splitext(os.path.basename(f))[0]
        try:
            reader = _open_with_password(ctx, f)
            total_pages = len(reader.pages)
            drop = set(parse_ranges(spec, total_pages, clip=True))
            keep = [i for i in range(total_pages) if i not in drop]
            if not keep:
                ctx.skip(f, "删除后没有任何剩余页面，已跳过")
                continue
            w = PdfWriter()
            for i in keep:
                w.add_page(reader.pages[i])
            out = unique_file(os.path.join(ctx.out_dir, f"{sanitize_name(stem)}_已删除{len(drop)}页.pdf"))
            write_pdf(w, out, compress=False)
            w.close()
            ctx.ok(src=f)
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"删除页 {os.path.basename(f)}")


# ================================================================ 5. 旋转
ROTATE_OP = OpDef(
    key="pdf_rotate",
    title="PDF 旋转",
    category="pdf_pages",
    accepts=("pdf",),
    mode="per_file",
    description="统一旋转页面方向。可只旋转指定页，留空则整篇旋转。",
    params=(
        Param("angle", "旋转角度", "choice", default="90",
              choices=(("90", "顺时针 90°"), ("180", "180°"), ("270", "顺时针 270°（逆时针 90°）"))),
        Param("pages", "仅旋转这些页", "range", default="", width=24,
              hint="留空 = 全部页；支持 1-3,5"),
    ),
)


@register(ROTATE_OP)
def run_rotate(ctx: OpContext, files: list[str], params: dict) -> None:
    angle = int(params.get("angle") or 90)
    spec = str(params.get("pages") or "").strip()
    ctx.set_total(len(files))
    for f in files:
        ctx.check()
        stem = os.path.splitext(os.path.basename(f))[0]
        try:
            reader = _open_with_password(ctx, f)
            total_pages = len(reader.pages)
            targets = set(parse_ranges(spec, total_pages, clip=True)) if spec else set(range(total_pages))
            w = PdfWriter()
            for i, page in enumerate(reader.pages):
                ctx.check()
                if i in targets:
                    page.rotate(angle)
                w.add_page(page)
            out = unique_file(os.path.join(ctx.out_dir, f"{sanitize_name(stem)}_旋转{angle}.pdf"))
            write_pdf(w, out, compress=False)
            w.close()
            ctx.ok(src=f)
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"旋转 {os.path.basename(f)}")


# ================================================================ 公共：密码
def _open_with_password(ctx: OpContext, path: str):
    """打开 PDF；遇到加密文件时向界面询问密码（最多两次）。"""
    from ..errors import PasswordRequired, WrongPassword
    from .pdfutil import open_reader

    attempts = 0
    password: str | None = None
    while True:
        try:
            return open_reader(path, password)
        except PasswordRequired:
            attempts += 1
            pwd = ctx.ask_password(path, confirm=False)
            if pwd is None:
                raise BrokenFile("用户取消了密码输入")
            password = pwd
        except WrongPassword:
            attempts += 1
            if attempts >= 3:
                raise BrokenFile("密码连续错误，跳过该文件")
            pwd = ctx.ask_password(path, confirm=True)
            if pwd is None:
                raise BrokenFile("用户取消了密码输入")
            password = pwd
