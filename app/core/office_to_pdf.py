"""Office 文档（Word / Excel / PPT）批量转 PDF —— LibreOffice headless。

三个必须踩过才不翻车的点（这也是本模块存在的意义）：

1) **用户配置目录锁**
   LibreOffice 默认共用同一个 UserInstallation 目录。多个进程并发时，
   后启动的进程会把任务丢给第一个进程后立即退出，于是"退出码 0 但没有产物"。
   解决：每次调用都通过 ``-env:UserInstallation=file://<随机目录>`` 隔离配置目录。

2) **启动开销远大于转换本身**
   soffice 冷启动约 1.5~3 秒。逐个文件调用等于把启动开销乘以文件数。
   解决：一次命令行传入多个文件（分批，避免超过系统命令行长度上限）。

3) **配置目录残留导致越用越慢**
   解决：转换结束（含异常/取消）后递归清理临时配置目录。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import time

from ..config import IS_WINDOWS
from ..deps import get as dep_get
from ..errors import Cancelled, DependencyMissing
from ..paths import sanitize_name, unique_dir
from .registry import OpContext, OpDef, Param, register

_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000) if IS_WINDOWS else 0
_MAX_ARGS_CHARS = 6000  # 保守阈值，Windows 命令行上限约 32k

OFFICE2PDF_OP = OpDef(
    key="office_to_pdf",
    title="Word / Excel / PPT 转 PDF",
    category="convert",
    accepts=("office",),
    mode="per_file",
    requires=("soffice",),
    description=(
        "调用本机 LibreOffice 进行转换，最大程度保留原始排版。\n"
        "支持 .doc/.docx/.docm/.rtf/.odt/.wps、.xls/.xlsx/.csv/.ods、.ppt/.pptx/.odp。"
    ),
    params=(
        Param("calc_pages", "超出以下页数时提示（0=不检查）", "int", default=0,
              min_value=0, max_value=10000),
        Param("keep_subdir", "每个文档单独建子文件夹放产物", "bool", default=False),
    ),
)


@register(OFFICE2PDF_OP)
def run_office_to_pdf(ctx: OpContext, files: list[str], params: dict) -> None:
    soffice = dep_get("soffice")
    if not soffice.available:
        raise DependencyMissing("soffice")

    keep_subdir = bool(params.get("keep_subdir"))
    ctx.set_total(len(files))
    ctx.log(f"LibreOffice：{soffice.path}"
            + (f"（v{soffice.version}）" if soffice.version else ""))

    batches = _chunk(files)
    ctx.log(f"共 {len(files)} 个文件，分 {len(batches)} 批转换。")

    for batch in batches:
        ctx.check()
        _convert_batch(ctx, soffice.path, batch, keep_subdir)


def _chunk(files: list[str]) -> list[list[str]]:
    batches: list[list[str]] = []
    cur: list[str] = []
    size = 0
    for f in files:
        if cur and size + len(f) > _MAX_ARGS_CHARS:
            batches.append(cur)
            cur, size = [], 0
        cur.append(f)
        size += len(f) + 3
    if cur:
        batches.append(cur)
    return batches


def _convert_batch(ctx: OpContext, soffice: str, batch: list[str],
                   keep_subdir: bool) -> None:
    profile = tempfile.mkdtemp(prefix="oda_lo_profile_")
    work = tempfile.mkdtemp(prefix="oda_lo_out_")
    # LibreOffice 要求 file:// URI 形式；Windows 需要 file:///C:/... 三斜杠 + 盘符
    norm = os.path.abspath(profile).replace("\\", "/")
    profile_uri = "file:///" + (norm[1:] if norm.startswith("/") else norm) if os.name == "nt" \
        else "file://" + norm

    args = [
        soffice,
        f"-env:UserInstallation={profile_uri}",
        "--headless", "--norestore", "--invisible",
        "--nolockcheck", "--nodefault", "--nofirststartwizard",
        "--convert-to", "pdf",   # 让 LO 按文档类型自动选过滤器
        "--outdir", work,
        *batch,
    ]

    proc = None
    cancelled = False
    try:
        proc = subprocess.Popen(
            args,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            creationflags=_NO_WINDOW,
            encoding="utf-8", errors="replace",
        )
        deadline = time.time() + 300 + 60 * len(batch)
        while True:
            if proc.poll() is not None:
                break
            if ctx.cancelled():
                _kill(proc)
                ctx.log("已取消，正在终止 LibreOffice …", "warn")
                cancelled = True
                break
            if time.time() > deadline:
                _kill(proc)
                ctx.log("LibreOffice 超时（可能是文件过大或弹出了对话框），已终止本批。", "error")
                break
            time.sleep(0.25)

        if cancelled:
            raise Cancelled("任务已被用户取消")

        out_text = ""
        try:
            out_text = (proc.stdout.read() if proc.stdout else "") or ""
        except Exception:
            pass

        produced = {}
        for name in os.listdir(work):
            if name.lower().endswith(".pdf"):
                produced[os.path.splitext(name)[0].lower()] = os.path.join(work, name)

        for f in batch:
            ctx.check()
            stem = os.path.splitext(os.path.basename(f))[0]
            src = produced.get(stem.lower())
            if src is None:
                # 有些版本会把结果命名为 "文件名.pdf" 之外的形式，做一次模糊匹配
                src = _fuzzy_match(produced, stem)
            if src is None:
                ctx.fail(f, "LibreOffice 未生成 PDF（文件可能加密、受保护或格式不受支持）")
                ctx.step(f"转换 {os.path.basename(f)}")
                continue
            try:
                from ..paths import unique_file

                target_dir = unique_dir(ctx.out_dir, sanitize_name(stem)) if keep_subdir else ctx.out_dir
                dst = unique_file(os.path.join(target_dir, sanitize_name(stem) + ".pdf"))
                shutil.move(src, dst)
                ctx.ok(dst, src=f)
            except Exception as exc:
                ctx.fail(f, f"移动产物失败：{exc}")
            ctx.step(f"转换 {os.path.basename(f)}")

        if out_text.strip():
            for line in out_text.strip().splitlines()[-4:]:
                ctx.log(f"  soffice: {line}", "debug")

    finally:
        if proc and proc.poll() is None:
            _kill(proc)
        shutil.rmtree(profile, ignore_errors=True)
        shutil.rmtree(work, ignore_errors=True)


def _fuzzy_match(produced: dict[str, str], stem: str) -> str | None:
    key = stem.lower()
    for k, v in produced.items():
        if k == key or k.startswith(key) or key.startswith(k):
            return v
    return None


def _kill(proc) -> None:
    try:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()
    except Exception:
        pass
