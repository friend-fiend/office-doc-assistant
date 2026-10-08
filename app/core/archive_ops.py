"""批量压缩 / 解压：ZIP 与 7z。

安全要点：
  · 解压时做 Zip-Slip 检查，拒绝任何解压后跑到目标目录之外的条目
    （路径穿越攻击，是解压功能最常见的安全漏洞）；
  · 遇到加密压缩包，提示输入密码（7z 支持，ZIP 仅支持 ZipCrypto/AES）。
"""
from __future__ import annotations

import os
import zipfile
from pathlib import Path

from ..errors import ParamError
from ..paths import long_path, sanitize_name, unique_dir, unique_file
from .registry import OpContext, OpDef, Param, register

ARCHIVE_COMPRESS_OP = OpDef(
    key="archive_compress",
    title="批量压缩",
    category="compress",
    accepts=("any",),
    mode="collection",
    description="把列表中的文件/文件夹打包成一个 ZIP 或 7z 压缩包。",
    params=(
        Param("format", "压缩格式", "choice", default="zip",
              choices=(("zip", "ZIP（兼容性最好）"), ("7z", "7z（体积更小）"))),
        Param("out_name", "输出文件名", "entry", default="打包结果", width=22),
        Param("level", "压缩级别", "choice", default="normal",
              choices=(("store", "仅打包（最快）"), ("normal", "普通"),
                       ("max", "最大压缩（最慢）"))),
        Param("password", "压缩密码（可留空）", "password", default="", width=18,
              hint="仅 7z 与标准 ZIP 支持"),
    ),
)

ARCHIVE_EXTRACT_OP = OpDef(
    key="archive_extract",
    title="批量解压",
    category="compress",
    accepts=("archive",),
    mode="per_file",
    description="把每个压缩包解压到独立子文件夹，自动防止路径穿越。",
    params=(
        Param("password", "统一解压密码", "password", default="", width=18,
              hint="留空则在遇到加密包时询问"),
    ),
)


@register(ARCHIVE_COMPRESS_OP)
def run_archive_compress(ctx: OpContext, files: list[str], params: dict) -> None:
    fmt = str(params.get("format") or "zip")
    out_name = sanitize_name(str(params.get("out_name") or "打包结果"))
    level = str(params.get("level") or "normal")
    password = str(params.get("password") or "")

    ctx.set_total(len(files))
    collected = 0
    if fmt == "7z":
        try:
            import py7zr
        except Exception as exc:
            raise ParamError(f"未安装 py7zr，无法生成 7z：{exc}") from exc

        out = unique_file(os.path.join(ctx.out_dir, out_name + ".7z"))
        filters = None
        if hasattr(py7zr, "FILTER_LZMA2"):
            filters = [{"id": py7zr.FILTER_LZMA2,
                        "preset": 9 if level == "max" else (0 if level == "store" else 6)}]
        with py7zr.SevenZipFile(out, "w", password=password or None, filters=filters) as zf:
            for f in files:
                ctx.check()
                try:
                    zf.write(f, arcname=os.path.basename(f))
                    collected += 1
                    ctx.ok(src=f)
                except Exception as exc:
                    ctx.fail(f, str(exc))
                ctx.step(f"压缩 {os.path.basename(f)}")
        ctx.metric("输出文件", out)
        ctx.metric("归档条目", collected)
        ctx.report.outputs.append(out)
        return

    out = unique_file(os.path.join(ctx.out_dir, out_name + ".zip"))
    mode = zipfile.ZIP_STORED if level == "store" else zipfile.ZIP_DEFLATED
    kwargs: dict = {"compression": mode}
    if level == "max":
        kwargs["compresslevel"] = 9
    elif level == "normal":
        kwargs["compresslevel"] = 6

    with zipfile.ZipFile(long_path(out), "w", **kwargs) as zf:
        for f in files:
            ctx.check()
            try:
                p = Path(f)
                if p.is_dir():
                    for root, _dirs, names in os.walk(p):
                        for name in names:
                            ctx.check()
                            full = os.path.join(root, name)
                            arc = os.path.relpath(full, p.parent)
                            zf.write(long_path(full), arc.replace(os.sep, "/"))
                            collected += 1
                else:
                    zf.write(long_path(f), os.path.basename(f))
                    collected += 1
                ctx.ok(src=f)
            except Exception as exc:
                ctx.fail(f, str(exc))
            ctx.step(f"压缩 {os.path.basename(f)}")

    ctx.metric("输出文件", out)
    ctx.metric("归档条目", collected)
    ctx.report.outputs.append(out)


@register(ARCHIVE_EXTRACT_OP)
def run_archive_extract(ctx: OpContext, files: list[str], params: dict) -> None:
    common_pwd = str(params.get("password") or "")
    ctx.set_total(len(files))

    for f in files:
        ctx.check()
        stem = sanitize_name(os.path.splitext(os.path.basename(f))[0])
        try:
            lower = f.lower()
            if lower.endswith(".zip"):
                target = _extract_zip(ctx, f, stem, common_pwd)
            elif lower.endswith(".7z"):
                target = _extract_7z(ctx, f, stem, common_pwd)
            elif lower.endswith((".tar", ".gz", ".tgz", ".bz2", ".xz")):
                target = _extract_tar(ctx, f, stem)
            else:
                ctx.skip(f, "不支持的压缩格式（仅支持 ZIP / 7z / tar.gz）")
                ctx.step(f"解压 {os.path.basename(f)}")
                continue
            if target:
                ctx.report.outputs.append(target)
            ctx.ok(src=f)
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"解压 {os.path.basename(f)}")


def _extract_tar(ctx: OpContext, path: str, stem: str) -> str:
    import tarfile

    target = unique_dir(ctx.out_dir, stem)
    try:
        with tarfile.open(long_path(path)) as tf:
            _safe_tar_extract(tf, target, ctx)
    except ParamError:
        raise
    except Exception as exc:
        raise ParamError(f"tar 解压失败：{exc}") from exc
    return target


def _extract_zip(ctx: OpContext, path: str, stem: str, password: str) -> str:
    target = unique_dir(ctx.out_dir, stem)
    pwd = password or None
    try:
        with zipfile.ZipFile(long_path(path)) as zf:
            encrypted = any(zi.flag_bits & 0x1 for zi in zf.infolist())
            if encrypted and not pwd:
                asked = ctx.ask_password(path)
                if asked is None:
                    raise ParamError("需要密码才能解压，用户已取消")
                pwd = asked
            _validate_zip_entries(zf, target)
            zf.extractall(long_path(target), pwd=pwd.encode("utf-8") if pwd else None)
    except zipfile.BadZipFile as exc:
        raise ParamError(f"压缩包损坏或不是有效的 ZIP：{exc}") from exc
    except RuntimeError as exc:
        raise ParamError(f"解压失败（密码错误？）：{exc}") from exc
    return target


def _validate_zip_entries(zf: zipfile.ZipFile, target: str) -> None:
    """Zip-Slip 防护：任何条目解析后必须仍位于 target 之内。"""
    base = os.path.realpath(target)
    for zi in zf.infolist():
        name = zi.filename.replace("\\", "/")
        if name.startswith("/") or ".." in Path(name).parts:
            raise ParamError(f"压缩包内含危险路径，已拒绝解压：{zi.filename}")
        resolved = os.path.realpath(os.path.join(base, *Path(name).parts))
        if not (resolved == base or resolved.startswith(base + os.sep)):
            raise ParamError(f"压缩包内含危险路径，已拒绝解压：{zi.filename}")


def _extract_7z(ctx: OpContext, path: str, stem: str, password: str) -> str:
    import py7zr

    target = unique_dir(ctx.out_dir, stem)
    pwd = password or None
    try:
        try:
            with py7zr.SevenZipFile(path, "r", password=pwd) as zf:
                _validate_7z_names(zf, target)
                zf.extractall(path=target)
        except py7zr.exceptions.PasswordRequired:
            asked = ctx.ask_password(path)
            if asked is None:
                raise ParamError("需要密码才能解压，用户已取消")
            with py7zr.SevenZipFile(path, "r", password=asked) as zf:
                _validate_7z_names(zf, target)
                zf.extractall(path=target)
    except ParamError:
        raise
    except Exception as exc:
        raise ParamError(f"7z 解压失败：{exc}") from exc
    return target


def _validate_7z_names(zf, target: str) -> None:
    base = os.path.realpath(target)
    try:
        names = zf.getnames()
    except Exception:
        return
    for name in names:
        clean = name.replace("\\", "/")
        if clean.startswith("/") or ".." in Path(clean).parts:
            raise ParamError(f"压缩包内含危险路径，已拒绝解压：{name}")
        resolved = os.path.realpath(os.path.join(base, *Path(clean).parts))
        if not (resolved == base or resolved.startswith(base + os.sep)):
            raise ParamError(f"压缩包内含危险路径，已拒绝解压：{name}")


def _safe_tar_extract(tf, target: str, ctx: OpContext) -> None:
    """tar 的 Python 3.12+ 已有 filter 参数，这里对旧版本做等价校验。"""
    base = os.path.realpath(target)
    for member in tf.getmembers():
        ctx.check()
        resolved = os.path.realpath(os.path.join(base, member.name))
        if not (resolved == base or resolved.startswith(base + os.sep)):
            raise ParamError(f"压缩包内含危险路径，已拒绝解压：{member.name}")
        if member.issym() or member.islnk():
            raise ParamError(f"压缩包内含链接条目，出于安全考虑已拒绝：{member.name}")
    try:
        tf.extractall(target, filter="data")  # Python 3.12+
    except TypeError:
        tf.extractall(target)
