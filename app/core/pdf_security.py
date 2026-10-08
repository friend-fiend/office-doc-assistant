"""PDF 加密 / 解密。

加密算法说明：
  · AES-256 / AES-128 —— 需要 ``cryptography`` 包（pypdf[crypto]），安全性最好；
  · RC4-128 —— 纯 Python 实现，无需额外依赖，兼容性最广但安全性弱。
程序会自动探测：装了 cryptography 就用 AES，否则退到 RC4-128 并提示用户。
"""
from __future__ import annotations

import os

from pypdf import PdfWriter

try:  # pypdf 6.x 在缺少 cryptography 时抛这个
    from pypdf.errors import DependencyError
except Exception:  # pragma: no cover
    class DependencyError(Exception):  # type: ignore[no-redef]
        pass

from ..errors import ParamError
from ..paths import sanitize_name, unique_file
from .pdfutil import open_reader, write_pdf
from .registry import OpContext, OpDef, Param, register


def _has_crypto() -> bool:
    try:
        import cryptography  # noqa: F401

        return True
    except Exception:
        return False


def _permissions(allow_print: bool, allow_copy: bool):
    """构造权限位。优先用 pypdf 6.x 的 UserAccessPermissions 枚举。

    pypdf 6.19 起 permissions_flag 期望枚举类型；直接传 int 仍然可用但语义不清，
    并且枚举名称更可读。拿不到枚举时退化为原始位运算。
    """
    try:
        from pypdf.constants import UserAccessPermissions as P

        perm = P.R2 | P.R3  # 基础保留位
        if allow_print:
            perm |= P.PRINT | P.PRINT_TO_REPRESENTATION
        if allow_copy:
            perm |= P.EXTRACT | P.EXTRACT_TEXT_AND_GRAPHICS
        return perm
    except Exception:
        pass

    flags = 0
    if allow_print:
        flags |= 0b0000000000000100
        flags |= 0b0000000010000000
    if allow_copy:
        flags |= 0b0000000000010000
    flags |= 0b0000000000000001
    return flags


ENCRYPT_OP = OpDef(
    key="pdf_encrypt",
    title="PDF 加密",
    category="compress",
    accepts=("pdf",),
    mode="per_file",
    confirm=True,
    confirm_text="将使用同一密码加密列表中的所有 PDF，请确认已记录该密码（忘记密码将无法恢复）。",
    description="为 PDF 设置打开密码，可选限制打印/复制。",
    params=(
        Param("user_pwd", "打开密码", "password", default="", width=20),
        Param("user_pwd2", "确认密码", "password", default="", width=20),
        Param("owner_pwd", "权限密码（可留空）", "password", default="", width=20,
              hint="留空则与打开密码相同"),
        Param("allow_print", "允许打印", "bool", default=True),
        Param("allow_copy", "允许复制文本", "bool", default=False),
        Param("algorithm", "加密算法", "choice", default="auto",
              choices=(("auto", "自动（优先 AES-128）"), ("AES-256", "AES-256"),
                       ("AES-128", "AES-128"), ("RC4-128", "RC4-128（兼容性最好）"))),
    ),
)


@register(ENCRYPT_OP)
def run_encrypt(ctx: OpContext, files: list[str], params: dict) -> None:
    pwd = str(params.get("user_pwd") or "")
    pwd2 = str(params.get("user_pwd2") or "")
    if not pwd:
        raise ParamError("请填写打开密码")
    if pwd != pwd2:
        raise ParamError("两次输入的密码不一致")
    owner = str(params.get("owner_pwd") or "") or pwd
    allow_print = bool(params.get("allow_print", True))
    allow_copy = bool(params.get("allow_copy", False))
    algo = str(params.get("algorithm") or "auto")

    has_crypto = _has_crypto()
    if algo == "auto":
        algo = "AES-128" if has_crypto else "RC4-128"
    if algo.startswith("AES") and not has_crypto:
        ctx.log("未安装 cryptography，AES 不可用 —— 自动改用 RC4-128。"
                "如需更强加密：pip install pypdf[crypto]", "warn")
        algo = "RC4-128"
    ctx.log(f"加密算法：{algo}")

    flags = _permissions(allow_print, allow_copy)

    ctx.set_total(len(files))
    for f in files:
        ctx.check()
        stem = os.path.splitext(os.path.basename(f))[0]
        try:
            reader = open_reader(f)  # 已加密的输入会抛 PasswordRequired，由上层提示
            w = PdfWriter()
            for page in reader.pages:
                w.add_page(page)
            used_algo = algo
            try:
                w.encrypt(user_password=pwd, owner_password=owner,
                          permissions_flag=flags, algorithm=algo)
            except DependencyError as exc:
                # AES 需要 cryptography>=3.1（pypdf[crypto]），缺失时降级
                ctx.log(f"{algo} 不可用（{exc}），已降级为 RC4-128。"
                        "如需更强加密：pip install pypdf[crypto]", "warn")
                used_algo = "RC4-128"
                w.encrypt(user_password=pwd, owner_password=owner,
                          permissions_flag=flags, algorithm="RC4-128")
            except Exception as exc:
                raise RuntimeError(f"加密失败：{exc}") from exc
            out = unique_file(os.path.join(ctx.out_dir, f"{sanitize_name(stem)}_已加密.pdf"))
            write_pdf(w, out, compress=False)
            w.close()
            ctx.metric("加密算法", used_algo)
            ctx.ok(src=f)
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"加密 {os.path.basename(f)}")


DECRYPT_OP = OpDef(
    key="pdf_decrypt",
    title="PDF 解密",
    category="compress",
    accepts=("pdf",),
    mode="per_file",
    description="去除已知密码的 PDF 打开限制，输出一份未加密副本。密码可在运行中逐个输入。",
    params=(
        Param("password", "统一密码", "password", default="", width=20,
              hint="留空则在处理每个文件时逐个询问密码"),
    ),
)


@register(DECRYPT_OP)
def run_decrypt(ctx: OpContext, files: list[str], params: dict) -> None:
    common = str(params.get("password") or "")
    ctx.set_total(len(files))
    for f in files:
        ctx.check()
        stem = os.path.splitext(os.path.basename(f))[0]
        try:
            reader = open_reader(f, common or None)
            if not reader.is_encrypted:
                ctx.skip(f, "文件未加密，无需解密")
                ctx.step(f"解密 {os.path.basename(f)}")
                continue
            w = PdfWriter()
            for page in reader.pages:
                w.add_page(page)
            out = unique_file(os.path.join(ctx.out_dir, f"{sanitize_name(stem)}_已解密.pdf"))
            write_pdf(w, out, compress=False)
            w.close()
            ctx.ok(src=f)
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"解密 {os.path.basename(f)}")
