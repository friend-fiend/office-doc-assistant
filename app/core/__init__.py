"""核心业务包。

导入本包即完成所有操作的注册（各模块通过 @register 自注册）。
"""
from __future__ import annotations

from .registry import (
    OPS,
    CATEGORIES,
    OpContext,
    OpDef,
    Param,
    categories,
    get_op,
    ops_of,
    register,
)

# 触发各操作模块的自注册（顺序无所谓，导入即注册）
from . import (  # noqa: E402,F401
    archive_ops,
    compare_ops,
    excel_ops,
    image_to_pdf,
    invoice_ops,
    office_to_pdf,
    pdf_compress,
    pdf_pages,
    pdf_security,
    pdf_stamp,
    pdf_to_image,
    print_ops,
    rename_ops,
)

__all__ = [
    "OPS",
    "CATEGORIES",
    "OpContext",
    "OpDef",
    "Param",
    "categories",
    "get_op",
    "ops_of",
    "register",
]
