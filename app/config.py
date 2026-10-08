"""全局常量：应用信息、扩展名集合、默认参数值。

所有模块统一从这里取值，避免魔法字符串散落各处。
"""
from __future__ import annotations

import sys
from pathlib import Path

# ---------------------------------------------------------------- 应用信息
APP_NAME = "办公文档一键处理助手"
APP_NAME_EN = "OfficeDocAssistant"
APP_ID = "OfficeDocAssistant"
VERSION = "1.0.2"
ORG_NAME = "OfficeDocAssistant"

IS_WINDOWS = sys.platform.startswith("win")
IS_MACOS = sys.platform == "darwin"
IS_LINUX = (not IS_WINDOWS) and (not IS_MACOS)

# ---------------------------------------------------------------- 扩展名
PDF_EXT = {".pdf"}
WORD_EXT = {".doc", ".docx", ".docm", ".dot", ".dotx", ".rtf", ".odt", ".wps"}
EXCEL_EXT = {".xls", ".xlsx", ".xlsm", ".xlt", ".xltx", ".csv", ".ods", ".et"}
PPT_EXT = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".pot", ".odp", ".dps"}
OFFICE_EXT = WORD_EXT | EXCEL_EXT | PPT_EXT
IMAGE_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
ARCHIVE_EXT = {".zip", ".7z", ".tar", ".gz", ".bz2", ".xz", ".tgz"}
TEXT_DOC_EXT = {".txt", ".md"}

# 扩展名集合 -> 显示名（用于校验提示）
EXT_GROUP_LABEL = {
    "pdf": "PDF 文件",
    "office": "Office 文档",
    "document": "Word / 文本",
    "image": "图片",
    "archive": "压缩包",
    "any": "任意文件",
}


def ext_set(group: str) -> set[str]:
    """把扩展名分组名解析成扩展名集合。"""
    return {
        "pdf": PDF_EXT,
        "office": OFFICE_EXT,
        "word": WORD_EXT,
        "excel": EXCEL_EXT,
        "ppt": PPT_EXT,
        "image": IMAGE_EXT,
        "archive": ARCHIVE_EXT,
        "document": WORD_EXT | PDF_EXT | TEXT_DOC_EXT,
        "any": set(),
    }.get(group, set())


# ---------------------------------------------------------------- 输出
OUTPUT_SUFFIX = "_输出"
TEMP_PREFIX = "oda_tmp_"

# ---------------------------------------------------------------- 压缩档位
COMPRESS_LEVELS = (
    ("light", "轻度（无损优先）", "仅结构优化、对象流重写、冗余清理"),
    ("medium", "中度（推荐）", "Ghostscript 重建，图片降采样 150 DPI，JPEG 质量 80"),
    ("strong", "强力（体积最小）", "Ghostscript 重建，图片降采样 96 DPI，JPEG 质量 60"),
)

# 档位 -> Ghostscript 参数
GS_PRESETS = {
    "medium": {"dpi": 150, "jpeg_q": 80, "color_dpi": 150, "gray_dpi": 150, "mono_dpi": 300},
    "strong": {"dpi": 96, "jpeg_q": 60, "color_dpi": 96, "gray_dpi": 96, "mono_dpi": 200},
}

# ---------------------------------------------------------------- 页码/水印
PAGE_NUM_FORMATS = (
    ("{n}", "1"),
    ("第 {n} 页", "第 1 页"),
    ("- {n} -", "- 1 -"),
    ("{n} / {total}", "1 / 10"),
    ("第 {n} 页 / 共 {total} 页", "第 1 页 / 共 10 页"),
    ("Page {n}", "Page 1"),
)

POSITIONS = (
    ("bottom-center", "底部居中"),
    ("bottom-right", "右下角"),
    ("bottom-left", "左下角"),
    ("top-center", "顶部居中"),
    ("top-right", "右上角"),
    ("top-left", "左上角"),
)

# ---------------------------------------------------------------- 打包资源定位
def resource_dir() -> Path:
    """返回只读资源目录（源码运行 / PyInstaller 冻结运行都适用）。"""
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", None)
        if base:
            return Path(base)
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent


def app_root() -> Path:
    """返回应用根目录（可执行文件所在目录）。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent
