"""Excel 合并与拆分（openpyxl）。

合并支持两种语义：
  · 每个文件一个工作表 —— 保留原始结构，工作簿之间互不干扰；
  · 追加到同一工作表 —— 适合多个结构相同的月报/明细表纵向拼接。
拆分则把一个工作簿的每个工作表导出成独立的 .xlsx。

拆分的样式策略（默认开启 ``preserve_format=True``）：
  · 公式真公式（``Cell.value = '=SUM(A1:A10)'``），不是字符串
  · 字体 / 填充 / 边框 / 对齐 / 数字格式
  · 列宽 / 行高
  · 合并单元格
  · 冻结窗格 / 打印设置（页面方向 / 纸张 / 缩放 / 打印标题）

注意：openpyxl 不能把 A 表的 ``Cell`` 对象 append 到 B 表，所以拆分的实现走
``new_ws.cell(...)`` 写入新单元格 + ``copy.copy(cell)`` 复制样式。这是
"保真拆分"的常用做法，速度比 ``iter_rows(values_only=True)`` 慢，但正确性
明显更好。
"""
from __future__ import annotations

import copy
import os

from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.utils import get_column_letter

from ..errors import ParamError
from ..paths import sanitize_name, unique_file
from .registry import OpContext, OpDef, Param, register

MAX_ROWS_PER_SHEET = 1_048_576


def _safe_sheet_title(title: str, used: set[str]) -> str:
    """Excel 工作表名限制 31 字符且不能含 : \\ / ? * [ ]。"""
    bad = ':\\/?*[]'
    clean = "".join("_" if ch in bad else ch for ch in str(title)).strip() or "Sheet"
    clean = clean[:31]
    base = clean
    idx = 2
    while clean.lower() in used:
        suffix = f"_{idx}"
        clean = base[: 31 - len(suffix)] + suffix
        idx += 1
    used.add(clean.lower())
    return clean


def _copy_sheet(src_ws, dst_ws) -> None:
    """把 ``src_ws`` 的数据、公式、样式、列宽行高、合并、冻结等复制到 ``dst_ws``。

    设计要点：
      1. 数据 / 公式：用 ``new_ws.cell(row, col, value=...)`` 写入（带坐标），
         因为 append() 不会保留公式和样式属性。
      2. 样式：``copy.copy(cell)`` 仅浅拷贝样式属性对象，跨工作表赋值是允许的。
      3. 行高 / 列宽：openpyxl 把它们存在 ``row_dimensions`` / ``column_dimensions``。
      4. 合并单元格：循环 ``ws.merged_cells.ranges``，目标表里重新 ``merge_cells``。
    """
    # ---------------- 数据 + 公式 + 样式
    for row in src_ws.iter_rows():
        for cell in row:
            if cell.value is None and cell.has_style is False:
                continue
            new_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy.copy(cell.font)
                new_cell.fill = copy.copy(cell.fill)
                new_cell.border = copy.copy(cell.border)
                new_cell.alignment = copy.copy(cell.alignment)
                new_cell.number_format = cell.number_format
                new_cell.protection = copy.copy(cell.protection)
            # 数据类型（formula 直接写 value 即可，openpyxl 已识别）

    # ---------------- 列宽
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width is not None:
            dst_ws.column_dimensions[col_letter].width = dim.width
        if dim.hidden:
            dst_ws.column_dimensions[col_letter].hidden = True
        if dim.outlineLevel:
            dst_ws.column_dimensions[col_letter].outlineLevel = dim.outlineLevel

    # ---------------- 行高（只在用户显式设过时才拷，避免给整列填无意义的默认值）
    for row_idx, dim in src_ws.row_dimensions.items():
        if dim.height is not None:
            dst_ws.row_dimensions[row_idx].height = dim.height
        if dim.hidden:
            dst_ws.row_dimensions[row_idx].hidden = True
        if dim.outlineLevel:
            dst_ws.row_dimensions[row_idx].outlineLevel = dim.outlineLevel

    # ---------------- 合并单元格
    for rng in list(src_ws.merged_cells.ranges):
        dst_ws.merge_cells(str(rng))

    # ---------------- 冻结窗格
    if src_ws.freeze_panes:
        dst_ws.freeze_panes = src_ws.freeze_panes

    # ---------------- 打印设置
    ps = src_ws.page_setup
    if any([ps.orientation, ps.paperSize, ps.fitToWidth, ps.fitToHeight, ps.scale]):
        dst_ws.page_setup.orientation = ps.orientation
        dst_ws.page_setup.paperSize = ps.paperSize
        dst_ws.page_setup.fitToWidth = ps.fitToWidth
        dst_ws.page_setup.fitToHeight = ps.fitToHeight
        if ps.scale:
            dst_ws.page_setup.scale = ps.scale

    pm = src_ws.page_margins
    if any([pm.left, pm.right, pm.top, pm.bottom, pm.header, pm.footer]):
        dst_ws.page_margins.left = pm.left
        dst_ws.page_margins.right = pm.right
        dst_ws.page_margins.top = pm.top
        dst_ws.page_margins.bottom = pm.bottom
        dst_ws.page_margins.header = pm.header
        dst_ws.page_margins.footer = pm.footer

    if src_ws.print_options.horizontalCentered or src_ws.print_options.verticalCentered:
        dst_ws.print_options.horizontalCentered = src_ws.print_options.horizontalCentered
        dst_ws.print_options.verticalCentered = src_ws.print_options.verticalCentered

    if src_ws.print_title_rows:
        dst_ws.print_title_rows = src_ws.print_title_rows
    if src_ws.print_title_cols:
        dst_ws.print_title_cols = src_ws.print_title_cols


EXCEL_MERGE_OP = OpDef(
    key="excel_merge",
    title="Excel 合并",
    category="batch",
    accepts=("excel",),
    mode="collection",
    description="把多个工作簿合并成一个 .xlsx，可选每文件一个表或全部追加到同一张表。",
    params=(
        Param("merge_mode", "合并方式", "choice", default="per_sheet",
              choices=(("per_sheet", "每个文件一个工作表（保留结构）"),
                       ("append", "全部追加到同一工作表（表头只保留一次）"))),
        Param("sheet_name", "追加模式的目标表名", "entry", default="汇总", width=18,
              visible_if=("merge_mode", ("append",))),
        Param("values_only", "只保留数值与文本（丢弃公式）", "bool", default=False),
        Param("skip_hidden_sheets", "跳过隐藏工作表", "bool", default=True),
    ),
)


@register(EXCEL_MERGE_OP)
def run_excel_merge(ctx: OpContext, files: list[str], params: dict) -> None:
    mode = str(params.get("merge_mode") or "per_sheet")
    values_only = bool(params.get("values_only"))
    skip_hidden = bool(params.get("skip_hidden_sheets", True))

    out_wb = Workbook()
    out_wb.remove(out_wb.active)  # 去掉默认空白表
    used: set[str] = set()
    ctx.set_total(len(files))

    append_ws = None
    header_written = False
    total_rows = 0

    for f in files:
        ctx.check()
        try:
            if f.lower().endswith(".csv"):
                ctx.skip(f, "CSV 请先用 Excel 打开另存为 .xlsx 后再合并")
                ctx.step(f"合并 {os.path.basename(f)}")
                continue

            wb = load_workbook(f, data_only=values_only, read_only=False)
            src_rows = 0

            if mode == "append":
                if append_ws is None:
                    append_ws = out_wb.create_sheet(
                        _safe_sheet_title(str(params.get("sheet_name") or "汇总"), used))
                for ws in wb.worksheets:
                    if skip_hidden and ws.sheet_state != "visible":
                        continue
                    # 必须用 values_only=True 取值：openpyxl 不允许把 A 表的 Cell
                    # 对象 append 到 B 表（会抛 "Cells cannot be copied from
                    # other worksheets"）。values_only 下公式单元格返回公式字符串，
                    # 数值单元格返回计算结果，语义正好符合需求。
                    for r_i, row in enumerate(ws.iter_rows(values_only=True), start=1):
                        ctx.check()
                        if r_i == 1 and header_written:
                            continue  # 后续文件跳过表头
                        append_ws.append(list(row))
                        src_rows += 1
                header_written = True
            else:
                for ws in wb.worksheets:
                    if skip_hidden and ws.sheet_state != "visible":
                        continue
                    title = _safe_sheet_title(f"{os.path.splitext(os.path.basename(f))[0]}-{ws.title}", used)
                    new_ws = out_wb.create_sheet(title)
                    for row in ws.iter_rows(values_only=True):
                        ctx.check()
                        new_ws.append(list(row))
                        src_rows += 1

            wb.close()
            total_rows += src_rows
            ctx.log(f"  {os.path.basename(f)}：{src_rows} 行")
            ctx.ok(src=f)
        except InvalidFileException as exc:
            ctx.fail(f, f"不是有效的 Excel 文件：{exc}")
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"合并 {os.path.basename(f)}")

    if not out_wb.sheetnames:
        raise ParamError("没有任何可合并的工作表，请检查文件格式")

    sheet_count = len(out_wb.sheetnames)
    out = unique_file(os.path.join(ctx.out_dir, "Excel合并结果.xlsx"))
    out_wb.save(out)
    out_wb.close()
    ctx.metric("工作表数", sheet_count)
    ctx.metric("总行数", total_rows)
    ctx.metric("输出文件", out)
    ctx.report.outputs.append(out)


EXCEL_SPLIT_OP = OpDef(
    key="excel_split",
    title="Excel 拆分",
    category="batch",
    accepts=("excel",),
    mode="per_file",
    description=(
        "把工作簿中的每个工作表导出为独立的 .xlsx 文件。"
        "默认保留公式（真公式而非字符串）、字体/填充/边框/对齐/数字格式、列宽行高、合并单元格、冻结窗格与打印设置。"
    ),
    params=(
        Param("each_subfolder", "每个工作簿单独建子文件夹", "bool", default=True),
        Param("preserve_format", "保留格式与公式（推荐）", "bool", default=True,
              hint="关闭后只导出数据，速度更快；推荐保持开启以获得真公式与样式。"),
        Param("skip_hidden_sheets", "跳过隐藏工作表", "bool", default=True),
    ),
)


@register(EXCEL_SPLIT_OP)
def run_excel_split(ctx: OpContext, files: list[str], params: dict) -> None:
    from ..paths import unique_dir

    sub = bool(params.get("each_subfolder", True))
    preserve = bool(params.get("preserve_format", True))   # 默认开启
    skip_hidden = bool(params.get("skip_hidden_sheets", True))
    ctx.set_total(len(files))

    for f in files:
        ctx.check()
        stem = sanitize_name(os.path.splitext(os.path.basename(f))[0])
        try:
            # data_only=False 才能拿到公式与样式；True 只会拿计算后的值（且丢失公式）
            wb = load_workbook(f, data_only=False)
            target = unique_dir(ctx.out_dir, stem) if sub else ctx.out_dir
            made = 0
            for ws in wb.worksheets:
                ctx.check()
                if skip_hidden and ws.sheet_state != "visible":
                    ctx.log(f"  跳过隐藏工作表：{ws.title}", "warn")
                    continue

                new_wb = Workbook()
                new_ws = new_wb.active
                new_ws.title = _safe_sheet_title(ws.title, set())

                if preserve:
                    _copy_sheet(ws, new_ws)
                else:
                    # 快速路径：仅数据，无公式无样式
                    for row in ws.iter_rows(values_only=True):
                        ctx.check()
                        new_ws.append(list(row))

                out = unique_file(os.path.join(target, f"{stem}_{sanitize_name(ws.title)}.xlsx"))
                new_wb.save(out)
                new_wb.close()
                made += 1
            wb.close()
            if sub:
                ctx.report.outputs.append(target)
            ctx.log(
                f"  {os.path.basename(f)}：拆出 {made} 个工作表"
                + ("（含格式与公式）" if preserve else "（仅数据）")
            )
            ctx.ok(src=f)
        except Exception as exc:
            ctx.fail(f, str(exc))
        ctx.step(f"拆分 {os.path.basename(f)}")