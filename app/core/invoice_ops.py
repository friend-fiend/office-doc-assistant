"""发票整理：把一批发票 PDF 解析成按项目归集的 Excel 汇总表。

设计要点（基于 2026-11 用 48 张真实样票的实测）：

1. 零 OCR。实测这批票全部有文本层：
   · 铁路电子客票 —— 出发站/车次/到达站/乘车日期/票价都在票面；
   · 数电普票（的士/网约车）—— 有专门的「出行人/出行日期/出发地/到达地」栏；
   · 数电专票（住宿）—— 酒店名就是销售方名称；
   · 数电普票（餐饮）—— 餐厅名就是销售方名称。
   扫描件/照片（无文本层）不猜，进「待确认」。

2. 提取用 PyMuPDF 的 ``get_text("text", sort=True)``（按坐标排序的阅读顺序）。
   默认的 ``get_text("text")`` 顺序是内容流顺序，铁路票会碎成单字，不能用。

3. 「项目归属」= 文件夹结构，不让用户命名：
   · 用户填了项目名 → 全部归该项目；
   · 所有文件同在一个目录 → 项目名 = 该目录名；
   · 文件分布在公共父目录的多个子目录 → 项目名 = 第 1 层子目录名；
   · 散在公共父目录根下的 → 「未归属」，绝不猜。

4. 一张票可以含多段行程（实测红冲网约车票 6 段）。明细行按「票」组织，
   多段行程用「；」连接写进行程列 —— 与用户现有报销表的人工习惯一致，
   也保证金额列按票求和不重不漏。

5. 铁路票有「乘车日期」和「开票日期」两个日期，报销口径用乘车日期；
   住宿/餐饮票面没有消费日期，只有开票日期，说明里注明。
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from ..paths import sanitize_name, unique_file
from .registry import OpContext, OpDef, Param, register

# ================================================================= 常量
RE_INVOICE_NO = re.compile(r"发票号码\s*[:：]\s*([0-9]{20})")
RE_INVOICE_NO_BARE = re.compile(r"(?<![0-9])([0-9]{20})(?![0-9])")
RE_INV_DATE = re.compile(r"开票日期\s*[:：]\s*(\d{4}年\d{2}月\d{2}日)")
RE_RIDE_DATE = re.compile(r"(\d{4}年\d{2}月\d{2}日)\s+(\d{2}:\d{2})开")
RE_STATION_LINE = re.compile(
    r"^\s*([\u4e00-\u9fa5A-Za-z]{2,14}站)\s+([DGKCZT]\d+)\s+([\u4e00-\u9fa5A-Za-z]{2,14}站)\s*$",
    re.M,
)
RE_RAIL_FARE = re.compile(r"票价\s*[:：]\s*[￥¥]\s*(-?[\d,]+\.?\d*)")
RE_PASSENGER = re.compile(r"(\d{6}\*+\d{4,8})\s+([\u4e00-\u9fa5]{2,4})")
RE_ITEM_LINE = re.compile(r"^\s*(\*\S+?\*\S+?)\s{2,}", re.M)
RE_TAX_TOTAL = re.compile(r"[（(]小写[）)]\s*[￥¥]?\s*(-?[\d,]+(?:\.\d{1,2})?)")
RE_MONEY = re.compile(r"[￥¥]\s*(-?[\d,]+\.\d{2})")
RE_ID_TAXNO = re.compile(r"识别号\s*[:：]\s*([0-9A-Z]{15,20})")
# 注意 lookahead 用 re.M 的 $ 而不是 \Z：公司名常在行尾，后面是 \n+下一行
# 的非空白字符，\s{2,} 和 \Z 都不成立，会导致第二个名称整体匹配失败。
RE_NAME = re.compile(r"名称\s*[:：]\s*(.+?)(?=\s{2,}|$)", re.M)
RE_TRIP_HEADER = re.compile(r"出行人\s+有效身份证件号\s+出行日期\s+出发地\s+到达地")
RE_FILENAME_MONEY = re.compile(r"-(\d+(?:\.\d+)?)")
# 出行信息行（数电普票的旅客运输栏）。实测 PyMuPDF sort 模式把每段行程
# 规整成一行、列间 ≥2 空格，直接按行正则提取即可，无需 words 坐标分列。
RE_TRIP_LINE = re.compile(
    r"^\s*([\u4e00-\u9fa5]{2,4})\s{2,}(\d{6}\*+\d{4,8})\s{2,}(\d{4}-\d{2}-\d{2})\s+"
    r"(.+?)\s{2,}(.+?)\s{2,}(\S+)\s{2,}(\S+)\s*$",
    re.M,
)

HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
HEAD_FONT = Font(bold=True)
WRAP = Alignment(wrap_text=True, vertical="top")

# 明细表列
DET_COLS = (
    "项目", "费用类别", "类型", "日期", "行程（出发→到达）", "酒店/餐厅",
    "销售方", "发票号码", "金额", "税额", "价税合计", "出行人", "文件名", "说明",
)
SUM_COLS = ("项目", "出行-张数", "出行-金额", "住宿-张数", "住宿-金额",
            "餐饮-张数", "餐饮-金额", "价税合计")
PEND_COLS = ("文件名", "归属项目", "原因", "已读到的信息")


# ================================================================= 数据结构
@dataclass
class Trip:
    date: str = ""
    origin: str = ""
    dest: str = ""
    seat: str = ""
    vehicle: str = ""


@dataclass
class Invoice:
    """一张票的解析结果（票级）。"""
    path: str
    project: str
    kind: str = ""          # 出行 / 住宿 / 餐饮 / 其他
    subtype: str = ""       # 动车 / 出租车·网约车 / 酒店 / 餐厅 / …
    invoice_no: str = ""
    invoice_date: str = ""  # 开票日期
    use_date: str = ""      # 乘车/消费日期（住宿餐饮常为空）
    amount: float | None = None
    tax: float | None = None
    total: float | None = None
    seller: str = ""
    buyer: str = ""
    passenger: str = ""
    trips: list[Trip] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)   # 说明（含待确认原因）

    @property
    def route(self) -> str:
        return "；".join(
            f"{t.origin}→{t.dest}" if t.origin or t.dest else (t.seat or "")
            for t in self.trips)

    def summary(self) -> str:
        bits = []
        if self.invoice_no:
            bits.append(f"号码 {self.invoice_no}")
        if self.total is not None:
            bits.append(f"价税合计 {self.total:.2f}")
        if self.seller:
            bits.append(f"销售方 {self.seller}")
        return "，".join(bits)


@dataclass
class Unreadable:
    path: str
    project: str
    reason: str
    seen: str = ""


def _f(s: str | None) -> float | None:
    if not s:
        return None
    try:
        return float(str(s).replace(",", ""))
    except ValueError:
        return None


# ================================================================= 项目归属
def assign_projects(files: list[str], explicit: str = "", level: int = 1) -> dict[str, str]:
    """把每个文件归到一个项目。规则见模块 docstring 第 3 条。"""
    if explicit:
        return {f: explicit for f in files}
    dirs = [os.path.dirname(os.path.abspath(f)) for f in files]
    out: dict[str, str] = {}
    if len(set(dirs)) == 1:
        name = os.path.basename(dirs[0].rstrip("\\/")) or "未归属"
        return {f: name for f in files}
    try:
        common = os.path.commonpath(dirs)
    except ValueError:  # 不同盘符
        return {f: "未归属" for f in files}
    for f, d in zip(files, dirs):
        rel = os.path.relpath(d, common)
        parts = [] if rel == "." else rel.split(os.sep)
        out[f] = parts[level - 1] if len(parts) >= level else "未归属"
    return out


# ================================================================= 解析
def _rail_ticket(text: str, inv: Invoice) -> None:
    inv.kind = "出行"
    inv.subtype = "动车"
    m = RE_STATION_LINE.search(text)
    if m:
        t = Trip(origin=m.group(1), dest=m.group(3), vehicle=f"列车 {m.group(2)}")
        dm = RE_RIDE_DATE.search(text)
        if dm:
            t.date = dm.group(1)
            inv.use_date = dm.group(1)
        seat = re.search(r"\d+车\d+[A-Z]号\s+(\S+)", text)
        if seat:
            t.seat = seat.group(1)
        inv.trips.append(t)
    else:
        inv.notes.append("未读到出发站/到达站行（请人工核对）")
    fm = RE_RAIL_FARE.search(text)
    if fm:
        inv.total = _f(fm.group(1))
        inv.amount = inv.total  # 铁路票票面不显示税额，金额即票价
        inv.notes.append("铁路票票面无税额列；进项税额按 9% 计算抵扣")
    pm = RE_PASSENGER.search(text)
    if pm:
        inv.passenger = pm.group(2)


def _seller_and_buyer(text: str, inv: Invoice) -> None:
    names = RE_NAME.findall(text)
    taxnos = RE_ID_TAXNO.findall(text)
    if len(names) >= 2:
        inv.buyer, inv.seller = names[0].strip(), names[1].strip()
    elif len(names) == 1:
        inv.buyer = names[0].strip()
        inv.notes.append("只读到购买方名称，销售方缺失")
    if len(taxnos) >= 2:
        inv.notes.append(f"税号 {taxnos[1]}")
    m = re.search(r"购买方名称[:：]\s*(\S+)", text)
    if m and not inv.buyer:
        inv.buyer = m.group(1)


def _trip_rows(text: str, inv: Invoice) -> None:
    """从「出行人/出发地/到达地」表格抽出行段。

    实测（48 张样票）：PyMuPDF sort 模式把每段行程规整成一行、列间 ≥2 空格，
    直接按行正则提取即可，不需要 words 坐标分列 —— 几何法反而会被
    「横跨日期列空间的长地址」击败。
    """
    rows = RE_TRIP_LINE.findall(text)
    for m in rows:
        t = Trip(
            date=m[2],
            origin=m[3].strip(),
            dest=m[4].strip(),
            seat=m[5].strip(),
            vehicle=m[6].strip(),
        )
        if not inv.passenger:
            inv.passenger = m[0].strip()
        inv.trips.append(t)
    if inv.trips:
        inv.use_date = inv.trips[0].date


def _general_ticket(page, text: str, inv: Invoice) -> None:
    """数电普票/专票：的士、住宿、餐饮（铁路票走 _rail_ticket）。"""
    item = RE_ITEM_LINE.search(text)
    item_name = item.group(1) if item else ""
    if "住宿" in item_name:
        inv.kind, inv.subtype = "住宿", "酒店"
    elif "餐饮" in item_name:
        inv.kind, inv.subtype = "餐饮", "餐厅"
    elif "交通运输" in item_name or "客运" in item_name:
        inv.kind, inv.subtype = "出行", "出租车·网约车"
    else:
        inv.kind, inv.subtype = "其他", item_name or "未知"
        inv.notes.append(f"未识别的项目名称：{item_name or '（空）'}，请人工确认")

    _seller_and_buyer(text, inv)
    if inv.kind == "住宿" and inv.seller:
        inv.notes.append(f"酒店名取自销售方：{inv.seller}")
        inv.subtype = "酒店"

    # 行 1：金额行（项目行里的 金额/税率/税额；合计行兜底）
    amounts = re.findall(r"[¥￥]\s*(-?[\d,]+\.\d{2})", text)
    rate = re.search(r"\*{0,1}[\u4e00-\u9fa5A-Za-z*]+?\s{2,}[\d.]+\s+\d+\s+([\d,]+\.\d{2})\s+(\d+%)\s+([\d,]+\.\d{2})", text)
    if rate:
        inv.amount, inv.tax = _f(rate.group(1)), _f(rate.group(3))
    if len(amounts) >= 2:
        inv.amount = _f(amounts[0])
        inv.tax = _f(amounts[1])
    tm = RE_TAX_TOTAL.search(text)
    if tm:
        inv.total = _f(tm.group(1))
    elif amounts:
        inv.total = _f(amounts[-1])

    _trip_rows(text, inv)


def parse_invoice(path: str, project: str) -> Invoice | Unreadable:
    """解析单张发票 PDF。只读；解析不出关键字段时返回 Unreadable 而不是猜。"""
    try:
        import pymupdf
    except ImportError:
        return Unreadable(path, project, "PyMuPDF 不可用")
    try:
        doc = pymupdf.open(path)
    except Exception as exc:
        return Unreadable(path, project, f"无法打开：{exc}")
    try:
        page = doc[0]
        text = page.get_text("text", sort=True)
        if not text.strip():
            return Unreadable(path, project, "PDF 无文本层（可能是扫描件），本功能不做 OCR")
        inv = Invoice(path=path, project=project)
        m = RE_INVOICE_NO.search(text) or RE_INVOICE_NO_BARE.search(text)
        inv.invoice_no = m.group(1) if m else ""
        dm = RE_INV_DATE.search(text)
        inv.invoice_date = dm.group(1) if dm else ""

        if "铁路电子客票" in text:
            _rail_ticket(text, inv)
        else:
            _general_ticket(page, text, inv)

        if not inv.invoice_no:
            inv.notes.append("未读到 20 位发票号码")
        if inv.total is None:
            inv.notes.append("未读到价税合计")
        # 文件名带负号但票面是正数 —— 红冲票符号疑点，必须人工确认
        base = os.path.basename(path)
        if base.count("-") and inv.total is not None and inv.total > 0:
            neg = re.search(r"-(\d+(?:\.\d+)?)", base)
            if neg and abs(_f(neg.group(1)) or 0) == inv.total:
                inv.notes.insert(0, f"文件名标注为 -{inv.total:.2f}（红冲？）但票面金额为正，请人工确认")
        return inv
    finally:
        doc.close()


# ================================================================= 确认文案
def _describe(files: list, params: dict) -> str:
    try:
        proj = assign_projects(files, str(params.get("project_name") or ""),
                               int(params.get("project_level") or 1))
    except Exception:
        proj = {f: "?" for f in files}
    buckets: dict[str, int] = {}
    for p in proj.values():
        buckets[p] = buckets.get(p, 0) + 1
    lines = [f"即将解析 {len(files)} 张发票并生成汇总 Excel："]
    for name, n in sorted(buckets.items()):
        lines.append(f"  · {name}：{n} 张")
    lines.append("输出：一个 xlsx（发票明细 / 项目汇总 / 待确认）")
    lines.append("解析是只读操作，不会修改发票文件。")
    return "\n".join(lines)


INVOICE_SUMMARY_OP = OpDef(
    key="invoice_summary",
    title="发票整理成表格",
    category="invoice",
    accepts=("pdf",),
    mode="collection",
    confirm=True,
    confirm_builder=_describe,
    output_kind="file",
    description=(
        "把一批发票 PDF 解析成按项目归集的 Excel：出行票自动取出发地→到达地，"
        "住宿票取酒店名称（销售方），餐饮票取餐厅名称；同一张发票多段行程合并为一行。\n"
        "项目归属按文件夹结构判定（一个项目一个文件夹），无需给发票命名。\n"
        "只支持有文本层的电子发票 PDF；扫描件/照片不做 OCR，会列入「待确认」。\n"
        "全程只读，不修改发票文件。"
    ),
    params=(
        Param("project_name", "项目名（可选）", "entry", default="", width=20,
              hint=("留空 = 按文件夹自动归属：同在一个目录时项目名取目录名，"
                    "分布在多个子目录时取第 1 层子目录名。填写后所有票归这一个项目。")),
        Param("project_level", "项目名取第几层文件夹", "choice", default="1",
              choices=(("1", "第 1 层子目录"), ("2", "第 2 层子目录")),
              visible_if=("project_name", ("",))),
        Param("dedupe", "按发票号码去重", "bool", default=True,
              hint="同一张发票被保存了两份时只计一次，重复的列入「待确认」。"),
        Param("check_amount", "金额自洽校验", "bool", default=True,
              hint="校验 金额+税额 ≈ 价税合计（误差 ≤ 0.02 元），不符的列入「待确认」。"),
        Param("split_by_project", "每个项目单独一张工作表", "bool", default=False,
              hint="默认所有票在一张「发票明细」表里（首列为项目，可筛选）。"
                   "勾选后另为每个项目生成独立工作表。"),
    ),
)


@register(INVOICE_SUMMARY_OP)
def run_invoice_summary(ctx: OpContext, files: list[str], params: dict) -> None:
    project_name = str(params.get("project_name") or "").strip()
    try:
        level = int(params.get("project_level") or 1)
    except (TypeError, ValueError):
        level = 1
    dedupe = bool(params.get("dedupe", True))
    check_amount = bool(params.get("check_amount", True))
    split_by_project = bool(params.get("split_by_project"))

    projects = assign_projects(files, project_name, level)
    ctx.set_total(len(files))
    ctx.metric("模式", "发票整理（只读解析）")

    invoices: list[Invoice] = []
    pending: list[Unreadable | Invoice] = []
    seen_no: dict[str, Invoice] = {}

    for f in files:
        ctx.check()
        result = parse_invoice(f, projects.get(f, "未归属"))
        ctx.step(f"解析 {os.path.basename(f)}")
        if isinstance(result, Unreadable):
            pending.append(result)
            ctx.fail(f, result.reason)
            continue
        if dedupe and result.invoice_no:
            first = seen_no.get(result.invoice_no)
            if first is not None:
                result.notes.insert(0, f"发票号码与「{os.path.basename(first.path)}」重复，未计入汇总")
                pending.append(result)
                ctx.skip(f, "发票号码重复")
                continue
            seen_no[result.invoice_no] = result
        if check_amount and result.total is not None:
            parts = [v for v in (result.amount, result.tax) if v is not None]
            if parts and abs(sum(parts) - result.total) > 0.02:
                result.notes.append(
                    f"金额自洽校验未通过：{result.amount}+{result.tax}≠{result.total}")
                pending.append(result)
                continue
        invoices.append(result)
        ctx.ok(src=f)

    # ------------------------------------------------------ 写 Excel
    wb = Workbook()
    wb.remove(wb.active)

    def _money_cell(ws, row, col, value):
        c = ws.cell(row=row, column=col, value=value)
        c.number_format = "#,##0.00"
        return c

    # ---- Sheet 1：发票明细
    ws = wb.create_sheet("发票明细")
    ws.append(list(DET_COLS))
    for c in range(1, len(DET_COLS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
    for inv in invoices:
        note = "；".join(inv.notes)
        ws.append([
            inv.project, inv.kind, inv.subtype,
            inv.use_date or inv.invoice_date,
            inv.route, (inv.seller if inv.kind in ("住宿", "餐饮") else ""),
            inv.seller, inv.invoice_no,
            inv.amount, inv.tax, inv.total, inv.passenger,
            os.path.basename(inv.path), note,
        ])
    widths = (10, 8, 12, 12, 34, 18, 22, 22, 10, 8, 10, 8, 30, 30)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        row[12].alignment = WRAP          # 文件名
        row[13].alignment = WRAP          # 说明
    ws.freeze_panes = "A2"
    last = ws.max_row
    if last > 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(DET_COLS))}{last}"
    # 金额列数字格式
    for row in ws.iter_rows(min_row=2, min_col=9, max_col=11):
        for c in row:
            c.number_format = "#,##0.00"

    # ---- 可选：每项目一张工作表
    if split_by_project:
        used: set[str] = set()
        groups: dict[str, list[Invoice]] = {}
        for inv in invoices:
            groups.setdefault(inv.project, []).append(inv)
        for name, items in sorted(groups.items()):
            sub = wb.create_sheet(_safe_sheet_title(f"明细-{name}", used))
            sub.append(list(DET_COLS))
            for c in range(1, len(DET_COLS) + 1):
                cell = sub.cell(row=1, column=c)
                cell.fill, cell.font = HEAD_FILL, HEAD_FONT
            for inv in items:
                sub.append([
                    inv.project, inv.kind, inv.subtype,
                    inv.use_date or inv.invoice_date, inv.route,
                    (inv.seller if inv.kind in ("住宿", "餐饮") else ""),
                    inv.seller, inv.invoice_no,
                    inv.amount, inv.tax, inv.total, inv.passenger,
                    os.path.basename(inv.path), "；".join(inv.notes),
                ])

    # ---- Sheet 2：项目汇总
    ws2 = wb.create_sheet("项目汇总")
    ws2.append(list(SUM_COLS))
    for c in range(1, len(SUM_COLS) + 1):
        cell = ws2.cell(row=1, column=c)
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
    buckets: dict[str, dict[str, list[float]]] = {}
    for inv in invoices:
        b = buckets.setdefault(inv.project, {"出行": [0, 0.0], "住宿": [0, 0.0], "餐饮": [0, 0.0]})
        if inv.kind in b and inv.total is not None:
            b[inv.kind][0] += 1
            b[inv.kind][1] += inv.total
    grand = 0.0
    for name in sorted(buckets):
        b = buckets[name]
        total = b["出行"][1] + b["住宿"][1] + b["餐饮"][1]
        grand += total
        ws2.append([name, b["出行"][0], b["出行"][1], b["住宿"][0],
                    b["住宿"][1], b["餐饮"][0], b["餐饮"][1], total])
    ws2.append(["总计",
                sum(buckets[n]["出行"][0] for n in buckets), None,
                sum(buckets[n]["住宿"][0] for n in buckets), None,
                sum(buckets[n]["餐饮"][0] for n in buckets), None, grand])
    for c in range(2, 9):
        ws2.cell(row=ws2.max_row, column=c).font = HEAD_FONT
    for row in ws2.iter_rows(min_row=2, min_col=3, max_col=8):
        for c in row:
            c.number_format = "#,##0.00"
    for i, w in enumerate((14, 9, 12, 9, 12, 9, 12, 12), 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # ---- Sheet 3：待确认
    if pending:
        ws3 = wb.create_sheet("待确认")
        ws3.append(list(PEND_COLS))
        for c in range(1, len(PEND_COLS) + 1):
            cell = ws3.cell(row=1, column=c)
            cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        for item in pending:
            if isinstance(item, Unreadable):
                ws3.append([os.path.basename(item.path), item.project,
                            item.reason, item.seen])
            else:
                ws3.append([os.path.basename(item.path), item.project,
                            "；".join(item.notes) or "（待人工确认）", item.summary()])
        for i, w in enumerate((32, 12, 40, 40), 1):
            ws3.column_dimensions[get_column_letter(i)].width = w
        for row in ws3.iter_rows(min_row=2):
            for c in row:
                c.alignment = WRAP

    out = unique_file(os.path.join(ctx.out_dir, "发票整理汇总.xlsx"))
    wb.save(out)
    ctx.report.outputs.append(out)
    ctx.metric("有效票数", len(invoices))
    ctx.metric("待确认", len(pending))
    ctx.ok(out)
    ctx.log(f"已生成汇总：{out}（有效 {len(invoices)} 张，待确认 {len(pending)} 张）")


def _safe_sheet_title(title: str, used: set[str]) -> str:
    """Excel 工作表名限 31 字符且不能含 : \\ / ? * [ ]。"""
    bad = ":\\/?*[]"
    clean = "".join("_" if ch in bad else ch for ch in str(title)).strip() or "Sheet"
    clean = clean[:31]
    base, idx = clean, 2
    while clean.lower() in used:
        suffix = f"_{idx}"
        clean = base[: 31 - len(suffix)] + suffix
        idx += 1
    used.add(clean.lower())
    return clean
