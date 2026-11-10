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

6. 「按日期」表是主表，一行 = 某项目的某一天：同一天的所有票合进一行的
   「费用明细」多行文本，金额按 交通/餐饮/住宿 分列。
   一张票含多天行程时（网约车月结，实测有 10 段），按天数均分到各天，
   明细里注明「整票 ¥X 含 N 天，本日第 k 份」——报销表本来就是按天汇总的，
   整笔钱砸在某一天会让那行虚高。均分只影响「按日期」表的分摊展示，
   分份金额之和严格等于票面价税合计（末份补分位差），「发票明细」仍是一票一行。
   日期未知的票列在表尾单独一段。都不丢，也不猜。
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field, replace

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
SUB_FILL = PatternFill("solid", fgColor="F2F2F2")
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
    qty: str = ""           # 票面数量（住宿票的间/晚数），读不到就留空，不猜
    trips: list[Trip] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)   # 说明（含待确认原因）
    # 只在「按日期」表里出现：跨天票按天均分后的第 part 份（共 parts 份），
    # 原始价税合计 orig。原始发票本身始终保留 share=None。
    share: tuple[int, int, float] | None = None

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


def _lodging_qty(text: str) -> str:
    """住宿票的「数量」（票面写作 间/晚/人）。读不到返回空串，绝不按文件名猜。"""
    for line in text.splitlines():
        if "住宿" not in line:
            continue
        parts = [p for p in re.split(r"\s{2,}", line.strip()) if p]
        # 期望：['*生产生活服务*住宿费', '间', '2 219.3396…', '438.68', '6%', '26.32']
        # 注意「数量」和「单价」之间可能只有一个空格，会被并成一段，所以取前导数字。
        if len(parts) >= 3 and len(parts[1]) <= 2 and not parts[1][:1].isdigit():
            m = re.match(r"(\d+)(?:\s|$)", parts[2])
            if m:
                return m.group(1)
    return ""


# ================================================================= 日期归一
WEEKDAYS = ("周一", "周二", "周三", "周四", "周五", "周六", "周日")


def iso_date(raw: str) -> str:
    """把 2026年07月08日 / 2026-07-08 / 2026.7.8 统一成 2026-07-08。

    识别不了就返回空串 —— 调用方据此把票放进「日期未知」，而不是随便猜一天。
    """
    if not raw:
        return ""
    m = re.search(r"(\d{4})\D{1,2}(\d{1,2})\D{1,2}(\d{1,2})", str(raw))
    if not m:
        return ""
    y, mo, d = (int(g) for g in m.groups())
    if not (1 <= mo <= 12 and 1 <= d <= 31):
        return ""
    return f"{y:04d}-{mo:02d}-{d:02d}"


def weekday_cn(iso: str) -> str:
    """2026-06-10 → 周三。"""
    import datetime as _dt
    try:
        y, m, d = (int(x) for x in iso.split("-"))
        return WEEKDAYS[_dt.date(y, m, d).weekday()]
    except (ValueError, TypeError, IndexError):
        return ""


def invoice_day(inv: "Invoice") -> tuple[str, bool]:
    """这张票该记到哪一天，以及日期是否只是开票日期。

    报销口径优先「消费/乘车日期」；住宿、餐饮票面没有消费日期，只能退回开票日期，
    第二个返回值 True 表示「日期依据较弱」，写表时要在备注里讲清楚。
    """
    d = iso_date(inv.use_date)
    if d:
        return d, False
    return iso_date(inv.invoice_date), True


def trip_days(inv: "Invoice") -> list[str]:
    """这张票的行程覆盖了哪几天（升序去重）。"""
    return sorted({t.date for t in inv.trips if t.date})


def day_shares(inv: "Invoice") -> list["Invoice"]:
    """把一张含多天行程的票按天均分成若干份，只供「按日期」表使用。

    报销表按天汇总，一张网约车月结票（实测一张含 10 段、跨 6 天）整笔钱
    砸在某一天会让那行虚高。这里按「天数」平分——不是按段数，因为用户
    是按天报销的，每天一份才好对着单据核。金额末份补分位差，
    保证各份之和严格等于票面价税合计，不多算不漏算。

    非跨天票、金额未读到的票原样返回（不拆，也不猜金额）。
    """
    days = trip_days(inv)
    if len(days) <= 1 or inv.total is None:
        return [inv]
    n = len(days)
    each = round(inv.total / n, 2)
    out: list[Invoice] = []
    for i, day in enumerate(days):
        amt = round(inv.total - each * (n - 1), 2) if i == n - 1 else each
        out.append(replace(
            inv,
            total=amt,
            use_date=day,          # 这一份就算这一天的花费，不再退回开票日期
            trips=[t for t in inv.trips if t.date == day],
            notes=list(inv.notes),
            share=(i + 1, n, inv.total),
        ))
    return out


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


def _join_cell(ws_: list) -> str:
    """把同一列里的若干 word 拼回一个单元格文本（贴合处不加空格）。"""
    out, prev_x1 = "", None
    for w in sorted(ws_, key=lambda x: x[0]):
        if out and prev_x1 is not None and w[0] - prev_x1 > 2:
            out += " "
        out += w[4]
        prev_x1 = w[2]
    return out.strip()


def _trip_rows(page, text: str, inv: Invoice) -> None:
    """从「出行人/日期/出发地/到达地」表格抽出行段。

    用 words 坐标做两件事，缺一不可：
    · **按 y 聚类**成行 —— 长地址折行时，续行要归到上一段行程；
    · **按表头 x 边界分列** —— 地址文字常从列左边界起排，比表头文字更靠左，
      所以边界取「左列右沿 ↔ 右列左沿」的中点，而不是表头的 x0。
    纯文本正则在这张票上会翻车：单元格被裁剪后列间只剩 1 个空格，
    `\\s{2,}` 分不开出发地与到达地，于是拼出「→无  其他」这种垃圾。
    """
    try:
        words = page.get_text("words")
    except Exception:
        words = []
    header = {}
    for w in words:
        if w[4] in ("出行人", "出行日期", "出发地", "到达地", "等级", "交通工具类型"):
            header[w[4]] = w
    if len(header) < 4 or "出发地" not in header or "到达地" not in header:
        _trip_rows_text(text, inv)          # 没有坐标信息时退回正则（老路径）
        return

    o, d = header["出发地"], header["到达地"]
    # 列归属用「最近的表头中心」：地址文字常从列左边界起排、甚至比表头更靠左，
    # 用「区间边界」会把出发地判给日期列。中心法在这批票上两种版面都稳。
    centers = [(k, (header[k][0] + header[k][2]) / 2) for k in
               ("出行日期", "出发地", "到达地", "等级", "交通工具类型") if k in header]
    slot = {"出发地": "origin", "到达地": "dest", "等级": "seat",
            "交通工具类型": "vehicle"}
    left = header["出行日期"][2] + 10 if "出行日期" in header else o[0] - 40

    top = max(w[3] for w in header.values()) + 1
    bottom = 1e9
    for w in words:
        if "价税合计" in w[4]:
            bottom = min(bottom, w[1])

    rows: list[dict] = []
    for w in sorted(words, key=lambda w: (w[1], w[0])):
        cy = (w[1] + w[3]) / 2
        if not (top < cy < bottom):
            continue
        if rows and abs(cy - rows[-1]["y"]) <= 6:
            rows[-1]["ws"].append(w)
            rows[-1]["y"] = (rows[-1]["y"] + cy) / 2
        else:
            rows.append({"y": cy, "ws": [w]})

    cur: Trip | None = None
    for row in rows:
        cols: dict[str, list] = {"origin": [], "dest": [], "seat": [], "vehicle": []}
        date, name = "", ""
        for w in row["ws"]:
            txt = w[4]
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", txt):
                date = txt
                continue
            if w[0] < left:
                if not name and re.fullmatch(r"[\u4e00-\u9fa5]{2,4}", txt):
                    name = txt
                continue
            key = min(centers, key=lambda kv: abs(kv[1] - (w[0] + w[2]) / 2))[0]
            if key in slot:
                cols[slot[key]].append(w)
        if date:
            seat = _join_cell(cols["seat"])
            vehicle = _join_cell(cols["vehicle"])
            # 有的票版式里没有「等级」列，等级值会落到最后一列 → 最后一段才是交通工具
            if "等级" not in header and vehicle and len(vehicle.split(" ")) > 1:
                bits = vehicle.split(" ")
                seat, vehicle = bits[0], bits[-1]
            cur = Trip(date=date, origin=_join_cell(cols["origin"]),
                       dest=_join_cell(cols["dest"]), seat=seat, vehicle=vehicle)
            if not inv.passenger and name:
                inv.passenger = name
            inv.trips.append(cur)
        elif cur is not None:
            # 折行续行：按列拼回同一段行程
            cur.origin += _join_cell(cols["origin"])
            cur.dest += _join_cell(cols["dest"])

    if inv.trips:
        inv.trips.sort(key=lambda t: (t.date, t.origin))
        inv.use_date = inv.trips[0].date


def _trip_rows_text(text: str, inv: Invoice) -> None:
    """无坐标时的兜底：按行正则（只对列间留 2+ 空格的票有效）。"""
    for m in RE_TRIP_LINE.findall(text):
        if not inv.passenger:
            inv.passenger = m[0].strip()
        inv.trips.append(Trip(date=m[2], origin=m[3].strip(), dest=m[4].strip(),
                              seat=m[5].strip(), vehicle=m[6].strip()))
    if inv.trips:
        inv.trips.sort(key=lambda t: (t.date, t.origin))
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
        inv.qty = _lodging_qty(text)

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

    _trip_rows(page, text, inv)


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


# ================================================================= 按日期汇总
KIND_DISPLAY = {"出行": "交通", "餐饮": "餐饮", "住宿": "住宿", "其他": "其他"}
KIND_ORDER = {"出行": 0, "餐饮": 1, "住宿": 2, "其他": 3}
# 这些是「解析过程说明」，属于审计信息，只该出现在「发票明细」的说明列。
# 「按日期」是给人照着抄报销单的，备注列只留真正需要人动手的事。
NOISE_NOTES = ("酒店名取自销售方", "铁路票票面无税额列", "进项税额按 9%", "税号 ", "销售方税号")


def detail_text(inv: Invoice) -> str:
    """一句话写清这张票，模仿人手写「费用明细」的习惯。

    例：``的士：海上五月花3期(西南门)→厦门北站(南进站口) ¥19.83``
         ``动车：厦门北站→宁德站 ¥139.00``
         ``住宿：…安嘉酒店有限公司（票面数量 2） ¥465.00``
    """
    money = f"¥{inv.total:,.2f}" if inv.total is not None else "金额未读到"
    if inv.kind == "出行":
        if inv.subtype == "动车":
            role = "动车"
        elif "出租" in inv.subtype or "网约" in inv.subtype:
            role = "的士"
        else:
            role = inv.subtype or "交通"
        trips = [t for t in inv.trips if t.date]
        if not trips:
            return f"{role}：{inv.route or '（未读到起终点）'} {money}"
        days = sorted({t.date for t in trips})
        if len(days) <= 1:
            routes = "；".join(f"{t.origin or '?'}→{t.dest or '?'}" for t in trips)
            text = f"{role}：{routes} {money}"
            if inv.share:
                part, parts, orig = inv.share
                text += (f"\n（整票 ¥{orig:,.2f} 含 {parts} 天行程，按天均分，"
                         f"本日第 {part}/{parts} 份）")
            return text
        # 一张票横跨多天（网约车月结）：逐段列日期，否则整段金额会被误读成某一天的车费
        lines = [f"{role}（{days[0][5:]}–{days[-1][5:]}，{len(trips)} 段）{money}"]
        lines += [f"  {t.date[5:]} {t.origin or '?'}→{t.dest or '?'}" for t in trips]
        return "\n".join(lines)
    if inv.kind == "住宿":
        qty = f"（票面数量 {inv.qty}）" if inv.qty else ""
        return f"住宿：{inv.seller or '（未读到酒店名）'}{qty} {money}"
    if inv.kind == "餐饮":
        return f"餐饮：{inv.seller or '（未读到餐厅名）'} {money}"
    return f"其他：{inv.seller or inv.subtype or '未识别'} {money}"


def ticket_count(items: list[Invoice]) -> int:
    """这一段里有几张原始发票。

    跨天票拆成多份后会落在多行，只有第 1 份算「新的一张」，
    否则合计张数会按天翻倍，对不上手里的单据数。
    """
    return sum(1 for i in items if not (i.share and i.share[0] > 1))


def _clean_notes(items: list[Invoice]) -> str:
    notes: list[str] = []
    for inv in items:
        for n in inv.notes:
            if n.startswith(NOISE_NOTES) or n in notes:
                continue
            notes.append(n)
    return "；".join(notes)


def _section_label(ws, row: int, ncol: int, text: str) -> None:
    ws.cell(row=row, column=1, value=text).font = Font(bold=True, color="C00000")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncol)


def _write_date_sheet(wb: Workbook, invoices: list[Invoice]) -> None:
    """「按日期」主表：一行 = 某项目某一天的全部花费（对齐用户的手工报销单）。"""
    import datetime as _dt

    present = {inv.kind for inv in invoices}
    money_cols = [("交通", "出行"), ("餐饮", "餐饮"), ("住宿", "住宿")]
    if "其他" in present:
        money_cols.append(("其他", "其他"))

    headers = (["序号", "日期", "星期", "项目", "费用类别", "费用明细"]
               + [label for label, _ in money_cols]
               + ["当日合计", "张数", "出行人", "备注"])
    ncol = len(headers)
    first_money = 7
    total_col = first_money + len(money_cols)

    ws = wb.create_sheet("按日期")
    projects = sorted({inv.project for inv in invoices})
    days = []
    for inv in invoices:
        d, weak = invoice_day(inv)
        if d:
            days.append(d)
    period = f"{min(days)} ~ {max(days)}" if days else "日期未读到"
    label = projects[0] if len(projects) == 1 else "发票整理汇总"
    ws.cell(row=1, column=1, value=f"{label} 费用明细（{period}）").font = Font(bold=True, size=13)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)

    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    groups: dict[tuple[str, str], list[Invoice]] = {}
    unknown: list[Invoice] = []
    for inv in invoices:
        base_day, _ = invoice_day(inv)
        if not base_day and not trip_days(inv):
            unknown.append(inv)
            continue
        # 一张票含多天行程（网约车月结）按天均分成多份，各归各天；
        # 非跨天票原样返回，行为不变。
        for part in day_shares(inv):
            pd, _ = invoice_day(part)
            if not pd:
                unknown.append(inv)
                break
            groups.setdefault((part.project, pd), []).append(part)

    row = 3
    seq = 0
    grand = {key: 0.0 for _, key in money_cols}
    grand_n = 0

    def put(r: int, values: list) -> None:
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.alignment = WRAP
            if first_money <= c < total_col or c == total_col:
                cell.number_format = "#,##0.00"

    def subtotal(r: int, tag: str, sums: dict[str, float], n: int) -> None:
        ws.cell(row=r, column=1, value=tag).font = HEAD_FONT
        for c in range(1, ncol + 1):
            ws.cell(row=r, column=c).fill = SUB_FILL
            ws.cell(row=r, column=c).font = HEAD_FONT
        for i, (_, key) in enumerate(money_cols):
            ws.cell(row=r, column=first_money + i, value=round(sums[key], 2)) \
                .number_format = "#,##0.00"
        ws.cell(row=r, column=total_col, value=round(sum(sums.values()), 2)) \
            .number_format = "#,##0.00"
        ws.cell(row=r, column=total_col + 1, value=n)

    def write_row(items: list[Invoice], day: str,
                  extra: list[str]) -> tuple[dict[str, float], int]:
        """写一行（一天），返回该行的分类金额与折算后的发票张数。"""
        nonlocal row, seq
        kinds = [KIND_DISPLAY.get(k, k) for k in
                 sorted({i.kind for i in items}, key=lambda k: KIND_ORDER.get(k, 9))]
        detail = "\n".join(detail_text(i) for i in items)
        people = "、".join(dict.fromkeys(i.passenger for i in items if i.passenger))
        auto = list(extra)
        weak = {i.kind for i in items if invoice_day(i)[1]}
        if "住宿" in weak:
            auto.append("住宿票面无人住日期，此行按开票日期归入，请按实际入住日调整")
        if "餐饮" in weak:
            auto.append("餐饮票面无消费日期，按开票日期归入")
        # 均分票必须在备注列点名：报销单审核时一眼能看出这一行的钱不是票面全额，
        # 否则「整票 111.75 / 当日 15.96」对不上，会被当成算错。
        shared = [i for i in items if i.share and i.share[1] > 1]
        if shared:
            bits = []
            for i in shared:
                _, parts, orig = i.share        # 已过滤，必有值
                bits.append(f"整票 ¥{orig:,.2f}÷{parts} 天，本日 ¥{(i.total or 0):,.2f}")
            auto.append("均分票：" + "；".join(bits))
        notes = "；".join(x for x in [_clean_notes(items), *auto] if x)
        amount = {key: 0.0 for _, key in money_cols}
        for i in items:
            if i.total is not None and i.kind in KIND_ORDER:
                amount[i.kind] += i.total
        seq += 1
        day_cell = None
        if day:
            y, m, d = (int(x) for x in day.split("-"))
            day_cell = _dt.date(y, m, d)
        put(row, [seq, day_cell, weekday_cn(day) if day else "", items[0].project,
                  "+".join(kinds), detail] + [round(amount[k], 2) for _, k in money_cols]
            + [round(sum(amount.values()), 2), len(items), people, notes])
        if day:
            ws.cell(row=row, column=2).number_format = "yyyy.m.d"
        ws.row_dimensions[row].height = min(15 * (detail.count("\n") + 1) + 4, 150)
        row += 1
        return amount, ticket_count(items)

    for proj in projects:
        p_sum = {key: 0.0 for _, key in money_cols}
        p_n = 0
        for day in sorted(d for p, d in groups if p == proj):
            items = sorted(groups[(proj, day)],
                           key=lambda i: (KIND_ORDER.get(i.kind, 9), i.invoice_no))
            amount, n = write_row(items, day, [])
            for _, key in money_cols:
                p_sum[key] += amount[key]
                grand[key] += amount[key]
            p_n += n
            grand_n += n
        if len(projects) > 1:
            subtotal(row, f"{proj} 小计", p_sum, p_n)
            row += 1

    if unknown:
        _section_label(ws, row, ncol, "以下发票未读到日期，请人工补录")
        row += 1
        for inv in unknown:
            seq += 1
            put(row, [seq, None, "", inv.project, KIND_DISPLAY.get(inv.kind, inv.kind),
                      detail_text(inv)]
                + [round(inv.total or 0, 2) if inv.kind == key else None
                   for _, key in money_cols]
                + [round(inv.total or 0, 2), 1, inv.passenger,
                   "；".join(x for x in ("未读到日期", _clean_notes([inv])) if x)])
            ws.row_dimensions[row].height = 32
            # 日期缺失的票仍计入合计：钱确实花出去了，不能因为日期不详就漏账
            for _, key in money_cols:
                if inv.kind == key:
                    grand[key] += inv.total or 0
            grand_n += 1
            row += 1

    subtotal(row, "合计", grand, grand_n)
    ws.cell(row=row, column=total_col, value=round(sum(grand.values()), 2)) \
        .number_format = "#,##0.00"
    split = [i for i in invoices if len(trip_days(i)) > 1 and i.total is not None]
    if split:
        ws.cell(row=row, column=ncol,
                value=(f"含 {len(split)} 张跨天票，金额已按天均分到各日，"
                       f"各日金额之和 = 票面价税合计；张数按原始票据数计")).alignment = WRAP

    for i, w in enumerate([5, 11, 6, 12, 11, 56] + [10] * len(money_cols)
                          + [11, 6, 10, 38], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A3"
    if row > 3:
        ws.auto_filter.ref = f"A2:{get_column_letter(ncol)}{row}"


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
    lines.append("「按日期」表把同一天的花费合并成一行（交通/餐饮/住宿 分列），")
    lines.append("一张票含多天行程时（如网约车月结），金额按天数均分到各日。")
    lines.append("并附发票明细、项目汇总、待确认。")
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
        "把一批发票 PDF 整理成 Excel：首张「按日期」表把同一天的花费合并成一行"
        "（交通/餐饮/住宿 分列 + 当日合计），出行票取出发地→到达地，"
        "住宿票取酒店名称，餐饮票取餐厅名称。\n"
        "一张票含多天行程时（如网约车月结票），金额按天数均分到各日，"
        "明细里注明整票金额与第几份，各日之和等于票面价税合计。\n"
        "另附「发票明细」（一张票一行，可溯源）、「项目汇总」和「待确认」。\n"
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

    # 全部按「项目 → 日期 → 类别」排序：用户按时间对账，文件顺序没有意义
    invoices.sort(key=lambda i: (i.project, invoice_day(i)[0] or "9999-99-99",
                                 KIND_ORDER.get(i.kind, 9), i.invoice_no))

    def _money_cell(ws, row, col, value):
        c = ws.cell(row=row, column=col, value=value)
        c.number_format = "#,##0.00"
        return c

    # ---- Sheet 1：按日期（主表，一行 = 某天全部花费）
    _write_date_sheet(wb, invoices)

    # ---- Sheet 2：发票明细（一张票一行，用于溯源）
    ws = wb.create_sheet("发票明细")
    ws.append(list(DET_COLS))
    for c in range(1, len(DET_COLS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
    for inv in invoices:
        note = "；".join(inv.notes)
        ws.append([
            inv.project, inv.kind, inv.subtype,
            invoice_day(inv)[0] or inv.use_date or inv.invoice_date,
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
                    invoice_day(inv)[0] or inv.use_date or inv.invoice_date,
                    inv.route,
                    (inv.seller if inv.kind in ("住宿", "餐饮") else ""),
                    inv.seller, inv.invoice_no,
                    inv.amount, inv.tax, inv.total, inv.passenger,
                    os.path.basename(inv.path), "；".join(inv.notes),
                ])

    # ---- Sheet 3：项目汇总
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

    # ---- Sheet 4：待确认
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
