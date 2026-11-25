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
   · 用户填了项目名 → 全部归该项目（散装 PDF 走这条最省事）；
   · 所有文件同在一个目录 → 项目名 = 该目录名；
   · 文件分布在公共父目录的多个子目录 → 项目名 = 第 1 层子目录名；
   · 子目录名只是费用类别（打的发票/住宿发票）→ 上溯一层取父目录名；
   · 目录名是桌面/下载/测试/输出 这类无业务含义的名字 → 「未归属」，
     绝不拿它冒充项目名；散在公共父目录根下的同样「未归属」。
   · 关键：用户若只选了一堆散 PDF（而不是整个项目文件夹），路径里就
     没有项目层级了，此时必须提醒他填项目名，而不是悄悄写个错的值。

4. 一张票可以含多段行程（实测红冲网约车票 6 段）。明细行按「票」组织，
   多段行程用「；」连接写进行程列 —— 与用户现有报销表的人工习惯一致，
   也保证金额列按票求和不重不漏。

5. 铁路票有「乘车日期」和「开票日期」两个日期，报销口径用乘车日期；
   住宿票的「入离日期」在备注栏（如 10-8至10-9,共1天），打车票的日期/起终点
   在配套的行程单里 —— 这两种才是真实消费日期。实在都没有时，才退回文件名里的
   日期并标注来源，最后才用开票日期（备注里会提示人工调整）。

6. 「按日期」表是主表，一行 = 某项目的某一天：同一天的所有票合进一行的
   「费用明细」多行文本，金额按 交通/餐饮/住宿 分列。
   一张票含多天行程时（网约车月结，实测有 10 段）按天数均分到各天；
   多晚住宿按晚均分。明细里注明「整票 ¥X 含 N 天，本日第 k 份」——
   报销表本来就是按天汇总的，整笔钱砸在某一天会让那行虚高。均分只影响
   「按日期」表的分摊展示，分份金额之和严格等于票面价税合计（末份补分位差），
   「发票明细」仍是一票一行。日期未知的票列在表尾单独一段。都不丢，也不猜。

7. 行程单、酒店结账单不是发票（没有发票号码和价税合计），单独解析为「附件」：
   按同名主干（交通-10-08-…-行程单 ↔ …-发票）优先、同金额+同目录次之，
   挂到对应发票上补全日期/起终点/酒店品牌/入住人，并单独列一张「附件」表。
   它们不再进「待确认」——那不是识别失败，是本来就不该按发票计。
"""
from __future__ import annotations

import datetime as _dt
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
# 住宿票在「备注」栏写明入离日期，这是票面上唯一的真实消费日期：
#   入离日期:10-8至10-9,共1天,共1间,订单姓名:郑向荣
RE_STAY_RANGE = re.compile(
    r"入离日期\s*[:：]\s*(?:(\d{4})\s*[-年]\s*)?(\d{1,2})\s*[-月]\s*(\d{1,2})\s*日?"
    r"\s*至\s*(?:(\d{4})\s*[-年]\s*)?(\d{1,2})\s*[-月]\s*(\d{1,2})\s*日?")
RE_STAY_ROOMS = re.compile(r"共\s*(\d+)\s*间")
RE_STAY_GUEST = re.compile(r"订单姓名\s*[:：]\s*([\u4e00-\u9fa5]{2,4})")
# 行程单：'· 行程起止日期：2026-10-08 至 2026-10-08'
RE_TRIP_RANGE = re.compile(
    r"行程起止日期\s*[:：]\s*(\d{4}-\d{2}-\d{2})\s*至\s*(\d{4}-\d{2}-\d{2})")
RE_TRIP_TOTAL = re.compile(r"合计\s*([\d,]+(?:\.\d{1,2})?)\s*元")
RE_TRIP_TIME = re.compile(r"(?:(\d{4})-)?(\d{1,2})-(\d{1,2})\s+\d{1,2}:\d{2}")
# 结账单
RE_FOLIO_IN = re.compile(r"入住日期\s*[:：]\s*(\d{4}-\d{2}-\d{2})")
RE_FOLIO_OUT = re.compile(r"离店日期\s*[:：]\s*(\d{4}-\d{2}-\d{2})")
RE_FOLIO_TOTAL = re.compile(r"消费合计\s*([\d,]+\.\d{1,2})")
# 文件名里的日期（交通-10-08-… / 住宿-10-08-…），仅在没有票面日期时兜底
RE_FILE_DATE = re.compile(
    r"(?:住宿|交通|打的|打车|动车|餐饮|出租车)[-_ ]?(?:20\d{2}[-_.年])?"
    r"(\d{1,2})[-_.月](\d{1,2})")

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
    # 住宿票的入住日与晚数（来自票面备注或结账单）。跨晚住宿按晚均分到「按日期」。
    stay_from: str = ""
    stay_nights: int = 0
    # 结账单上的酒店品牌（如「星程南昌新建吾悦广场酒店」）；
    # 发票销售方是开票主体（如「南昌百业酒店管理有限公司」），报销时品牌更好认。
    brand: str = ""
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


@dataclass
class Attachment:
    """发票的配套单据：网约车行程单、酒店结账单。

    它们不是发票（没有发票号码、没有价税合计），以前会被当成「读不出的发票」
    塞进待确认，既制造噪音，又浪费了票面上唯一可靠的消费日期与起终点。
    现在单独解析，按「同名文件 → 同金额+同目录」匹配到发票，用于补全
    日期 / 起终点 / 酒店品牌 / 入住人。
    """
    path: str
    project: str
    kind: str                      # 行程单 / 结账单
    total: float | None = None
    dates: list[str] = field(default_factory=list)
    trips: list[Trip] = field(default_factory=list)
    hotel: str = ""
    guest: str = ""
    stay_from: str = ""
    stay_nights: int = 0
    matched: str = ""              # 匹配到的发票文件名
    note: str = ""


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
    try:
        y, m, d = (int(x) for x in iso.split("-"))
        return WEEKDAYS[_dt.date(y, m, d).weekday()]
    except (ValueError, TypeError, IndexError):
        return ""


def add_days(iso: str, n: int) -> str:
    """2026-10-08 + 1 → 2026-10-09。"""
    try:
        y, m, d = (int(x) for x in iso.split("-"))
        return (_dt.date(y, m, d) + _dt.timedelta(days=n)).isoformat()
    except (ValueError, TypeError, IndexError):
        return ""


def _stay_dates(text: str, inv: "Invoice") -> None:
    """住宿票备注栏的「入离日期:10-8至10-9,共1天」是票面唯一的真实住店日期。

    只写月日时年份取开票年份；跨年（12 月住、1 月开票）时离店日进下一年。
    """
    m = RE_STAY_RANGE.search(text)
    if not m:
        return
    ref = iso_date(inv.invoice_date)
    base_year = int(ref[:4]) if ref else _dt.date.today().year
    y1 = int(m.group(1)) if m.group(1) else base_year
    y2 = int(m.group(4)) if m.group(4) else y1
    mo1, d1 = int(m.group(2)), int(m.group(3))
    mo2, d2 = int(m.group(5)), int(m.group(6))
    try:
        start = _dt.date(y1, mo1, d1)
        if y2 == y1 and mo2 < mo1 - 6:
            y2 += 1
        end = _dt.date(y2, mo2, d2)
    except ValueError:
        return
    nights = (end - start).days
    if nights <= 0:
        nights = 1
    inv.stay_from = start.isoformat()
    inv.stay_nights = nights
    if not iso_date(inv.use_date):
        inv.use_date = inv.stay_from
    if not inv.qty:
        inv.qty = str(nights)
    inv.notes.append(
        f"入住 {start:%m-%d} 至 {end:%m-%d}，共 {nights} 晚（取自票面备注）")


def apply_filename_date(inv: "Invoice") -> None:
    """票面实在没有消费日期时，用文件名里的日期兜底，并在说明里写明来源。

    只认「住宿-10-08-…」这种带业务前缀的命名，避免把发票号码、开票日期
    里的数字误当消费日期。优先级低于票面/行程单/结账单。
    """
    if iso_date(inv.use_date):
        return
    m = RE_FILE_DATE.search(os.path.basename(inv.path))
    if not m:
        return
    mo, d = int(m.group(1)), int(m.group(2))
    if not (1 <= mo <= 12 and 1 <= d <= 31):
        return
    ref = iso_date(inv.invoice_date)
    year = int(ref[:4]) if ref else _dt.date.today().year
    # 1 月开票、12 月消费 → 消费在上一年
    if ref and mo > int(ref[5:7]) + 2:
        year -= 1
    try:
        day = _dt.date(year, mo, d).isoformat()
    except ValueError:
        return
    inv.use_date = day
    inv.notes.append("日期取自文件名（票面未印消费日期），请核对")


def sync_use_date(inv: "Invoice") -> None:
    """把补全后的行程/住店日期同步成这张票的消费日期。"""
    if inv.kind == "住宿" and inv.stay_from:
        inv.use_date = inv.stay_from
        return
    days = [t.date for t in inv.trips if t.date]
    if days:
        inv.use_date = min(days)


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
    """这张票的行程覆盖了哪几天（升序去重）。

    住宿票按「晚」算：10-20 入住、10-23 离店 = 住 3 晚（20/21/22），
    费用按晚均分，而不是把 3 晚房费全压在入住当天。
    """
    if inv.kind == "住宿" and inv.stay_from and inv.stay_nights > 1:
        return [d for d in (add_days(inv.stay_from, i)
                            for i in range(inv.stay_nights)) if d]
    return sorted({t.date for t in inv.trips if t.date})


def day_shares(inv: "Invoice") -> list["Invoice"]:
    """把一张含多天行程的票按天均分成若干份，只供「按日期」表使用。

    报销表按天汇总，一张网约车月结票（实测一张含 10 段、跨 6 天）整笔钱
    砸在某一天会让那行虚高。这里按「天数」平分——不是按段数，因为用户
    是按天报销的，每天一份才好对着单据核。金额末份补分位差，
    保证各份之和严格等于票面价税合计，不多算不漏算。

    住宿票同理，按「晚」平分多晚房费；单晚不拆。

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
# 常见「按费用类别分文件夹」的目录名。若第 1 层全是这类名字，说明它们不是项目，
# 真正的项目是它们的父文件夹（如「洪盾 10 月报销」），否则项目汇总会被拆成
# 「住宿发票 / 动车票 / 打的发票」三行，没法按项目对账。
COST_DIRS = frozenset({
    "住宿", "住宿发票", "住宿费", "酒店", "酒店发票",
    "交通", "交通费", "交通发票", "打车", "打的", "打的发票", "出租车",
    "动车", "动车票", "火车", "火车票", "高铁", "高铁票", "机票",
    "餐饮", "餐费", "餐饮发票", "发票", "票据", "其他",
})

# 明显不是项目名的目录。散装 PDF 常常就摊在桌面/下载/测试/输出目录里，
# 拿它当项目名会写出「测试」「输出_3」这种没有业务含义的值 —— 宁可标
# 「未归属」并在确认弹窗里提醒用户填，也不悄悄编一个名字。
_NOISE_DIR_EXACT = frozenset({
    "desktop", "桌面", "downloads", "下载", "documents", "文档",
    "pictures", "图片", "temp", "tmp", "test", "测试",
    "新建文件夹", "newfolder", "未命名文件夹", "untitledfolder",
    "发票", "票据", "报销", "报销单", "报销凭证", "扫描件",
})
_NOISE_DIR_PATTERN = re.compile(
    r"^(?:输出|output|导出|export|发票整理成表格)(?:输出)?\d*$"
    r"|^(?:temp|tmp|test|测试|新建文件夹)\d*$",
    re.I,
)


def _is_noise_dir(name: str) -> bool:
    """目录名是否属于「系统/临时/输出」这类不能当项目名的名字。"""
    n = re.sub(r"[\s_\-]", "", name).lower()
    if not n or n in _NOISE_DIR_EXACT:
        return True
    return bool(_NOISE_DIR_PATTERN.fullmatch(n)) or bool(
        re.fullmatch(r"tmp[a-z0-9]{6,}", n))  # tempfile.mkdtemp 之类


def _clean_project_name(name: str, parent: str) -> str:
    """把目录名洗成可用的项目名；不可信的一律「未归属」，不猜。"""
    if not name:
        return "未归属"
    if name in COST_DIRS:
        # 「打的发票 / 住宿发票」只是费用分类，往上取一层才是项目。
        return parent if parent and not _is_noise_dir(parent) else "未归属"
    if _is_noise_dir(name):
        return "未归属"
    return name


def assign_projects(files: list[str], explicit: str = "", level: int = 1) -> dict[str, str]:
    """把每个文件归到一个项目。规则见模块 docstring 第 3 条。"""
    if explicit:
        return {f: explicit for f in files}
    dirs = [os.path.dirname(os.path.abspath(f)) for f in files]
    out: dict[str, str] = {}
    if len(set(dirs)) == 1:
        d = dirs[0].rstrip("\\/")
        name = os.path.basename(d)
        parent = os.path.basename(os.path.dirname(d).rstrip("\\/"))
        p = _clean_project_name(name, parent)
        return {f: p for f in files}
    try:
        common = os.path.commonpath(dirs)
    except ValueError:  # 不同盘符
        return {f: "未归属" for f in files}
    common_parent = os.path.basename(common.rstrip("\\/"))
    for f, d in zip(files, dirs):
        rel = os.path.relpath(d, common)
        parts = [] if rel == "." else rel.split(os.sep)
        if len(parts) < level:
            out[f] = "未归属"          # 散在公共父目录根下，绝不猜
            continue
        name = parts[level - 1]
        parent = parts[level - 2] if level >= 2 else common_parent
        out[f] = _clean_project_name(name, parent)
    # 兼容：若整批推断出来仍全是费用类别，统一降级取公共父目录名。
    if set(out.values()) <= COST_DIRS and common_parent and not _is_noise_dir(common_parent):
        return {f: common_parent for f in files}
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


# ================================================================= 附件（行程单 / 结账单）
def _is_invoice_text(text: str) -> bool:
    """票面是否是一张发票：有发票号码，或铁路电子客票。"""
    return bool(RE_INVOICE_NO.search(text)) or "铁路电子客票" in text


def _is_trip_receipt(text: str) -> bool:
    """网约车行程单：标题可能折行，认「行程起止日期」或表格列名更稳。"""
    return ("行程起止日期" in text
            or ("上车时间" in text and "起点" in text and "终点" in text)
            or "行程单" in text)


def _is_folio(text: str) -> bool:
    """酒店结账单：有入住/离店日期与消费合计。"""
    return (("入住日期" in text and "离店日期" in text)
            or ("结账单" in text and "消费合计" in text))


def _receipt_trips(page, year: int) -> list[Trip]:
    """从行程单表格里抽起终点。

    行程单的地址列会折行（实测「南昌西站-南广场-…上车点」跨两行），
    所以按 y 聚类成行、按表头列中心分列，续行并回上一段行程——
    和发票里的出行信息表同一套做法。
    """
    try:
        words = page.get_text("words")
    except Exception:
        return []
    heads = {}
    for w in words:
        if w[4] in ("城市", "起点", "终点", "里程", "上车时间"):
            heads.setdefault(w[4], w)
    if "起点" not in heads or "终点" not in heads:
        return []

    def cx(w):
        return (w[0] + w[2]) / 2

    xo, xd = cx(heads["起点"]), cx(heads["终点"])
    b_meta = (cx(heads["城市"]) + xo) / 2 if "城市" in heads else xo - 40
    b_dest = (xo + xd) / 2
    b_after = (xd + cx(heads["里程"])) / 2 if "里程" in heads else xd + 40
    top = max(heads["起点"][3], heads["终点"][3])

    rows: list[dict] = []
    for w in sorted(words, key=lambda w: (w[1], w[0])):
        cy = (w[1] + w[3]) / 2
        if cy <= top:
            continue
        if "页码" in w[4]:
            break
        if rows and abs(cy - rows[-1]["y"]) <= 5:
            rows[-1]["ws"].append(w)
            rows[-1]["y"] = (rows[-1]["y"] + cy) / 2
        else:
            rows.append({"y": cy, "ws": [w]})

    trips: list[Trip] = []
    cur: Trip | None = None
    for row in rows:
        line = " ".join(w[4] for w in sorted(row["ws"], key=lambda w: w[0]))
        m = RE_TRIP_TIME.search(line)
        origin, dest = [], []
        for w in row["ws"]:
            c = cx(w)
            if c < b_meta:
                continue
            if c < b_dest:
                origin.append(w)
            elif c < b_after:
                dest.append(w)
        if m:
            mo, dd = int(m.group(2)), int(m.group(3))
            try:
                day = _dt.date(int(m.group(1)) if m.group(1) else year, mo, dd).isoformat()
            except ValueError:
                day = ""
            cur = Trip(date=day, origin=_join_cell(origin), dest=_join_cell(dest))
            trips.append(cur)
        elif cur is not None:
            cur.origin += _join_cell(origin)
            cur.dest += _join_cell(dest)
    return [t for t in trips if t.origin or t.dest]


def _parse_trip_receipt(path: str, project: str, page, text: str) -> Attachment:
    att = Attachment(path=path, project=project, kind="行程单")
    m = RE_TRIP_RANGE.search(text)
    year_m = re.search(r"申请日期\s*[:：]\s*(\d{4})", text)
    year = int(year_m.group(1)) if year_m else _dt.date.today().year
    if m:
        year = int(m.group(1)[:4])
    tm = RE_TRIP_TOTAL.search(text)
    if tm:
        att.total = _f(tm.group(1))
    att.trips = _receipt_trips(page, year)
    dates = sorted({t.date for t in att.trips if t.date})
    if not dates and m:
        start, end = m.group(1), m.group(2)
        try:
            a = _dt.date.fromisoformat(start)
            b = _dt.date.fromisoformat(end)
            if 0 <= (b - a).days <= 62:
                dates = [(a + _dt.timedelta(days=i)).isoformat()
                         for i in range((b - a).days + 1)]
        except ValueError:
            pass
    att.dates = dates
    return att


def _parse_folio(path: str, project: str, text: str) -> Attachment:
    att = Attachment(path=path, project=project, kind="结账单")
    tm = RE_FOLIO_TOTAL.search(text)
    if tm:
        att.total = _f(tm.group(1))
    im, om = RE_FOLIO_IN.search(text), RE_FOLIO_OUT.search(text)
    if im:
        att.stay_from = im.group(1)
        att.dates = [att.stay_from]
        if om:
            try:
                nights = (_dt.date.fromisoformat(om.group(1))
                          - _dt.date.fromisoformat(im.group(1))).days
                att.stay_nights = max(1, nights)
            except ValueError:
                att.stay_nights = 1
        else:
            att.stay_nights = 1
    for line in text.splitlines():
        s = line.strip()
        if s and "结账单" not in s:
            att.hotel = s
            break
    gm = re.search(r"客人姓名\s*[:：]\s*([\u4e00-\u9fa5]{2,4})", text)
    if gm:
        att.guest = gm.group(1)
    return att


def parse_document(path: str, project: str) -> "Invoice | Attachment | Unreadable":
    """先判断这份 PDF 是什么，再交给对应解析器。

    发票 → Invoice；行程单/结账单 → Attachment；读不了或不是这三类 → Unreadable。
    """
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
        if _is_invoice_text(text):
            return parse_invoice(path, project)
        if _is_trip_receipt(text):
            return _parse_trip_receipt(path, project, page, text)
        if _is_folio(text):
            return _parse_folio(path, project, text)
        return Unreadable(path, project,
                          "不是发票：未读到发票号码/价税合计，也不像行程单或结账单")
    finally:
        doc.close()


def _stem_key(path: str) -> str:
    """去掉 发票/行程单/结账单(-1) 后缀，得到配对用的文件名主干。"""
    s = os.path.splitext(os.path.basename(path))[0]
    return re.sub(r"[-_ ]?(?:发票|行程单|结账单|账单)(?:[-_ ]?\d+)?$", "", s).strip()


def link_attachments(invoices: list[Invoice], attachments: list[Attachment]) -> None:
    """把附件挂到对应发票上，补全票面缺失的日期/起终点/酒店/入住人。

    匹配顺序：同名主干（交通-10-08-…-行程单 ↔ 交通-10-08-…-发票）
    → 同目录 + 同金额 + 同类别。同名优先，因为同一天可能有两张同金额的的士票。
    """
    by_stem = {_stem_key(i.path): i for i in invoices}
    for att in attachments:
        inv = by_stem.get(_stem_key(att.path))
        if inv is not None and inv.project != att.project:
            inv = None
        if inv is None:
            def _cands(same_dir: bool):
                return [i for i in invoices
                        if i.total is not None and att.total is not None
                        and abs(i.total - att.total) < 0.011
                        and i.project == att.project
                        and (not same_dir or os.path.dirname(i.path) == os.path.dirname(att.path))
                        and ((att.kind == "行程单" and i.kind == "出行")
                             or (att.kind == "结账单" and i.kind == "住宿"))]
            cands = _cands(True) or _cands(False)
            if len(cands) == 1:
                inv = cands[0]
            elif len(cands) > 1:
                att.note = "同金额发票有多张，无法确定对应关系，请人工核对"
                continue

        if inv is None:
            att.note = "未匹配到发票（同名或同金额+同目录均未命中）"
            continue

        att.matched = os.path.basename(inv.path)
        if att.kind == "行程单":
            if att.trips:
                known = {(t.date, t.origin, t.dest) for t in inv.trips}
                for t in att.trips:
                    if (t.date, t.origin, t.dest) not in known:
                        inv.trips.append(t)
                        known.add((t.date, t.origin, t.dest))
                inv.trips.sort(key=lambda t: (t.date, t.origin))
            if att.dates:
                days = att.dates + ([inv.use_date] if iso_date(inv.use_date) else [])
                inv.use_date = min(days)
            if att.trips or att.dates:
                inv.notes.append("日期/起终点取自行程单")
        else:
            used = False
            if att.stay_from and not inv.stay_from:
                inv.stay_from = att.stay_from
                inv.stay_nights = att.stay_nights
                used = True
            if att.stay_from and not iso_date(inv.use_date):
                inv.use_date = att.stay_from
                used = True
            if att.hotel and not inv.brand:
                inv.brand = att.hotel
            if att.guest and not inv.passenger:
                inv.passenger = att.guest
            if used:
                inv.notes.append("入住日期取自结账单")


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
    if inv.kind == "住宿":
        _stay_dates(text, inv)
        if inv.seller:
            inv.notes.append(f"酒店名取自销售方：{inv.seller}")
        inv.subtype = "酒店"
        if not inv.qty:
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
        if not _is_invoice_text(text):
            return Unreadable(path, project,
                              "不是发票：未读到发票号码/价税合计（行程单、结账单会单独归入附件）")
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
NOISE_NOTES = ("酒店名取自销售方", "铁路票票面无税额列", "进项税额按 9%", "税号 ", "销售方税号",
               "日期/起终点取自行程单", "入住日期取自结账单")


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
        name = inv.brand or inv.seller or "（未读到酒店名）"
        text = f"住宿：{name}{qty} {money}"
        if inv.share:
            part, parts, orig = inv.share
            text += (f"\n（整票 ¥{orig:,.2f} 含 {parts} 晚，按晚均分，"
                     f"本晚第 {part}/{parts} 份）")
        return text
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
            auto.append("住宿票面未读到入离日期，此行暂按开票日期归入，请按实际入住日调整")
        if "餐饮" in weak:
            auto.append("餐饮票面无消费日期，按开票日期归入")
        # 均分票必须在备注列点名：报销单审核时一眼能看出这一行的钱不是票面全额，
        # 否则「整票 111.75 / 当日 15.96」对不上，会被当成算错。
        shared = [i for i in items if i.share and i.share[1] > 1]
        if shared:
            bits = []
            for i in shared:
                _, parts, orig = i.share        # 已过滤，必有值
                night = i.kind == "住宿"
                unit = "晚" if night else "天"
                bits.append(f"整票 ¥{orig:,.2f}÷{parts} {unit}，本"
                            f"{'晚' if night else '日'} ¥{(i.total or 0):,.2f}")
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
                value=(f"含 {len(split)} 张跨天票/多晚住宿，金额已按天（晚）均分到各日，"
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
    lines = [f"即将解析 {len(files)} 个 PDF 并生成汇总 Excel（发票计入金额，"
             f"行程单/结账单只用于补全信息）："]
    lines.append("项目归属（按文件夹自动判定）：")
    for name, n in sorted(buckets.items()):
        lines.append(f"  · {name}：{n} 个")
    if any(p in ("未归属", "?") for p in buckets):
        lines.append("⚠ 有文件没识别出项目名：如果这些票属于同一个项目，请在左侧")
        lines.append("   「项目名（可选）」里填一个名字后重新运行；或改用「添加项目文件夹」")
        lines.append("   选整个项目文件夹，项目名会自动带上。")
    lines.append("「按日期」表把同一天的花费合并成一行（交通/餐饮/住宿 分列），")
    lines.append("一张票含多天行程或住多晚时，金额按天（晚）均分到各日。")
    lines.append("并附发票明细、项目汇总、附件（行程单/结账单）与待确认。")
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
        "一张票含多天行程时（如网约车月结票），金额按天数均分到各日；"
        "多晚住宿按晚均分。明细里注明整票金额与第几份，各日之和等于票面价税合计。\n"
        "网约车行程单、酒店结账单会自动配对到对应发票，用来补全真实日期、"
        "起终点和酒店品牌，并单独列一张「附件」表；它们不是发票，不计入票数。\n"
        "另附「发票明细」（一张票一行，可溯源）、「项目汇总」和「待确认」。\n"
        "项目归属按文件夹结构判定：建议用「添加项目文件夹」选整个项目文件夹，\n"
        "或直接拖入整个文件夹；若手上是一堆散装 PDF，请在参数里填「项目名（可选）」。\n"
        "桌面/下载/测试/输出 这类目录名不会被当成项目名，会标「未归属」提醒你补填。\n"
        "只支持有文本层的电子发票 PDF；扫描件/照片不做 OCR，会列入「待确认」。\n"
        "全程只读，不修改发票文件。"
    ),
    params=(
        Param("project_name", "项目名（可选）", "entry", default="", width=20,
              hint=("留空 = 按文件夹自动归属（一个项目一个文件夹，建议用"
                    "「添加项目文件夹」或拖入整个文件夹）。\n"
                    "如果只上传了一堆散装 PDF，路径里没有项目信息，"
                    "请在这里填一个项目名；填写后所有票统一归这个项目。")),
        Param("project_level", "项目名取第几层文件夹", "choice", default="1",
              choices=(("1", "第 1 层子目录"), ("2", "第 2 层子目录")),
              visible_if=("project_name", ("",))),
        Param("dedupe", "按发票号码去重", "bool", default=True,
              hint="同一张发票被保存了两份时只计一次，重复的列入「待确认」。"),
        Param("check_amount", "金额自洽校验", "bool", default=True,
              hint="校验 金额+税额 ≈ 价税合计（误差 ≤ 0.02 元），不符的列入「待确认」。"),
        Param("use_filename_date", "票面无日期时用文件名日期兜底", "bool", default=True,
              hint="住宿/餐饮票面常常只印开票日期；打车票的日期在行程单里。"
                   "两者都读不到时，认「交通-10-08-…」这类文件名里的月日，并在说明里标注来源。"),
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
    use_filename_date = bool(params.get("use_filename_date", True))

    projects = assign_projects(files, project_name, level)
    ctx.set_total(len(files))
    ctx.metric("模式", "发票整理（只读解析）")

    invoices: list[Invoice] = []
    attachments: list[Attachment] = []
    pending: list[Unreadable | Invoice] = []
    seen_no: dict[str, Invoice] = {}

    for f in files:
        ctx.check()
        result = parse_document(f, projects.get(f, "未归属"))
        ctx.step(f"解析 {os.path.basename(f)}")
        if isinstance(result, Attachment):
            attachments.append(result)
            ctx.skip(f, f"{result.kind}（发票附件，不单独计入）")
            continue
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

    # 行程单/结账单补全日期与起终点；仍缺日期时才用文件名兜底
    link_attachments(invoices, attachments)
    for inv in invoices:
        sync_use_date(inv)
        if use_filename_date:
            apply_filename_date(inv)

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
            inv.route, ((inv.brand or inv.seller) if inv.kind in ("住宿", "餐饮") else ""),
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
                    ((inv.brand or inv.seller) if inv.kind in ("住宿", "餐饮") else ""),
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

    # ---- Sheet 4：附件（行程单/结账单）：不是发票，用来补全发票信息
    if attachments:
        wsA = wb.create_sheet("附件")
        cols = ("文件名", "类型", "项目", "匹配到的发票", "金额", "日期", "说明")
        wsA.append(list(cols))
        for c in range(1, len(cols) + 1):
            cell = wsA.cell(row=1, column=c)
            cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        for att in attachments:
            note = att.note or ("已匹配，用于补全日期/起终点" if att.kind == "行程单"
                                else "已匹配，用于补全入住信息")
            wsA.append([os.path.basename(att.path), att.kind, att.project,
                        att.matched or "（未匹配）", att.total,
                        "、".join(d[5:] for d in att.dates), note])
        for i, w in enumerate((34, 10, 12, 34, 10, 22, 40), 1):
            wsA.column_dimensions[get_column_letter(i)].width = w
        for row in wsA.iter_rows(min_row=2):
            for c in row:
                c.alignment = WRAP
        for row in wsA.iter_rows(min_row=2, min_col=5, max_col=5):
            for c in row:
                c.number_format = "#,##0.00"

    # ---- Sheet 5：待确认
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
    ctx.metric("附件", len(attachments))
    ctx.metric("待确认", len(pending))
    ctx.ok(out)
    ctx.log(f"已生成汇总：{out}（有效 {len(invoices)} 张，附件 {len(attachments)} 份，"
            f"待确认 {len(pending)} 张）")


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
