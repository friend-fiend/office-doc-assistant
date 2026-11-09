"""发票整理功能回归测试。

原则：不依赖用户桌面上的任何真实样票 —— 测试用 reportlab 合成发票 PDF
（CJK 字体用内置的 STSong-Light CID 字体，PyMuPDF 可正常提取文本）。
默认零外部副作用，产物全部进临时目录。
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import load_workbook  # noqa: E402

from app.cancel import CancelToken  # noqa: E402
from app.core import OPS  # noqa: E402
from app.core.invoice_ops import (  # noqa: E402
    Invoice,
    Unreadable,
    assign_projects,
    parse_invoice,
)
from app.core.registry import OpContext  # noqa: E402
from app.engine import TaskReport  # noqa: E402

op = OPS["invoice_summary"]


class StubBus:
    def log(self, *a, **k):
        pass

    def progress(self, *a, **k):
        pass

    def ask_password(self, *a, **k):
        return None


def make_ctx(out_dir: str) -> OpContext:
    os.makedirs(out_dir, exist_ok=True)
    return OpContext(StubBus(), CancelToken(), op, out_dir,
                     TaskReport(op_title=op.title))


# ---------------------------------------------------------------- 合成发票
def _pdf(path: str, rows: list[tuple[float, list[tuple[float, str]]]]) -> None:
    """分列绘制合成发票 PDF（贴近真实数电票版面：每列独立文本、列间大空隙）。

    rows: [(y, [(x, text), ...]), ...]，y 自页顶往下。
    整行绘制会让空格被 PyMuPDF 压缩成单空格，与真实票的多空格列距不符。
    """
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas

    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    c = canvas.Canvas(path, pagesize=(560, 300))
    c.setFont("STSong-Light", 7)
    for y, cells in rows:
        for x, text in cells:
            c.drawString(x, 300 - y, text)
    c.save()


RAIL = [
    (30, [(40, "发票号码: 26949134178001753245"), (300, "开票日期: 2026年09月06日")]),
    (60, [(40, "厦门北站"), (140, "D3234"), (200, "宁德站")]),
    (80, [(40, "2026年06月10日 07:08开"), (160, "05车10C号"), (240, "二等座")]),
    (100, [(40, "票价: ￥139.00")]),
    (120, [(40, "3706021967****3518"), (140, "林锦明")]),
    (150, [(40, "购买方名称: 厦门市博格管理咨询有限公司")]),
    (180, [(40, "电子发票（铁路电子客票）")]),
]
TAXI = [
    (30, [(60, "发票号码： 26327000001223909548"), (330, "电子发票（普通发票）")]),
    (50, [(60, "开票日期： 2026年07月04日")]),
    (70, [(20, "购"), (35, "名称：厦门市博格管理咨询有限公司"),
          (270, "销"), (285, "名称：苏州市吉利优行电子科技有限公司")]),
    (90, [(20, "息"), (35, "统一社会信用代码/纳税人识别号：913502036123127852"),
          (270, "息"), (285, "统一社会信用代码/纳税人识别号：91320594MA1MFD7F31")]),
    (110, [(30, "*交通运输服务*客运服务费"), (200, "19.252427"), (250, "1"),
           (280, "19.25"), (330, "3%"), (370, "0.58")]),
    (130, [(30, "合    计"), (280, "¥19.25"), (370, "¥0.58")]),
    (150, [(20, "出行人"), (60, "有效身份证件号"), (130, "出行日期"), (180, "出发地"),
           (300, "到达地"), (420, "等级"), (460, "交通工具类型")]),
    (170, [(20, "林锦明"), (60, "370602********3518"), (130, "2026-06-18"),
           (180, "海上五月花3期(西南门)"), (300, "厦门北站(南进站口)"),
           (420, "特价车"), (460, "出租车")]),
    (190, [(30, "价税合计（大写）"), (150, "壹拾玖圆捌角叁分"), (300, "（小写）"), (340, "¥19.83")]),
]
HOTEL = [
    (30, [(40, "电子发票（增值税专用发票）"), (280, "发票号码： 26352000001641721951")]),
    (50, [(280, "开票日期： 2026年07月08日")]),
    (70, [(20, "买"), (35, "名称：厦门市博格管理咨询有限公司"),
          (280, "售"), (295, "名称：宁德市东侨经济开发区安嘉酒店有限公司")]),
    (90, [(20, "息"), (35, "统一社会信用代码/纳税人识别号：913502036123127852"),
          (280, "息"), (295, "统一社会信用代码/纳税人识别号：91350901MA32XUTR18")]),
    (110, [(30, "*生产生活服务*住宿费"), (200, "间"), (240, "2"),
           (270, "219.34"), (330, "6%"), (370, "26.32")]),
    (130, [(30, "合    计"), (270, "¥438.68"), (370, "¥26.32")]),
    (150, [(30, "价税合计（大写）"), (140, "肆佰陆拾伍圆整"), (300, "（小写）"), (340, "¥ 465.00")]),
]


def main() -> int:
    tmp = tempfile.mkdtemp(prefix="oda_invoice_test_")
    failures = 0

    def check(label: str, cond: bool, detail: str = "") -> None:
        nonlocal failures
        if cond:
            print(f"  [OK]   {label}")
        else:
            failures += 1
            print(f"  [FAIL] {label} {detail}")

    try:
        # ========== 1. 铁路电子客票
        p_rail = os.path.join(tmp, "动车样例.pdf")
        _pdf(p_rail, RAIL)
        r = parse_invoice(p_rail, "测试项目")
        assert isinstance(r, Invoice), f"铁路票解析失败：{getattr(r, 'reason', '?')}"
        check("铁路票：号码", r.invoice_no == "26949134178001753245", r.invoice_no)
        check("铁路票：类别/类型", (r.kind, r.subtype) == ("出行", "动车"), f"{r.kind}/{r.subtype}")
        check("铁路票：乘车日期优先于开票日期", r.use_date == "2026年06月10日", r.use_date)
        check("铁路票：起终点", r.route == "厦门北站→宁德站", r.route)
        check("铁路票：出行人", r.passenger == "林锦明", r.passenger)
        check("铁路票：票价", r.total == 139.0, str(r.total))

        # ========== 2. 的士/网约车普票（含出行行）
        p_taxi = os.path.join(tmp, "的士样例.pdf")
        _pdf(p_taxi, TAXI)
        r2 = parse_invoice(p_taxi, "测试项目")
        assert isinstance(r2, Invoice), "的士票解析失败"
        check("的士票：类别/类型", (r2.kind, r2.subtype) == ("出行", "出租车·网约车"), f"{r2.kind}/{r2.subtype}")
        check("的士票：出行行起终点", "海上五月花3期(西南门)" in r2.route
              and "厦门北站(南进站口)" in r2.route, r2.route)
        check("的士票：销售方", r2.seller == "苏州市吉利优行电子科技有限公司", r2.seller)
        check("的士票：金额/税额/合计", (r2.amount, r2.tax, r2.total) == (19.25, 0.58, 19.83),
              f"{r2.amount}/{r2.tax}/{r2.total}")

        # ========== 3. 住宿专票：酒店名 = 销售方
        p_hotel = os.path.join(tmp, "住宿样例.pdf")
        _pdf(p_hotel, HOTEL)
        r3 = parse_invoice(p_hotel, "测试项目")
        assert isinstance(r3, Invoice), "住宿票解析失败"
        check("住宿票：类别", r3.kind == "住宿", r3.kind)
        check("住宿票：酒店名取自销售方", r3.seller == "宁德市东侨经济开发区安嘉酒店有限公司", r3.seller)
        check("住宿票：价税合计", r3.total == 465.0, str(r3.total))

        # ========== 4. 无文本层 → 明确拒绝而不是猜
        p_blank = os.path.join(tmp, "扫描件样例.pdf")
        _pdf(p_blank, [])
        r4 = parse_invoice(p_blank, "测试项目")
        check("无文本层：进待确认", isinstance(r4, Unreadable) and "扫描件" in r4.reason,
              getattr(r4, "reason", str(r4)))

        # ========== 5. 项目归属
        sub1 = os.path.join(tmp, "projA"); os.makedirs(sub1, exist_ok=True)
        sub2 = os.path.join(tmp, "projB"); os.makedirs(sub2, exist_ok=True)
        f1 = os.path.join(sub1, "a.pdf"); shutil.copy(p_rail, f1)
        f2 = os.path.join(sub2, "b.pdf"); shutil.copy(p_taxi, f2)
        f3 = os.path.join(tmp, "散票.pdf"); shutil.copy(p_hotel, f3)
        proj = assign_projects([f1, f2, f3])
        check("归属：子目录名即项目", proj[f1] == "projA" and proj[f2] == "projB", str(proj))
        check("归属：根下散票未归属", proj[f3] == "未归属", str(proj))
        single = os.path.join(sub1, "c.pdf"); shutil.copy(p_taxi, single)
        proj2 = assign_projects([f1, single])
        check("归属：全在同一目录→目录名", proj2[f1] == "projA", str(proj2))
        check("归属：显式项目名覆盖", assign_projects([f1, f2], "宁波项目")[f1] == "宁波项目", "")

        # ========== 6. 端到端：去重 + 汇总表
        f4 = os.path.join(sub1, "dup.pdf"); shutil.copy(p_taxi, f4)   # 与 f2 同号码
        ctx = make_ctx(os.path.join(tmp, "out"))
        from app.core.invoice_ops import run_invoice_summary
        ctx.set_total(4)
        run_invoice_summary(ctx, [f1, f2, f4, f3], {
            "project_name": "", "project_level": "1",
            "dedupe": True, "check_amount": True, "split_by_project": False,
        })
        out_xlsx = ctx.report.outputs[0]
        wb = load_workbook(out_xlsx)
        check("端到端：三张工作表", wb.sheetnames == ["发票明细", "项目汇总", "待确认"], str(wb.sheetnames))
        ws = wb["发票明细"]
        check("端到端：有效票 3 张（重复票被剔除）", ws.max_row == 4, f"实际 {ws.max_row - 1} 行")
        first = [c.value for c in ws[2]]
        check("端到端：明细首列是项目", first[0] == "projA", str(first[0]))
        check("端到端：行程列含起终点", "厦门北站" in str(first[4]), str(first[4]))
        ws3 = wb["待确认"]
        dup_note = " ".join(str(c.value) for c in ws3[2])
        check("端到端：重复票列入待确认", "重复" in dup_note, dup_note)
        ws2 = wb["项目汇总"]
        # 表头 + projA/projB/未归属 三个项目行 + 总计行 = 5
        check("端到端：汇总行数", ws2.max_row == 5, f"{ws2.max_row}")
        total_row = [c.value for c in ws2[ws2.max_row]]
        check("端到端：总计金额 139+19.83+465",
              abs((total_row[-1] or 0) - 623.83) < 0.01, str(total_row[-1]))

        # ========== 7. 确认文案（不弹窗，直接调 builder）
        text = op.confirm_builder([f1, f2, f3], {"project_name": "", "project_level": "1"})
        check("确认文案：列出项目与张数", "projA" in text and "1 张" in text, text[:60])

        print()
        if failures:
            print(f"失败 {failures} 项")
            return 1
        print("全部通过")
        return 0
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
