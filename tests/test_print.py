"""PDF 打印 / 导出 PDF 验证（Windows）。

**本测试默认不会向任何打印机提交作业。**

历史上这个文件闯过祸：它直接用"系统默认打印机"跑端到端打印，于是在开发机上
每执行一次就往真实打印机送 4 个作业（a4×1 + a5×2 + a3×1，共 8 页）。
用户侧看到的现象就是"我什么都没干，打印机自己开始出纸"。
所以现在所有真实打印都必须**显式**满足两个条件：

    ODA_TEST_REAL_PRINT=1        # 明确表示"我知道这会出纸"
    ODA_TEST_PRINTER=<打印机名>  # 明确指定打到哪台，不用系统默认

缺一个就跳过真实打印段，只跑演练模式。CI / 无人值守环境天然不满足，
不可能误出纸。

覆盖：
  1. pywin32 + 平台守卫
  2. 打印机枚举（虚拟设备必须被过滤）
  3. 两个 OpDef 注册（pdf_print + pdf_export）
  4. DEVMODE 字节偏移与 PyDEVMODE 交叉校验
  5. 每作业 DEVMODE 生效验证（只读 GetDeviceCaps，不打印）
  6. 未授权路径必须被拒绝（本次事故的回归测试）
  7. 演练模式端到端（不出纸、产出预览 PDF）
  8. 导出 PDF（含合并）—— 无对话框、纸张按需求生效
  9. ParamPanel 构建 per_file_config UI + 覆盖取值
 10. 【可选，需显式开启】真实打印一张
"""
from __future__ import annotations

import os
import sys
import time

import tkinter as tk

THIS = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(THIS)
sys.path.insert(0, ROOT)


def main() -> int:
    print("=" * 70)
    print("PDF 打印 + 导出 PDF 验证")
    print("=" * 70)

    from app.cancel import CancelToken
    from app.core import print_ops as po
    from app.core.registry import OPS, OpContext
    from app.errors import PrintNotAuthorized, TaskReport

    real_print = str(os.environ.get("ODA_TEST_REAL_PRINT", "")).strip() == "1"
    target_printer = str(os.environ.get("ODA_TEST_PRINTER", "")).strip()
    if real_print:
        print(f"!! ODA_TEST_REAL_PRINT=1 —— 本段会真实出纸，目标打印机={target_printer!r}")

    class StubBus:
        def log(self, t, lvl="info"):
            print(f"    [{lvl[:4]}] {t[:110]}")

        def progress(self, d, t, lbl=""):
            pass

        def state(self, t):
            pass

        def done(self, report):
            pass

        def ask_password(self, p, confirm=False):
            return None

    def make_ctx(op, out_dir, **kw):
        os.makedirs(out_dir, exist_ok=True)
        return OpContext(StubBus(), CancelToken(), op, out_dir,
                         TaskReport(op_title=op.title), **kw)

    # ========== 1. 平台与依赖
    assert po.PYWIN32_OK, "pywin32 不可用"
    print(f"[OK] pywin32 + IS_WINDOWS={po.IS_WINDOWS}")

    # ========== 2. 打印机枚举 + 虚拟设备过滤
    real = po.list_printers(real_only=True)
    all_p = po.list_printers(real_only=False)
    virtual = po.list_virtual_pdf_printers()
    print(f"[OK] 真实打印机 {len(real)} 台 / 全部 {len(all_p)} / 虚拟设备 {len(virtual)}")
    print(f"    real   : {real}")
    print(f"    virtual: {virtual}")
    assert real, "没有任何真实打印机（本机装的是纯虚拟打印机？）"
    for v in virtual:
        assert v not in real, f"虚拟设备 {v!r} 未被过滤"
    # 这几个是键盘上的"常见坑"，必须被识别为虚拟设备
    for must_filter in ("Microsoft Print to PDF", "导出为WPS PDF",
                        "Microsoft XPS Document Writer", "Fax"):
        hit = must_filter.lower() in [v.lower() for v in all_p]
        if hit:
            assert must_filter not in real, f"{must_filter} 必须被过滤掉"

    default = po.default_printer()
    print(f"[OK] 系统默认打印机：{default!r}（仅作参考，不再被测试直接使用）")

    # ========== 3. OpDef 注册
    assert "pdf_print" in OPS, "OpDef pdf_print 未注册"
    assert "pdf_export" in OPS, "OpDef pdf_export 未注册"
    op_print = OPS["pdf_print"]
    op_export = OPS["pdf_export"]
    assert op_print.hardware, "pdf_print 必须标记 hardware=True（物理动作闸门依赖它）"
    assert op_print.confirm, "hardware 功能必须 confirm=True"
    assert callable(op_print.confirm_builder), "pdf_print 应提供动态确认文案"
    assert any(p.key == "dry_run" for p in op_print.params), "pdf_print 应提供「仅演练不打印」"
    print(f"[OK] pdf_print hardware=True confirm=True，参数 {len(op_print.params)} 项")
    print(f"[OK] pdf_export 参数 {len(op_export.params)} 项")

    fc_print = next(p for p in op_print.params if p.key == "file_configs")
    fc_export = next(p for p in op_export.params if p.key == "file_configs")
    assert len(fc_print.axes) == 5, f"pdf_print axes 应有 5，实际 {len(fc_print.axes)}"
    assert len(fc_export.axes) == 2, f"pdf_export axes 应有 2，实际 {len(fc_export.axes)}"
    print(f"[OK] pdf_print axes={len(fc_print.axes)}, pdf_export axes={len(fc_export.axes)}")

    # ========== 4. DEVMODE 偏移与 PyDEVMODE 交叉校验
    # 我们自己按字节偏移改 DEVMODE，偏移写错会静默改到别的字段（例如把
    # dmColor 写到 dmYResolution 上），所以必须和 pywin32 的解析结果对齐。
    import win32print

    probe = real[0]
    h = win32print.OpenPrinter(probe)
    try:
        py_dm = win32print.GetPrinter(h, 2)["pDevMode"]
        want = (py_dm.PaperSize, py_dm.Orientation, py_dm.Copies, py_dm.Color, py_dm.Duplex)
    finally:
        win32print.ClosePrinter(h)

    raw = po._read_default_devmode_bytes(probe)

    def _i16(buf, off):
        return int.from_bytes(buf[off:off + 2], "little", signed=True)

    got = (_i16(raw, po._OFF_DM_PAPER_SIZE), _i16(raw, po._OFF_DM_ORIENTATION),
           _i16(raw, po._OFF_DM_COPIES), _i16(raw, po._OFF_DM_COLOR),
           _i16(raw, po._OFF_DM_DUPLEX))
    assert got == want, f"DEVMODE 偏移解析不一致：自算 {got} vs pywin32 {want}"
    print(f"[OK] DEVMODE 偏移校验通过：PaperSize/Orientation/Copies/Color/Duplex = {got}")

    raw2 = po._job_devmode_bytes(probe, po.PrintConfig(paper="a5", orientation="landscape"))
    assert _i16(raw2, po._OFF_DM_PAPER_SIZE) == po.PAPER_IDS["a5"], "作业 DEVMODE 纸张未生效"
    assert _i16(raw2, po._OFF_DM_ORIENTATION) == 2, "作业 DEVMODE 方向未生效"
    assert _i16(raw2, po._OFF_DM_COPIES) == 1, "作业 DEVMODE 份数必须固定为 1（份数靠多次作业）"
    raw3 = po._read_default_devmode_bytes(probe)
    assert raw3 == raw, "构造作业 DEVMODE 时污染了系统默认设置！"
    # 关键：构造作业 DEVMODE 绝不能污染系统默认
    h = win32print.OpenPrinter(probe)
    try:
        now = win32print.GetPrinter(h, 2)["pDevMode"]
        assert (now.PaperSize, now.Orientation) == (want[0], want[1]), \
            "构造作业 DEVMODE 时污染了系统默认设置！"
    finally:
        win32print.ClosePrinter(h)
    print("[OK] 作业 DEVMODE 已生效，且未改动系统默认设置")

    # ========== 5. 每作业 DC 纸张生效（只读 GetDeviceCaps）
    import win32con

    dc_a4, ok_a4, note_a4 = po._create_printer_dc(
        probe, po.PrintConfig(paper="a4", orientation="portrait"))
    cap_a4 = (dc_a4.GetDeviceCaps(win32con.PHYSICALWIDTH),
              dc_a4.GetDeviceCaps(win32con.PHYSICALHEIGHT))
    del dc_a4
    dc_a5, ok_a5, note_a5 = po._create_printer_dc(
        probe, po.PrintConfig(paper="a5", orientation="landscape"))
    cap_a5 = (dc_a5.GetDeviceCaps(win32con.PHYSICALWIDTH),
              dc_a5.GetDeviceCaps(win32con.PHYSICALHEIGHT))
    del dc_a5
    print(f"[OK] 每作业 DC：A4纵={cap_a4}（自定义 DEVMODE={ok_a4}）  "
          f"A5横={cap_a5}（自定义 DEVMODE={ok_a5}）")
    # 是否所有真实打印机都吃这套 DEVMODE —— 抽查一台即可，但必须至少一台成功
    any_applied = ok_a4 or ok_a5
    assert any_applied, f"没有任何打印机接受自定义 DEVMODE（降级原因：{note_a4} / {note_a5}）"
    if ok_a4 and ok_a5:
        assert cap_a4 != cap_a5, "作业 DEVMODE 没有真正改变物理纸张尺寸"
        # A4 纵 = 210×297mm，A5 横 = 210×148mm → 宽相同、高约 2:1
        assert abs(cap_a4[0] - cap_a5[0]) <= cap_a4[0] * 0.05, \
            f"A4纵/A5横 宽度应相同（都是 210mm），实得 {cap_a4[0]} vs {cap_a5[0]}"
        assert abs(cap_a4[1] - cap_a5[1] * 2) <= cap_a4[1] * 0.05, \
            f"A4纵/A5横 高度应约 2:1（297mm vs 148mm），实得 {cap_a4[1]} vs {cap_a5[1]}"
        print("[OK] A4 纵向与 A5 横向的物理尺寸符合预期（宽同为 210mm、高约 2:1），"
              "per-job DEVMODE 真的生效")
    else:
        print(f"[警告] 该驱动退回了默认纸张，降级原因：{note_a4 or note_a5}")

    # ========== 6. 授权闸门：无人值守路径必须被拒绝
    sandbox = os.path.join(ROOT, "_printtest_full", f"run-{time.strftime('%Y%m%d-%H%M%S')}")
    os.makedirs(sandbox, exist_ok=True)
    pdfs = _make_pdfs(sandbox)
    print(f"[OK] 测试 PDF 3 份：{sandbox}")

    from pypdf import PdfReader

    _ = [PdfReader(p).pages for p in pdfs]  # 读得动即可

    base_params = {
        "printer": probe,
        "page_range": "",
        "scaling": "fit",
        "scale_pct": 100,
        "n_up": 1,
        "dpi": 96,
        "dry_run": False,
        "file_configs": {
            "default": {"paper": "A4", "orientation": "portrait", "copies": 1,
                        "duplex": "none", "color": "color"},
            "overrides": {"1": {"paper": "A5", "orientation": "landscape", "copies": 2},
                          "2": {"paper": "A3", "orientation": "landscape", "copies": 1}},
        },
    }

    ctx_deny = make_ctx(op_print, os.path.join(sandbox, "deny_out"))  # allow_real_print 默认 False
    try:
        po._do_print(ctx_deny, pdfs, dict(base_params))
    except PrintNotAuthorized as exc:
        print(f"[OK] 未授权打印被拒绝：{str(exc).splitlines()[0][:70]}…")
    else:
        raise AssertionError("未授权的打印竟然被执行了 —— 本次事故的回归防线失效！")

    # ========== 4.5 没有真实打印机时，打印功能必须"还在"，且讲清楚为什么打不了
    # 1.0.2 在这里写了 `if not printers: return None`，导致只有虚拟打印设备的
    # 机器上整个「PDF 批量打印」按钮从侧栏消失 —— 真实用户反馈：
    # 「我看打印功能怎么不见了」。这条用例把该场景永久钉死。
    _orig_lp = (po.list_printers, po.list_virtual_pdf_printers, po.default_printer)
    try:
        po.list_printers = lambda real_only=True: []          # 一台真打印机都没有
        po.list_virtual_pdf_printers = lambda: ["Microsoft Print to PDF"]
        po.default_printer = lambda: "Microsoft Print to PDF"

        op_nopr = po._build_print_op()
        assert op_nopr is not None, (
            "没有真实打印机时 _build_print_op 返回 None —— 打印功能整块消失，"
            "这就是「打印功能不见了」的根因")
        assert "未检测到" in op_nopr.title, f"标题应说明现状：{op_nopr.title!r}"
        _pp = next(p for p in op_nopr.params if p.key == "printer")
        assert _pp.default == "" and _pp.choices and _pp.choices[0][1] == "", \
            f"打印机参数应退化为占位空值：{(_pp.default, _pp.choices)!r}"
        assert "虚拟" in op_nopr.description, "描述里必须告诉用户虚拟设备为什么被隐藏"
        print(f"[OK] 无真实打印机：功能仍在，标题 = {op_nopr.title!r}")

        # 执行路径：非演练 + 未选打印机 → 必须给出"怎么办"的明确解释
        ctx_nopr = make_ctx(op_nopr, os.path.join(sandbox, "noprinter_out"),
                            allow_real_print=True)
        try:
            po._do_print(ctx_nopr, pdfs[:1], {"printer": "", "dry_run": False})
        except RuntimeError as exc:
            assert ("虚拟" in str(exc) and "导出" in str(exc)), \
                f"报错文案应同时说明原因与出路：{exc}"
            print(f"[OK] 无真实打印机执行被拦下并解释：{str(exc).splitlines()[0][:60]}…")
        else:
            raise AssertionError("机器上没有真实打印机，竟然没拦住打印请求！")

        # 防线：脚本/旧配置塞进来一台虚拟设备名 → 点名拒绝，绝不静默出 PDF
        try:
            po._do_print(make_ctx(op_print, os.path.join(sandbox, "virtual_out"),
                                  allow_real_print=True),
                         pdfs[:1],
                         {"printer": "Microsoft Print to PDF", "dry_run": False})
        except ValueError as exc:
            assert "虚拟" in str(exc), f"应点名这是虚拟设备：{exc}"
            print("[OK] 虚拟设备被点名拒绝（脚本传参也拦得住）")
        else:
            raise AssertionError("把 Microsoft Print to PDF 传进去竟然没拦 —— 会复刻 1.0.1 事故！")
    finally:
        (po.list_printers,
         po.list_virtual_pdf_printers,
         po.default_printer) = _orig_lp

    # 授权 + 演练模式 → 允许执行，但仍然不出纸
    ctx_dry = make_ctx(op_print, os.path.join(sandbox, "dry_out"),
                       allow_real_print=True, dry_run=True)
    po._do_print(ctx_dry, pdfs, dict(base_params))
    dry_files = sorted(os.listdir(ctx_dry.out_dir))
    assert ctx_dry.report.ok == 3, f"演练应成功 3 个，实际 {ctx_dry.report.ok}"
    assert len(dry_files) == 3, f"演练应产出 3 份预览 PDF，实际 {dry_files}"
    assert all("预览" in f for f in dry_files), f"预览文件命名异常：{dry_files}"
    # 演练产物纸张应按各自配置：a4→A4纵(595x842)，a5→A5横(595x420)，a3→A3横(1191x842)
    import pymupdf

    expect = {"a4": (595, 842), "a5": (595, 420), "a3": (1191, 842)}
    for f in dry_files:
        doc = pymupdf.open(os.path.join(ctx_dry.out_dir, f))
        w_pt, h_pt = doc[0].rect.width, doc[0].rect.height
        doc.close()
        key = f[:2].lower()
        ew, eh = expect[key]
        assert abs(w_pt - ew) < 6 and abs(h_pt - eh) < 6, \
            f"{f} 应为 {key.upper()} 期望 {ew}x{eh}pt，实得 {w_pt:.0f}x{h_pt:.0f}"
    print("[OK] 演练模式：产出 3 份预览 PDF，纸张 A4纵 / A5横 / A3横 全部正确，"
          "且未向打印机提交任何作业")
    print(f"[OK] 演练指标：{dict(ctx_dry.report.metrics)}")

    # 参数级演练（界面勾选"仅演练不打印"走的就是这条）
    ctx_dry2 = make_ctx(op_print, os.path.join(sandbox, "dry2_out"), allow_real_print=True)
    p2 = dict(base_params)
    p2["dry_run"] = True
    po._do_print(ctx_dry2, pdfs, p2)
    assert ctx_dry2.report.ok == 3
    print("[OK] params.dry_run=True 同样不会出纸")

    # ========== 7. 导出 PDF（不合并）
    out_dir2 = os.path.join(sandbox, "export_out")
    ctx2 = make_ctx(op_export, out_dir2)
    po._do_export_pdf(ctx2, pdfs, {
        "merge_all": False, "n_up": 1, "fit_mode": "fit", "dpi": 96,
        "file_configs": {
            "default": {"paper": "A4", "orientation": "portrait", "copies": 1,
                        "duplex": "none", "color": "color"},
            "overrides": {"1": {"paper": "A5", "orientation": "landscape"}},
        },
    })
    assert ctx2.report.ok == 3, f"导出应成功 3 个，实际 {ctx2.report.ok}"
    for f in sorted(os.listdir(out_dir2)):
        doc = pymupdf.open(os.path.join(out_dir2, f))
        w_pt, h_pt = doc[0].rect.width, doc[0].rect.height
        doc.close()
        if "a5" in f:
            assert abs(w_pt - 595) < 5 and abs(h_pt - 420) < 5, f"{f} 应为 A5 横"
        else:
            assert abs(w_pt - 595) < 5 and abs(h_pt - 842) < 5, f"{f} 应为默认 A4 纵"
    print("[OK] 导出 PDF（不合并）：3 份，纸张 A4纵 / A5横 / 默认 A4纵 全部正确")

    # ========== 8. 导出 PDF（合并 + 拼版）
    out_dir3 = os.path.join(sandbox, "merge_out")
    ctx3 = make_ctx(op_export, out_dir3)
    po._do_export_pdf(ctx3, pdfs, {
        "merge_all": True, "n_up": 2, "fit_mode": "fit", "dpi": 96,
        "file_configs": {
            "default": {"paper": "A4", "orientation": "portrait", "copies": 1,
                        "duplex": "none", "color": "color"},
            "overrides": {"1": {"paper": "A5", "orientation": "landscape"}},
        },
    })
    assert ctx3.report.ok == 1
    merged = os.path.join(out_dir3, "merged.pdf")
    assert os.path.exists(merged), f"合并文件未生成：{os.listdir(out_dir3)}"
    doc = pymupdf.open(merged)
    n_pages = len(doc)
    doc.close()
    assert n_pages == 3, f"合并 PDF 应 3 页（n_up=2 拼版后），实得 {n_pages}"
    print(f"[OK] 合并导出 PDF：单文件 {n_pages} 页（n_up=2 拼版后）")

    # ========== 9. GUI ParamPanel 构建 + 覆盖机制
    from app.ui.parampanel import ParamPanel
    from app.ui.theme import apply_theme

    fake_files = ['quarterly.pdf', 'manual_a5.pdf', 'big_a3.pdf']

    def get_files():
        return list(fake_files)

    root = tk.Tk()
    apply_theme(root)
    pp = ParamPanel(root, get_files=get_files)
    pp.pack(fill="both", expand=True)
    pp.build(op_print)

    vals = pp.get_values()
    fc = vals["file_configs"]
    assert fc["default"] == {"paper": "A4", "orientation": "portrait",
                             "copies": 1, "duplex": "none", "color": "color"}, \
        f"default 配置错位：{fc['default']}"
    assert fc["overrides"] == {}, f"应有空覆盖，实得 {fc['overrides']}"
    assert vals["dry_run"] is False, "默认应为非演练（要真打必须用户主动勾选演练）"
    print("[OK] GUI default 配置 label→value 还原正确，dry_run 默认 False")

    container = pp._per_file_containers["file_configs"]
    df = container.winfo_children()[0]
    df._overrides["data"]["1"] = {"paper": "A5", "orientation": "landscape", "copies": 3}
    container.refresh_files()
    vals2 = pp.get_values()["file_configs"]
    assert vals2["overrides"]["1"]["paper"] == "A5"
    assert vals2["overrides"]["1"]["orientation"] == "landscape"
    assert vals2["overrides"]["1"]["copies"] == 3
    print("[OK] 行内覆盖可设置 + 取值正确")

    # ========== 9b. 确认文案必须把真实动作讲清楚
    text_dry = op_print.confirm_builder(fake_files, {"printer": probe, "dry_run": True,
                                                    "file_configs": fc})
    text_real = op_print.confirm_builder(fake_files, {"printer": probe, "dry_run": False,
                                                     "file_configs": fc})
    assert "演练" in text_dry and probe not in text_dry.split("打印机")[0]
    assert "真实打印" in text_real and probe in text_real
    print("[OK] 确认文案：演练与真实打印的措辞不同，真实打印会点名打印机")

    fake_files.append("extra.pdf")
    pp.refresh_files()
    print("[OK] 文件列表变更同步 OK")
    root.destroy()

    # ========== 10. 引擎层：默认 TaskRequest 也必须被拦住
    # 这条守的是"任何脚本/自检/定时任务直接用引擎跑打印"的场景 ——
    # TaskRequest 的默认值就是 False，所以它们天生拿不到授权。
    from app.engine import TaskManager, TaskRequest

    req = TaskRequest(op_key="pdf_print", files=pdfs[:1],
                      params={"printer": probe, "dry_run": False,
                              "file_configs": fc},
                      output_parent=sandbox)
    assert req.allow_real_print is False, "TaskRequest.allow_real_print 默认必须是 False"
    mgr = TaskManager()
    mgr.start(req)
    mgr.join(timeout=60)
    assert not mgr.running, "任务应已结束"
    done = mgr.bus.poll(limit=200)
    reports = [m[1] for m in done if m[0] == "done"]
    assert reports, "应收到 done 消息"
    rep = reports[0]
    assert rep.ok == 0 and rep.failed == 1, \
        f"未授权打印应被拒绝（ok=0 failed=1），实际 ok={rep.ok} failed={rep.failed}"
    reason = rep.failures[0].reason if rep.failures else ""
    assert "授权" in reason or "确认" in reason, f"拒绝原因不明确：{reason!r}"
    print(f"[OK] 引擎层：默认请求被拒绝且给出明确原因 —— {reason.splitlines()[0][:50]}…")

    # ========== 11. 【可选】真实打印一张
    print("-" * 70)
    if not real_print:
        print("[跳过] 真实打印（默认不做）。如确需验证出纸，请显式指定：")
        print("        ODA_TEST_REAL_PRINT=1 ODA_TEST_PRINTER=\"<打印机名>\" "
              "python tests/test_print.py")
    elif not target_printer:
        print("[跳过] 设了 ODA_TEST_REAL_PRINT=1 但没给 ODA_TEST_PRINTER，"
              "为不误用到系统默认打印机，仍然跳过。")
    else:
        assert target_printer in po.list_printers(), \
            f"{target_printer!r} 不在真实打印机列表中"
        print(f"!! 即将向 {target_printer!r} 真实打印 1 页……")
        ctx_real = make_ctx(op_print, os.path.join(sandbox, "real_out"),
                            allow_real_print=True)
        p_real = dict(base_params)
        p_real["printer"] = target_printer
        p_real["dpi"] = 96
        p_real["file_configs"] = {
            "default": {"paper": "A4", "orientation": "portrait", "copies": 1,
                        "duplex": "none", "color": "color"},
            "overrides": {},
        }
        po._do_print(ctx_real, pdfs[:1], p_real)
        assert ctx_real.report.ok == 1, f"真实打印失败：{ctx_real.report.failures}"
        print("[OK] 真实打印 1 份完成（请到打印机取纸）")

    print("\n" + "=" * 70)
    print("全部 PDF 打印 + 导出验证通过")
    print(f"日志目录：{sandbox}")
    print("=" * 70)
    return 0


def _make_pdfs(sandbox: str) -> list[str]:
    from reportlab.pdfgen import canvas as rl_canvas

    pdfs = []
    for name in ("a4.pdf", "a5.pdf", "a3.pdf"):
        p = os.path.join(sandbox, name)
        c = rl_canvas.Canvas(p)
        for i in range(2):
            c.drawString(72, 700, f"{name} page {i + 1}")
            c.showPage()
        c.save()
        pdfs.append(p)
    return pdfs


if __name__ == "__main__":
    sys.exit(main())
