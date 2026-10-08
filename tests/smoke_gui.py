"""GUI 构建冒烟测试。

不进入 mainloop，而是手动构造整个界面并把 20 个功能全部切换一遍 ——
这样能覆盖所有参数面板的控件生成代码（Tkinter 的拼写错误只有真正创建控件时才会暴露）。

用法：python tests/smoke_gui.py
"""
from __future__ import annotations

import os
import sys
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

failures: list[str] = []


def step(label: str, fn):
    try:
        fn()
        print(f"  [OK]   {label}")
        return True
    except Exception as exc:
        failures.append(f"{label}: {type(exc).__name__}: {exc}")
        print(f"  [FAIL] {label}: {type(exc).__name__}: {exc}")
        traceback.print_exc()
        return False


def wait_manager_idle(win, root, timeout: float = 60.0) -> bool:
    """等上一个用例的工作线程彻底退出。

    为什么必须显式等：`manager.running` 是**全局**状态，而取消类用例结束后
    工作线程还要跑一小段收尾（取消检查点 → 释放 DC → 置位结束标志）。
    如果下一个用例直接断言 `not manager.running`，就会变成"依赖上一个用例
    的线程调度速度"的竞态 —— 机器空闲时过、有负载时随机失败。
    这里显式把前置条件等干净，断言才有意义（也才稳定）。
    """
    import time as _time

    deadline = _time.time() + timeout
    while win.manager.running and _time.time() < deadline:
        root.update()
        _time.sleep(0.02)
    # 再抽几帧，让消息泵把上个任务的 done 消息处理掉
    for _ in range(5):
        root.update()
        _time.sleep(0.01)
    return not win.manager.running


def main() -> int:
    import tkinter as tk

    print("=" * 72)
    print("GUI 构建冒烟测试")
    print("=" * 72)

    try:
        from app.ui.app import MainWindow, create_root
    except Exception:
        print("无法导入界面模块：")
        traceback.print_exc()
        return 2

    try:
        root, dnd = create_root()
        print(f"  [OK]   根窗口创建成功（拖拽支持：{'可用' if dnd else '不可用'}）")
    except Exception as exc:
        print(f"  [FAIL] 无法创建 Tk 根窗口（可能是无图形环境）：{exc}")
        print("\n  → 本机没有可用的图形显示，跳过 GUI 测试。这不代表代码有问题。")
        return 0

    root.withdraw()  # 不弹窗干扰
    win = None

    def build():
        nonlocal win
        win = MainWindow(root, dnd_available=dnd)
        root.update_idletasks()
        root.update()

    if not step("构造主窗口", build):
        root.destroy()
        return 1

    from app.core import OPS

    print(f"\n  遍历全部 {len(OPS)} 个功能（构建对应参数面板 + 取值 + 校验）：")
    for key, op in OPS.items():
        def switch(k=key):
            win._select_op(k)
            root.update_idletasks()
            root.update()
            values = win.param_panel.get_values()
            assert isinstance(values, dict), "get_values 必须返回 dict"
            win.param_panel.validate()
            # 依赖缺失时按钮应为禁用，依赖齐全且无文件时也应禁用（没有文件）
            assert str(win.run_btn["state"]) == "disabled", \
                f"没有文件时执行按钮应为禁用，实际为 {win.run_btn['state']}"
        step(f"{op.category}/{key}（{op.title}）", switch)

    # 文件列表交互
    def filelist_ops():
        import tempfile

        d = tempfile.mkdtemp(prefix="oda_smoke_")
        paths = []
        for i in range(4):
            p = os.path.join(d, f"测试{i + 1}.pdf")
            with open(p, "wb") as fh:
                fh.write(b"%PDF-1.4\n%%EOF\n")
            paths.append(p)
        win.filelist.set_filter(("pdf",), "PDF 文件")
        win.filelist.add_files(paths)
        root.update()
        assert win.filelist.count() == 4, f"应加入 4 个文件，实际 {win.filelist.count()}"
        # 去重
        win.filelist.add_files(paths)
        root.update()
        assert win.filelist.count() == 4, "重复添加应被去重"
        # 排序
        before = win.filelist.get_files()
        win.filelist.tree.selection_set(win.filelist.tree.get_children()[0])
        win.filelist.move_selected(1)
        root.update()
        after = win.filelist.get_files()
        assert after[0] != before[0] or len(before) < 2, "下移应改变顺序"
        assert set(after) == set(before), "排序不应丢失文件"
        # 类型过滤
        win.filelist.set_filter(("image",), "图片")
        win.filelist.add_files([os.path.join(d, "x.txt")])
        assert win.filelist.count() == 4, "类型不符的文件不应被加入"

    step("文件列表：添加/去重/排序/类型过滤", filelist_ops)

    def rename_preview():
        from app.core.rename_ops import build_preview

        files = win.filelist.get_files()
        win._select_op("batch_rename")
        rows = build_preview(files, win.param_panel.get_values())
        assert len(rows) == len(files), "预览行数应与文件数一致"
        assert all(r.new_name.endswith(".pdf") for r in rows), "扩展名应保留"

    step("重命名预览计算", rename_preview)

    def param_roundtrip():
        win._select_op("pdf_page_number")
        root.update()
        vals = win.param_panel.get_values()
        assert vals.get("pnum_format"), "页码格式应有默认值"
        assert isinstance(vals.get("skip_first"), int), "skip_first 应转成 int"
        assert isinstance(vals.get("margin_mm"), int), "margin_mm 应转成 int"

        win._select_op("pdf_watermark")
        root.update()
        vals = win.param_panel.get_values()
        assert isinstance(vals.get("wm_opacity"), float), "透明度应转成 float"
        assert vals.get("wm_angle") in ("0", "45", "-45", "90")

        win._select_op("pdf_compress")
        root.update()
        vals = win.param_panel.get_values()
        assert vals.get("level") in ("light", "medium", "strong"), \
            f"下拉框应还原成内部值，实际 {vals.get('level')!r}（说明中文标签映射有问题）"

    step("参数取值与类型转换（含下拉框中文标签还原）", param_roundtrip)

    def visibility_toggle():
        win._select_op("batch_rename")
        root.update()
        # 默认 original_seq：应看到 seq_sep；find_text 应隐藏
        panel = win.param_panel
        assert "seq_sep" in panel._rows, "seq_sep 行应存在"
        assert panel._rows["seq_sep"].winfo_ismapped(), "original_seq 模式下 seq_sep 应可见"
        assert not panel._rows["find_text"].winfo_ismapped(), "original_seq 模式下 find_text 应隐藏"
        # 切到 find_replace
        cb = panel._widgets["name_mode"]
        cb.set("查找替换")
        panel._apply_visibility()
        root.update()
        assert panel._rows["find_text"].winfo_ismapped(), "find_replace 模式下 find_text 应可见"
        assert not panel._rows["seq_sep"].winfo_ismapped(), "find_replace 模式下 seq_sep 应隐藏"

    step("参数联动显隐（visible_if）", visibility_toggle)

    def validation():
        win._select_op("pdf_extract")
        root.update()
        win.param_panel._vars["pages"].set("abc")
        msg = win.param_panel.validate()
        assert msg, "非法页码范围应被校验拦截"
        win.param_panel._vars["pages"].set("1-3,5")
        assert win.param_panel.validate() == "", "合法页码范围不应报错"

        win._select_op("pdf_page_number")
        root.update()
        win.param_panel._vars["font_size"].set("999")
        assert win.param_panel.validate(), "超出上限的字号应被拦截"
        win.param_panel._vars["font_size"].set("12")
        assert win.param_panel.validate() == ""

    step("参数校验（页码格式 / 数值范围）", validation)

    def dialogs():
        from app.errors import Failure, TaskReport

        rep = TaskReport(op_title="测试任务", output_dir=os.path.expanduser("~"),
                         ok=3, failed=1, duration=1.23)
        rep.metrics["压缩前"] = "1 MB"
        rep.failures.append(Failure("C:/x/坏文件.pdf", "PDF 结构损坏", "error"))
        from app.ui.app import _RenamePreviewDialog, _ResultDialog

        d1 = _ResultDialog(root, rep, lambda: None)
        root.update()
        d1.destroy()
        # 重命名预览对话框
        from app.core.rename_ops import RenameRow

        rows = [RenameRow("C:/a/1.pdf", "成品_1_01.pdf", "ok"),
                RenameRow("C:/a/2.pdf", "成品_1_01.pdf", "conflict", "批次内出现重名")]
        d2 = _RenamePreviewDialog(root, rows)
        root.update()
        d2.destroy()

    step("结果对话框 / 重命名预览对话框", dialogs)

    def settings_window():
        from app.ui.settings_win import SettingsWindow

        sw = SettingsWindow(root, on_saved=lambda s: None)
        root.update()
        sw.destroy()

    step("设置窗口", settings_window)

    def resize():
        for geo in ("1020x700", "1400x900", "1100x760"):
            root.geometry(geo)
            root.update_idletasks()
            root.update()
        # 切换若干功能后再强制布局，验证没有 geometry manager 冲突
        for k in ("pdf_merge", "image_to_pdf", "doc_compare", "batch_rename"):
            win._select_op(k)
            root.update()
        assert win.run_btn.winfo_exists(), "布局切换后控件应仍然存在"

    step("窗口尺寸变化与功能切换（布局稳定性）", resize)

    def async_task_end_to_end():
        """真实跑一次异步任务，验证「工作线程 → 队列 → 主线程」整条链路。

        这是对跨线程安全性的正面验证：日志由工作线程产生，必须全部经由队列
        到达主线程；若哪里误用了 root.after 直调，这里会抛
        RuntimeError: main thread is not in main loop。
        """
        import io
        import tempfile
        import time as _time

        from pypdf import PdfReader
        from reportlab.pdfgen import canvas

        d = tempfile.mkdtemp(prefix="oda_async_")
        src = os.path.join(d, "待旋转.pdf")
        b = io.BytesIO()
        c = canvas.Canvas(b)
        for i in range(3):
            c.drawString(72, 720, f"page {i + 1}")
            c.showPage()
        c.save()
        with open(src, "wb") as fh:
            fh.write(b.getvalue())

        win.filelist.set_filter(("pdf",), "PDF 文件")
        win.filelist.set_files([src])
        win._select_op("pdf_rotate")
        root.update()
        win.param_panel._vars["angle"].set("顺时针 90°")
        win.param_panel._widgets["angle"].set("顺时针 90°")
        root.update()

        win._start_task()
        deadline = _time.time() + 60
        while win._busy and _time.time() < deadline:
            root.update()
            _time.sleep(0.02)
        # 收尾：让消息泵把 done 消息处理掉
        for _ in range(10):
            root.update()
            _time.sleep(0.02)

        assert not win._busy, "任务应在超时前结束"
        rep = win._last_report
        assert rep is not None, "应收到结果报告"
        assert rep.ok == 1 and rep.failed == 0, f"任务应成功 1 个，实际 {rep.ok}/{rep.failed}"

        out = os.path.join(rep.output_dir, os.listdir(rep.output_dir)[0])
        assert PdfReader(out).pages[0].rotation == 90, "输出 PDF 应已旋转 90°"

        # 关闭自动弹出的结果对话框，避免影响后续测试
        from app.ui.app import _ResultDialog

        for child in list(root.winfo_children()):
            if isinstance(child, _ResultDialog):
                child.grab_release()
                child.destroy()
        root.update()

        log_text = win.log_text.get("1.0", "end")
        assert "开始执行" in log_text, "主线程日志区应收到工作线程的日志"

    step("异步任务端到端（线程/队列/取消链路）", async_task_end_to_end)

    def print_dry_run_end_to_end():
        """打印功能的界面链路：必须"确认了才跑"，且演练模式一张纸都不出。

        这条用例守的是 1.0.1 那次事故：
          · 用户没确认 → 必须什么都不发生；
          · 确认了但勾了演练 → 只出预览 PDF，不碰打印机。
        所以下方**绝不能**出现"真的打印"的用例 —— 自动化测试永远不该出纸。
        """
        import io
        import tempfile
        import time as _time

        from tkinter import messagebox

        from app.core.registry import OPS
        from reportlab.pdfgen import canvas

        if "pdf_print" not in OPS:
            print("        （本机没有真实打印机，跳过打印界面链路）")
            return

        d = tempfile.mkdtemp(prefix="oda_print_gui_")
        pdfs = []
        for name in ("甲.pdf", "乙.pdf"):
            p = os.path.join(d, name)
            b = io.BytesIO()
            c = canvas.Canvas(b)
            c.drawString(72, 720, name)
            c.showPage()
            c.save()
            with open(p, "wb") as fh:
                fh.write(b.getvalue())
            pdfs.append(p)

        win.filelist.set_filter(("pdf",), "PDF 文件")
        win.filelist.set_files(pdfs)
        win._select_op("pdf_print")
        root.update()

        # ---- 0) 前置条件：先等上一个用例（取消链路）的线程彻底退干净 ----
        # 否则下面那句 `assert not manager.running` 测的其实是"上一个用例跑得够快吗"，
        # 而不是"点「否」之后有没有启动任务" —— 这是真实出现过的随机失败。
        assert wait_manager_idle(win, root), \
            "前置条件不满足：上一个用例结束后仍有工作线程在跑"

        # ---- 1) 用户在确认框里点"否" → 必须什么都不发生
        seen = {}

        def _no(title, msg, **kw):
            seen["q"] = msg
            return False

        orig = messagebox.askyesno
        messagebox.askyesno = _no
        try:
            win._start_task()
            root.update()
            assert not win._busy, "用户点了「否」，任务不该启动"
            assert not win.manager.running, "用户点了「否」，不该有工作线程在跑"
            assert "打印机" in seen.get("q", ""), \
                "确认框必须点名真实打印机，而不是笼统问一句"
            assert "份" in seen.get("q", ""), "确认框必须给出合计份数/张数"
            print("        （用户点「否」→ 未启动任务，符合预期）")

            # ---- 2) 点"是" + 勾选演练 → 只出预览 PDF
            win.param_panel._vars["dry_run"].set(True)
            root.update()
            messagebox.askyesno = lambda *a, **k: True
            win._start_task()
            deadline = _time.time() + 120
            while win._busy and _time.time() < deadline:
                root.update()
                _time.sleep(0.02)
            for _ in range(10):
                root.update()
                _time.sleep(0.02)

            assert not win._busy, "演练任务应在超时前结束"
            rep = win._last_report
            assert rep is not None and rep.ok == 2, \
                f"演练应成功 2 份，实际 {None if rep is None else (rep.ok, rep.failed)}"
            out_files = sorted(os.listdir(rep.output_dir))
            assert len(out_files) == 2, f"应产出 2 份预览 PDF，实际 {out_files}"
            assert all("预览" in f for f in out_files), f"命名异常：{out_files}"
            assert "演练" in rep.metrics.get("模式", ""), \
                f"结果里应写明这是演练：{rep.metrics}"
            print(f"        （演练产出：{out_files}）")

            from app.ui.app import _ResultDialog

            for child in list(root.winfo_children()):
                if isinstance(child, _ResultDialog):
                    child.grab_release()
                    child.destroy()
            root.update()

            # 收尾：把 dry_run 恢复默认，避免影响其它用例
            win.param_panel._vars["dry_run"].set(False)
        finally:
            messagebox.askyesno = orig

    step("打印：用户未确认不执行 + 演练模式不出纸", print_dry_run_end_to_end)

    root.destroy()

    print("-" * 72)
    if failures:
        print(f"失败 {len(failures)} 项：")
        for f in failures:
            print("  ·", f)
        print("=" * 72)
        return 1
    print("全部通过")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    sys.exit(main())
