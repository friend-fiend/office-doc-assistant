"""针对两个高风险点的专项测试：Tkinter 线程安全 + 环境守卫。

为什么单独一个文件
------------------
这两个问题的共同点是 **不会在功能自检里暴露**：
  · 线程问题只在真实并发下偶发（随机崩溃 / 随机卡死）
  · Tkinter 缺失只会让程序"启动不起来"，而自检根本不需要界面

所以必须用专门的、针对性的测试把边界条件钉死。

用法：python tests/threading_and_guard.py
退出码：0 = 全部通过
"""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import time
import traceback

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _ROOT)

failures: list[str] = []
checks = 0


def check(label: str, fn) -> bool:
    global checks
    checks += 1
    try:
        fn()
        print(f"  [OK]   {label}")
        return True
    except Exception as exc:
        failures.append(f"{label}: {type(exc).__name__}: {exc}")
        print(f"  [FAIL] {label}: {type(exc).__name__}: {exc}")
        traceback.print_exc()
        return False


def section(title: str) -> None:
    print(f"\n{'─' * 72}\n{title}\n{'─' * 72}")


# ====================================================================== 工具
def _after_ids(root) -> set:
    """当前挂起的 after 回调 id 集合。Tk 在不同版本下返回 str 或 tuple。"""
    raw = root.tk.call("after", "info")
    if not raw:
        return set()
    if isinstance(raw, str):
        return {raw}
    return set(raw)


def _clear_afters(root) -> None:
    for aid in list(_after_ids(root)):
        try:
            root.after_cancel(aid)
        except Exception:
            pass


def _pump_once(win, root) -> None:
    """跑一次消息泵并清掉它排的下一次，避免回调累积互相干扰。"""
    win._pump()
    _clear_afters(root)


class FakeRunner:
    """假的 TaskRunner，用来模拟"任务还在跑"的状态。"""

    def __init__(self, alive: bool = True):
        self._alive = alive
        self.join_timeouts: list = []

    def is_alive(self) -> bool:
        return self._alive

    def join(self, timeout=None) -> None:
        self.join_timeouts.append(timeout)


# ================================================================ 一、线程安全
def test_threading() -> None:
    section("一、Tkinter 线程安全与消息泵")

    try:
        from app.ui.app import MAX_MSGS_PER_TICK, MainWindow, create_root
        from app.engine import MSG_DONE, MSG_LOG, MSG_PASSWORD, MSG_PROGRESS
    except Exception:
        print("无法导入界面模块：")
        traceback.print_exc()
        return

    try:
        root, dnd = create_root()
    except Exception as exc:
        print(f"  [SKIP] 无图形环境，跳过（{exc}）")
        return

    root.withdraw()
    win = MainWindow(root, dnd_available=dnd)
    root.update_idletasks()
    _clear_afters(root)  # __init__ 里排了 _startup，测试里不需要它

    # ---------------------------------------------------------- 1. after 排程
    def t_pump_reschedules():
        _clear_afters(root)
        win._closing = False
        before = _after_ids(root)
        win._pump()
        after = _after_ids(root)
        _clear_afters(root)
        assert len(after - before) == 1, (
            f"正常运行时消息泵必须重新排程自己，实际新增 {len(after - before)} 个"
        )

    def t_pump_stops_after_close():
        _clear_afters(root)
        win._closing = True
        before = _after_ids(root)
        win._pump()
        after = _after_ids(root)
        assert after == before, (
            "关窗后不得再排程 —— 否则回调会打到已销毁的控件上，"
            f"实际新增 {len(after - before)} 个"
        )

    check("消息泵正常运行时自我重新排程", t_pump_reschedules)
    check("关窗后消息泵停止排程（不会打到已销毁控件）", t_pump_stops_after_close)
    win._closing = False

    # ---------------------------------------------------------- 2. 处理量上限
    def t_tick_limit():
        # 塞 5 倍上限的消息，一个 tick 必须只处理上限条
        total = MAX_MSGS_PER_TICK * 5
        for i in range(total):
            win.manager.bus.queue.put((MSG_LOG, f"洪水消息 {i}", "info"))
        seen = []
        orig = win._log
        win._log = lambda text, level="info": seen.append(text)
        try:
            _pump_once(win, root)
        finally:
            win._log = orig
        assert len(seen) == MAX_MSGS_PER_TICK, (
            f"单 tick 处理量必须是 {MAX_MSGS_PER_TICK}，实际 {len(seen)}"
        )
        # 剩下的必须还在队列里（消息不能丢）
        left = win.manager.bus.queue.qsize()
        assert left == total - MAX_MSGS_PER_TICK, f"剩余消息数不对：{left}"
        # 后续 tick 能排空
        win.manager.bus.queue.queue.clear()

    check(f"单 tick 处理量有上限（{MAX_MSGS_PER_TICK} 条），不会拖死界面", t_tick_limit)

    def t_backlog_no_loss():
        n = MAX_MSGS_PER_TICK * 3
        win.manager.bus.queue.queue.clear()
        for i in range(n):
            win.manager.bus.queue.put((MSG_LOG, f"m{i}", "info"))
        seen = []
        orig = win._log
        win._log = lambda text, level="info": seen.append(text)
        try:
            for _ in range(10):
                _pump_once(win, root)
        finally:
            win._log = orig
        assert len(seen) == n, f"多轮 tick 后应处理完全部 {n} 条，实际 {len(seen)}"

    check("多轮 tick 后消息不丢失（积压最终被排空）", t_backlog_no_loss)

    # ---------------------------------------------------------- 3. 进度合并
    def t_progress_coalesced():
        win.manager.bus.queue.queue.clear()
        n = 300
        for i in range(1, n + 1):
            win.manager.bus.queue.put((MSG_PROGRESS, i, n, ""))

        from tkinter import ttk

        calls: list = []
        orig = ttk.Progressbar.configure

        def spy(self, *a, **k):
            calls.append((a, k))
            return orig(self, *a, **k)

        ttk.Progressbar.configure = spy
        try:
            _pump_once(win, root)
        finally:
            ttk.Progressbar.configure = orig

        assert len(calls) == 1, (
            f"同一 tick 内多条进度消息应合并为一次重绘，实际重绘 {len(calls)} 次"
        )
        got = float(win.progress["value"])
        assert abs(got - 100.0) < 0.01, f"合并后应取最后一条进度（100%），实际 {got}"

    check("进度消息合并为一次重绘，且取最后一条", t_progress_coalesced)

    # ------------------------------------------------- 4. 非主线程写控件的绊线
    def t_tripwire():
        win._ui_queue.queue.clear()
        win.manager.bus.queue.queue.clear()
        before = win.log_text.get("1.0", "end").strip()

        def worker():
            # 故意绕过消息泵直接写日志控件 —— 这正是要拦住的错误用法
            win._append("非法直写", "error")

        th = threading.Thread(target=worker)
        th.start()
        th.join(timeout=5)

        mid = win.log_text.get("1.0", "end").strip()
        assert mid == before, "非主线程不应能直接改动日志控件"
        assert win._ui_queue.qsize() >= 1, "绊线应把消息转入队列"

        _pump_once(win, root)
        after = win.log_text.get("1.0", "end")
        assert "非法直写" in after, "主线程消费队列后应正常显示该条日志"

    check("绊线拦截非主线程直接写控件并自动改走队列", t_tripwire)

    # ------------------------------------------------------- 5. _drain 有上限
    def t_drain_limit():
        from app.ui.app import _drain

        q: "queue.Queue" = queue.Queue()
        for i in range(1000):
            q.put((i,))
        got = _drain(q, limit=10)
        assert len(got) == 10, f"_drain 应最多取 10 条，实际 {len(got)}"
        assert q.qsize() == 990, f"剩余应为 990，实际 {q.qsize()}"
        rest = _drain(q, limit=10_000)
        assert len(rest) == 990

    check("_drain 遵守上限且不丢消息", t_drain_limit)

    root.destroy()


# ================================================================ 二、密码握手
def test_password_handshake() -> None:
    section("二、密码握手：正常流程与超时竞态")

    from app.engine import MSG_PASSWORD, MessageBus

    def t_normal():
        bus = MessageBus()
        box_holder: dict = {}
        result: dict = {}

        def worker():
            result["pwd"] = bus.ask_password("secret.pdf")

        th = threading.Thread(target=worker)
        th.start()

        # 主线程模拟：轮询到请求 -> 用户输入 -> 交付
        msg = bus.queue.get(timeout=5)
        assert msg[0] == MSG_PASSWORD
        assert msg[1] == "secret.pdf"
        box = msg[3]
        box_holder["box"] = box
        bus.deliver_password("hunter2", box=box)
        th.join(timeout=5)
        assert result.get("pwd") == "hunter2", f"应拿到密码，实际 {result.get('pwd')!r}"

    def t_timeout_then_late_password():
        """复现原始缺陷：超时作废后，迟到的密码不得唤醒任何人、也不该弹框。"""
        bus = MessageBus()
        bus.password_timeout = 0.3
        result: dict = {}

        def worker():
            result["pwd"] = bus.ask_password("late.pdf")

        th = threading.Thread(target=worker)
        th.start()
        msg = bus.queue.get(timeout=5)
        box = msg[3]

        th.join(timeout=5)
        assert result["pwd"] is None, "超时应返回 None"
        assert box.get("expired") is True, "超时后 box 必须被标记为作废"

        # 迟到的交付：必须静默忽略，不能抛异常，也不能留下悬挂状态
        bus.deliver_password("too-late", box=box)
        assert bus._pending_password is None, "作废后不应有悬挂的待处理请求"

    check("正常流程：worker 阻塞等待，主线程交付密码后唤醒", t_normal)
    check("超时作废后迟到的密码被忽略（不会弹出孤儿对话框）",
          t_timeout_then_late_password)

    def t_ui_skips_expired():
        """界面侧也要跳过已作废的请求，否则用户会看到"输了没反应"的对话框。"""
        try:
            from app.ui.app import MainWindow, create_root
            from app.engine import MSG_PASSWORD
            import app.ui.app as uiapp
        except Exception:
            raise AssertionError("无法导入界面模块")

        root, dnd = create_root()
        root.withdraw()
        try:
            win = MainWindow(root, dnd_available=dnd)
            _clear_afters(root)

            called = {"dialog": 0}
            orig = uiapp._ask_string

            def fake_dialog(*a, **k):
                called["dialog"] += 1
                return "x"

            uiapp._ask_string = fake_dialog
            try:
                expired_box = {"expired": True, "value": None}
                win._handle_password_request((MSG_PASSWORD, "x.pdf", False, expired_box))
                assert called["dialog"] == 0, "已作废的请求不得弹出对话框"

                live_box = {"expired": False, "value": None}
                win._handle_password_request((MSG_PASSWORD, "x.pdf", False, live_box))
                assert called["dialog"] == 1, "有效请求应正常弹框"
            finally:
                uiapp._ask_string = orig
        finally:
            root.destroy()

    check("界面跳过已作废的密码请求（不弹孤儿对话框）", t_ui_skips_expired)


# ================================================================ 三、关窗流程
def test_close_flow() -> None:
    section("三、关窗流程：等任务收尾，不留下残缺输出")

    import app.ui.app as uiapp
    from app.ui.app import MainWindow, create_root

    def _make_window():
        root, dnd = create_root()
        root.withdraw()
        win = MainWindow(root, dnd_available=dnd)
        _clear_afters(root)
        return root, win

    def t_cancel_then_close():
        root, win = _make_window()
        orig_yesno = uiapp.messagebox.askyesno
        orig_destroy = root.destroy
        calls = {"cancel": 0, "join": [], "destroy": 0}

        try:
            win._busy = True
            win.manager._runner = FakeRunner(alive=False)  # 取消后能及时停下
            win.manager.cancel = lambda: calls.__setitem__("cancel", calls["cancel"] + 1)
            real_join = win.manager.join

            def spy_join(timeout=None):
                calls["join"].append(timeout)
                return real_join(timeout)

            win.manager.join = spy_join
            uiapp.messagebox.askyesno = lambda *a, **k: True  # 首次确认"退出"
            root.destroy = lambda: calls.__setitem__("destroy", calls["destroy"] + 1)

            win._on_close()

            assert calls["cancel"] == 1, "应先发出取消请求"
            assert calls["join"], "必须给工作线程留出收尾时间（join）"
            assert calls["join"][0] and calls["join"][0] > 0, "join 超时必须为正数"
            assert win._closing is True, "关窗时必须先置位 _closing"
            assert calls["destroy"] == 1, "正常路径应销毁窗口"
        finally:
            uiapp.messagebox.askyesno = orig_yesno
            root.destroy = orig_destroy
            try:
                root.destroy()
            except Exception:
                pass

    def t_still_running_user_declines():
        root, win = _make_window()
        orig_yesno = uiapp.messagebox.askyesno
        orig_destroy = root.destroy
        answers = iter([True, False])  # 先同意退出，再拒绝强制退出
        calls = {"destroy": 0}

        try:
            win._busy = True
            win.manager._runner = FakeRunner(alive=True)  # 永远停不下来
            win.manager.cancel = lambda: None
            uiapp.messagebox.askyesno = lambda *a, **k: next(answers)
            root.destroy = lambda: calls.__setitem__("destroy", calls["destroy"] + 1)

            win._on_close()

            assert calls["destroy"] == 0, "用户拒绝强制退出时不得销毁窗口"
            assert win._closing is False, "未退出就不该置位 _closing"
        finally:
            uiapp.messagebox.askyesno = orig_yesno
            root.destroy = orig_destroy
            try:
                root.destroy()
            except Exception:
                pass

    def t_user_declines_first_prompt():
        root, win = _make_window()
        orig_yesno = uiapp.messagebox.askyesno
        orig_destroy = root.destroy
        calls = {"cancel": 0, "destroy": 0}

        try:
            win._busy = True
            win.manager.cancel = lambda: calls.__setitem__("cancel", calls["cancel"] + 1)
            uiapp.messagebox.askyesno = lambda *a, **k: False  # 第一问就取消
            root.destroy = lambda: calls.__setitem__("destroy", calls["destroy"] + 1)

            win._on_close()

            assert calls["cancel"] == 0, "用户取消退出时不该打断正在跑的任务"
            assert calls["destroy"] == 0, "窗口应保持打开"
        finally:
            uiapp.messagebox.askyesno = orig_yesno
            root.destroy = orig_destroy
            try:
                root.destroy()
            except Exception:
                pass

    check("有任务时关窗：先取消 → join 等待 → 再销毁", t_cancel_then_close)
    check("任务停不下来且用户拒绝强退：窗口保持打开", t_still_running_user_declines)
    check("用户在首次确认时选择否：任务不被打断", t_user_declines_first_prompt)


# ============================================================== 四、环境守卫
def test_env_guard() -> None:
    section("四、Tkinter 缺失时的原生报错通道")

    def t_check_tkinter_shape():
        from app.native_alert import check_tkinter

        ok, detail = check_tkinter()
        assert isinstance(ok, bool), "check_tkinter 必须返回 (bool, str)"
        assert isinstance(detail, str) and detail, "诊断信息不能为空"
        if not ok:
            raise AssertionError(f"本环境应能 import tkinter，但报告不可用：{detail}")

    def t_alert_never_raises():
        from app.native_alert import alert, error, warn

        os.environ["ODA_SILENT_ALERT"] = "1"
        try:
            for kind in ("error", "warning", "info", "未知类型"):
                alert("标题", "正文", kind)
            error("t", "x" * 20000)      # 超长必须被截断而不是炸掉
            warn("t", "")                 # 空正文
            alert("", "")                 # 空标题
            alert("中文/emoji 🎉 标题", "换行\n第二行\ttab \"引号\" \\反斜杠")
        finally:
            os.environ.pop("ODA_SILENT_ALERT", None)

    def t_truncate():
        from app.native_alert import _MAX_TEXT, _truncate

        assert _truncate("短") == "短"
        long_text = "字" * (_MAX_TEXT + 500)
        out = _truncate(long_text)
        assert len(out) < len(long_text), "超长文本必须被截断"
        assert "截断" in out, "截断后应说明原因"

    def t_applescript_quoting():
        from app.native_alert import _applescript_quote

        # AppleScript 不支持 \n 转义，换行必须写成字符串拼接
        got = _applescript_quote('a"b\\c\nd')
        assert got.startswith('"') and got.endswith('"'), got
        assert '\\"' in got, f"双引号应被转义：{got}"
        assert "\\\\" in got, f"反斜杠应被转义：{got}"
        assert "& return &" in got, f"换行应转为 AppleScript 拼接：{got}"

    def t_help_texts():
        from app.native_alert import python_where_help, tkinter_help_text

        for txt in (tkinter_help_text(), python_where_help()):
            assert txt.strip(), "帮助文本不能为空"
        assert sys.executable in python_where_help(), "应指出当前解释器路径"

    check("check_tkinter 返回 (bool, 诊断信息)", t_check_tkinter_shape)
    check("原生弹窗在任何输入下都不抛异常", t_alert_never_raises)
    check("超长正文被截断而非崩溃", t_truncate)
    check("AppleScript 字符串转义正确（换行用拼接）", t_applescript_quoting)
    check("帮助文本包含可操作指引", t_help_texts)


# ============================================================== 五、体检脚本
def test_preflight() -> None:
    section("五、环境体检脚本（--check）")

    def t_report_logic():
        from tools.preflight import FAIL, OK, WARN, Report

        r = Report()
        r.add("硬性项", OK, hard=True)
        r.add("软性项", WARN)
        assert not r.hard_failed, "只有 WARN 时不应算硬性失败"

        r2 = Report()
        r2.add("硬性项", FAIL, hard=True)
        assert len(r2.hard_failed) == 1, "硬性 FAIL 必须被识别"
        text = r2.render()
        # 结论文案不再写死"程序无法启动"：硬性缺失里既有真的起不来的（Tkinter），
        # 也有"能起来但关键功能废掉"的（PyMuPDF 渲染引擎）。
        # 断言改成校验更有用的性质：结论必须点名是哪一项、且明确要求修复后才能分发。
        assert "硬性缺失" in text, "有硬性失败时结论应明确"
        assert "硬性项" in text.split("结论：")[-1], "结论必须点名具体缺失项"
        assert "请勿打包分发" in text, "结论必须给出可执行的要求"

        r3 = Report()
        r3.add("一切正常", OK)
        assert "全部功能可用" in r3.render()

    def t_preflight_runs():
        from tools.preflight import main as pf_main

        code = pf_main(verbose=False)
        assert code in (0, 1), f"体检退出码应为 0 或 1，实际 {code}"

    check("体检结论的硬性/软性判定正确", t_report_logic)
    check("体检脚本可正常跑完并返回退出码", t_preflight_runs)


# ========================================================== 六、真实子进程验证
def test_subprocess_guard() -> None:
    section("六、真实启动：缺 Tkinter 时的退出码与提示")

    embedded = sys.executable  # 当前跑测试的解释器（有 tkinter）
    candidates: list[tuple[str, str]] = []
    for label, exe in (
        ("托管 Python 3.13（实测无 tkinter）",
         r"C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe"),
    ):
        if os.path.exists(exe):
            candidates.append((label, exe))

    if not candidates:
        print("  [SKIP] 未找到用于对照的解释器")
        return

    env = dict(os.environ)
    env["ODA_SILENT_ALERT"] = "1"  # 避免原生弹窗阻塞自动化测试

    for label, exe in candidates:
        def run(exe=exe, label=label):
            proc = subprocess.run(
                [exe, "main.py"],
                cwd=_ROOT,
                env=env,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=180,
            )
            out = (proc.stdout or "") + (proc.stderr or "")
            # 无 tkinter 的解释器：必须优雅退出，而不是抛 traceback
            if "No module named 'tkinter'" in out or "No module named \"tkinter\"" in out:
                assert proc.returncode == 2, (
                    f"缺少 Tkinter 时应返回退出码 2，实际 {proc.returncode}"
                )
                assert "缺少 Tkinter" in out, "应给出明确的中文提示"
                assert "Traceback" not in out, (
                    "不应把原始 traceback 直接甩给用户：\n" + out[:1500]
                )
                assert "python.org" in out or "installer" in out, "应给出可操作的解决办法"
            else:
                # 有 tkinter 的解释器：本测试只验证它不会被误判为缺 Tkinter
                assert "缺少 Tkinter" not in out, f"{label} 被误判为缺少 Tkinter"

        check(f"缺 Tkinter 时给出友好提示并返回 2（{label}）", run)

    def t_check_exit_code():
        proc = subprocess.run(
            [embedded, "main.py", "--check"],
            cwd=_ROOT, env=env, capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=300,
        )
        assert "环境体检" in proc.stdout, "--check 应打印体检报告"
        assert proc.returncode in (0, 1), f"退出码异常：{proc.returncode}"

    check("--check 子进程可正常输出体检报告", t_check_exit_code)


# ======================================================================== 入口
def main() -> int:
    print("=" * 72)
    print("线程安全与环境守卫 专项测试")
    print(f"解释器：{sys.executable}")
    print(f"线程：{threading.current_thread().name}")
    print("=" * 72)

    test_threading()
    test_password_handshake()
    test_close_flow()
    test_env_guard()
    test_preflight()
    test_subprocess_guard()

    print("\n" + "=" * 72)
    if failures:
        print(f"存在 {len(failures)} 项失败（共 {checks} 项）：")
        for f in failures:
            print(f"  · {f}")
        print("=" * 72)
        return 1
    print(f"全部通过（共 {checks} 项）")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    sys.exit(main())
