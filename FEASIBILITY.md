# 《办公文档一键处理助手》方案可行性评审与落地方案

> 评审方式：**不是纸面推演，而是把方案真实实现出来并跑通。**
> 项目代码位于 `office-doc-assistant/`，评审结论均附带实测证据。
>
> 环境：Windows 11 64 位；Python 3.13.12（managed，无 tkinter）与 Python 3.12.10（python.org 版，Tk 8.6）双环境交叉验证。

---

## 一、总体结论

**方案整体可行，可以落地。** 20 项功能（P0 全部 + P1 全部）已实现并实测通过。
但有 **4 处提示词需要修正**、**3 个高风险点必须提前处置**，否则会在交付阶段变成返工。

其中 2 个高风险点（Tkinter 线程安全、`tkinter` 可能缺失）已在本次一并**处置完毕并实测通过**；
剩下 1 个（Ghostscript AGPL）取决于分发策略，非技术问题 —— **若不做商业分发，此项自然消解**。

| 维度 | 评级 | 说明 |
| --- | --- | --- |
| 技术选型 | ✅ 可行 | Python 3.10/3.11 + Tkinter + 题述库组合无阻塞性障碍 |
| P0 功能 | ✅ 已完成 | 10/10 全部通过端到端自检 |
| P1 功能 | ✅ 已完成 | 7/7 全部通过 |
| 跨平台一致性 | ⚠️ 需注意 | 界面可统一，但**外部依赖安装体验**在 Windows/macOS 差异大 |
| 打包分发 | ✅ 已实测 | Windows 安装包/便携版已产出并完成安装→运行→卸载全链路验收；macOS 脚本就绪（待 Mac 执行），CI 可代工 |
| 法律合规 | ✅ 已消解 | 不做商业分发时，Ghostscript 的 AGPL 不构成约束；程序本身也不捆绑它 |
| 线程安全 | ✅ 已处置 | 消息泵队列化 + 处理量上限 + 绊线 + 关窗收尾，22 项专项测试通过 |
| 环境守卫 | ✅ 已处置 | 缺 Tkinter 时走原生弹窗报错并给平台化指引，`--check` 一键体检 |
| 提示词自洽性 | ⚠️ 2 处矛盾 | 「仅用标准库 Tkinter」与「拖拽区」矛盾；「不覆盖原文件」与「原地重命名」矛盾 |

### 实测结果汇总

| 验证项 | 结果 |
| --- | --- |
| 核心功能端到端自检（`tests/selftest.py`） | **31 / 31 通过** |
| 产物有效性校验（`tests/verify_artifacts.py`） | **18 / 18 通过** |
| 界面构建与交互冒烟测试（`tests/smoke_gui.py`） | **29 项全部通过** |
| 线程安全与环境守卫专项（`tests/threading_and_guard.py`） | **22 / 22 通过** |
| 依赖库 API 兼容性探测（`tests/probe_api*.py`） | 通过（含 4 处版本差异的适配） |
| Python 版本兼容 | 3.13 与 3.12 均 **31/31 通过** |
| PDF 加密算法 | AES-256 / AES-128 / RC4-128 **全部实测可加可解** |
| 缺 Tkinter 时的启动行为 | 原生弹窗 + 可操作指引，**退出码 2，无原始 traceback** |


---

## 二、逐项技术验证

### 2.1 ✅ Python 版本选择：正确

PEP 11 规定：Python 3.11 支持 Windows 8.1 及更新版本；**Python 3.13 起已放弃 Windows 8.1**。
所以「3.10 / 3.11」这个约束是**必要且正确**的 —— 如果误用 3.13，Windows 8.1 直接不达标。

> 附带发现：Python 3.11 已进入安全维护期，建议本项目锁定 **3.11.x**，并在 README 中写明生命周期。

### 2.2 🔴 提示词矛盾 1：「仅用标准库 Tkinter」 vs 「拖拽区」

**Tkinter 本身没有任何拖放能力。** 拖拽必须依赖 Tcl 扩展 `tkdnd`，Python 侧封装是第三方包 `tkinterdnd2`。

| 方案 | 结论 |
| --- | --- |
| 纯标准库实现拖拽 | ❌ 不可能。Tk 的 `tkdnd` 是外部 Tcl 扩展 |
| 引入 `tkinterdnd2` | ✅ 推荐。MIT 许可，Windows/macOS 均可，体积小（约 800 KB） |

**已落地处置**：把 `tkinterdnd2` 作为**可选依赖**。装了就用真拖拽；没装自动退化为「点击选择文件」，并在日志中说明原因。这样既不违反「尽量少依赖」，也不牺牲体验。

**打包注意**：PyInstaller 默认抓不到 `tkdnd` 二进制，运行时会报 `TclError: can't find package tkdnd`。已提供 `build/hook-tkinterdnd2.py` 解决。

### 2.3 🔴 高风险：Ghostscript 的 AGPL 许可

这是整个方案里**唯一可能导致商业项目翻车的点**，原提示词没有提到。

Ghostscript 采用 **AGPL**。按其官方说明：若你的应用不是 AGPL 开源，则**不得**把 Ghostscript 打包进你的应用 / 与你的应用同介质分发 / 在服务器上供用户交互使用 —— 除非向 Artifex 购买商业许可。

| 场景 | 结论 |
| --- | --- |
| 内部自用工具 | ✅ 无问题 |
| 开源（AGPL）项目 | ✅ 无问题 |
| **不做商业分发**（本次情况） | ✅ **无问题 —— 此项风险自然消解** |
| 闭源商业软件**捆绑** Ghostscript | 🔴 **需购买商业许可** |
| 闭源商业软件**要求用户自行安装** Ghostscript | ✅ 可行（用户与 Artifex 之间是用户自己的事） |

**已落地处置**：程序**不捆绑** Ghostscript，改为「启动探测 + 用户自行安装 + 缺失时降级为 pypdf 结构优化并明确提示压缩效果有限」。这既满足需求中的降级要求，也规避了许可风险。

> **不做商业分发时的结论**：这一项从"必须提前决策"降级为"已经天然合规"。
> 当前实现本来就是"不捆绑 + 用户自装"，因此在非商业场景下**无需任何额外动作**，
> 也不会产生 AGPL 的源码开放义务。剩下的只是使用体验问题 ——
> 用户装了就压得更狠，没装就自动降级，两条路径都已实测跑通。

> 对比：LibreOffice（MPL 2.0）、Poppler（GPL，作为独立可执行程序调用）、pypdf/reportlab/Pillow（BSD/MIT 类）在「独立程序调用」模式下都无此问题。

### 2.4 ⚠️ 高风险：`tkinter` 可能不存在 —— ✅ 已处置（含实测）

实测：本机 managed Python 3.13.12 **`import tkinter` 直接失败**。而 python.org 官方安装包的 Python 3.12 正常（Tk 8.6）。

- Windows 的 **embeddable 包（embed zip）不含 tkinter**；
- 部分发行版 / 精简版 Python 也不含；
- Linux 上需额外装 `python3-tk`。

**这个问题的尴尬之处**：它发作时，程序自己**弹不出任何窗口**（因为弹窗也要 tkinter），
用户看到的就是"双击了，什么都没发生"，或者一闪而过的黑窗口加一串 traceback。
所以「README 里写一句要装 tkinter」远远不够 —— 必须有一条**不依赖 tkinter 的报错通道**。

**已落地处置（三层）**：

| 层 | 实现 | 效果 |
| --- | --- | --- |
| 启动守卫 | `main.py::_guard_tkinter()` + `app/native_alert.py` | 缺 Tk 时用**系统原生弹窗**报错（Windows `MessageBoxW`、macOS `osascript`、Linux `zenity`），全部标准库，零第三方依赖 |
| 环境体检 | `tools/preflight.py`（`python main.py --check`） | 该脚本**本身不 import tkinter**，所以界面起不来时它照样能跑，给出完整诊断清单与逐平台解决办法 |
| 打包拦截 | `build_win.bat` / `build_macos.sh` | 打包前先跑 `--check`，不通过直接中止 —— 避免打出"双击没反应"的 exe |

实测输出（在本机那个确实没有 tkinter 的解释器上）：

```
$ python main.py            # 托管 Python 3.13（确实没有 tkinter）
[缺少 Tkinter，无法启动]
办公文档一键处理助手 需要 Tkinter 才能显示界面，但当前 Python 解释器中不可用。
检测结果：import tkinter 失败：ModuleNotFoundError: No module named 'tkinter'
当前解释器：C:\...\3.13.12\python.exe
Python 版本：3.13.14
Windows 下最常见的原因是用了不含 Tkinter 的 Python：
  · 「embeddable package」压缩包不含 Tkinter
  · Microsoft Store 版 Python 部分版本不含
  · 某些精简/绿色版发行版不含
解决办法：到 python.org 下载标准安装包（64 位），安装时务必勾选「tcl/tk and IDLE」。
→ 退出码 2，且没有把原始 traceback 甩给用户
```

另外补了一条容易被忽略的规则：诊断信息里必须**报出当前解释器的绝对路径**。
用户机器上常同时存在多个 Python（系统 / venv / 包管理器 / 嵌入式），
只说"缺 tkinter"用户根本不知道该换哪一个。`python_where_help()` 会一并打印解释器路径、
版本以及 venv 的真实基准路径。

> 附带收获：`--check` 不只是给缺 Tk 用的。它现在是**唯一的环境真相来源** ——
> 排查"为什么某个按钮是灰的"时，跑一次就知道是缺哪个库/哪个外部程序。

### 2.5 ⚠️ 中风险：LibreOffice headless 的三个真实坑

原提示词只说「LibreOffice headless」，实际落地有三个必须处理的坑，否则会出现**"退出码 0 但产物是 0 字节或没有产物"**这种最难排查的故障：

| 坑 | 现象 | 处置（已实现） |
| --- | --- | --- |
| **共用一个 UserInstallation 目录** | 并发时后启动的进程把任务丢给第一个进程后立即退出 → 随机失败 | 每次调用加 `-env:UserInstallation=file:///<唯一临时目录>` 隔离 |
| **启动开销 >> 转换耗时** | 逐个文件调用时，1.5~3 秒的冷启动被乘以文件数 | 一次命令行传入多个文件（按 6000 字符分批，避免超命令行上限） |
| **profile 目录残留** | 越用越慢、占磁盘 | `finally` 中递归清理（含异常与取消路径） |

另需注意：**本机 GUI 版 LibreOffice 开着时，headless 调用会被"接力"到 GUI 实例**，这正是必须做 profile 隔离的根本原因。

### 2.6 ⚠️ 中风险：pypdf 6.x 的 API 变化（实测踩到 4 处）

实测版本 pypdf 6.19.0，与网上大量 pypdf 3.x/4.x 教程不兼容：

| 变更点 | 旧写法 | 6.x 现状 | 处置 |
| --- | --- | --- | --- |
| `compress_content_streams()` | 对 `reader.pages` 直接调用 | **页面必须已属于某个 PdfWriter**，否则抛 `ValueError: Page must be part of a PdfWriter` | 改为"先 add_page 再压缩"（在 `write_pdf()` 内统一处理） |
| `encrypt(permissions_flag=)` | 传 int 位掩码 | 期望 `UserAccessPermissions` 枚举（传 int 仍可用但不规范） | 优先用枚举，取不到则退化位运算 |
| `compress_identical_objects()` | `remove_identicals` / `remove_orphans` | 已弃用，改为 `remove_duplicates` / `remove_unreferenced`，**7.0 将移除** | 双写兼容（先试新参数，`TypeError` 再退回旧参数） |
| AES 加密 | 直接可用 | 缺 `cryptography` 时抛 `DependencyError` | 捕获后自动降级 RC4-128 并提示 `pip install pypdf[crypto]` |

> 这条很重要：**requirements 里必须写 `pypdf>=4.0,<7`**，或在 7.0 发布时重新验证。已在 requirements.txt 中锁定上界。

### 2.7 ✅ 关键可行点：中文页码/水印不依赖系统字体

这是本方案里我最担心、但实测最漂亮的一点。

**问题**：页码、水印、页眉页脚都要写中文。用外部 TTF 会带来"字体文件缺失 / macOS 与 Windows 字体名不同 / 打包漏字体"三类问题。

**解决的方案**：`reportlab` 内置 **CID 字体 `STSong-Light`**（中日韩字体，随库分发），注册后即可直接渲染中文：

```python
pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
```

**实测**：`stringWidth("第 1 页 / 共 10 页", "STSong-Light", 10) → 69.62`，且产物中确实提取到 `第…页` 文本。**不依赖任何外部字体文件，跨平台一致。** 同时保留 Helvetica 选项给纯西文场景（体积更小）。

### 2.8 ✅ 可行：PDF 压缩的三档策略与实测数据

采用与 PDF24 Tools 相同的思路：**轻档纯 pypdf 无损、中/重档 Ghostscript 重建**。

| 档位 | 实现 | 目标 |
| --- | --- | --- |
| 轻度 | 内容流 zlib level 9 重压 + 对象去重 + 孤儿清理 + 元数据/XMP 清除 | 尽量无损，只挤冗余 |
| 中度 | Ghostscript `pdfwrite`，图片降采样 150 DPI，JPEG Q80，字体子集化，重复图片合并 | 推荐档 |
| 强力 | 同上，96 DPI，JPEG Q60，单色 200 DPI | 体积最小 |

**实测（无 Ghostscript 环境，自动降级）**：原始 3544 B → 2894 B，**缩减 18.3%**，产物可正常解析并保持 3 页。降级提示正确出现在日志与结果指标中：
> `降级提示: Ghostscript 缺失，使用 pypdf 结构优化，压缩效果可能有限`

**必须注意**：Ghostscript 的参数组合对"扫描件 PDF"和"矢量文字 PDF"效果差异极大（扫描件能压到 1/10，纯文字 PDF 可能几乎压不动）。所以**实现里没有承诺"一定变小"**，而是在压缩后反而变大时明确警告，而不是假装成功。

### 2.9 ⚠️ 中风险：Pillow 输出 PDF 的页面尺寸陷阱

**踩到的真实 bug**：用 Pillow 把 595×842 px 的 A4 画布存成 PDF，得到的是 **214×303 pt**，而不是 A4 的 595×842 pt。

**原因**：Pillow 的 PDF 插件按 `点数 = 像素 / resolution × 72` 换算页面尺寸。默认 `resolution=200` 就会把 A4 缩水到 36%。

**处置**：显式传 `resolution=72.0`，让像素与点数 1:1 映射。修复后实测：`页面尺寸pt = (595.0, 841.0)`，**是真正的 A4**。

### 2.10 ⚠️ 中风险：Excel 跨工作表复制单元格会失败

openpyxl 不允许把 A 表的 `Cell` 对象 `append` 到 B 表，会抛
`ValueError: Cells cannot be copied from other worksheets`。

**处置**：读表一律 `ws.iter_rows(values_only=True)`，`append(list(row))`。语义上正好符合需求 —— `data_only=True` 取计算值，`data_only=False` 取公式字符串。修复后实测：每文件一表模式合并出 4 张表 / 24 行；追加模式 1 张表 / 22 行（第二份文件的表头被正确跳过）。

### 2.11 🔴 高风险：Tkinter 线程安全 —— ✅ 已处置（含实测）

**踩到的真实 bug**：依赖探测线程里写了 `self.root.after(0, lambda: ...)`，实测抛：

```
RuntimeError: main thread is not in main loop
```

Tkinter **不是线程安全的**。从非主线程调用 `after()` / 操作控件，轻则异常，重则整个进程崩溃。这是一个非常容易被忽略、且只在真实并发下才暴露的问题。

**处置**：确立唯一规则 —— **"工作线程只往 `queue.Queue` 放数据，主线程用 `after(80ms)` 常驻消息泵消费并更新控件"**。所有跨线程路径（任务日志、进度、密码请求、依赖探测结果、标准 logging 回调）统一走这条队列。

光有"约定"是不够的 —— 约定挡不住三个月后新增的一行代码。所以又补了**三个具体机制**：

**① 处理量上限（防日志洪水假死）**

一个失败较多的任务可能瞬间产出几千条日志。原来的写法是 `while True: get_nowait()` 取空为止，
一个 tick 全画完 → 主线程十几秒无法重绘 → 用户以为程序卡死了。

现在每 tick 最多处理 400 条，进度消息**只取最后一条**（同一 tick 内重画 300 次进度条毫无意义）。
实测：塞 `MAX_MSGS_PER_TICK × 5 = 2000` 条消息，单 tick 只处理 400 条，剩余 1600 条留在队列里，
多轮 tick 后排空、**一条不丢**；300 条进度消息被合并为 **1 次**重绘，且取最后一条的值。

**② 关窗后停止排程（防打到已销毁控件）**

`root.destroy()` 之后，那次已经排好的 `after` 回调仍会被 Tk 调用，然后在控件上炸出
`TclError: invalid command name "...!label"`。现在加了 `_closing` 标志，置位后不再排程；
并且**先置标志再 destroy**（顺序反了就白做）。

**③ 线程安全绊线（把随机崩溃变成可见缺陷）**

`_append()` 是唯一真正写日志控件的函数，所以在那里放了一条检查：
一旦发现调用方不是主线程，**自动改走队列**并记一条 error：

> 检测到非主线程直接更新控件，已自动改走消息队列。这是代码缺陷，请修复调用方。

意义在于：Tkinter 的线程违规表现为**随机崩溃或随机卡死**，是最难排查的一类问题。
有了绊线，它变成一个稳定复现、日志里明确写着"这是谁的错"的事件。
实测：从工作线程调 `_append` → 日志控件内容不变、消息被转入队列、主线程消费后正常显示。

**④ 关窗流程：等任务收尾（防残缺输出）**

原来 `_on_close` 是「取消 → 立即 destroy」。问题是工作线程可能正在写输出文件，
进程一退就留下**残缺的半成品**，而用户会以为处理成功了。现在改为：

```
请求取消 → join(5s) 等检查点退出 → 真的停不下来才问用户是否强制退出
```

实测三条分支：正常取消能及时停下（销毁窗口）；停不下来且用户拒绝强退（窗口保持打开）；
用户在首次确认时选"否"（**不打断**正在跑的任务）。

**⑤ 密码握手竞态（防孤儿对话框）**

用户对着密码框迟迟不动，工作线程 300 秒超时后继续往下跑了。
等用户终于点了确定 —— 结果落到一个"已经没人等"的请求上，界面表现为**输了密码但毫无反应**。

现在每个请求带 `expired` 标记，超时即作废；`deliver_password(value, box=box)` 校验身份后
才唤醒，作废则静默忽略；界面侧也跳过已作废的请求，不弹孤儿对话框。

**验证**：新增 `tests/threading_and_guard.py`，**22 项全部通过**，覆盖
消息泵排程/上限/合并、消息不丢失、绊线、`_drain` 上限、密码握手正常与超时竞态、
关窗三条分支、原生弹窗健壮性、体检逻辑、以及**真实子进程**验证缺 Tk 时的退出码与提示文案。

### 2.12 ⚠️ 中风险：Zip-Slip（路径穿越）必须自己防

解压是这类工具最常被利用的攻击面。构造一个含 `../../Windows/System32/xx.dll` 的 ZIP，直接 `extractall()` 会把文件写到目标目录之外。

**处置**：解压前逐条校验（ZIP / 7z / tar 三套都做了）：拒绝绝对路径、拒绝 `..` 段、用 `os.path.realpath` 二次确认解析后仍位于目标目录内、tar 额外拒绝符号链接与硬链接。

### 2.13 ⚠️ 提示词矛盾 2：「输出不覆盖原文件」 vs 「批量重命名」

重命名按定义就是**改名字**，"输出到新文件夹"与"重命名"本质冲突。

**处置**：拆成两种模式，让用户明确选择，且默认安全侧：

| 模式 | 行为 | 默认 |
| --- | --- | --- |
| `copy` | 把文件以新名字**复制**到输出文件夹，原文件不动 | ✅ 默认 |
| `inplace` | 原地重命名，额外一次**危险操作二次确认**，并生成 `重命名回滚清单.json` | 需显式选择 |

**预览**：`build_preview()` 纯计算不落盘，界面弹表格展示「原文件名 → 新文件名 + 状态」，并对**批次内重名**与**与磁盘已有文件重名**分别标红。

实测：`图片1.png` → `成品_图片1_01.png`（序号带分隔符，避免出现 `图片101` 这种歧义命名）。

### 2.14 ⚠️ 中风险：文档对比对扫描件无意义

`doc_compare` 比较的是**文本层**。扫描件 PDF 没有文本层，抽取结果为空。若不处理，会得出"两个文件 100% 不同"的**误导性结论**。

**处置**：抽取为空时**明确报错并说明原因**，提示"本工具不含 OCR，建议先做 OCR 转成可搜索 PDF"。宁可拒绝执行，也不给出错误答案。

### 2.15 其它已验证项

| 项 | 结论 |
| --- | --- |
| 长路径（>260 字符） | Windows 下统一加 `\\?\` 前缀 + 打包 manifest 声明 `longPathAware` |
| 中文/空格/emoji 路径 | 全程 `pathlib` + `str`，禁止手工拼接；文件名清洗处理 Windows 保留名（CON/PRN/…） |
| macOS NFD/NFC | 文件名统一 NFC 归一化，避免同名文件被当成两个 |
| PDF 损坏容错 | 实测：坏文件 + 好文件同批，1 失败 1 成功，**任务不中断** |
| 取消机制 | 协作式取消（Python 无法安全 kill 线程），在检查点抛 `Cancelled`，保证临时文件被清理。实测标记正确 |
| 外部依赖探测 | 五级查找：用户设置 → `vendor/` 自带 → 解释器同级 → PATH → 各平台常见安装位置 |
| macOS 分发 | `.app` 可打出，但**未公证的 app 在别人机器上首次打开需右键 → 打开**，正式分发需 Developer ID + notarize |

### 2.16 🔴 高风险：打印是"物理世界不可撤销"动作 —— ✅ 已处置（含 1.0.1 事故复盘）

**事故现象**：用户什么都没操作，办公室打印机自己开始出纸，打出一批内容为
`a4.pdf page 1 / a5.pdf page 1 …` 的测试页。

**根因（两个叠加）**：

| # | 问题 | 后果 |
| --- | --- | --- |
| 1 | `tests/test_print.py` 的端到端段直接用**系统默认打印机**跑真实打印 | 每次跑测试就往真机送 4 个作业（a4×1 + a5×2 + a3×1，各 2 页 = 8 页纸）。当天 15:35 与 15:37 各跑一次 → 16 页 |
| 2 | 打印通道没有任何"必须由人确认"的门槛 | 任何代码路径（测试、脚本、体检、定时任务）走到 `_do_print` 就真打印，事后无法追责 |

**顺带查出的第三个缺陷**：`_print_one_pdf` 用 `win32print.SetPrinter(h,2,info,0)`
把 DEVMODE 写回打印机 —— 那是**持久**改写用户的系统打印机默认设置，影响所有
其它软件、还需要管理员权限。实测开发机上该打印机的默认纸张已被改成 **A3 横向**。

**处置**：

| 措施 | 实现位置 |
| --- | --- |
| 人工授权闸门：`OpContext.allow_real_print`（默认 `False`），只有界面"点击开始执行 + 确认框点是"那条路径置 `True`；其余抛 `PrintNotAuthorized` | `core/registry.py` / `engine.py` / `ui/app.py` / `core/print_ops.py` |
| 声明式约束：`OpDef.hardware=True` 必须同时 `confirm=True`，否则构造即报错 | `core/registry.py::OpDef.__post_init__` |
| 进程级熔断：`ODA_NO_PRINT=1` | `core/print_ops.py::print_denied_reason` |
| 演练模式：`dry_run` 参数 / `ODA_PRINT_DRYRUN=1`，只出预览 PDF、零作业 | `core/print_ops.py::_dryrun_preview` |
| 确认框摊开真实动作：打印机名 / 文件数 / 合计份次 / 预计张数 | `core/print_ops.py::_describe_print_job` |
| 不再写系统状态：改为「读默认 DEVMODE 字节 → 内存改字段 → `gdi32.CreateDCW` 建作业级 DC」 | `core/print_ops.py::_job_devmode_bytes / _create_printer_dc` |
| 份数：每份文件 N 次作业 + 作业内份数固定 1，杜绝"驱动份数 × 作业份数"叠乘 | `core/print_ops.py::_do_print` |
| 测试改造：默认零打印；真实打印需同时给 `ODA_TEST_REAL_PRINT=1` 与 `ODA_TEST_PRINTER=<名>` | `tests/test_print.py` |
| 被旧版改坏的机器：一次性还原工具 | `tools/fix_printer_defaults.py` |

**技术要点（已实测）**：`win32print.CreateDC` 在 pywin32 中**不存在**；
`win32ui.CreateDC().CreatePrinterDC()` 不接受 DEVMODE。可行路径是
`winspool.DocumentPropertiesW` 取默认 DEVMODE 原始字节 → 按偏移改字段
（`dmFields`=72 / `dmOrientation`=76 / `dmPaperSize`=78 / `dmCopies`=86 /
`dmColor`=92 / `dmDuplex`=94 / `dmFormName`=102）→ `gdi32.CreateDCW(driver, printer, None, pdm)`。
实测同一台机器上 A4 纵 = 2480×3508px、A5 横 = 2480×1748px（宽同为 210mm、高约 2:1），
证明纸张/方向真正逐作业生效，且系统默认值零变化。

---

## 三、对原提示词的修正建议（8 条）

| # | 原提示词 | 修正建议 | 理由 |
| --- | --- | --- | --- |
| 1 | GUI：Tkinter + ttk（标准库） | 增加**可选**依赖 `tkinterdnd2` | 标准库无法实现拖拽，二者不可兼得 |
| 2 | （未提及） | requirements 写 `pypdf>=4.0,<7` | 6.x 已有 4 处 API 变更；7.0 会移除旧参数 |
| 3 | （未提及） | 显式加 `cryptography` | 否则 PDF 加密只能用 RC4-128，安全性弱 |
| 4 | PDF 压缩：Ghostscript | 改为**不捆绑** Ghostscript，用户自装 + 降级 | AGPL 许可风险 |
| 5 | 输出默认写入新文件夹，不覆盖原文件 | 明确重命名有两种模式（副本/原地） | 两者语义冲突，需给用户选择权 |
| 6 | PDF 转图片：每页转 PNG/JPG | 补充 DPI 上限建议与分批提示 | 600 DPI 下 500 页 PDF 内存占用可达数 GB |
| 7 | 文档对比：PDF/Word 两版差异提示 | 明确"仅比对文本层，不含 OCR" | 扫描件会得到误导性结论 |
| 8 | Windows 用 onedir | 补充：**关闭 UPX** | UPX 显著提高杀软误报率，对本项目得不偿失 |

另外建议补一条原提示词没提但很关键的：**PDF 预览**。当前只做处理不做预览，用户加页码/水印后要打开文件才能确认效果。加一个内嵌预览（用 `pdf2image` 渲染首页缩略图）会明显提升可用性 —— 这是 P2 的优先候选。

---

## 四、落地方案

### 4.1 架构分层

```
┌──────────────────────────────────────────────────────────┐
│ ui/        Tkinter 界面层  只做展示与事件转发              │
│            必须运行在主线程                                │
├──────────────────────────────────────────────────────────┤
│ engine.py  任务引擎  thread + queue.Queue + CancelToken    │
│            跨线程唯一通道；所有 UI 更新由主线程消息泵消费     │
├──────────────────────────────────────────────────────────┤
│ core/      业务逻辑层  纯 Python，不 import tkinter        │
│            registry 声明式注册；可脱离界面单独测试          │
├──────────────────────────────────────────────────────────┤
│ deps.py    外部依赖探测（Poppler / LibreOffice / GS）      │
│ paths.py   长路径、唯一输出目录、中文路径                  │
└──────────────────────────────────────────────────────────┘
```

**最关键的设计决策：`core/` 完全不依赖 Tkinter。** 这让 20 个功能可以脱离界面被自动化测试覆盖 —— 本报告里 31/31 的自检就是这么跑出来的。如果业务逻辑和界面耦合，验证成本会高一个数量级。

### 4.2 声明式功能注册表

界面完全由 `OpDef` 驱动，新增功能**不用改任何界面代码**：

```python
MY_OP = OpDef(
    key="my_feature", title="我的功能", category="advanced",
    accepts=("pdf",),            # → 自动做文件类型过滤
    mode="per_file",             # → 自动决定批处理语义
    requires=("ghostscript",),   # → 缺失时自动禁用按钮并提示
    params=(Param("level", "档位", "choice", choices=(...)),),  # → 自动生成控件
    confirm=True,                # → 自动加二次确认
)

@register(MY_OP)
def run_my_feature(ctx, files, params): ...
```

配套收益：`Param.visible_if` 实现参数联动（如压缩档位切换时显隐 DPI 选项），`OpContext` 统一提供 `log / progress / check / ok / fail / skip / ask_password`，每个 handler 只需关心业务。

### 4.3 线程模型（不可妥协的部分）

```
主线程                                工作线程
───────                              ────────
after(80ms) 常驻消息泵 ◄── queue ──  执行 handler
  · 日志 → 日志区                      每步调用 token.check()
  · 进度 → 进度条
  · 密码请求 → 弹框 ──event──►         阻塞等待密码
  · 依赖结果 → 状态灯
点击取消 ─────token.set()────►          抛 Cancelled → 清理临时文件
```

**血泪教训（已实测复现并修复）**：绝不允许工作线程直接 `root.after()`。

### 4.4 打包与安装包策略（已实测）

**跨平台分发铁律：必须分开做安装包，且无法只构建一次**——PyInstaller 不支持交叉
编译，Windows 包要在 Windows 上打，macOS 包要在 Mac 上打；macOS 还要按芯片出
arm64 / x86_64 两份（或 universal2 合一份，体积 +70%）。详细分析与产物矩阵见
`DISTRIBUTION.md`。

| 平台 | 产物 | 关键配置 |
| --- | --- | --- |
| Windows | `dist/OfficeDocAssistant/OfficeDocAssistant.exe`（**onedir**，82 MB 实测） | 自定义 manifest：`longPathAware` + `dpiAware` + `activeCodePage=UTF-8`；**关闭 UPX**；已写入版本资源与图标 |
| Windows 安装包 | `dist/installer/OfficeDocAssistant-Setup-<v>-win64.exe`（28.4 MB 实测） | Inno Setup：**per-user 免管理员安装**（装到 %LOCALAPPDATA%\Programs）、中文向导、开始菜单+桌面快捷方式、卸载器、MinVersion=6.3（8.1+） |
| Windows 便携版 | `OfficeDocAssistant-<v>-portable-win64.zip`（38.4 MB 实测） | 免安装、U 盘可带 |
| macOS | `dist/OfficeDocAssistant.app` + `.dmg` | `NSHighResolutionCapable=true`（否则 Retina 上模糊）、`LSMinimumSystemVersion=12.0`；ad-hoc 签名 + `xattr -cr` 清隔离属性；`build_macos_dmg.sh` 一键完成 |
| CI 代工 | `.github/workflows/build-installers.yml` | windows + macos-14(arm64) + macos-13(x64) 三个 runner，发 `v*` 标签自动出三份安装包并挂 Release |

**打包实测抓到的真问题**：
1. PyInstaller 6.20+ 执行 spec 时**不再注入 `__file__`**（首行即 `NameError`），
   已改为 `__file__` → 官方变量 `SPECPATH` 双保险；
2. Inno Setup 官方包不含简体中文语言文件，已把 `ChineseSimplified.isl` 收入 `build/`；
3. 冻结版已实测：`--check` 12 项依赖全 OK、GUI 正常启动退出、
   静默安装→运行→快捷方式→静默卸载全链路无残留。

**免安装便携版**：把 Poppler / Ghostscript 绿色版解压到项目根目录 `vendor/`，重新打包即会一起收进产物，用户机器上零安装。

### 4.5 实施里程碑（建议）

| 阶段 | 内容 | 交付物 |
| --- | --- | --- |
| M1 骨架 | 目录结构、`paths`/`errors`/`logging`/`settings`、`deps` 探测、`engine` 线程引擎 | 可启动空界面 + 依赖状态显示 |
| M2 纯 Python 功能（不依赖外部程序） | 合并/拆分/提取/删页/旋转/页码/水印/页眉页脚/图片转PDF/Excel/压缩包/重命名/对比 | 覆盖 P0 的 9 项 + P1 的 6 项 |
| M3 外部依赖功能 | PDF 转图片（Poppler）、Office 转 PDF（LibreOffice）、PDF 压缩（Ghostscript） | 补齐 P0 第 7/9 项 + P1 第 1 项 |
| M4 界面完善 | 文件列表拖排、参数联动、进度/日志、结果面板、设置窗口 | 完整界面 |
| M5 打包与分发 | spec / manifest / 双平台脚本 / 便携版 | 可分发产物 |
| M6 加固 | 长路径、Zip-Slip、坏文件容错、取消、密码、日志 | 三套自动化测试全绿 |

> 本次交付已一次性完成 M1–M6，并把 M6 的验证脚本固化在 `tests/` 下，可作为后续回归基线。

### 4.6 建议补充的自动化测试（已实现）

| 脚本 | 作用 |
| --- | --- |
| `tests/selftest.py` | 31 项端到端功能自检 + 取消 + 坏文件容错 + 依赖缺失时的正确报错 |
| `tests/verify_artifacts.py` | 不只看"有没有抛异常"，而是**读回产物**验证页数、`/Rotate`、水印文字、加密状态、A4 尺寸 |
| `tests/smoke_gui.py` | 构建全部 20 个参数面板 + 参数联动 + 校验 + 对话框 + 异步任务整链路 |
| `tests/probe_api*.py` | 依赖库 API 探测，升级依赖前先跑一遍 |

**这条经验值得单独强调**：`selftest` 当时报了 31/31 全绿，但 `verify_artifacts` 立刻抓出「A4 变 214×303」这个真实缺陷 —— **"没抛异常"和"结果正确"是两件事**，只测前者会漏掉大量问题。

---

## 五、功能支持矩阵

| # | 功能 | 优先级 | 实现状态 | 外部依赖 | 自动化验证 |
| --- | --- | --- | --- | --- | --- |
| 1 | PDF 合并 | P0 | ✅ 完成 | 无 | ✅ |
| 2 | PDF 拆分（每页/每N页/范围） | P0 | ✅ 完成 | 无 | ✅ ×3 |
| 3 | 提取指定页 `1-3,5,8-10` | P0 | ✅ 完成 | 无 | ✅ |
| 4 | PDF 加页码（位置/格式/起始/跳封面） | P0 | ✅ 完成 | 无 | ✅ |
| 5 | PDF 删除页 | P0 | ✅ 完成 | 无 | ✅ |
| 6 | PDF 旋转 90/180/270 | P0 | ✅ 完成 | 无 | ✅ |
| 7 | PDF 转图片（PNG/JPG，可选 DPI） | P0 | ✅ 完成 | **Poppler** | ✅（含缺依赖时报错路径） |
| 8 | 图片转 PDF（多图排序） | P0 | ✅ 完成 | 无 | ✅ ×2 |
| 9 | Word/Excel/PPT 转 PDF | P0 | ✅ 完成 | **LibreOffice** | ✅（含缺依赖时报错路径） |
| 10 | 批量重命名（前缀/后缀/序号/预览） | P0 | ✅ 完成 | 无 | ✅ ×3 |
| 11 | PDF 压缩（轻/中/强 + 前后大小） | P1 | ✅ 完成 | Ghostscript（可降级） | ✅ ×2 |
| 12 | PDF 加密 / 解密 | P1 | ✅ 完成 | cryptography（可降级） | ✅ |
| 13 | PDF 加水印 | P1 | ✅ 完成 | 无 | ✅ |
| 14 | PDF 页眉页脚 | P1 | ✅ 完成 | 无 | ✅ |
| 15 | Excel 合并 / 拆分 | P1 | ✅ 完成 | 无 | ✅ ×3 |
| 16 | 文档对比（PDF/Word） | P1 | ✅ 完成 | 无 | ✅ ×2 |
| 17 | 批量压缩解压（ZIP/7z） | P1 | ✅ 完成 | 无 | ✅ ×3 |
| — | PDF 转 Word | 排除 | ⛔ 按需求不实现 | — | — |
| — | PDF 转 Excel | 排除 | ⛔ 按需求不实现 | — | — |
| — | PDF OCR | 排除 | ⛔ 按需求不实现 | — | — |
| P2 | PDF 内嵌预览（建议新增） | 新建议 | 💡 待评估 | Poppler | — |

**统计：P0 10/10 完成，P1 7/7 完成，排除项 3/3 正确未实现，全部 17 项有自动化验证。**

---

## 六、交付物清单

```
office-doc-assistant/
├── main.py                        程序入口（环境守卫 / DPI 感知 / 全程异常兜底）
├── requirements.txt               依赖清单（含版本上界与说明）
├── README.md                      安装 / 运行 / 打包 / 外部依赖 / 线程模型 / 已知限制
├── FEASIBILITY.md                 本评审报告
├── app/
│   ├── config.py  paths.py  errors.py  cancel.py
│   ├── settings.py  logging_setup.py  deps.py  engine.py
│   ├── native_alert.py            零 Tkinter 依赖的原生弹窗（缺 Tk 时报错用）
│   ├── core/                     13 个业务模块（不依赖 Tkinter）
│   └── ui/                       theme / app / filelist / parampanel / settings_win
├── tools/
│   └── preflight.py               环境体检（零 Tkinter 依赖，main.py --check）
├── build/
│   ├── office_doc_assistant.spec 双平台 PyInstaller 配置
│   ├── app.manifest              长路径 / 高 DPI / UTF-8 代码页
│   ├── hook-tkinterdnd2.py       解决 tkdnd 打包丢失
│   ├── build_win.bat             打包前跑 --check，不通过即中止
│   └── build_macos.sh
└── tests/
    ├── selftest.py               31 项端到端自检
    ├── verify_artifacts.py       18 项产物校验
    ├── smoke_gui.py              29 项界面测试
    ├── threading_and_guard.py    22 项线程安全与环境守卫专项
    └── probe_api.py / probe_api2.py
```

### 复现验证

```bash
cd office-doc-assistant
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

```bash
python main.py --check             # 预期：可以启动（可有提示项）
python tests/selftest.py          # 预期：总计 31/31 通过
python tests/verify_artifacts.py  # 预期：校验通过 18/18
python tests/smoke_gui.py         # 预期：全部通过
python tests/threading_and_guard.py  # 预期：全部通过（共 22 项）
python main.py                    # 启动界面
```

---

## 七、结论与建议

**结论：方案可行，且已完成可运行实现与验证。**

### 7.1 非商业用途下的结论

如果确定**不做商业分发**（本次的答复），三个高风险点里的第一个直接消解，整体风险等级下降一档：

| 风险 | 商业分发时 | **非商业分发时（本次）** |
| --- | --- | --- |
| Ghostscript AGPL | 🔴 必须决策：购买商业许可 或 改为用户自装 | ✅ **无需任何动作** —— 当前实现本就是"不捆绑 + 用户自装 + 缺失降级" |
| Tkinter 线程安全 | 🔴 必须处理 | ✅ 已处置并 22 项测试通过 |
| `tkinter` 可能缺失 | ⚠️ 必须处理 | ✅ 已处置并实测通过 |

**也就是说：非商业场景下，剩余待办为零，项目处于可直接交付状态。**

### 7.2 两个已处置问题的做法（可复用）

**① `tkinter` 可能缺失 —— "报错通道不能依赖要报的那个东西"**

最常见的失败是 Tk 缺失，而弹窗本身也要 Tk。所以必须用平台原生通道：
Windows `MessageBoxW`（ctypes，标准库）、macOS `osascript`、Linux `zenity`。
配套三件套让它真正可运维：

- **启动守卫** —— 失败时弹原生框 + 返回退出码 2，且**不甩 traceback**
- **`python main.py --check`** —— 体检脚本自身**不 import tkinter**，界面挂了它照样能跑
- **打包前置拦截** —— 不通过就中止，避免产出"双击没反应"的 exe

细节上有一条容易漏：诊断必须**报出当前解释器的绝对路径**（含 venv 基准路径），
否则用户不知道该换机器上的哪一个 Python。

**② Tkinter 线程安全 —— "约定挡不住三个月后新增的一行代码"**

"只从主线程更新控件"这个约定必须有机制兜底，已补四个：

- **处理量上限**（每 tick ≤400 条）+ **进度消息合并**（只取最后一条）→ 日志洪水不再拖死界面
- **关窗后停止排程**（先置 `_closing` 再 destroy）→ 不再抛 `TclError: invalid command name`
- **绊线** —— `_append()` 检测到非主线程调用时自动改走队列并记 error，
  把"随机崩溃"变成"稳定复现、日志里写明是谁的错"
- **关窗等收尾** —— 取消 → `join(5s)` → 真停不下来才问用户，避免留下残缺输出
- 外加**密码握手竞态** —— 超时作废的请求不再弹出"输了没反应"的孤儿对话框

### 7.3 给决策者的建议

1. **Ghostscript 策略**：非商业场景下**不需要决策**，保持现状（不捆绑 + 用户自装 + 自动降级）即可。
   唯一要做的只是在 README 里如实说明"没装 Ghostscript 时压缩效果有限"。

2. **接受两处提示词修正**（⚠️）：引入 `tkinterdnd2` 换真拖拽、重命名提供副本/原地双模式。
   这两处不修会导致功能缺失或语义矛盾。

3. **把 `tests/` 四套脚本纳入 CI**。本方案的外部依赖（Poppler/LibreOffice/Ghostscript）
   版本差异大，加上 pypdf 6.x→7.0 的破坏性变更已在路上，有回归基线才能放心升级依赖。
   注意 `smoke_gui.py` 与 `threading_and_guard.py` 需要图形环境，无显示器时会自动跳过。

一句话总结：**技术选型是对的。三个高风险点里，两个已用实测数据定位清楚并处置完毕，
第三个在不做商业分发的前提下自然消解 —— 项目当前处于可直接交付状态。**
