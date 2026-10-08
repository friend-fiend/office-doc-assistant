# 办公文档一键处理助手

跨平台（Windows 8.1 / 10 / 11 与 macOS 12+）的桌面文档批处理工具。
纯 Python + Tkinter 实现，无需安装 Office，所有输出写入新文件夹，**不覆盖原文件**。

---

## 一、功能一览

### P0（全部已实现）

| 功能 | 说明 |
| --- | --- |
| PDF 合并 | 按列表顺序合并多个 PDF，可为每个来源生成书签 |
| PDF 拆分 | 每页一个 / 每 N 页一个 / 按自定义页码范围 |
| 提取指定页 | `1-3,5,8-10` 语法，可把多文件结果再合并成一个 |
| PDF 加页码 | 位置、格式（`{n}` / `{total}`）、起始编号、跳过封面、字号、边距 |
| PDF 删除页 | 删除指定页后另存为新文件 |
| PDF 旋转 | 90° / 180° / 270°，可只旋转指定页 |
| PDF 转图片 | 每页导出 PNG/JPG，可选 72–600 DPI（需 Poppler） |
| 图片转 PDF | 多图排序合成，A4/A4 横向/Letter/原尺寸，白底透明处理 |
| Word/Excel/PPT 转 PDF | 调用 LibreOffice，保留原始排版（需 LibreOffice） |
| 批量重命名 | 前缀/后缀/序号/查找替换，**执行前可预览**，自动检测冲突 |

### P1（全部已实现）

| 功能 | 说明 |
| --- | --- |
| PDF 压缩 | 轻度（无损结构优化）/ 中度（150 DPI）/ 强力（96 DPI），显示前后大小与缩减比例 |
| PDF 加密 / 解密 | AES-256 / AES-128 / RC4-128 可选，权限位控制打印与复制 |
| PDF 加水印 | 文字水印，字号/颜色/透明度/旋转/平铺/置于内容下方 |
| PDF 页眉页脚 | 占位符 `{n}` `{total}` `{date}` `{filename}`，左中右对齐 |
| Excel 合并 / 拆分 | 每文件一表或追加到同一表；按工作表拆出独立文件 |
| 文档对比 | PDF / Word / 文本行级 diff，输出相似度与 HTML 红绿高亮报告 |
| 批量压缩解压 | ZIP / 7z / tar.gz，内置路径穿越（Zip-Slip）防护 |
| PDF 批量打印 | 走 Windows 打印队列，每份文件独立配置纸张/方向/份数/双面/颜色；支持「仅演练不打印」 |
| PDF 导出 PDF | 按指定纸张/方向/拼版合成新 PDF，**不经过打印系统**、无任何对话框 |

### 明确不实现

按需求排除：**PDF 转 Word**、**PDF 转 Excel**、**PDF OCR**。
`doc_compare` 在遇到扫描件（无文本层）时会明确提示"本工具不含 OCR，建议先做 OCR"，而不是给出误导性结论。

---

## 二、界面说明

```
┌──────────────────────────────────────────────────────────────┐
│ 顶栏：标题 · Poppler/LibreOffice/Ghostscript 状态 · 设置 · 关于 │
├──────────┬───────────────────────────────────────────────────┤
│ 左侧导航  │ 拖拽区（或点击选择文件）                            │
│  5 个分类 │ 文件列表：拖动排序 / 移除 / 清空 / 上移 / 下移       │
│  20 个功能│ 参数区：随功能自动切换                              │
│          │ 操作栏：输出目录 · 开始执行 · 取消 · 打开输出文件夹   │
│          │ 进度条 + 状态                                       │
│          │ 日志区（可保存 / 清空）                             │
└──────────┴───────────────────────────────────────────────────┘
```

- 左侧 5 个分类：**PDF 页面操作 / 格式转换 / 批量文件 / 压缩加密 / 高级工具**
- 参数区由功能定义自动生成，不需要手动找设置项
- 快捷键：`Ctrl+O` 添加文件、`F5` 执行、`Esc` 取消、`Ctrl+,` 设置
- 全中文界面；Windows 与 macOS 使用同一套 ttk 主题，行高与字体统一，控件不会错位

---

## 三、安装与运行

### 3.1 环境要求

| 项目 | 要求 |
| --- | --- |
| Python | **3.10 / 3.11**（3.12 也已验证可用；**必须用 python.org 官方安装包**，自带 tkinter） |
| 图形库 | Tkinter（官方安装包自带；部分精简版/管理器版 Python **不含 tkinter**） |
| 系统 | Windows 8.1 / 10 / 11（64 位）、macOS 12+ |

> ⚠️ 常见坑：某些发行版（如 Windows 的嵌入式包、部分 Linux 包管理器的 python）**不包含 tkinter**。

### 3.2 第一步永远是体检（强烈推荐）

```bash
python main.py --check      # 简明清单：哪些能用、哪些要装
python main.py --doctor     # 附带详细路径与版本
```

输出示例：

```
[ OK ] Python 版本
[ OK ] Tkinter（界面必需）
[ OK ] pypdf（PDF 读/写/合并/拆分）
[警告] Poppler（PDF 转图片）
        未检测到（影响：PDF 转图片）
        → Windows：下载 Release-Poppler-xx.xx.x-0.zip，解压到 C:\Program Files\poppler …
[ OK ] 配置目录可写
--------------------------------------------------------------------
结论：可以启动，但有 1 项提示（部分功能会禁用或降级）。
```

**退出码**：`0` = 可以启动（可能有可选功能被禁用）；`1` = 存在硬性缺失，无法启动。

体检脚本本身**不依赖 Tkinter**，所以哪怕界面已经起不来了，它照样能跑并告诉你原因。

**没有 Tkinter 时程序不会甩你一脸 traceback**，而是弹出系统原生对话框（Windows 用
`MessageBoxW`、macOS 用 `osascript`，均不依赖 Tkinter），内容包含：

- 具体缺什么、当前用的是哪个解释器（含 venv 的真实基准路径）
- 针对你所在平台的解决办法（含 `brew install python-tk` / `apt install python3-tk` 等）
- 退出码 `2`，方便脚本判断

### 3.3 从源码运行

```bash
# 1) 取代码
cd office-doc-assistant

# 2) 建虚拟环境
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS:
source .venv/bin/activate

# 3) 装依赖
pip install -r requirements.txt

# 4) 先体检，再运行
python main.py --check
python main.py
```

### 3.4 运行前自检（可选但推荐）

```bash
# 核心功能端到端自检（含取消、损坏文件容错）
python tests/selftest.py

# 产物有效性校验（读取输出文件，确认页数/水印/加密真的生效）
python tests/verify_artifacts.py

# 界面构建冒烟测试（把 20 个功能的参数面板全建一遍）
python tests/smoke_gui.py

# 线程安全与环境守卫专项测试（消息泵上限、密码握手竞态、关窗收尾、缺 Tkinter 提示）
python tests/threading_and_guard.py
```

> 后两个需要图形环境。无显示器（如 SSH / CI）时会自动跳过并说明原因，不算失败。

---

## 四、外部依赖安装

程序启动时会自动探测下列三项，**检测不到的功能按钮会被禁用**，其余功能照常可用。
也可以把绿色版解压到程序目录下的 `vendor/` 文件夹，程序会自动识别（便于做免安装便携版）。

### 4.1 Poppler —— PDF 转图片

| 平台 | 安装方式 |
| --- | --- |
| Windows | 下载 [poppler-windows Releases](https://github.com/oschwartz10612/poppler-windows/releases) 的 `Release-xx.xx.x-0.zip`，解压到 `C:\Program Files\poppler`，在「设置」里把 `Library\bin` 目录填进去 |
| macOS | `brew install poppler`（默认在 `/opt/homebrew/bin`，程序会自动找到） |

### 4.2 LibreOffice —— Word/Excel/PPT 转 PDF

| 平台 | 安装方式 |
| --- | --- |
| Windows | 从 [官网](https://www.libreoffice.org/download/download-libreoffice/) 下载 `.msi`，一路默认安装 |
| macOS | 下载 `.dmg`，把 LibreOffice.app 拖进 `/Applications` |

### 4.3 Ghostscript —— PDF 深度压缩

| 平台 | 安装方式 |
| --- | --- |
| Windows | 下载 [gs10.0x.x-win64.exe](https://ghostscript.com/releases/gsdnld.html) 安装（建议勾选添加到 PATH） |
| macOS | `brew install ghostscript` |

**缺失时的行为**：PDF 压缩自动降级为 pypdf 结构优化，并在日志和结果里提示
「压缩效果可能有限」——不会静默失败，也不会假装压好了。

> ⚖️ **许可提醒**：Ghostscript 采用 **AGPL** 许可。如果本工具要**对外商业分发**，
> 直接打包捆绑 Ghostscript 会触发 AGPL 的源码开放义务，需向 Artifex 购买商业许可，
> 或改为让用户自行安装（本程序默认就是这种模式）。详见评审报告。

### 4.4 拖拽支持（可选）

Tkinter 本身不支持从资源管理器拖入文件，需要 `tkinterdnd2`（已列入 requirements）：

```bash
pip install tkinterdnd2
```

未安装时程序正常运行，拖拽区自动退化为「点击选择文件」，并在日志里说明原因。

---

## 五、打包发布

> 跨平台分发要点：**Windows 包和 macOS 包必须分开做，且无法只构建一次**
> （PyInstaller 不支持交叉编译；macOS 还要按 Intel / Apple Silicon 出两份）。
> 完整的分发决策（安装包矩阵、签名/公证、升级、真机验收清单）见 `DISTRIBUTION.md`。

### 5.1 Windows：一键出齐三件套

```bat
build\build_win.bat
```

脚本依次完成：环境体检 → 建干净构建环境 → 装依赖 → 核心自检 → PyInstaller 打包
→ **冻结产物复检** → 编译安装包 / 便携版 zip。产物在 `dist\installer\`：

| 产物 | 说明 |
| --- | --- |
| `OfficeDocAssistant-Setup-1.0.0-win64.exe` | 安装包（Inno Setup）：**per-user 免管理员安装**、中文向导、开始菜单+桌面快捷方式、自带卸载器 |
| `OfficeDocAssistant-1.0.0-portable-win64.zip` | 便携版：解压即用，可放 U 盘 |
| `..\OfficeDocAssistant\` | onedir 原始产物（82 MB） |

- **打包前会自动跑 `main.py --check`**，环境体检不通过直接中止 ——
  缺 tkinter 时打出来的 exe 是"双击没反应"的，必须提前拦住
- 打包完成后还会用**冻结版自己再跑一遍 `--check`**，产物坏了就不出安装包
- 使用 **onedir** 而非 onefile：冷启动快、异常退出不留临时文件、不易被杀软误报
- `app.manifest` 已声明 `longPathAware`（长路径）、`dpiAware`（高 DPI）与 UTF-8 代码页
- exe 已写入版本信息与图标（版本信息由 spec 现场生成到 `build/version_info.txt`，图标在 `assets/app.ico`）
- 找不到 Inno Setup 时自动跳过 Setup.exe，仅产出便携版并提示
- **版本号只有一个权威来源**：`build/office_doc_assistant.spec` 顶部的 `VERSION`。
  `version_info.txt` 由 spec 按它现场生成，`build_win.bat` / `make_portable_win.sh` /
  `build_macos_dmg.sh` 从它读出来，CI 传给 ISCC 的 `/DAppVersion=` 也来自它。
  改版本只需改一处。
  > 这里踩过一次真实的坑：`version_info.txt` 原本是手写文件，spec 里又写了一份
  > `VERSION`，1.0.1→1.0.2 时只改了字符串表里的 `"1.0.2.0"`，ffi 的
  > `filevers`/`prodvers` 元组还是 `(1,0,1,0)`。于是右键属性显示 1.0.2.0，
  > 而 `win32api.GetFileVersionInfo` 读出 1.0.1.0 —— 资产盘点与升级判断读的是后者。
  > 现在由 spec 生成，并且 `--check` 里新增「版本资源一致性」交叉校验，
  > 冻结产物上不一致会直接判失败，坏包出不去。

### 5.2 macOS（.app + .dmg）

```bash
chmod +x build/build_macos_dmg.sh
./build/build_macos_dmg.sh            # 本机架构
./build/build_macos_dmg.sh universal2 # Intel + Apple Silicon 合体包
```

产物：`dist/OfficeDocAssistant.app` 与 `dist/installer/*.dmg`

- **打包前会用打包环境自己的解释器跑 `main.py --check`**，不通过即中止
- 脚本会自动 `xattr -cr` 清理隔离属性并做 ad-hoc 签名，本机可直接 `open`
- 正式对外分发需要 Developer ID 签名 + 公证（脚本末尾已给出命令）
- `build_macos.sh`（不带 DMG）与旧流程仍保留

### 5.3 没有 Mac？让 CI 代工

仓库已内置 `.github/workflows/build-installers.yml`：推到 GitHub 后在 Actions
手动触发（或推 `v*` 标签），会用 Windows + macOS(arm64) + macOS(x64) 三个云端
runner 产出三份安装包，发标签时自动挂到 Release。签名/公证环节已写好，
配置 secrets 后自动启用。

### 5.4 把外部工具一起分发（可选）

把 Poppler / Ghostscript 的绿色版解压到项目根目录的 `vendor/`（可放 `vendor/poppler/`、`vendor/ghostscript/`），
再执行打包脚本，程序会把它们一起收进产物 —— 用户机器上无需任何额外安装。
注意 Ghostscript 的 AGPL 许可约束（见 4.3 节）。

---

## 六、安全与错误处理

- **不覆盖原文件**：所有输出写到新建的 `<功能名>_输出` 文件夹，重名自动追加 `_2`、`_3`
- **二次确认**：删除页、加密、原地重命名、清空列表都会先弹确认框
- **原地重命名可回滚**：会生成 `重命名回滚清单.json`
- **中文与长路径**：全程 `pathlib` + Windows `\\?\` 前缀；打包 manifest 声明 `longPathAware`
- **异步与取消**：工作线程执行、主线程更新界面；取消是协作式的，会等当前文件处理完再退出，保证临时文件被清理
- **密码处理**：加密 PDF 会弹出密码框，最多重试 3 次；7z/ZIP 加密包同理
- **坏文件不中断任务**：损坏的 PDF 被记录到失败明细后跳过，同批次的好文件照常处理
- **解压安全**：ZIP/7z/tar 解压前逐条校验路径，拒绝 `../` 与绝对路径（Zip-Slip 防护），tar 还会拒绝符号链接
- **日志**：滚动写入 `%LOCALAPPDATA%\OfficeDocAssistant\app.log`（macOS 为 `~/Library/Logs/`）

### 6.1 线程模型（改动界面代码前必读）

**Tkinter 不是线程安全的。** 工作线程直接调 `root.after()` 会抛
`RuntimeError: main thread is not in main loop`；直接改控件则可能**随机崩溃或随机卡死**。

所以本程序只有一条规则：

```
工作线程 / 依赖探测线程 / logging 回调
        │
        ├─ TaskManager.bus  ──┐          （只放数据，绝不碰控件）
        └─ self._ui_queue   ──┤
                              ▼
                    _pump()（主线程，after 80ms 常驻）
                              │
                              ▼
                        更新控件
```

消息泵有三条硬性约束，都在 `app/ui/app.py` 里实现了：

| 约束 | 为什么 |
| --- | --- |
| 每 tick 最多处理 400 条 | 一个日志量巨大的任务会把队列瞬间塞满；一次全画完界面会假死几十秒 |
| 进度消息只取最后一条 | 同一 tick 内重画 300 次进度条毫无意义，纯粹浪费主线程时间 |
| 关窗后停止排程 | 否则已排队的回调会打到已销毁的控件上，抛 `TclError: invalid command name` |

另外 `_append()` 里有一条**绊线**：一旦检测到非主线程直接写日志控件，会自动改走队列
并记录一条 error。这样潜在的随机崩溃会变成一个**可见、可定位的缺陷报告**，而不是
上线后偶发。

关窗流程也不是直接 `destroy()`：先请求取消 → `join(5s)` 等任务收尾 → 真的停不下来
才询问用户是否强制退出。直接退出会留下**残缺的半成品输出文件**，而用户会以为处理成功了。

### 6.2 密码握手的竞态

加密 PDF 需要向用户要密码，这个交互是"工作线程阻塞等待 + 主线程弹框"。
有一个容易忽略的竞态：用户一直不操作，工作线程的等待超时（300 秒）后继续往下跑了，
等用户终于点了确定 —— 结果落到一个"已经没人等"的请求上，界面上表现为**输了密码但毫无反应**。

处理办法是给每个密码请求一个带 `expired` 标记的 box：超时即作废，
`deliver_password(value, box=box)` 会先校验身份，作废则静默忽略；
界面侧也会跳过已作废的请求，不弹出孤儿对话框。

### 6.3 打印安全模型（1.0.2 起）

**背景**：1.0.1 有一个真实事故 —— 用户什么都没操作，打印机自己开始出纸。
两个叠加的原因：

1. `tests/test_print.py` 用**系统默认打印机**做端到端验证，等于每次跑测试就
   往真机送 4 个作业（8 页纸）；
2. 打印通道没有任何"必须由人确认"的门槛，任何代码路径（测试、脚本、体检、
   定时任务）只要走到 `_do_print` 就会真打印。

**现在的三道闸门**：

| 机制 | 说明 |
| --- | --- |
| `OpContext.allow_real_print`（默认 `False`） | 真实打印的授权标志。全程序**只有** `ui/app.py` 里"用户点开始执行 + 确认框点是"这一条路径会把它置 `True`；其余一律抛 `PrintNotAuthorized` |
| `OpDef.hardware = True` | 强约束：标记为物理动作的功能必须同时 `confirm = True`（`OpDef.__post_init__` 里报错拦截），否则"界面上点一下"和"脚本调一次"就没区别了 |
| `ODA_NO_PRINT=1` | 进程级熔断，一键关掉本进程内所有真实打印（CI / 排障用） |

另外新增 **演练模式**（界面勾选「仅演练不打印」，或 `ODA_PRINT_DRYRUN=1`）：
完全不碰打印机，把每份文件按当前配置渲染成预览 PDF 落到输出目录。

**不再修改系统打印机默认设置**：旧版用 `win32print.SetPrinter` 把 DEVMODE
写回打印机 —— 那是**持久**改写用户的系统默认，会连带影响所有其它软件（实测
开发机的默认纸张就是这么被改成 A3 横向的），还要管理员权限。
现在改为「读默认 DEVMODE 的原始字节 → 在内存里改纸张/方向/份数 → `gdi32.CreateDCW`
建一个只属于本次作业的 DC」，不做任何系统写入、不需要管理员权限。
被旧版改坏的机器可以用 `python tools/fix_printer_defaults.py --restore "<打印机名>"`
还原。

**份数**按"每份文件提交 N 次作业、作业内份数固定 1"实现：不依赖驱动是否支持
`DM_COPIES`，也不会出现"驱动份数 × 作业份数"叠乘打出一堆纸。

---

## 七、目录结构

```
office-doc-assistant/
├── main.py                      # 入口：DPI 感知、异常兜底、启动界面
├── requirements.txt
├── README.md
├── FEASIBILITY.md               # 可行性评审与落地方案
├── app/
│   ├── config.py                # 常量：扩展名、压缩档位、字体方案
│   ├── paths.py                 # 长路径、唯一输出目录、打开文件夹、大小格式化
│   ├── errors.py                # 异常体系与 TaskReport
│   ├── cancel.py                # 协作式取消令牌
│   ├── settings.py              # 设置持久化（跨平台配置目录）
│   ├── logging_setup.py         # 滚动日志 + 转发到界面
│   ├── native_alert.py          # 不依赖 Tkinter 的原生弹窗（缺 Tk 时报错用）
│   ├── deps.py                  # Poppler / LibreOffice / Ghostscript 探测
│   ├── engine.py                # 线程 + 队列任务引擎、消息总线
│   ├── core/                    # 业务逻辑（不依赖 Tkinter，可单独测试）
│   │   ├── registry.py          # OpDef / Param 声明式注册表
│   │   ├── pdfutil.py           # 读取器、页码解析、中文字体、覆盖层
│   │   ├── pdf_pages.py         # 合并/拆分/提取/删页/旋转
│   │   ├── pdf_stamp.py         # 页码/水印/页眉页脚
│   │   ├── pdf_compress.py      # 三档压缩（pypdf + Ghostscript）
│   │   ├── pdf_security.py      # 加密/解密
│   │   ├── pdf_to_image.py      # PDF 转图片（pdf2image + Poppler）
│   │   ├── image_to_pdf.py      # 图片转 PDF（Pillow）
│   │   ├── office_to_pdf.py     # Office 转 PDF（LibreOffice headless）
│   │   ├── excel_ops.py         # Excel 合并/拆分（openpyxl）
│   │   ├── archive_ops.py       # ZIP/7z 压缩解压（含 Zip-Slip 防护）
│   │   ├── rename_ops.py        # 批量重命名 + 预览
│   │   └── compare_ops.py       # 文档对比（difflib）
│   └── ui/
│       ├── theme.py             # ttk 主题、跨平台字体与行高
│       ├── app.py               # 主窗口、消息泵、结果/预览对话框
│       ├── filelist.py          # 文件列表（拖动排序/删除/清空）
│       ├── parampanel.py        # 由 Param 声明动态生成参数控件
│       └── settings_win.py      # 设置与依赖状态
├── build/
│   ├── office_doc_assistant.spec # PyInstaller 双平台配置
│   ├── version_info.txt          # exe 版本资源（打包时由 spec 生成，勿手改）
│   ├── make_icons.py             # 生成 ico/icns/png 图标
│   ├── app.manifest             # 长路径/DPI/UTF-8 声明
│   ├── hook-tkinterdnd2.py
│   ├── installer_win.iss        # Inno Setup 安装包脚本（per-user、中文向导）
│   ├── ChineseSimplified.isl    # 安装向导简体中文语言文件
│   ├── build_win.bat            # 一键：体检→自检→打包→复检→安装包+便携版
│   ├── make_portable_win.sh     # 便携版 zip
│   ├── build_macos_dmg.sh       # macOS 一键：体检→自检→打包→DMG
│   └── build_macos.sh
├── tools/
│   └── preflight.py             # 环境体检（零 Tkinter 依赖，可独立运行）
├── .github/workflows/
│   └── build-installers.yml     # 三平台 CI：Windows + macOS arm64/x64
├── DISTRIBUTION.md              # 分发方案：安装包矩阵/签名/升级/验收清单
├── LICENSE-NOTICES.txt          # 第三方许可声明（随安装包分发）
└── tests/
    ├── selftest.py              # 核心功能端到端自检
    ├── verify_artifacts.py      # 产物有效性校验
    ├── smoke_gui.py             # 界面构建冒烟测试
    ├── threading_and_guard.py   # 线程安全 + 环境守卫专项测试
    ├── probe_api.py             # 依赖库 API 探测
    └── probe_api2.py
```

---

## 八、扩展一个新功能

因为界面完全由 `OpDef` 驱动，新增功能只需两步：

```python
# 1) 在 app/core/xxx.py 里写 handler
from .registry import OpContext, OpDef, Param, register

MY_OP = OpDef(
    key="my_feature", title="我的功能", category="advanced",
    accepts=("pdf",), mode="per_file",
    params=(Param("n", "次数", "int", default=1, min_value=1, max_value=9),),
)

@register(MY_OP)
def run_my_feature(ctx: OpContext, files: list[str], params: dict) -> None:
    ctx.set_total(len(files))
    for f in files:
        ctx.check()                       # 取消检查点
        try:
            ...                           # 干活
            ctx.ok(out_path, src=f)       # 计成功
        except Exception as exc:
            ctx.fail(f, str(exc))         # 计失败，不中断批次
        ctx.step()                        # 进度 +1
```

```python
# 2) 在 app/core/__init__.py 的 import 列表里加上模块名
```

界面上的分类按钮、参数控件、类型过滤、依赖禁用会自动出现。

---

## 九、已知限制

1. **PDF 转图片**依赖 Poppler 与 `pdf2image`，大文件（>500 页）在 600 DPI 下内存占用较高，建议分批。
2. **文档对比**比较的是文本层。扫描件没有文本层，会提示改用 OCR —— 本工具不内置 OCR。
3. **旧版 `.doc`** 无法用 python-docx 直接读取文本，文档对比建议先另存为 `.docx`。
4. **LibreOffice 转换**对复杂排版的还原度取决于 LibreOffice 本身，个别图表/公式可能与 Office 有差异。
5. **PDF 加密**使用 RC4-128 时安全性较弱（这是 PDF 规范的旧算法）；装了 `cryptography` 后会自动使用 AES-128。需要 AES-256 请在加密功能的算法选项里显式选择。
6. macOS 未做公证的 `.app` 在别人机器上首次打开需右键 → 打开。
