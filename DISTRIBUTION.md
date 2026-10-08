# 分发方案 —— 怎么把它变成"别人电脑上也能装的软件"

> 回答两个问题：**还有哪些因素要考虑？** 和 **要不要分开做安装包？**
> 结论先行：**必须分开，而且没法只构建一次**。Windows 包和 macOS 包要各自在对应
> 系统上打出来；macOS 还要额外考虑 Intel / Apple Silicon 两种芯片。
> 本文所有"已验证"的条目都附了实测数据，没验证的也明确标出。

---

## 一、为什么必须分开做（能不能只打一个包？不能）

| 限制 | 说明 |
| --- | --- |
| PyInstaller 不支持交叉编译 | Windows 机器上打不出 .app，Mac 上也打不出 .exe。因为打包本质是把当前平台的 Python 解释器、二进制依赖、bootloader 复制进去 |
| macOS 芯片分叉 | Apple Silicon (M1–M4) 是 arm64，老 Mac 是 x86_64，**互相不兼容**。要么各打一份（推荐，包小），要么 universal2 合一份（体积 +70%，且必须装 universal2 版 Python） |
| 安装包格式分叉 | Windows 用 .exe 安装器（Inno Setup）；macOS 用 .app + .dmg 镜像。没有"通用格式" |
| Apple 政策 | macOS 构建只能在 macOS 上完成（工具链和签名要求） |

**最终产物矩阵：**

| 产物 | 覆盖人群 | 状态 |
| --- | --- | --- |
| `OfficeDocAssistant-Setup-1.0.0-win64.exe` | Windows 8.1/10/11 x64 | ✅ 本机已产出并完成安装→运行→卸载全链路验收 |
| `OfficeDocAssistant-1.0.0-portable-win64.zip` | 免安装/U 盘场景 | ✅ 本机已产出（38.4 MB） |
| `OfficeDocAssistant-<v>-macOS-arm64.dmg` | M1–M4 芯片 Mac | ⚠️ 脚本就绪，需在 Mac 上执行 |
| `OfficeDocAssistant-<v>-macOS-x86_64.dmg` | Intel 芯片 Mac | ⚠️ 脚本就绪，需在 Mac 上执行 |

---

## 二、本次实测结果（Windows，本机真实执行）

| 验证项 | 结果 |
| --- | --- |
| PyInstaller onedir 打包 | ✅ 58 秒完成，产物 82 MB |
| 冻结版 `--check` 体检 | ✅ 12 项依赖全部 OK，退出码 0 |
| tkdnd（拖拽二进制）收集 | ✅ `_internal/tkinterdnd2/tkdnd/win-x64/` 完整 |
| cryptography（AES） | ✅ 完整打入 |
| 冻结版 GUI 启动 | ✅ 进程存活、窗口标题正确、正常退出 |
| Inno Setup 编译安装包 | ✅ 20 秒，28.4 MB（压缩比 87→28） |
| 静默安装 | ✅ 装入 `%LOCALAPPDATA%\Programs`，无 UAC 弹窗 |
| 安装版运行 | ✅ GUI 正常启动退出 |
| 快捷方式 | ✅ 开始菜单 + 桌面均创建 |
| 静默卸载 | ✅ exe、两个快捷方式全部清除，无残留 |
| 便携版 zip | ✅ 38.4 MB，含使用说明 |

**打包过程中抓到的真问题**（又是"没实际跑过就不知道"系列）：

1. **PyInstaller 6.20+ 执行 spec 时不再注入 `__file__`**——原 spec 第一行就 `NameError` 崩掉。
   已改为优先 `__file__`、回退官方变量 `SPECPATH` 的双保险写法。
2. Inno Setup 官方包**不含简体中文语言文件**（属非官方翻译集），已把
   `ChineseSimplified.isl` 收进 `build/` 目录，离线可用。
3. Git-Bash 调 Inno 时 `/LOG=/tmp/...` 会被转成非法路径导致安装器退出码 1
   （冒充"安装失败"）。静默安装参数要用 Windows 原生路径写法。

## 三、怎么在"没有 Mac"的情况下产出 macOS 包

两条路，推荐第二条：

1. **找一台 Mac**：把项目拷过去，运行 `./build/build_macos_dmg.sh`（参数可选
   `arm64` / `x86_64` / `universal2`）。脚本自带体检、自检、打包、去隔离属性、
   ad-hoc 签名、`hdiutil` 制 DMG，全程只用 macOS 自带工具。
2. **用 CI 代工（零成本）**：项目里已配好 `.github/workflows/build-installers.yml`。
   把仓库推上 GitHub，打开 Actions 点 Run（或推一个 `v*` 标签），会用
   windows-latest + macos-14 + macos-13 三个云端 runner 同时产出
   **Windows 安装包 + 两份 macOS DMG**，发标签时自动挂到 Release 页。
   GitHub Actions 公共仓库免费，私有仓库每月也有免费额度，这个量级够用。

## 四、其余必须考虑的因素（按优先级）

### 4.1 🔴 外部依赖不在安装包里 —— 最大的体验缺口

Poppler / LibreOffice / Ghostscript **不会**被打进安装包（原因见 FEASIBILITY.md
的许可分析）。接收者拿到安装包后：20 项功能里 17 项开箱即用，
**PDF 转图片、Office 转 PDF、深度压缩这 3 项需要用户自己装对应工具**。

| 方案 | 体积 | 体验 | 合规 | 建议 |
| --- | --- | --- | --- | --- |
| A. 维持现状：程序内提示 + 设置页指路 | 28 MB | 3 项功能要手动装工具 | ✅ 无风险 | **起步用这个** |
| B. 把 Poppler（GPL-2.0）作为可选组件捆进安装包 | +30 MB | PDF 转图片开箱即用 | 非商业可用，需附带其许可与源码链接 | 第二步可做 |
| C. 捆绑 LibreOffice | +600 MB | 全功能开箱 | MPL-2.0 允许 | 不建议，体积失控 |
| D. Windows 有 MS Office 时走 COM 直转 Word/Excel/PPT | 0 | 保真度比 LibreOffice 还高 | ✅ | 值得做的后续增强（需先在装有 Office 的机器上实测） |

Ghostscript 无论如何都不捆绑（AGPL，捆绑闭源程序需商业许可）；
缺失时程序已自动降级并明示。

### 4.2 🟠 签名问题 —— "双击能不能打开"的分水岭

**非商业 ≠ 免签名，只是"没有签名也能用，但多两步操作"：**

| 平台 | 不签名的后果 | 用户的绕法 | 彻底解决 |
| --- | --- | --- | --- |
| Windows | SmartScreen 蓝屏警告"已保护你的电脑"；部分杀软可能误报 | 点"更多信息→仍要运行"；被误报时加入白名单 | OV 代码签名证书约 ¥1000–3000/年，或 Azure Trusted Signing（更便宜） |
| macOS | Gatekeeper 直接拦，"无法验证开发者" | 右键→打开；或终端 `xattr -cr /Applications/OfficeDocAssistant.app` | Apple Developer 账号 $99/年 + 公证（CI 流程已留好接口，配 3 个 secret 即启用） |

内部十几个人的团队分发：**不签，写清绕法**即可，无法律问题。
对外给陌生用户：建议至少 Windows 买证书，macOS 没有公证几乎必然流失用户。

### 4.3 🟡 版本与升级

- `AppId`（installer_win.iss 里那串 GUID）**发布后就永远不能改**，否则升级会装成两份。
- 现阶段升级方式 = 装新版覆盖（设置保存在 `%LOCALAPPDATA%\OfficeDocAssistant`，不会被卸载/覆盖清掉）。
- 以后想加"检查更新"：启动时 GET 一个版本接口比对即可，先不要上自动更新框架。

### 4.4 🟡 真机验收清单（分发前最后一道关）

在干净环境（虚拟机即可）各过一遍：

- [ ] Windows 10 / 11：安装 → 合并/压缩/重命名各跑一次 → 卸载
- [ ] Windows 8.1（如真要支持）：重点验证 Python 3.11 构建产物（3.13 构建的不支持 8.1）
- [ ] 中文路径、桌面被 OneDrive 接管的机器
- [ ] 无管理员权限的账号：验证"仅为我安装"路径
- [ ] macOS 12（Intel）与最新版（Apple Silicon）各一台
- [ ] 断网运行：全部功能应可离线使用（本程序无任何网络请求）

### 4.5 🟢 其余已就绪、无需动作的项

| 项 | 现状 |
| --- | --- |
| 用户数据隐私 | 全部本地处理，无遥测，README 已声明 |
| 卸载残留 | 卸载有意保留用户设置目录（重装可继承），文档已注明 |
| 第三方许可 | `LICENSE-NOTICES.txt` 已随安装包分发 |
| 杀软误报抑制 | 已用 onedir（非 onefile）、关闭 UPX、写入完整版本资源，误报概率已压到最低 |
| 高 DPI / 长路径 | manifest 已声明 |
| 图标 | 已生成并验证（ico/icns/png，`build/make_icons.py` 可复现） |

---

## 五、你现在的三条上手路径

1. **立刻发给同事（Windows）**：直接用 `dist/installer/` 里的两个文件，
   已验收。发 `Setup.exe` 或便携 zip 均可。
2. **要 macOS 包**：推 GitHub → Actions 跑一次 → 下载三份安装包。
3. **要更正式**：买代码签名证书 + Apple 开发者账号，CI 的签名/公证环节
   已写好，填入 secrets 即自动生效。

## 六、构建文件索引

| 文件 | 用途 |
| --- | --- |
| `build/office_doc_assistant.spec` | PyInstaller 双平台配置（含实测修复） |
| `build/version_info.txt` | exe 属性页的版本信息（**由 spec 按 `VERSION` 自动生成，请勿手改**） |
| `build/make_icons.py` | 生成 ico/icns/png 图标 |
| `build/build_win.bat` | Windows 一键：体检→自检→打包 |
| `build/installer_win.iss` | Inno Setup 安装包脚本（per-user、中文向导） |
| `build/make_portable_win.sh` | 便携版 zip |
| `build/build_macos_dmg.sh` | macOS 一键：体检→自检→打包→DMG |
| `build/ChineseSimplified.isl` | 安装向导简体中文语言文件 |
| `.github/workflows/build-installers.yml` | 三平台 CI（含可选签名/公证） |
