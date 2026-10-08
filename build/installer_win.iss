; =====================================================================
;  办公文档一键处理助手 —— Windows 安装包（Inno Setup 6.3+）
;
;  设计决策（为什么这样做）
;  -------------------------------
;  1. PrivilegesRequired=lowest      → 默认"仅为我安装"，不弹 UAC，
;     装到 %LOCALAPPDATA%\Programs\OfficeDocAssistant。
;     这也是 VS Code / Chrome 的默认做法：学校/公司电脑没有管理员权限也能装。
;  2. MinVersion=6.3                 → 最低 Windows 8.1，与需求一致。
;  3. lzma2/max + solid              → 安装包体积比 zip 小 30% 以上。
;  4. 中文向导                       → 依赖同目录的 ChineseSimplified.isl。
;
;  用法
;  ----
;     先跑 build_win.bat 产出 dist\OfficeDocAssistant\，然后：
;     iscc build\installer_win.iss
;     产物：dist\installer\OfficeDocAssistant-Setup-<版本>-win64.exe
;     版本号可在命令行覆盖：iscc /DAppVersion=1.2.0 build\installer_win.iss
; =====================================================================

; 版本号默认值。权威来源是 build\office_doc_assistant.spec 里的 VERSION，
; build_win.bat 会自动读出来并用 /DAppVersion 传进来；这里只是手动调用时的兜底。
#ifndef AppVersion
  #define AppVersion "1.0.2"
#endif

#define AppName "办公文档一键处理助手"
#define AppNameEn "OfficeDocAssistant"
#define AppExe "OfficeDocAssistant.exe"

[Setup]
AppId={{8E6F4C21-9A3D-4B5E-9C7F-0DA1B2C3E4F5}
; 注意：AppId 一旦对外发布就【永远不要再改】，否则升级会装成两份
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
; 注意：AppVersion 是给人看的字符串，不会自动写进 Setup.exe 的版本资源。
; 不写下面这几行，右键安装包看到的 FileVersion 是 0.0.0.0，
; 用户和企业分发时无法区分版本，也更容易被判为可疑文件。
VersionInfoVersion={#AppVersion}
VersionInfoProductVersion={#AppVersion}
VersionInfoTextVersion={#AppVersion}
VersionInfoProductName={#AppName}
VersionInfoProductTextVersion={#AppVersion}
VersionInfoCompany={#AppNameEn}
VersionInfoDescription={#AppName} 安装程序
AppPublisher={#AppNameEn}
DefaultDirName={autopf}\{#AppNameEn}
DisableDirPage=no
DisableProgramGroupPage=yes
OutputDir=..\dist\installer
OutputBaseFilename={#AppNameEn}-Setup-{#AppVersion}-win64
SetupIconFile=..\assets\app.ico
UninstallDisplayIcon={app}\{#AppExe}
Compression=lzma2/max
SolidCompression=yes
LZMAUseSeparateProcess=yes
LZMADictionarySize=65536
WizardStyle=modern
; Windows 8.1 = NT 6.3。低于此版本的系统直接提示不支持
MinVersion=6.3
; 默认"仅为我安装"：不弹 UAC、无需管理员
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
ShowLanguageDialog=no
; 覆盖安装时检测正在运行的程序并提示关闭
CloseApplications=yes

[Languages]
Name: "chinesesimplified"; MessagesFile: "ChineseSimplified.isl"

[Messages]
; 非官方语言包可能落后于 Inno 新增的键，缺的回落英文即可，这里再补齐最显眼的几句
WelcomeLabel2=这将安装【办公文档一键处理助手】到你的电脑。%n%n一款离线的 PDF / Office / 图片 / 压缩包批量处理工具：所有处理都在本机完成，文件不会上传到任何服务器。

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式(&D)"; \
    GroupDescription: "附加任务："; Flags: checkedonce

[Files]
; 打包产物整体拷入（onedir 目录必须完整分发，不能只拷 exe）
Source: "..\dist\{#AppNameEn}\*"; DestDir: "{app}"; \
    Flags: ignoreversion recursesubdirs createallsubdirs
; 附带说明文档
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE-NOTICES.txt"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
; 功能详解（浏览器打开，含每个功能的参数说明与限制）
Source: "..\docs\USER_GUIDE.html"; DestDir: "{app}\docs"; \
    Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"; \
    Comment: "PDF 合并/拆分/压缩、格式转换、批量重命名等一站式处理"
Name: "{autoprograms}\使用说明"; Filename: "{app}\docs\USER_GUIDE.html"; \
    Comment: "查看 24 个功能的详细说明"; \
    Flags: createonlyiffileexists
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; \
    Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "立即运行 {#AppName}"; \
    Flags: nowait postinstall skipifsilent

[UninstallRun]
; 无需额外清理：用户设置存于 %LOCALAPPDATA%\OfficeDocAssistant，
; 卸载时有意保留（重装后配置还在）。要彻底清除可手动删该目录。

[UninstallDelete]
; 清理运行期可能产生的临时文件（不影响用户设置目录）
Type: filesandordirs; Name: "{app}\_selftest"
