; Inno Setup Script for 利息计算器
; 生成Windows安装包

#define MyAppName "利息计算器"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "利息计算器"
#define MyAppURL ""
#define MyAppExeName "利息计算器.exe"

[Setup]
; 应用基本信息
AppId={{B5E9A7D8-8C4F-4E6B-9A3D-2F7E1C5D4B8A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

; 安装目录
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes

; 输出设置
OutputDir=..\installer
OutputBaseFilename=利息计算器_Setup_v{#MyAppVersion}

; 压缩设置
Compression=lzma2
SolidCompression=yes
LZMAUseSeparateProcess=yes

; Windows版本要求
MinVersion=10.0

; 权限
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; 安装向导设置
WizardStyle=modern
DisableWelcomePage=no

; 卸载设置
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
; 主程序及依赖
Source: "dist\利息计算器\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; 如有桌面快捷方式需要的图标
; Source: "interest_calculator.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; 开始菜单
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autoprograms}\{#MyAppName} - 卸载"; Filename: "{uninstallexe}"

; 桌面快捷方式
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

; 快速启动
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
; 安装完成后运行
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Registry]
; 注册表：记住窗口位置（可选）
Root: HKCU; Subkey: "Software\利息计算器"; Flags: uninsdeletekey

[UninstallDelete]
; 卸载时删除用户数据（可选）
Type: filesandordirs; Name: "{localappdata}\利息计算器"

[Code]
// 检查是否有其他实例在运行
function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  Result := True;
  // 尝试关闭运行中的程序
  Exec('cmd.exe', '/c taskkill /F /IM 利息计算器.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;