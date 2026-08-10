#define MyAppName "抖音智能发布中心 v3.0.4 补丁"
#define MyAppVersion "3.0.4"
#define MyAppPublisher "百业信息"
#define MyAppExeName "DouyinPublisher.exe"
#ifndef HotfixAppId
#define HotfixAppId "{{E4C4B9AC-D87F-4A1C-A082-F5CBDF962760}"
#endif
#ifndef HotfixOutputDir
#define HotfixOutputDir "..\release"
#endif
#ifndef HotfixOutputBaseFilename
#define HotfixOutputBaseFilename "DouyinPublisher_Hotfix_v" + MyAppVersion
#endif

[Setup]
AppId={#HotfixAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={code:GetInstallDir}
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir={#HotfixOutputDir}
OutputBaseFilename={#HotfixOutputBaseFilename}
SetupIconFile=..\app\assets\app_logo.ico
Compression=lzma2/normal
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
CloseApplicationsFilter=DouyinPublisher.exe
RestartApplications=no
Uninstallable=no
CreateUninstallRegKey=no
VersionInfoVersion={#MyAppVersion}.0
VersionInfoProductName={#MyAppName}
VersionInfoCompany={#MyAppPublisher}

[Languages]
Name: "chinesesimp"; MessagesFile: "ChineseSimplified.isl"

[Files]
Source: "..\build\hotfix\dist\DouyinPublisher\DouyinPublisher.exe"; DestDir: "{app}"; Flags: ignoreversion

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "启动抖音智能发布中心"; Flags: nowait postinstall skipifsilent

[Code]
const
  RequiredVersion = '3.0.4.0';
  ProductUninstallKey = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{6D54D2A3-4431-4F35-8D43-A3405BC99F56}_is1';

function GetInstallDir(Param: String): String;
var
  Value: String;
begin
  Value := ExpandConstant('{param:HOTFIXTARGET|}');
  if Value <> '' then
  begin
    Result := RemoveBackslashUnlessRoot(Value);
    exit;
  end;
  if RegQueryStringValue(HKCU, ProductUninstallKey, 'InstallLocation', Value) and (Value <> '') then
  begin
    Result := RemoveBackslashUnlessRoot(Value);
    exit;
  end;
  Result := ExpandConstant('{localappdata}\Programs\DouyinPublisher');
end;

function InitializeSetup(): Boolean;
var
  TargetExe: String;
  ExistingVersion: String;
begin
  Result := False;
  TargetExe := AddBackslash(GetInstallDir('')) + '{#MyAppExeName}';
  if not FileExists(TargetExe) then
  begin
    MsgBox('没有找到已安装的抖音智能发布中心：' + TargetExe + #13#10 +
      '请先安装正式版 v3.0.4。', mbError, MB_OK);
    exit;
  end;
  if not GetVersionNumbersString(TargetExe, ExistingVersion) then
  begin
    MsgBox('无法读取现有主程序版本，补丁已停止。', mbError, MB_OK);
    exit;
  end;
  if ExistingVersion <> RequiredVersion then
  begin
    MsgBox('此补丁仅适用于 v3.0.4。当前主程序版本：' + ExistingVersion, mbError, MB_OK);
    exit;
  end;
  Result := True;
end;
