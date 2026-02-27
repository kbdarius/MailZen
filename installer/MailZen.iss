; MailZen InnoSetup Installer Script
; Requires Inno Setup 6+ (https://jrsoftware.org/isinfo.php)
;
; Build steps:
;   1. Run build-installer.ps1  (publishes app + compiles this script)
;   2. Resulting installer: installer\MailZenSetup.exe

#define MyAppName "MailZen"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "MailZen"
#define MyAppURL "https://github.com/mailzen"
#define MyAppExeName "MailZen.exe"

[Setup]
AppId={{B7A3E2F1-4D5C-6E7F-8A9B-0C1D2E3F4A5B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\installer
OutputBaseFilename=MailZenSetup
SetupIconFile=..\src\EmailManage.App\Resources\mailzen.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}
LicenseFile=
; If no icon file exists yet, comment out SetupIconFile line above

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; All published files from the self-contained output
Source: "..\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Comment: "Smart Email Cleanup"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; Comment: "Smart Email Cleanup"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
// Check if Outlook is installed before proceeding
function InitializeSetup(): Boolean;
var
  OutlookPath: String;
begin
  Result := True;
  if not RegQueryStringValue(HKEY_LOCAL_MACHINE,
    'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE',
    '', OutlookPath) then
  begin
    if MsgBox('Microsoft Outlook Desktop was not detected on this computer.' + #13#10 + #13#10 +
              'MailZen requires Outlook Desktop to manage your email.' + #13#10 + #13#10 +
              'Do you want to continue installing anyway?',
              mbConfirmation, MB_YESNO) = IDNO then
    begin
      Result := False;
    end;
  end;
end;
