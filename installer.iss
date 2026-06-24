; BhugtanEase Installer Script (Inno Setup)
; ----------------------------------------------------------------
; Yeh script "BhugtanEase.exe" (jo build.bat se banti hai) ko
; ek professional Setup.exe installer mein convert karta hai.
; Start Menu shortcut, Desktop shortcut, Uninstaller sab include hain.
; ----------------------------------------------------------------

#define MyAppName "BhugtanEase"
#define MyAppVersion "1.0"
#define MyAppPublisher "BhugtanEase"
#define MyAppExeName "BhugtanEase.exe"

[Setup]
AppId={{B6E5B6B5-2C2A-4B7E-9C2D-BHUGTANEASE01}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; Output installer file ka naam aur location
OutputDir=installer_output
OutputBaseFilename=BhugtanEase_Setup
SetupIconFile=BhugtanEase.ico
Compression=lzma
SolidCompression=yes
; Windows 7/8/10/11 sab support — purana wizard style, no admin-only restriction
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
; Windows 7 SP1 aur upar
MinVersion=6.1sp1

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Desktop par shortcut banao"; GroupDescription: "Additional icons:"

[Files]
Source: "BhugtanEase.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; IconFilename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent
