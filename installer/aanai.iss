; Inno Setup script for AanAI
; Builds a single installer .exe

#define MyAppName "AanAI"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "AanAI"
#define MyAppURL "https://example.com"
#define MyAppExeName "AanAI.exe"

[Setup]
AppId={{A2F3B6A2-9D7A-4E8C-9E2E-1B1D6D0AF9B1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableDirPage=no
DisableProgramGroupPage=no
OutputDir=..\installer\dist
OutputBaseFilename={#MyAppName}-Setup
Compression=lzma
SolidCompression=yes
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
SetupLogging=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Copy everything from the PyInstaller dist folder into the app directory
Source: "..\dist\AanAI\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion
; Optionally include an .env.example
Source: "..\.env.example"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist; Tasks: createenv

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked
Name: "createenv"; Description: "Install a sample .env file"; GroupDescription: "Optional components:"; Flags: unchecked

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
