#define MyAppName "Prenomina"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "Humanio"
#define MyAppExeName "Prenomina.exe"

[Setup]
AppId={{8B18F3E0-2255-4676-A710-6B2E5EB9C780}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=output
OutputBaseFilename=Prenomina Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear icono en el escritorio"; GroupDescription: "Accesos directos:"; Flags: checkedonce

[Files]
Source: "..\dist\Prenomina.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\.env.example"; DestDir: "{userdocs}\Prenomina"; DestName: ".env.example"; Flags: ignoreversion

[Dirs]
Name: "{userdocs}\Prenomina"
Name: "{userdocs}\Prenomina\input"
Name: "{userdocs}\Prenomina\output"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall skipifsilent
