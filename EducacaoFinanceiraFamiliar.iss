#define MyAppName "Educação Financeira Familiar"
#define MyAppVersion "2.1.1"
#define MyAppPublisher "Educação Financeira Familiar"
#define MyAppExeName "EducacaoFinanceiraFamiliar.exe"

[Setup]
AppId={{A9B5DF10-26D4-4F93-95A4-72A38D7C8E71}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Educacao Financeira Familiar
DefaultGroupName={#MyAppName}
OutputDir=installer-output
OutputBaseFilename=EducacaoFinanceiraFamiliar-Setup-x64
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\{#MyAppExeName}

[Files]
Source: "dist\EducacaoFinanceiraFamiliar.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Educação Financeira Familiar"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Educação Financeira Familiar"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na área de trabalho"; GroupDescription: "Atalhos adicionais:"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir Educação Financeira Familiar"; Flags: postinstall nowait skipifsilent
