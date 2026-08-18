#define AppName "Cabecalho Logo Inserter"
#define AppExeName "Cabecalho.LogoInserter.exe"

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

[Setup]
AppId={{D29762F5-29EA-43BF-8874-CF306E36429A}

AppName={#AppName}
AppVersion={#AppVersion}

DefaultDirName={localappdata}\Programs\{#AppName}

DefaultGroupName={#AppName}

PrivilegesRequired=lowest

DisableProgramGroupPage=yes
UsePreviousAppDir=yes

LicenseFile=docs\TERMS.txt

ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

WizardStyle=modern

DisableWelcomePage=no
DisableReadyPage=no
DisableFinishedPage=no

OutputDir=dist
OutputBaseFilename=Cabecalho.LogoInserter-Setup-{#AppVersion}

Compression=lzma2
SolidCompression=yes

Uninstallable=yes

UninstallDisplayName={#AppName}
UninstallDisplayIcon={app}\{#AppExeName}

CreateUninstallRegKey=yes

CloseApplications=yes
RestartApplications=no

VersionInfoDescription={#AppName}
VersionInfoProductName={#AppName}
VersionInfoProductVersion={#AppVersion}

[Tasks]

Name: "desktopicon"; \
    Description: "Criar atalho na Área de Trabalho"; \
    GroupDescription: "Atalhos adicionais:"; \
    Flags: unchecked

[Files]

Source: "publish\*"; \
    DestDir: "{app}"; \
    Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]

Name: "{autoprograms}\{#AppName}"; \
    Filename: "{app}\{#AppExeName}"; \
    WorkingDir: "{app}"

Name: "{autodesktop}\{#AppName}"; \
    Filename: "{app}\{#AppExeName}"; \
    WorkingDir: "{app}"; \
    Tasks: desktopicon

[Run]

Filename: "{app}\{#AppExeName}"; \
    Description: "Executar {#AppName}"; \
    WorkingDir: "{app}"; \
    Flags: nowait postinstall skipifsilent

[UninstallDelete]

Type: dirifempty; Name: "{app}"