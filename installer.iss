#define MyAppName "Gamboa Transcriptor"
#define MyAppPublisher "Gamboa Desarrollos"
#define MyAppExeName "GamboaTranscriptor.exe"
#define MyAppVersion "1.0.0"

[Setup]
AppId={{A28F9803-5681-489C-8E31-1FCBAEFB32E1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\GamboaTranscriptor
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=no
OutputDir=installer_output
OutputBaseFilename=GamboaTranscriptor_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
SetupIconFile=assets\icon.ico
PrivilegesRequired=admin

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Files]
Source: "dist\GamboaTranscriptor\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el escritorio"; GroupDescription: "Accesos directos:"; Flags: unchecked

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID = wpFinished then
  begin
    WizardForm.FinishedLabel.Caption :=
      WizardForm.FinishedLabel.Caption + #13#10#13#10 +
      'Para generar minutas, instale Ollama y descargue qwen3:8b.';
  end;
end;
