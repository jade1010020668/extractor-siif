; Instalador de Actas Privadas (Inno Setup 6). Lo compila build.ps1.
; Se instala para el usuario actual: NO requiere permisos de administrador.
#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#ifndef DistDir
  #error Falta /DDistDir=...
#endif
#ifndef OutDir
  #define OutDir "."
#endif

[Setup]
AppId={{7C1E3F52-4B7A-4D3A-9C55-2F0B6A1D8E41}
AppName=Actas Privadas
AppVersion={#AppVersion}
AppPublisher=Actas Privadas
DefaultDirName={localappdata}\Programs\ActaPrivada
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir={#OutDir}
OutputBaseFilename=ActaPrivada-Setup
Compression=lzma2/ultra64
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
UninstallDisplayName=Actas Privadas

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "{#DistDir}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{autoprograms}\Actas Privadas"; Filename: "{app}\ActaPrivada.cmd"; WorkingDir: "{app}"; IconFilename: "{app}\python\python.exe"
Name: "{autodesktop}\Actas Privadas"; Filename: "{app}\ActaPrivada.cmd"; WorkingDir: "{app}"; IconFilename: "{app}\python\python.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\ActaPrivada.cmd"; Description: "Abrir Actas Privadas ahora"; Flags: postinstall nowait skipifsilent shellexec

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
; Los modelos y la configuración (%LOCALAPPDATA%\ActaPrivada) NO se borran al desinstalar.
