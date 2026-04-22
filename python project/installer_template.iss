; ============================================================
; PLANTILLA INNO SETUP - Reutilizable para apps Python/PyInstaller
; Reemplaza los valores marcados con <ESTO> según tu proyecto
; ============================================================

[Setup]
AppName=<NOMBRE DE LA APP>
AppVersion=<VERSION, ej: 1.0.0>
AppPublisher=<TU NOMBRE O EMPRESA>
DefaultDirName={autopf}\<NOMBRE DE LA APP>
DefaultGroupName=<NOMBRE DE LA APP>
OutputDir=installer_output
OutputBaseFilename=<NOMBRE>_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

; Opcional: agrega un ícono al instalador mismo (.ico obligatorio, no .png)
; SetupIconFile=images\icono.ico

; Opcional: imagen de bienvenida (izquierda del wizard, 164x314 px)
; WizardImageFile=images\wizard_banner.bmp

[Files]
; Copia todo el contenido de la carpeta generada por PyInstaller (--onedir)
Source: "dist\<NOMBRE DE LA APP>\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Acceso directo en el menú inicio
Name: "{group}\<NOMBRE DE LA APP>"; Filename: "{app}\<NOMBRE DE LA APP>.exe"

; Acceso directo en el escritorio
Name: "{commondesktop}\<NOMBRE DE LA APP>"; Filename: "{app}\<NOMBRE DE LA APP>.exe"

[Run]
; Opción para abrir la app al terminar la instalación
Filename: "{app}\<NOMBRE DE LA APP>.exe"; Description: "Abrir <NOMBRE DE LA APP> ahora"; Flags: nowait postinstall skipifsilent

; ============================================================
; NOTAS PARA MacOS:
; Inno Setup es solo para Windows. En Mac usá:
;   - create-dmg (npm): genera un .dmg con drag-to-Applications
;   - pyinstaller --onefile o --onedir genera el .app
;   - Comando: npx create-dmg 'dist/MiApp.app' dist/
; ============================================================
