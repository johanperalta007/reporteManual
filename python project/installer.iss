[Setup]
AppName=Reporte Pricing
AppVersion=1.0
AppPublisher=ADL - JP
DefaultDirName={autopf}\Reporte Pricing
DefaultGroupName=Reporte Pricing
OutputDir=installer_output
OutputBaseFilename=Reporte_Pricing_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Files]
Source: "dist\Reporte Pricing\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Reporte Pricing"; Filename: "{app}\Reporte Pricing.exe"
Name: "{commondesktop}\Reporte Pricing"; Filename: "{app}\Reporte Pricing.exe"

[Run]
Filename: "{app}\Reporte Pricing.exe"; Description: "Abrir Reporte Pricing ahora"; Flags: nowait postinstall skipifsilent
