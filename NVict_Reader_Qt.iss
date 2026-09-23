; NVict Reader Installer Script - PySide6-editie
; Gebaseerd op NVict_Reader.iss (tkinter-build), aangepast voor de
; PyInstaller-onedir-map van de Qt-app (dist\NVict_Reader_Qt\...).
;
; Zelfde AppId als de tkinter-installer: zo herkent Windows/Inno een
; bestaande tkinter-installatie bij het bijwerken als "upgrade" en wordt
; die (via de bestaande [Code]-sectie hieronder) eerst netjes
; gedeïnstalleerd voordat de nieuwe versie erover heen wordt gezet -
; zelfde mechanisme als tussen tkinter-versies, nu ook over de techstack-
; wissel heen. De geïnstalleerde exe heet ook na installatie nog steeds
; "NVict Reader.exe" (met spatie, zoals de tkinter-build) zodat bestaande
; snelkoppelingen, bestandskoppelingen en de "sluit lopende versie af"-
; herkenning ongewijzigd blijven werken.
;
; Wordt aangeroepen door _nvict_build/release.py met /DVERSION=... en /Ssigntool=...
; Handmatig compileren kan ook: ISCC NVict_Reader_Qt.iss /DVERSION=3.0

#ifndef VERSION
  #define VERSION "3.0"
#endif

#define MyAppName "NVict Reader"
#define MyAppVersion VERSION
#define MyAppPublisher "NVict Service"
#define MyAppURL "https://www.nvict.nl"
#define MyAppExeName "NVict Reader.exe"

[Setup]
; --- SIGNING ---
; De setup.exe wordt NA Inno gesigned door release.py, ALS er een werkende
; signing-methode is geconfigureerd (SSL.com is opgezegd - zie
; _nvict_build/README.md voor het alternatief zodra dat gekozen is).
; De geembedde uninstaller wordt niet gesigned - zelfde gedrag als de
; tkinter-installer.

; APP INFORMATIE
; Zelfde AppId-GUID als NVict_Reader.iss - bewust, voor een schone
; in-place-upgrade over een bestaande tkinter-installatie heen.
AppId={{B8F3C9D2-1E4A-4F5B-9C3D-2A1B8E7F6C5D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

; INSTALLATIE DIRECTORIES
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes

; OUTPUT INSTELLINGEN
OutputDir=Output
OutputBaseFilename=NVict_Reader_Setup
Compression=lzma
SolidCompression=yes

; PRIVILEGES
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog
UsedUserAreasWarning=no

; ICONEN
SetupIconFile=favicon.ico
UninstallDisplayIcon={app}\favicon.ico

; VERSIE INFO (zichtbaar in Eigenschappen van het setup bestand)
VersionInfoVersion={#MyAppVersion}.0.0
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription=NVict Reader Setup
VersionInfoCopyright=Copyright (c) 2026 {#MyAppPublisher}
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}.0.0

; WIZARD SETTINGS
WizardStyle=modern

; SLUIT DRAAIENDE VERSIE AF VOOR INSTALLATIE (zowel een al-lopende Qt-
; versie als een eventueel nog draaiende tkinter-versie - beide zijn na
; installatie "NVict Reader.exe")
CloseApplications=yes
CloseApplicationsFilter=NVict Reader.exe
RestartApplications=no

[Languages]
Name: "dutch"; MessagesFile: "compiler:Languages\Dutch.isl"

[Files]
; MAIN APPLICATION - one-folder (onedir) PyInstaller-build van de Qt-app.
; De hoofd-exe wordt hernoemd naar "NVict Reader.exe" (met spatie, zelfde
; naam als de tkinter-build) zodat bestaande snelkoppelingen/koppelingen
; na een upgrade blijven werken; de rest van de PyInstaller-map (incl.
; _internal met alle DLL's/data) gaat mee zoals hij is.
Source: "dist\NVict_Reader_Qt\NVict_Reader_Qt.exe"; DestDir: "{app}"; DestName: "NVict Reader.exe"; Flags: ignoreversion
Source: "dist\NVict_Reader_Qt\*"; DestDir: "{app}"; Excludes: "NVict_Reader_Qt.exe"; Flags: ignoreversion recursesubdirs createallsubdirs

; ICONEN
Source: "favicon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "PDF_File_icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\favicon.ico"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Registry]
; --- 1. OPRUIMEN OUDE VERSIES (Schoonmaak) ---
Root: HKCU; Subkey: "Software\Classes\NVictReader.PDF"; Flags: deletekey uninsdeletekey dontcreatekey
Root: HKCU; Subkey: "Software\Classes\Applications\NVict_Reader.py"; Flags: deletekey uninsdeletekey dontcreatekey
Root: HKCU; Subkey: "Software\Classes\Applications\NVict Reader.exe"; Flags: deletekey uninsdeletekey dontcreatekey
Root: HKCU; Subkey: "Software\Classes\.pdf\OpenWithProgids"; ValueType: none; ValueName: "NVictReader.PDF"; Flags: deletevalue
Root: HKCU; Subkey: "Software\NVict Service"; Flags: deletekey uninsdeletekey dontcreatekey
Root: HKCU; Subkey: "Software\RegisteredApplications"; ValueType: none; ValueName: "NVictReader"; Flags: deletevalue

; --- 2. DE PROGID (Bestandstype definitie) ---
Root: HKA; Subkey: "Software\Classes\NVictReader.PDF"; ValueType: string; ValueData: "NVict Reader PDF Document"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\NVictReader.PDF\DefaultIcon"; ValueType: string; ValueData: "{app}\PDF_File_icon.ico,0"
Root: HKA; Subkey: "Software\Classes\NVictReader.PDF\shell\open\command"; ValueType: string; ValueData: """{app}\{#MyAppExeName}"" ""%1"""
Root: HKA; Subkey: "Software\Classes\NVictReader.PDF\shell\print"; ValueType: string; ValueData: "Afdrukken"
Root: HKA; Subkey: "Software\Classes\NVictReader.PDF\shell\print\command"; ValueType: string; ValueData: """{app}\{#MyAppExeName}"" --print ""%1"""

; --- 3. FILE ASSOCIATIONS (Koppeling maken) ---
Root: HKA; Subkey: "Software\Classes\.pdf\OpenWithProgids"; ValueType: string; ValueName: "NVictReader.PDF"; ValueData: ""; Flags: uninsdeletevalue

; --- 4. CAPABILITIES (Voor Windows Standaard Apps lijst) ---
Root: HKA; Subkey: "Software\NVict Service\NVict Reader\Capabilities"; ValueType: string; ValueName: "ApplicationName"; ValueData: "NVict Reader"
Root: HKA; Subkey: "Software\NVict Service\NVict Reader\Capabilities"; ValueType: string; ValueName: "ApplicationDescription"; ValueData: "Bekijk en bewerk PDF bestanden."
Root: HKA; Subkey: "Software\NVict Service\NVict Reader\Capabilities\FileAssociations"; ValueType: string; ValueName: ".pdf"; ValueData: "NVictReader.PDF"
Root: HKA; Subkey: "Software\RegisteredApplications"; ValueType: string; ValueName: "NVictReader"; ValueData: "Software\NVict Service\NVict Reader\Capabilities"; Flags: uninsdeletevalue

; --- 5. FRIENDLY NAME FIX ---
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}"; ValueType: string; ValueName: "FriendlyAppName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}"; ValueType: string; ValueName: "ApplicationCompany"; ValueData: "{#MyAppPublisher}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\DefaultIcon"; ValueType: string; ValueData: "{app}\{#MyAppExeName},0"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\shell\open\command"; ValueType: string; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Flags: uninsdeletekey

; --- 6. CONTEXT MENU (Rechtermuisknop optie) ---
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\NVictReader"; ValueType: string; ValueName: ""; ValueData: "Open met NVict Reader"
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\NVictReader"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\{#MyAppExeName},0"
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\NVictReader\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""

[Code]
// Code sectie ongewijzigd t.o.v. NVict_Reader.iss - deïnstalleert een
// bestaande installatie (tkinter of een eerdere Qt-versie, beide onder
// dezelfde AppId) automatisch bij een upgrade.
function GetUninstallString(): String;
var
  sUnInstPath: String;
  sUnInstallString: String;
begin
  sUnInstPath := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{#emit SetupSetting("AppId")}_is1');
  sUnInstallString := '';
  if not RegQueryStringValue(HKLM, sUnInstPath, 'UninstallString', sUnInstallString) then
    RegQueryStringValue(HKCU, sUnInstPath, 'UninstallString', sUnInstallString);
  Result := sUnInstallString;
end;

function IsUpgrade(): Boolean;
begin
  Result := (GetUninstallString() <> '');
end;

function UnInstallOldVersion(): Integer;
var
  sUnInstallString: String;
  iResultCode: Integer;
begin
  Result := 0;
  sUnInstallString := GetUninstallString();
  if sUnInstallString <> '' then begin
    sUnInstallString := RemoveQuotes(sUnInstallString);
    if Exec(sUnInstallString, '/SILENT /NORESTART /SUPPRESSMSGBOXES','', SW_HIDE, ewWaitUntilTerminated, iResultCode) then
      Result := 3
    else
      Result := 2;
  end else
    Result := 1;
end;

procedure CleanupOldMEIFolders();
var
  TempDir: String;
  FindRec: TFindRec;
  FolderPath: String;
begin
  // Verwijder oude PyInstaller _MEI extractiemappen uit temp
  // Deze veroorzaken "Failed to load Python DLL" bij updates
  TempDir := ExpandConstant('{tmp}\..');
  if FindFirst(TempDir + '\_MEI*', FindRec) then
  begin
    try
      repeat
        if (FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then
        begin
          FolderPath := TempDir + '\' + FindRec.Name;
          DelTree(FolderPath, True, True, True);
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if (CurStep=ssInstall) then
  begin
    // Ruim oude PyInstaller temp-mappen op
    CleanupOldMEIFolders();
    if (IsUpgrade()) then
    begin
      UnInstallOldVersion();
    end;
  end;
end;
