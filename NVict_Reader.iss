; NVict Reader Installer Script met RELATIEVE paden
; Versie 2.0 — gemigreerd naar _nvict_build pipeline
;
; Wordt aangeroepen door _nvict_build/release.py met /DVERSION=... en /Ssigntool=...
; Handmatig compileren kan ook: ISCC NVict_Reader.iss /DVERSION=2.3

#ifndef VERSION
  #define VERSION "2.3"
#endif

#define MyAppName "NVict Reader"
#define MyAppVersion VERSION
#define MyAppPublisher "NVict Service"
#define MyAppURL "https://www.nvict.nl"
#define MyAppExeName "NVict Reader.exe"

[Setup]
; --- SIGNING ---
; De setup.exe wordt NA Inno gesigned door release.py (CodeSignTool).
; De geembedde uninstaller wordt niet gesigned — zelfde gedrag als het
; oude Release_Complete_v3_0.bat script.

; APP INFORMATIE
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

; SLUIT DRAAIENDE VERSIE AF VOOR INSTALLATIE
CloseApplications=yes
CloseApplicationsFilter=NVict Reader.exe
RestartApplications=no

[Languages]
Name: "dutch"; MessagesFile: "compiler:Languages\Dutch.isl"

[Files]
; MAIN APPLICATION FILES - Relatief pad naar je al ondertekende EXE!
Source: "dist\NVict_Reader.exe"; DestDir: "{app}"; DestName: "NVict Reader.exe"; Flags: ignoreversion

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
// Code sectie ongewijzigd voor automatische de-installatie van oude versies
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
