; DI.2 -- per-user installer for Frazaro. Copies the built add-in plus
; its two required companion files (prelude.vla, english.vla) under
; %AppData%\Frazaro -- no admin rights needed (PrivilegesRequired=lowest)
; -- and registers a real Windows uninstall entry that removes exactly
; what this installer placed.
;
; Also auto-registers the add-in with Excel via the OPEN/OPENn
; mechanism under HKCU\...\Excel\Options -- the same registry entry the
; Add-Ins dialog itself writes when you tick a box, so no separate
; "Browse to Frazaro.xlam and check it" step is needed. Live-verified
; on this machine (not assumed from documentation): wrote a throwaway
; OPEN value pointing at a real, minimal test .xlam, launched a fresh
; Excel via COM automation, and confirmed Application.AddIns lists it
; with Installed=True (matching the built-in add-ins that are NOT
; ticked, which read Installed=False) -- then confirmed OPEN1 (a
; second, numbered slot) works identically and that Excel deduplicates
; by path rather than double-loading. Cleaned up afterwards; no
; standing state was left from that test.
;
; Deliberately does NOT touch the Trust Center macro-security setting
; (needed only for Compile, never for Interpret) -- that one setting is
; specifically walled off by Microsoft from installer automation for
; individual users, on purpose (it's the setting a malicious installer
; would most want to flip silently), so no installer can script around
; it. It stays the one honest manual step (POST_INSTALL.txt, shown on
; the wizard's finish page) -- same doctrine DEPLOY.md already applies
; to locking the VBA project.
;
; Build with: ISCC.exe installer\Frazaro.iss  (or open in the Inno
; Setup IDE and press Compile). Requires Frazaro_English.xlam AND
; version.iss to already exist (run VlaBuildAddin first - it writes
; both) and scripts\prelude.vla / scripts\polyglotta\english.vla to
; exist (they ship in the
; repo already).
; Output: installer\output\FrazaroSetup.exe -- sign it afterwards with
; sign_installer.ps1 (see DEPLOY.md's "Code signing" section for why
; the .xlam itself isn't signed this way).

#define MyAppName "Frazaro"
; DI.3a: MyAppVersion now comes from version.iss (VLA_Build.bas's
; VlaWriteInstallerVersion, generated from VLA.VLA_RELEASE_VERSION on
; every build) instead of a literal here - the literal sat at "1.0",
; untouched, for this installer's entire life, because nothing forced a
; human to remember two places. Run VlaBuildAddin at least once before
; building this installer, or version.iss will not exist yet.
#include "version.iss"
#define MyAppPublisher "Frazaro"

[Setup]
AppId={{7329BF19-5722-4CE6-B885-E6B61B6615C7}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={userappdata}\Frazaro
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=output
OutputBaseFilename=FrazaroSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
InfoAfterFile=POST_INSTALL.txt
UninstallDisplayName={#MyAppName}

[Files]
; EDITIONMANIFEST.2/3: the build's own output renamed Frazaro.xlam ->
; Frazaro_English.xlam (this installer stays English-only, per DEPLOY.md
; - EDITION-ESPANOL's own job to give it a per-edition pass), and
; english.vla moved into scripts\polyglotta\ (owner, "keep language
; files together") - the installed DestDir stays flat scripts\ either
; way, unrelated to this repo's own source-tree organization.
Source: "..\Frazaro_English.xlam"; DestDir: "{app}"; DestName: "Frazaro.xlam"; Flags: ignoreversion
Source: "..\scripts\prelude.vla"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "..\scripts\polyglotta\english.vla"; DestDir: "{app}\scripts"; Flags: ignoreversion

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Code]
// Every installed Office version that has an Excel\Options key -- more
// than one can coexist (e.g. a Click-to-Run 16.0 alongside a legacy
// MSI install), so registration/unregistration loops over all of them
// rather than assuming 16.0.
function GetOfficeExcelVersions(): TArrayOfString;
var
  Names: TArrayOfString;
  Versions: TArrayOfString;
  I: Integer;
begin
  SetArrayLength(Versions, 0);
  if RegGetSubkeyNames(HKCU, 'Software\Microsoft\Office', Names) then
  begin
    for I := 0 to GetArrayLength(Names) - 1 do
    begin
      if RegKeyExists(HKCU, 'Software\Microsoft\Office\' + Names[I] + '\Excel\Options') then
      begin
        SetArrayLength(Versions, GetArrayLength(Versions) + 1);
        Versions[GetArrayLength(Versions) - 1] := Names[I];
      end;
    end;
  end;
  Result := Versions;
end;

function OpenSlotName(Index: Integer): String;
begin
  // The first slot is the bare, unnumbered "OPEN"; later ones are
  // "OPEN1", "OPEN2", ... -- live-verified naming, not assumed.
  if Index = 0 then
    Result := 'OPEN'
  else
    Result := 'OPEN' + IntToStr(Index);
end;

// Writes XlamPath into the next free OPEN/OPENn slot under this Office
// version's Excel\Options key -- or does nothing if XlamPath is
// already registered in some slot, so re-running the installer (an
// upgrade, a repair) never mints a duplicate entry.
procedure RegisterExcelAddin(const Version, XlamPath: String);
var
  OptKey: String;
  Existing: String;
  I: Integer;
begin
  OptKey := 'Software\Microsoft\Office\' + Version + '\Excel\Options';
  // 0..50: a generous safety bound, inlined because this Pascal Script
  // dialect rejected a local const block here (compile-tested).
  for I := 0 to 50 do
  begin
    if RegQueryStringValue(HKCU, OptKey, OpenSlotName(I), Existing) then
    begin
      if Existing = XlamPath then
        Exit; // already registered in this slot -- nothing to do
    end
    else
    begin
      RegWriteStringValue(HKCU, OptKey, OpenSlotName(I), XlamPath);
      Exit;
    end;
  end;
end;

// Removes exactly the slot(s) that hold XlamPath, leaving every other
// OPEN/OPENn entry (any other add-in registered the same way) intact.
procedure UnregisterExcelAddin(const Version, XlamPath: String);
var
  OptKey: String;
  Existing: String;
  I: Integer;
begin
  OptKey := 'Software\Microsoft\Office\' + Version + '\Excel\Options';
  for I := 0 to 50 do
  begin
    if RegQueryStringValue(HKCU, OptKey, OpenSlotName(I), Existing) then
    begin
      if Existing = XlamPath then
        RegDeleteValue(HKCU, OptKey, OpenSlotName(I));
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  Versions: TArrayOfString;
  I: Integer;
  XlamPath: String;
begin
  if CurStep = ssPostInstall then
  begin
    XlamPath := ExpandConstant('{app}\Frazaro.xlam');
    Versions := GetOfficeExcelVersions();
    for I := 0 to GetArrayLength(Versions) - 1 do
      RegisterExcelAddin(Versions[I], XlamPath);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Versions: TArrayOfString;
  I: Integer;
  XlamPath: String;
begin
  if CurUninstallStep = usUninstall then
  begin
    XlamPath := ExpandConstant('{app}\Frazaro.xlam');
    Versions := GetOfficeExcelVersions();
    for I := 0 to GetArrayLength(Versions) - 1 do
      UnregisterExcelAddin(Versions[I], XlamPath);
  end;
end;
