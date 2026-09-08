<#
build_installer.ps1 - build and sign FrazaroSetup.exe: one command, verified.

Added 2026-09-08 after 0.5.0, 0.5.1 and 0.5.2 all shipped the two .xlam
add-ins and no installer. Nothing was broken; nothing was even wrong. The
installer simply was not named anywhere a release had to pass through:
DEPLOY.md's day-end sequence did not list it, and tools/release.ps1 did not
know it existed, so it could not miss it. This script is now the "build and
sign the installer" step of that sequence, and release.ps1's check 9 refuses
to publish unless this script's output is present, fresh, versioned, and
signed. A step that a script refuses without cannot be forgotten.

What it does, in order, refusing by name at the first failure:
  1. Frazaro_English.xlam exists beside the repository root and is newer
     than every file under src/ and scripts/ - release.ps1's own freshness
     rule, applied here because the installer EMBEDS that file (Frazaro.iss,
     [Files]) and so can never legitimately be fresher than it.
  2. installer\version.iss exists and declares the same version as
     VLA_RELEASE_VERSION in src/VLA.bas. VlaBuildAddin writes version.iss
     (DI.3a); a mismatch means the constant was bumped and the add-in was
     not rebuilt afterwards - the exact slip that would ship an installer
     whose Windows DisplayVersion lies.
  3. ISCC.exe (Inno Setup 6) is where DEPLOY.md documents it, or -IsccPath.
  4. Runs ISCC over installer\Frazaro.iss -> installer\output\FrazaroSetup.exe.
  5. Runs installer\sign_installer.ps1 (self-signed, CN=Frazaro Dev Signing -
     see DEPLOY.md "Code signing" for what that does and does not buy).
  6. Verifies the result: the .exe exists, is newer than the add-in it
     embeds, and carries that certificate's Authenticode signature. Unlike
     the VBA project's lock and signature, an .exe signature IS readable
     (Get-AuthenticodeSignature), so this is a check, not an attestation.

Host-independent: no Excel, no COM, no network. Safe to run repeatedly.
installer\output\ is a build product and is never committed.

Usage:  powershell -File installer\build_installer.ps1 [-IsccPath <path>]
Exit code: 0 built and signed; 1 with the first failing precondition named.
#>
param(
    [string]$IsccPath = "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
)
$ErrorActionPreference = 'Stop'
$root       = Split-Path -Parent $PSScriptRoot
$xlam       = Join-Path $root 'Frazaro_English.xlam'
$iss        = Join-Path $PSScriptRoot 'Frazaro.iss'
$versionIss = Join-Path $PSScriptRoot 'version.iss'
$exe        = Join-Path $PSScriptRoot 'output\FrazaroSetup.exe'
$signer     = 'CN=Frazaro Dev Signing'

function Fail([string]$msg) { Write-Host "NOT BUILT - $msg"; exit 1 }

# 1. the add-in the installer embeds: present and fresh
if (-not (Test-Path $xlam)) { Fail "Frazaro_English.xlam not found beside the repository root - run VlaBuildAddin first" }
$xlamItem = Get-Item $xlam
$newest = Get-ChildItem (Join-Path $root 'src'), (Join-Path $root 'scripts') -Recurse -File |
          Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($newest -and $xlamItem.LastWriteTime -lt $newest.LastWriteTime) {
    Fail "Frazaro_English.xlam ($($xlamItem.LastWriteTime)) is older than $($newest.FullName) ($($newest.LastWriteTime)) - rebuild the add-in before the installer"
}

# 2. version.iss agrees with the source constant
$constLine = Get-Content (Join-Path $root 'src\VLA.bas') |
             Where-Object { $_ -match 'Public Const VLA_RELEASE_VERSION As String = "([^"]+)"' } | Select-Object -First 1
$srcVer = if ($constLine -match '"([^"]+)"') { $Matches[1] } else { '' }
if (-not (Test-Path $versionIss)) { Fail "installer\version.iss is missing - VlaBuildAddin writes it (DI.3a); run the build" }
$m = [regex]::Match((Get-Content $versionIss -Raw), '#define MyAppVersion "([^"]+)"')
$issVer = if ($m.Success) { $m.Groups[1].Value } else { '' }
if ($issVer -ne $srcVer) {
    Fail "installer\version.iss declares '$issVer' but src/VLA.bas says '$srcVer' - the add-in was not rebuilt after the bump; run VlaBuildAddin, then this script"
}

# 3. the compiler
if (-not (Test-Path $IsccPath)) { Fail "ISCC.exe not found at $IsccPath - install Inno Setup 6, or pass -IsccPath" }

# 4. build
Write-Host "Building FrazaroSetup.exe $issVer from $iss ..."
& $IsccPath /Q $iss
if ($LASTEXITCODE -ne 0) { Fail "ISCC.exe exited $LASTEXITCODE" }
if (-not (Test-Path $exe)) { Fail "ISCC reported success but $exe is not there - check OutputDir/OutputBaseFilename in Frazaro.iss" }

# 5. sign
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'sign_installer.ps1') -InstallerPath $exe
if ($LASTEXITCODE -ne 0) { Fail "sign_installer.ps1 exited $LASTEXITCODE" }

# 6. verify what was produced, the same way release.ps1 will
$exeItem = Get-Item $exe
if ($exeItem.LastWriteTime -lt $xlamItem.LastWriteTime) { Fail "the built installer is older than the add-in it embeds - clock skew? rebuild" }
$sig  = Get-AuthenticodeSignature $exe
$subj = if ($sig.SignerCertificate) { $sig.SignerCertificate.Subject } else { '' }
if ($sig.Status -eq 'NotSigned' -or $sig.Status -eq 'HashMismatch' -or $subj -ne $signer) {
    Fail "signature check failed: status $($sig.Status), signer '$subj' (expected $signer)"
}

Write-Host ("BUILT AND SIGNED: {0}" -f $exe)
Write-Host ("  {0:N0} bytes, version {1}, signer {2}" -f $exeItem.Length, $issVer, $subj)
Write-Host ("  signature status {0} - '{1}' is the expected, honest status for a self-signed certificate (DEPLOY.md, Code signing)" -f $sig.Status, $sig.Status)
exit 0
