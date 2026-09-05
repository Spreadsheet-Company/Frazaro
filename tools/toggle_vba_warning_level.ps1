# DI.1 Pass 2 -- closes the one verification gap generate_vba_signing_cert.ps1
# left open: whether the "Enable Content / trust this publisher" flow
# actually behaves as documented for a self-signed certificate.
#
# The blocker on a dev machine is VBAWarnings=1 ("Enable all macros" --
# no warning bar ever shows, signed or not), confirmed live on this
# machine rather than assumed. This script flips it to 2 ("Disable
# with notification") so a real trust-bar test is possible WITHOUT a
# different machine, then flips it back afterwards -- a developer
# should not be left running with macro warnings on by accident.
#
# Usage:
#   pwsh -File tools\toggle_vba_warning_level.ps1            # flip to test mode
#   ... do the manual signing step, open the signed .xlam, observe ...
#   pwsh -File tools\toggle_vba_warning_level.ps1 -Restore   # flip back

param(
    [switch]$Restore
)

$ErrorActionPreference = 'Stop'
$sec = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$stateFile = "$env:TEMP\vla_vbawarnings_prior.txt"

if (-not $Restore) {
    if (-not (Test-Path $sec)) {
        Write-Error "Registry key not found: $sec -- is Office 16.0 installed under this user profile?"
        exit 1
    }
    $current = (Get-ItemProperty $sec -Name VBAWarnings -ErrorAction SilentlyContinue).VBAWarnings
    if ($null -eq $current) {
        Write-Error "VBAWarnings has no value set -- unexpected. Not changing anything; check Trust Center > Macro Settings by hand."
        exit 1
    }
    Set-Content -Path $stateFile -Value $current -Encoding ascii
    Set-ItemProperty -Path $sec -Name VBAWarnings -Value 2 -Type DWord
    Write-Host "VBAWarnings: $current -> 2 (Disable with notification). Prior value saved to $stateFile."
    Write-Host ""
    Write-Host "Now:"
    Write-Host "  1. Run tools\generate_vba_signing_cert.ps1 if you haven't already."
    Write-Host "  2. Sign Frazaro.xlam's VBA project in the VBE (Tools > Digital Signature)."
    Write-Host "  3. Close Excel entirely, then reopen the signed file fresh."
    Write-Host "  4. Look for the security bar: does it offer 'Enable Content' AND a"
    Write-Host "     distinct 'Trust all documents from this publisher' option (not just"
    Write-Host "     a one-time Enable Content)?"
    Write-Host "  5. Close and reopen a second time: does the bar NOT reappear?"
    Write-Host "  6. Run this script again with -Restore when done."
} else {
    if (-not (Test-Path $stateFile)) {
        Write-Error "No saved prior value at $stateFile -- either -Restore was run without a prior flip, or the state file was removed. Set Trust Center > Macro Settings back to 'Enable all macros' by hand if this machine should have it, or check File > Options > Trust Center yourself before assuming."
        exit 1
    }
    $prior = (Get-Content -Path $stateFile -Raw).Trim()
    Set-ItemProperty -Path $sec -Name VBAWarnings -Value ([int]$prior) -Type DWord
    Remove-Item -Path $stateFile -Force
    Write-Host "VBAWarnings restored to $prior."
}
