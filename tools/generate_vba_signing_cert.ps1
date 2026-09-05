# DI.1 Pass 2 -- generate (or confirm) the self-signed certificate used
# to sign Frazaro's VBA project.
#
# Deliberately a SEPARATE certificate from installer/sign_installer.ps1's
# "CN=Frazaro Dev Signing" -- that one Authenticode-signs the installer
# .exe; this one signs the VBA project itself, a different mechanism
# with a different trust surface (Excel's macro-security model, not
# Windows SmartScreen). Sharing one cert across both would blur which
# trust conversation a given signature is even answering.
#
# What this script does NOT do, and cannot do: select this certificate
# in the VBE's Tools -> Digital Signature dialog and actually sign a
# project. VBIDE's VBProject COM interface exposes no signing member at
# all (enumerated directly: Application, BuildFileName, Collection,
# Description, FileName, HelpContextID, HelpFile, MakeCompiledFile,
# Mode, Name, Parent, Protection, References, SaveAs, Saved, Type, VBE
# -- nothing else), the same "VBA offers no API for this" limitation
# VLA_Build.bas already documents for project locking. That step is
# manual, every release, in the VBE itself -- see DEPLOY.md's "Signing
# the VBA project" section for exactly what to click and what to expect
# afterward.
#
# Usage: pwsh -File tools\generate_vba_signing_cert.ps1

$ErrorActionPreference = 'Stop'

$subject = 'CN=Frazaro VBA Signing'
$existing = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.Subject -eq $subject -and $_.NotAfter -gt (Get-Date) } |
    Select-Object -First 1

if ($existing) {
    $cert = $existing
    Write-Host "Reusing existing certificate:"
} else {
    $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $subject `
        -CertStoreLocation Cert:\CurrentUser\My -KeyUsage DigitalSignature `
        -KeyAlgorithm RSA -KeyLength 2048 -NotAfter (Get-Date).AddYears(5)
    Write-Host "Minted new certificate:"
}

Write-Host "  Subject:    $($cert.Subject)"
Write-Host "  Thumbprint: $($cert.Thumbprint)"
Write-Host "  Expires:    $($cert.NotAfter)"
Write-Host ""
Write-Host "Next (manual, every release -- see DEPLOY.md 'Signing the VBA project'):"
Write-Host "  1. Open the built add-in's VBA project in the VBE (Alt+F11)."
Write-Host "  2. Tools -> Digital Signature... -> Choose Certificate..."
Write-Host "  3. Select the certificate with subject '$subject' above."
Write-Host "  4. OK, then save the workbook."
