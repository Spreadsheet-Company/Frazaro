# DI.1 -- self-sign the compiled installer executable.
#
# Targets the installer (FrazaroSetup.exe), not Frazaro.xlam: Office's
# OOXML package format has no registered Authenticode Subject Interface
# Package on Windows, so Set-AuthenticodeSignature simply cannot sign
# an .xlam at all -- confirmed empirically ("The form specified for the
# subject is not one supported or known by the specified trust
# provider"), not a missing-cert or missing-tool problem. A real .exe is
# a standard, fully-supported Authenticode target, so signing moves here
# instead. See DEPLOY.md's "Code signing" section for what a self-signed
# certificate does and does not buy (short version: identity continuity
# across builds and tamper-evidence: it does NOT clear SmartScreen or
# Excel's own macro-trust prompts without the certificate being
# explicitly trusted on the target machine first).
#
# Usage: pwsh -File installer\sign_installer.ps1 [-InstallerPath <path>]
# Run after ISCC.exe has produced the installer.

param(
    [string]$InstallerPath = "$PSScriptRoot\output\FrazaroSetup.exe"
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $InstallerPath)) {
    Write-Error "Installer not found at $InstallerPath -- build it first (ISCC.exe installer\Frazaro.iss)."
    exit 1
}

$subject = 'CN=Frazaro Dev Signing'
$existing = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
    Where-Object { $_.Subject -eq $subject -and $_.NotAfter -gt (Get-Date) } |
    Select-Object -First 1

if ($existing) {
    $cert = $existing
} else {
    $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $subject `
        -CertStoreLocation Cert:\CurrentUser\My -KeyUsage DigitalSignature `
        -KeyAlgorithm RSA -KeyLength 2048 -NotAfter (Get-Date).AddYears(5)
    Write-Host "Minted new signing certificate: $($cert.Thumbprint)"
}

$sig = Set-AuthenticodeSignature -FilePath $InstallerPath -Certificate $cert -HashAlgorithm SHA256

if ($sig.Status -eq 'NotSigned') {
    Write-Error "Signing did not apply: $($sig.StatusMessage)"
    exit 1
}

# 'UnknownError' / 'NotTrusted' here is the expected, honest outcome for
# a self-signed cert with no chain to a trusted root -- the signature is
# real, the trust conversation with Windows/the target machine is not.
Write-Host "Signed $InstallerPath"
Write-Host "  Certificate: $($cert.Subject) ($($cert.Thumbprint))"
Write-Host "  Status: $($sig.Status) -- $($sig.StatusMessage)"
exit 0
