<#
check_version_twin.ps1 - one corpus, one version.

WHY: the repository now carries two implementations of one language, the VBA
reference in src/ and the Cargo workspace (core/, cli/), and one corpus in
scripts/ that both are held to. A version is the corpus's, not a product's:
a product that has not caught up is absent from a release's artifact list,
it never gets a number of its own. So VLA_RELEASE_VERSION in src/VLA.bas and
[workspace.package] version in Cargo.toml must be the same string, and this
check fails when they drift. Both are bumped by hand at release (DI.3a) and
nowhere else; release.ps1 runs every check_*.ps1 before it will tag, so a
release with the twins apart cannot be cut.

check_hash_twin.ps1's shape: two sources of one fact, read without Excel,
compared, and the comparison is the whole check.

House style (tools/check_*.ps1): PowerShell 5.1, host-independent, nothing
discovered that is not also pinned. Exit 0 clean, exit 1 with the two values.

Usage:  powershell -File tools\check_version_twin.ps1
        powershell -File tools\check_version_twin.ps1 -VbaPath x -CargoPath y   # testability
#>
param(
    [string]$VbaPath = '',
    [string]$CargoPath = ''
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$vba   = if ($VbaPath   -ne '') { $VbaPath }   else { Join-Path $root 'src/VLA.bas' }
$cargo = if ($CargoPath -ne '') { $CargoPath } else { Join-Path $root 'Cargo.toml' }

$failures = New-Object System.Collections.Generic.List[string]

# ---- the VBA side: the one Public Const that names the release ----
$vbaVersion = $null
foreach ($line in Get-Content $vba) {
    if ($line -match '^\s*Public Const VLA_RELEASE_VERSION As String = "([^"]+)"') { $vbaVersion = $Matches[1]; break }
}
if (-not $vbaVersion) { $failures.Add("${vba}: no 'Public Const VLA_RELEASE_VERSION As String = `"...`"' line") }

# ---- the Cargo side: [workspace.package] version, and only that table's ----
$cargoVersion = $null
$table = ''
foreach ($line in Get-Content $cargo) {
    $t = $line.Trim()
    if ($t -match '^\[(.+)\]$') { $table = $Matches[1]; continue }
    if ($table -eq 'workspace.package' -and $t -match '^version\s*=\s*"([^"]+)"') { $cargoVersion = $Matches[1]; break }
}
if (-not $cargoVersion) { $failures.Add("${cargo}: no 'version = `"...`"' under [workspace.package]") }

# ---- the comparison, which is the whole check ----
if ($vbaVersion -and $cargoVersion -and $vbaVersion -ne $cargoVersion) {
    $failures.Add("VLA_RELEASE_VERSION is $vbaVersion but Cargo.toml's workspace version is $cargoVersion; bump both by hand at release (DI.3a)")
}
foreach ($v in @($vbaVersion, $cargoVersion)) {
    if ($v -and $v -notmatch '^\d+\.\d+\.\d+$') { $failures.Add("'$v' is not MAJOR.MINOR.PATCH (SD-14)") }
}

if ($failures.Count -gt 0) {
    Write-Host "FAIL: $($failures.Count) problem(s)"
    $failures | ForEach-Object { Write-Host "  - $_" }
    exit 1
}
Write-Host "OK: VLA_RELEASE_VERSION and Cargo.toml agree on $vbaVersion"
exit 0
