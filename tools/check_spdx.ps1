<#
check_spdx.ps1 - SIG.0's mechanical check.

The licence map is REUSE.toml (reuse.software): path globs -> SPDX
identifiers, so licence territories need no directory moves. This script
verifies three things, none of which needs Excel or any host:

  1. Every tracked source file resolves to exactly one licence through
     REUSE.toml's [[annotations]] tables (later tables win, matching the
     REUSE specification's ordering rule).
  2. Every identifier REUSE.toml uses has its full text in LICENSES/<id>.txt.
  3. The load-bearing in-file headers are present and agree with the map:
     VLA_Runtime.bas carries 0BSD INSIDE its injectable region (the text
     that is copied into customers' workbooks, so the notice travels with
     the code); every phrasebook carries MPL-2.0 as its first line, since
     MPL is per-file and an organization copying the file must see it.

House style (tools/check_*.ps1): PowerShell 5.1, host-independent, a
hardcoded and reviewable baseline, never wired into VlaSelfTest. Exit 0
clean, exit 1 with every failure listed.

Usage:  powershell -File tools\check_spdx.ps1
#>

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

# ---- baseline: in-file headers that must exist, and what they must say ----
$requiredHeaders = @(
    @{ Path = 'src/VLA_Runtime.bas';                 Id = '0BSD';       Comment = "'" },
    @{ Path = 'scripts/prelude.vla';                 Id = 'Apache-2.0'; Comment = ';' },
    @{ Path = 'scripts/polyglotta/english.vla';      Id = 'MPL-2.0';    Comment = ';' },
    @{ Path = 'scripts/polyglotta/espanol.vla';      Id = 'MPL-2.0';    Comment = ';' },
    @{ Path = 'scripts/polyglotta/alien.vla';        Id = 'MPL-2.0';    Comment = ';' },
    @{ Path = 'scripts/polyglotta/dansk.vla';        Id = 'MPL-2.0';    Comment = ';' },
    @{ Path = 'scripts/polyglotta/deutsche.vla';     Id = 'MPL-2.0';    Comment = ';' },
    @{ Path = 'scripts/polyglotta/esperanto.vla';    Id = 'MPL-2.0';    Comment = ';' },
    @{ Path = 'scripts/polyglotta/francais.vla';     Id = 'MPL-2.0';    Comment = ';' },
    @{ Path = 'scripts/polyglotta/latin.vla';        Id = 'MPL-2.0';    Comment = ';' },
    @{ Path = 'scripts/polyglotta/pirate.vla';       Id = 'MPL-2.0';    Comment = ';' }
)

# Source extensions the map must cover. Binaries, build outputs, and the
# tracked dev workbook are deliberately outside the check.
$sourceExt = @('.bas', '.cls', '.frm', '.vla', '.ps1', '.md', '.toml', '.txt', '.iss', '.vba', '.lisp', '.pl')

$failures = New-Object System.Collections.Generic.List[string]

# ---- 1. parse REUSE.toml (the subset this project writes) ----
$reusePath = Join-Path $root 'REUSE.toml'
if (-not (Test-Path $reusePath)) { Write-Host 'FAIL: REUSE.toml missing'; exit 1 }

$annotations = @()
$current = $null
foreach ($line in Get-Content $reusePath) {
    $t = $line.Trim()
    if ($t -eq '[[annotations]]') {
        if ($current) { $annotations += $current }
        $current = @{ Paths = @(); Id = $null }
        continue
    }
    if (-not $current) { continue }
    if ($t -match '^path\s*=\s*\[(.*)\]$') {
        $current.Paths = @([regex]::Matches($Matches[1], '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
    } elseif ($t -match '^path\s*=\s*"([^"]+)"$') {
        $current.Paths = @($Matches[1])
    } elseif ($t -match '^SPDX-License-Identifier\s*=\s*"([^"]+)"$') {
        $current.Id = $Matches[1]
    }
}
if ($current) { $annotations += $current }
if ($annotations.Count -eq 0) { $failures.Add('REUSE.toml: no [[annotations]] tables found') }

function ConvertTo-GlobRegex([string]$glob) {
    $r = [regex]::Escape($glob)
    $r = $r -replace '\\\*\\\*/', '(.*/)?'   # **/  -> any directory depth
    $r = $r -replace '\\\*\\\*', '.*'         # **   -> anything
    $r = $r -replace '\\\*', '[^/]*'          # *    -> within one segment
    return '^' + $r + '$'
}

# ---- 2. every identifier has its text ----
$ids = $annotations | ForEach-Object { $_.Id } | Where-Object { $_ } | Sort-Object -Unique
foreach ($id in $ids) {
    $p = Join-Path $root ("LICENSES/" + $id + ".txt")
    if (-not (Test-Path $p)) { $failures.Add("LICENSES/$id.txt missing (REUSE.toml names $id)") }
}

# ---- 3. every tracked source file resolves to one licence ----
Push-Location $root
try { $tracked = git ls-files } finally { Pop-Location }
$resolved = @{}
foreach ($f in $tracked) {
    $ext = [System.IO.Path]::GetExtension($f).ToLowerInvariant()
    if ($sourceExt -notcontains $ext) { continue }
    if ($f -like 'LICENSES/*' -or $f -eq 'LICENSE') { continue }
    $match = $null
    foreach ($a in $annotations) {
        foreach ($g in $a.Paths) {
            if ($f -match (ConvertTo-GlobRegex $g)) { $match = $a.Id }   # later tables win
        }
    }
    if (-not $match) { $failures.Add("uncovered: $f matches no REUSE.toml annotation") }
    else { $resolved[$f] = $match }
}

# ---- 4. load-bearing in-file headers ----
foreach ($h in $requiredHeaders) {
    $p = Join-Path $root $h.Path
    if (-not (Test-Path $p)) { $failures.Add("missing file for required header: $($h.Path)"); continue }
    $head = Get-Content $p -TotalCount 12
    $want = $h.Comment + ' SPDX-License-Identifier: ' + $h.Id
    $has = $false
    foreach ($l in $head) { if ($l.Trim() -eq $want) { $has = $true } }
    if (-not $has) { $failures.Add("header: $($h.Path) lacks '$want' in its first 12 lines") }
    if ($resolved.ContainsKey($h.Path) -and $resolved[$h.Path] -ne $h.Id) {
        $failures.Add("disagreement: $($h.Path) header says $($h.Id), REUSE.toml resolves $($resolved[$h.Path])")
    }
}

# VLA_Runtime.bas: the header must sit ABOVE the inject boundary.
$rt = Join-Path $root 'src/VLA_Runtime.bas'
if (Test-Path $rt) {
    $lines = Get-Content $rt
    $hdr = -1; $bnd = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($hdr -lt 0 -and $lines[$i] -match 'SPDX-License-Identifier: 0BSD') { $hdr = $i }
        if ($bnd -lt 0 -and $lines[$i] -match 'EN_RUNTIME INJECT BOUNDARY') { $bnd = $i }
    }
    if ($hdr -ge 0 -and $bnd -ge 0 -and $hdr -gt $bnd) { $failures.Add('VLA_Runtime.bas: 0BSD header is BELOW the inject boundary; it would not travel with the injected text') }
}

# ---- report ----
$summary = ($resolved.Values | Group-Object | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', '
Write-Host "check_spdx: $($resolved.Count) source files resolved ($summary)"
if ($failures.Count -gt 0) {
    Write-Host "FAIL: $($failures.Count) problem(s)"
    $failures | ForEach-Object { Write-Host "  - $_" }
    exit 1
}
Write-Host 'OK: every source file has one licence, every licence has its text, every load-bearing header is in place'
exit 0
