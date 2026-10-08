<#
check_crate_package.ps1 - the crates can be packaged: what cargo package
would refuse, read without cargo.

WHY: crates.io receives a crate's directory and nothing beside it. `cargo
package` builds the tarball it has made, to prove the crate stands alone,
and the first time the core's was tried (2026-10-05) it did not: four
`include_str!` paths reached from core/src/ up into scripts/ for the
exported data tables, which the workspace build had resolved on every
commit and the tarball could not. The tables moved into core/data/, and
this check holds the shape that keeps the crates packageable:

  1. every `include_str!`/`include_bytes!` outside a module's test region
     (the `#[cfg(test)] mod tests` at a file's end, which cargo package does
     not compile, and which reaches the corpus in scripts/ on purpose) names
     a file that exists INSIDE its own crate's directory;
  2. each member's manifest names what crates.io shows a visitor -
     description, readme (a file beside the manifest), license and
     repository - its own or inherited from [workspace.package];
  3. the members are the ones the root Cargo.toml lists, and that many.

The ratchets job runs on Windows without a Rust toolchain, so this reads the
sources; `cargo package --workspace` is the rehearsal before a publish
(docs/DEPLOY.md, "Publishing the crates"), and the core CI job's build is
the other half.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, a hardcoded and
reviewable baseline. Pinned: the crate-bound includes found never go below
their floor (a scan whose pattern stopped matching would otherwise pass by
finding nothing), and the member count. Exit 0 clean, exit 1 with every
problem named.

Usage:  powershell -File tools\check_crate_package.ps1
        powershell -File tools\check_crate_package.ps1 -Root <dir>   # testability
#>
param([string]$Root = '')

$ErrorActionPreference = 'Stop'
$root = if ($Root -ne '') { (Resolve-Path $Root).Path } else { Split-Path -Parent $PSScriptRoot }
$root = $root.TrimEnd('\', '/')
$sep = [System.IO.Path]::DirectorySeparatorChar

# ---- the baseline: raise when a member or a build-time include is added ----
$expectedMembers = 3   # vla-lang, core, cli (PORT.4; PORT.12 added the language, 2026-10-08)
$includeFloor    = 7   # vla-lang/src: headtable.vla, messages.vla (PORT.12); core/src: messages.vla, words.vla, names.vla (PORT.5, PORT.6); cli/src: prelude.vla, english.vla (2026-10-05, the door's built-in pair)
$requiredKeys    = @('description', 'readme', 'license', 'repository')

$failures = New-Object System.Collections.Generic.List[string]

# ---- 3. the members and the inheritable keys, from the root manifest ----
$members = @()
$workspaceKeys = @{}
$table = ''
foreach ($line in @(Get-Content (Join-Path $root 'Cargo.toml'))) {
    $t = $line.Trim()
    if ($t -match '^\[(.+)\]$') { $table = $Matches[1]; continue }
    if ($table -eq 'workspace' -and $t -match '^members\s*=\s*\[(.*)\]') {
        $members = @([regex]::Matches($Matches[1], '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
    }
    if ($table -eq 'workspace.package' -and $t -match '^([a-z-]+)\s*=') { $workspaceKeys[$Matches[1]] = $true }
}
if ($members.Count -ne $expectedMembers) {
    $failures.Add("Cargo.toml lists $($members.Count) workspace member(s), not the $expectedMembers pinned here")
}

$found = 0
foreach ($m in $members) {
    $crateDir = [System.IO.Path]::GetFullPath((Join-Path $root $m)).TrimEnd('\', '/')
    $manifest = Join-Path $crateDir 'Cargo.toml'
    if (-not (Test-Path $manifest)) { $failures.Add("$m/Cargo.toml is missing"); continue }

    # ---- 2. the manifest's [package] keys ----
    $keys = @{}
    $table = ''
    foreach ($line in @(Get-Content $manifest)) {
        $t = $line.Trim()
        if ($t -match '^\[(.+)\]$') { $table = $Matches[1]; continue }
        if ($table -ne 'package') { continue }
        if ($t -match '^([a-z-]+)\.workspace\s*=\s*true') {
            if ($workspaceKeys.ContainsKey($Matches[1])) { $keys[$Matches[1]] = '(workspace)' }
            else { $failures.Add("$m/Cargo.toml inherits '$($Matches[1])', which [workspace.package] does not define") }
        } elseif ($t -match '^([a-z-]+)\s*=\s*"([^"]*)"') {
            $keys[$Matches[1]] = $Matches[2]
        }
    }
    foreach ($k in $requiredKeys) {
        if (-not $keys.ContainsKey($k)) { $failures.Add("$m/Cargo.toml: [package] has no '$k', which crates.io shows for every crate") }
    }
    if ($keys.ContainsKey('readme') -and $keys['readme'] -ne '(workspace)' -and -not (Test-Path (Join-Path $crateDir $keys['readme']))) {
        $failures.Add("$m/Cargo.toml: readme = `"$($keys['readme'])`" names no file beside the manifest")
    }

    # ---- 1. every build-time include stays inside the crate ----
    foreach ($f in Get-ChildItem -Path (Join-Path $crateDir 'src') -Filter '*.rs' -Recurse) {
        $lines = @(Get-Content $f.FullName)
        $relFile = $f.FullName.Substring($root.Length + 1)
        $testStart = $lines.Count
        for ($i = 0; $i -lt $lines.Count - 1; $i++) {
            if ($lines[$i] -match '^#\[cfg\(test\)\]\s*$' -and $lines[$i + 1] -match '^mod\s+\w+') { $testStart = $i; break }
        }
        for ($i = 0; $i -lt $testStart; $i++) {
            foreach ($mt in [regex]::Matches($lines[$i], 'include_(?:str|bytes)!\(\s*"([^"]+)"\s*\)')) {
                $found++
                $rel = $mt.Groups[1].Value
                $target = [System.IO.Path]::GetFullPath((Join-Path $f.DirectoryName $rel))
                if (-not $target.StartsWith($crateDir + $sep)) {
                    $failures.Add("${relFile}:$($i + 1): include of `"$rel`" leaves the crate ($m/), which cargo package cannot ship")
                } elseif (-not (Test-Path $target)) {
                    $failures.Add("${relFile}:$($i + 1): include of `"$rel`" names no file")
                }
            }
        }
    }
}
if ($found -lt $includeFloor) {
    $failures.Add("$found crate-bound build-time include(s) found, below the floor of ${includeFloor}: the scan's pattern has stopped matching, or an include was removed (then lower the floor deliberately, with the reason)")
}

if ($failures.Count -gt 0) {
    Write-Host "FAIL: $($failures.Count) problem(s)"
    $failures | ForEach-Object { Write-Host "  - $_" }
    exit 1
}
Write-Host "OK: $($members.Count) crates packageable - $found build-time include(s), each inside its crate; every manifest names description, readme, license and repository"
exit 0
