<#
check_crate_package.ps1 - the crates can be packaged, and the language's
package can test itself: what cargo package would refuse, read without cargo.

WHY: crates.io receives a crate's directory and nothing beside it. `cargo
package` builds the tarball it has made, to prove the crate stands alone,
and the first time the core's was tried (2026-10-05) it did not: four
`include_str!` paths reached from core/src/ up into scripts/ for the
exported data tables, which the workspace build had resolved on every
commit and the tarball could not. The tables moved into core/data/, and
this check holds the shape that keeps the crates packageable:

  1. every `include_str!`/`include_bytes!` outside a module's test region
     (the `#[cfg(test)] mod tests` at a file's end, which cargo package does
     not compile) names a file that exists INSIDE its own crate's directory;
  2. each member's manifest names what crates.io shows a visitor -
     description, readme (a file beside the manifest), license and
     repository - its own or inherited from [workspace.package];
  3. the members are the ones the root Cargo.toml lists, and that many;
  4. (VLA.1, 2026-10-10) in vla-lang, the tests' includes too: the test
     region of every file under src/ and every file under tests/ names a
     file inside the crate, so that `cargo test` runs in the published
     package. Until VLA.1 twelve test includes reached scripts/, and the
     package's suite did not compile; the nine files now travel as copies
     under vla-lang/testdata/, which check_crate_data.ps1 holds to their
     sources. The core's and the door's tests may still read scripts/ on
     purpose: their packages are not asked to test themselves.

The ratchets job runs on Windows without a Rust toolchain, so this reads the
sources; `cargo package --workspace` is the rehearsal before a publish
(docs/DEPLOY.md, "Publishing the crates"), and the core CI job's build, and
its run of vla-lang's tests inside the unpacked package, are the other half.

-Control proves the check on a scratch copy of the three crates: the clean
copy passes, and each of four mutants fails - a vla-lang test include
reaching scripts/ again, an integration test under vla-lang/tests/ reaching
out of the crate, a core build-time include leaving its crate, and a watched
test include respelled out of the scan's pattern (`include_str! (`, which
Rust accepts), so that the floor is what catches a scan that stopped seeing.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, a hardcoded and
reviewable baseline. Pinned: the crate-bound includes found never go below
their floors (a scan whose pattern stopped matching would otherwise pass by
finding nothing), and the member count. Exit 0 clean, exit 1 with every
problem named.

Usage:  powershell -File tools\check_crate_package.ps1
        powershell -File tools\check_crate_package.ps1 -Control
        powershell -File tools\check_crate_package.ps1 -Root <dir>   # testability
#>
param(
    [string]$Root = '',
    [switch]$Control
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$sep = [System.IO.Path]::DirectorySeparatorChar

# ---- the baseline: raise when a member or an include is added ----
$expectedMembers  = 3   # vla-lang, core, cli (PORT.4; PORT.12 added the language, 2026-10-08)
$includeFloor     = 7   # vla-lang/src: headtable.vla, messages.vla (PORT.12); core/src: messages.vla, words.vla, names.vla (PORT.5, PORT.6); cli/src: prelude.vla, english.vla (2026-10-05, the door's built-in pair)
$selfTesting      = @('vla-lang')   # VLA.1: the crates whose own package runs their tests
$testIncludeFloor = 35  # VLA.1, 2026-10-10: vla-lang's test includes, each inside the crate - in src/, expand.rs 2, reader.rs 2, refers.rs 2, machine/mod.rs 6; in tests/, laws.rs 2, strangers.rs 8; VLA.3: expand_golden.rs 3; VLA.2: walks.rs 1; KERNEL.47: formats_golden.rs 2; VLA.4: formula_syntax.rs 3; KERNEL.50: appearance.rs 2, appearance_golden.rs 2
$requiredKeys     = @('description', 'readme', 'license', 'repository')
$includePattern   = 'include_(?:str|bytes)!\(\s*"([^"]+)"\s*\)'

function Get-TestStart([string[]]$lines) {
    for ($i = 0; $i -lt $lines.Count - 1; $i++) {
        if ($lines[$i] -match '^#\[cfg\(test\)\]\s*$' -and $lines[$i + 1] -match '^mod\s+\w+') { return $i }
    }
    return $lines.Count
}

# Every include on lines [from, to) of a file, each held inside the crate.
function Test-Includes($failures, [string]$root, [string]$crateDir, [string]$m, $file, [string[]]$lines, [int]$from, [int]$to, [string]$what) {
    $count = 0
    $relFile = $file.FullName.Substring($root.Length + 1)
    for ($i = $from; $i -lt $to; $i++) {
        foreach ($mt in [regex]::Matches($lines[$i], $includePattern)) {
            $count++
            $rel = $mt.Groups[1].Value
            $target = [System.IO.Path]::GetFullPath((Join-Path $file.DirectoryName $rel))
            if (-not $target.StartsWith($crateDir + $sep)) {
                $failures.Add("${relFile}:$($i + 1): $what include of `"$rel`" leaves the crate ($m/), which its package cannot ship")
            } elseif (-not (Test-Path $target)) {
                $failures.Add("${relFile}:$($i + 1): $what include of `"$rel`" names no file")
            }
        }
    }
    return $count
}

function Get-Failures([string]$root) {
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
    $testFound = 0
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

        # ---- 1. every build-time include stays inside the crate; 4. so do a self-testing crate's test includes ----
        $selfTest = $selfTesting -contains $m
        foreach ($f in Get-ChildItem -Path (Join-Path $crateDir 'src') -Filter '*.rs' -Recurse) {
            $lines = @(Get-Content $f.FullName)
            $testStart = Get-TestStart $lines
            $found += Test-Includes $failures $root $crateDir $m $f $lines 0 $testStart 'a build-time'
            if ($selfTest) {
                $testFound += Test-Includes $failures $root $crateDir $m $f $lines $testStart $lines.Count 'a test'
            }
        }
        $testsDir = Join-Path $crateDir 'tests'
        if ($selfTest -and (Test-Path $testsDir)) {
            foreach ($f in Get-ChildItem -Path $testsDir -Filter '*.rs' -Recurse) {
                $lines = @(Get-Content $f.FullName)
                $testFound += Test-Includes $failures $root $crateDir $m $f $lines 0 $lines.Count 'a test'
            }
        }
    }
    if ($found -lt $includeFloor) {
        $failures.Add("$found crate-bound build-time include(s) found, below the floor of ${includeFloor}: the scan's pattern has stopped matching, or an include was removed (then lower the floor deliberately, with the reason)")
    }
    if ($testFound -lt $testIncludeFloor) {
        $failures.Add("$testFound test include(s) found in $($selfTesting -join ', '), below the floor of ${testIncludeFloor}: the scan's pattern has stopped matching, or a test include was removed (then lower the floor deliberately, with the reason)")
    }
    return @{ Failures = $failures; Found = $found; TestFound = $testFound; Members = $members.Count }
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_package_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        # A scratch copy of what the check reads: the root manifest and, for
        # each crate, its manifest, page, sources, tables, fixtures and tests.
        Copy-Item -Path (Join-Path $repoRoot 'Cargo.toml') -Destination (Join-Path $tmp 'Cargo.toml')
        foreach ($m in @('vla-lang', 'core', 'cli')) {
            New-Item -ItemType Directory -Path (Join-Path $tmp $m) -Force | Out-Null
            foreach ($part in @('Cargo.toml', 'README.md', 'src', 'data', 'testdata', 'tests')) {
                $src = Join-Path (Join-Path $repoRoot $m) $part
                if (Test-Path $src) { Copy-Item -Path $src -Destination (Join-Path (Join-Path $tmp $m) $part) -Recurse -Force }
            }
        }
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $clean = Get-Failures $tmp
        $verdicts = New-Object System.Collections.Generic.List[string]
        $ok = ($clean.Failures.Count -eq 0)
        $verdicts.Add(("the clean copy: {0} problem(s), {1} build-time and {2} test include(s)" -f $clean.Failures.Count, $clean.Found, $clean.TestFound))
        $mutants = @(
            @{ Name = 'a vla-lang test include reaching scripts/ again'; File = 'vla-lang/src/reader.rs'; Find = 'include_str!("../testdata/prelude.vla")'; Replace = 'include_str!("../../scripts/prelude.vla")' },
            @{ Name = 'an integration test reaching out of the crate'; File = 'vla-lang/tests/control_mutant.rs'; Find = ''; Replace = "const OUTSIDE: &str = include_str!(`"../../scripts/prelude.vla`");`n" },
            @{ Name = 'a core build-time include leaving its crate'; File = 'core/src/messages.rs'; Find = 'include_str!("../data/messages.vla")'; Replace = 'include_str!("../../scripts/messages.vla")' },
            @{ Name = 'a watched test include respelled out of the pattern'; File = 'vla-lang/src/refers.rs'; Find = 'include_str!("../testdata/refers_golden.txt")'; Replace = 'include_str! ("../testdata/refers_golden.txt")' }
        )
        foreach ($mu in $mutants) {
            $path = Join-Path $tmp $mu.File
            $isNew = ($mu.Find -eq '')
            $before = if ($isNew) { $null } else { [System.IO.File]::ReadAllText($path) }
            if ($isNew) {
                New-Item -ItemType Directory -Path (Split-Path -Parent $path) -Force | Out-Null
                [System.IO.File]::WriteAllText($path, $mu.Replace, $utf8)
            } else {
                $hits = [regex]::Matches($before, [regex]::Escape($mu.Find)).Count
                if ($hits -ne 1) { $ok = $false; $verdicts.Add(("{0}: its anchor occurs {1} time(s) in {2}, not once, so the mutant was not applied" -f $mu.Name, $hits, $mu.File)); continue }
                [System.IO.File]::WriteAllText($path, $before.Replace($mu.Find, $mu.Replace), $utf8)
            }
            $r = Get-Failures $tmp
            if ($isNew) { Remove-Item -Force $path } else { [System.IO.File]::WriteAllText($path, $before, $utf8) }
            if ($r.Failures.Count -ge 1) { $verdicts.Add(("{0}: fails, as it should ({1}: {2})" -f $mu.Name, $r.Failures.Count, $r.Failures[0])) }
            else { $ok = $false; $verdicts.Add(("{0}: PASSED but should fail" -f $mu.Name)) }
        }
        foreach ($v in $verdicts) { Write-Output ('control: ' + $v) }
        if ($ok) { Write-Output 'OK: control: the clean copy passes and each of the four mutants fails'; exit 0 }
        foreach ($p in $clean.Failures) { Write-Output ('  clean: ' + $p) }
        Write-Output 'FAIL: the control did not behave as the header says'
        exit 1
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}

$root = if ($Root -ne '') { (Resolve-Path $Root).Path } else { $repoRoot }
$root = $root.TrimEnd('\', '/')
$result = Get-Failures $root
if ($result.Failures.Count -gt 0) {
    Write-Host "FAIL: $($result.Failures.Count) problem(s)"
    $result.Failures | ForEach-Object { Write-Host "  - $_" }
    exit 1
}
Write-Host "OK: $($result.Members) crates packageable - $($result.Found) build-time include(s), each inside its crate; $($result.TestFound) test include(s) in $($selfTesting -join ', '), each inside the crate; every manifest names description, readme, license and repository"
exit 0
