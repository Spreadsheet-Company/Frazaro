<#
check_crate_data.ps1 - the files a crate carries as copies are their sources,
byte for byte.

WHY: `cargo install frazaro` builds the door from its crate directory and
nothing beside it, so the prelude and the English phrasebook the door carries
(cli/data/, embedded at build time, used when --prelude or --phrasebook is
not given) are COPIES of scripts/prelude.vla and
scripts/polyglotta/english.vla, the corpus files the add-in, the web page and
the treaty read. Two copies of one file with nothing mechanical to hold them
together is the drift check_data_exports.ps1 was written for: a prelude edit
lands in scripts/, every test passes, and the published door translates with
yesterday's prelude. This check compares each pair byte for byte and names
the one copy command that mends a drift. The copies are never edited by hand.

ALSO HELD, on a built door (-Impl, or target/debug/frazaro beside the tree,
as the other -Impl checks find it): the door with no flags translates the
corpus program exactly as it does with the corpus files named, so that the
embedding itself, not only the bytes on disk, is what the flags would give.
SKIPPED where no door was built, exit 0, since CI always has one.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; a hardcoded, reviewable baseline (the pair count). Exit 0 clean,
exit 1 with every drift named.

Usage:  powershell -File tools\check_crate_data.ps1
        powershell -File tools\check_crate_data.ps1 -Impl target\debug\frazaro.exe
        powershell -File tools\check_crate_data.ps1 -Root <dir>   # testability
#>
param(
    [string]$Impl = '',
    [string]$Root = ''
)

$ErrorActionPreference = 'Stop'
$root = if ($Root -ne '') { (Resolve-Path $Root).Path } else { Split-Path -Parent $PSScriptRoot }
$root = $root.TrimEnd('\', '/')

# ---- the baseline: the pairs, copy then source; raise the count with a pair ----
$pairs = @(
    @{ Copy = 'cli/data/prelude.vla'; Source = 'scripts/prelude.vla' },
    @{ Copy = 'cli/data/english.vla'; Source = 'scripts/polyglotta/english.vla' }
)
$expectedPairs = 2

$failures = New-Object System.Collections.Generic.List[string]
if ($pairs.Count -ne $expectedPairs) { $failures.Add("the table holds $($pairs.Count) pair(s), not the $expectedPairs pinned here") }

function Get-FirstDifference([byte[]]$a, [byte[]]$b) {
    $n = [Math]::Min($a.Length, $b.Length)
    for ($i = 0; $i -lt $n; $i++) { if ($a[$i] -ne $b[$i]) { return $i } }
    if ($a.Length -ne $b.Length) { return $n }
    return -1
}

# ---- 1. each copy is its source, byte for byte ----
$sha = [System.Security.Cryptography.SHA256]::Create()
foreach ($p in $pairs) {
    $copyPath = Join-Path $root $p.Copy
    $sourcePath = Join-Path $root $p.Source
    if (-not (Test-Path $sourcePath)) { $failures.Add("$($p.Source) is missing"); continue }
    if (-not (Test-Path $copyPath)) { $failures.Add("$($p.Copy) is missing - run: Copy-Item $($p.Source) $($p.Copy)"); continue }
    $a = [System.IO.File]::ReadAllBytes($copyPath)
    $b = [System.IO.File]::ReadAllBytes($sourcePath)
    if ([Convert]::ToBase64String($sha.ComputeHash($a)) -ne [Convert]::ToBase64String($sha.ComputeHash($b))) {
        $at = Get-FirstDifference $a $b
        $failures.Add("$($p.Copy) differs from $($p.Source) at byte $at ($($a.Length) and $($b.Length) bytes) - run: Copy-Item $($p.Source) $($p.Copy)")
    }
}

# ---- 2. a built door with no flags is the door with the corpus files named ----
function Invoke-Door([string]$exe, [string[]]$arguments) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $exe
    $psi.Arguments = (($arguments | ForEach-Object { '"' + $_ + '"' }) -join ' ')
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $proc = [System.Diagnostics.Process]::Start($psi)
    $out = New-Object System.IO.MemoryStream
    $proc.StandardOutput.BaseStream.CopyTo($out)
    $err = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()
    return @{ Code = $proc.ExitCode; Bytes = $out.ToArray(); Err = $err.Trim() }
}

if ($Impl -eq '') {
    foreach ($p in @('target/debug/frazaro.exe', 'target/release/frazaro.exe', 'target/debug/frazaro', 'target/release/frazaro')) {
        $candidate = Join-Path $root $p
        if (Test-Path $candidate) { $Impl = $candidate; break }
    }
}
$doorNote = 'no built door, so the embedding itself was not run (cargo build --workspace writes one; CI always has it)'
if ($Impl -ne '' -and (Test-Path $Impl)) {
    $program = Join-Path $root 'scripts/instructions.txt'
    $prelude = Join-Path $root 'scripts/prelude.vla'
    $english = Join-Path $root 'scripts/polyglotta/english.vla'
    $bare = Invoke-Door $Impl @('translate-vla', $program)
    $named = Invoke-Door $Impl @('translate-vla', $program, '--prelude', $prelude, '--phrasebook', $english)
    if ($bare.Code -ne 0 -or $named.Code -ne 0) {
        $failures.Add("the door exited $($bare.Code) with no flags ($($bare.Err)) and $($named.Code) with the corpus files named ($($named.Err)) on scripts/instructions.txt")
    } elseif ((Get-FirstDifference $bare.Bytes $named.Bytes) -ge 0) {
        $failures.Add("the door translates scripts/instructions.txt differently with no flags and with --prelude scripts/prelude.vla --phrasebook scripts/polyglotta/english.vla: its built-in pair is not the corpus pair (copy, then cargo build --workspace)")
    } else {
        $doorNote = "the built door translates the corpus program the same with no flags and with the corpus files named ($($bare.Bytes.Length) bytes)"
    }
}

if ($failures.Count -gt 0) {
    Write-Host "FAIL: $($failures.Count) problem(s)"
    $failures | ForEach-Object { Write-Host "  - $_" }
    exit 1
}
Write-Host "OK: $($pairs.Count) copies under cli/data/ are their sources byte for byte; $doorNote"
exit 0
