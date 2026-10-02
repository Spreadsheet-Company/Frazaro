<#
check_prove_floors.ps1 - the core's passing phrasebook proofs, per file,
never go down.

WHY: PORT.6 ports English to the core (core/) slice by slice, and the
treaty's third oracle (conformance/README.md) is every test-success and
test-fail form of every source phrasebook under scripts/polyglotta/, run by
`<impl> prove <phrasebook.vla>`. Until the whole grammar is there, a file's
proofs pass in part; what must never happen is that a slice lands and fewer
of them pass than did yesterday. So this check holds, per file, the number
of proofs the built door passes, as a hardcoded, reviewable floor in the
house style (check_compile_prefix.ps1 held the compile golden's matched
prefix the same way), raised by hand as the matcher and the grammars land,
until each floor is its file's whole count and tools/prove.ps1 scores the
oracle as passed.

WHAT IT RUNS: `<impl> prove <file>` for each file in the baseline below, on
the built door at target/debug/frazaro.exe (or target/release), or on -Impl
when given. The last line of the output is read as the treaty defines it:
`PASS n/n` or `FAIL k/n`, k the proofs that passed, n the proofs the file
holds. An exit of 3 is "not attempted" and counts as 0 passed, which the
floors of a slice that attempts nothing allow. A tree with no door says
SKIPPED and exits 0, as check_core_imports.ps1 does, since CI always has it.

ALSO HELD: n itself. The treaty's count (its amendment of 2026-10-02) is the
forms the loader runs after a phrasebook's own generators have expanded:
read from the <name>_expanded.vla export beside the source when there is
one, whose source-hash stamp must match the source, and from the source's
own top-level forms otherwise. An implementation whose n differs has counted
something else, and fails here however many it passed.

-Control proves the check on two fake implementations written to a scratch
directory, against floors set to each file's whole count: one that answers
`PASS n/n` for every file (must pass), and one that answers `FAIL n-1/n`
wherever n is above 0 (must fail on exactly those files).

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; exit 0 clean, exit 1 with every file below its floor named. The
door's stderr is read under ErrorActionPreference Continue, since a native
command's stderr line is a terminating error under Stop.

Usage:  powershell -File tools\check_prove_floors.ps1
        powershell -File tools\check_prove_floors.ps1 -Impl C:\path\to\frazaro.exe
        powershell -File tools\check_prove_floors.ps1 -Control
#>
param(
    [string]$Impl = '',
    [switch]$Control
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# --- the floors: proofs passed per source phrasebook, never lowered --------
# 2026-10-02, PORT.6 step 0: 0 everywhere. The core attempts no proof yet
# (frazaro prove exits 3); each floor rises as the matcher (6d) and the
# grammars (6e) land, to its file's whole count (482, 30, 20 x 6).
# 2026-10-02, LX.15: alien.vla has no floor - a library of macros a program
# includes, not a phrasebook (no rule, no proof), which prove.ps1 inventories
# and never scores.
$floors = [ordered]@{
    'scripts/polyglotta/dansk.vla'     = 0
    'scripts/polyglotta/deutsche.vla'  = 0
    'scripts/polyglotta/english.vla'   = 0
    'scripts/polyglotta/espanol.vla'   = 0
    'scripts/polyglotta/esperanto.vla' = 0
    'scripts/polyglotta/francais.vla'  = 0
    'scripts/polyglotta/latin.vla'     = 0
    'scripts/polyglotta/pirate.vla'    = 0
}

# The treaty's proof form, at the top level of a line (prove.ps1 counts the
# same pattern the same way).
$proofForm = '^\(test-(success|fail)\b'

function Get-FormCount([string]$path) {
    return @(Select-String -Path $path -Pattern $proofForm).Count
}

# SHA-256 over a file's non-whitespace bytes, upper-case hex: the stamp an
# expanded phrasebook export carries (check_rule_coverage.ps1 reads it the
# same way).
function Get-NonWhitespaceSha256([string]$path) {
    $raw = [System.IO.File]::ReadAllBytes($path)
    $packed = New-Object byte[] $raw.Length
    $k = 0
    foreach ($b in $raw) {
        if ($b -ne 9 -and $b -ne 10 -and $b -ne 13 -and $b -ne 32) { $packed[$k] = $b; $k++ }
    }
    $out = New-Object byte[] $k
    [Array]::Copy($packed, $out, $k)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $digest = $sha.ComputeHash($out) } finally { $sha.Dispose() }
    return @{ Hex = (([BitConverter]::ToString($digest)) -replace '-', ''); Count = [long]$k }
}

# Oracle 3's n for one source phrasebook: from a fresh expanded export beside
# it, else from the source itself. A stale export is a failure to report.
function Get-ProofCount([string]$relFile) {
    $file = Join-Path $repoRoot $relFile
    $dir = Split-Path -Parent $file
    $base = [System.IO.Path]::GetFileNameWithoutExtension($file)
    $expanded = Join-Path $dir ($base + '_expanded.vla')
    if (-not (Test-Path $expanded)) { return @{ Count = (Get-FormCount $file); Stale = '' } }
    $stamp = (Get-Content -LiteralPath $expanded -TotalCount 2)[1]
    $m = [regex]::Match($stamp, 'source-hash:\s*sha256:([0-9A-Fa-f]{64}) over (\d+) non-whitespace bytes')
    $have = Get-NonWhitespaceSha256 $file
    if (-not $m.Success -or $m.Groups[1].Value.ToUpperInvariant() -ne $have.Hex -or [long]$m.Groups[2].Value -ne $have.Count) {
        return @{ Count = 0; Stale = "$($base)_expanded.vla is stale for $relFile - re-export it (Export Expanded Phrasebook)" }
    }
    return @{ Count = (Get-FormCount $expanded); Stale = '' }
}

# `<impl> prove <file>`: its stdout as text and its exit code, the door's
# stderr set aside (a refusal there is the PASS/FAIL line's business).
function Invoke-Prove([string]$impl, [string]$file) {
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $global:LASTEXITCODE = 0
        $lines = & $impl prove $file 2>$null
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $old
    }
    if ($null -eq $code) { $code = 0 }
    $text = if ($null -eq $lines) { '' } else { (@($lines) | ForEach-Object { [string]$_ }) -join "`n" }
    return @{ Stdout = $text; ExitCode = [int]$code }
}

# One file: what the implementation passed, and whether its n is the corpus's.
function Measure-File([string]$impl, [string]$relFile) {
    $pc = Get-ProofCount $relFile
    $r = @{ File = $relFile; Expected = $pc.Count; Passed = 0; Attempted = $true; Problem = '' }
    if ($pc.Stale -ne '') { $r.Problem = $pc.Stale; return $r }
    $run = Invoke-Prove $impl (Join-Path $repoRoot $relFile)
    if ($run.ExitCode -eq 3) { $r.Attempted = $false; return $r }
    $last = ($run.Stdout.TrimEnd("`n", "`r") -split "`n") | Select-Object -Last 1
    $m = [regex]::Match($last, '^(PASS|FAIL) (\d+)/(\d+)$')
    if (-not $m.Success) { $r.Problem = "the last line '$last' is not PASS n/n or FAIL k/n (exit $($run.ExitCode))"; return $r }
    $r.Passed = [int]$m.Groups[2].Value
    $n = [int]$m.Groups[3].Value
    if ($n -ne $pc.Count) { $r.Problem = "counted $n proof(s) where the corpus holds $($pc.Count)" }
    elseif ($m.Groups[1].Value -eq 'PASS' -and ($run.ExitCode -ne 0 -or $r.Passed -ne $n)) { $r.Problem = "PASS $($r.Passed)/$n with exit $($run.ExitCode)" }
    return $r
}

# Every file against a floors table. Returns the report lines and the
# failures together, and writes nothing to the pipeline itself.
function Measure-All([string]$impl, $table) {
    $lines = New-Object System.Collections.Generic.List[string]
    $failed = New-Object System.Collections.Generic.List[string]
    foreach ($relFile in $table.Keys) {
        $m = Measure-File $impl $relFile
        $floor = $table[$relFile]
        $label = $relFile -replace '^scripts/', ''
        if ($m.Problem -ne '') {
            $lines.Add(("  {0,-30} {1}  FAIL" -f $label, $m.Problem))
            $failed.Add("$label - $($m.Problem)")
        } elseif (-not $m.Attempted) {
            $mark = if ($floor -gt 0) { '  FAIL' } else { '' }
            $lines.Add(("  {0,-30} not attempted (exit 3)  of {1}, floor {2}{3}" -f $label, $m.Expected, $floor, $mark))
            if ($floor -gt 0) { $failed.Add("$label - not attempted, but its floor is $floor") }
        } elseif ($m.Passed -lt $floor) {
            $lines.Add(("  {0,-30} passed {1,4} of {2,4}, BELOW THE FLOOR of {3}  FAIL" -f $label, $m.Passed, $m.Expected, $floor))
            $failed.Add("$label - passed $($m.Passed) of $($m.Expected), below its floor of $floor")
        } else {
            $note = if ($m.Passed -gt $floor) { ' - raise the floor in this file' } else { '' }
            $lines.Add(("  {0,-30} passed {1,4} of {2,4} (floor {3}){4}" -f $label, $m.Passed, $m.Expected, $floor, $note))
        }
    }
    return @{ Lines = $lines; Failed = $failed }
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_floors_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        # A fake that passes everything and a mutant one short, each counting
        # n as the treaty counts it (the expanded export where there is one).
        $count = @"
`$f = `$a[1]
`$exp = Join-Path (Split-Path -Parent `$f) ([System.IO.Path]::GetFileNameWithoutExtension(`$f) + '_expanded.vla')
`$counted = if (Test-Path `$exp) { `$exp } else { `$f }
`$n = @(Select-String -Path `$counted -Pattern '$proofForm').Count
"@
        $fake = Join-Path $tmp 'fake.ps1'
        [System.IO.File]::WriteAllText($fake, "param([Parameter(ValueFromRemainingArguments=`$true)][string[]]`$a)`n$count`nWrite-Output `"PASS `$n/`$n`"`nexit 0`n")
        $mut = Join-Path $tmp 'mutant.ps1'
        [System.IO.File]::WriteAllText($mut, "param([Parameter(ValueFromRemainingArguments=`$true)][string[]]`$a)`n$count`nif (`$n -eq 0) { Write-Output 'PASS 0/0'; exit 0 }`nWrite-Output `"FAIL `$(`$n - 1)/`$n`"`nexit 1`n")
        # The control's floors: each file's whole count.
        $whole = [ordered]@{}
        $above = 0
        foreach ($relFile in $floors.Keys) {
            $pc = Get-ProofCount $relFile
            if ($pc.Stale -ne '') { Write-Output "FAIL: $($pc.Stale)"; exit 1 }
            $whole[$relFile] = $pc.Count
            if ($pc.Count -gt 0) { $above++ }
        }
        Write-Output 'control: the fake (PASS n/n everywhere) against floors of n'
        $r1 = Measure-All $fake $whole
        $r1.Lines | ForEach-Object { Write-Output $_ }
        Write-Output 'control: the mutant (FAIL n-1/n wherever n is above 0)'
        $r2 = Measure-All $mut $whole
        $r2.Lines | ForEach-Object { Write-Output $_ }
        if ($r1.Failed.Count -eq 0 -and $r2.Failed.Count -eq $above -and $above -gt 0) {
            Write-Output "OK: the fake passed every floor; the mutant failed all $above file(s) with proofs"
            exit 0
        }
        Write-Output "FAIL: fake failures $($r1.Failed.Count) (want 0); mutant failures $($r2.Failed.Count) (want $above)"
        exit 1
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}

if ($Impl -eq '') {
    foreach ($candidate in @('target/debug/frazaro.exe', 'target/release/frazaro.exe', 'target/debug/frazaro', 'target/release/frazaro')) {
        $p = Join-Path $repoRoot $candidate
        if (Test-Path $p) { $Impl = $p; break }
    }
}
if ($Impl -eq '' -or -not (Test-Path $Impl)) {
    Write-Output 'SKIPPED: no built door at target/debug/frazaro.exe or target/release/frazaro.exe (cargo build --workspace writes it; CI always has it)'
    exit 0
}

Write-Output '=== PASSING PHRASEBOOK PROOFS PER FILE NEVER GO DOWN (PORT.6, oracle 3) ==='
$res = Measure-All $Impl $floors
$res.Lines | ForEach-Object { Write-Output $_ }
Write-Output ''
if ($res.Failed.Count -eq 0) {
    Write-Output '=== CHECK: clean - every file passes at least its floor of proofs ==='
    exit 0
}
Write-Output "=== CHECK: $($res.Failed.Count) file(s) below their floor ==="
$res.Failed | ForEach-Object { Write-Output "  $_" }
exit 1
