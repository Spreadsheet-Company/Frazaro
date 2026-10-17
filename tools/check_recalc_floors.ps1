<#
check_recalc_floors.ps1 - KERNEL.20's oracle: the machine's recalculation
of four fixtures measured on the owner's machine beside Excel's of the same
files, held to floors that never go down; and, on every push where the
runner is built, the fixtures' values held.

WHY: KERNEL.20 (docs/BETA_ROADMAP.md: "Oracle: check_recalc_floors.ps1, the
timings as floors that never go down, in the house's measure-first form";
web/CALLOSUM.md section 12.6 calls the fastest spreadsheet a prediction
until it is measured). Its sixth slice built the measurement
(docs/BETA_REARVIEW.md, its scoping): core/examples/recalc.rs builds each
program of scripts/recalc/bench/ as frazaro build builds it and times the
machine computing it, under Excel's library, after holding every value;
tools/bench_recalc.ps1 runs it on the owner's machine and
tools/bench_recalc_excel.ps1 times Excel on the same files. A timing is a
measurement on the owner's machine, so the floors are the owner's run of
the two, pasted below.

WHAT IT HOLDS, from the baseline pasted below:
  - a date and two environment lines, each saying, as its instrument
    writes it, that it ran opted out of Windows' power throttling (an
    invisible or background process is slowed by about 40% a few seconds
    in), the native half on the performance cores and at high priority, and
    Excel's half by Excel and not the instrument's stand-in;
  - four native rows, one a fixture, each measured in at least three
    rounds, its clock within 10% across them and from the run's start, its
    per-second numbers agreeing with its times within half a percent;
  - eight Excel rows, one a fixture and a mode (all threads, as Excel
    ships, and one): each fixture's file the same bytes in both halves, by
    its SHA-256, and Excel's sums of the program's range, as built and as
    edited, equal to the machine's in both modes, so the two computed the
    same values;
  - three flag rows, Excel's reading of the build golden's fixture written
    with fullCalcOnLoad, with Excel 365's calcId, and with neither, one of
    the three verdicts each;
  - the fixtures' programs, each one's SHA-256 over LF bytes equal to the
    one measured, so a program edited without a new measurement fails by
    name;
  - the floors, a dated history, each fixture's whole frames in formulas a
    second and its edits a second: each entry's floor at least the entry
    before it unless the entry gives its reason; each native row's numbers
    at least the latest floors. The floor is the run's number less a margin
    of 15%: measured over three rounds, a fixture's medians moved up to 11%
    between processes (chain's whole frame, 3.87 to 4.31 ms over three runs,
    2026-10-10, the slice's build), and Alonzo's 10% would not hold it;
  - the history appended to and never pasted over: every entry the last
    commit's copy of this file holds is still here, in its order, at the
    head (SKIPPED where git, the commit or the file in it is not there to
    ask), since a paste over an entry would erase it and could lower a
    floor with no reason written (Alonzo's frame floors lost an entry so);
  - the bar: Excel's numbers printed beside the machine's, with the ratio;
    a measurement, not gated, since Excel's numbers move with its version.
An empty baseline fails: a floor with no number is no floor.

WITH A BUILT RUNNER (target/release/examples/recalc, or -Runner): its check
mode, every value held (each fixture's first frame to the values its build
folded, its edited state to Calc, every frame of Life to a Life written by
hand), and each fixture's sums equal to the baseline's, so the values
measured are the values computed today. A file's bytes may move with the
stamp (a phrasebook or the prelude edited) while its values do not; that is
printed and not failed. SKIPPED without a runner: the ratchets job builds
none.

-Control proves the judgment: a clean baseline passes, and so does one
with a floor lowered for a written reason; an empty baseline, a native row
missing, an Excel row missing, a file not the same bytes in both halves,
sums that differ, a clock that moved, too few rounds, per-second numbers
that disagree with the times, a native half not opted out, Excel's half
from the stand-in, a floor lowered with no reason, a row under its floor, a
program edited, a flag row with no verdict and a fixture Excel could not
open must each fail; the history,
read from text as the commit's copy is: an entry added after the committed
ones passes, and so does a history unchanged, while an entry pasted over a
committed one and a committed entry edited in place must each fail; then
fake runners: a clean one passes, and one whose values differ, one whose
sums differ and one that leaves out a fixture must each fail.

House style (tools/check_*.ps1): PowerShell 5.1 and pwsh alike, host-free,
no network; a hardcoded, reviewable baseline; exit 0 clean, exit 1 with
every problem named.

Usage:  powershell -File tools\check_recalc_floors.ps1
        powershell -File tools\check_recalc_floors.ps1 -Runner target\release\examples\recalc.exe
        powershell -File tools\check_recalc_floors.ps1 -Control
#>
param(
    [string]$Root = '',
    [string]$Runner = '',
    [switch]$Control
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if ($Root -eq '') { $Root = $repoRoot }
$onWindows = [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT

# --- the baseline, hand-maintained: the owner's run of the two instruments ---
# Paste the native half's block (tools/bench_recalc.ps1) over $baselineDate,
# $baselineNative and its four rows, Excel's half's
# (tools/bench_recalc_excel.ps1) over $baselineExcel and its eleven rows, and
# the history the native half prints over $floorHistory, whole: it is this
# file's history with the run's entry added last. A later run slower than a
# floor is a finding to record in docs/BETA_REARVIEW.md, never a number to
# paste over a faster one.
$baselineDate = '2026-10-17'
$baselineNative = 'Intel(R) Core(TM) Ultra 7 265KF, 20 logical processors | Microsoft Windows 11 Home 10.0.26200.0 | rustc 1.99.0 (b940084d7 2026-09-28) | frazaro-core 0.9.4 | power throttling opted out: the runner yes | high priority: the runner yes | on the performance cores, 8 of 20 logical processors (efficiency class 1): the runner yes'
$baselineExcel = 'Excel''s half, 2026-10-17: Intel(R) Core(TM) Ultra 7 265KF | Microsoft Windows 11 Home 10.0.26200.0 | Excel 16.0 build 20430 | all threads 20; one thread on the performance cores, mask 0xC03C3 | power throttling opted out: yes'
$baseline = @(
    # the native rows:
    @{ Half = 'native'; Fixture = 'column'; Formulas = 60000; Sha256 = '97C22A696D71AA680D997C3778FDA6960D4945A972FFDE5A6F8D729F8CFF4304'; Rounds = 3; Load = 208.5; First = 7.39; Whole = 0.468; Edit = 1.456; Calc = 214.8; WholePerSecond = 128095645; EditsPerSecond = 686.8; SumBuilt = 1100080000; SumEdited = 1425100000; ClockMs = 0.172; ClockSpread = 6.0; ClockDrift = -3.4 }
    @{ Half = 'native'; Fixture = 'chain'; Formulas = 39999; Sha256 = 'DC691ACEB3B881FB705CC6CA0CDB48E83C209495535EF67CEC5AF029FB9D77E4'; Rounds = 3; Load = 123.3; First = 4.73; Whole = 0.547; Edit = 1.331; Calc = 116.1; WholePerSecond = 73177827; EditsPerSecond = 751.3; SumBuilt = 1200039974; SumEdited = 1200059975; ClockMs = 0.170; ClockSpread = 6.3; ClockDrift = -4.9 }
    @{ Half = 'native'; Fixture = 'sums'; Formulas = 28014; Sha256 = '9606B24A73D6BD8D60056F39567ECF9F5E0A3BA118A137CDF998E513468349B8'; Rounds = 3; Load = 105.9; First = 3.35; Whole = 0.767; Edit = 1.496; Calc = 142.5; WholePerSecond = 36509840; EditsPerSecond = 668.4; SumBuilt = 14800794; SumEdited = 14900794; ClockMs = 0.174; ClockSpread = 5.7; ClockDrift = -2.4 }
    @{ Half = 'native'; Fixture = 'life'; Formulas = 62964; Sha256 = '9C3909801C7219371ED2A431960AE8C2C070694D529841688D867F73271DAB70'; Rounds = 3; Load = 1352.9; First = 8.74; Whole = 1.750; Edit = 2.071; Calc = 2841.4; WholePerSecond = 35985598; EditsPerSecond = 482.8; SumBuilt = 23683; SumEdited = 9603; ClockMs = 0.171; ClockSpread = 5.9; ClockDrift = -4.1 }
    # Excel's half, its flag first:
    @{ Half = 'excel-flag'; File = 'calcpr_built'; C2 = 15; Saved = 'True'; Verdict = 'computed on opening' }
    @{ Half = 'excel-flag'; File = 'calcpr_calcid'; C2 = 15; Saved = 'True'; Verdict = 'computed on opening' }
    @{ Half = 'excel-flag'; File = 'calcpr_none'; C2 = 15; Saved = 'True'; Verdict = 'computed on opening' }
    @{ Half = 'excel'; Fixture = 'column'; Threads = 'all'; ThreadCount = 20; Sha256 = '97C22A696D71AA680D997C3778FDA6960D4945A972FFDE5A6F8D729F8CFF4304'; Open = 149.3; Whole = 2.563; Rebuild = 11.623; Edit = 2.767; Copy = 0.0; ComFloor = 0.256; SumBuilt = 1100080000; SumEdited = 1425100000 }
    @{ Half = 'excel'; Fixture = 'column'; Threads = 'one'; ThreadCount = 1; Sha256 = '97C22A696D71AA680D997C3778FDA6960D4945A972FFDE5A6F8D729F8CFF4304'; Open = 149.3; Whole = 2.945; Rebuild = 12.349; Edit = 2.389; Copy = 0.0; ComFloor = 0.244; SumBuilt = 1100080000; SumEdited = 1425100000 }
    @{ Half = 'excel'; Fixture = 'chain'; Threads = 'all'; ThreadCount = 20; Sha256 = 'DC691ACEB3B881FB705CC6CA0CDB48E83C209495535EF67CEC5AF029FB9D77E4'; Open = 125.9; Whole = 3.362; Rebuild = 7.022; Edit = 2.482; Copy = 0.0; ComFloor = 0.231; SumBuilt = 1200039974; SumEdited = 1200059975 }
    @{ Half = 'excel'; Fixture = 'chain'; Threads = 'one'; ThreadCount = 1; Sha256 = 'DC691ACEB3B881FB705CC6CA0CDB48E83C209495535EF67CEC5AF029FB9D77E4'; Open = 125.9; Whole = 2.591; Rebuild = 6.562; Edit = 0.851; Copy = 0.0; ComFloor = 0.226; SumBuilt = 1200039974; SumEdited = 1200059975 }
    @{ Half = 'excel'; Fixture = 'sums'; Threads = 'all'; ThreadCount = 20; Sha256 = '9606B24A73D6BD8D60056F39567ECF9F5E0A3BA118A137CDF998E513468349B8'; Open = 66.9; Whole = 1.947; Rebuild = 7.140; Edit = 1.834; Copy = 0.0; ComFloor = 0.265; SumBuilt = 14800794; SumEdited = 14900794 }
    @{ Half = 'excel'; Fixture = 'sums'; Threads = 'one'; ThreadCount = 1; Sha256 = '9606B24A73D6BD8D60056F39567ECF9F5E0A3BA118A137CDF998E513468349B8'; Open = 66.9; Whole = 3.670; Rebuild = 8.846; Edit = 3.871; Copy = 0.0; ComFloor = 0.250; SumBuilt = 14800794; SumEdited = 14900794 }
    @{ Half = 'excel'; Fixture = 'life'; Threads = 'all'; ThreadCount = 20; Sha256 = '9C3909801C7219371ED2A431960AE8C2C070694D529841688D867F73271DAB70'; Open = 2422.8; Whole = 4.807; Rebuild = 1156.417; Edit = 4.375; Copy = 78.2; ComFloor = 0.220; SumBuilt = 23683; SumEdited = 9603; CopyWay = 'paste values'; CopyRefused = 'the range itself: Specified cast is not valid.' }
    @{ Half = 'excel'; Fixture = 'life'; Threads = 'one'; ThreadCount = 1; Sha256 = '9C3909801C7219371ED2A431960AE8C2C070694D529841688D867F73271DAB70'; Open = 2422.8; Whole = 24.594; Rebuild = 1171.006; Edit = 25.163; Copy = 77.8; ComFloor = 0.237; SumBuilt = 23683; SumEdited = 9603; CopyWay = 'paste values'; CopyRefused = 'the range itself: Specified cast is not valid.' }
)

# --- the floors, a dated history that never goes down ---
# One entry a measurement, one a line inside the parentheses, a comma after
# the entry before it: @{ Date = 'yyyy-mm-dd'; Reason = ''; Floors = @{
# 'column whole' = n; 'column edit' = n; ... } }, whole frames in formulas a
# second and edits a second. An entry lower than the one before it in any
# floor needs its Reason.
$floorHistory = @(
    @{ Date = '2026-10-11'; Reason = ''; Floors = @{ 'column whole' = 47700000; 'column edit' = 353; 'chain whole' = 10600000; 'chain edit' = 318; 'sums whole' = 5560000; 'sums edit' = 176; 'life whole' = 15400000; 'life edit' = 314 } },
    @{ Date = '2026-10-12'; Reason = 'KERNEL.28, at the owner''s ruling: column edit and chain edit, rows the item does not move, lowered to this run less the margin; the unchanged code (5607910) took 2.83 to 3.23 and 2.92 to 3.58 ms an edit in nine runs on 2026-10-12, against the 2.41 and 2.67 of the run of 2026-10-11 the floors were set from (docs/BETA_REARVIEW.md, KERNEL.28)'; Floors = @{ 'column whole' = 52500000; 'column edit' = 272; 'chain whole' = 11900000; 'chain edit' = 233; 'sums whole' = 7040000; 'sums edit' = 223; 'life whole' = 15400000; 'life edit' = 314 } },
    @{ Date = '2026-10-16'; Reason = ''; Floors = @{ 'column whole' = 55200000; 'column edit' = 478; 'chain whole' = 14000000; 'chain edit' = 248; 'sums whole' = 7080000; 'sums edit' = 241; 'life whole' = 28500000; 'life edit' = 368 } },
    @{ Date = '2026-10-16'; Reason = ''; Floors = @{ 'column whole' = 56800000; 'column edit' = 478; 'chain whole' = 63000000; 'chain edit' = 609; 'sums whole' = 7210000; 'sums edit' = 247; 'life whole' = 28500000; 'life edit' = 380 } },
    @{ Date = '2026-10-17'; Reason = ''; Floors = @{ 'column whole' = 108000000; 'column edit' = 583; 'chain whole' = 63000000; 'chain edit' = 638; 'sums whole' = 31000000; 'sums edit' = 568; 'life whole' = 30500000; 'life edit' = 410 } }
)

# --- the programs measured, each one's SHA-256 over LF bytes ---
# 2026-10-10, KERNEL.20's sixth slice: the four programs as the slice wrote
# them, scripts/recalc/bench/<name>.txt.
$programs = [ordered]@{
    'column' = '751885133C249C25C8C4BB3854877A72E6054B2E0945BB2C1931C3FF05D55E2E'
    'chain'  = '72FA6E5DB80DF71796473E21B89432C598AD8719C473D8955B24B796417A8A5A'
    'sums'   = 'A339D809DE2493BBA3A129FC3FC4A5AEC458A5A6F1B5FB590DB04D2CCB3F7348'
    'life'   = 'C2B7AC34005AB88CFE2F4CE98FAC3C704CCF26E8AEED90BE221BF800D63F1716'
}

# --- the rules ---
$fixtures = @('column', 'chain', 'sums', 'life')
$modes = @('all', 'one')
$flagFiles = @('calcpr_built', 'calcpr_calcid', 'calcpr_none')
$verdicts = @('kept the folded value', 'computed on opening', 'neither')
$minRounds = 3
$clockMax = 10.0
$agreeWithin = 0.005

function Get-LfDigest([string]$path) {
    $text = [System.IO.File]::ReadAllText($path) -replace "`r`n", "`n"
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($text)) } finally { $sha.Dispose() }
    return (($hash | ForEach-Object { $_.ToString('X2') }) -join '')
}

function Test-Near([double]$a, [double]$b) {
    if ($b -eq 0) { return ($a -eq 0) }
    return ([math]::Abs($a / $b - 1.0) -le $agreeWithin)
}

# The whole judgment of a baseline, its history and the programs as they
# stand; the problems, one a line.
function Test-Floors($date, $native, $excel, $rows, $history, $measured, $now) {
    $problems = New-Object System.Collections.Generic.List[string]
    $rows = @($rows | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0 -or $date -eq '') {
        $problems.Add('the baseline is empty: run tools\bench_recalc.ps1 and tools\bench_recalc_excel.ps1 on the owner''s machine and paste both blocks')
        return $problems.ToArray()
    }
    # The environments.
    if ($native -notmatch 'power throttling opted out: the runner yes') { $problems.Add('the native half did not run opted out of Windows'' power throttling: its environment line does not say so') }
    if ($native -notmatch 'high priority: the runner yes') { $problems.Add('the native half did not run at high priority: its environment line does not say so') }
    if ($native -notmatch 'on the performance cores, .*: the runner yes' -and $native -notmatch 'one kind of core') { $problems.Add('the native half did not run on the performance cores: its environment line does not say so') }
    if ($excel -notmatch 'power throttling opted out: yes') { $problems.Add('Excel''s half did not run opted out of Windows'' power throttling: its environment line does not say so') }
    if ($excel -match 'STAND-IN' -or $excel -eq '') { $problems.Add('Excel''s half is the instrument''s stand-in, or missing: no Excel was measured') }
    # The native rows.
    $nat = @{}
    foreach ($r in @($rows | Where-Object { $_.Half -eq 'native' })) {
        if ($nat.ContainsKey($r.Fixture)) { $problems.Add("native $($r.Fixture): two rows"); continue }
        $nat[$r.Fixture] = $r
    }
    foreach ($f in $fixtures) {
        if (-not $nat.ContainsKey($f)) { $problems.Add("native $($f): no row"); continue }
        $r = $nat[$f]
        if ([int]$r.Rounds -lt $minRounds) { $problems.Add("native $($f): measured in $($r.Rounds) round(s), fewer than $minRounds") }
        if ([math]::Abs([double]$r.ClockSpread) -gt $clockMax -or [math]::Abs([double]$r.ClockDrift) -gt $clockMax) { $problems.Add("native $($f): the clock moved, spread $($r.ClockSpread)% and drift $($r.ClockDrift)%, more than $clockMax%") }
        if (-not (Test-Near ([double]$r.WholePerSecond) ([double]$r.Formulas / ([double]$r.Whole / 1000.0)))) { $problems.Add("native $($f): $($r.WholePerSecond) formulas a second is not $($r.Formulas) formulas in $($r.Whole) ms") }
        if (-not (Test-Near ([double]$r.EditsPerSecond) (1000.0 / [double]$r.Edit))) { $problems.Add("native $($f): $($r.EditsPerSecond) edits a second is not one in $($r.Edit) ms") }
    }
    # Excel's rows: the same bytes, the same sums.
    $xl = @{}
    foreach ($r in @($rows | Where-Object { $_.Half -eq 'excel' })) {
        $k = "$($r.Fixture) $($r.Threads)"
        if ($xl.ContainsKey($k)) { $problems.Add("excel $($k): two rows"); continue }
        $xl[$k] = $r
    }
    foreach ($f in $fixtures) {
        foreach ($mode in $modes) {
            $k = "$f $mode"
            if (-not $xl.ContainsKey($k)) { $problems.Add("excel $($k): no row"); continue }
            $r = $xl[$k]
            if ($r.ContainsKey('Failed')) { $problems.Add("excel $($k): Excel could not measure it: $($r.Failed)"); continue }
            if (-not $nat.ContainsKey($f)) { continue }
            $n = $nat[$f]
            if ($r.Sha256 -ne $n.Sha256) { $problems.Add("excel $($k): measured $($r.Sha256), not the native half's file $($n.Sha256)") }
            if ([double]$r.SumBuilt -ne [double]$n.SumBuilt) { $problems.Add("excel $($k): summed $($r.SumBuilt) as built where the machine summed $($n.SumBuilt)") }
            if ([double]$r.SumEdited -ne [double]$n.SumEdited) { $problems.Add("excel $($k): summed $($r.SumEdited) edited where the machine summed $($n.SumEdited)") }
            if ($mode -eq 'one' -and [int]$r.ThreadCount -ne 1) { $problems.Add("excel $($k): $($r.ThreadCount) threads, not one") }
        }
    }
    # The flag.
    $flags = @($rows | Where-Object { $_.Half -eq 'excel-flag' })
    foreach ($file in $flagFiles) {
        $match = @($flags | Where-Object { $_.File -eq $file })
        if ($match.Count -ne 1) { $problems.Add("the flag: $($match.Count) row(s) for $file, not one"); continue }
        if ($verdicts -notcontains $match[0].Verdict) { $problems.Add("the flag: $file's verdict '$($match[0].Verdict)' is none of the three") }
    }
    # The programs.
    foreach ($f in $fixtures) {
        if (-not $now.Contains($f)) { $problems.Add("the program $($f): not found"); continue }
        if ($now[$f] -ne $measured[$f]) { $problems.Add("the program $($f): its SHA-256 over LF is $($now[$f]), not the measured $($measured[$f]); measure it again") }
    }
    # The floors.
    $history = @($history | Where-Object { $null -ne $_ })
    if ($history.Count -eq 0) { $problems.Add('the floor history is empty'); return $problems.ToArray() }
    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($f in $fixtures) { $keys.Add("$f whole"); $keys.Add("$f edit") }
    for ($i = 0; $i -lt $history.Count; $i++) {
        $e = $history[$i]
        foreach ($k in $keys) {
            if ($null -eq $e.Floors -or -not $e.Floors.ContainsKey($k)) { $problems.Add("the floor history's entry of $($e.Date): no floor '$k'"); continue }
            if ($i -gt 0 -and $history[$i - 1].Floors.ContainsKey($k) -and [double]$e.Floors[$k] -lt [double]$history[$i - 1].Floors[$k] -and [string]::IsNullOrWhiteSpace($e.Reason)) {
                $problems.Add("the floor history's entry of $($e.Date): '$k' is $($e.Floors[$k]), under the $($history[$i - 1].Floors[$k]) before it, with no reason written")
            }
        }
    }
    $latest = $history[$history.Count - 1]
    foreach ($f in $fixtures) {
        if (-not $nat.ContainsKey($f)) { continue }
        $n = $nat[$f]
        if ($latest.Floors.ContainsKey("$f whole") -and [double]$n.WholePerSecond -lt [double]$latest.Floors["$f whole"]) { $problems.Add("native $($f): $($n.WholePerSecond) formulas a second, under the floor of $($latest.Floors["$f whole"])") }
        if ($latest.Floors.ContainsKey("$f edit") -and [double]$n.EditsPerSecond -lt [double]$latest.Floors["$f edit"]) { $problems.Add("native $($f): $($n.EditsPerSecond) edits a second, under the floor of $($latest.Floors["$f edit"])") }
    }
    return $problems.ToArray()
}

# The history's entries as text, one a line, from a copy of this file.
function Get-HistoryLines([string]$text) {
    $entries = New-Object System.Collections.Generic.List[string]
    $inside = $false
    foreach ($line in ($text -split "`r?`n")) {
        if (-not $inside) { if ($line -match '^\$floorHistory = @\(\s*$') { $inside = $true }; continue }
        if ($line -match '^\)\s*$') { break }
        $t = $line.Trim().TrimEnd(',').Trim()
        if ($t -ne '' -and -not $t.StartsWith('#')) { $entries.Add($t) }
    }
    return $entries.ToArray()
}

# Every committed entry still here, in its order, at the head.
function Test-Appended($committed, $current) {
    $problems = New-Object System.Collections.Generic.List[string]
    $committed = @($committed | Where-Object { $null -ne $_ }); $current = @($current | Where-Object { $null -ne $_ })
    for ($i = 0; $i -lt $committed.Count; $i++) {
        if ($i -ge $current.Count -or $current[$i] -ne $committed[$i]) {
            $problems.Add("the floor history's committed entry $($i + 1) is not where the last commit had it: an entry was pasted over or edited; put it back and add the new entry after it")
        }
    }
    return $problems.ToArray()
}

# A runner's check mode judged against the baseline's sums.
function Test-Runner([string[]]$lines, $rows) {
    $problems = New-Object System.Collections.Generic.List[string]
    $info = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($l in $lines) {
        if ($l -match '^\(recalc-differs ') { $problems.Add("the runner: $l") }
        if ($l -match '^\(recalc-flag-failed ') { $problems.Add("the runner: $l") }
        $m = [regex]::Match($l, '^\(recalc-check "([a-z]+)" \(sha256 "([0-9A-F]{64})"\) \(formulas (\d+)\) \(held (\d+)\) \(sum-built ([-0-9.e+]+)\) \(sum-edited ([-0-9.e+]+)\)\)$')
        if ($m.Success) { $seen[$m.Groups[1].Value] = @{ Sha256 = $m.Groups[2].Value; Held = [long]$m.Groups[4].Value; SumBuilt = [double]$m.Groups[5].Value; SumEdited = [double]$m.Groups[6].Value } }
    }
    if (-not ($lines -match '^\(recalc-check-done 4\)$')) { $problems.Add('the runner did not finish its check mode over the four fixtures') }
    foreach ($f in $fixtures) {
        if (-not $seen.ContainsKey($f)) { $problems.Add("the runner: no check row for $f"); continue }
        $s = $seen[$f]
        if ($s.Held -le 0) { $problems.Add("the runner: $f held no value") }
        $n = @($rows | Where-Object { $null -ne $_ -and $_.Half -eq 'native' -and $_.Fixture -eq $f })
        if ($n.Count -ne 1) { continue }
        if ([double]$n[0].SumBuilt -ne $s.SumBuilt -or [double]$n[0].SumEdited -ne $s.SumEdited) { $problems.Add("the runner: $f sums $($s.SumBuilt) and $($s.SumEdited) today where the baseline measured $($n[0].SumBuilt) and $($n[0].SumEdited)") }
        if ($n[0].Sha256 -ne $s.Sha256) { $info.Add("  $($f): today's file is $($s.Sha256.Substring(0, 12))..., the measured one $($n[0].Sha256.Substring(0, 12))...; its values are the measured ones (the stamp or the writer moved, not the computation)") }
    }
    return @{ Problems = $problems; Info = $info }
}

function Invoke-Runner([string]$exe) {
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        if ($exe.EndsWith('.ps1')) {
            $shell = if (Get-Command powershell -ErrorAction SilentlyContinue) { 'powershell' } else { 'pwsh' }
            $out = & $shell -NoProfile -ExecutionPolicy Bypass -File $exe --check 2>$null
        } else {
            $out = & $exe --check 2>$null
        }
    } finally { $ErrorActionPreference = $old }
    return @($out | ForEach-Object { [string]$_ })
}

# ------------------------------------------------------------------ -Control
if ($Control) {
    $failures = New-Object System.Collections.Generic.List[string]
    $okEnvNative = 'cpu | os | rustc | frazaro-core 0.8.1 | power throttling opted out: the runner yes | high priority: the runner yes | on the performance cores, 8 of 20 logical processors (efficiency class 1): the runner yes'
    $okEnvExcel = 'cpu | os | Excel 16.0 build 19127 | all threads 20; one thread on the performance cores, mask 0xC03C3 | power throttling opted out: yes'
    function New-Clean {
        $rows = New-Object System.Collections.Generic.List[object]
        $sums = @{ 'column' = @(1100080000, 1425100000); 'chain' = @(1200039974, 1200059975); 'sums' = @(14800794, 14900794); 'life' = @(23683, 9603) }
        $sha = @{ 'column' = ('A' * 64); 'chain' = ('B' * 64); 'sums' = ('C' * 64); 'life' = ('D' * 64) }
        foreach ($f in $fixtures) {
            $rows.Add(@{ Half = 'native'; Fixture = $f; Formulas = 60000; Sha256 = $sha[$f]; Rounds = 3; Load = 200.0; First = 15.0; Whole = 3.0; Edit = 5.0; Calc = 200.0; WholePerSecond = 20000000; EditsPerSecond = 200.0; SumBuilt = $sums[$f][0]; SumEdited = $sums[$f][1]; ClockMs = 0.182; ClockSpread = 1.0; ClockDrift = 0.5 })
            foreach ($mode in $modes) {
                $rows.Add(@{ Half = 'excel'; Fixture = $f; Threads = $mode; ThreadCount = $(if ($mode -eq 'one') { 1 } else { 20 }); Sha256 = $sha[$f]; Open = 900.0; Whole = 4.0; Rebuild = 9.0; Edit = 3.0; Copy = 0.0; ComFloor = 0.05; SumBuilt = $sums[$f][0]; SumEdited = $sums[$f][1] })
            }
        }
        foreach ($file in $flagFiles) { $rows.Add(@{ Half = 'excel-flag'; File = $file; C2 = 15; Saved = 'False'; Verdict = 'computed on opening' }) }
        $floors = @{}
        foreach ($f in $fixtures) { $floors["$f whole"] = 18000000; $floors["$f edit"] = 180 }
        $measured = [ordered]@{}; $now = [ordered]@{}
        foreach ($k in $programs.Keys) { $measured[$k] = $programs[$k]; $now[$k] = $programs[$k] }
        return @{ Rows = $rows; History = @(@{ Date = '2026-10-10'; Reason = ''; Floors = $floors }); Native = $okEnvNative; Excel = $okEnvExcel; Programs = $measured; Now = $now }
    }
    function Get-Problems($c) { return (Test-Floors '2026-10-10' $c.Native $c.Excel $c.Rows $c.History $c.Programs $c.Now) }
    function Expect([string]$what, $c, [bool]$clean) {
        $p = @(Get-Problems $c)
        if ($clean -and $p.Count -ne 0) { $failures.Add("control: $what should pass and fails: $($p -join ' / ')") }
        elseif (-not $clean -and $p.Count -eq 0) { $failures.Add("control: $what should fail and passes") }
        else { Write-Output ("  control: {0}: {1}" -f $what, $(if ($clean) { 'passes' } else { "fails ($($p.Count))" })) }
    }
    Expect 'a clean baseline' (New-Clean) $true
    $c = New-Clean; $prev = @{}; foreach ($k in $c.History[0].Floors.Keys) { $prev[$k] = $c.History[0].Floors[$k] + 1000 }; $c.History = @(@{ Date = '2026-10-09'; Reason = ''; Floors = $prev }, @{ Date = '2026-10-10'; Reason = 'a correctness fix that costs speed'; Floors = $c.History[0].Floors }); Expect 'a floor lowered for a written reason' $c $true
    $c = New-Clean; $c.Rows = @(); Expect 'an empty baseline' $c $false
    $c = New-Clean; $c.Rows = @($c.Rows | Where-Object { -not ($_.Half -eq 'native' -and $_.Fixture -eq 'life') }); Expect 'a native row missing' $c $false
    $c = New-Clean; $c.Rows = @($c.Rows | Where-Object { -not ($_.Half -eq 'excel' -and $_.Fixture -eq 'chain' -and $_.Threads -eq 'one') }); Expect 'an Excel row missing' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'excel' -and $_.Fixture -eq 'sums' -and $_.Threads -eq 'all' })[0].Sha256 = ('E' * 64); Expect 'a file not the same bytes in both halves' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'excel' -and $_.Fixture -eq 'life' -and $_.Threads -eq 'one' })[0].SumEdited = 9604; Expect 'sums that differ' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq 'column' })[0].ClockSpread = 12.0; Expect 'a clock that moved' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq 'chain' })[0].Rounds = 1; Expect 'too few rounds' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq 'sums' })[0].WholePerSecond = 21000000; Expect 'formulas a second that disagree with the time' $c $false
    $c = New-Clean; $c.Native = $c.Native -replace 'the runner yes \| high', 'the runner no | high'; Expect 'a native half not opted out' $c $false
    $c = New-Clean; $c.Excel = $c.Excel + ' | THE STAND-IN, NOT EXCEL'; Expect 'Excel''s half from the stand-in' $c $false
    $c = New-Clean; $prev = @{}; foreach ($k in $c.History[0].Floors.Keys) { $prev[$k] = $c.History[0].Floors[$k] + 1000 }; $c.History = @(@{ Date = '2026-10-09'; Reason = ''; Floors = $prev }, $c.History[0]); Expect 'a floor lowered with no reason' $c $false
    $c = New-Clean; $c.History[0].Floors['life edit'] = 250; Expect 'a row under its floor' $c $false
    $c = New-Clean; $c.Now['chain'] = ('0' * 64); Expect 'a program edited' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'excel-flag' })[1].Verdict = 'maybe'; Expect 'a flag row with no verdict' $c $false
    $c = New-Clean; $c.Rows = @($c.Rows | Where-Object { -not ($_.Half -eq 'excel' -and $_.Fixture -eq 'life') }) + @(@{ Half = 'excel'; Fixture = 'life'; Threads = 'all'; Sha256 = ('D' * 64); Failed = 'Unable to get the Open property of the Workbooks class' }, @{ Half = 'excel'; Fixture = 'life'; Threads = 'one'; Sha256 = ('D' * 64); Failed = 'Unable to get the Open property of the Workbooks class' }); Expect 'a fixture Excel could not open' $c $false
    # The history, from text.
    $committed = Get-HistoryLines "`$floorHistory = @(`r`n    @{ Date = '2026-10-10'; Reason = ''; Floors = @{ 'column whole' = 1 } },`r`n    @{ Date = '2026-10-11'; Reason = ''; Floors = @{ 'column whole' = 2 } }`r`n)"
    $cases = @(
        @{ What = 'an entry added after the committed ones'; Clean = $true; Text = "`$floorHistory = @(`n    @{ Date = '2026-10-10'; Reason = ''; Floors = @{ 'column whole' = 1 } },`n    @{ Date = '2026-10-11'; Reason = ''; Floors = @{ 'column whole' = 2 } },`n    @{ Date = '2026-10-12'; Reason = ''; Floors = @{ 'column whole' = 3 } }`n)" },
        @{ What = 'a history unchanged'; Clean = $true; Text = "`$floorHistory = @(`n    @{ Date = '2026-10-10'; Reason = ''; Floors = @{ 'column whole' = 1 } },`n    @{ Date = '2026-10-11'; Reason = ''; Floors = @{ 'column whole' = 2 } }`n)" },
        @{ What = 'an entry pasted over a committed one'; Clean = $false; Text = "`$floorHistory = @(`n    @{ Date = '2026-10-11'; Reason = ''; Floors = @{ 'column whole' = 2 } },`n    @{ Date = '2026-10-12'; Reason = ''; Floors = @{ 'column whole' = 3 } }`n)" },
        @{ What = 'a committed entry edited in place'; Clean = $false; Text = "`$floorHistory = @(`n    @{ Date = '2026-10-10'; Reason = ''; Floors = @{ 'column whole' = 1 } },`n    @{ Date = '2026-10-11'; Reason = ''; Floors = @{ 'column whole' = 1 } }`n)" }
    )
    foreach ($case in $cases) {
        $p = @(Test-Appended $committed (Get-HistoryLines $case.Text))
        if ($case.Clean -and $p.Count -ne 0) { $failures.Add("control: the history, $($case.What), should pass and fails") }
        elseif (-not $case.Clean -and $p.Count -eq 0) { $failures.Add("control: the history, $($case.What), should fail and passes") }
        else { Write-Output ("  control: the history, {0}: {1}" -f $case.What, $(if ($case.Clean) { 'passes' } else { 'fails' })) }
    }
    # Fake runners: scripts printing a check mode.
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('recalc_floors_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $clean = New-Clean
        $good = New-Object System.Collections.Generic.List[string]
        foreach ($f in $fixtures) {
            $n = @($clean.Rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq $f })[0]
            $good.Add(('(recalc-check "{0}" (sha256 "{1}") (formulas 60000) (held 120000) (sum-built {2}) (sum-edited {3}))' -f $f, $n.Sha256, $n.SumBuilt, $n.SumEdited))
        }
        $good.Add('(recalc-check-flags 3)'); $good.Add('(recalc-check-done 4)')
        $fakes = @(
            @{ What = 'a clean runner'; Clean = $true; Lines = $good.ToArray() },
            @{ What = 'a runner whose values differ'; Clean = $false; Lines = @('(recalc-differs "chain: Output!R5C2: the machine holds 1, Calc 2")') },
            @{ What = 'a runner whose sums differ'; Clean = $false; Lines = @($good.ToArray() | ForEach-Object { $_ -replace '\(sum-edited 9603\)', '(sum-edited 9600)' }) },
            @{ What = 'a runner that leaves out a fixture'; Clean = $false; Lines = @($good.ToArray() | Where-Object { $_ -notmatch '"sums"' }) }
        )
        $k = 0
        foreach ($fake in $fakes) {
            $k++
            $script = Join-Path $tmp "fake$k.ps1"
            $body = ($fake.Lines | ForEach-Object { "Write-Output '" + ($_ -replace "'", "''") + "'" }) -join "`r`n"
            [System.IO.File]::WriteAllText($script, $body + "`r`n", (New-Object System.Text.UTF8Encoding($false)))
            $judged = Test-Runner (Invoke-Runner $script) $clean.Rows
            $n = $judged.Problems.Count
            if ($fake.Clean -and $n -ne 0) { $failures.Add("control: $($fake.What) should pass and fails: $($judged.Problems -join ' / ')") }
            elseif (-not $fake.Clean -and $n -eq 0) { $failures.Add("control: $($fake.What) should fail and passes") }
            else { Write-Output ("  control: {0}: {1}" -f $fake.What, $(if ($fake.Clean) { 'passes' } else { "fails ($n)" })) }
        }
    } finally { Remove-Item -Recurse -Force -LiteralPath $tmp -ErrorAction SilentlyContinue }
    if ($failures.Count -gt 0) { foreach ($f in $failures) { Write-Output "FAIL: $f" }; exit 1 }
    Write-Output 'OK: check_recalc_floors -Control: every clean case passes and every mutant fails'
    exit 0
}

# ------------------------------------------------------------------ the check
$problems = New-Object System.Collections.Generic.List[string]
$now = [ordered]@{}
foreach ($f in $fixtures) {
    $path = Join-Path $Root "scripts/recalc/bench/$f.txt"
    if (Test-Path -LiteralPath $path) { $now[$f] = Get-LfDigest $path }
}
foreach ($p in @(Test-Floors $baselineDate $baselineNative $baselineExcel $baseline $floorHistory $programs $now)) { $problems.Add($p) }

# The history against the last commit's copy of this file.
$appended = 'SKIPPED (no git)'
$git = Get-Command git -ErrorAction SilentlyContinue
if ($git) {
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $committedText = & git -C $Root show 'HEAD:tools/check_recalc_floors.ps1' 2>$null | Out-String; $gitExit = $LASTEXITCODE } finally { $ErrorActionPreference = $old }
    if ($gitExit -eq 0 -and $committedText) {
        $self = [System.IO.File]::ReadAllText((Join-Path $Root 'tools/check_recalc_floors.ps1'))
        foreach ($p in @(Test-Appended (Get-HistoryLines $committedText) (Get-HistoryLines $self))) { $problems.Add($p) }
        $appended = 'held against the last commit''s copy'
    } else { $appended = 'SKIPPED (not in the last commit yet)' }
}

# The runner's check mode.
$runnerWord = ''
$info = @()
if ($Runner -eq '') { $Runner = Join-Path $Root ('target/release/examples/recalc' + $(if ($onWindows) { '.exe' } else { '' })) }
if (Test-Path -LiteralPath $Runner) {
    $judged = Test-Runner (Invoke-Runner $Runner) $baseline
    foreach ($p in $judged.Problems) { $problems.Add($p) }
    $info = @($judged.Info)
    $runnerWord = 'the runner''s check mode: every value held, the sums the measured ones'
} else { $runnerWord = 'SKIPPED the runner (none built)' }

# The bar: Excel beside the machine.
$rows = @($baseline | Where-Object { $null -ne $_ })
if ($rows.Count -gt 0) {
    Write-Output ("check_recalc_floors: the baseline of {0}, the machine against Excel on the same files:" -f $baselineDate)
    foreach ($f in $fixtures) {
        $n = @($rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq $f })
        $a = @($rows | Where-Object { $_.Half -eq 'excel' -and $_.Fixture -eq $f -and $_.Threads -eq 'all' })
        $o = @($rows | Where-Object { $_.Half -eq 'excel' -and $_.Fixture -eq $f -and $_.Threads -eq 'one' })
        if ($n.Count -ne 1 -or $a.Count -ne 1 -or $o.Count -ne 1) { continue }
        $ratio = { param($x, $y) if ([double]$y -le 0) { 'n/a' } elseif ([double]$x -le [double]$y) { '{0:F1}x faster' -f ([double]$y / [double]$x) } else { '{0:F1}x slower' -f ([double]$x / [double]$y) } }
        Write-Output ("  {0}: whole {1} ms, Excel {2} ms on {3} threads ({4}), {5} ms on one ({6}); the edit {7} ms, Excel {8} and {9} ms ({10}, {11})" -f $f, $n[0].Whole, $a[0].Whole, $a[0].ThreadCount, (& $ratio $n[0].Whole $a[0].Whole), $o[0].Whole, (& $ratio $n[0].Whole $o[0].Whole), $n[0].Edit, $a[0].Edit, $o[0].Edit, (& $ratio $n[0].Edit $a[0].Edit), (& $ratio $n[0].Edit $o[0].Edit))
    }
    foreach ($r in @($rows | Where-Object { $_.Half -eq 'excel-flag' })) { Write-Output ("  the flag: {0}, Output!C2 read {1} before any calculation: {2}; saved {3}" -f $r.File, $r.C2, $r.Verdict, $r.Saved) }
}
foreach ($i in $info) { Write-Output $i }
if ($problems.Count -gt 0) {
    foreach ($p in $problems) { Write-Output "FAIL: $p" }
    exit 1
}
Write-Output ("OK: check_recalc_floors: four fixtures measured against Excel, the floors held, the history {0}; {1}" -f $appended, $runnerWord)
exit 0
