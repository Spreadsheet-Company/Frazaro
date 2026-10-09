<#
check_recalc_golden.ps1 - the core reproduces the recalc goldens whole; each
golden's line count and its agree count never go down; the fixed order and
the verdict pins hold in the golden itself; the study's workbooks agree.

WHY: KERNEL.7 computes a workbook's formulas in the language's evaluator
(vla-lang/src/calc/) under Excel's library (core/src/excel.rs), the declared
subset, and `frazaro calc` prints each formula cell's computed value against
the value the host saved into the file, with a verdict: agree, differ or
unchecked. That is the treaty's twelfth oracle (conformance/README.md, the
amendment of 2026-10-08 for KERNEL.7): the goldens under scripts/recalc/ are
the door's own output, and the one fixture Excel 365 itself saved,
scripts/reflect/saved.xlsx, is the row whose every verdict must be agree.
The VBA reference has no recalculation of its own, Excel is its evaluator,
so Excel's saved values are the reference (CALLOSUM section 7, decision 3),
under one stated tolerance, fifteen significant digits, Excel's documented
precision. What this check holds is that the evaluator never drifts from
what agreed without a regenerated golden committed beside the change, and
that a second implementation prints the same lines.

ALSO HELD, read off each golden itself: its line count, as a floor that
never goes down; its agree count, likewise, where a fixture carries cached
values; the fixed order the amendment fixes - every cycle row first, then
one calc row per formula cell, a sheet's rows together and in row-major
order within it, each row sheet, address, the formula, the computed value
(a datum, or (not-computed "<name>") or (not-computed cycle)), the cached
value (a datum or none) and the verdict; and the pin each fixture earns:
the Excel-saved fixture every row agree, a fixture whose cached values a
script computed (two lineages of evaluation) no row differ, and the door's
own build, which carries no cached values, every row none unchecked.

THE STUDY'S WORKBOOKS (scripts/study/, KERNEL.3) are instruments held by
their own check and not treaty fixtures, so no golden is stored for them;
what is held here is their agreement alone: with a door, `frazaro calc
<workbook> --counts` must show at least the agree count below and no
differ, the PowerShell evaluator that wrote their cached values and the Rust
one agreeing on every formula. SKIPPED without a door, and SKIPPED while
scripts/study/ is not in the tree (KERNEL.3 lands it); held from the day it
is, and a workbook missing from a present folder fails.

WHAT IT RUNS: `<impl> calc <fixture>`, the contract's own command, on the
built door at target/debug/frazaro.exe (or target/release), or on -Impl
when given; its stdout is compared to the golden after line endings are
normalized to LF and trailing blank lines dropped, and the first differing
line is named with both texts. A tree with no door says SKIPPED for that
half and exits 0, as check_view_golden.ps1 does, since CI always has it;
the floors, the order and the pins are checked either way.

-Control proves the check on fakes written to a scratch directory: one that
prints the golden of each fixture and the study's agreement counts (must
pass every row); one that changes one value in the first calc row it prints
and reports one differ (must fail every golden row naming line 1 and every
study row); a copy of the saved golden with one agree turned to differ (the
pin must fail it); and a copy with a cycle row appended after the calc rows
(the order pin must fail it, and pass the golden as it is).

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no
COM, no network; exit 0 clean, exit 1 with every problem named. ASCII only,
since a BOM-less .ps1 is read as ANSI by PowerShell 5.1.

Usage:  powershell -File tools\check_recalc_golden.ps1
        powershell -File tools\check_recalc_golden.ps1 -Impl C:\path\to\frazaro.exe
        powershell -File tools\check_recalc_golden.ps1 -Control
#>
param(
    [string]$Impl = '',
    [switch]$Control
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# --- the goldens, their floors and their pins ---------------------------------
# 2026-10-08, KERNEL.7: saved 5 (Excel 365's save of the first build golden:
# the three arithmetic formulas, the shared children and IFS, every row
# agree); fixture 29 (the reader's fixture, every reference kind: 23 agree,
# six not computed and named - a structured reference, SEQUENCE, INDIRECT,
# a link into another workbook, OFFSET, a range where one value is wanted;
# no differ); build_fixture 5 (the core's own build, no cached values, every
# row none unchecked); opendocument 29 (the fixture written again in ODF's
# terms, read through the OpenDocument reader: 24 agree, five not computed,
# no differ); subset 106 (the declared subset's fixture, scripts/recalc/
# subset.txt built by the door, one formula per function and per coercion
# claimed, no cached values); subset_saved 106 (the owner's save of it from
# Excel 365, 2026-10-08: 102 agree, none differing, the four it does not
# compute named - VLOOKUP, CONCATENATE, the cell reading VLOOKUP, and the
# range where one value is wanted; it corrected two rules the evening it
# was saved, a number's General spelling and a text read through a
# reference in a logical test); edges 16 (scripts/recalc/edges.txt built by
# the door, the boundaries those two corrections chose, no cached values
# until an Excel save of it, edges_saved.xlsx, joins this table at
# no-differ). Raise a floor when a regenerated golden is longer; lower it
# only with a regenerated golden that is shorter, and say why.
#
# Pin: 'agree' - every verdict agree (a host saved the file and the subset
# computes all of it); 'no-differ' - no verdict differ (a host's save that
# holds formulas outside the subset, or a script's cached values, two
# lineages agreeing); 'unsaved' - every row none unchecked (the door's own
# build).
$goldens = @(
    @{ Name = 'saved';         Fixture = 'scripts/reflect/saved.xlsx';        Golden = 'scripts/recalc/saved_calc.vla';         Floor = 5;   Pin = 'agree';     AgreeFloor = 5 },
    @{ Name = 'fixture';       Fixture = 'scripts/reflect/fixture.xlsx';      Golden = 'scripts/recalc/fixture_calc.vla';       Floor = 29;  Pin = 'no-differ'; AgreeFloor = 23 },
    @{ Name = 'build_fixture'; Fixture = 'scripts/build/fixture_golden.xlsx'; Golden = 'scripts/recalc/build_fixture_calc.vla'; Floor = 5;   Pin = 'unsaved';   AgreeFloor = 0 },
    @{ Name = 'opendocument';  Fixture = 'scripts/reflect/opendocument.ods';  Golden = 'scripts/recalc/opendocument_calc.vla';  Floor = 29;  Pin = 'no-differ'; AgreeFloor = 24 },
    @{ Name = 'subset';        Fixture = 'scripts/recalc/subset.xlsx';        Golden = 'scripts/recalc/subset_calc.vla';        Floor = 106; Pin = 'unsaved';   AgreeFloor = 0 },
    @{ Name = 'subset_saved';  Fixture = 'scripts/recalc/subset_saved.xlsx';  Golden = 'scripts/recalc/subset_saved_calc.vla';  Floor = 106; Pin = 'no-differ'; AgreeFloor = 102 },
    @{ Name = 'edges';         Fixture = 'scripts/recalc/edges.xlsx';         Golden = 'scripts/recalc/edges_calc.vla';         Floor = 16;  Pin = 'unsaved';   AgreeFloor = 0 }
)

# --- the study's workbooks: agreement alone, no golden stored -----------------
# 2026-10-08, KERNEL.7: find_a and find_b 80 formulas each, build_a and
# build_b 8 each, every one agreeing with the cached value
# tools/build_study_fixture.ps1 computed; differ is pinned at 0.
$agreements = @(
    @{ Name = 'find_a';  Fixture = 'scripts/study/find_a.xlsx';  Agree = 80 },
    @{ Name = 'find_b';  Fixture = 'scripts/study/find_b.xlsx';  Agree = 80 },
    @{ Name = 'build_a'; Fixture = 'scripts/study/build_a.xlsx'; Agree = 8 },
    @{ Name = 'build_b'; Fixture = 'scripts/study/build_b.xlsx'; Agree = 8 }
)
# The study folder is KERNEL.3's; until it is committed the agreement is
# skipped, never failed, so this check stands on its own fixtures alone.
$studyPresent = Test-Path (Join-Path $repoRoot 'scripts/study')

# A text as its lines: line endings normalized, trailing blank lines dropped.
function Get-Lines([string]$text) {
    $t = $text -replace "`r`n", "`n"
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($l in ($t -split "`n")) { $lines.Add($l) }
    while ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq '') { $lines.RemoveAt($lines.Count - 1) }
    return ,$lines
}

# The relation a line is a row of, or '' when the line is not a row.
function Get-Relation([string]$line) {
    $m = [regex]::Match($line, '^\(([a-z]+) ')
    if (-not $m.Success) { return '' }
    return $m.Groups[1].Value
}

# The quoted strings of a form, in order, unescaped as WriteDatum escapes.
function Get-Strings([string]$line) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($m in [regex]::Matches($line, '"((?:[^"\\]|\\.)*)"')) { $out.Add(($m.Groups[1].Value -replace '\\(.)', '$1')) }
    return ,$out
}

# An A1 address as (row, column), or $null.
function Get-RowCol([string]$addr) {
    $m = [regex]::Match($addr, '^([A-Z]+)([0-9]+)$')
    if (-not $m.Success) { return $null }
    $col = 0
    foreach ($ch in $m.Groups[1].Value.ToCharArray()) { $col = $col * 26 + ([int][char]$ch - [int][char]'A' + 1) }
    return @([int]$m.Groups[2].Value, $col)
}

# A calc row's six fields, read off the line after its three quoted strings:
# the computed datum, the cached datum and the verdict, or $null when the
# line is not shaped as a calc row.
function Get-CalcFields([string]$line) {
    $m = [regex]::Match($line, '^\(calc "(?:[^"\\]|\\.)*" "(?:[^"\\]|\\.)*" "(?:[^"\\]|\\.)*" (.*) (agree|differ|unchecked)\)$')
    if (-not $m.Success) { return $null }
    $rest = $m.Groups[1].Value
    $cached = ''
    if ($rest -match ' none$') { $cached = 'none'; $computed = $rest.Substring(0, $rest.Length - 5) }
    elseif ($rest -match ' (\(error "[^"]*"\)|\(date "[^"]*"\)|"(?:[^"\\]|\\.)*"|true|false|-?[0-9][0-9.eE+-]*)$') {
        $cached = $Matches[1]; $computed = $rest.Substring(0, $rest.Length - $cached.Length - 1)
    } else { return $null }
    return @{ Computed = $computed; Cached = $cached; Verdict = $m.Groups[2].Value }
}

# The fixed order and the row shapes, read off a golden's lines, and the
# counts of each verdict: problems (empty when they hold) and the counts.
function Test-Order($lines) {
    $problems = New-Object System.Collections.Generic.List[string]
    $counts = @{ agree = 0; differ = 0; unchecked = 0; cycles = 0; rows = 0 }
    $inCalc = $false
    $sheet = $null
    $seenSheets = @{}
    $lastCell = $null
    $n = 0
    foreach ($line in $lines) {
        $n++
        $rel = Get-Relation $line
        if ($rel -eq 'cycle') {
            if ($inCalc) { $problems.Add("line ${n}: a cycle row after the calc rows began: $line") }
            $counts.cycles++
            if ((Get-Strings $line).Count -lt 1) { $problems.Add("line ${n}: a cycle row naming no cell: $line") }
            continue
        }
        if ($rel -ne 'calc') { $problems.Add("line ${n}: not a row of the calc record: $line"); continue }
        $inCalc = $true
        $counts.rows++
        $strings = Get-Strings $line
        if ($strings.Count -lt 3) { $problems.Add("line ${n}: a calc row short of its sheet, address and formula: $line"); continue }
        if (-not $strings[2].StartsWith('=')) { $problems.Add("line ${n}: the formula does not begin with =: $line") }
        $f = Get-CalcFields $line
        if ($null -eq $f) { $problems.Add("line ${n}: not shaped as a calc row (computed, cached, verdict): $line"); continue }
        $counts[$f.Verdict]++
        $notComputed = $f.Computed.StartsWith('(not-computed ')
        if ($f.Verdict -ne 'unchecked' -and ($notComputed -or $f.Cached -eq 'none')) { $problems.Add("line ${n}: a verdict of $($f.Verdict) with nothing to compare: $line") }
        if ($f.Verdict -eq 'unchecked' -and -not $notComputed -and $f.Cached -ne 'none' -and -not $f.Cached.StartsWith('(date ')) { $problems.Add("line ${n}: unchecked though both values are there: $line") }
        if ($strings[0] -ne $sheet) {
            if ($seenSheets.ContainsKey($strings[0])) { $problems.Add("line ${n}: the sheet's rows are not together: $line") }
            $seenSheets[$strings[0]] = $true
            $sheet = $strings[0]; $lastCell = $null
        }
        $rc = Get-RowCol $strings[1]
        if ($null -eq $rc) { $problems.Add("line ${n}: not an address: $line") }
        elseif ($null -ne $lastCell -and ($rc[0] -lt $lastCell[0] -or ($rc[0] -eq $lastCell[0] -and $rc[1] -le $lastCell[1]))) { $problems.Add("line ${n}: a cell out of row-major order: $line") }
        $lastCell = $rc
    }
    return @{ Problems = $problems; Counts = $counts }
}

# The pin a fixture earns, against the counts read off its golden.
function Test-Pin($g, $counts) {
    $problems = New-Object System.Collections.Generic.List[string]
    switch ($g.Pin) {
        'agree'     { if ($counts.rows -ne $counts.agree) { $problems.Add("$($g.Name): a host saved this fixture, so every row must agree; $($counts.agree) of $($counts.rows) do") } }
        'no-differ' { if ($counts.differ -ne 0) { $problems.Add("$($g.Name): $($counts.differ) row(s) differ from the cached values") } }
        'unsaved'   { if ($counts.rows -ne $counts.unchecked -or $counts.agree -ne 0 -or $counts.differ -ne 0) { $problems.Add("$($g.Name): the door's own build has no cached values, so every row must be unchecked") } }
    }
    if ($counts.agree -lt $g.AgreeFloor) { $problems.Add("$($g.Name): $($counts.agree) agree rows, below the floor of $($g.AgreeFloor)") }
    return ,$problems
}

# `<impl> <args>` with the door's stderr set aside: its exit code and every
# line of its stdout.
function Invoke-Door([string]$impl, [string[]]$doorArgs) {
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $global:LASTEXITCODE = 0
        $lines = & $impl @doorArgs 2>$null
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $old
    }
    if ($null -eq $code) { $code = 0 }
    $text = if ($null -eq $lines) { '' } else { (@($lines) | ForEach-Object { [string]$_ }) -join "`n" }
    return @{ ExitCode = [int]$code; Text = $text }
}

# One implementation against one golden: a list of problems, empty when its
# rows are the golden whole.
function Measure-Impl([string]$impl, $g) {
    $problems = New-Object System.Collections.Generic.List[string]
    $golden = Join-Path $repoRoot $g.Golden
    $fixture = Join-Path $repoRoot $g.Fixture
    if (-not (Test-Path $fixture)) { $problems.Add("$($g.Name): no fixture at $fixture"); return ,$problems }
    $r = Invoke-Door $impl @('calc', $fixture)
    if ($r.ExitCode -eq 3) { $problems.Add("$($g.Name): calc not attempted (exit 3), but the golden exists"); return ,$problems }
    if ($r.ExitCode -ne 0) { $problems.Add("$($g.Name): calc exited $($r.ExitCode)"); return ,$problems }
    $want = Get-Lines ([System.IO.File]::ReadAllText($golden))
    $got = Get-Lines $r.Text
    $n = [Math]::Min($want.Count, $got.Count)
    for ($i = 0; $i -lt $n; $i++) {
        if ($want[$i] -ne $got[$i]) {
            $problems.Add("$($g.Name): differs at line $($i + 1) of $($want.Count): golden $($want[$i]) / door $($got[$i])")
            return ,$problems
        }
    }
    if ($want.Count -ne $got.Count) { $problems.Add("$($g.Name): the door printed $($got.Count) lines, the golden holds $($want.Count); they agree to line $n") }
    return ,$problems
}

# One implementation against one study workbook's agreement: problems, and
# the counts line it printed.
function Measure-Agreement([string]$impl, $a) {
    $problems = New-Object System.Collections.Generic.List[string]
    $fixture = Join-Path $repoRoot $a.Fixture
    if (-not (Test-Path $fixture)) { $problems.Add("$($a.Name): no workbook at $fixture"); return @{ Problems = $problems; Line = '' } }
    $r = Invoke-Door $impl @('calc', $fixture, '--counts')
    if ($r.ExitCode -ne 0) { $problems.Add("$($a.Name): calc --counts exited $($r.ExitCode)"); return @{ Problems = $problems; Line = $r.Text } }
    $m = [regex]::Match($r.Text, 'agree (\d+) differ (\d+)')
    if (-not $m.Success) { $problems.Add("$($a.Name): no agree and differ counts in: $($r.Text)"); return @{ Problems = $problems; Line = $r.Text } }
    $agree = [int]$m.Groups[1].Value; $differ = [int]$m.Groups[2].Value
    if ($agree -lt $a.Agree) { $problems.Add("$($a.Name): $agree agree, below the floor of $($a.Agree)") }
    if ($differ -ne 0) { $problems.Add("$($a.Name): $differ row(s) differ from the cached values the builder computed") }
    return @{ Problems = $problems; Line = $r.Text }
}

# Every golden's floor, order and pin: problems, and lines to print.
function Test-Goldens() {
    $problems = New-Object System.Collections.Generic.List[string]
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($g in $goldens) {
        $path = Join-Path $repoRoot $g.Golden
        if (-not (Test-Path $path)) { $problems.Add("$($g.Name): no golden at $path"); continue }
        $rows = Get-Lines ([System.IO.File]::ReadAllText($path))
        if ($rows.Count -lt $g.Floor) {
            $problems.Add("$($g.Name): the golden holds $($rows.Count) lines, below its floor of $($g.Floor)")
            $lines.Add("  $($g.Name): $($rows.Count) lines, BELOW THE FLOOR of $($g.Floor)  FAIL")
        } elseif ($rows.Count -gt $g.Floor) {
            $lines.Add("  $($g.Name): $($rows.Count) lines, above the floor of $($g.Floor) - raise the floor in this file")
        } else {
            $lines.Add("  $($g.Name): $($g.Floor) lines, at its floor")
        }
        $order = Test-Order $rows
        foreach ($p in $order.Problems) { $problems.Add("$($g.Name): $p"); $lines.Add("  $($g.Name): $p  FAIL") }
        $c = $order.Counts
        if ($order.Problems.Count -eq 0) { $lines.Add("  $($g.Name): the fixed order holds; $($c.cycles) cycle(s), $($c.rows) calc rows: agree $($c.agree) differ $($c.differ) unchecked $($c.unchecked)") }
        $pin = Test-Pin $g $c
        foreach ($p in $pin) { $problems.Add($p); $lines.Add("  $p  FAIL") }
        if ($pin.Count -eq 0) { $lines.Add("  $($g.Name): the pin holds ($($g.Pin); agree floor $($g.AgreeFloor))") }
    }
    return @{ Problems = $problems; Lines = $lines }
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_recalc_golden_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        # Both fakes answer `calc <fixture>` with the golden of the fixture's
        # base name, and `calc <fixture> --counts` with the agreement floor.
        $map = ($goldens | ForEach-Object { "'" + [System.IO.Path]::GetFileNameWithoutExtension($_.Fixture) + "' { '" + ((Join-Path $repoRoot $_.Golden) -replace "'", "''") + "' }" }) -join '; '
        $agreeMap = ($agreements | ForEach-Object { "'" + $_.Name + "' { " + $_.Agree + " }" }) -join '; '
        $common = "param([Parameter(ValueFromRemainingArguments=`$true)][string[]]`$a)`n" +
                  "`$base = [System.IO.Path]::GetFileNameWithoutExtension(`$a[1])`n" +
                  "if (`$a -contains '--counts') { `$agree = switch (`$base) { $agreeMap; default { 0 } }; `$differ = @@D@@; Write-Output `"calc: formulas `$agree computed `$agree not-computed 0 cycles 0 agree `$agree differ `$differ unchecked 0 read 0.0 ms calc 0.0 ms`"; exit 0 }`n" +
                  "`$which = switch (`$base) { $map; default { exit 3 } }`n" +
                  "`$lines = [System.IO.File]::ReadAllText(`$which) -replace `"`r`n`", `"`n`" -split `"`n`"`n"
        $fake = Join-Path $tmp 'fake.ps1'
        [System.IO.File]::WriteAllText($fake, ($common -replace "@@D@@", "0") + "`$lines | ForEach-Object { Write-Output `$_ }`nexit 0`n")
        # The mutant changes one value in the first calc row it prints, and
        # reports one differ in its counts.
        $mut = Join-Path $tmp 'mutant.ps1'
        $flip = "for (`$i = 0; `$i -lt `$lines.Count; `$i++) { if (`$lines[`$i] -like '(calc *') { `$lines[`$i] = `$lines[`$i] -replace ' (agree|differ|unchecked)\)`$', ' X `$1)'; break } }`n"
        [System.IO.File]::WriteAllText($mut, ($common -replace "@@D@@", "1") + $flip + "`$lines | ForEach-Object { Write-Output `$_ }`nexit 0`n")
        $fakeProblems = 0; $mutantProblems = 0; $named = $true
        foreach ($g in $goldens) {
            $r1 = Measure-Impl $fake $g
            $r2 = Measure-Impl $mut $g
            $fakeProblems += $r1.Count; $mutantProblems += $r2.Count
            $rows = Get-Lines ([System.IO.File]::ReadAllText((Join-Path $repoRoot $g.Golden)))
            $firstCalc = 0
            for ($i = 0; $i -lt $rows.Count; $i++) { if ($rows[$i] -like '(calc *') { $firstCalc = $i + 1; break } }
            if ($r2.Count -ne 1 -or -not ($r2[0] -like "*differs at line $firstCalc of*")) { $named = $false }
            Write-Output ("control, {0}: the fake has {1} problem(s); the mutant {2}: {3}" -f $g.Name, $r1.Count, $r2.Count, ($r2 -join '; '))
        }
        $fakeAgree = 0; $mutantAgree = 0
        if ($studyPresent) {
            foreach ($a in $agreements) {
                $m1 = Measure-Agreement $fake $a
                $m2 = Measure-Agreement $mut $a
                $fakeAgree += $m1.Problems.Count; $mutantAgree += $m2.Problems.Count
                Write-Output ("control, {0}: the fake's agreement has {1} problem(s); the mutant's {2}" -f $a.Name, $m1.Problems.Count, $m2.Problems.Count)
            }
        } else {
            $mutantAgree = $agreements.Count
            Write-Output 'control: the study folder is not in this tree; the agreement cases are skipped'
        }
        # The pins: the saved golden as it is; a copy with one agree turned
        # to differ; a copy with a cycle row appended after the calc rows.
        $rows = Get-Lines ([System.IO.File]::ReadAllText((Join-Path $repoRoot $goldens[0].Golden)))
        $o0 = Test-Order $rows
        $p0 = Test-Pin $goldens[0] $o0.Counts
        $flipped = New-Object System.Collections.Generic.List[string]
        $done = $false
        foreach ($l in $rows) { if (-not $done -and $l -like '* agree)') { $flipped.Add(($l -replace ' agree\)$', ' differ)')); $done = $true } else { $flipped.Add($l) } }
        $o1 = Test-Order $flipped
        $p1 = Test-Pin $goldens[0] $o1.Counts
        $late = New-Object System.Collections.Generic.List[string]
        foreach ($l in $rows) { $late.Add($l) }
        $late.Add('(cycle "Output!B3" "Output!C2")')
        $o2 = Test-Order $late
        Write-Output ("control: the saved golden's order {0} and pin {1} problem(s); one agree turned to differ: order {2}, pin {3}: {4}; a cycle row after the calc rows: order {5}: {6}" -f $o0.Problems.Count, $p0.Count, $o1.Problems.Count, $p1.Count, ($p1 -join '; '), $o2.Problems.Count, ($o2.Problems -join '; '))
        if ($fakeProblems -eq 0 -and $named -and $mutantProblems -eq $goldens.Count -and $fakeAgree -eq 0 -and $mutantAgree -eq $agreements.Count -and $o0.Problems.Count -eq 0 -and $p0.Count -eq 0 -and $o1.Problems.Count -eq 0 -and $p1.Count -ge 1 -and $o2.Problems.Count -ge 1) {
            Write-Output 'OK: the check passes every golden and every agreement, fails the mutant on each naming its line, fails a differ where every row must agree, and catches a cycle row out of order'
            exit 0
        }
        Write-Output 'FAIL: the control did not behave as the header says'
        exit 1
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}

Write-Output '=== THE RECALC GOLDENS: REPRODUCED WHOLE, NEVER SHORTER, IN THE FIXED ORDER, EVERY SAVED VALUE AGREEING (KERNEL.7, oracle 12) ==='
$failed = New-Object System.Collections.Generic.List[string]
$tg = Test-Goldens
$tg.Lines | ForEach-Object { Write-Output $_ }
foreach ($p in $tg.Problems) { $failed.Add($p) }

# --- the door ---------------------------------------------------------------
$door = $Impl
if ($door -eq '') {
    foreach ($candidate in @('target/debug/frazaro.exe', 'target/release/frazaro.exe', 'target/debug/frazaro', 'target/release/frazaro')) {
        $p = Join-Path $repoRoot $candidate
        if (Test-Path $p) { $door = $p; break }
    }
}
if ($door -eq '') {
    Write-Output '  SKIPPED: no door built (target/debug/frazaro.exe) and no -Impl given; the floors, the order and the pins were checked'
} else {
    if (-not [System.IO.Path]::IsPathRooted($door)) { $door = Join-Path (Get-Location) $door }
    Write-Output "  door: $door"
    foreach ($g in $goldens) {
        $r = Measure-Impl $door $g
        if ($r.Count -eq 0) { Write-Output "  $($g.Name): the door reproduces the golden whole" }
        foreach ($p in $r) { $failed.Add($p); Write-Output "  $p  FAIL" }
    }
    if ($studyPresent) {
        foreach ($a in $agreements) {
            $m = Measure-Agreement $door $a
            if ($m.Problems.Count -eq 0) { Write-Output "  $($a.Name): agrees ($($m.Line -replace '^calc: ', '' -replace ' read .*$', ''))" }
            foreach ($p in $m.Problems) { $failed.Add($p); Write-Output "  $p  FAIL" }
        }
    } else {
        Write-Output '  SKIPPED: the study workbooks (scripts/study/, KERNEL.3) are not in this tree; their agreement is held where they are'
    }
}

if ($failed.Count -gt 0) {
    Write-Output ''
    Write-Output "FAIL: $($failed.Count) problem(s)"
    $failed | ForEach-Object { Write-Output "  - $_" }
    exit 1
}
Write-Output 'OK: every recalc golden at or above its floor, in the fixed order, its pin holding, reproduced whole by the door where one is built, and the study''s workbooks agreeing'
exit 0
