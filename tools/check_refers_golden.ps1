<#
check_refers_golden.ps1 - the refers golden agrees with its fixture, every
case names its cell and ends in its R1C1 record, every record is one of
the eleven kinds, and the distinct kinds reached never go down.

WHY: AXM.7 (the formula-reference reader, src/VLA_Refers.bas) is held to
the core's port (PORT.8, slice 8b) through a golden in the token golden's
shape: scripts/refers.txt is the fixture (cases under "=== <name>" lines,
each the cell holding the formula, then the formula), and
scripts/refers_golden.txt is the reference's reading of each
(VlaWriteRefersGolden, VLA_Tests.bas): one <kind><TAB><the token as
written><TAB><the second field of its refers row> record per reference in
formula order, or NONE, then one R1C1<TAB><the formula in R1C1 relative to
the cell> record. The core's own test compares its reading with the golden
whole. That test proves agreement; nothing proves coverage, and until 8b
lands nothing reads the pair at all: a case that drifted into another
kind, a golden regenerated against a fixture that lost a case, or a kind
the fixture no longer reaches would all leave the core's test green. This
check holds the shape and the coverage, host-free, on every push:

  1. the golden's headers are the fixture's, in order, and every fixture
     case is a cell line (Model!B3 or 'Q1 Data'!C5) then a formula that
     begins with =;
  2. every golden case is one or more kind records, or exactly one NONE,
     then exactly one R1C1 record, last; a BAD HOME record (the writer's
     word for a cell line it could not read) fails;
  3. every kind is one of the eleven: cell, range, column, row, name,
     structured, external, 3d, spill, unreadable, broken;
  4. the number of distinct kinds is at or above the floor below, the
     house style's hardcoded, reviewable number.

-Control proves the check on scratch copies of the golden: the real pair
must pass at the real floor; a golden with one kind misspelled, one with a
case's R1C1 record dropped, one with its last case dropped, and one with
every spill record turned into a cell record (one kind fewer) must each
fail.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; exit 0 clean, exit 1 with every problem named.

Usage:  powershell -File tools\check_refers_golden.ps1
        powershell -File tools\check_refers_golden.ps1 -Control
#>
param(
    [switch]$Control
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# --- the floor: distinct kinds the golden reaches, never lowered --------
# 2026-10-04, AXM.7: all eleven, from 85 cases, on the golden written by
# hand ahead of the reference's first run.
$floor = 11

$kinds = @('cell', 'range', 'column', 'row', 'name', 'structured', 'external', '3d', 'spill', 'unreadable', 'broken')
$fixturePath = Join-Path $repoRoot 'scripts/refers.txt'
$goldenPath  = Join-Path $repoRoot 'scripts/refers_golden.txt'
# Sheet!A1 or 'Q1 Data'!A1: an unquoted sheet name, or a quoted one with
# its apostrophes doubled, then a cell with optional $ marks.
$homePattern = '^(''(?:[^'']|'''')+''|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$'

function Get-Lines([string]$path) {
    $text = [System.IO.File]::ReadAllText($path)
    $text = $text -replace "`r`n", "`n" -replace "`r", "`n"
    return @($text -split "`n")
}

# The fixture's cases: the header, the cell line, the formula's first line.
function Read-FixtureCases([string]$path) {
    $cases = @()
    $current = $null
    foreach ($ln in (Get-Lines $path)) {
        if ($ln.StartsWith('=== ')) {
            if ($null -ne $current) { $cases += $current }
            $current = @{ Header = $ln; Home = $null; Formula = $null }
        } elseif ($null -ne $current -and $ln -ne '') {
            if ($null -eq $current.Home) { $current.Home = $ln }
            elseif ($null -eq $current.Formula) { $current.Formula = $ln }
        }
    }
    if ($null -ne $current) { $cases += $current }
    return $cases
}

# The golden's cases: the header and its record lines, blank lines dropped.
function Read-GoldenCases([string]$path) {
    $cases = @()
    $current = $null
    foreach ($ln in (Get-Lines $path)) {
        if ($ln.StartsWith('=== ')) {
            if ($null -ne $current) { $cases += $current }
            $current = @{ Header = $ln; Records = @() }
        } elseif ($null -ne $current -and $ln -ne '') {
            $current.Records += $ln
        }
    }
    if ($null -ne $current) { $cases += $current }
    return $cases
}

# Every problem in a fixture/golden pair, with the distinct kinds it reaches.
function Test-Pair([string]$fixture, [string]$golden, [int]$minKinds) {
    $problems = @()
    $fcases = @(Read-FixtureCases $fixture)
    $gcases = @(Read-GoldenCases $golden)
    if ($fcases.Count -ne $gcases.Count) {
        $problems += ("the fixture has {0} cases and the golden {1}" -f $fcases.Count, $gcases.Count)
    }
    $n = [Math]::Min($fcases.Count, $gcases.Count)
    for ($i = 0; $i -lt $n; $i++) {
        if ($fcases[$i].Header -ne $gcases[$i].Header) {
            $problems += ("case {0}: the fixture says '{1}', the golden '{2}'" -f ($i + 1), $fcases[$i].Header, $gcases[$i].Header)
        }
    }
    foreach ($c in $fcases) {
        if ($null -eq $c.Home -or $c.Home -notmatch $homePattern) {
            $problems += ("'{0}': its first line is not a cell (Model!B3 or 'Q1 Data'!C5): '{1}'" -f $c.Header, $c.Home)
        }
        if ($null -eq $c.Formula -or -not $c.Formula.StartsWith('=')) {
            $problems += ("'{0}': no formula beginning with = after its cell line" -f $c.Header)
        }
    }
    $distinct = @{}
    foreach ($c in $gcases) {
        $recs = @($c.Records)
        if ($recs.Count -lt 2) {
            $problems += ("'{0}': {1} record(s), want at least a reference or NONE and then R1C1" -f $c.Header, $recs.Count)
            continue
        }
        $last = $recs[$recs.Count - 1]
        if (-not $last.StartsWith("R1C1`t")) {
            $problems += ("'{0}': the last record is not R1C1: '{1}'" -f $c.Header, $last)
        }
        $body = @($recs[0..($recs.Count - 2)])
        $r1c1Inside = @($body | Where-Object { $_.StartsWith("R1C1`t") }).Count
        if ($r1c1Inside -gt 0) {
            $problems += ("'{0}': an R1C1 record before the last" -f $c.Header)
        }
        if ($body.Count -eq 1 -and $body[0] -eq 'NONE') { continue }
        foreach ($rec in $body) {
            if ($rec -eq 'NONE') {
                $problems += ("'{0}': NONE beside other records" -f $c.Header)
                continue
            }
            $f = $rec -split "`t"
            if ($f.Count -ne 3) {
                $problems += ("'{0}': a record with {1} field(s), want kind, text as written, spelled: '{2}'" -f $c.Header, $f.Count, $rec)
                continue
            }
            if ($kinds -notcontains $f[0]) {
                $problems += ("'{0}': the kind '{1}' is not one of the eleven" -f $c.Header, $f[0])
                continue
            }
            if ($f[1] -eq '' -or $f[2] -eq '') {
                $problems += ("'{0}': an empty field in '{1}'" -f $c.Header, $rec)
                continue
            }
            $distinct[$f[0]] = $true
        }
    }
    if ($distinct.Count -lt $minKinds) {
        $problems += ("{0} distinct kinds, below the floor of {1}" -f $distinct.Count, $minKinds)
    }
    return @{ Problems = $problems; Distinct = $distinct.Count; Cases = $gcases.Count }
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_rfgolden_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $real = Test-Pair $fixturePath $goldenPath $floor
        $lines = Get-Lines $goldenPath
        $firstRecord = -1
        $firstR1C1 = -1
        $lastHeader = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($firstRecord -lt 0 -and $lines[$i] -match "^[a-z0-9]+`t") { $firstRecord = $i }
            if ($firstR1C1 -lt 0 -and $lines[$i].StartsWith("R1C1`t")) { $firstR1C1 = $i }
            if ($lines[$i].StartsWith('=== ')) { $lastHeader = $i }
        }
        # (a) one kind misspelled
        $a = @($lines | ForEach-Object { $_ })
        $f = $a[$firstRecord] -split "`t"
        $f[0] = $f[0] + '-misspelt'
        $a[$firstRecord] = $f -join "`t"
        $pa = Join-Path $tmp 'a.txt'
        [System.IO.File]::WriteAllText($pa, ($a -join "`n"), $utf8)
        # (b) one case's R1C1 record dropped
        $b = @($lines[0..($firstR1C1 - 1)]) + @($lines[($firstR1C1 + 1)..($lines.Count - 1)])
        $pb = Join-Path $tmp 'b.txt'
        [System.IO.File]::WriteAllText($pb, ($b -join "`n"), $utf8)
        # (c) the last case dropped
        $c = @($lines[0..($lastHeader - 1)])
        $pc = Join-Path $tmp 'c.txt'
        [System.IO.File]::WriteAllText($pc, ($c -join "`n"), $utf8)
        # (d) every spill record turned into a cell record: one kind fewer
        $d = @($lines | ForEach-Object { $_ -replace "^spill`t", "cell`t" })
        $pd = Join-Path $tmp 'd.txt'
        [System.IO.File]::WriteAllText($pd, ($d -join "`n"), $utf8)
        $ra = Test-Pair $fixturePath $pa $real.Distinct
        $rb = Test-Pair $fixturePath $pb $real.Distinct
        $rc = Test-Pair $fixturePath $pc $real.Distinct
        $rd = Test-Pair $fixturePath $pd $real.Distinct
        Write-Output ("control: the real pair has {0} problem(s) at {1} kinds; the misspelt kind {2}, the dropped R1C1 {3}, the dropped case {4}, the lost kind {5}" -f $real.Problems.Count, $real.Distinct, $ra.Problems.Count, $rb.Problems.Count, $rc.Problems.Count, $rd.Problems.Count)
        if ($real.Problems.Count -eq 0 -and $ra.Problems.Count -ge 1 -and $rb.Problems.Count -ge 1 -and $rc.Problems.Count -ge 1 -and $rd.Problems.Count -ge 1) {
            Write-Output 'OK: the check passes the golden and fails each mutant'
            exit 0
        }
        Write-Output 'FAIL: the control did not behave'
        foreach ($p in $real.Problems) { Write-Output ('  real: ' + $p) }
        exit 1
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}

$r = Test-Pair $fixturePath $goldenPath $floor
if ($r.Problems.Count -eq 0) {
    if ($r.Distinct -gt $floor) {
        Write-Output ("OK: {0} cases, {1} distinct kinds, above the floor of {2} - raise the floor in this file" -f $r.Cases, $r.Distinct, $floor)
    } else {
        Write-Output ("OK: {0} cases, every one a cell and a formula, {1} distinct kinds (floor {2})" -f $r.Cases, $r.Distinct, $floor)
    }
    exit 0
}
Write-Output ("FAIL: the refers golden has {0} problem(s):" -f $r.Problems.Count)
foreach ($p in $r.Problems) { Write-Output ('  ' + $p) }
exit 1
