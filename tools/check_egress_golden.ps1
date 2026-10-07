<#
check_egress_golden.ps1 - the egress golden agrees with its fixture, every
record is REFUSED with a listed name or WRITTEN, the VBA's list and the
core's are one, every listed name is reached, and the list never shrinks.

WHY: SEC.15 (the formula sink) refuses a formula that reaches outside the
workbook on its own by name, from one list per implementation: the
`listed = "..."` line of VlaFormulaEgress (src/VLA_Runtime.bas, above the
inject boundary, so the compiled program carries it) and the EGRESS_NAMES
array of core/src/egress.rs. The two are held to one golden in the token
golden's shape: scripts/egress.txt is the fixture (cases under "=== <name>"
lines, each one text as the sink receives it), and scripts/egress_golden.txt
is the reference's reading of each (VlaWriteEgressGolden, VLA_Tests.bas):
one record, REFUSED<TAB><the function named, or DDE for a DDE link> or
WRITTEN. The core's own test compares its reading with the golden whole.
That test proves agreement; nothing proves coverage, and nothing but this
script reads the two lists side by side: a name added to one list and not
the other, a name on the list that no case reaches, a golden regenerated
against a fixture that lost a case, or a list quietly shortened would all
leave the core's test green. This check holds the shape, the twin and the
coverage, host-free, on every push:

  1. the golden's headers are the fixture's, in order, and every fixture
     case has at least one line of text;
  2. every golden case is exactly one record, WRITTEN or REFUSED<TAB><name>,
     the name one of the list's upper-cased, or DDE;
  3. the VBA's list and the core's list hold the same names (read from the
     source, never retyped here);
  4. every name of the list is reached by a REFUSED record, and DDE by one
     (the link shape), so a name is never on the list unproven;
  5. the list holds at least the floor below, the house style's hardcoded,
     reviewable number.

-Control proves the check on scratch copies: the real files must pass at
the real floor; a golden with one record misspelled, one with its last case
dropped, one with every WEBSERVICE record turned WRITTEN (a name no case
reaches), and a VBA source with one name taken off its list (the twin
broken) must each fail.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; exit 0 clean, exit 1 with every problem named.

Usage:  powershell -File tools\check_egress_golden.ps1
        powershell -File tools\check_egress_golden.ps1 -Control
#>
param(
    [switch]$Control
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# --- the floor: names on the list, never lowered --------------------------
# 2026-10-05, SEC.15: 23, every one reached by a case of the 55, the golden
# written by hand ahead of the reference's first run.
$floor = 23

$fixturePath = Join-Path $repoRoot 'scripts/egress.txt'
$goldenPath  = Join-Path $repoRoot 'scripts/egress_golden.txt'
$vbaPath     = Join-Path $repoRoot 'src/VLA_Runtime.bas'
$corePath    = Join-Path $repoRoot 'core/src/egress.rs'

function Get-Lines([string]$path) {
    $text = [System.IO.File]::ReadAllText($path)
    $text = $text -replace "`r`n", "`n" -replace "`r", "`n"
    return @($text -split "`n")
}

# The fixture's cases: the header and whether any text follows it.
function Read-FixtureCases([string]$path) {
    $cases = @()
    $current = $null
    foreach ($ln in (Get-Lines $path)) {
        if ($ln.StartsWith('=== ')) {
            if ($null -ne $current) { $cases += $current }
            $current = @{ Header = $ln; HasText = $false }
        } elseif ($null -ne $current -and $ln -ne '') {
            $current.HasText = $true
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

# The VBA's list: the words of the `listed = " ... "` line inside
# VlaFormulaEgress, folded.
function Read-VbaList([string]$path) {
    $names = @()
    $inProc = $false
    foreach ($ln in (Get-Lines $path)) {
        if ($ln -match '^\s*Public Function VlaFormulaEgress\b') { $inProc = $true; continue }
        if ($inProc -and $ln -match '^\s*End Function') { break }
        if ($inProc -and $ln -match '^\s*listed\s*=\s*"([^"]*)"') {
            foreach ($w in ($Matches[1] -split '\s+')) {
                if ($w.Length -gt 0) { $names += $w.ToLowerInvariant() }
            }
        }
    }
    return $names
}

# The core's list: the quoted strings of the EGRESS_NAMES array, folded.
function Read-CoreList([string]$path) {
    $text = [System.IO.File]::ReadAllText($path)
    $m = [regex]::Match($text, 'EGRESS_NAMES:\s*&\[&str\]\s*=\s*&\[(.*?)\];', 'Singleline')
    if (-not $m.Success) { return @() }
    $names = @()
    foreach ($q in [regex]::Matches($m.Groups[1].Value, '"([^"]*)"')) {
        $names += $q.Groups[1].Value.ToLowerInvariant()
    }
    return $names
}

# Every problem in a fixture/golden/VBA/core set, with the names reached.
function Test-Egress([string]$fixture, [string]$golden, [string]$vba, [string]$core, [int]$minNames) {
    $problems = @()
    $fcases = @(Read-FixtureCases $fixture)
    $gcases = @(Read-GoldenCases $golden)
    $vbaList = @(Read-VbaList $vba)
    $coreList = @(Read-CoreList $core)
    if ($vbaList.Count -eq 0) { $problems += "no listed = `"...`" line inside VlaFormulaEgress in $vba" }
    if ($coreList.Count -eq 0) { $problems += "no EGRESS_NAMES array in $core" }
    $onlyVba = @($vbaList | Where-Object { $coreList -notcontains $_ })
    $onlyCore = @($coreList | Where-Object { $vbaList -notcontains $_ })
    foreach ($n in $onlyVba) { $problems += "'$n' is on the VBA's list and not the core's" }
    foreach ($n in $onlyCore) { $problems += "'$n' is on the core's list and not the VBA's" }
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
        if (-not $c.HasText) { $problems += ("'{0}': no text after its header" -f $c.Header) }
    }
    $upper = @{}
    foreach ($w in $vbaList) { $upper[$w.ToUpperInvariant()] = $true }
    $reached = @{}
    foreach ($c in $gcases) {
        $recs = @($c.Records)
        if ($recs.Count -ne 1) {
            $problems += ("'{0}': {1} record(s), want exactly one" -f $c.Header, $recs.Count)
            continue
        }
        $rec = $recs[0]
        if ($rec -eq 'WRITTEN') { continue }
        $f = $rec -split "`t"
        if ($f.Count -ne 2 -or $f[0] -ne 'REFUSED') {
            $problems += ("'{0}': a record that is neither WRITTEN nor REFUSED<TAB><name>: '{1}'" -f $c.Header, $rec)
            continue
        }
        if ($f[1] -ne 'DDE' -and -not $upper.ContainsKey($f[1])) {
            $problems += ("'{0}': REFUSED names '{1}', which is not on the list" -f $c.Header, $f[1])
            continue
        }
        $reached[$f[1]] = $true
    }
    foreach ($w in $vbaList) {
        if (-not $reached.ContainsKey($w.ToUpperInvariant())) {
            $problems += "'$w' is on the list and no case of the golden reaches it"
        }
    }
    if (-not $reached.ContainsKey('DDE')) { $problems += 'no case of the golden reaches a DDE link' }
    if ($vbaList.Count -lt $minNames) {
        $problems += ("{0} names on the list, below the floor of {1}" -f $vbaList.Count, $minNames)
    }
    return @{ Problems = $problems; Names = $vbaList.Count; Cases = $gcases.Count; Reached = $reached.Count }
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_egress_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $real = Test-Egress $fixturePath $goldenPath $vbaPath $corePath $floor
        $lines = Get-Lines $goldenPath
        $firstRefused = -1
        $lastHeader = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($firstRefused -lt 0 -and $lines[$i].StartsWith("REFUSED`t")) { $firstRefused = $i }
            if ($lines[$i].StartsWith('=== ')) { $lastHeader = $i }
        }
        # (a) one record misspelled
        $a = @($lines | ForEach-Object { $_ })
        $a[$firstRefused] = 'REFUSE' + $a[$firstRefused].Substring(7)
        $pa = Join-Path $tmp 'a.txt'
        [System.IO.File]::WriteAllText($pa, ($a -join "`n"), $utf8)
        # (b) the last case dropped
        $b = @($lines[0..($lastHeader - 1)])
        $pb = Join-Path $tmp 'b.txt'
        [System.IO.File]::WriteAllText($pb, ($b -join "`n"), $utf8)
        # (c) every WEBSERVICE record turned WRITTEN: a listed name no case reaches
        $c = @($lines | ForEach-Object { if ($_ -eq "REFUSED`tWEBSERVICE") { 'WRITTEN' } else { $_ } })
        $pc = Join-Path $tmp 'c.txt'
        [System.IO.File]::WriteAllText($pc, ($c -join "`n"), $utf8)
        # (d) the VBA's list with one name taken off: the twin broken
        $vbaText = [System.IO.File]::ReadAllText($vbaPath)
        $count = [regex]::Matches($vbaText, ' filterxml ').Count
        if ($count -ne 1) { Write-Output "FAIL: the control expected ' filterxml ' once in the VBA's list, found $count"; exit 1 }
        $pd = Join-Path $tmp 'd.bas'
        [System.IO.File]::WriteAllText($pd, $vbaText.Replace(' filterxml ', ' '), $utf8)
        $ra = Test-Egress $fixturePath $pa $vbaPath $corePath $real.Names
        $rb = Test-Egress $fixturePath $pb $vbaPath $corePath $real.Names
        $rc = Test-Egress $fixturePath $pc $vbaPath $corePath $real.Names
        $rd = Test-Egress $fixturePath $goldenPath $pd $corePath $real.Names
        Write-Output ("control: the real set has {0} problem(s) at {1} names; the misspelt record {2}, the dropped case {3}, the unreached name {4}, the broken twin {5}" -f $real.Problems.Count, $real.Names, $ra.Problems.Count, $rb.Problems.Count, $rc.Problems.Count, $rd.Problems.Count)
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

$r = Test-Egress $fixturePath $goldenPath $vbaPath $corePath $floor
if ($r.Problems.Count -eq 0) {
    if ($r.Names -gt $floor) {
        Write-Output ("OK: {0} cases, {1} names on both lists, every one reached, above the floor of {2} - raise the floor in this file" -f $r.Cases, $r.Names, $floor)
    } else {
        Write-Output ("OK: {0} cases, {1} names on both lists, every one reached and DDE too (floor {2})" -f $r.Cases, $r.Names, $floor)
    }
    exit 0
}
Write-Output ("FAIL: the egress golden has {0} problem(s):" -f $r.Problems.Count)
foreach ($p in $r.Problems) { Write-Output ('  ' + $p) }
exit 1
