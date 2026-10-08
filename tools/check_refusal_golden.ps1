<#
check_refusal_golden.ps1 - the refusal golden agrees with its fixture, every
case refuses, and the distinct refusal ids it reaches never go down.

WHY: PORT.6 (slice 6f) holds the core's English port to every refusal id the
sentence engine raises on the path, in the situation that raises it, with its
text: scripts/refusals.txt is the fixture (cases under "=== program <name>"
and "=== phrasebook <name>" lines, each named for the id it means to reach),
scripts/refusals_golden.txt the reference's reading of each
(VlaWriteRefusalGolden, VLA_Tests.bas: one REFUSED<TAB><line><TAB><id><TAB>
<text> record per case, or TRANSLATED or LOADED where it does not refuse),
and the core's own test (core/src/english/refusals.rs) compares its reading
with the golden whole. That test proves agreement; nothing proves coverage:
a case whose sentence drifts into translating, a golden regenerated against
a fixture that lost a case, or an id the fixture no longer reaches would all
leave that test green. This check holds the shape and the coverage,
host-free, on every push:

  1. the golden's headers are the fixture's, in order, one record each;
  2. every record is REFUSED - a case that translates or loads reaches no
     refusal, and is mended or removed, never kept;
  3. every id in the golden is a (message <id> ...) of core/data/messages.vla
     or vla-lang/data/messages.vla, the catalogue's two halves since PORT.12;
  4. the number of distinct ids is at or above the floor below, the house
     style's hardcoded, reviewable number, raised by hand as cases are added.

The ids the golden cannot reach are named, with their reasons, in the
treaty's amendment of 2026-10-02 for this slice (conformance/README.md): the
door's file and gate refusals, G-RENDER's three, the two overwrite refusals
of file-writing commands, one that reaches a caller only wrapped, and two
defensive arms registration keeps unreachable.

-Control proves the check on scratch copies of the golden: the real pair
must pass at the real floor; a golden with one id misspelled, one with a
record turned to TRANSLATED, and one with its last case dropped must each
fail.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; exit 0 clean, exit 1 with every problem named.

Usage:  powershell -File tools\check_refusal_golden.ps1
        powershell -File tools\check_refusal_golden.ps1 -Control
#>
param(
    [switch]$Control
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# --- the floor: distinct refusal ids the golden reaches, never lowered -------
# 2026-10-02, PORT.6 slice 6f: 113 of the 125 ids the English modules raise,
# from 115 cases, on the golden the core predicted ahead of the reference's
# first run.
$floor = 113

$fixturePath  = Join-Path $repoRoot 'scripts/refusals.txt'
$goldenPath   = Join-Path $repoRoot 'scripts/refusals_golden.txt'
# The catalogue's two halves since PORT.12: the core's and the language's, one VBA source.
$messagesPaths = @((Join-Path $repoRoot 'core/data/messages.vla'), (Join-Path $repoRoot 'vla-lang/data/messages.vla'))

function Get-Lines([string]$path) {
    $text = [System.IO.File]::ReadAllText($path)
    $text = $text -replace "`r`n", "`n" -replace "`r", "`n"
    return @($text -split "`n")
}

# The fixture's headers, in order.
function Read-FixtureHeaders([string]$path) {
    $headers = @()
    foreach ($ln in (Get-Lines $path)) {
        if ($ln.StartsWith('=== ')) { $headers += $ln }
    }
    return $headers
}

# The golden's cases: the header, the record line, the record count, and the
# id when the record is REFUSED. A refusal's text may run on over more lines
# (the catalogue has templates with line breaks); those are neither headers
# nor records.
function Read-GoldenCases([string]$path) {
    $cases = @()
    $current = $null
    foreach ($ln in (Get-Lines $path)) {
        if ($ln.StartsWith('=== ')) {
            if ($null -ne $current) { $cases += $current }
            $current = @{ Header = $ln; Record = ''; Id = ''; Records = 0 }
        } elseif ($null -ne $current) {
            if ($ln -match '^(REFUSED|TRANSLATED|LOADED)(\t|$)') {
                $current.Records = $current.Records + 1
                if ($current.Record -eq '') {
                    $current.Record = $ln
                    if ($ln.StartsWith('REFUSED')) {
                        $f = $ln -split "`t"
                        if ($f.Count -ge 4) { $current.Id = $f[2] }
                    }
                }
            }
        }
    }
    if ($null -ne $current) { $cases += $current }
    return $cases
}

function Read-MessageIds([string[]]$paths) {
    $ids = @{}
    foreach ($path in $paths) {
        foreach ($ln in (Get-Lines $path)) {
            if ($ln -match '^\(message ([a-z0-9-]+) ') { $ids[$matches[1]] = $true }
        }
    }
    return $ids
}

# Every problem in a fixture/golden pair, with the distinct ids it reaches.
function Test-Pair([string]$fixture, [string]$golden, [string[]]$messages, [int]$minIds) {
    $problems = @()
    $headers = @(Read-FixtureHeaders $fixture)
    $cases = @(Read-GoldenCases $golden)
    $ids = Read-MessageIds $messages
    if ($headers.Count -ne $cases.Count) {
        $problems += ("the fixture has {0} cases and the golden {1}" -f $headers.Count, $cases.Count)
    }
    $n = [Math]::Min($headers.Count, $cases.Count)
    for ($i = 0; $i -lt $n; $i++) {
        if ($headers[$i] -ne $cases[$i].Header) {
            $problems += ("case {0}: the fixture says '{1}', the golden '{2}'" -f ($i + 1), $headers[$i], $cases[$i].Header)
        }
    }
    $distinct = @{}
    foreach ($c in $cases) {
        if ($c.Records -ne 1) {
            $problems += ("'{0}': {1} record(s), want exactly one" -f $c.Header, $c.Records)
            continue
        }
        if ($c.Id -eq '') {
            $problems += ("'{0}': {1} - a case the reference does not refuse reaches no refusal; mend it or remove it" -f $c.Header, $c.Record)
            continue
        }
        if (-not $ids.ContainsKey($c.Id)) {
            $problems += ("'{0}': the id '{1}' is not in core/data/messages.vla or vla-lang/data/messages.vla" -f $c.Header, $c.Id)
        }
        $distinct[$c.Id] = $true
    }
    if ($distinct.Count -lt $minIds) {
        $problems += ("{0} distinct refusal ids, below the floor of {1}" -f $distinct.Count, $minIds)
    }
    return @{ Problems = $problems; Distinct = $distinct.Count; Cases = $cases.Count }
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_rgolden_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $real = Test-Pair $fixturePath $goldenPath $messagesPaths $floor
        $lines = Get-Lines $goldenPath
        $firstRecord = -1
        $lastHeader = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($firstRecord -lt 0 -and $lines[$i].StartsWith('REFUSED')) { $firstRecord = $i }
            if ($lines[$i].StartsWith('=== ')) { $lastHeader = $i }
        }
        # (a) one id misspelled
        $a = @($lines | ForEach-Object { $_ })
        $f = $a[$firstRecord] -split "`t"
        $f[2] = $f[2] + '-misspelt'
        $a[$firstRecord] = $f -join "`t"
        $pa = Join-Path $tmp 'a.txt'
        [System.IO.File]::WriteAllText($pa, ($a -join "`n"), $utf8)
        # (b) one record that does not refuse
        $b = @($lines | ForEach-Object { $_ })
        $b[$firstRecord] = 'TRANSLATED'
        $pb = Join-Path $tmp 'b.txt'
        [System.IO.File]::WriteAllText($pb, ($b -join "`n"), $utf8)
        # (c) the last case dropped
        $c = @($lines[0..($lastHeader - 1)])
        $pc = Join-Path $tmp 'c.txt'
        [System.IO.File]::WriteAllText($pc, ($c -join "`n"), $utf8)
        $ra = Test-Pair $fixturePath $pa $messagesPaths $real.Distinct
        $rb = Test-Pair $fixturePath $pb $messagesPaths $real.Distinct
        $rc = Test-Pair $fixturePath $pc $messagesPaths $real.Distinct
        Write-Output ("control: the real pair has {0} problem(s) at {1} ids; the misspelt id {2}, the translated case {3}, the dropped case {4}" -f $real.Problems.Count, $real.Distinct, $ra.Problems.Count, $rb.Problems.Count, $rc.Problems.Count)
        if ($real.Problems.Count -eq 0 -and $ra.Problems.Count -ge 1 -and $rb.Problems.Count -ge 1 -and $rc.Problems.Count -ge 1) {
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

$r = Test-Pair $fixturePath $goldenPath $messagesPaths $floor
if ($r.Problems.Count -eq 0) {
    if ($r.Distinct -gt $floor) {
        Write-Output ("OK: {0} cases, every one refused, {1} distinct refusal ids, above the floor of {2} - raise the floor in this file" -f $r.Cases, $r.Distinct, $floor)
    } else {
        Write-Output ("OK: {0} cases, every one refused, {1} distinct refusal ids (floor {2})" -f $r.Cases, $r.Distinct, $floor)
    }
    exit 0
}
Write-Output ("FAIL: the refusal golden has {0} problem(s):" -f $r.Problems.Count)
foreach ($p in $r.Problems) { Write-Output ('  ' + $p) }
exit 1
