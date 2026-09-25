<#
optimize3_search_twin.ps1 - the forty generated search problems of
VLA_Tests_Query.TestOptimizeSearch, generated again here by the same seeded
generator and answered by the same exhaustive walk, host-independent.

WHY. TestOptimizeSearch holds OPTIMIZE's search to one claim - its answer
is the first world, lexicographic with TRUE first, that an exhaustive walk
finds - over forty problems a seeded generator writes. The pins compare
the search with a walk written in the same module, so a defect shared by
the two could cancel out; and a floor pin says how many of the forty have
a world, which is a number that has to come from somewhere other than the
code it checks. This computes it here, from the generator's definition,
and -Control fails when the VBA pin's own number is not this one.

THE GENERATOR, as VLA_Tests_Query.RandomSearchProblem writes it (read that
procedure, not this summary, when they disagree): Park and Miller's
minimal standard generator in Doubles (seed x 16807 mod 2^31 - 1), seed
20260924; per problem 3 to 9 atoms; up to two counters, each atom a member
by a coin flip (one drawn if none was), at least drawn from 0 to the
members, at most drawn from the least to the members or, one time in
three, none; up to six clauses of one to three literals over distinct
atoms, each sign a coin flip.

House style: PowerShell 5.1, host-independent, no network, no Excel.
Usage:  powershell -File tools\optimize3_search_twin.ps1 [-Control]
Exit 0 when -Control holds (always 0 without it); 1 otherwise.
#>
param([switch]$Control)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

$script:seed = 20260924.0
function Get-LcgNext {
    $x = $script:seed * 16807.0
    $script:seed = $x - [math]::Floor($x / 2147483647.0) * 2147483647.0
    return $script:seed / 2147483647.0
}
function Get-LcgInt([int]$n) {
    $v = [int][math]::Floor((Get-LcgNext) * $n)
    if ($v -ge $n) { $v = $n - 1 }
    return $v
}

function New-Problem {
    $n = 3 + (Get-LcgInt 7)
    $counters = New-Object System.Collections.Generic.List[object]
    $nCtr = Get-LcgInt 3
    for ($g = 1; $g -le $nCtr; $g++) {
        $mem = New-Object System.Collections.Generic.List[int]
        for ($a = 1; $a -le $n; $a++) {
            if ((Get-LcgInt 2) -eq 0) { $mem.Add($a) }
        }
        if ($mem.Count -eq 0) { $mem.Add(1 + (Get-LcgInt $n)) }
        $m = $mem.Count
        $lo = Get-LcgInt ($m + 1)
        if ((Get-LcgInt 3) -eq 0) { $hi = -1 } else { $hi = $lo + (Get-LcgInt ($m - $lo + 1)) }
        $counters.Add([pscustomobject]@{ Members = $mem.ToArray(); Lo = $lo; Hi = $hi })
    }
    $clauses = New-Object System.Collections.Generic.List[object]
    $nCl = Get-LcgInt 7
    for ($c = 1; $c -le $nCl; $c++) {
        $nLit = 1 + (Get-LcgInt 3)
        $lits = New-Object System.Collections.Generic.List[int]
        for ($j = 1; $j -le $nLit; $j++) {
            $at = 1 + (Get-LcgInt $n)
            $dup = $false
            foreach ($x in $lits) { if ([math]::Abs($x) -eq $at) { $dup = $true } }
            if (-not $dup) {
                if ((Get-LcgInt 2) -eq 0) { $lits.Add($at) } else { $lits.Add(-$at) }
            }
        }
        $clauses.Add($lits.ToArray())
    }
    return [pscustomobject]@{ N = $n; Counters = $counters; Clauses = $clauses }
}

function Get-FirstWorld($p) {
    $n = $p.N
    $total = [int][math]::Pow(2, $n)
    $v = New-Object 'int[]' ($n + 1)
    for ($k = 0; $k -lt $total; $k++) {
        for ($a = 1; $a -le $n; $a++) {
            $bit = [int][math]::Pow(2, $n - $a)
            if (($k -band $bit) -eq 0) { $v[$a] = 1 } else { $v[$a] = -1 }
        }
        $ok = $true
        foreach ($cl in $p.Clauses) {
            $sat = $false
            foreach ($lit in $cl) {
                if ($lit -gt 0) { if ($v[$lit] -eq 1) { $sat = $true } }
                else { if ($v[-$lit] -eq -1) { $sat = $true } }
            }
            if (-not $sat) { $ok = $false; break }
        }
        if ($ok) {
            foreach ($ct in $p.Counters) {
                $t = 0
                foreach ($m in $ct.Members) { if ($v[$m] -eq 1) { $t++ } }
                if ($t -lt $ct.Lo) { $ok = $false; break }
                if ($ct.Hi -ne -1 -and $t -gt $ct.Hi) { $ok = $false; break }
            }
        }
        if ($ok) {
            $s = ''
            for ($a = 1; $a -le $n; $a++) { if ($v[$a] -eq 1) { $s += 'T' } else { $s += 'F' } }
            return $s
        }
    }
    return $null
}

$found = 0
$none = 0
Write-Output '=== OPTIMIZE.3 SEARCH TWIN: the forty generated problems, walked ==='
for ($p = 1; $p -le 40; $p++) {
    $prob = New-Problem
    $w = Get-FirstWorld $prob
    $ctrText = ($prob.Counters | ForEach-Object { ('{0}:{1}:{2}' -f ($_.Members -join ' '), $_.Lo, $_.Hi) }) -join ';'
    $clText = ($prob.Clauses | ForEach-Object { $_ -join ' ' }) -join ';'
    if ($null -eq $w) {
        $none++
        Write-Output ('  {0,2}  n={1}  none        clauses [{2}]  counters [{3}]' -f $p, $prob.N, $clText, $ctrText)
    } else {
        $found++
        Write-Output ('  {0,2}  n={1}  {2,-10}  clauses [{3}]  counters [{4}]' -f $p, $prob.N, $w, $clText, $ctrText)
    }
}
Write-Output ''
Write-Output ("{0} with a world, {1} with none" -f $found, $none)

if ($Control) {
    $testPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'src\VLA_Tests_Query.bas'
    $text = [System.IO.File]::ReadAllText($testPath)
    $m = [regex]::Match($text, 'nFound = (\d+) And nNone = (\d+),')
    if (-not $m.Success) {
        Write-Output 'FAIL: the floor pin (nFound = N And nNone = M) is not in VLA_Tests_Query.bas - renamed or reworded?'
        exit 1
    }
    $pinFound = [int]$m.Groups[1].Value
    $pinNone = [int]$m.Groups[2].Value
    if (-not ($pinFound -eq $found -and $pinNone -eq $none)) {
        Write-Output ("FAIL: TestOptimizeSearch pins {0} with a world and {1} with none; the twin computes {2} and {3}" -f $pinFound, $pinNone, $found, $none)
        exit 1
    }
    Write-Output ("=== CONTROL: the VBA floor pin ({0} and {1}) is the twin's own count ===" -f $pinFound, $pinNone)
}
exit 0
