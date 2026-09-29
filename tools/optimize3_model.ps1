<#
optimize3_model.ps1 - OPTIMIZE.3's grounding cost model, re-fitted on the
engine that exists after DATALOG.14, with its superlinear term FITTED rather
than assumed.

WHY. OPTIMIZE.3 sets the ceilings a formula and a command may ground up to.
OPTIMIZE.0's paper model (time = a*F + b*J, a = 0.16 ms per source row,
b = 0.01 ms per produced row) was calibrated on an engine DATALOG.14 then
changed underneath it: its own first coefficient is, to within noise, the
per-source-row dictionary DATALOG.14 deleted. The OPTIMIZE.3 pre-flight
(tools/optimize3_preflight_steps.md, archived when OPTIMIZE.3 began; its
findings are in BETA_REARVIEW.md's OPTIMIZE.3 entry) re-timed the ladder
and found the old model tenfold pessimistic below ~100,000 peak rows and up
to 1.84x OPTIMISTIC past ~300,000. This refits it.

THE DATA. The pre-flight kept one number per rung - measured/model - and
four explicit seconds. This script ports archive/VLA_DiagO0.bas's own
BuildModel/RuleItem/RuleEnd to recompute each rung's F (source rows), J
(produced rows, every round) and peak (largest single step), multiplies the
ratio back by the old model to recover the measured seconds, and checks the
recovery against the four seconds the pre-flight wrote down explicitly
(9.039, 3.461, 8.063, 18.719). -Control fails if any is off by more than 1%.

RE-TIMED the same day by OPTIMIZE.3 slice 1's harness (tools/VLA_DiagO3.bas,
DATALOG as a VBA call on in-memory relations): ten of the eleven rungs it
shares reproduce within 0.91-1.07. The eleventh, A10 10x1 k3, took 17.148 s -
September's figure, not the 8.063 s fitted here - so the fit is 2.1x
optimistic at that rung, whose 279,300-row step is 5.6 times the formula's
per-step ceiling. The fit is kept as the pre-flight's record because the
ceiling does not depend on it: with 17.148 s in that rung's place, without
the rung, or with the harness's rungs added, the formula ceiling costs
0.88-0.92 s. BETA_REARVIEW.md's OPTIMIZE.3 entry, slice 1, has the table.

THE FIT, relative least squares (each rung weighted by 1/seconds, so a
0.05 s rung counts as much as a 9 s one):

    t = c + b*J + s*J*(peak / 100,000)

    c = 15 ms   a fixed cost per calculation; OPTIMIZE.0 measured 12 ms
    b = 7.6 us  per produced row
    s = 2.2 us  per produced row per 100,000 rows in the largest step -
                Scripting.Dictionary growth: O0DictCost measured about
                1 us per key per 100,000 keys, and a produced row pays
                about two dictionary operations

A source-row term was tried and fits to ZERO (-0.0004 ms): DATALOG.14
deleted it. That is why the pre-flight's two regimes - shapes bound by the
rows they READ and by the rows they PRODUCE - collapse into one model: after
DATALOG.14 both cost what they produce, and differ only in how many rows
they produce per row read.

-Control also holds the model to itself across families: fit on the A8
triples plus the native shapes and predict A10's quadruples, then the
reverse, each prediction within 0.7x-1.3x (measured 0.73-1.25).

WHAT IT PRINTS without -Control: the rungs, the fit, and the ceiling table -
how many produced rows a given time buys in each regime - from which
OPTIMIZE.3's formula ceiling (100,000 produced rows, 50,000 in any one step,
the owner's call 2026-09-24) was set: about 0.9 s by this model, 1.1 s
allowing the worst cross-family error.

SCOPE. This models DATALOG's own evaluator. OPTIMIZE.3 grounds through its
integer grounder (VLA_Datalog.DatalogGroundRules) instead, which this model
does not describe: the ladder re-measures that, and the entry says "re-
measure before raising" the ceilings on it.

House style: PowerShell 5.1, host-independent, no network, no Excel.
Usage:  powershell -File tools\optimize3_model.ps1 [-Control]
Exit 0 when -Control holds (always 0 without it); 1 otherwise.
#>
param([switch]$Control)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

function Get-Comb([int]$n, [int]$r) {
    if ($r -lt 0 -or $r -gt $n) { return 0.0 }
    $v = 1.0
    for ($i = 1; $i -le $r; $i++) { $v = $v * ($n - $r + $i) / $i }
    return [math]::Floor($v + 0.5)
}

# --- archive/VLA_DiagO0.bas's BuildModel, for the shapes the pre-flight kept ---
# Item kinds: 1 a positive atom over a Table, 2 over a derived predicate,
# 3 anything else. Round 2 re-runs the items ahead of every Table position
# before it meets the empty delta; the head's rows are counted twice.
$script:mF = 0.0; $script:mJ = 0.0; $script:mPeak = 0.0
$script:items = New-Object System.Collections.Generic.List[object]
function Add-TablesRead([double]$n) { $script:mJ += $n }
function Start-Rule() { $script:items = New-Object System.Collections.Generic.List[object] }
function Add-RuleItem([int]$kind, [double]$src, [double]$out) { $script:items.Add(@($kind, $src, $out)) }
function Complete-Rule([double]$headRows) {
    foreach ($it in $script:items) {
        $script:mF += $it[1]; $script:mJ += $it[2]
        if ($it[2] -gt $script:mPeak) { $script:mPeak = $it[2] }
    }
    $script:mJ += 2.0 * $headRows
    for ($p = 0; $p -lt $script:items.Count; $p++) {
        if ($script:items[$p][0] -eq 1) {
            for ($i = 0; $i -lt $p; $i++) { $script:mF += $script:items[$i][1]; $script:mJ += $script:items[$i][2] }
        } elseif ($script:items[$p][0] -eq 2) {
            for ($i = 0; $i -lt $script:items.Count; $i++) { $script:mF += $script:items[$i][1]; $script:mJ += $script:items[$i][2] }
            $script:mJ += $headRows
        }
    }
}
function Get-ModelCounts([int]$stepNo, [int]$nP, [int]$nW, [int]$k) {
    $script:mF = 0.0; $script:mJ = 0.0; $script:mPeak = 0.0
    $nS = 21.0 * $nW
    $pool = $nP * $nS
    $pruned = 18.0 * $nP * $nW
    $pw = [double]$nP * $nW
    switch ($stepNo) {
        1 { Add-TablesRead ($nS + $nP); Start-Rule; Add-RuleItem 1 $nS $nS; Add-RuleItem 1 $nP $pool; Complete-Rule $pool }
        4 { Add-TablesRead ($nS + $pruned); Start-Rule; Add-RuleItem 1 $nS $nS; Add-RuleItem 3 $pruned $nS; Complete-Rule $nS }
        5 { Add-TablesRead ($pw + $pruned); Start-Rule; Add-RuleItem 1 $pw $pw; Add-RuleItem 3 $pruned $pw; Complete-Rule $pw }
        6 {
            Add-TablesRead ($nS - 1 + $pool); Start-Rule
            Add-RuleItem 1 ($nS - 1) ($nS - 1)
            Add-RuleItem 1 $pool (($nS - 1) * $nP)
            Add-RuleItem 1 $pool (($nS - 1) * $nP)
            Complete-Rule (($nS - 1) * $nP)
        }
        default {
            if ($stepNo -eq 8) { $groups = $nS; $g = $nP } else { $groups = $pw; $g = 21 }
            Add-TablesRead $pool
            Start-Rule
            Add-RuleItem 1 $pool ($groups * $g)
            for ($j = 2; $j -le $k + 1; $j++) {
                Add-RuleItem 1 $pool ($groups * $g * (Get-Comb $g ($j - 1)))
                Add-RuleItem 3 0 ($groups * (Get-Comb $g $j))
            }
            Complete-Rule ($groups * (Get-Comb $g ($k + 1)))
        }
    }
    return [pscustomobject]@{ F = $script:mF; J = $script:mJ; Peak = $script:mPeak }
}

# --- the pre-flight's record, 2026-09-24 ---------------------------------
# Native shapes: seconds as recorded. A8 and A10: the measured/model ratio
# the pre-flight kept per rung; -1 seconds means "recover it from the ratio".
# A known second beside a ratio is the check that the recovery is right.
$rows = New-Object System.Collections.Generic.List[object]
function Add-Rung([string]$tag, [int]$s, [int]$p, [int]$w, [int]$k, [double]$ratio, [double]$secs, [double]$known) {
    $c = Get-ModelCounts $s $p $w $k
    $old = (0.16 * $c.F + 0.01 * $c.J) / 1000.0
    if ($secs -lt 0) { $secs = $ratio * $old }
    $script:rows.Add([pscustomobject]@{ Tag = $tag; Step = $s; F = $c.F; J = $c.J; Peak = $c.Peak; Old = $old; Secs = $secs; Known = $known })
}
Add-Rung 'A1 50x4'  1 50 4 2 0 0.074 -1
Add-Rung 'A4 50x4'  4 50 4 2 0 0.047 -1
Add-Rung 'A5 50x4'  5 50 4 2 0 0.047 -1
Add-Rung 'A6 50x4'  6 50 4 2 0 0.188 -1
Add-Rung 'A8 5x1'   8 5 1 2 0.40 -1 -1
Add-Rung 'A8 7x1'   8 7 1 2 0.41 -1 -1
Add-Rung 'A8 10x1'  8 10 1 2 0.51 -1 -1
Add-Rung 'A8 14x1'  8 14 1 2 0.65 -1 -1
Add-Rung 'A8 20x1'  8 20 1 2 0.77 -1 -1
Add-Rung 'A8 5x4'   8 5 4 2 0.26 -1 -1
Add-Rung 'A8 7x4'   8 7 4 2 0.37 -1 -1
Add-Rung 'A8 10x4'  8 10 4 2 0.50 -1 -1
Add-Rung 'A8 14x4'  8 14 4 2 0.68 -1 -1
Add-Rung 'A8 20x4'  8 20 4 2 1.03 -1 9.039
Add-Rung 'A10 5x1 k1'   10 5 1 1 0.58 -1 -1
Add-Rung 'A10 5x1 k2'   10 5 1 2 0.82 -1 -1
Add-Rung 'A10 10x1 k2'  10 10 1 2 0.85 -1 -1
Add-Rung 'A10 5x1 k3'   10 5 1 3 1.10 -1 3.461
Add-Rung 'A10 10x1 k3'  10 10 1 3 1.28 -1 8.063
Add-Rung 'A10 1x1 k5'   10 1 1 5 1.84 -1 18.719

# --- least squares, each rung weighted by 1/seconds ------------------------
function Solve-Linear([double[][]]$A, [double[]]$b) {
    $n = $b.Length
    $M = New-Object 'double[][]' $n
    for ($i = 0; $i -lt $n; $i++) {
        $row = New-Object 'double[]' ($n + 1)
        for ($j = 0; $j -lt $n; $j++) { $row[$j] = $A[$i][$j] }
        $row[$n] = $b[$i]
        $M[$i] = $row
    }
    for ($col = 0; $col -lt $n; $col++) {
        $piv = $col
        for ($r2 = $col + 1; $r2 -lt $n; $r2++) {
            $cand = [math]::Abs($M[$r2][$col]); $best = [math]::Abs($M[$piv][$col])
            if ($cand -gt $best) { $piv = $r2 }
        }
        $tmp = $M[$col]; $M[$col] = $M[$piv]; $M[$piv] = $tmp
        for ($r2 = 0; $r2 -lt $n; $r2++) {
            if ($r2 -ne $col) {
                $f = $M[$r2][$col] / $M[$col][$col]
                for ($j = $col; $j -le $n; $j++) { $M[$r2][$j] -= $f * $M[$col][$j] }
            }
        }
    }
    $x = New-Object 'double[]' $n
    for ($i = 0; $i -lt $n; $i++) { $x[$i] = $M[$i][$n] / $M[$i][$i] }
    return ,$x
}

# The model's three features. Every element is parenthesised on its own:
# in PowerShell `,` binds tighter than arithmetic.
function Get-Features($r) { return ,@(1.0, ($r.J / 1000.0), (($r.J / 1000.0) * ($r.Peak / 100000.0))) }

function Get-Fit($rowsIn) {
    $AtA = New-Object 'double[][]' 3
    for ($i = 0; $i -lt 3; $i++) { $AtA[$i] = New-Object 'double[]' 3 }
    $Atb = New-Object 'double[]' 3
    foreach ($r in $rowsIn) {
        $x = Get-Features $r
        $wgt = 1.0 / ($r.Secs * $r.Secs)
        for ($i = 0; $i -lt 3; $i++) {
            for ($j = 0; $j -lt 3; $j++) { $AtA[$i][$j] += $x[$i] * $x[$j] * $wgt }
            $Atb[$i] += $x[$i] * $r.Secs * $wgt
        }
    }
    return ,(Solve-Linear $AtA $Atb)
}

function Get-Prediction($theta, $r) {
    $x = Get-Features $r
    return $theta[0] * $x[0] + $theta[1] * $x[1] + $theta[2] * $x[2]
}

$failures = New-Object System.Collections.Generic.List[string]

Write-Output '=== OPTIMIZE.3 GROUNDING MODEL: the pre-flight ladder, recovered ==='
'{0,-13} {1,8} {2,10} {3,9} {4,8}' -f 'rung', 'F', 'J', 'peak', 'seconds' | Write-Output
foreach ($r in $rows) {
    '{0,-13} {1,8:N0} {2,10:N0} {3,9:N0} {4,8:N3}' -f $r.Tag, $r.F, $r.J, $r.Peak, $r.Secs | Write-Output
    if ($r.Known -gt 0) {
        $off = [math]::Abs($r.Secs / $r.Known - 1.0)
        if ($off -gt 0.01) {
            $failures.Add(("{0}: recovered {1:N3} s against the recorded {2:N3} s - the port of BuildModel is wrong" -f $r.Tag, $r.Secs, $r.Known))
        }
    }
}

$theta = Get-Fit $rows
Write-Output ''
Write-Output ('fit: c = {0:N1} ms, b = {1:N2} us a produced row, s = {2:N2} us a produced row per 100,000 peak rows' -f `
    ($theta[0] * 1000.0), ($theta[1] * 1000.0), ($theta[2] * 1000.0))
$sumsq = 0.0
$worstUnder = 99.0
foreach ($r in $rows) {
    $pred = Get-Prediction $theta $r
    $ratio = $pred / $r.Secs
    $sumsq += [math]::Log($ratio) * [math]::Log($ratio)
    if ($r.Secs -ge 0.1 -and $ratio -lt $worstUnder) { $worstUnder = $ratio }
}
$rms = [math]::Sqrt($sumsq / $rows.Count)
Write-Output ('rms log error {0:N3}; worst optimism above 0.1 s: predicted {1:N2} of measured' -f $rms, $worstUnder)

# A source-row term, for the record: it fits to about zero.
function Get-FeaturesWithSource($r) { return ,@(1.0, ($r.F / 1000.0), ($r.J / 1000.0), (($r.J / 1000.0) * ($r.Peak / 100000.0))) }
$A4 = New-Object 'double[][]' 4
for ($i = 0; $i -lt 4; $i++) { $A4[$i] = New-Object 'double[]' 4 }
$b4 = New-Object 'double[]' 4
foreach ($r in $rows) {
    $x = Get-FeaturesWithSource $r
    $wgt = 1.0 / ($r.Secs * $r.Secs)
    for ($i = 0; $i -lt 4; $i++) {
        for ($j = 0; $j -lt 4; $j++) { $A4[$i][$j] += $x[$i] * $x[$j] * $wgt }
        $b4[$i] += $x[$i] * $r.Secs * $wgt
    }
}
$theta4 = Solve-Linear $A4 $b4
Write-Output ('with a source-row term: a = {0:N4} ms a source row - DATALOG.14 deleted it' -f $theta4[1])

Write-Output ''
Write-Output '--- across families: fit on one, predict the other ---'
$fitA8 = Get-Fit @($rows | Where-Object { $_.Step -ne 10 })
$fitA10 = Get-Fit @($rows | Where-Object { $_.Step -ne 8 })
foreach ($pair in @(@('A10 from A8', $fitA8, 10), @('A8 from A10', $fitA10, 8))) {
    $lo = 99.0; $hi = 0.0
    foreach ($r in ($rows | Where-Object { $_.Step -eq $pair[2] })) {
        $ratio = (Get-Prediction $pair[1] $r) / $r.Secs
        if ($ratio -lt $lo) { $lo = $ratio }
        if ($ratio -gt $hi) { $hi = $ratio }
    }
    Write-Output ('  {0}: predicted {1:N2} to {2:N2} of measured' -f $pair[0], $lo, $hi)
    if (-not ($lo -ge 0.7 -and $hi -le 1.3)) {
        $failures.Add(("{0}: a prediction left 0.7x-1.3x ({1:N2} to {2:N2}) - the model no longer transfers across shape families" -f $pair[0], $lo, $hi))
    }
}

Write-Output ''
Write-Output '--- what a given time buys, in produced rows (J) and rows in the largest step (peak) ---'
Write-Output '    produce-bound shapes (subsets, triples) produce about 2.35 rows per step row; read-bound (the pair relation) about 6'
foreach ($target in @(0.5, 1.0, 2.0, 10.0)) {
    foreach ($perPeak in @(2.35, 6.0)) {
        $lo = 1.0; $hiP = 5000000.0
        for ($it = 0; $it -lt 80; $it++) {
            $mid = ($lo + $hiP) / 2.0
            $j = $perPeak * $mid
            $t = $theta[0] + $theta[1] * ($j / 1000.0) + $theta[2] * ($j / 1000.0) * ($mid / 100000.0)
            if ($t -gt $target) { $hiP = $mid } else { $lo = $mid }
        }
        Write-Output ('  {0,5:N1} s   J/peak {1,4:N2}:  peak {2,9:N0}   J {3,10:N0}' -f $target, $perPeak, $lo, ($perPeak * $lo))
    }
}
$jFormula = 100000.0; $peakFormula = 50000.0
$tFormula = $theta[0] + $theta[1] * ($jFormula / 1000.0) + $theta[2] * ($jFormula / 1000.0) * ($peakFormula / 100000.0)
Write-Output ''
Write-Output ('the formula ceiling, 100,000 produced rows with 50,000 in the largest step: {0:N2} s by this model, {1:N2} s at the worst cross-family optimism' -f $tFormula, ($tFormula / 0.8))

if ($Control) {
    Write-Output ''
    # Every range below is written as what MUST hold, so a NaN - which
    # fails every comparison - fails the check. Written the other way
    # round ("fail if below or above"), a singular fit's NaN coefficients
    # passed all four: this control's own mutation run found it, by
    # zeroing the superlinear feature.
    if (-not ($rms -le 0.2)) { $failures.Add(("the fit's rms log error is {0:N3}, not at most 0.2" -f $rms)) }
    if (-not ($theta[0] -ge 0.005 -and $theta[0] -le 0.03)) { $failures.Add(("the fixed cost fits at {0:N1} ms, not within 5-30" -f ($theta[0] * 1000.0))) }
    if (-not ($theta[1] -ge 0.005 -and $theta[1] -le 0.010)) { $failures.Add(("the per-row cost fits at {0:N2} us, not within 5-10" -f ($theta[1] * 1000.0))) }
    if (-not ($theta[2] -ge 0.001 -and $theta[2] -le 0.004)) { $failures.Add(("the superlinear term fits at {0:N2} us, not within 1-4" -f ($theta[2] * 1000.0))) }
    if (-not ([math]::Abs($theta4[1]) -le 0.002)) { $failures.Add(("the source-row term fits at {0:N4} ms, not about zero" -f $theta4[1])) }
    if ($failures.Count -gt 0) {
        foreach ($f in $failures) { Write-Output "FAIL: $f" }
        Write-Output "=== CONTROL: FAILED - $($failures.Count) problem(s) ==="
        exit 1
    }
    Write-Output '=== CONTROL: the recovery reproduces every recorded second within 1%, and the fit holds across families ==='
}
exit 0
