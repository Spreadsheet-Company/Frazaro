# SPDX-License-Identifier: 0BSD
#
# DATALOG.14 - a standalone cost model of VLA_Datalog's semi-naive fixpoint.
#
# WHY THIS EXISTS. DATALOG.14 names two causes for a closure's cost and asks
# which one the re-run ladder's numbers actually blame. The answer is
# arithmetic, and hand arithmetic over a hundred fixpoint rounds is exactly
# the kind of arithmetic that is confidently wrong. So this script SIMULATES
# the engine - the same rule order, the same round loop, the same per-body-
# position delta substitution, the same filter-then-join-then-project chain -
# and COUNTS the operations whose unit costs DATALOG.12's pass 2 measured.
#
# It computes no times of its own. It counts, and multiplies by ONE measured
# constant: 0.157 ms for CreateObject("Scripting.Dictionary") + CompareMode
# (DATALOG.12 pass 2). Everything else that engine does per tuple is smaller
# by two orders of magnitude and is left in the residual on purpose.
#
# THE CONTROL. -Control replays the shapes DATALOG.13's live pass measured
# and prints predicted beside measured. If the ratio is not near-constant
# across shapes and sizes, this model is wrong and nothing it says about a
# proposed change should be believed.
#
# It calls nothing in Frazaro and touches no workbook. Host-independent.

[CmdletBinding()]
param(
    [ValidateSet('shipped', 'planfilter', 'identity', 'freeindex', 'both')]
    [string]$Mode = 'shipped',
    [ValidateSet('scan', 'wide', 'chain')]
    [string]$Shape = 'chain',
    [int]$Rows = 100,
    [switch]$MagicSets,
    [switch]$Control,
    [switch]$Forks,
    [switch]$Causes
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# The one measured unit cost, DATALOG.12 pass 2, the owner's machine.
$script:MS_PER_DICTNEW = 0.157

# ---------------------------------------------------------------- counters

$script:cDictNew = 0     # VlaDictNew / RelNew's NewIndex / BuildJoinIndex
$script:cTryAdd = 0      # RelTryAdd: one TupleKey build + Exists + Add
$script:cProbe = 0       # RelJoin: one PartialKey build + one bucket lookup
$script:cTupleTouch = 0  # one tuple examined by a filter, dictionary or not

# DATALOG.14 names two causes. These split every tuple a later round touches
# between them, so the entry's "separable in principle" can be answered with
# a number. Round 1 is neither: it is the bootstrap both causes share.
$script:cBoot = 0        # round 1, any position
$script:cRescan = 0      # a later round, a NON-delta position: cause 2
$script:cPairs = 0       # a later round, the DELTA position: cause 1
$script:ctxRound = 1
$script:ctxIsDelta = $false

function Reset-Counters {
    $script:cDictNew = 0
    $script:cTryAdd = 0
    $script:cProbe = 0
    $script:cTupleTouch = 0
    $script:cBoot = 0
    $script:cRescan = 0
    $script:cPairs = 0
    $script:ctxRound = 1
    $script:ctxIsDelta = $false
}

function Add-CauseTouch {
    param([int]$Count)
    if ($script:ctxRound -le 1) { $script:cBoot += $Count }
    elseif ($script:ctxIsDelta) { $script:cPairs += $Count }
    else { $script:cRescan += $Count }
}

# ---------------------------------------------------------------- relations
#
# VLA_Relation.RelNew: a 3-item record (arity, tuples, index). The index is
# a Scripting.Dictionary, so constructing a relation costs one CreateObject.

function New-Rel {
    param([int]$Arity)
    $script:cDictNew++
    [pscustomobject]@{
        Arity  = $Arity
        Tuples = [System.Collections.Generic.List[object]]::new()
        Keys   = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    }
}

# RelTryAdd: TupleKey, then Exists, then Add. Returns $true iff new.
function Add-Tuple {
    param($Rel, [string[]]$Tuple)
    $script:cTryAdd++
    $key = [string]::Join([char]31, $Tuple)
    if ($Rel.Keys.Contains($key)) { return $false }
    [void]$Rel.Keys.Add($key)
    [void]$Rel.Tuples.Add($Tuple)
    return $true
}

# ---------------------------------------------------------------- atoms
#
# An atom is @{ Pred = 'name'; Args = @( @{Var=$true; Text='X'}, ... ) },
# mirroring ParseAtom's record. A "constraining" atom is one AtomMatches can
# reject a tuple for: it carries a constant, or names one variable twice.
# An atom that constrains nothing passes every tuple of every relation.

function New-Atom {
    # NOT $Args: that is a PowerShell automatic variable, and a parameter of
    # that name binds nothing - every atom came back with no arguments at all,
    # silently, and the first symptom was a null column index four functions
    # away. The trap is in the project's own notes; this is it again.
    param([string]$Pred, [string[]]$ArgTexts)
    $argRecs = [System.Collections.Generic.List[object]]::new()
    foreach ($a in $ArgTexts) {
        # A leading quote marks a constant, exactly as IsVariableAtom reads it.
        if ($a.StartsWith('"')) {
            [void]$argRecs.Add(@{ Var = $false; Text = $a.Substring(1) })
        } else {
            [void]$argRecs.Add(@{ Var = $true; Text = $a })
        }
    }
    @{ Kind = 'pos'; Pred = $Pred; Args = $argRecs }
}

function Test-AtomConstrains {
    param($Atom)
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($a in $Atom.Args) {
        if (-not $a.Var) { return $true }
        if (-not $seen.Add($a.Text)) { return $true }
    }
    return $false
}

# AtomMatches, as shipped: one fresh Scripting.Dictionary per tuple.
function Test-AtomMatches {
    param($Atom, [string[]]$Tup)
    $seen = @{}
    for ($i = 0; $i -lt $Atom.Args.Count; $i++) {
        $a = $Atom.Args[$i]
        $v = $Tup[$i]
        if ($a.Var) {
            $nm = $a.Text.ToUpperInvariant()
            if ($seen.ContainsKey($nm)) {
                if ($seen[$nm] -cne $v) { return $false }
            } else {
                $seen[$nm] = $v
            }
        } else {
            if ($a.Text -cne $v) { return $false }
        }
    }
    return $true
}

# FilterAtomRelation, in three shapes.
#   shipped    - today: a fresh dictionary per tuple, always a full copy.
#   planfilter - the atom's checks compiled once; still a full copy.
#   identity   - planfilter, plus: an atom that constrains nothing returns
#                the relation itself, untouched.
function Select-AtomRelation {
    param($Rel, $Atom, [string]$FilterMode)

    # Counted before the early return on purpose: the cause split is a
    # property of the program and its data - what the SHIPPED filter reads -
    # not of the variant being modelled, so every variant reports the same
    # split and only the totals move.
    Add-CauseTouch -Count $Rel.Tuples.Count

    if ($FilterMode -eq 'identity' -and -not (Test-AtomConstrains $Atom)) {
        return $Rel
    }

    # An ORACLE index: a per-predicate index on the atom's bound positions
    # that is already built, never maintained and never paid for. It cannot
    # exist, and that is the point - it is the CEILING of what PROLOG.29's
    # sibling could buy here, so if the ceiling is low the shape is settled
    # without building it.
    $oracle = ($FilterMode -eq 'freeindex' -and (Test-AtomConstrains $Atom))

    $out = New-Rel -Arity $Rel.Arity
    foreach ($t in $Rel.Tuples) {
        $ok = Test-AtomMatches -Atom $Atom -Tup $t
        if ($oracle -and -not $ok) { continue }
        $script:cTupleTouch++
        if ($FilterMode -eq 'shipped' -or $FilterMode -eq 'freeindex') {
            $script:cDictNew++          # AtomMatches' own `seen`
        }
        if ($ok) { [void](Add-Tuple -Rel $out -Tuple $t) }
    }
    return $out
}

# ---------------------------------------------------------------- join
#
# RelJoin: build a hash index on whichever side has fewer tuples, probe with
# the other, emit left++right for every matching pair. BuildJoinIndex costs
# one CreateObject; each indexed row and each probe costs one PartialKey.

function Join-Rel {
    param($Left, [int[]]$LeftCols, $Right, [int[]]$RightCols)

    $buildOnLeft = ($Left.Tuples.Count -le $Right.Tuples.Count)
    if ($buildOnLeft) {
        $buildRel = $Left;  $buildCols = $LeftCols
        $probeRel = $Right; $probeCols = $RightCols
    } else {
        $buildRel = $Right; $buildCols = $RightCols
        $probeRel = $Left;  $probeCols = $LeftCols
    }

    $script:cDictNew++                  # BuildJoinIndex's own CreateObject
    $idx = @{}
    foreach ($t in $buildRel.Tuples) {
        $script:cProbe++
        $k = Get-PartialKey -Tup $t -Cols $buildCols
        if (-not $idx.ContainsKey($k)) {
            $idx[$k] = [System.Collections.Generic.List[object]]::new()
        }
        [void]$idx[$k].Add($t)
    }

    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($p in $probeRel.Tuples) {
        $script:cProbe++
        $k = Get-PartialKey -Tup $p -Cols $probeCols
        if (-not $idx.ContainsKey($k)) { continue }
        foreach ($b in $idx[$k]) {
            if ($buildOnLeft) {
                [void]$out.Add(($b + $p))
            } else {
                [void]$out.Add(($p + $b))
            }
        }
    }
    # `,` or PowerShell unrolls the list on return, and a join that produced
    # exactly ONE row would come back as that row - the caller would then
    # iterate its cells as if they were rows. The wide-org and scan controls
    # never build a one-row join, so they passed green over this defect; the
    # chain hits it in its third round. A control only covers the construct
    # it actually contains.
    return , $out
}

function Get-PartialKey {
    param([string[]]$Tup, [int[]]$Cols)
    if ($Cols.Count -eq 0) { return '' }
    $sb = [System.Text.StringBuilder]::new()
    foreach ($c in $Cols) { [void]$sb.Append($Tup[$c]).Append([char]31) }
    return $sb.ToString()
}

# ---------------------------------------------------------------- rule body
#
# EvalRuleBody: colOf (one dictionary) and RelUnit (one relation, so one
# more) before the first body item; then, per positive body item, a filter,
# a join and a projection.

function Invoke-RuleBody {
    param($Rule, [int]$DeltaPos, $Relations, $Deltas, [string]$FilterMode)

    $script:cDictNew++                  # colOf
    $accum = New-Rel -Arity 0           # RelUnit
    [void](Add-Tuple -Rel $accum -Tuple @())
    $colOf = @{}                        # variable name -> 0-based column

    $bi = 0
    foreach ($atom in $Rule.Body) {
        $bi++

        if ($atom.Kind -eq 'cmp') {
            # A comparison filters accum in place, per row, with no dictionary
            # and no predicate lookup.
            $kept = New-Rel -Arity $accum.Arity
            foreach ($t in $accum.Tuples) {
                $script:cTupleTouch++
                $col = $colOf[$atom.Var.ToUpperInvariant()]
                if ([double]$t[$col] -gt [double]$atom.Bound) {
                    [void](Add-Tuple -Rel $kept -Tuple $t)
                }
            }
            $accum = $kept
            if ($accum.Tuples.Count -eq 0) { return $null }
            continue
        }

        $script:ctxIsDelta = ($bi -eq $DeltaPos)
        if ($bi -eq $DeltaPos) {
            if (-not $Deltas.ContainsKey($atom.Pred)) { return $null }
            $src = $Deltas[$atom.Pred]
        } else {
            if (-not $Relations.ContainsKey($atom.Pred)) { return $null }
            $src = $Relations[$atom.Pred]
        }

        $filtered = Select-AtomRelation -Rel $src -Atom $atom -FilterMode $FilterMode
        if ($filtered.Tuples.Count -eq 0) { return $null }

        # BuildJoinPlan: a variable already in colOf is a join column; one
        # seen for the first time in this atom becomes a new output column.
        $leftCols = [System.Collections.Generic.List[int]]::new()
        $rightCols = [System.Collections.Generic.List[int]]::new()
        $newCols = [System.Collections.Generic.List[int]]::new()
        $newNames = [System.Collections.Generic.List[string]]::new()
        $seenHere = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        for ($i = 0; $i -lt $atom.Args.Count; $i++) {
            $a = $atom.Args[$i]
            if (-not $a.Var) { continue }
            if (-not $seenHere.Add($a.Text)) { continue }
            $nm = $a.Text.ToUpperInvariant()
            if ($colOf.ContainsKey($nm)) {
                [void]$leftCols.Add($colOf[$nm])
                [void]$rightCols.Add($i)
            } else {
                [void]$newCols.Add($i)
                [void]$newNames.Add($nm)
            }
        }

        $accumArity = $colOf.Count
        $joined = Join-Rel -Left $accum -LeftCols $leftCols.ToArray() -Right $filtered -RightCols $rightCols.ToArray()

        # ProjectAfterJoin: keep accum's columns, append the atom's new ones.
        $projected = New-Rel -Arity ($accumArity + $newCols.Count)
        foreach ($row in $joined) {
            $nt = [string[]]::new($accumArity + $newCols.Count)
            for ($i = 0; $i -lt $accumArity; $i++) { $nt[$i] = $row[$i] }
            for ($i = 0; $i -lt $newCols.Count; $i++) {
                $nt[$accumArity + $i] = $row[$accumArity + $newCols[$i]]
            }
            [void](Add-Tuple -Rel $projected -Tuple $nt)
        }
        for ($i = 0; $i -lt $newNames.Count; $i++) {
            $colOf[$newNames[$i]] = $accumArity + $i
        }
        $accum = $projected
        if ($accum.Tuples.Count -eq 0) { return $null }
    }

    # The head projection, one row out per surviving accum row.
    $outp = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $accum.Tuples) {
        $head = [string[]]::new($Rule.Head.Args.Count)
        for ($i = 0; $i -lt $Rule.Head.Args.Count; $i++) {
            $a = $Rule.Head.Args[$i]
            if ($a.Var) { $head[$i] = $row[$colOf[$a.Text.ToUpperInvariant()]] }
            else { $head[$i] = $a.Text }
        }
        [void]$outp.Add($head)
    }
    return , $outp                      # the same unrolling trap as Join-Rel's
}

function Invoke-OneRulePass {
    param($Rule, [int]$DeltaPos, $Relations, $Deltas, $NewDeltas, [string]$FilterMode)

    $derived = Invoke-RuleBody -Rule $Rule -DeltaPos $DeltaPos -Relations $Relations -Deltas $Deltas -FilterMode $FilterMode
    if ($null -eq $derived -or $derived.Count -eq 0) { return $false }

    $headPred = $Rule.Head.Pred
    if (-not $Relations.ContainsKey($headPred)) {
        $Relations[$headPred] = New-Rel -Arity $Rule.Head.Args.Count
    }
    $full = $Relations[$headPred]
    if (-not $NewDeltas.ContainsKey($headPred)) {
        $NewDeltas[$headPred] = New-Rel -Arity $Rule.Head.Args.Count
    }
    $newRel = $NewDeltas[$headPred]

    $foundNew = $false
    foreach ($t in $derived) {
        if (Add-Tuple -Rel $full -Tuple $t) {
            [void](Add-Tuple -Rel $newRel -Tuple $t)
            $foundNew = $true
        }
    }
    return $foundNew
}

function Invoke-Fixpoint {
    param($Rules, $Relations, [string]$FilterMode)

    $script:cDictNew++                  # emptyDeltas
    $emptyDeltas = @{}
    $script:cDictNew++                  # newDeltas
    $newDeltas = @{}

    $anyNew = $false
    foreach ($r in $Rules) {
        if (Invoke-OneRulePass -Rule $r -DeltaPos 0 -Relations $Relations -Deltas $emptyDeltas -NewDeltas $newDeltas -FilterMode $FilterMode) {
            $anyNew = $true
        }
    }

    $round = 1
    while ($anyNew) {
        $round++
        $script:ctxRound = $round
        if ($round -gt 10000) { throw 'MAX_ROUNDS' }
        $deltas = $newDeltas
        $script:cDictNew++
        $newDeltas = @{}
        $anyNew = $false
        foreach ($r in $Rules) {
            for ($bi = 1; $bi -le $r.Body.Count; $bi++) {
                if ($r.Body[$bi - 1].Kind -eq 'cmp') { continue }
                if (Invoke-OneRulePass -Rule $r -DeltaPos $bi -Relations $Relations -Deltas $deltas -NewDeltas $newDeltas -FilterMode $FilterMode) {
                    $anyNew = $true
                }
            }
        }
    }
    return $round
}

# ---------------------------------------------------------------- fixtures
#
# The generator's own closed forms, from archive/VLA_Diag12.bas' FillOrg and
# FillBills - rebuilt here so the model is measured against the same data the
# live ladder ran on.

function New-Program {
    param([string]$ShapeName, [int]$N, [switch]$Magic)

    $relations = @{}

    if ($ShapeName -eq 'scan') {
        # Bills(Bill, Vendor, Amount): amount = ((i*37) mod 100) * 200.
        $bills = New-Rel -Arity 3
        for ($i = 1; $i -le $N; $i++) {
            $amount = ((($i * 37) % 100) * 200)
            [void](Add-Tuple -Rel $bills -Tuple @("B$i", "V$($i % 50)", "$amount"))
        }
        $relations['bills12x1'] = $bills
        $rules = @(
            @{ Head = (New-Atom 'big' @('Bill'))
               Body = @((New-Atom 'bills12x1' @('Bill', 'VlaCol2', 'Amount')),
                        @{ Kind = 'cmp'; Var = 'Amount'; Bound = 10000; Pred = '>'; Args = @() }) },
            @{ Head = (New-Atom 'vla-ask-big' @('Bill'))
               Body = @((New-Atom 'big' @('Bill'))) }
        )
        return @{ Relations = $relations; Rules = $rules; Query = 'vla-ask-big' }
    }

    # wide and chain share FillOrg: row i is E(i+1) -> E(boss).
    $tbl = New-Rel -Arity 2
    for ($i = 1; $i -le $N; $i++) {
        if ($ShapeName -eq 'chain') { $boss = $i } else { $boss = [int][math]::Floor(($i - 1) / 10) + 1 }
        [void](Add-Tuple -Rel $tbl -Tuple @("E$($i + 1)", "E$boss"))
    }
    $tableName = if ($ShapeName -eq 'chain') { 'chain12x8' } else { 'reports12x7' }
    $relations[$tableName] = $tbl

    if ($Magic) {
        # The bound argument pushed in before the closure materialises: the
        # second argument of every vla-any- atom is the constant the question
        # asked about, so only pairs ending at E1 are ever built.
        $rules = @(
            @{ Head = (New-Atom 'reports-to' @('Person', 'Boss'))
               Body = @((New-Atom $tableName @('Person', 'Boss'))) },
            @{ Head = (New-Atom 'vla-any-reports--to' @('X', '"E1'))
               Body = @((New-Atom 'reports-to' @('X', '"E1'))) },
            @{ Head = (New-Atom 'vla-any-reports--to' @('X', '"E1'))
               Body = @((New-Atom 'reports-to' @('X', 'Z')),
                        (New-Atom 'vla-any-reports--to' @('Z', '"E1'))) },
            @{ Head = (New-Atom 'vla-ask-reports-to' @('Who'))
               Body = @((New-Atom 'vla-any-reports--to' @('Who', '"E1'))) }
        )
    } else {
        # Exactly what the grammar writes: the rules range first (the
        # reports-to rule), then the question's own two closure rules and its
        # narrowing rule - prolog-ask-place's TEXTJOIN(rules) & " " & question.
        $rules = @(
            @{ Head = (New-Atom 'reports-to' @('Person', 'Boss'))
               Body = @((New-Atom $tableName @('Person', 'Boss'))) },
            @{ Head = (New-Atom 'vla-any-reports--to' @('X', 'Y'))
               Body = @((New-Atom 'reports-to' @('X', 'Y'))) },
            @{ Head = (New-Atom 'vla-any-reports--to' @('X', 'Y'))
               Body = @((New-Atom 'reports-to' @('X', 'Z')),
                        (New-Atom 'vla-any-reports--to' @('Z', 'Y'))) },
            @{ Head = (New-Atom 'vla-ask-reports-to' @('Who'))
               Body = @((New-Atom 'vla-any-reports--to' @('Who', '"E1'))) }
        )
    }
    return @{ Relations = $relations; Rules = $rules; Query = 'vla-ask-reports-to' }
}

function Invoke-Model {
    param([string]$ShapeName, [int]$N, [string]$FilterMode, [switch]$Magic)

    Reset-Counters
    $p = New-Program -ShapeName $ShapeName -N $N -Magic:$Magic
    $rounds = Invoke-Fixpoint -Rules $p.Rules -Relations $p.Relations -FilterMode $FilterMode

    $answer = 0
    if ($p.Relations.ContainsKey($p.Query)) { $answer = $p.Relations[$p.Query].Tuples.Count }
    $pairs = 0
    if ($p.Relations.ContainsKey('vla-any-reports--to')) {
        $pairs = $p.Relations['vla-any-reports--to'].Tuples.Count
    }

    [pscustomobject]@{
        Shape       = $ShapeName
        N           = $N
        FilterMode  = $FilterMode
        Magic       = [bool]$Magic
        Rounds      = $rounds
        Answer      = $answer
        Pairs       = $pairs
        DictNew     = $script:cDictNew
        TryAdd      = $script:cTryAdd
        ProbeKeys   = $script:cProbe
        TupleTouch  = $script:cTupleTouch
        Bootstrap   = $script:cBoot
        Rescan      = $script:cRescan
        PairTouch   = $script:cPairs
        PredictedS  = [math]::Round(($script:cDictNew * $script:MS_PER_DICTNEW) / 1000.0, 3)
    }
}

# ---------------------------------------------------------------- control

# The post-DATALOG.13 ladder, from BETA_ROADMAP1.md's own table. These are
# the numbers the model has to reproduce before anything it says is evidence.
$script:MEASURED = @(
    @{ Shape = 'scan';  N = 1000;  Seconds = 0.383;  Answer = 490 },
    @{ Shape = 'scan';  N = 10000; Seconds = 3.813;  Answer = 4900 },
    @{ Shape = 'wide';  N = 100;   Seconds = 0.281;  Answer = 100 },
    @{ Shape = 'wide';  N = 300;   Seconds = 0.992;  Answer = 300 },
    @{ Shape = 'wide';  N = 1000;  Seconds = 3.344;  Answer = 1000 },
    @{ Shape = 'wide';  N = 3000;  Seconds = 11.281; Answer = 3000 },
    @{ Shape = 'chain'; N = 100;   Seconds = 4.125;  Answer = 100 }
)

# The second term, and why it has to be here. The first model priced ONLY
# CreateObject, which accounted for 82-89% of every measured point - green,
# and useless for predicting an after, because every shape under
# consideration removes exactly that term and leaves the unpriced 11-18% as
# the whole remaining cost. A model calibrated on a term that then goes away
# cannot extrapolate past its own removal. So the rest of a tuple's work -
# one filter touch, one RelTryAdd, one join key - is fitted here as one
# constant over the same seven points, and reported with its spread.
function Get-TupleOps {
    param($R)
    return $R.TupleTouch + $R.TryAdd + $R.ProbeKeys
}

function Get-FittedTupleCost {
    $fits = [System.Collections.Generic.List[double]]::new()
    foreach ($m in $script:MEASURED) {
        $r = Invoke-Model -ShapeName $m.Shape -N $m.N -FilterMode 'shipped'
        $residualMs = ($m.Seconds * 1000.0) - ($r.DictNew * $script:MS_PER_DICTNEW)
        $ops = Get-TupleOps -R $r
        if ($ops -gt 0) { [void]$fits.Add($residualMs / $ops) }
    }
    return , $fits
}

if ($Control) {
    Write-Output 'CONTROL - the model against DATALOG.13''s own measured ladder'
    Write-Output ''
    Write-Output 'shape  rows    answer  ok    rounds  pairs   dictNew   tupleOps   dictOnly  measured  share  ms/op'
    Write-Output '-----  ------  ------  ----  ------  ------  --------  ---------  --------  --------  -----  ------'
    $shares = [System.Collections.Generic.List[double]]::new()
    $fits = [System.Collections.Generic.List[double]]::new()
    foreach ($m in $script:MEASURED) {
        $r = Invoke-Model -ShapeName $m.Shape -N $m.N -FilterMode 'shipped'
        $ok = if ($r.Answer -eq $m.Answer) { 'yes' } else { "NO($($r.Answer))" }
        $ops = Get-TupleOps -R $r
        $share = $r.PredictedS / $m.Seconds
        $perOp = (($m.Seconds * 1000.0) - ($r.DictNew * $script:MS_PER_DICTNEW)) / $ops
        [void]$shares.Add($share)
        [void]$fits.Add($perOp)
        '{0,-5}  {1,6}  {2,6}  {3,-4}  {4,6}  {5,6}  {6,8}  {7,9}  {8,8}  {9,8}  {10,5}  {11,6}' -f `
            $r.Shape, $r.N, $r.Answer, $ok, $r.Rounds, $r.Pairs, $r.DictNew, $ops, `
            $r.PredictedS, $m.Seconds, [math]::Round($share, 3), [math]::Round($perOp, 4) | Write-Output
    }
    $lo = ($shares | Measure-Object -Minimum).Minimum
    $hi = ($shares | Measure-Object -Maximum).Maximum
    $flo = ($fits | Measure-Object -Minimum).Minimum
    $fhi = ($fits | Measure-Object -Maximum).Maximum
    $favg = ($fits | Measure-Object -Average).Average
    Write-Output ''
    Write-Output ("CreateObject alone covers {0:P1} to {1:P1} of measured, over {2} points." -f $lo, $hi, $shares.Count)
    Write-Output ("The residual fits {0:N4} ms per tuple operation (spread {1:N4} to {2:N4})." -f $favg, $flo, $fhi)
    if (($hi - $lo) -le 0.15) {
        Write-Output 'CONTROL GREEN on term one: one measured constant tracks every shape and size.'
    } else {
        Write-Output 'CONTROL RED: the share is not constant. The model is wrong; do not trust its forks.'
    }
    if (($fhi - $flo) -le ($favg * 0.75)) {
        Write-Output 'CONTROL GREEN on term two: one fitted constant covers the residual within a factor under two.'
    } else {
        Write-Output 'CONTROL AMBER on term two: the residual per operation varies by more than a factor of two,'
        Write-Output '  so an AFTER predicted from it is an order of magnitude, not a number. Say so when quoting it.'
    }
    return
}

if ($Forks) {
    $fits = Get-FittedTupleCost
    $perOp = ($fits | Measure-Object -Average).Average
    Write-Output 'THE FORKS, each replayed over the same fixtures.'
    Write-Output ("Two-term model: {0:N3} ms per CreateObject (measured) + {1:N4} ms per tuple operation (fitted)." -f $script:MS_PER_DICTNEW, $perOp)
    Write-Output 'An AFTER is an estimate with the second term doing most of the work. Treat it as a size, not a number.'
    Write-Output ''
    $cases = @(
        @{ Shape = 'chain'; N = 100;  Measured = 4.125 },
        @{ Shape = 'chain'; N = 250;  Measured = 0 },
        @{ Shape = 'wide';  N = 1000; Measured = 3.344 },
        @{ Shape = 'wide';  N = 3000; Measured = 11.281 },
        @{ Shape = 'scan';  N = 1000; Measured = 0.383 }
    )
    Write-Output 'shape  rows    variant               pairs   rounds  dictNew   tupleOps   estimate  gain    measured'
    Write-Output '-----  ------  --------------------  ------  ------  --------  ---------  --------  ------  --------'
    foreach ($c in $cases) {
        $variants = @(
            @{ Name = 'shipped';            Filter = 'shipped';    Magic = $false },
            @{ Name = 'compiled filter';    Filter = 'planfilter'; Magic = $false },
            @{ Name = 'identity skip';      Filter = 'identity';   Magic = $false },
            @{ Name = 'free index (ceiling)'; Filter = 'freeindex';  Magic = $false },
            @{ Name = 'magic sets only';    Filter = 'shipped';    Magic = $true },
            @{ Name = 'magic + identity';   Filter = 'identity';   Magic = $true }
        )
        $base = $null
        foreach ($v in $variants) {
            if ($c.Shape -eq 'scan' -and $v.Magic) { continue }
            $r = Invoke-Model -ShapeName $c.Shape -N $c.N -FilterMode $v.Filter -Magic:$v.Magic
            $est = (($r.DictNew * $script:MS_PER_DICTNEW) + ((Get-TupleOps -R $r) * $perOp)) / 1000.0
            if ($null -eq $base) { $base = $est }
            $gain = if ($est -gt 0) { [math]::Round($base / $est, 1) } else { 0 }
            $meas = if ($v.Name -eq 'shipped' -and $c.Measured -gt 0) { '{0:N3}' -f $c.Measured } else { '' }
            '{0,-5}  {1,6}  {2,-20}  {3,6}  {4,6}  {5,8}  {6,9}  {7,8:N3}  {8,5}x  {9,8}' -f `
                $r.Shape, $r.N, $v.Name, $r.Pairs, $r.Rounds, $r.DictNew, (Get-TupleOps -R $r), `
                $est, $gain, $meas | Write-Output
        }
        Write-Output ''
    }
    return
}

if ($Causes) {
    Write-Output 'FORK 2 - which of DATALOG.14''s two causes the numbers blame.'
    Write-Output ''
    Write-Output 'Every tuple the shipped filter reads, split three ways. Round 1 is the'
    Write-Output 'bootstrap both causes share and is charged to neither.'
    Write-Output ''
    Write-Output 'shape  rows    bootstrap  rescan(cause2)  pairs(cause1)  rescan%  pairs%'
    Write-Output '-----  ------  ---------  --------------  -------------  -------  ------'
    foreach ($m in $script:MEASURED) {
        $r = Invoke-Model -ShapeName $m.Shape -N $m.N -FilterMode 'shipped'
        $later = $r.Rescan + $r.PairTouch
        if ($later -eq 0) { $later = 1 }
        '{0,-5}  {1,6}  {2,9}  {3,14}  {4,13}  {5,6:P0}  {6,6:P0}' -f `
            $r.Shape, $r.N, $r.Bootstrap, $r.Rescan, $r.PairTouch, `
            ($r.Rescan / $later), ($r.PairTouch / $later) | Write-Output
    }
    return
}

$modes = if ($Mode -eq 'both') { @('shipped', 'identity') } else { @($Mode) }
foreach ($m in $modes) {
    Invoke-Model -ShapeName $Shape -N $Rows -FilterMode $m -Magic:$MagicSets
}
