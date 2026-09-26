<#
check_datalog_per_tuple_alloc.ps1 - DATALOG.14's mechanical pin: nothing
allocates a dictionary once per TUPLE on DATALOG's per-row path.

DATALOG.14: after DATALOG.13 removed the TypeName guard, the largest term
left in the engine was a fresh Scripting.Dictionary built for every tuple a
filter tested - 0.157 ms a call (DATALOG.12 pass 2), about 21,000 of them in
a closure over a hundred-link chain. Two sites did it: the atom filter, whose
`seen` answered a question that depends only on the atom's own shape, and
ComputeAggregateGroups, whose `seenLocal` did the same given colOf. Both now
work the answer out ONCE - into an atom plan and an aggregate plan - and walk
it per tuple, through one shared PlanMatches, with no allocation at all.

WHY A STATIC SCAN: the regression is SILENT, the same reasoning
check_vladict_guard.ps1 gives. A VlaDictNew put back inside a tuple loop
returns identical answers, so the pure, DSL and host suites all stay green
and every parity program still agrees; only a live timing on a real-sized
Table would notice, and nobody runs one by habit. The same is true of
deleting FilterAtomRelation's identity return or unwiring the specialisation:
the answers are correct either way, and only the ladder would say.

WHAT IS CHECKED:

  Rule A - no allocation at all, anywhere in the procedure. PlanMatches runs
  once per tuple in its entirety, so any VlaDictNew or CreateObject in it is
  per-tuple by construction.

  Rule B - no allocation INSIDE a tuple loop. These legitimately allocate
  once per call (a result relation, a groups dictionary) and must not do it
  again per row. Each entry names the loop header that opens its own walk,
  and a procedure whose header no longer matches FAILS rather than quietly
  passing with nothing pinned; the scan tracks For/Next depth, so a nested
  loop inside a tuple loop is covered too.

  Rule C - the load-bearing lines are still there, verbatim. A pin that only
  forbids things cannot notice a fix being deleted.

  Rule D - OPTIMIZE.3's integer grounder (VLA_Datalog.DatalogGroundRules
  and VLA_Relation's symbol table): its ROW loops hold no object at all -
  no VlaDict call, CreateObject, New, Array( or Collection - and each
  procedure has exactly the number of row loops its entry pins, so a loop
  renamed out of the pattern fails rather than going unwatched. The pinned
  count is not decoration: without it, this rule's own mutation control
  renamed one of IntJoinAtom's three row loops and passed. Mutation-
  controlled 2026-09-24 with five planted defects (a dictionary lookup in
  the probe loop, Array() in the comparison loop, a New Collection in the
  head loop, a Collection in the interning loop, a renamed row loop), all
  caught, and an unmutated copy clean.

WHAT IS NOT CHECKED: allocation per RULE PASS or per ROUND. EvalRuleBody's
colOf, RunFixpointForRules' newDeltas and RelNew's own index are one
dictionary per call, not per row, and DATALOG.12's arithmetic puts all of
them together at about a tenth of what one tuple loop used to cost. Nor
VLA_Sql.bas or VLA_Prolog.bas: neither has been measured on a per-row path,
and a pin nobody measured is a guess with a red light.

BASELINE: hand-maintained below, in check_vladict_guard.ps1's reviewable
shape. Adding a procedure is a deliberate, reviewed act.

NOT wired into VlaSelfTest, the same reasoning the other tools/check_*.ps1
scans give: a step a human runs, not a per-run gate.

Usage:  powershell -File tools\check_datalog_per_tuple_alloc.ps1 [-SrcDir <folder>]
Exit code: 0 if no listed procedure allocates per tuple and every pinned
line is present; 1 otherwise.
#>

param(
    [string]$SrcDir = ''
)

$ErrorActionPreference = 'Stop'

if ($SrcDir -eq '') {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $SrcDir   = Join-Path $repoRoot 'src'
}

# --- The pinned procedures, hand-maintained. ---
# 2026-09-20, DATALOG.14.
$ruleA = @(
    @{ Module = 'VLA_Datalog.bas'; Proc = 'PlanMatches' }
)
# Each entry names the loop header that opens its per-tuple walk, because
# not every one of them walks a Relation. ProjectAfterJoin iterates RelJoin's
# return value, which is a bare Collection of tuple arrays and NOT a Relation
# record - its own header says so, and a first run of this check that assumed
# RelTuples( everywhere reported that as a missing loop. Naming the header
# per procedure keeps the baseline hardcoded and reviewable rather than
# clever, and makes a rewritten loop fail loudly instead of silently
# un-pinning the procedure.
$tupleLoop = 'RelTuples\s*\('
$ruleB = @(
    @{ Module = 'VLA_Datalog.bas'; Proc = 'FilterAtomRelation';     Loop = $tupleLoop },
    @{ Module = 'VLA_Datalog.bas'; Proc = 'ComputeAggregateGroups'; Loop = $tupleLoop },
    @{ Module = 'VLA_Datalog.bas'; Proc = 'FilterOutMatching';      Loop = $tupleLoop },
    @{ Module = 'VLA_Datalog.bas'; Proc = 'ProjectAfterJoin';       Loop = 'For\s+Each\s+\w+\s+In\s+joined\b' },
    @{ Module = 'VLA_Relation.bas'; Proc = 'BuildJoinIndex';        Loop = $tupleLoop },
    @{ Module = 'VLA_Relation.bas'; Proc = 'RelJoin';               Loop = $tupleLoop }
)
$ruleC = @(
    @{ Module = 'VLA_Datalog.bas'; Proc = 'ParseAtom'
       Line = 'atom.Add BuildAtomPlan(args)'
       Why  = 'every atom must carry its match plan, or the filter has nothing to read' },
    @{ Module = 'VLA_Datalog.bas'; Proc = 'FilterAtomRelation'
       Line = 'If Not AtomPlanConstrains(atom) Then'
       Why  = 'an atom that constrains nothing must get its relation back whole, not copied' },
    @{ Module = 'VLA_Datalog.bas'; Proc = 'FilterAtomRelation'
       Line = 'PlanUnpack AtomPlan(atom), used, pos, cmp, txt'
       Why  = 'the plan''s arrays must be unpacked ONCE, outside the tuple loop - unpacking per tuple copies three Variant arrays a row, a smaller version of the cost this item removed' },
    @{ Module = 'VLA_Datalog.bas'; Proc = 'ComputeAggregateGroups'
       Line = 'PlanUnpack tuplePlan, used, pos, cmp, txt'
       Why  = 'the same, on the aggregate path' },
    @{ Module = 'VLA_Datalog.bas'; Proc = 'ComputeAggregateGroups'
       Line = 'Set tuplePlan = BuildAggregatePlan(atom, colOf, kind, valuePos)'
       Why  = 'the aggregate plan must be built once per call, outside the tuple loop' },
    @{ Module = 'VLA_Datalog.bas'; Proc = 'DatalogRunForms'
       Line = 'PushBoundArguments rules, relations, queryName, queryAtom'
       Why  = 'the question''s bound argument must still be pushed in before the fixpoint' }
)

# --- Rule D, added 2026-09-24 by OPTIMIZE.3's slice 1: the INTEGER GROUNDER.
# VLA_Datalog.DatalogGroundRules evaluates OPTIMIZE's rules over Long arrays
# (VLA_Relation's VlaSymbols ids) so that a produced row costs a few array
# writes instead of the ~7.6 us a Variant row, a string key and a dictionary
# entry cost EvalRuleBody. That is the whole of what it is for, and losing it
# is SILENT in the same way rule B's regression is: put a dictionary lookup,
# a Collection or an Array() back inside a row loop and every answer is
# still right - the pins that run each rule both ways still agree - and only
# a timing on a real-sized pool would say. So in these procedures a ROW LOOP
# may hold no object at all, which is a wider pattern than rule B's: it also
# forbids a VlaDict lookup, New and Array(. The per-atom PLAN loops (For p =
# 1 To k), which do read colOf, are deliberately not named here - they run
# once per body atom, not once per row.
# OPTIMIZE.3 slice 3 added the two procedures that COUNT rows without making
# them - IntJoinCount (a join's exact size, from the keys: the smaller side's
# rows, then the larger's) and IntAtomMatchCount (an atom's own matches, for
# the planned order) - and they are the same hot loops over the same rows.
# PlanBodyItems loops over a body's ITEMS, not its rows, and is not named.
$ruleD = @(
    @{ Module = 'VLA_Datalog.bas';  Proc = 'IntJoinAtom';       Loop = 'For\s+(ri|e|pr)\s*='; Count = 3 },
    @{ Module = 'VLA_Datalog.bas';  Proc = 'IntJoinCount';      Loop = 'For\s+(e|pr)\s*='; Count = 2 },
    @{ Module = 'VLA_Datalog.bas';  Proc = 'IntAtomMatchCount'; Loop = 'For\s+ri\s*='; Count = 1 },
    @{ Module = 'VLA_Datalog.bas';  Proc = 'IntAntiJoinAtom';   Loop = 'For\s+(ri|r)\s*='; Count = 2 },
    @{ Module = 'VLA_Datalog.bas';  Proc = 'IntCompareFilter';  Loop = 'For\s+r\s*='; Count = 1 },
    @{ Module = 'VLA_Datalog.bas';  Proc = 'IntTextFilter';     Loop = 'For\s+r\s*='; Count = 1 },
    @{ Module = 'VLA_Datalog.bas';  Proc = 'GroundRuleInteger'; Loop = 'For\s+r\s*=\s*1\s+To\s+accN'; Count = 1 },
    @{ Module = 'VLA_Relation.bas'; Proc = 'VlaSymInternRelation'; Loop = $tupleLoop; Count = 1 }
)
$objectPattern = '\b(VlaDict\w*\s*\(|CreateObject\s*\(|New\s+\w|Array\s*\(|Collection\b)'

$allocPattern = '\b(VlaDictNew\s*\(|CreateObject\s*\()'
$endPattern   = '^\s*End\s+(Sub|Function|Property)\b'

function Test-IsComment([string]$line) {
    $t = $line.TrimStart()
    return ($t.Length -eq 0 -or $t.StartsWith("'"))
}

function Get-ProcBody([string[]]$lines, [string]$proc) {
    $startPattern = '^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function)\s+' + [regex]::Escape($proc) + '\s*\('
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match $startPattern) { $start = $i; break }
    }
    if ($start -lt 0) { return $null }
    $body = New-Object System.Collections.Generic.List[object]
    for ($i = $start; $i -lt $lines.Count; $i++) {
        $body.Add([pscustomobject]@{ No = $i + 1; Text = $lines[$i] })
        if ($i -gt $start -and $lines[$i] -match $endPattern) { break }
    }
    return ,$body
}

$failed = New-Object System.Collections.Generic.List[string]
$cache  = @{}

function Get-Lines([string]$module) {
    if (-not $cache.ContainsKey($module)) {
        $path = Join-Path $SrcDir $module
        if (-not (Test-Path -LiteralPath $path)) {
            $failed.Add("${module}: file not found at $path")
            return $null
        }
        $cache[$module] = [System.IO.File]::ReadAllLines($path)
    }
    return ,$cache[$module]
}

Write-Output '=== PER-TUPLE ALLOCATION PIN (DATALOG.14) ==='
Write-Output ''
Write-Output '--- rule A: these run once per tuple entirely, so they may not allocate at all ---'

foreach ($g in $ruleA) {
    $lines = Get-Lines $g.Module
    if ($null -eq $lines) { continue }
    $body = Get-ProcBody $lines $g.Proc
    $label = "$($g.Module)::$($g.Proc)"
    if ($null -eq $body) {
        $failed.Add("${label}: procedure not found - renamed or removed? Update this check's list deliberately.")
        Write-Output "  FAIL  $label"
        continue
    }
    $hits = New-Object System.Collections.Generic.List[string]
    foreach ($l in $body) {
        if (Test-IsComment $l.Text) { continue }
        if ($l.Text -match $allocPattern) { $hits.Add("line $($l.No): $($l.Text.Trim())") }
    }
    if ($hits.Count -gt 0) {
        foreach ($h in $hits) { $failed.Add("${label}: allocates on a per-tuple path - $h") }
        Write-Output "  FAIL  $label"
    } else {
        Write-Output "  ok    $label"
    }
}

Write-Output ''
Write-Output '--- rule B: these may allocate once per call, never inside a tuple loop ---'

foreach ($g in $ruleB) {
    $lines = Get-Lines $g.Module
    if ($null -eq $lines) { continue }
    $body = Get-ProcBody $lines $g.Proc
    $label = "$($g.Module)::$($g.Proc)"
    if ($null -eq $body) {
        $failed.Add("${label}: procedure not found - renamed or removed? Update this check's list deliberately.")
        Write-Output "  FAIL  $label"
        continue
    }
    $depth = 0
    $tupleDepth = -1
    $loops = 0
    $hits = New-Object System.Collections.Generic.List[string]
    foreach ($l in $body) {
        if (Test-IsComment $l.Text) { continue }
        $t = $l.Text
        if ($t -match '^\s*For\s+Each\b' -or $t -match '^\s*For\s+\w+\s*=') {
            $depth++
            if ($tupleDepth -lt 0 -and $t -match $g.Loop) {
                $tupleDepth = $depth
                $loops++
            }
        } elseif ($t -match '^\s*Next\b') {
            if ($tupleDepth -eq $depth) { $tupleDepth = -1 }
            $depth--
        } elseif ($tupleDepth -ge 0 -and $t -match $allocPattern) {
            $hits.Add("line $($l.No): $($t.Trim())")
        }
    }
    if ($hits.Count -gt 0) {
        foreach ($h in $hits) { $failed.Add("${label}: allocates inside a tuple loop - $h") }
        Write-Output "  FAIL  $label"
    } elseif ($loops -eq 0) {
        $failed.Add("${label}: no loop matching its pinned header found. This pin is watching a loop that no longer exists - if the procedure was rewritten, re-aim the check deliberately.")
        Write-Output "  FAIL  $label"
    } else {
        Write-Output "  ok    $label  ($loops tuple loop(s))"
    }
}

Write-Output ''
Write-Output '--- rule C: the load-bearing lines are still present ---'

foreach ($g in $ruleC) {
    $lines = Get-Lines $g.Module
    if ($null -eq $lines) { continue }
    $body = Get-ProcBody $lines $g.Proc
    $label = "$($g.Module)::$($g.Proc)"
    if ($null -eq $body) {
        $failed.Add("${label}: procedure not found - renamed or removed? Update this check's list deliberately.")
        Write-Output "  FAIL  $label"
        continue
    }
    $found = $false
    foreach ($l in $body) {
        if (Test-IsComment $l.Text) { continue }
        if ($l.Text.Contains($g.Line)) { $found = $true; break }
    }
    if ($found) {
        Write-Output "  ok    $label  ($($g.Line))"
    } else {
        $failed.Add("${label}: the pinned line '$($g.Line)' is gone - $($g.Why)")
        Write-Output "  FAIL  $label"
    }
}

Write-Output ''
Write-Output '--- rule D: the integer grounder''s row loops hold no object at all (OPTIMIZE.3) ---'

foreach ($g in $ruleD) {
    $lines = Get-Lines $g.Module
    if ($null -eq $lines) { continue }
    $body = Get-ProcBody $lines $g.Proc
    $label = "$($g.Module)::$($g.Proc)"
    if ($null -eq $body) {
        $failed.Add("${label}: procedure not found - renamed or removed? Update this check's list deliberately.")
        Write-Output "  FAIL  $label"
        continue
    }
    $depth = 0
    $rowDepth = -1
    $loops = 0
    $hits = New-Object System.Collections.Generic.List[string]
    foreach ($l in $body) {
        if (Test-IsComment $l.Text) { continue }
        $t = $l.Text
        if ($t -match '^\s*For\s+Each\b' -or $t -match '^\s*For\s+\w+\s*=') {
            $depth++
            if ($rowDepth -lt 0 -and $t -match $g.Loop) {
                $rowDepth = $depth
                $loops++
            }
        } elseif ($t -match '^\s*Next\b') {
            if ($rowDepth -eq $depth) { $rowDepth = -1 }
            $depth--
        } elseif ($rowDepth -ge 0 -and $t -match $objectPattern) {
            $hits.Add("line $($l.No): $($t.Trim())")
        }
    }
    if ($hits.Count -gt 0) {
        foreach ($h in $hits) { $failed.Add("${label}: an object inside a row loop - $h") }
        Write-Output "  FAIL  $label"
    } elseif ($loops -ne $g.Count) {
        # The COUNT is pinned, not merely "at least one": IntJoinAtom has
        # three row loops, and this check's own mutation control renamed one
        # of them out of the pattern and passed, because the other two still
        # matched - the renamed loop was then watched by nothing.
        $failed.Add("${label}: $loops row loop(s) match the pinned header where $($g.Count) should. A loop was renamed, added or removed - if the procedure was rewritten, re-aim the check deliberately.")
        Write-Output "  FAIL  $label"
    } else {
        Write-Output "  ok    $label  ($loops row loop(s))"
    }
}

Write-Output ''
if ($failed.Count -gt 0) {
    Write-Output '=== CHECK: FAILED ==='
    foreach ($f in $failed) { Write-Output "  $f" }
    exit 1
}
$total = $ruleA.Count + $ruleB.Count + $ruleC.Count + $ruleD.Count
Write-Output "=== CHECK: clean - $($ruleA.Count) allocation-free procedure(s), $($ruleB.Count) with no allocation inside a tuple loop, $($ruleC.Count) pinned line(s) present, $($ruleD.Count) integer procedure(s) with object-free row loops ($total rules) ==="
exit 0
