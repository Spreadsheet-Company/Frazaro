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
if ($failed.Count -gt 0) {
    Write-Output '=== CHECK: FAILED ==='
    foreach ($f in $failed) { Write-Output "  $f" }
    exit 1
}
$total = $ruleA.Count + $ruleB.Count + $ruleC.Count
Write-Output "=== CHECK: clean - $($ruleA.Count) allocation-free procedure(s), $($ruleB.Count) with no allocation inside a tuple loop, $($ruleC.Count) pinned line(s) present ($total rules) ==="
exit 0
