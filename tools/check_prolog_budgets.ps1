<#
check_prolog_budgets.ps1 - PROLOG.28's mechanical check.

WHY THIS EXISTS. PROLOG.28 split one ceiling into the two budgets it had
been standing for. DEPTH is the nesting of SolveGoalList frames - the
native stack the old 120-step ceiling was introduced to protect - and WORK
is every candidate tried, which grows with the Tables. It also made the
term walkers loop on a term's LAST position instead of recursing into it,
so that a list as long as work allows no longer nests one frame per
element. Every one of those properties can be undone by an ordinary-looking
edit that no test will catch on a green run:

  - a new `Exit Sub` in SolveGoalList skips the count-out at its single
    exit, so mDepth drifts upward and every later frame of the same query
    is refused as too deep - a confidently wrong refusal, and only on the
    queries that happen to take that exit;
  - a new candidate loop that charges work but checks the old ceiling, or
    checks none, or checks depth, gives a runaway or a big Table the wrong
    refusal or none;
  - a generator's up-front bound written as `> PROLOG_MAX_WORK` instead of
    `> (PROLOG_MAX_WORK - 1)` lets a range of exactly the budget pass and
    die at the work ceiling a value later, naming the budget rather than
    the range (PROLOG.9's off-by-one, now on the new budget);
  - a walker "simplified" back to `For i = 1 To lst.Count` with a
    recursive call inside recurses into the tail again, and the stack runs
    out only on the day someone gathers a real Table into one list.

WHAT IT CHECKS, eight rules, each read from the source under test:

  A. The retired names are gone: no `PROLOG_MAX_STEPS` outside a comment
     in any module, and no `prolog-step-ceiling` raised or catalogued.
  B. SolveGoalList counts depth in as its first statement, refuses past
     PROLOG_MAX_DEPTH with prolog-depth-ceiling, has NO `Exit Sub` at all,
     exactly one `leave:` label, and counts out on the line after it,
     immediately before `End Sub`.
  C. mDepth is assigned in exactly two places: SolveGoalList (in and out)
     and PrologRun (zeroed, before anything is solved).
  D. Every `stepsTaken = stepsTaken + 1` is followed, on its next code
     line, by the work check and nothing else - the one refusal a unit of
     work can reach.
  E. Every comparison against PROLOG_MAX_WORK is either the work check
     itself (`stepsTaken > PROLOG_MAX_WORK`) or an up-front generator bound
     (`> (PROLOG_MAX_WORK - 1)`).
  F. Each of the nine walkers named below contains a `Do` loop, and never
     calls itself from inside a `For` loop that runs to a `.Count` with no
     `- 1` - the one shape that recurses into the last position.
  G. The two new refusals and the cell-length refusal are catalogued in
     VLA_Messages.bas with the slots the raise sites fill.

The walker list is the reviewable baseline: a walker that starts seeing
long lists must be added here on purpose, and a rename breaks the run
loudly rather than silently checking nothing.

House style (tools/check_*.ps1): PowerShell 5.1, host-independent, a
hard-coded reviewable baseline, no live Excel needed. Never wired into
VlaSelfTest.
#>
param([string]$Root = (Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference = 'Stop'

$prologPath   = Join-Path $Root 'src\VLA_Prolog.bas'
$unifyPath    = Join-Path $Root 'src\VLA_Unify.bas'
$vlaPath      = Join-Path $Root 'src\VLA.bas'
$messagesPath = Join-Path $Root 'src\VLA_Messages.bas'
$walkers = @(
    @{ File = $prologPath; Name = 'FreshenTerm' },
    @{ File = $prologPath; Name = 'ResolveTermDeep' },
    @{ File = $prologPath; Name = 'TermIsGroundDeep' },
    @{ File = $prologPath; Name = 'ContractListsInto' },
    @{ File = $prologPath; Name = 'QuotedVersusBareClass' },
    @{ File = $unifyPath;  Name = 'UnifyTwoWay' },
    @{ File = $unifyPath;  Name = 'TermsIdentical' },
    @{ File = $unifyPath;  Name = 'EnvOccurs' },
    @{ File = $vlaPath;    Name = 'WriteDatum' }
)

$failures = New-Object System.Collections.Generic.List[string]
function Fail([string]$rule, [string]$what) { $script:failures.Add("$rule  $what"); Write-Output "  FAIL $rule  $what" }
function IsComment([string]$l) { return ($l -match "^\s*'") }
function CodeOf([string]$l) {
    # the line with any trailing comment removed - a ' inside a "string" does
    # not start one, so strings are skipped first
    $out = New-Object System.Text.StringBuilder; $inStr = $false
    foreach ($ch in $l.ToCharArray()) {
        if ($ch -eq '"') { $inStr = -not $inStr }
        if (-not $inStr -and $ch -eq "'") { break }
        [void]$out.Append($ch)
    }
    return $out.ToString()
}
function Get-Proc([string[]]$lines, [string]$name) {
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match ('^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function)\s+' + [regex]::Escape($name) + '\s*\(')) { $start = $i; break }
    }
    if ($start -lt 0) { return $null }
    for ($j = $start + 1; $j -lt $lines.Count; $j++) {
        if ($lines[$j] -match '^\s*End\s+(Sub|Function)\s*$') { return [pscustomobject]@{ Start = $start; End = $j } }
    }
    return $null
}
function NextCode([string[]]$lines, [int]$from) {
    for ($k = $from; $k -lt $lines.Count; $k++) { if (-not (IsComment $lines[$k]) -and $lines[$k].Trim().Length -gt 0) { return $k } }
    return -1
}

$pl = Get-Content -LiteralPath $prologPath
$ml = Get-Content -LiteralPath $messagesPath

# ---- A. the retired names --------------------------------------------------
Write-Output 'A. the retired ceiling is gone'
foreach ($f in (Get-ChildItem -LiteralPath (Join-Path $Root 'src') -Filter '*.bas')) {
    $n = 0
    foreach ($l in (Get-Content -LiteralPath $f.FullName)) {
        if (IsComment $l) { continue }
        $c = CodeOf $l
        if ($c -match '\bPROLOG_MAX_STEPS\b') { Fail 'A' "$($f.Name): PROLOG_MAX_STEPS in code: $($l.Trim())" }
        if ($c -match '"prolog-step-ceiling"') { Fail 'A' "$($f.Name): prolog-step-ceiling raised or catalogued: $($l.Trim())" }
    }
}

# ---- B. SolveGoalList's one way in and one way out -------------------------
Write-Output 'B. SolveGoalList counts depth in once and out once'
$sgl = Get-Proc $pl 'SolveGoalList'
if ($null -eq $sgl) { Fail 'B' 'SolveGoalList not found' } else {
    # the signature may continue over several lines; the body starts after the one ending in ")"
    $bodyStart = $sgl.Start
    while ($bodyStart -lt $sgl.End -and (CodeOf $pl[$bodyStart]).TrimEnd() -match '_$') { $bodyStart++ }
    $first = NextCode $pl ($bodyStart + 1)
    if ($pl[$first].Trim() -ne 'mDepth = mDepth + 1') { Fail 'B' "first statement is not the count-in: $($pl[$first].Trim())" }
    $second = NextCode $pl ($first + 1)
    if ($pl[$second].Trim() -ne 'If mDepth > PROLOG_MAX_DEPTH Then VLA_Messages.RaiseMsg "prolog-depth-ceiling", "depth", PROLOG_MAX_DEPTH') { Fail 'B' "second statement is not the depth check: $($pl[$second].Trim())" }
    $labels = 0; $exits = 0
    for ($i = $sgl.Start; $i -le $sgl.End; $i++) {
        if (IsComment $pl[$i]) { continue }
        $c = CodeOf $pl[$i]
        if ($c -match '\bExit\s+Sub\b') { $exits++; Fail 'B' ("line {0}: Exit Sub skips the count-out: {1}" -f ($i + 1), $pl[$i].Trim()) }
        if ($c -match '^\s*leave:\s*$') { $labels++ }
    }
    if ($labels -ne 1) { Fail 'B' "expected exactly one leave: label, found $labels" }
    if ($pl[$sgl.End - 2].Trim() -ne 'leave:' -or $pl[$sgl.End - 1].Trim() -ne 'mDepth = mDepth - 1') { Fail 'B' 'the last two lines before End Sub are not "leave:" and the count-out' }
    if ($exits -eq 0 -and $labels -eq 1) { Write-Output '  ok  count-in first, one exit, count-out last' }
}

# ---- C. who may touch mDepth -----------------------------------------------
Write-Output 'C. mDepth is assigned only by SolveGoalList and PrologRun'
$pr = Get-Proc $pl 'PrologRun'
$zeroed = $false
for ($i = 0; $i -lt $pl.Count; $i++) {
    if (IsComment $pl[$i]) { continue }
    $c = CodeOf $pl[$i]
    if ($c -notmatch '^\s*mDepth\s*=') { continue }
    $inSgl = ($i -gt $sgl.Start -and $i -lt $sgl.End)
    $inPr = ($null -ne $pr -and $i -gt $pr.Start -and $i -lt $pr.End)
    if ($inPr -and $c.Trim() -eq 'mDepth = 0') { $zeroed = $true; continue }
    if (-not $inSgl) { Fail 'C' ("line {0}: mDepth assigned outside SolveGoalList: {1}" -f ($i + 1), $pl[$i].Trim()) }
}
if (-not $zeroed) { Fail 'C' 'PrologRun never zeroes mDepth' }
else {
    # zeroed before the solver is first called
    $zeroAt = -1; $solveAt = -1
    for ($i = $pr.Start; $i -le $pr.End; $i++) {
        $c = CodeOf $pl[$i]
        if ($zeroAt -lt 0 -and $c.Trim() -eq 'mDepth = 0') { $zeroAt = $i }
        if ($solveAt -lt 0 -and $c -match '^\s*(ParseProgram|SolveGoalList)\b') { $solveAt = $i }
    }
    if ($zeroAt -gt $solveAt) { Fail 'C' 'PrologRun zeroes mDepth only after parsing or solving has begun' } else { Write-Output '  ok  zeroed first in PrologRun, counted only in SolveGoalList' }
}

# ---- D. every unit of work reaches the work check ---------------------------
Write-Output 'D. every charge of work is followed by the work check'
$workCheck = 'If stepsTaken > PROLOG_MAX_WORK Then VLA_Messages.RaiseMsg "prolog-work-ceiling", "work", PROLOG_MAX_WORK'
$charges = 0
for ($i = 0; $i -lt $pl.Count; $i++) {
    if (IsComment $pl[$i]) { continue }
    if ((CodeOf $pl[$i]).Trim() -ne 'stepsTaken = stepsTaken + 1') { continue }
    $charges++
    $n = NextCode $pl ($i + 1)
    if ($pl[$n].Trim() -ne $workCheck) { Fail 'D' ("line {0}: a charge of work not followed by the work check: next is {1}" -f ($i + 1), $pl[$n].Trim()) }
}
if ($charges -eq 0) { Fail 'D' 'no charge of work found at all - the scan is reading the wrong file' } else { Write-Output "  ok  $charges charge(s), each checked" }

# ---- E. every comparison against the budget ---------------------------------
Write-Output 'E. every comparison against PROLOG_MAX_WORK is the check or an up-front bound'
$bounds = 0
for ($i = 0; $i -lt $pl.Count; $i++) {
    if (IsComment $pl[$i]) { continue }
    $c = CodeOf $pl[$i]
    if ($c -notmatch '\bPROLOG_MAX_WORK\b') { continue }
    if ($c -match 'Const\s+PROLOG_MAX_WORK') { continue }
    $rest = $c
    $rest = $rest.Replace('stepsTaken > PROLOG_MAX_WORK', '')
    $rest = $rest.Replace('"work", PROLOG_MAX_WORK', '')
    $rest = $rest.Replace('"max", PROLOG_MAX_WORK', '')
    if ($rest -match '>\s*\(PROLOG_MAX_WORK - 1\)') { $bounds++; $rest = $rest -replace '>\s*\(PROLOG_MAX_WORK - 1\)', '' }
    if ($rest -match '\bPROLOG_MAX_WORK\b') { Fail 'E' ("line {0}: PROLOG_MAX_WORK used another way: {1}" -f ($i + 1), $pl[$i].Trim()) }
}
Write-Output "  ok-or-listed  $bounds up-front generator bound(s) at PROLOG_MAX_WORK - 1"

# ---- F. the walkers loop on the last position --------------------------------
Write-Output 'F. the nine walkers loop on the last position'
$cache = @{}
foreach ($w in $walkers) {
    if (-not $cache.ContainsKey($w.File)) { $cache[$w.File] = Get-Content -LiteralPath $w.File }
    $lines = $cache[$w.File]
    $pr2 = Get-Proc $lines $w.Name
    if ($null -eq $pr2) { Fail 'F' "$($w.Name): not found in $(Split-Path -Leaf $w.File)"; continue }
    $hasDo = $false; $bad = 0
    $forStack = New-Object System.Collections.ArrayList     # $true = a For over every position
    for ($i = $pr2.Start + 1; $i -lt $pr2.End; $i++) {
        if (IsComment $lines[$i]) { continue }
        $c = (CodeOf $lines[$i]).Trim()
        if ($c -match '^Do(\s+While\b.*)?$') { $hasDo = $true }
        if ($c -match '^For\s+\w+\s*=\s*.+\s+To\s+(.+)$') {
            $bound = $Matches[1].Trim()
            [void]$forStack.Add(($bound -match '\.Count\s*$'))
            continue
        }
        if ($c -match '^Next\b') { if ($forStack.Count -gt 0) { $forStack.RemoveAt($forStack.Count - 1) }; continue }
        if ($c -match ('\b' + [regex]::Escape($w.Name) + '\b') -and $c -notmatch ('^' + [regex]::Escape($w.Name) + '\s*=')) {
            if ($forStack -contains $true) { $bad++; Fail 'F' ("{0} line {1}: calls itself inside a For over every position: {2}" -f $w.Name, ($i + 1), $c) }
        }
    }
    if (-not $hasDo) { Fail 'F' "$($w.Name): no Do loop - the last position is not walked iteratively" }
    if ($hasDo -and $bad -eq 0) { Write-Output ("  ok  {0}" -f $w.Name) }
}

# ---- H. every list builder is held to the list budget -----------------------
# A list is a chain of cons cells and VBA releases such a chain recursively,
# so a builder that forgets the budget can hand back a value that crashes
# Excel when it is dropped - which is what happened live before this budget
# existed. The four builders are the reviewable baseline.
Write-Output 'H. every list builder checks PROLOG_MAX_LIST and refuses by name'
$listSites = 0
for ($i = 0; $i -lt $pl.Count; $i++) {
    if (IsComment $pl[$i]) { continue }
    $c2 = CodeOf $pl[$i]
    if ($c2 -notmatch '>\s*PROLOG_MAX_LIST\b') { continue }
    $listSites++
    $n2 = NextCode $pl ($i + 1)
    if ($pl[$n2] -notmatch 'RaiseMsg "prolog-list-ceiling"') { Fail 'H' ("line {0}: a list-budget check that does not refuse by name: {1}" -f ($i + 1), $pl[$n2].Trim()) }
}
if ($listSites -lt 4) { Fail 'H' "expected the four list builders (findall, a text split, append's join, a written-out list) to check the budget; found $listSites" }
else { Write-Output "  ok  $listSites builder(s), each refusing by name" }

# ---- G. the refusals exist with their slots ---------------------------------
Write-Output 'G. the refusals are catalogued with the slots their raise sites fill'
$need = @{ 'prolog-depth-ceiling' = @('{depth}'); 'prolog-work-ceiling' = @('{work}'); 'prolog-value-too-long-for-a-cell' = @('{var}', '{length}'); 'prolog-list-ceiling' = @('{form}', '{count}', '{max}') }
foreach ($id in $need.Keys) {
    $line = $ml | Where-Object { $_ -match ('AddMsg m, "' + [regex]::Escape($id) + '"') } | Select-Object -First 1
    if (-not $line) { Fail 'G' "$id is not catalogued"; continue }
    foreach ($slot in $need[$id]) { if ($line -notmatch [regex]::Escape($slot)) { Fail 'G' "$id has no $slot slot" } }
}
if (-not ($failures | Where-Object { $_ -like 'G*' })) { Write-Output '  ok  four refusals, every slot present' }

Write-Output ''
if ($failures.Count -eq 0) {
    Write-Output '=== CHECK: clean - depth counted once in and once out, every unit of work checked, and every walker loops on the last position ==='
} else {
    Write-Output "=== CHECK: $($failures.Count) problem(s) with PROLOG's budgets ==="
    $failures | ForEach-Object { Write-Output "  $_" }
    Write-Output ''
    Write-Output 'PROLOG.28 split depth from work and made the term walkers loop on a list''s tail. Each rule above guards one edit that would quietly undo that: see this script''s header.'
}
exit ([Math]::Min($failures.Count, 1))
