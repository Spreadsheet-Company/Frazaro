<#
check_prolog_reserved_names.ps1 - PROLOG.8's mechanical check.

WHY THIS EXISTS. PROLOG reserves predicate names by FORWARD DECLARATION:
a name is refused as a user-defined predicate before the goal that gives
it meaning is built, so a knowledge base written today can never be
silently broken when the goal ships. That discipline creates three
separate places one set of names has to be written down, and every one of
them can drift from the others independently:

  1. IsReservedPredicateName  - what the parser REFUSES as a predicate
                                name in a (fact ...) or (rule ...);
  2. SolveGoalList            - what the solver actually DISPATCHES;
  3. prolog-reserved-predicate-name's own catalogue text - what the user
                                is TOLD the reserved set is.

Each pairwise disagreement is its own distinct, quiet defect:

  reserved but not dispatched  - the name is forbidden AND does nothing;
  dispatched but not reserved  - a user can define a predicate that the
                                 solver then shadows, and their own facts
                                 are silently unreachable;
  reserved but not enumerated  - the refusal lists a set that is not the
                                 set it just refused from, which is the
                                 confidently-wrong-answer class this
                                 project holds to be worse than a crash
                                 (IN.15's own finding).

WHY A SCRIPT AND NOT A HAND COUNT. PROLOG.8's own roadmap entry says the
dispatch is a three-arm chain "IsReservedPredicateName already reserves
four names for". Both halves are stale as of PROLOG.7: there are four
predName-keyed arms (is/not/findall plus the comparison table) and TEN
reserved names, since PROLOG.7 delegated six more from Case Else. The
entry for PROLOG.7 was wrong in the same way about its own message count
(two claimed, five actual), and the check written BEFORE that fix is the
only reason it was caught. This one is written first too.

WHAT IT CHECKS.

  A. Set agreement. The reserved set is assembled from the code alone -
     IsReservedPredicateName's own literal Case list, plus every table
     function it delegates to from Case Else. Every name in that set must
     appear as a whole token in prolog-reserved-predicate-name's text,
     and every operator-shaped token in that text must be in the set. Both
     directions, so neither adding a name without telling the user nor
     removing one and leaving it advertised can pass.

  B. Count words. The refusal says "the six comparisons". A count word
     standing in front of a group noun is a hand count embedded in prose,
     and goes stale exactly the way the roadmap entry above did. Each one
     named in $countPhrases below is checked against the real size of the
     table it describes.

  C. Dispatch reachability. Every reserved name must actually be reached
     by SolveGoalList - a literal name through its own `predName = "..."`
     arm, a table through its own `<Table>(predName) <> ""` arm, or, for
     the one name dispatched structurally rather than by predName, through
     the baseline exemption recorded in $structuralDispatch below.

  D. Dispatch legitimacy - rule C read the other way round. Every
     predName-keyed arm SolveGoalList actually HAS must belong to the
     reserved set: a `predName = "..."` arm whose name no code path
     reserves, or a `<Table>(predName) <> ""` arm over a table
     IsReservedPredicateName does not delegate to, is the second of the
     three defects named above - dispatched but not reserved, so a user
     may DEFINE that predicate, have their definition silently shadowed
     by the solver, and never be told.

     PROLOG.9 added this. Rules A-C as PROLOG.8 shipped them walked the
     reserved set outward to the other two places and never walked the
     dispatch back, so the direction was unguarded despite the header
     above having always claimed it: a `predName = "bogusundeclared"` arm
     spliced into SolveGoalList passed the whole script clean, exit 0,
     verified by mutation before this rule was written. Nothing in the
     shipped code was wrong - every arm was reserved - but PROLOG.9 adds
     two arms at once (the six type tests through their own table, and
     `between` as a literal name), which is exactly the change that would
     have walked through the gap.

The baselines below are lists of NAMES, not patterns: a table that joins
the reserved set must be added here on purpose, and a rename breaks the
run loudly rather than silently scanning nothing.

House style (tools/check_*.ps1): PowerShell 5.1, host-independent, a
hard-coded reviewable baseline, no live Excel needed. Never wired into
VlaSelfTest - it reads source text, not runtime behaviour.

Usage:  powershell -File tools\check_prolog_reserved_names.ps1
Exit code: 0 clean, 1 if the three places disagree about the reserved set.
#>

$ErrorActionPreference = 'Stop'

$repoRoot     = Split-Path -Parent $PSScriptRoot
$prologPath   = Join-Path $repoRoot 'src\VLA_Prolog.bas'
$messagesPath = Join-Path $repoRoot 'src\VLA_Messages.bas'

$reservedMsgId = 'prolog-reserved-predicate-name'

# ---- baseline: the one name dispatched structurally, not by predName ----
# Cut is the ONLY non-object goal ValidateBodyItem can let through to
# SolveGoalList, so its arm must run BEFORE GoalPredName (which assumes a
# Collection) and therefore cannot be a `predName = "!"` test at all. It is
# exempt from rule C by name, with the reason recorded, rather than by a
# pattern loose enough to excuse a genuinely undispatched name.
$structuralDispatch = @{
    '!' = "SolveGoalList's own non-object arm, which must precede GoalPredName"
}

# ---- baseline: count words in the refusal, and the table each describes -
# Key is the group noun as it appears in the text; value is the table
# function whose size the preceding count word must equal.
# NOTE on phrasing: the count word must sit IMMEDIATELY before the group
# noun, because rule B matches '<word> <noun>'. "the six ISO spellings"
# is checked; "their six bare ISO spellings" would find "bare" in front
# of the noun, report "not a count word", and silently check nothing.
$countPhrases = @{
    'comparisons'          = 'ComparisonOpFor'
    'term-matching goals'  = 'UnificationOpFor'
    'type tests'           = 'TypeTestKindFor'
    'ISO spellings'        = 'TypeTestIsoSpellingFor'
    'list goals'           = 'ListGoalKindFor'
}

$numberWords = @{
    'one' = 1; 'two' = 2; 'three' = 3; 'four' = 4; 'five' = 5; 'six' = 6
    'seven' = 7; 'eight' = 8; 'nine' = 9; 'ten' = 10; 'eleven' = 11; 'twelve' = 12
}

# A token made only of the characters Prolog operator names are spelled
# from. Used for the message -> code direction of rule A: it is what lets
# the scan tell an advertised operator apart from ordinary prose.
$operatorShaped = '^[<>=\\!:]+$'

$failures = New-Object System.Collections.Generic.List[string]

foreach ($p in @($prologPath, $messagesPath)) {
    if (-not (Test-Path -LiteralPath $p)) {
        Write-Output "FAIL: missing source file $p"
        exit 1
    }
}

# ---- procedure body extraction ------------------------------------------
# Whole-line VBA comments are dropped: these procedures carry heavy prose,
# and a name quoted inside a comment is documentation, not a Case arm.
function Get-ProcBody {
    param([string[]]$Lines, [string]$Name)

    $startPattern = '^\s*(Private\s+|Public\s+)?(Sub|Function)\s+' + [regex]::Escape($Name) + '\s*\('
    $start = -1
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Lines[$i] -match $startPattern) { $start = $i; break }
    }
    if ($start -lt 0) { return $null }

    for ($j = $start + 1; $j -lt $Lines.Count; $j++) {
        if ($Lines[$j] -match '^\s*End\s+(Sub|Function)\s*$') {
            return ,@($Lines[$start..$j] | Where-Object { $_ -notmatch "^\s*'" })
        }
    }
    return $null
}

# Every quoted literal appearing in a Case arm's own SELECTOR. The trailing
# `: Name = "value"` of a single-line Case body is stripped first, so a
# table function written as `Case "=<":  ComparisonOpFor = "<="` yields the
# Prolog spelling it matches on and never the substrate spelling it
# returns. The strip requires an identifier and an `=` after the colon, so
# the colons inside a name like "=:=" are never mistaken for it.
function Get-CaseLiterals {
    param([string[]]$Body)

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($line in $Body) {
        if ($line -notmatch '^\s*Case\s') { continue }
        $selector = $line -replace ':\s*[A-Za-z_]\w*\s*=.*$', ''
        foreach ($m in [regex]::Matches($selector, '"([^"]*)"')) {
            $names.Add($m.Groups[1].Value)
        }
    }
    return ,@($names | Sort-Object -Unique)
}

$prologLines = Get-Content -LiteralPath $prologPath

Write-Output '=== PROLOG.8 - reserved set, dispatch and refusal text must agree ==='
Write-Output "Source:   $prologPath"
Write-Output "Messages: $messagesPath"
Write-Output ''

# ---- 1. assemble the reserved set from the code -------------------------
$reservedBody = Get-ProcBody -Lines $prologLines -Name 'IsReservedPredicateName'
if ($null -eq $reservedBody) {
    Write-Output 'FAIL: IsReservedPredicateName not found in VLA_Prolog.bas (renamed or removed - update this baseline deliberately)'
    exit 1
}

$literalNames = Get-CaseLiterals -Body $reservedBody

# Case Else delegates to one table function per group. The call shape is
# the single source of truth for WHICH tables take part, so it is read out
# of the code rather than listed here.
$tableNames = @([regex]::Matches(($reservedBody -join "`n"), '([A-Za-z_]\w*)\s*\(\s*predName\s*\)\s*<>\s*""') |
    ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)

Write-Output '--- the reserved set, assembled from the code ---'
Write-Output ("  IsReservedPredicateName literal Case arms : {0}" -f ($literalNames -join ' '))
Write-Output ("  delegated table function(s)              : {0}" -f (($tableNames -join ', ')))

$reserved   = New-Object System.Collections.Generic.List[string]
$sourceOf   = @{}
$tableSizes = @{}

foreach ($n in $literalNames) { $reserved.Add($n); $sourceOf[$n] = 'literal' }

if ($tableNames.Count -eq 0) {
    $failures.Add('IsReservedPredicateName delegates to no table function - the Case Else single-source pattern is gone, which this check exists to hold in place')
}

foreach ($t in $tableNames) {
    $tbody = Get-ProcBody -Lines $prologLines -Name $t
    if ($null -eq $tbody) {
        $failures.Add("table function '$t' is delegated to by IsReservedPredicateName but not found in VLA_Prolog.bas")
        Write-Output ("  {0,-22} NOT FOUND" -f $t)
        continue
    }
    $tnames = Get-CaseLiterals -Body $tbody
    $tableSizes[$t] = $tnames.Count
    Write-Output ("  {0,-22} {1,2} name(s): {2}" -f $t, $tnames.Count, ($tnames -join ' '))
    if ($tnames.Count -eq 0) {
        $failures.Add("table function '$t' contributes no names - it matched nothing, which is itself a failure")
    }
    foreach ($n in $tnames) {
        if ($reserved -contains $n) {
            $failures.Add("'$n' is contributed twice (once as $($sourceOf[$n]), once by $t) - a name must have exactly one home")
        } else {
            $reserved.Add($n)
            $sourceOf[$n] = $t
        }
    }
}

$reservedSorted = @($reserved | Sort-Object -Unique)
Write-Output ''
Write-Output ("  RESERVED SET: {0} name(s) - {1}" -f $reservedSorted.Count, ($reservedSorted -join ' '))

# ---- 2. the refusal's own text ------------------------------------------
$msgText = $null
foreach ($line in (Get-Content -LiteralPath $messagesPath)) {
    if ($line -match '^\s*AddMsg\s+m,\s*"([^"]+)"\s*,\s*[^,]+,\s*"[^"]*"\s*,\s*"(.*)"\s*$') {
        if ($Matches[1] -eq $reservedMsgId) { $msgText = $Matches[2] }
    }
}

Write-Output ''
Write-Output "--- rule A: $reservedMsgId must enumerate exactly that set ---"

if ($null -eq $msgText) {
    $failures.Add("$reservedMsgId is not defined in VLA_Messages.bas")
    Write-Output '  UNDEFINED'
} else {
    # Tokens are split on whitespace and on the "/" the text uses to run
    # short names together ("is/not/findall/!"); surrounding prose
    # punctuation is trimmed from each end. A name is advertised only if it
    # survives as a WHOLE token - a bare Contains() would let "<" pass on
    # the strength of any "=<" already in the sentence.
    $rawTokens = $msgText -split '[\s/]+'
    $tokens = New-Object System.Collections.Generic.List[string]
    foreach ($t in $rawTokens) {
        $trimmed = $t.Trim(",.;()'`"")
        if ($trimmed.Length -gt 0) { $tokens.Add($trimmed) }
    }

    foreach ($n in $reservedSorted) {
        if ($tokens -contains $n) {
            Write-Output ("  {0,-6} ok         advertised (from {1})" -f $n, $sourceOf[$n])
        } else {
            Write-Output ("  {0,-6} MISSING    reserved (from {1}) but never named in the refusal" -f $n, $sourceOf[$n])
            $failures.Add("'$n' is reserved but the refusal text does not name it - the message enumerates a set that is not the set it refuses from")
        }
    }

    foreach ($t in @($tokens | Sort-Object -Unique)) {
        if ($t -notmatch $operatorShaped) { continue }
        if ($reservedSorted -contains $t) { continue }
        Write-Output ("  {0,-6} STALE      advertised as reserved but no code path reserves it" -f $t)
        $failures.Add("'$t' is advertised by the refusal but is not in the reserved set - a user told they cannot use a name they can")
    }

    Write-Output ''
    Write-Output '--- rule B: a count word in the refusal must match its table ---'
    foreach ($noun in ($countPhrases.Keys | Sort-Object)) {
        $table = $countPhrases[$noun]
        $m = [regex]::Match($msgText, '\b([A-Za-z]+)\s+' + [regex]::Escape($noun) + '\b')
        if (-not $m.Success) {
            Write-Output ("  {0,-14} (no count word in front of it - nothing to go stale)" -f $noun)
            continue
        }
        $word = $m.Groups[1].Value.ToLower()
        if (-not $numberWords.ContainsKey($word)) {
            Write-Output ("  {0,-14} (preceded by '{1}', not a count word - nothing to check)" -f $noun, $word)
            continue
        }
        if (-not $tableSizes.ContainsKey($table)) {
            $failures.Add("count phrase '$word $noun' is checked against table '$table', which contributed no names this run")
            Write-Output ("  {0,-14} TABLE MISSING ({1})" -f $noun, $table)
            continue
        }
        if ($numberWords[$word] -eq $tableSizes[$table]) {
            Write-Output ("  {0,-14} ok         says '{1}', {2} has {3}" -f $noun, $word, $table, $tableSizes[$table])
        } else {
            Write-Output ("  {0,-14} STALE      says '{1}' ({2}), {3} has {4}" -f $noun, $word, $numberWords[$word], $table, $tableSizes[$table])
            $failures.Add("the refusal says '$word $noun' but $table holds $($tableSizes[$table]) - a hand count in prose that has gone stale")
        }
    }
}

# ---- 3. dispatch reachability -------------------------------------------
Write-Output ''
Write-Output '--- rule C: every reserved name must be reached by SolveGoalList ---'

$solveBody = Get-ProcBody -Lines $prologLines -Name 'SolveGoalList'
if ($null -eq $solveBody) {
    $failures.Add('SolveGoalList not found in VLA_Prolog.bas (renamed or removed - update this baseline deliberately)')
    Write-Output '  SolveGoalList NOT FOUND'
} else {
    $solveText = $solveBody -join "`n"

    foreach ($t in $tableNames) {
        $pattern = [regex]::Escape($t) + '\s*\(\s*predName\s*\)\s*<>\s*""'
        if ($solveText -match $pattern) {
            Write-Output ("  {0,-22} ok         dispatched as a table arm" -f $t)
        } else {
            Write-Output ("  {0,-22} UNDISPATCHED" -f $t)
            $failures.Add("table '$t' contributes reserved names but SolveGoalList has no '$t(predName) <> """"' arm - every one of its names is forbidden AND inert")
        }
    }

    foreach ($n in $literalNames) {
        if ($structuralDispatch.ContainsKey($n)) {
            Write-Output ("  {0,-22} ok         exempt: {1}" -f $n, $structuralDispatch[$n])
            continue
        }
        $pattern = 'predName\s*=\s*"' + [regex]::Escape($n) + '"'
        if ($solveText -match $pattern) {
            Write-Output ("  {0,-22} ok         dispatched by name" -f $n)
        } else {
            Write-Output ("  {0,-22} UNDISPATCHED" -f $n)
            $failures.Add("'$n' is reserved but SolveGoalList never dispatches it - the name is forbidden as a predicate AND does nothing")
        }
    }

    # ---- rule D: and nothing BUT a reserved name may be dispatched ------
    # The two arm shapes are read straight out of SolveGoalList's own text
    # and matched back against the sets rule C walked outward from. An
    # assignment (`predName = GoalPredName(...)`) cannot match: the pattern
    # requires a quoted literal on the right-hand side. The empty string is
    # skipped so a defensive `If predName = "" Then` guard, should one ever
    # be added, is not reported as an undeclared predicate name.
    Write-Output ''
    Write-Output '--- rule D: and every arm SolveGoalList has must be a reserved one ---'

    $dispatchedLiterals = @([regex]::Matches($solveText, 'predName\s*=\s*"([^"]*)"') |
        ForEach-Object { $_.Groups[1].Value } | Where-Object { $_.Length -gt 0 } | Sort-Object -Unique)
    $dispatchedTables = @([regex]::Matches($solveText, '([A-Za-z_]\w*)\s*\(\s*predName\s*\)\s*<>\s*""') |
        ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)

    if ($dispatchedLiterals.Count -eq 0 -and $dispatchedTables.Count -eq 0) {
        $failures.Add('SolveGoalList has no predName-keyed dispatch arm of either shape - the chain this rule reads has been restructured, so update this check deliberately rather than letting it scan nothing')
        Write-Output '  NO ARMS FOUND - the dispatch chain no longer has the shape this rule reads'
    }

    foreach ($t in $dispatchedTables) {
        if ($tableNames -contains $t) {
            Write-Output ("  {0,-22} ok         table arm, delegated to by IsReservedPredicateName" -f $t)
        } else {
            Write-Output ("  {0,-22} UNRESERVED" -f $t)
            $failures.Add("SolveGoalList dispatches table '$t' but IsReservedPredicateName does not delegate to it - every name that table holds can be DEFINED by a user and is then silently shadowed by the solver")
        }
    }

    foreach ($n in $dispatchedLiterals) {
        if ($reservedSorted -contains $n) {
            Write-Output ("  {0,-22} ok         reserved (from {1})" -f $n, $sourceOf[$n])
        } else {
            Write-Output ("  {0,-22} UNRESERVED" -f $n)
            $failures.Add("SolveGoalList dispatches '$n' but no code path reserves it - a user may define that predicate, and their own facts are then silently unreachable")
        }
    }
}

Write-Output ''
if ($failures.Count -eq 0) {
    Write-Output '=== CHECK: clean - the reserved set, its dispatch and its refusal text all name the same predicates ==='
} else {
    Write-Output "=== CHECK: $($failures.Count) disagreement(s) about the reserved set ==="
    $failures | ForEach-Object { Write-Output "  $_" }
    Write-Output ''
    Write-Output 'PROLOG reserves a predicate name by forward declaration, so the set lives in three places at once: what the parser refuses, what the solver dispatches, and what the refusal tells the user. Any two of them disagreeing is a silent defect - a name forbidden but inert, a name dispatched over a user predicate that was never refused, or a refusal enumerating a set it did not refuse from. Fix the one that is wrong; do not widen this check to accommodate it.'
}
exit ([Math]::Min($failures.Count, 1))
