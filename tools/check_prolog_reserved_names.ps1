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

  E. Derived near-misses. AliasSpellingFor's own header says its class is
     DERIVED, by two rules over the reserved set - swap every hyphen for
     an underscore (or the reverse), and drop a trailing question mark -
     and is therefore "closed and countable". Closed only if something
     re-derives it: the derivation was done BY HAND at 44 names, and by
     52 it had gone stale twice without anything noticing. `whole` (from
     PROLOG.17's `whole?`) was a silent unknown predicate, which is the
     exact dead end the alias table exists to remove. So every derived
     spelling of every reserved name must itself be reserved - in any
     table; which table owns it is the code's decision, not this rule's.

     Rule 1 applies to WORD-SHAPED names only - names that begin with a
     letter. Run over `->` it produces `_>`, which is an artifact of the
     rule rather than a spelling anyone types; the rule was always about
     joining the WORDS of a name, and an operator has none. Proved by
     mutation: lifting the restriction reports `_>` and nothing else.

     PROLOG.18 added this, and ran it RED on `whole` before fixing it.

  F. The catalogue fits the VBA editor. VBE holds a physical source line
     to 1023 characters and SPLITS a longer one on import, mid-word, into a
     syntax error that stops the whole project compiling - every suite,
     not just PROLOG's. prolog-reserved-predicate-name grows with every
     PROLOG item, sat at 1004 characters at v0.5.5, and PROLOG.18 took it
     to 1196: the owner's compile caught it, not any check. It is now
     written across `_` continuation lines (the parser below joins them),
     and this rule holds EVERY physical line of VLA_Messages.bas to the
     limit, reporting the longest so the headroom is visible in a green
     run. Run RED on the 1196-character line before the split.

  G. No reserved name reaches TermPredName. ValidateBodyItem sends every
     goal it has no arm for into TermPredName, which RECORDS AN ARITY; a
     reserved name can never be defined, so the only thing that arity can
     do is refuse a program writing the name at two arities as
     prolog-arity-mismatch - blaming the author's arities for a goal that
     is refused on sight, and pre-empting the refusal that says why. The
     class is read out of ValidateBodyItem's own arm conditions, never
     listed; the catch-all `IsReservedPredicateName(headWord)` arm must be
     the LAST one, or the solving arms after it lose their shape checks.

     PROLOG.24 added this, and ran it RED on 28 names first: `!`, `list`,
     the two control spellings, the nine ISO spellings, the four
     number-type names and the eleven alias spellings. PROLOG.19 had given
     the impure goals an arm and no other refusing family had one.

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

# ---- baseline: names dispatched structurally, not by predName -----------
# EMPTY since PROLOG.24, and the history is the reason to keep the slot.
# Cut is the ONLY non-object goal ValidateBodyItem can let through to
# SolveGoalList, so its arm runs BEFORE GoalPredName (which assumes a
# Collection) and is not a `predName = "!"` test - and `!` used to be
# exempt from rule C here for that reason. The exemption hid a real
# defect: cut written in PARENTHESES, `(! ...)`, is a compound goal named
# "!" that reaches the name-keyed arms, and nothing dispatched it, so it
# fell through to the clauseDict lookup and failed silently - reserved AND
# inert, the first defect this script's header names. PROLOG.24 gave it a
# `predName = "!"` arm refusing it by name, so rule C now checks `!` like
# any other literal and deleting that arm turns it red. A name added here
# must be one with NO name-keyed form at all, with the reason recorded.
$structuralDispatch = @{}

# ---- baseline: count words in the refusal, and the table each describes -
# Key is the group noun as it appears in the text; value is the table
# function whose size the preceding count word must equal.
# NOTE on phrasing: the count word must sit IMMEDIATELY before the group
# noun, because rule B matches '<word> <noun>'. "the nine ISO spellings"
# is checked; "their nine bare ISO spellings" would find "bare" in front
# of the noun, report "not a count word", and silently check nothing.
#
# PROLOG.15 found the OTHER half of that hole, and it is worse: a count
# with no group noun at all. prolog-type-test-iso-spelling read "all six
# of them end in a question mark" and there were nine, in a message this
# rule never looks at - rule B reads prolog-reserved-predicate-name and
# nothing else. That text now carries no count at all, which is the only
# version that cannot go stale; the same treatment is the right fix
# anywhere else a bare count turns up in prose.
# NOTE on collision: rule B matches '<word> <noun>' and Regex.Match returns
# the FIRST hit anywhere in the text, so a group noun that CONTAINS another
# group noun would silently steal its check. PROLOG.15's own deferred family
# is therefore called 'number-type names' and not, say, 'deferred type
# tests' - the latter contains 'type tests', would match "deferred type
# tests" with 'deferred' in front of it, report "not a count word", and
# check TypeTestKindFor's real size against nothing at all.
$countPhrases = @{
    'comparisons'          = 'ComparisonOpFor'
    'term-matching goals'  = 'UnificationOpFor'
    'type tests'           = 'TypeTestKindFor'
    'ISO spellings'        = 'TypeTestIsoSpellingFor'
    'list goals'           = 'ListGoalKindFor'
    'number-type names'    = 'TypeTestDeferredFor'
    'alias spellings'      = 'AliasSpellingFor'
    # PROLOG.14's own table. Called 'control spellings' and deliberately
    # NOT 'ISO control spellings', which would contain the existing key
    # 'ISO spellings' and, per the collision note above, silently steal
    # its check - the scan would match "...ISO spellings" inside it, find
    # 'control' sitting in front of the noun, report "not a count word",
    # and hold TypeTestIsoSpellingFor's real size against nothing at all.
    # Checked for containment against all seven existing keys, both
    # directions, before choosing it.
    'control spellings'    = 'ControlIsoSpellingFor'
    # PROLOG.18's own table. Checked for containment against all eight
    # existing keys, both directions, before choosing it: 'text goals'
    # contains none of them and none contains it. 'list goals' is the near
    # neighbour and the two share only the word 'goals', which is not a
    # key on its own.
    'text goals'           = 'TextGoalKindFor'
    # PROLOG.19's own table. Checked for containment against all nine
    # existing keys, both directions, before choosing it. PHRASED WITH NO
    # COUNT, on purpose, and the entry is here anyway: the family holds
    # more names than $numberWords can read (it stops at 'twelve'). Measured,
    # not assumed: "the thirty impure goals" captures 'thirty', reaches the
    # "not a count word" branch below and passes having checked nothing;
    # "the thirty-nine impure goals" captures only 'nine' (the regex stops at
    # the hyphen) and fails, but as 9 against the table's real size - loud,
    # and for the wrong reason. With no count there is nothing to go stale;
    # with this entry, the day someone adds one is the day it is looked at.
    'impure goals'         = 'ImpureGoalKindFor'
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
# PROLOG.18: physical lines ending in ` _` are joined into one logical
# statement first, and the `" & "` seams between the literals of a split
# template are removed, so the refusal reads the same whether it is written
# on one line or several. It HAS to be several - see rule F - and without
# this join the id would simply not be found (a loud failure, but one that
# stops every other rule here from checking anything).
#
# The SEAM removal is the quieter half and was proved separately, because
# a mutation deleting it came back GREEN on the shipped text - every split
# there falls between clauses. It matters when a split lands between a
# count word and its noun (`"the eight " & _` / `"text goals ..."`): with
# the seam left in, rule B finds no count word, reports "nothing to go
# stale" and passes blind; with it removed, the same wrong count is
# caught. Shown on a synthetic copy of the catalogue, both ways.
$messagesLines = Get-Content -LiteralPath $messagesPath
$logicalLines = New-Object System.Collections.Generic.List[string]
$pending = ''
foreach ($line in $messagesLines) {
    if ($line -match '\s_\s*$' -and -not $line.TrimStart().StartsWith("'")) {
        $pending += ($line -replace '\s_\s*$', ' ')
        continue
    }
    $logicalLines.Add($pending + $line)
    $pending = ''
}
$msgText = $null
foreach ($line in $logicalLines) {
    if ($line -match '^\s*AddMsg\s+m,\s*"([^"]+)"\s*,\s*[^,]+,\s*"[^"]*"\s*,\s*"(.*)"\s*$') {
        if ($Matches[1] -eq $reservedMsgId) { $msgText = ($Matches[2] -replace '"\s*&\s*"', '') }
    }
}

Write-Output ''
Write-Output "--- rule A: $reservedMsgId must enumerate exactly that set ---"

if ($null -eq $msgText) {
    $failures.Add("$reservedMsgId is not defined in VLA_Messages.bas")
    Write-Output '  UNDEFINED'
} else {
    # PROLOG.14: the scan is scoped to the PARENTHESISED CATALOGUE - the
    # "(is/not/findall/... spelling it does use)" span - and not to the
    # whole sentence, which is prose either side of it.
    #
    # This is not tidiness; it closes a real hole that PROLOG.14 was about
    # to walk straight through. Rule A's forward direction asks whether a
    # reserved name appears as a whole token ANYWHERE in the text, and the
    # sentence ends "...can't be used as a predicate name in a (fact ...)
    # or (rule ...)". That trailing `or` is ordinary English, but it is
    # also a whole token - so the moment `or` became a reserved name, the
    # refusal would have been credited with advertising it while saying
    # nothing about it at all. Verified before the fix: `or` occurred once
    # in the text and zero times in the catalogue.
    #
    # It is a live hazard rather than a one-off, because a good half of the
    # reserved set is already made of ordinary English words - is, not,
    # list, between, length, member, append, reverse, number, atom, ground,
    # var, callable, compound, integer, float - any of which a future
    # sentence could mention in passing and silently satisfy this rule.
    #
    # The catalogue is found by matching the first "(" to its own closing
    # ")", so a name written with parentheses inside it stays inside, and
    # the "(fact ...)" the sentence ends on stays outside.
    $catOpen = $msgText.IndexOf('(')
    $catClose = -1
    if ($catOpen -ge 0) {
        $depth = 0
        for ($j = $catOpen; $j -lt $msgText.Length; $j++) {
            if ($msgText[$j] -eq '(') { $depth++ }
            elseif ($msgText[$j] -eq ')') {
                $depth--
                if ($depth -eq 0) { $catClose = $j; break }
            }
        }
    }
    if ($catOpen -lt 0 -or $catClose -lt 0) {
        $failures.Add("$reservedMsgId has no parenthesised catalogue to read - rule A would otherwise scan the whole sentence, where ordinary prose words silently satisfy it")
        $catalogue = ''
    } else {
        $catalogue = $msgText.Substring($catOpen + 1, $catClose - $catOpen - 1)
    }

    # Tokens are split on whitespace and on the "/" the text uses to run
    # short names together ("is/not/findall/!"); surrounding prose
    # punctuation is trimmed from each end. A name is advertised only if it
    # survives as a WHOLE token - a bare Contains() would let "<" pass on
    # the strength of any "=<" already in the sentence.
    $rawTokens = $catalogue -split '[\s/]+'
    $tokens = New-Object System.Collections.Generic.List[string]
    foreach ($t in $rawTokens) {
        $trimmed = $t.Trim(",.;()'`"")
        if ($trimmed.Length -gt 0) { $tokens.Add($trimmed) }
    }

    # A rule that looks at nothing passes by not looking. The catalogue has
    # carried well over a hundred tokens since PROLOG.13, so a scan that
    # collapses to a handful means the extraction above found the wrong
    # span, not that the refusal got shorter - and every name would then
    # report MISSING rather than the check quietly passing. Reported
    # either way, so the number is visible in a green run too.
    Write-Output ("  catalogue: {0} char(s), {1} token(s) examined" -f $catalogue.Length, $tokens.Count)
    if ($tokens.Count -lt 40) {
        $failures.Add("rule A examined only $($tokens.Count) token(s) of $reservedMsgId - the catalogue span looks wrong, and a scan this small cannot have checked the reserved set")
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

# ---- rule F: every physical line of the catalogue fits the VBA editor ----
Write-Output ''
Write-Output '--- rule F: no physical line of VLA_Messages.bas may exceed the VBA editor''s 1023 characters ---'
$vbeMaxLine = 1023
$longest = 0; $longestAt = 0; $lineNo = 0
foreach ($line in $messagesLines) {
    $lineNo++
    if ($line.Length -gt $longest) { $longest = $line.Length; $longestAt = $lineNo }
    if ($line.Length -gt $vbeMaxLine) {
        Write-Output ("  line {0,5}: {1} characters  TOO LONG" -f $lineNo, $line.Length)
        $failures.Add("VLA_Messages.bas line $lineNo is $($line.Length) characters; the VBA editor splits anything over $vbeMaxLine on import, mid-word, into a syntax error that stops the whole project compiling - continue the statement across lines with ' _'")
    }
}
# A rule that looks at nothing passes by not looking.
if ($lineNo -lt 500) { $failures.Add("rule F examined only $lineNo line(s) of VLA_Messages.bas - the file was not read") }
Write-Output ("  {0} line(s) examined; longest is line {1} at {2} of {3} characters" -f $lineNo, $longestAt, $longest, $vbeMaxLine)

# ---- rule E: every derived near-miss spelling must itself be reserved ----
# The two derivation rules AliasSpellingFor's own header states, applied
# mechanically rather than by hand: see the header above for why a hand
# derivation went stale. A candidate equal to its source (a name with no
# hyphen, underscore or trailing '?') is not a candidate.
Write-Output ''
Write-Output '--- rule E: every derived near-miss spelling must itself be reserved ---'

$derivedFrom = @{}
$wordShaped = 0
foreach ($n in $reservedSorted) {
    $cands = New-Object System.Collections.Generic.List[string]
    if ($n -cmatch '^[a-z]') {
        $wordShaped++
        if ($n.Contains('-')) { $cands.Add($n.Replace('-', '_')) }
        if ($n.Contains('_')) { $cands.Add($n.Replace('_', '-')) }
    }
    if ($n.EndsWith('?')) { $cands.Add($n.Substring(0, $n.Length - 1)) }
    # The rules are applied ONCE, never composed, and that is not an
    # omission. A composed spelling - `is_list` from `is-list?` by way of
    # `is_list?` - can only be reached through a first-order one, and this
    # rule requires every first-order spelling to be reserved; once it is,
    # it is a reserved name in its own right and is expanded on its own
    # turn, finding the composed spelling there. So composing can never
    # change the verdict. A first version composed anyway, and deleting
    # that step changed nothing - which is the honest evidence it was
    # never load-bearing, so it is gone rather than kept as a "guard".
    foreach ($c in $cands) {
        if ($c -ceq $n) { continue }
        if (-not $derivedFrom.ContainsKey($c)) { $derivedFrom[$c] = $n }
    }
}

# A rule that looks at nothing passes by not looking. Most of the reserved
# set is word-shaped and a good share of it carries a hyphen or a '?', so a
# derivation that collapses to a handful means the reserved set above was
# not read, not that the class got smaller.
Write-Output ("  {0} reserved name(s) examined, {1} word-shaped; {2} derived spelling(s)" -f $reservedSorted.Count, $wordShaped, $derivedFrom.Count)
if ($wordShaped -lt 30 -or $derivedFrom.Count -lt 15) {
    $failures.Add("rule E examined $wordShaped word-shaped name(s) and derived $($derivedFrom.Count) spelling(s) - too few to have read the reserved set, so the derivation cannot be trusted")
}

foreach ($c in ($derivedFrom.Keys | Sort-Object)) {
    if ($reservedSorted -ccontains $c) {
        Write-Output ("  {0,-20} ok         reserved (from {1}), derived from {2}" -f $c, $sourceOf[$c], $derivedFrom[$c])
    } else {
        Write-Output ("  {0,-20} UNRESERVED derived from {1}" -f $c, $derivedFrom[$c])
        $failures.Add("'$c' is a near-miss spelling of the reserved '$($derivedFrom[$c])' but is not itself reserved - a user who types it gets a silent unknown predicate instead of a refusal naming the real spelling")
    }
}

# ---- rule G: no reserved name may reach TermPredName ---------------------
# PROLOG.24. ValidateBodyItem sends every goal it has no arm for into
# TermPredName, which RECORDS AN ARITY - and a reserved name can never be
# defined, so the only thing a recorded arity can do is refuse a program
# that uses the name at two arities, as prolog-arity-mismatch, blaming the
# author's arities for a goal that is refused on sight whatever its shape.
# The class is read out of ValidateBodyItem's own arm conditions, never
# listed here: a literal `headWord = "..."`, a table `<Table>(headWord) <>
# ""`, or the catch-all `IsReservedPredicateName(headWord)`. The catch-all
# must be the LAST such arm - placed earlier, it would skip the shape
# checks of every solving arm below it, which every bad-shape pin would
# then catch, but this rule says why.
Write-Output ''
Write-Output '--- rule G: ValidateBodyItem must keep every reserved name out of TermPredName ---'
$vbiBody = Get-ProcBody -Lines $prologLines -Name 'ValidateBodyItem'
if ($null -eq $vbiBody) {
    $failures.Add('ValidateBodyItem not found in VLA_Prolog.bas (renamed or removed - update this baseline deliberately)')
    Write-Output '  ValidateBodyItem NOT FOUND'
} else {
    $armConds = New-Object System.Collections.Generic.List[string]
    foreach ($line in $vbiBody) {
        if ($line -match '^\s*(If|ElseIf)\s+(.*headWord.*?)\s+Then\s*(.*)$') { $armConds.Add($Matches[2]) }
    }
    $covered = New-Object System.Collections.Generic.List[string]
    # EVERY catch-all is recorded, not only the last one seen: a second,
    # early copy would skip the shape checks of every solving arm between
    # it and the real one, and a single "where is the catch-all" index
    # overwritten per arm reads only the last copy and passes. Measured: that
    # first version went GREEN on exactly that mutant.
    $catchAlls = New-Object System.Collections.Generic.List[int]
    for ($ai = 0; $ai -lt $armConds.Count; $ai++) {
        $cond = $armConds[$ai]
        foreach ($m in [regex]::Matches($cond, 'headWord\s*=\s*"([^"]*)"')) { $covered.Add($m.Groups[1].Value) }
        foreach ($m in [regex]::Matches($cond, '([A-Za-z_]\w*)\s*\(\s*headWord\s*\)\s*<>\s*""')) {
            $tb = $m.Groups[1].Value
            foreach ($n in $reservedSorted) { if ($sourceOf[$n] -eq $tb) { $covered.Add($n) } }
        }
        if ($cond -match 'IsReservedPredicateName\s*\(\s*headWord\s*\)') { $catchAlls.Add($ai) }
    }
    # A rule that looks at nothing passes by not looking.
    Write-Output ("  {0} headWord-keyed arm(s) read" -f $armConds.Count)
    if ($armConds.Count -lt 10) { $failures.Add("rule G read only $($armConds.Count) headWord-keyed arm(s) in ValidateBodyItem - the arm chain was not found, so nothing was checked") }
    if ($catchAlls.Count -gt 0) {
        $early = @($catchAlls | Where-Object { $_ -ne $armConds.Count - 1 })
        if ($early.Count -gt 0) {
            foreach ($e in $early) {
                $failures.Add("ValidateBodyItem has an IsReservedPredicateName(headWord) catch-all at arm $($e + 1) of $($armConds.Count), not the last - every solving arm after it loses its shape check")
            }
            Write-Output ("  catch-all IS NOT only the last arm (found at arm(s) {0} of {1})" -f (($catchAlls | ForEach-Object { $_ + 1 }) -join ', '), $armConds.Count)
        } else {
            Write-Output '  catch-all IsReservedPredicateName(headWord) is the last arm - every reserved name is kept out'
        }
    } else {
        $leaks = @($reservedSorted | Where-Object { $covered -notcontains $_ })
        foreach ($n in $leaks) {
            Write-Output ("  {0,-20} REACHES TermPredName (from {1})" -f $n, $sourceOf[$n])
            $failures.Add("'$n' is reserved but ValidateBodyItem has no arm for it, so TermPredName records its arity - using it at two arities is refused as prolog-arity-mismatch instead of by its own refusal")
        }
        if ($leaks.Count -eq 0) { Write-Output ("  all {0} reserved name(s) have an arm" -f $reservedSorted.Count) }
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
