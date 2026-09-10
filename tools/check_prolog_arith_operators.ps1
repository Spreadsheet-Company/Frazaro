<#
check_prolog_arith_operators.ps1 - PROLOG.17's mechanical check.

WHY THIS EXISTS, and it is two defects rather than one.

DEFECT ONE - THE SILENT SEAM. VLA_Relation.ComputeArithmetic is shared
substrate with THREE callers, one in each of VLA_Prolog, VLA_Sql and
VLA_Datalog. Until PROLOG.17 its `Select Case op` had NO `Case Else`, and
`ok = True` was set unconditionally after it - so an operator that reached
it without a Case returned EMPTY and reported SUCCESS. That was latent
rather than live (each engine gated its own operator set first), but the
moment an item adds an operator to one engine's gate and forgets the
substrate, three engines compute Empty and call it a valid answer. The
failure is a wrong NUMBER in a SQL or DATALOG cell, which looks like an
answer; this project holds that worse than a crash. The Case Else now
catches it at runtime. RULE C below catches it before the code ships.

DEFECT TWO - THE HAND-WRITTEN LIST IN PROSE. The refusal
prolog-arith-unknown-operator names the operators it recognizes in its own
text. Before this item it read "only +, -, *, and / are supported" and
NOTHING would have noticed that going stale: check_prolog_reserved_names
reads a different message entirely, and its own $operatorShaped pattern
(^[<>=\\!:]+$) matches none of + - * /. A user would have been told an
operator was unsupported in the same breath as the engine supporting it.

WHAT IT CHECKS.

  RULE A  Every operator in VLA_Relation's ArithOpArity table is named
          in the refusal's prose list, and vice versa. Both directions:
          an operator missing from the prose is one a user is wrongly
          told does not exist; one in the prose but not the table is one
          they are told about and cannot use.

  RULE B  ArithOpArity is the ONLY place the set is spelled. It lives in
          VLA_Relation, beside the two functions that implement the
          operators, because all THREE engines gate their own arithmetic
          and a per-engine copy would be three more places to forget.
          The set used to be written three times (ValidateArithExpr,
          EvalArithTerm, ComputeArithmetic) and the roadmap entry that
          scoped this work knew of only two. This rule scans VLA_Prolog's
          two arithmetic procedures for a second literal list, which is
          how parse-time and runtime drift apart.

  RULE C  THE SEAM, and the reason this file is worth more than the rest
          of it put together. Every BINARY operator in the table must
          have a Case in VLA_Relation.ComputeArithmetic, and every UNARY
          one a Case in ComputeArithmeticUnary. This is the check that
          makes "add it to the table and forget the substrate"
          impossible.

  RULE D  Both substrate functions must still carry a `Case Else`. A
          future edit that removes one re-opens defect one silently.

House style (tools/check_*.ps1): PowerShell 5.1, host-independent, a
hard-coded reviewable baseline, no live Excel needed. Never wired into
VlaSelfTest - it reads source text, not runtime behaviour.

Every rule COUNTS WHAT IT EXAMINED and fails if that collapses to zero: a
rule that looks at nothing otherwise passes by not looking, which is how
a renamed procedure turns a ratchet into decoration.

Usage:  powershell -File tools\check_prolog_arith_operators.ps1
Exit code: 0 clean, 1 if the operator set disagrees with itself anywhere.
#>

$ErrorActionPreference = 'Stop'

$repoRoot     = Split-Path -Parent $PSScriptRoot
$prologPath   = Join-Path $repoRoot 'src\VLA_Prolog.bas'
$relationPath = Join-Path $repoRoot 'src\VLA_Relation.bas'
$messagesPath = Join-Path $repoRoot 'src\VLA_Messages.bas'

# ---- baseline: the reviewable names this check is written against -----
# Deliberately NAMES, not patterns: a rename must break this run loudly
# rather than let it silently scan nothing.
$tableProc     = 'ArithOpArity'
$binaryProc    = 'ComputeArithmetic'
$unaryProc     = 'ComputeArithmeticUnary'
$scannedProcs  = @('ValidateArithExpr', 'EvalArithTerm')
$refusalIds    = @('prolog-arith-unknown-operator', 'datalog-unknown-arithmetic-operator')
$prosePrefix   = 'the ones it knows are '

# ---- baseline: rule E exemptions, by exact query text and with reason --
# One assertion legitimately puts a REAL operator inside a query about an
# unknown one, because what it is testing is the NESTED FORM in operator
# position - `((+ 1 2) 3 4)` is refused for having a form where a symbol
# belongs, and the `+` is incidental, inside the nested form rather than
# operating. Exempted BY NAME with its reason (the $rawIndexExempt
# precedent in check_test_assertion_safety.ps1) rather than by loosening
# rule E until it stops noticing the real ones.
$ruleEexempt = @{
    '(query (is X ((+ 1 2) 3 4)))' =
        'the operator position holds a NESTED FORM; the + is inside it, not operating, and ValidateArithExpr refuses on the IsObject(lst.Item(1)) guard before any operator lookup happens'
    '(query (is X (+ () 3)))' =
        'the + is the OUTER operator and is perfectly valid; the refusal is about the EMPTY OPERAND () nested inside it, which PROLOG.17 re-pointed from the wrong-arity message to this one because a form with no operator has no arity to want'
    '(fact (thing (% 5 2))) (rule (compute X) (thing Y) (is X (+ Y 1))) (query (compute X))' =
        'the + is the valid OUTER operator; what must stay unrecognized is the % the variable Y dereferences to. This is the query whose earlier spelling used mod, computed a real answer, returned a spilled array and made the test CStr-crash the whole live run'
}

foreach ($p in @($prologPath, $relationPath, $messagesPath)) {
    if (-not (Test-Path -LiteralPath $p)) {
        Write-Output "FAIL: missing source file $p"
        exit 1
    }
}

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

# Every double-quoted token appearing in a `Case ...` line of a body.
function Get-CaseTokens {
    param([string[]]$Body)
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($line in $Body) {
        if ($line -notmatch '^\s*Case\b') { continue }
        foreach ($m in [regex]::Matches($line, '"([^"]*)"')) {
            $out.Add($m.Groups[1].Value)
        }
    }
    return $out
}

$failures = New-Object System.Collections.Generic.List[string]
$prologLines   = Get-Content -LiteralPath $prologPath
$relationLines = Get-Content -LiteralPath $relationPath

Write-Output '=== PROLOG.17 - one arithmetic operator set, agreed everywhere ==='
Write-Output "Prolog:   $prologPath"
Write-Output "Relation: $relationPath"
Write-Output "Messages: $messagesPath"
Write-Output ''

# ---- the table -------------------------------------------------------
$tableBody = Get-ProcBody -Lines $relationLines -Name $tableProc
if ($null -eq $tableBody) {
    Write-Output "FAIL: '$tableProc' not found in VLA_Relation.bas (renamed or removed - update this baseline deliberately)"
    exit 1
}

$arityOf = @{}
$pendingOps = New-Object System.Collections.Generic.List[string]
foreach ($line in $tableBody) {
    if ($line -match '^\s*Case\b') {
        foreach ($m in [regex]::Matches($line, '"([^"]*)"')) { $pendingOps.Add($m.Groups[1].Value) }
    }
    if ($line -match ($tableProc + '\s*=\s*(\d+)')) {
        $n = [int]$Matches[1]
        foreach ($op in $pendingOps) { $arityOf[$op] = $n }
        $pendingOps.Clear()
    }
}

if ($arityOf.Keys.Count -eq 0) {
    $failures.Add("$tableProc yielded NO operators - the parse matched nothing, which is itself a failure")
}

$binaryOps = @($arityOf.Keys | Where-Object { $arityOf[$_] -eq 2 } | Sort-Object)
$unaryOps  = @($arityOf.Keys | Where-Object { $arityOf[$_] -eq 1 } | Sort-Object)

Write-Output '--- the one table ---'
Write-Output ("  {0,-20} {1,2} binary: {2}" -f $tableProc, $binaryOps.Count, ($binaryOps -join ' '))
Write-Output ("  {0,-20} {1,2} unary : {2}" -f '', $unaryOps.Count, ($unaryOps -join ' '))

# ---- RULE A: the table vs the refusal's own prose --------------------
Write-Output ''
Write-Output '--- rule A: the refusal must advertise exactly the table ---'

# PROLOG.17 opened the set in THREE engines, and PROLOG and DATALOG each
# advertise it in their own refusal text. Both are checked, from one loop,
# because a second prose list with no rule over it is precisely the defect
# this rule exists for - and the DATALOG one was added by the same item
# that added this check, so it would have been the first to rot.
$msgTextOf = @{}
foreach ($line in (Get-Content -LiteralPath $messagesPath)) {
    foreach ($id in $refusalIds) {
        if ($line -match ('^\s*AddMsg\s+m,\s*"' + [regex]::Escape($id) + '"\s*,\s*[^,]+,\s*"[^"]*"\s*,\s*"(.*)"\s*$')) {
            $msgTextOf[$id] = $Matches[1]
        }
    }
}

$ruleAexamined = 0
foreach ($refusalId in $refusalIds) {
$msgText = ''
if ($msgTextOf.ContainsKey($refusalId)) { $msgText = $msgTextOf[$refusalId] }
if ($msgText -eq '') {
    $failures.Add("$refusalId - not found in VLA_Messages.bas (renamed, or its AddMsg shape changed)")
    Write-Output "  $refusalId NOT FOUND"
} else {
    $ix = $msgText.IndexOf($prosePrefix, [System.StringComparison]::OrdinalIgnoreCase)
    if ($ix -lt 0) {
        $failures.Add("$refusalId - its text no longer contains the phrase '$prosePrefix', so this rule can no longer find the list it is meant to check")
        Write-Output "  prose list NOT FOUND (the phrase '$prosePrefix' is gone)"
    } else {
        $tail = $msgText.Substring($ix + $prosePrefix.Length)
        $dot  = $tail.IndexOf('.')
        # '.' terminates the sentence, but no operator name contains one,
        # so the first '.' is safe as the end of the list.
        if ($dot -ge 0) { $tail = $tail.Substring(0, $dot) }
        $advertised = @($tail -split '\s+' | Where-Object { $_ -ne '' } | Sort-Object)
        $tabled     = @($arityOf.Keys | Sort-Object)
        $ruleAexamined = $ruleAexamined + $advertised.Count

        Write-Output ("  {0,-34} advertises {1}: {2}" -f $refusalId, $advertised.Count, ($advertised -join ' '))

        if ($advertised.Count -eq 0) {
            $failures.Add("$refusalId - advertises NO operators; rule A examined nothing for it")
        }
        foreach ($op in $tabled) {
            if ($advertised -notcontains $op) {
                $failures.Add("'$op' is in $tableProc but $refusalId does not advertise it - a user is told an operator does not exist while the engine supports it")
                Write-Output ("  {0,-10} MISSING FROM PROSE of {1}" -f $op, $refusalId)
            }
        }
        foreach ($op in $advertised) {
            if ($tabled -notcontains $op) {
                $failures.Add("'$op' is advertised by $refusalId but is not in $tableProc - a user is told about an operator they cannot use")
                Write-Output ("  {0,-10} ADVERTISED BUT UNTABLED by {1}" -f $op, $refusalId)
            }
        }
    }
}
}
if ($ruleAexamined -eq 0) {
    $failures.Add('rule A examined zero advertised operators across every refusal - it looked at nothing')
}

# ---- RULE B: no second literal operator list in VLA_Prolog -----------
Write-Output ''
Write-Output '--- rule B: ArithOpArity is the only place the set is spelled ---'
$ruleBexamined = 0
foreach ($proc in $scannedProcs) {
    $body = Get-ProcBody -Lines $prologLines -Name $proc
    if ($null -eq $body) {
        $failures.Add("procedure '$proc' not found in VLA_Prolog.bas (renamed or removed - update this baseline deliberately)")
        Write-Output ("  {0,-20} NOT FOUND" -f $proc)
        continue
    }
    $ruleBexamined = $ruleBexamined + $body.Count
    $toks = @(Get-CaseTokens -Body $body)
    $opLike = @($toks | Where-Object { $arityOf.ContainsKey($_) })
    if ($opLike.Count -gt 0) {
        $failures.Add("$proc spells operator name(s) $($opLike -join ' ') in its own Case arms - ask $tableProc instead, or parse-time and runtime will drift apart")
        Write-Output ("  {0,-20} SECOND LIST: {1}" -f $proc, ($opLike -join ' '))
    } else {
        Write-Output ("  {0,-20} ok - {1} line(s), no operator literals" -f $proc, $body.Count)
    }
}
if ($ruleBexamined -eq 0) {
    $failures.Add('rule B examined zero lines - it scanned nothing and would pass by not looking')
}

# ---- RULE C: THE SEAM -----------------------------------------------
Write-Output ''
Write-Output '--- rule C: every tabled operator has a Case in the shared substrate ---'

$binBody = Get-ProcBody -Lines $relationLines -Name $binaryProc
$unBody  = Get-ProcBody -Lines $relationLines -Name $unaryProc
$seamExamined = 0

if ($null -eq $binBody) {
    $failures.Add("'$binaryProc' not found in VLA_Relation.bas - the shared substrate this check exists to guard")
} elseif ($null -eq $unBody) {
    $failures.Add("'$unaryProc' not found in VLA_Relation.bas - the shared substrate this check exists to guard")
} else {
    $binCases = @(Get-CaseTokens -Body $binBody)
    $unCases  = @(Get-CaseTokens -Body $unBody)
    foreach ($op in $binaryOps) {
        $seamExamined++
        if ($binCases -contains $op) {
            Write-Output ("  {0,-10} binary  ok - has a Case in {1}" -f $op, $binaryProc)
        } else {
            Write-Output ("  {0,-10} binary  NO CASE in {1}" -f $op, $binaryProc)
            $failures.Add("binary operator '$op' is tabled in $tableProc but has NO Case in $binaryProc - it would reach the shared substrate, fall to Case Else, and be refused as unknown at runtime (before that Case Else existed it computed EMPTY and reported SUCCESS)")
        }
    }
    foreach ($op in $unaryOps) {
        $seamExamined++
        if ($unCases -contains $op) {
            Write-Output ("  {0,-10} unary   ok - has a Case in {1}" -f $op, $unaryProc)
        } else {
            Write-Output ("  {0,-10} unary   NO CASE in {1}" -f $op, $unaryProc)
            $failures.Add("unary operator '$op' is tabled in $tableProc but has NO Case in $unaryProc - same seam, same failure")
        }
    }
}
if ($seamExamined -eq 0) {
    $failures.Add('rule C examined zero operators - the seam this file exists for was not looked at')
}

# ---- RULE D: the Case Else must survive ------------------------------
Write-Output ''
Write-Output '--- rule D: the substrate keeps its Case Else ---'
foreach ($pair in @(@($binaryProc, $binBody), @($unaryProc, $unBody))) {
    $name = $pair[0]
    $body = $pair[1]
    if ($null -eq $body) { continue }
    $hasElse = @($body | Where-Object { $_ -match '^\s*Case\s+Else\b' }).Count
    if ($hasElse -gt 0) {
        Write-Output ("  {0,-24} ok - Case Else present" -f $name)
    } else {
        Write-Output ("  {0,-24} NO Case Else" -f $name)
        $failures.Add("$name has no 'Case Else' - an operator reaching it without a Case would return EMPTY and report SUCCESS, which is the defect this whole file exists to prevent")
    }
}

# ---- RULE E: no test may use a REAL operator as its example of a fake one
# The class rule G in check_test_assertion_safety.ps1 does NOT cover, found
# by hand during PROLOG.17 and mechanized here so the next widening cannot
# repeat it. Rule G catches an expected RENDERING the writer can no longer
# emit; this catches an INPUT whose MEANING changed underneath a test. Two
# assertions used `(mod 5 2)` as their example of an unrecognized operator
# and were correct for four items - until `mod` became real, at which point
# they asserted that a working operator does not work. The expected text
# ("isn't an arithmetic operator") is still perfectly emittable, so nothing
# about the assertion looks stale; only the input gave it away.
Write-Output ''
Write-Output '--- rule E: no test uses a tabled operator as its example of an unknown one ---'
$testsPath = Join-Path $repoRoot 'src\VLA_Tests_Query.bas'
$ruleEexamined = 0
if (-not (Test-Path -LiteralPath $testsPath)) {
    $failures.Add("missing $testsPath - rule E could not run")
} else {
    $tn = 0
    foreach ($line in (Get-Content -LiteralPath $testsPath)) {
        if ($line -notmatch 'isn''t an arithmetic operator') { continue }
        $tn++
    }
    # The refusal is asserted on a PRIOR line's PROLOG(...) call, so the
    # scan walks the file keeping the last query text seen.
    $lastQuery = ''
    $lastQueryLine = 0
    $n = 0
    $testLines = @(Get-Content -LiteralPath $testsPath)
    foreach ($line in $testLines) {
        $n++
        $qm = [regex]::Match($line, 'PROLOG\("(.*?)"\)')
        if ($qm.Success) { $lastQuery = $qm.Groups[1].Value; $lastQueryLine = $n }
        # TWO triggers, not one. The first is the refusal TEXT an
        # assertion expects. The second is the word a test uses to
        # DESCRIBE itself - and it is the one that matters, because it
        # caught a case the first could not: a REGRESSION test citing
        # "an unrecognized operator symbol" asserted a DIFFERENT message
        # ("isn't one", the not-numeric refusal) while still depending on
        # `mod` being unrecognized. When `mod` became real that query
        # started SUCCEEDING, PROLOG returned a spilled ARRAY, and the
        # test's own `CStr(...)` raised a type mismatch that KILLED the
        # run - the exact outcome a test suite must never produce, and it
        # reached the owner's live pass.
        if ($line -notmatch "isn't an arithmetic operator" -and $line -notmatch 'unrecogni[sz]ed') { continue }
        # ...and the trigger must belong to an assertion ABOUT THE REFUSAL
        # STRING. `InStr(1, r, ...)` is how this suite tests refusal text,
        # so its presence on the trigger line or the next one is what
        # separates a NEGATIVE test (the class this rule polices) from a
        # POSITIVE one whose prose merely mentions the history - the twin
        # asserting that `mod` NOW COMPUTES says "used to cite as
        # unrecognized" and must not be flagged for saying so.
        $isRefusalAssertion = $false
        if ($line -match 'InStr\(1,\s*r,') { $isRefusalAssertion = $true }
        if ($n -lt $testLines.Count) {
            if ($testLines[$n] -match 'InStr\(1,\s*r,') { $isRefusalAssertion = $true }
        }
        if (-not $isRefusalAssertion) { continue }
        $ruleEexamined++
        if ($lastQuery -eq '') { continue }
        if ($ruleEexempt.ContainsKey($lastQuery)) {
            Write-Output ("  line {0,-6} exempt     {1}" -f $lastQueryLine, $ruleEexempt[$lastQuery])
            continue
        }
        foreach ($op in ($arityOf.Keys | Sort-Object)) {
            # the operator in HEAD position of a parenthesised form
            if ($lastQuery -match ('\(\s*' + [regex]::Escape($op) + '\s')) {
                $failures.Add("VLA_Tests_Query.bas line $lastQueryLine uses '$op' as its example of an operator PROLOG does not recognize, but '$op' IS in $tableProc - the assertion now claims a working operator does not work")
                Write-Output ("  line {0,-6} STALE INPUT: uses '{1}', which is now a real operator" -f $lastQueryLine, $op)
            }
        }
    }
    if ($ruleEexamined -eq 0) {
        $failures.Add("rule E found NO assertion on the unknown-operator refusal - it examined nothing and would pass by not looking")
        Write-Output '  EXAMINED NOTHING'
    } else {
        Write-Output ("  examined {0} assertion(s) on the unknown-operator refusal" -f $ruleEexamined)
    }
}

# ---- RULE F: SQL's own keyword map must land in the table -------------
# SQL spells the operators differently (LEAST for min, POWER for **,
# because MIN( and MAX( were already aggregate keywords), so it carries a
# TRANSLATION rather than a copy of the set. A translation can still point
# somewhere that does not exist, and the consequence is worse than a
# missing Case: ParsePrimary asks ArithOpArity for the arity, an unknown
# name answers 0, and the parser would then take the BINARY branch for a
# function it knows nothing about. Checked at the target end, so SQL is
# free to spell the names its own way but never free to invent operators.
Write-Output ''
Write-Output '--- rule F: every SQL scalar-function keyword maps to a tabled operator ---'
$sqlPath = Join-Path $repoRoot 'src\VLA_Sql.bas'
$ruleFexamined = 0
if (-not (Test-Path -LiteralPath $sqlPath)) {
    $failures.Add("missing $sqlPath - rule F could not run")
} else {
    $sqlBody = Get-ProcBody -Lines (Get-Content -LiteralPath $sqlPath) -Name 'ScalarFuncOpFor'
    if ($null -eq $sqlBody) {
        $failures.Add("'ScalarFuncOpFor' not found in VLA_Sql.bas (renamed or removed - update this baseline deliberately)")
        Write-Output '  ScalarFuncOpFor NOT FOUND'
    } else {
        foreach ($line in $sqlBody) {
            $m = [regex]::Match($line, 'ScalarFuncOpFor\s*=\s*"([^"]*)"')
            if (-not $m.Success) { continue }
            $target = $m.Groups[1].Value
            $ruleFexamined++
            if ($arityOf.ContainsKey($target)) {
                Write-Output ("  -> {0,-10} ok - tabled, arity {1}" -f $target, $arityOf[$target])
            } else {
                Write-Output ("  -> {0,-10} NOT IN TABLE" -f $target)
                $failures.Add("VLA_Sql's ScalarFuncOpFor maps a keyword to '$target', which is not in $tableProc - ArithOpArity would answer 0 and the parser would read the call as BINARY")
            }
        }
        if ($ruleFexamined -eq 0) {
            $failures.Add('rule F found no keyword mappings in ScalarFuncOpFor - it examined nothing')
            Write-Output '  EXAMINED NOTHING'
        }
    }
}

Write-Output ''
if ($failures.Count -eq 0) {
    Write-Output '=== CHECK: clean - one operator set, and the table, the prose and the shared substrate all agree on it ==='
} else {
    Write-Output "=== CHECK: $($failures.Count) disagreement(s) about the arithmetic operator set ==="
    $failures | ForEach-Object { Write-Output "  $_" }
    Write-Output ''
    Write-Output 'ComputeArithmetic is shared by PROLOG, SQL and DATALOG. An operator that exists in one place and not another does not fail loudly - it produces a number, or an Empty, that looks like an answer.'
}
exit ([Math]::Min($failures.Count, 1))
