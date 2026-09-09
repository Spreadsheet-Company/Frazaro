<#
check_prolog_form_attribution.ps1 - PROLOG.7's mechanical check.

WHY THIS EXISTS. PROLOG.5.1 shipped arithmetic as exactly one form,
`(is Var Expr)`, so the refusals its evaluator raises were written naming
that form in their own prose: "(is ...) expected a number but found
'{value}', which isn't one." PROLOG.7 gives that SAME evaluator a second
caller - the six comparison goals (<, >, =<, >=, =:=, =\=) - and every one
of those refusals then fires for a user who never wrote `(is ...)` at all.
Being told about a form you did not write is a confidently wrong answer,
which this project holds to be worse than a crash (IN.15's own finding).

WHAT IT CHECKS. Every message id raised from inside the shared arithmetic
procedures must name the offending form through the {form} placeholder,
filled in by whichever caller is actually running, rather than hard-coding
one form's spelling into its own text. Two conditions per message, both
required:

  1. the text contains the {form} placeholder; and
  2. the text contains no hard-coded form mention - a parenthesised head
     followed by an ellipsis, as in "(is ...)". A worked example such as
     "(+ X Y)" carries no ellipsis and is deliberately NOT matched, so a
     message may still teach with a concrete example.

WHY A SCRIPT AND NOT A HAND COUNT. PROLOG.7's own roadmap entry named two
such messages. There are five: the entry missed prolog-arith-divide-by-zero
outright, and both parse-time refusals (prolog-arith-unknown-operator,
prolog-arith-wrong-arity), which a comparison reaches the moment its own
operands are validated the way `(is ...)`'s already are. On IN.15 an
identical hand count was off by eight of sixteen, and a check written
BEFORE the fix is the only reason that was caught. This one was written
first too, and its first run listed exactly the work to do.

The scanned-procedure list below is the reviewable baseline. It is
deliberately a list of NAMES, not a pattern: a procedure that starts
raising these refusals must be added here on purpose, and a rename breaks
the run loudly rather than silently scanning nothing.

House style (tools/check_*.ps1): PowerShell 5.1, host-independent, a
hard-coded reviewable baseline, no live Excel needed. Never wired into
VlaSelfTest - it reads source text, not runtime behaviour.

Usage:  powershell -File tools\check_prolog_form_attribution.ps1
Exit code: 0 clean, 1 if any shared-arithmetic refusal names a form itself.
#>

$ErrorActionPreference = 'Stop'

$repoRoot     = Split-Path -Parent $PSScriptRoot
$prologPath   = Join-Path $repoRoot 'src\VLA_Prolog.bas'
$messagesPath = Join-Path $repoRoot 'src\VLA_Messages.bas'

# ---- baseline: the procedures whose refusals must stay form-neutral ----
# Both are reached by `(is ...)` today and by a comparison goal from
# PROLOG.7 onward. EvalArithTerm is the runtime walk; ValidateArithExpr is
# the parse-time shape check that runs over the same operand expressions.
$scannedProcs = @(
    'ValidateArithExpr',
    'EvalArithTerm'
)

# ---- baseline: ids that serve more than one form without being raised --
# ---- from inside the shared evaluator, so the scan above cannot find them
# prolog-comparison-bad-shape is raised by ValidateBodyItem's own
# comparison arm, which serves all six operators from one Case. It is
# listed here rather than by adding ValidateBodyItem to the scan above,
# because that procedure ALSO raises prolog-is-bad-shape and its
# not/findall siblings - and those are correctly form-specific, each
# raised by one arm about one form. Scanning the whole procedure would
# flag them wrongly; naming this id keeps the distinction deliberate.
#
# prolog-unification-bad-shape joins on the identical terms at PROLOG.8:
# raised by ValidateBodyItem's own term-matching arm, which serves all
# four of =, \=, == and \== from one Case, and correspondingly unable to
# name a form of its own.
# prolog-type-test-bad-shape joins on the identical terms at PROLOG.9:
# raised by ValidateBodyItem's own type-test arm, which serves all SIX of
# var, nonvar, atom, number, atomic and compound from one Case. Six forms
# through one raise site is the widest fan-out of the three, so naming a
# form in its own text would be wrong five times out of six.
#
# prolog-type-test-iso-spelling serves the six BARE ISO spellings from
# one raise site in SolveGoalList - it exists to tell a Prolog author
# that (atom X) is written (atom? X) here - so it too must take the form
# from its caller rather than naming one of the six.
#
# prolog-list-bad-shape and prolog-list-not-a-list join at PROLOG.13, and
# they are the widest fan-out yet: both are raised for all SIX list goals
# (length, member, nth, append, reverse, sum-list) from one site each -
# ValidateBodyItem's own list arm for the arity refusal, ListTermToItems
# for the not-a-proper-list one. The six do not even share an arity
# (length/2, member/2, nth/3, append/3, reverse/2, sum-list/2), so the
# arity message must take BOTH its form and its count from the caller;
# naming a form in either text would be wrong five times out of six.
$multiFormIds = @(
    'prolog-comparison-bad-shape',
    'prolog-unification-bad-shape',
    'prolog-type-test-bad-shape',
    'prolog-type-test-iso-spelling',
    'prolog-list-bad-shape',
    'prolog-list-not-a-list'
)

$formToken          = '{form}'
$literalFormPattern = '\([^)]*\.\.\.\)'

$failures = New-Object System.Collections.Generic.List[string]

foreach ($p in @($prologPath, $messagesPath)) {
    if (-not (Test-Path -LiteralPath $p)) {
        Write-Output "FAIL: missing source file $p"
        exit 1
    }
}

# ---- 1. pull each scanned procedure's body ------------------------------
# Whole-line VBA comments are dropped first: these procedures carry heavy
# prose, and a message id quoted inside a comment is documentation, not a
# raise site.
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

$prologLines = Get-Content -LiteralPath $prologPath
$raisedBy    = @{}   # id -> list of procedures that raise it

Write-Output '=== PROLOG.7 - shared arithmetic refusals must not name a form ==='
Write-Output "Source:   $prologPath"
Write-Output "Messages: $messagesPath"
Write-Output ''
Write-Output '--- scanned procedures (the reviewable baseline) ---'

foreach ($proc in $scannedProcs) {
    $body = Get-ProcBody -Lines $prologLines -Name $proc
    if ($null -eq $body) {
        $failures.Add("procedure '$proc' not found in VLA_Prolog.bas (renamed or removed - update this script's baseline deliberately)")
        Write-Output ("  {0,-20} NOT FOUND" -f $proc)
        continue
    }
    $ids = @($body |
        Select-String -Pattern 'RaiseMsg\s+"([^"]+)"' -AllMatches |
        ForEach-Object { $_.Matches } |
        ForEach-Object { $_.Groups[1].Value } |
        Sort-Object -Unique)
    Write-Output ("  {0,-20} {1,3} line(s), {2} distinct refusal id(s)" -f $proc, $body.Count, $ids.Count)
    foreach ($id in $ids) {
        if (-not $raisedBy.ContainsKey($id)) { $raisedBy[$id] = New-Object System.Collections.Generic.List[string] }
        $raisedBy[$id].Add($proc)
    }
}

if ($raisedBy.Keys.Count -eq 0) {
    $failures.Add('no refusal ids found in any scanned procedure - the scan matched nothing, which is itself a failure')
}

# Each declared multi-form id must actually be raised somewhere in
# VLA_Prolog.bas: a baseline entry for an id no code raises any more is
# stale, and silently carrying it would let this list rot into decoration.
$prologText = $prologLines -join "`n"
foreach ($id in $multiFormIds) {
    if ($prologText -notmatch ('RaiseMsg\s+"' + [regex]::Escape($id) + '"')) {
        $failures.Add("$id - named in this script's multi-form baseline but raised nowhere in VLA_Prolog.bas (stale entry)")
        continue
    }
    if (-not $raisedBy.ContainsKey($id)) { $raisedBy[$id] = New-Object System.Collections.Generic.List[string] }
    $raisedBy[$id].Add('declared multi-form')
}

# ---- 2. resolve each id to its catalogue text ---------------------------
# errNum is matched as [^,]+ rather than \d+ : one AddMsg in the catalogue
# passes a named constant (vla-interpreter-only-handler) instead of a
# literal, and a parser that assumed digits would silently skip it.
$msgText = @{}
foreach ($line in (Get-Content -LiteralPath $messagesPath)) {
    if ($line -match '^\s*AddMsg\s+m,\s*"([^"]+)"\s*,\s*[^,]+,\s*"[^"]*"\s*,\s*"(.*)"\s*$') {
        $msgText[$Matches[1]] = $Matches[2]
    }
}

Write-Output ''
Write-Output '--- refusals reachable from more than one form ---'

foreach ($id in ($raisedBy.Keys | Sort-Object)) {
    $procs = ($raisedBy[$id] | Sort-Object -Unique) -join ', '

    if (-not $msgText.ContainsKey($id)) {
        $failures.Add("$id - raised in $procs but not defined in VLA_Messages.bas")
        Write-Output ("  {0,-34} UNDEFINED  (raised in {1})" -f $id, $procs)
        continue
    }

    $text     = $msgText[$id]
    $hasForm  = $text.Contains($formToken)
    $literals = @([regex]::Matches($text, $literalFormPattern) | ForEach-Object { $_.Value } | Sort-Object -Unique)

    $problems = @()
    if (-not $hasForm)         { $problems += "no $formToken placeholder" }
    if ($literals.Count -gt 0) { $problems += ('hard-codes ' + ($literals -join ' ')) }

    if ($problems.Count -eq 0) {
        Write-Output ("  {0,-34} ok         ({1})" -f $id, $procs)
    } else {
        $why = $problems -join '; '
        Write-Output ("  {0,-34} NAMES A FORM: {1}" -f $id, $why)
        Write-Output ("      text: {0}" -f $text)
        $failures.Add("$id - $why (raised in $procs)")
    }
}

Write-Output ''
if ($failures.Count -eq 0) {
    Write-Output '=== CHECK: clean - every shared arithmetic refusal names its form through {form}, never in its own text ==='
} else {
    Write-Output "=== CHECK: $($failures.Count) refusal(s) name a form they cannot know ==="
    $failures | ForEach-Object { Write-Output "  $_" }
    Write-Output ''
    Write-Output 'These procedures serve more than one Prolog form. A refusal that spells one form into its own text tells a user about a form they may never have written. Take the form name as a {form} argument from the caller instead.'
}
exit ([Math]::Min($failures.Count, 1))
