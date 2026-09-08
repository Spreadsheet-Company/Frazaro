<#
check_emitter_coverage.ps1 - AS.2's own coverage report.

"Which Select Case arms no pin exercises, per backend" - BETA_ROADMAP.md's
AS.2. Extends AS.8's own proven approach (check_backend_parity.ps1) rather
than inventing a new one: same pin-detection over VLA_Tests*.bas
(AssertVla/TryTranspile/VlaTranspile for the emitter, VlaInterpret/
VlaEvalExpression for the interpreter, same logical-statement-joining so a
VBA line-continuation doesn't split a pin in two). The new piece is the
DENOMINATOR: not VLA_HeadTable.bas's ~65-row catalog (AS.8's own source -
not 1:1 with Select Case arms, since some arms are aliases or purely
internal), but the literal Case arms in six specific dispatch functions,
read directly from their own bounded source text.

WHY THESE SIX, NOT EVERY "Case " IN EITHER FILE: a blind search returns
125 hits in VLA.bas, 175 in VLA_Interpreter.bas - almost entirely noise
for this question. Excluded, by name: tokenizer character dispatch
(Case " ", Case "("); the LISTOPS/quasiquote sub-dispatch inside
Substitute (macro-expansion-time, a different subsystem); ResolveExcel-
Constant's own lookup table (xlleft -> -4131, dozens of arms - a data
table, not a coverage question); small helpers (OpDisplay, EmitParams's
byval/byref modifiers); EmitFormula (mostly a REFUSAL list - which forms
formula mode explicitly rejects - success-path coverage doesn't apply the
same way, mirrors AS.1's own test-fail-is-separate precedent).

THE FUNCTIONS: EmitTop/EmitStmt/EmitExpr (VLA.bas, emitter); ExecTop/
ExecStmt/EvalExpr/EvalDynamicHead/TryEvalBuiltin/TryRuntimeHelper
(VLA_Interpreter.bas, interpreter) - ExecTop has no Select Case of its
own at all, reported explicitly every run rather than silently skipped:
the emitter splits module-level declarations (EmitTop) from statement
execution (EmitStmt); the interpreter's ExecTop doesn't mirror that split
structurally, only in what the two backends ultimately accept.
TryEvalBuiltin/TryRuntimeHelper were missed on this item's own first
scoping pass (a manually-bounded awk range check silently absorbed their
arms into what looked like one large EvalDynamicHead) - this script's own
function-boundary detection, driven by real Sub/Function declarations
rather than a hand-picked line range, caught the mistake before it ever
shipped. A third sibling in that same span, ResolveGlobalReceiver, is
NOT included - see the $dispatchFuncs comment below for why the
head-position heuristic is structurally wrong for it specifically, not
merely imprecise.

A COMMA-GROUPED ARM (Case "exit-for", "exit-do") is one code path and is
counted as ONE arm here, covered if ANY of its literals appears in head
position in the pin corpus.

WHAT COUNTS AS A "PIN": identical to check_backend_parity.ps1's own
contract - see that script's header for the full reasoning. Two risks
specific to THIS heuristic, stated up front rather than found mid-build:
(1) ResolveHeadAlias remaps some spellings before the Select Case ever
runs, so a pin written with an alias spelling can fail to head-position-
match the canonical arm string - a possible false negative, not a false
claim of coverage. (2) This extracts EVERY "Case "..."" line within a
function's own bounded text, including any Select Case nested inside it
for an unrelated sub-dispatch (e.g. EvalExpr's own error-object member-
access arms, "description"/"number"/"source") - a slight over-count of
the true head-dispatch-only arm total, stated plainly rather than hidden,
the same tolerance AS.8's own heuristic already carries elsewhere in this
project. Good enough for a coverage REPORT, not a parser.

Usage:  pwsh -File tools/check_emitter_coverage.ps1
#>

param(
    # CO.6 additions, both off by default. With neither passed this
    # script prints and exits exactly as it always has - the ratchet's
    # own behaviour is untouched, and that was verified by diffing a
    # full run against a baseline captured before this param block
    # existed.
    #
    # -ListArms       print EVERY dispatch arm, one per line, as
    #                 "<function><TAB><arm>", instead of the coverage
    #                 report. CO.6's grammar-since seed needs the arm
    #                 INVENTORY, not the uncovered subset, and reusing
    #                 this file's own Get-CaseArmGroups is the whole
    #                 point: a second copy of that parsing in a seeder
    #                 script is exactly the divergence this project
    #                 keeps getting burned by.
    # -SourceDir      read VLA.bas / VLA_Interpreter.bas from here
    #                 instead of <repo>/src, so the same parser can be
    #                 pointed at a historical tag extracted to a temp
    #                 directory and date each arm. Named SourceDir, NOT
    #                 SrcDir: PowerShell variables are case-insensitive,
    #                 so a -SrcDir parameter would silently BE the
    #                 existing $srcDir rather than override it.
    [switch]$ListArms,
    [string]$SourceDir
)

$ErrorActionPreference = 'Stop'

$repoRoot    = Split-Path -Parent $PSScriptRoot
$srcDir      = if ($SourceDir) { $SourceDir } else { Join-Path $repoRoot 'src' }
$emitterFile = Join-Path $srcDir 'VLA.bas'
$interpFile  = Join-Path $srcDir 'VLA_Interpreter.bas'
$testFiles   = @('VLA_Tests.bas', 'VLA_Tests_Grammar.bas', 'VLA_Tests_Host.bas') |
               ForEach-Object { Join-Path $srcDir $_ } | Where-Object { Test-Path $_ }

# --- Pass 1: bound each named dispatch function's own source text -----------
function Get-FunctionLines([string[]]$lines, [string]$funcName) {
    $startIdx = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match "^(Private|Public)\s+(Function|Sub)\s+$([regex]::Escape($funcName))\(") {
            $startIdx = $i
            break
        }
    }
    if ($startIdx -lt 0) {
        Write-Error "Function/Sub '$funcName' not found - has it been renamed or moved? Update `$dispatchFuncs."
    }
    $endIdx = $lines.Count - 1
    for ($i = $startIdx + 1; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^(Private|Public)\s+(Function|Sub)\s') {
            $endIdx = $i - 1
            break
        }
    }
    return $lines[$startIdx..$endIdx]
}

# --- Pass 2: join VBA line-continuations into logical statements -------------
# check_backend_parity.ps1's own Get-LogicalStatements, applied here to a
# bounded function body instead of a whole file.
function Get-LogicalStatements([string[]]$lines) {
    $stmts = New-Object System.Collections.Generic.List[string]
    $buf = ''
    foreach ($line in $lines) {
        $buf += $line
        if ($line -match '_\s*$') {
            $buf += "`n"
            continue
        }
        $stmts.Add($buf)
        $buf = ''
    }
    if ($buf.Length -gt 0) { $stmts.Add($buf) }
    return $stmts
}

# --- Pass 3: extract Case arm-groups from a function's own logical stmts ----
# Each element of the returned list is itself a list of one-or-more string
# literals (a comma-grouped arm) - "exit-for", "exit-do" is ONE arm-group,
# not two.
function Get-CaseArmGroups([string[]]$logicalStmts) {
    $groups = New-Object System.Collections.Generic.List[string[]]
    foreach ($stmt in $logicalStmts) {
        if ($stmt -notmatch '^\s*Case\s+"') { continue }
        $lits = [regex]::Matches($stmt, '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
        if ($lits.Count -gt 0) { $groups.Add(@($lits)) }
    }
    return $groups
}

function Test-HeadPosition([string]$blob, [string]$symbol) {
    $esc = [regex]::Escape($symbol)
    return [regex]::IsMatch($blob, '\(' + $esc + '[\s\)]')
}

# --- Pins: identical contract to check_backend_parity.ps1 --------------------
$emitterCallNames = @('AssertVla\(', 'TryTranspile\(', 'VlaTranspile\(')
$interpCallNames  = @('VlaInterpret\(', 'VlaEvalExpression\(')
$emitterRe = ($emitterCallNames -join '|')
$interpRe  = ($interpCallNames  -join '|')

function Get-LogicalStatementsFromFile([string]$path) {
    return Get-LogicalStatements (Get-Content -LiteralPath $path)
}

$emitterText = New-Object System.Text.StringBuilder
$interpText  = New-Object System.Text.StringBuilder
foreach ($f in $testFiles) {
    foreach ($stmt in (Get-LogicalStatementsFromFile $f)) {
        if ($stmt -match $emitterRe) { [void]$emitterText.AppendLine($stmt) }
        if ($stmt -match $interpRe)  { [void]$interpText.AppendLine($stmt) }
    }
}
$emitterBlob = $emitterText.ToString()
$interpBlob  = $interpText.ToString()

# --- The six dispatch functions, per backend ---------------------------------
$emitterLines = Get-Content -LiteralPath $emitterFile
$interpLines  = Get-Content -LiteralPath $interpFile

$dispatchFuncs = @(
    @{ Name = 'EmitTop';         File = 'VLA.bas';             Lines = $emitterLines; Blob = $emitterBlob },
    @{ Name = 'EmitStmt';        File = 'VLA.bas';             Lines = $emitterLines; Blob = $emitterBlob },
    @{ Name = 'EmitExpr';        File = 'VLA.bas';             Lines = $emitterLines; Blob = $emitterBlob },
    @{ Name = 'ExecTop';              File = 'VLA_Interpreter.bas'; Lines = $interpLines; Blob = $interpBlob },
    @{ Name = 'ExecStmt';             File = 'VLA_Interpreter.bas'; Lines = $interpLines; Blob = $interpBlob },
    @{ Name = 'EvalExpr';             File = 'VLA_Interpreter.bas'; Lines = $interpLines; Blob = $interpBlob },
    @{ Name = 'EvalDynamicHead';      File = 'VLA_Interpreter.bas'; Lines = $interpLines; Blob = $interpBlob },
    # Found by this script's own correct boundary detection, not the
    # session's first-pass manual read: EvalDynamicHead's true body ends
    # at its own next function, well short of where a mis-bounded awk
    # range check assumed - TryEvalBuiltin/TryRuntimeHelper (same shape,
    # Select Case on a known string reached from head position) sit
    # between there and EvalPositionalArgs.
    # NOT included, deliberately: ResolveGlobalReceiver, one function
    # earlier in that same span - checked its own call sites (all four)
    # directly, and its "name" argument is always sliced out of a DOTTED
    # STRING (Left$(plainName, dotAt-1)), never read from a form's own
    # head position the way every function above is. The head-position
    # heuristic is structurally wrong for it, not merely imprecise - it
    # would report false negatives across this function's entire arm
    # list, not real gaps, so it's excluded rather than left in to
    # mislead. A genuine denominator for it would need a different
    # heuristic (a bare dotted-atom scan), scoped separately if ever
    # needed - not silently folded into this one.
    @{ Name = 'TryEvalBuiltin';        File = 'VLA_Interpreter.bas'; Lines = $interpLines; Blob = $interpBlob },
    @{ Name = 'TryRuntimeHelper';      File = 'VLA_Interpreter.bas'; Lines = $interpLines; Blob = $interpBlob }
)

if (-not $ListArms) {
    Write-Output '=== EMITTER-CASE COVERAGE (heuristic, head-position text scan) ==='
}
Write-Output ''

$totalArms = 0
$totalUncovered = 0
foreach ($fn in $dispatchFuncs) {
    $body = Get-FunctionLines $fn.Lines $fn.Name
    $logical = Get-LogicalStatements $body
    $groups = Get-CaseArmGroups $logical

    if ($ListArms) {
        # Inventory mode: every arm, covered or not. A comma-grouped arm
        # stays ONE line joined by " | ", the same way this script
        # already reports it - the ledger dates a code path, not a
        # spelling.
        foreach ($group in $groups) {
            Write-Output ("{0}`t{1}" -f $fn.Name, ($group -join ' | '))
        }
        continue
    }

    if ($groups.Count -eq 0) {
        Write-Output "--- $($fn.Name) ($($fn.File)): no Select Case dispatch found (0 arms) ---"
        Write-Output ''
        continue
    }

    $uncovered = New-Object System.Collections.Generic.List[string]
    foreach ($group in $groups) {
        $covered = $false
        foreach ($lit in $group) {
            if (Test-HeadPosition $fn.Blob $lit) { $covered = $true; break }
        }
        if (-not $covered) { $uncovered.Add(($group -join ' | ')) }
    }

    $totalArms += $groups.Count
    $totalUncovered += $uncovered.Count

    Write-Output "--- $($fn.Name) ($($fn.File)): $($uncovered.Count)/$($groups.Count) arm(s) with no pin ---"
    $uncovered | ForEach-Object { Write-Output "  $_" }
    Write-Output ''
}

if (-not $ListArms) {
    $totalCovered = $totalArms - $totalUncovered
    Write-Output "=== SUMMARY: $totalCovered/$totalArms dispatch arms have a pin (per-function breakdown above) ==="
}
