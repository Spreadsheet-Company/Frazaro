<#
check_engine_call_names.ps1 - U.30's mechanical pin.

U.30: a program's own name is never one the code Frazaro writes calls by
name. Before it, "To len of x:" was accepted, and inside that program every
"length of" called it: compiled VBA writes len(...) unqualified, where a
module's own procedure answers before VBA's library, and the interpreter's
EvalDynamicHead looks up the program's procedures before its builtins. An
action called trim changed every "is empty" (len(trim(...))), an action
called range every cell its program names. A variable or parameter of such
a name broke the compiled call alone, since it shadows VBA's or Excel's
inside its procedure while Interpret answers from its builtins. So
VLA_SentenceEngine's CheckName refuses a name in IsEngineCallName's list
wherever a name is made (english-engine-call-name).

WHY A STATIC SCAN: that list is a copy of four sources elsewhere in the
engine. A builtin, a function word or an Excel object added to one of them
later would reopen U.30 without a failing test, since nothing but this list
connects them.

WHAT IS CHECKED:
  A. IsEngineCallName's list equals the union of four sources, each read
     from the code:
       1. TryEvalBuiltin's Case arms (src/VLA_Interpreter.bas) - VBA's
          functions, as the interpreter answers them.
       2. RegisterBuiltinFuncWords's AddFnEntry targets
          (src/VLA_SentenceEngine.bas) - what English's own function words
          compile to.
       3. The heads EvalDynamicHead answers natively (Excel's range, cells,
          rows, columns, worksheets and workbooks, make-button, and
          application from its WorksheetFunction arms: a dotted arm gives
          its first segment) and the roots ResolveGlobalReceiver resolves
          (src/VLA_Interpreter.bas).
       4. The words the emitter writes itself: a capitalized name directly
          followed by ( or . inside a string literal of src/VLA.bas
          (Array(, Debug.Print, LBound(, UBound(, ThisWorkbook., Len(),
          less the runtime's own Vla helpers, the all-capitals Excel
          functions a deflambda formula carries, and the reviewed
          exceptions below; and a lowercase dotted root inside a string
          literal of src/VLA_SentenceEngine.bas (err.description).
     A name in a source and not in the list, and a name in the list with no
     source, are both failures, each named.
  B. HOW MANY names each source gives, pinned below, so a regex that
     silently stopped matching cannot pass A by reading nothing.
  C. Every target of a built-in function word is one the interpreter
     answers: a TryEvalBuiltin arm, or a Public procedure of
     src/VLA_Runtime.bas above its INJECT BOUNDARY (the helper manifest's
     reach). A function word whose target nothing answers runs under Run and
     stops under Interpret, naming the target: U.30 found six (absolute of,
     month of, year of, day of, hour of, minute of) and gave each an arm.

WHAT IS NOT CHECKED: the runtime's own Vla helpers, and the VBA functions a
phrasebook's rules and macros call directly (replace, timeserial, ...),
which come with the phrasebook, not the engine; and a program's action named
like a macro or a core form, which the macro or the form then takes. Both
are recorded in U.30's entry in docs/BETA_REARVIEW.md.

Host-independent: reads files only. NOT wired into VlaSelfTest - the house
ratchet shape, a step a human and CI run.

Usage:  powershell -File tools\check_engine_call_names.ps1 [-RepoRoot <dir>]
Exit code: 0 if clean; 1 if anything above fails.
#>

param(
    [string]$RepoRoot
)

$ErrorActionPreference = 'Stop'
if (-not $RepoRoot) { $RepoRoot = Split-Path -Parent $PSScriptRoot }

# --- Baseline, hand-maintained. ---
# 2026-09-30, U.30: how many distinct names each source gives. A source
# that grows fails A until its names are in the list, and then fails here
# until the count is raised - both on purpose, so that adding a name the
# generated code calls stays a reviewed act.
$pins = [ordered]@{
    'TryEvalBuiltin arms'            = 20
    'function-word targets'          = 16
    'EvalDynamicHead heads'          = 8
    'ResolveGlobalReceiver roots'    = 10
    'VLA.bas emitted words'          = 6
    'VLA_SentenceEngine dotted roots' = 1
}

# Reviewed exceptions to source 4: "<name>" = "why it is not emitted".
$notEmitted = @{
    'Scripting' = 'the transpiler''s own CreateObject("Scripting.Dictionary"), a ProgID in its code, never written into a program'
}

$interpPath  = Join-Path $RepoRoot (Join-Path 'src' 'VLA_Interpreter.bas')
$enginePath  = Join-Path $RepoRoot (Join-Path 'src' 'VLA_SentenceEngine.bas')
$vlaPath     = Join-Path $RepoRoot (Join-Path 'src' 'VLA.bas')
$runtimePath = Join-Path $RepoRoot (Join-Path 'src' 'VLA_Runtime.bas')

$failed = New-Object System.Collections.Generic.List[string]

foreach ($p in @($interpPath, $enginePath, $vlaPath, $runtimePath)) {
    if (-not (Test-Path -LiteralPath $p)) {
        Write-Output "  missing: $p"
        exit 1
    }
}

function Remove-VbaComment([string]$line) {
    # The code part of a VBA line: text before a ' that is not inside a
    # string. A line that is all comment comes back empty.
    $inString = $false
    for ($k = 0; $k -lt $line.Length; $k++) {
        $c = $line[$k]
        if ($c -eq '"') { $inString = -not $inString; continue }
        if ($c -eq "'" -and -not $inString) { return $line.Substring(0, $k) }
    }
    return $line
}

function Get-StringLiterals([string]$line) {
    # The string literals on the code part of a VBA line, "" read as one ".
    $lits = New-Object System.Collections.Generic.List[string]
    $inString = $false
    $cur = New-Object System.Text.StringBuilder
    for ($k = 0; $k -lt $line.Length; $k++) {
        $c = $line[$k]
        if ($c -eq '"') {
            if ($inString -and ($k + 1) -lt $line.Length -and $line[$k + 1] -eq '"') {
                [void]$cur.Append('"'); $k++; continue
            }
            if ($inString) { $lits.Add($cur.ToString()); [void]$cur.Clear() }
            $inString = -not $inString
            continue
        }
        if (-not $inString -and $c -eq "'") { break }
        if ($inString) { [void]$cur.Append($c) }
    }
    return ,$lits.ToArray()
}

function Get-ProcCode([string[]]$lines, [string]$name) {
    # The code lines (comments removed) of the procedure called $name, or
    # $null when there is none.
    $head = '^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Sub|Function)\s+' + [regex]::Escape($name) + '\s*\('
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match $head) { $start = $i; break }
    }
    if ($start -lt 0) { return $null }
    $code = New-Object System.Collections.Generic.List[string]
    for ($i = $start + 1; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*End\s+(?:Sub|Function)\b') { break }
        $c = Remove-VbaComment $lines[$i]
        if ($c.Trim().Length -gt 0) { $code.Add($c) }
    }
    return ,$code.ToArray()
}

function Get-QuotedCaseArms([string[]]$code) {
    # Every quoted literal on a Case line, in order, folded.
    $arms = New-Object System.Collections.Generic.List[string]
    foreach ($c in $code) {
        if ($c -match '^\s*Case\s+"') {
            foreach ($m in [regex]::Matches($c, '"([^"]+)"')) { $arms.Add($m.Groups[1].Value.ToLowerInvariant()) }
        }
    }
    return ,$arms.ToArray()
}

function New-NameSet { return New-Object System.Collections.Generic.SortedSet[string] }

Write-Output '=== NAMES THE GENERATED CODE CALLS ARE NEVER A PROGRAM''S NAMES (U.30) ==='

$interpLines  = [IO.File]::ReadAllText($interpPath) -split "`r?`n"
$engineLines  = [IO.File]::ReadAllText($enginePath) -split "`r?`n"
$vlaLines     = [IO.File]::ReadAllText($vlaPath) -split "`r?`n"
$runtimeLines = [IO.File]::ReadAllText($runtimePath) -split "`r?`n"

$sources = [ordered]@{}

# --- 1. TryEvalBuiltin's arms. ---
$teb = Get-ProcCode $interpLines 'TryEvalBuiltin'
$s1 = New-NameSet
if ($null -eq $teb) { $failed.Add('TryEvalBuiltin not found in src/VLA_Interpreter.bas') }
else { foreach ($a in (Get-QuotedCaseArms $teb)) { [void]$s1.Add($a) } }
$sources['TryEvalBuiltin arms'] = $s1

# --- 2. RegisterBuiltinFuncWords's targets. ---
$rbf = Get-ProcCode $engineLines 'RegisterBuiltinFuncWords'
$s2 = New-NameSet
$words = New-Object System.Collections.Generic.List[object]
if ($null -eq $rbf) { $failed.Add('RegisterBuiltinFuncWords not found in src/VLA_SentenceEngine.bas') }
else {
    foreach ($c in $rbf) {
        $m = [regex]::Match($c, 'AddFnEntry\s+mFn(?:Of|Nullary)\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"')
        if ($m.Success) {
            $target = $m.Groups[2].Value.ToLowerInvariant()
            [void]$s2.Add($target)
            $words.Add([pscustomobject]@{ Word = $m.Groups[1].Value.ToLowerInvariant(); Target = $target })
        }
    }
}
$sources['function-word targets'] = $s2

# --- 3. The heads EvalDynamicHead answers natively, and the global roots. ---
$edh = Get-ProcCode $interpLines 'EvalDynamicHead'
$s3 = New-NameSet
if ($null -eq $edh) { $failed.Add('EvalDynamicHead not found in src/VLA_Interpreter.bas') }
else {
    foreach ($a in (Get-QuotedCaseArms $edh)) {
        $dot = $a.IndexOf('.')
        if ($dot -gt 0) { [void]$s3.Add($a.Substring(0, $dot)) } else { [void]$s3.Add($a) }
    }
}
$sources['EvalDynamicHead heads'] = $s3

$rgr = Get-ProcCode $interpLines 'ResolveGlobalReceiver'
$s3b = New-NameSet
if ($null -eq $rgr) { $failed.Add('ResolveGlobalReceiver not found in src/VLA_Interpreter.bas') }
else { foreach ($a in (Get-QuotedCaseArms $rgr)) { [void]$s3b.Add($a) } }
$sources['ResolveGlobalReceiver roots'] = $s3b

# --- 4. The words the emitter writes itself. ---
$s4 = New-NameSet
$excepted = New-NameSet
foreach ($line in $vlaLines) {
    if ($line.TrimStart().StartsWith("'")) { continue }
    foreach ($lit in (Get-StringLiterals $line)) {
        foreach ($m in [regex]::Matches($lit, '(?<![A-Za-z0-9_.])([A-Z][A-Za-z]+)(?=[(.])')) {
            $n = $m.Groups[1].Value
            if ($n -cmatch '^[A-Z]+$') { continue }            # an Excel function in formula text
            if ($n -match '^Vla') { continue }                 # the runtime's own helpers
            if ($notEmitted.ContainsKey($n)) { [void]$excepted.Add($n); continue }
            [void]$s4.Add($n.ToLowerInvariant())
        }
    }
}
$sources['VLA.bas emitted words'] = $s4

$s4b = New-NameSet
foreach ($line in $engineLines) {
    if ($line.TrimStart().StartsWith("'")) { continue }
    foreach ($lit in (Get-StringLiterals $line)) {
        foreach ($m in [regex]::Matches($lit, '(?<![A-Za-z0-9_.\-])([a-z][a-z]+)\.[a-z]')) {
            [void]$s4b.Add($m.Groups[1].Value)
        }
    }
}
$sources['VLA_SentenceEngine dotted roots'] = $s4b

# --- The list itself. ---
$iecn = Get-ProcCode $engineLines 'IsEngineCallName'
$listSet = New-NameSet
if ($null -eq $iecn) { $failed.Add('IsEngineCallName not found in src/VLA_SentenceEngine.bas') }
else {
    $listLine = $iecn | Where-Object { $_ -match '^\s*list\s*=\s*"' } | Select-Object -First 1
    if ($null -eq $listLine) { $failed.Add('IsEngineCallName has no list = "..." line') }
    else {
        $lm = [regex]::Match($listLine, '"([^"]*)"')
        foreach ($w in ($lm.Groups[1].Value -split '\s+')) {
            if ($w.Length -gt 0) { [void]$listSet.Add($w.ToLowerInvariant()) }
        }
    }
}

# --- A and B. ---
Write-Output ''
Write-Output 'Sources (names read / pinned):'
$union = New-NameSet
foreach ($k in $sources.Keys) {
    $set = $sources[$k]
    foreach ($n in $set) { [void]$union.Add($n) }
    $want = $pins[$k]
    $mark = if ($set.Count -eq $want) { '' } else { '  FAIL' }
    Write-Output ("  {0,-34} {1,3} / {2,3}{3}" -f $k, $set.Count, $want, $mark)
    Write-Output ("      {0}" -f (($set | ForEach-Object { $_ }) -join ' '))
    if ($set.Count -ne $want) {
        $failed.Add("$k gives $($set.Count) name(s); pinned $want")
    }
}
if ($excepted.Count -gt 0) {
    Write-Output ("  reviewed, not emitted: {0}" -f (($excepted | ForEach-Object { $_ }) -join ' '))
}
$unusedExceptions = $notEmitted.Keys | Where-Object { -not $excepted.Contains($_) }
foreach ($u in $unusedExceptions) {
    $failed.Add("reviewed exception '$u' no longer matches anything in src/VLA.bas - drop it from `$notEmitted")
}

Write-Output ''
Write-Output ("IsEngineCallName's list: {0} names; the sources' union: {1}" -f $listSet.Count, $union.Count)
$missing = @($union | Where-Object { -not $listSet.Contains($_) })
$extra   = @($listSet | Where-Object { -not $union.Contains($_) })
foreach ($n in $missing) {
    $from = @($sources.Keys | Where-Object { $sources[$_].Contains($n) }) -join ', '
    Write-Output "  missing from the list: $n   (from $from)  FAIL"
    $failed.Add("'$n' is called by the generated code ($from) but IsEngineCallName does not refuse it")
}
foreach ($n in $extra) {
    Write-Output "  in the list, no source: $n  FAIL"
    $failed.Add("'$n' is in IsEngineCallName's list but no source gives it - drop it, or the scan broke")
}

# --- C. Every built-in function word's target is answered by the interpreter. ---
$helpers = New-Object System.Collections.Generic.HashSet[string]
$boundary = $false
foreach ($line in $runtimeLines) {
    if ($line -match '^\s*''\s*===\s*EN_RUNTIME INJECT BOUNDARY') { $boundary = $true; break }
    $m = [regex]::Match($line, '^\s*Public\s+(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)')
    if ($m.Success) { [void]$helpers.Add($m.Groups[1].Value.ToLowerInvariant()) }
}
if (-not $boundary) { $failed.Add('src/VLA_Runtime.bas has no EN_RUNTIME INJECT BOUNDARY line') }

Write-Output ''
Write-Output 'Built-in function words and what answers their targets under Interpret:'
foreach ($w in $words) {
    if ($s1.Contains($w.Target)) {
        $how = 'TryEvalBuiltin'
    } elseif ($helpers.Contains($w.Target)) {
        $how = 'VLA_Runtime helper'
    } else {
        $how = $null
    }
    if ($null -ne $how) {
        Write-Output ("  {0,-12} -> {1,-14} {2}" -f $w.Word, $w.Target, $how)
    } else {
        Write-Output ("  {0,-12} -> {1,-14} nothing  FAIL" -f $w.Word, $w.Target)
        $failed.Add("the function word '$($w.Word)' compiles to '$($w.Target)', which no interpreter tier answers: give TryEvalBuiltin an arm for it")
    }
}

Write-Output ''
if ($failed.Count -eq 0) {
    Write-Output "=== CHECK: clean - all $($union.Count) names the generated code calls are refused as a program's names, and every built-in function word runs under Interpret ==="
} else {
    Write-Output "=== CHECK: $($failed.Count) problem(s) ==="
    $failed | ForEach-Object { Write-Output "  $_" }
    Write-Output 'A name the code Frazaro writes calls by name must be refused wherever a program makes one (U.30 in docs/BETA_REARVIEW.md): add it to IsEngineCallName''s list in src/VLA_SentenceEngine.bas, then raise its source''s pin above.'
}
exit ([Math]::Min($failed.Count, 1))
