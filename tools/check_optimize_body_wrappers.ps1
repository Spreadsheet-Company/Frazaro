<#
=== OPTIMIZE.2 - the one thing VLA_Optimize.bas has to know about
    DATALOG's body grammar, held to DATALOG's own copy of it ===

WHY THIS EXISTS. OPTIMIZE.2 rewrites each (forbid ...)/(require ...)
into an ordinary DATALOG rule whose head collects the rows that break
it. To build that head it has to know which body items BIND a variable
and which do not - because CheckRuleSafety refuses a rule head naming a
variable no positive body atom bound, and the rule in question is one
this engine wrote, not one the user did. So VLA_Optimize.CollectBoundVars
carries a copy of DATALOG's own body-wrapper list:

    not, count, sum, let, textjoin,
    >, <, <=, >=, =, <>,
    text-starts-with, text-ends-with, text-contains

and treats everything else as a positive atom.

THE DEFECT THIS CATCHES. DATALOG gains a fifteenth wrapper - some later
item's own aggregate or test - and nobody edits VLA_Optimize.bas.
OPTIMIZE then reads that wrapper as a positive atom, puts its inner
variables into a generated head, and the user meets
`datalog-unsafe-head-variable` naming `vla-check-2`, about a rule they
never wrote, on a program that is perfectly well formed. Two independent
lists that must agree, with nothing mechanical holding them together, is
the actual defect - check_devrig_mods_parity.ps1's own words, in a new
place.

WHAT IS CHECKED

  1. The wrapper words VLA_Datalog.bas's rule-body Select Case names -
     read from the Case lines between the `wrapperWord` assignment and
     that Select Case's End Select - are EXACTLY the words
     VLA_Optimize.bas's CollectBoundVars names across its two
     non-positive arms.

  2. CheckConsequent names the same set, since a require's consequent
     is held to "a row, not a test" against the identical list.

  2a. OPTIMIZE.3 slice 2 added two more copies, held to the same set:
     BodyReadKind, which reads what a body item reads (a choice program
     must know whether a constraint touches a chosen row, and where a
     negated one is), and CheckPoolAtom, which refuses a wrapper where a
     pool or a (per ...) row belongs. A fifteenth DATALOG wrapper read as
     a positive atom there would put its inner names into a grounding
     rule the user never wrote - the same defect, two new places.

  3. Each list is sorted and compared as a set, so a reordering inside
     either file is not a failure - only a word present in one and
     absent from the other.

Exit code: 0 when all five sets agree; 1 otherwise, naming each word
and the file it is missing from.

MUTATION CONTROL: -Control plants each of three defects in memory (a
word dropped from OPTIMIZE's collector, a word added to it, and a word
added to DATALOG's Select Case) and requires all three to be caught.
#>

[CmdletBinding()]
param(
    [switch]$Control
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$datalogPath = Join-Path $repo 'src\VLA_Datalog.bas'
$optimizePath = Join-Path $repo 'src\VLA_Optimize.bas'

# A VBA line continuation (a trailing " _") splits one logical line
# across several physical ones, and a Case arm listing fourteen words is
# exactly the kind that gets split. Read without joining them, this
# check's own first run reported three text tests as missing from
# VLA_Optimize.bas when they were sitting on the continuation line - the
# check reporting its own defect over a known-good tree.
# check_optimize_parity.ps1 joins them for the same reason.
function Join-VbaContinuations {
    param([string[]]$Lines)
    $joined = New-Object System.Collections.Generic.List[string]
    $pending = $null
    foreach ($line in $Lines) {
        $text = if ($null -eq $pending) { $line } else { $pending + ' ' + $line.TrimStart() }
        if ($text -match '\s_\s*$') {
            $pending = ($text -replace '\s_\s*$', '')
        } else {
            $joined.Add($text)
            $pending = $null
        }
    }
    if ($null -ne $pending) { $joined.Add($pending) }
    return , $joined.ToArray()
}

function Get-CaseWords {
    param([string[]]$Lines)
    $words = New-Object System.Collections.Generic.List[string]
    foreach ($line in $Lines) {
        $t = $line.Trim()
        if ($t -notmatch '^Case\s') { continue }
        if ($t -match '^Case\s+Else') { continue }
        foreach ($m in [regex]::Matches($t, '"([^"]*)"')) {
            $words.Add($m.Groups[1].Value)
        }
    }
    return , ($words | Sort-Object -Unique)
}

# --- 1. DATALOG's own rule-body wrapper Select Case ------------------
# Bounded by the wrapperWord assignment above it and the first End
# Select below, so a different Select Case elsewhere in the file cannot
# be read by accident.
function Get-DatalogWrappers {
    param([string]$Path)
    $lines = Get-Content -LiteralPath $Path
    $lines = Join-VbaContinuations -Lines $lines
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match 'wrapperWord\s*=\s*VLA_Identity\.Fold') { $start = $i; break }
    }
    if ($start -lt 0) { throw "could not find the wrapperWord assignment in $Path" }
    $end = -1
    for ($i = $start; $i -lt $lines.Count; $i++) {
        if ($lines[$i].Trim() -eq 'End Select') { $end = $i; break }
    }
    if ($end -lt 0) { throw "could not find the End Select after wrapperWord in $Path" }
    return Get-CaseWords -Lines $lines[$start..$end]
}

# --- 2/3. OPTIMIZE's two copies --------------------------------------
# Each is the Select Case inside one named Sub; the arms wanted are the
# ones that are NOT the positive-atom fallthrough, which is Case Else in
# CollectBoundVars and absent in CheckConsequent.
function Get-OptimizeWrappers {
    param([string]$Path, [string]$SubName)
    $lines = Get-Content -LiteralPath $Path
    $lines = Join-VbaContinuations -Lines $lines
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match "^Private (Sub|Function) $([regex]::Escape($SubName))\b") { $start = $i; break }
    }
    if ($start -lt 0) { throw "could not find Private Sub $SubName in $Path" }
    $end = -1
    for ($i = $start; $i -lt $lines.Count; $i++) {
        if ($lines[$i].Trim() -eq 'End Select') { $end = $i; break }
    }
    if ($end -lt 0) { throw "could not find the End Select inside $SubName in $Path" }
    return Get-CaseWords -Lines $lines[$start..$end]
}

function Compare-Sets {
    param([string[]]$Expected, [string[]]$Actual, [string]$ExpectedName, [string]$ActualName)
    $problems = New-Object System.Collections.Generic.List[string]
    foreach ($w in $Expected) {
        if ($Actual -notcontains $w) {
            $problems.Add("'$w' is in $ExpectedName but not in $ActualName")
        }
    }
    foreach ($w in $Actual) {
        if ($Expected -notcontains $w) {
            $problems.Add("'$w' is in $ActualName but not in $ExpectedName")
        }
    }
    return , $problems
}

$datalog = Get-DatalogWrappers -Path $datalogPath
$collector = Get-OptimizeWrappers -Path $optimizePath -SubName 'CollectBoundVars'
$consequent = Get-OptimizeWrappers -Path $optimizePath -SubName 'CheckConsequent'
$readKind = Get-OptimizeWrappers -Path $optimizePath -SubName 'BodyReadKind'
$poolAtom = Get-OptimizeWrappers -Path $optimizePath -SubName 'CheckPoolAtom'

if ($Control) {
    $cases = @(
        @{ name = 'a word dropped from the collector'; d = $datalog; o = ($collector | Where-Object { $_ -ne 'sum' }) }
        @{ name = 'a word added to the collector'; d = $datalog; o = (@($collector) + 'invented-wrapper') }
        @{ name = 'a word added to DATALOG'; d = (@($datalog) + 'text-matches'); o = $collector }
    )
    $caught = 0
    foreach ($c in $cases) {
        $p = Compare-Sets -Expected $c.d -Actual $c.o -ExpectedName 'DATALOG' -ActualName 'OPTIMIZE'
        if ($p.Count -gt 0) {
            Write-Output "CONTROL caught: $($c.name) - $($p -join '; ')"
            $caught++
        } else {
            Write-Output "CONTROL MISSED: $($c.name)"
        }
    }
    if ($caught -ne $cases.Count) {
        Write-Output "check_optimize_body_wrappers: CONTROL FAILED, $caught of $($cases.Count) caught"
        exit 1
    }
    Write-Output "check_optimize_body_wrappers: CONTROL passed, $caught of $($cases.Count) caught"
    exit 0
}

$problems = New-Object System.Collections.Generic.List[string]
foreach ($p in (Compare-Sets -Expected $datalog -Actual $collector -ExpectedName 'VLA_Datalog.bas' -ActualName 'VLA_Optimize.CollectBoundVars')) {
    $problems.Add($p)
}
foreach ($p in (Compare-Sets -Expected $datalog -Actual $consequent -ExpectedName 'VLA_Datalog.bas' -ActualName 'VLA_Optimize.CheckConsequent')) {
    $problems.Add($p)
}
foreach ($p in (Compare-Sets -Expected $datalog -Actual $readKind -ExpectedName 'VLA_Datalog.bas' -ActualName 'VLA_Optimize.BodyReadKind')) {
    $problems.Add($p)
}
foreach ($p in (Compare-Sets -Expected $datalog -Actual $poolAtom -ExpectedName 'VLA_Datalog.bas' -ActualName 'VLA_Optimize.CheckPoolAtom')) {
    $problems.Add($p)
}

if ($problems.Count -gt 0) {
    Write-Output "check_optimize_body_wrappers: FAILED"
    foreach ($p in $problems) { Write-Output "  $p" }
    exit 1
}

Write-Output "check_optimize_body_wrappers: OK - $($datalog.Count) body wrappers, and all five lists agree"
Write-Output "  $($datalog -join ' ')"
exit 0
