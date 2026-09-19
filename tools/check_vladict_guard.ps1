<#
check_vladict_guard.ps1 - DATALOG.13's mechanical pin: no TypeName in the
dictionary wrappers every engine reads.

DATALOG.13: DATALOG.12's live passes measured `TypeName(d) = "Dictionary"`
at 0.148 ms a call on a late-bound Scripting.Dictionary - as much as
creating one - and a DATALOG row made about a dozen such calls, 85% of its
cost. `TypeOf d Is Collection` measures 0.000 ms. The fix inverted every
guard to test for the fallback instead.

WHY A STATIC SCAN: the regression this guards against is SILENT. A
`TypeName` put back into one of these procedures gives exactly the same
answers, so the pure, DSL and host suites all stay green; only a live
timing on a real-sized Table would notice, and nobody runs one by habit.

WHAT IS CHECKED, per procedure in the list below:
  - the procedure exists in its module (a rename must not disarm the pin);
  - no non-comment line in its body calls `TypeName(`;
  - its body carries the expected guard, verbatim. The VlaDict family's
    guard is `If VlaDictIsFallback(d) Then`, and that helper's own is
    `If d Is Nothing Then`, tested BEFORE its TypeOf: Nothing keeps the
    fallback it always had (VlaDictHas answers False, VlaDictGet misses
    loudly), which VLA_Interpreter's module-scope read relies on. The
    join index is never handed Nothing, so its guard is
    `TypeOf idx Is Collection` alone;
  - no body joins `Is Nothing` and `TypeOf` with `Or` or `And` on one
    line. TypeOf RAISES error 91 on Nothing, and VBA evaluates every
    operand, so that one-liner raises on the very Nothing it names - the
    first build shipped it and the owner's live pass caught it
    (2026-09-18, eval of an unbound name in a fresh session).

WHAT IS NOT CHECKED: VlaCount (VLA_Runtime.bas). It asks whether a value is
ANY countable object, a Dictionary included, and TypeOf cannot name a
late-bound Scripting.Dictionary without a reference, so it keeps its
TypeName by decision (DATALOG.13's scoping, 2026-09-18). Nor any other
TypeName in the repository: most ask of values that are not COM objects,
and none is on a per-row path that has been measured.

BASELINE: hand-maintained below, in the reviewable shape of
check_raise_ratchet.ps1's `$ceilings`. Adding a procedure is a deliberate,
reviewed act.

NOT wired into VlaSelfTest, the same reasoning the other tools/check_*.ps1
scans give: a step a human runs, not a per-run gate.

Usage:  powershell -File tools\check_vladict_guard.ps1 [-SrcDir <folder>]
Exit code: 0 if every listed procedure holds its guard and no TypeName;
1 otherwise.
#>

param(
    [string]$SrcDir = ''
)

$ErrorActionPreference = 'Stop'

if ($SrcDir -eq '') {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $SrcDir   = Join-Path $repoRoot 'src'
}

# --- The guarded procedures, hand-maintained. ---
# 2026-09-18, DATALOG.13: the five VlaDict wrappers, the helper they ask,
# and the join index's two readers of it.
$vlaDictGuard = 'If VlaDictIsFallback(d) Then'
$helperGuard  = 'If d Is Nothing Then'
$joinGuard    = 'TypeOf idx Is Collection'
$guarded = @(
    @{ Module = 'VLA_Runtime.bas';  Proc = 'VlaDictIsFallback'; Guard = $helperGuard },
    @{ Module = 'VLA_Runtime.bas';  Proc = 'VlaDictSet';    Guard = $vlaDictGuard },
    @{ Module = 'VLA_Runtime.bas';  Proc = 'VlaDictGet';    Guard = $vlaDictGuard },
    @{ Module = 'VLA_Runtime.bas';  Proc = 'VlaDictHas';    Guard = $vlaDictGuard },
    @{ Module = 'VLA_Runtime.bas';  Proc = 'VlaDictKeys';   Guard = $vlaDictGuard },
    @{ Module = 'VLA_Runtime.bas';  Proc = 'VlaDictPairs';  Guard = $vlaDictGuard },
    @{ Module = 'VLA_Relation.bas'; Proc = 'JoinIndexPut';  Guard = $joinGuard },
    @{ Module = 'VLA_Relation.bas'; Proc = 'JoinIndexGet';  Guard = $joinGuard }
)

$endPattern = '^\s*End\s+(Sub|Function|Property)\b'

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

Write-Output '=== VLADICT GUARD PIN (DATALOG.13) ==='
Write-Output ''

foreach ($g in $guarded) {
    $path = Join-Path $SrcDir $g.Module
    if (-not (Test-Path -LiteralPath $path)) {
        $failed.Add("$($g.Module): file not found at $path")
        continue
    }
    if (-not $cache.ContainsKey($g.Module)) {
        $cache[$g.Module] = [System.IO.File]::ReadAllLines($path)
    }
    $body = Get-ProcBody $cache[$g.Module] $g.Proc
    $label = "$($g.Module)::$($g.Proc)"
    if ($null -eq $body) {
        $failed.Add("${label}: procedure not found - renamed or removed? Update this check's list deliberately.")
        continue
    }
    $hasGuard = $false
    $typeNameLines = New-Object System.Collections.Generic.List[string]
    $oneLiners = New-Object System.Collections.Generic.List[string]
    foreach ($l in $body) {
        if (Test-IsComment $l.Text) { continue }
        if ($l.Text.Contains($g.Guard)) { $hasGuard = $true }
        if ($l.Text -match '\bTypeName\s*\(') { $typeNameLines.Add("line $($l.No): $($l.Text.Trim())") }
        if ($l.Text -match '\bIs\s+Nothing\s+(Or|And)\s+(Not\s+)?\(?\s*TypeOf\b') { $oneLiners.Add("line $($l.No): $($l.Text.Trim())") }
    }
    foreach ($o in $oneLiners) {
        $failed.Add("${label}: Is Nothing and TypeOf on one line - TypeOf raises 91 on Nothing and VBA evaluates both - $o")
    }
    if ($typeNameLines.Count -gt 0) {
        foreach ($t in $typeNameLines) { $failed.Add("${label}: calls TypeName - $t") }
    }
    if (-not $hasGuard) {
        $failed.Add("${label}: the guard '$($g.Guard)' is missing")
    }
    if ($typeNameLines.Count -eq 0 -and $oneLiners.Count -eq 0 -and $hasGuard) {
        Write-Output "  ok    $label"
    } else {
        Write-Output "  FAIL  $label"
    }
}

Write-Output ''
if ($failed.Count -gt 0) {
    Write-Output '=== CHECK: FAILED ==='
    foreach ($f in $failed) { Write-Output "  $f" }
    exit 1
}
Write-Output "=== CHECK: clean - all $($guarded.Count) guarded procedure(s) test for the fallback with TypeOf, and none calls TypeName ==="
exit 0
