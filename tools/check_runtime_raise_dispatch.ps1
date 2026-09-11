<#
check_runtime_raise_dispatch.ps1 - IN.15's mechanical pin.

IN.15: `Application.Run` does not propagate a target macro's `Err.Raise` to
the caller's handler. `tools/VLA_Diag2.bas` scenario 1 proves it standalone,
with zero project state on the stack, and `VLA_Interpreter.bas`'s own
`TryRuntimeHelper` header records the live catch. So a refusal raised inside
a `VLA_Runtime` helper that the interpreter reaches through the generic
`Application.Run` tier does not arrive as a Frazaro modal - it breaks into
the VBE as "Run-time error '5'" with a Debug button, in direct violation of
`LX.8`'s refuse-in-words doctrine.

The fix is per-helper: a native `Case` in `TryRuntimeHelper` calls the
procedure directly, which puts the raise back on the ordinary VBA call stack
where `InterpretProgram`'s handler catches it. `IN.11`/`IN.12`/`SEC.8` each
added such a `Case` one at a time, by hand, after a crash was reported.

WHAT THIS SCRIPT IS FOR: making that count mechanical instead of
hand-maintained. The boundary is not arbitrary - a helper needs a native
`Case` exactly when IT CAN RAISE - and a boundary that is remembered rather
than pinned is one a ninth helper crosses silently. Written before the eight
(then sixteen; see below) `Case`s it demands, so its first run was red and
enumerated them.

WHAT IS CHECKED: every public `Vla*` procedure in `VLA_Runtime.bas` that can
raise - directly, or through a private procedure in the same module that
raises - must have a native `Case "<foldedname>"` in `TryRuntimeHelper`.

TRANSITIVE, NOT JUST DIRECT, AND THAT IS THE POINT. Scoping IN.15 by hand
counted eight helpers with a raise site in their own body. The transitive
closure finds SIXTEEN: `RequirePivotTableByName` (private) raises
`rt-pivot-not-found`, and twelve public pivot helpers call it - eight of
which have no raise site of their own and were therefore missed by the hand
count. "Pivot table name misspelled" is as reachable as any refusal in this
module. A direct-sites-only scan would have passed the day the eight landed
and left those eight crashing.

GOVERNED SET: public procedures whose name begins with `Vla`, declared ABOVE
the `EN_RUNTIME INJECT BOUNDARY`. That is not a convenience - it is the
same set `VlaHelperManifest` itself enumerates, because the manifest reads
`RuntimeSourceText()`, which trims at that boundary. Machinery below the
boundary (`VlaInjectRuntime`, `VlaHelperManifest`, `VlaRuntimeInjectText`,
`VlaRuntimeSheetName`) is build/deploy surface the interpreter never
dispatches to, is absent from the manifest, and is reported informationally
rather than demanded.

WHAT IS NOT CHECKED: that a native `Case` passes the right arguments in the
right order, or that its arity guard is correct. Those are read, and covered
by the pure suite's own dispatch tests; this script pins existence, which is
the property that was silently drifting.

A LINE IS A RAISE SITE if it is not wholly a comment and contains
`RaiseRuntimeMsg`, `RaiseMsg`, or `Err.Raise` - the same "trimmed line does
not start with an apostrophe" rule `check_raise_ratchet.ps1` documents, and
inherited deliberately so both scripts agree on what a raise looks like.

BASELINE: hand-maintained below, the reviewable shape
`check_word_automation_security.ps1`'s `$baseline` and
`check_raise_ratchet.ps1`'s `$ceilings` already use. A helper that JOINS the
raising set is not a failure by itself - it fails only if it has no `Case` -
but it is reported loudly, so growing the set stays a deliberate act.

NOT wired into VlaSelfTest, same reasoning F.12/F.14/AS.8/SEC.13 gave for
the other tools/check_*.ps1 scans: a version-close step a human runs, not a
per-run gate. Host-independent - no Excel, no COM, text only.

Usage:  powershell -File tools\check_runtime_raise_dispatch.ps1
Exit code: 0 if every raising governed helper has a native Case; 1 if any
does not.
#>

$ErrorActionPreference = 'Stop'

$repoRoot    = Split-Path -Parent $PSScriptRoot
$srcDir      = Join-Path $repoRoot 'src'
$runtimeFile = Join-Path $srcDir 'VLA_Runtime.bas'
$interpFile  = Join-Path $srcDir 'VLA_Interpreter.bas'

foreach ($f in @($runtimeFile, $interpFile)) {
    if (-not (Test-Path $f)) { Write-Error "Missing source file: $f" }
}

# --- Known raising governed helpers, hand-maintained. ---
# 2026-09-08, IN.15: sixteen. Eight raise in their own body (VlaColor,
# VlaDictGet, VlaFreezePanes, VlaFillSeries and the four pivot helpers that
# validate a kind/direction/function argument); eight more raise only
# through RequirePivotTableByName's rt-pivot-not-found. The four helpers
# that already had native Cases before this script existed
# (VlaCheckSheetName, VlaCheckSheetAbsent, VlaCheckRangeName, VlaSendMail)
# are in the set too - they are raising helpers, they simply were already
# dispatched correctly.
$baseline = @(
    'VlaCheckRangeName',
    'VlaCheckSheetAbsent',
    'VlaCheckSheetName',
    'VlaColor',
    'VlaDictGet',
    'VlaFillSeries',
    'VlaFreezePanes',
    'VlaNumberFormatCode',   # G-FORMAT slice 2, 2026-09-10: refuses a bad decimal count
    'VlaPivotAddValues',
    'VlaPivotChangeSource',
    'VlaPivotClear',
    'VlaPivotDelete',
    'VlaPivotRefresh',
    'VlaPivotRename',
    'VlaPivotSetBlankLine',
    'VlaPivotSetOrientation',
    'VlaPivotSetRowLayout',
    'VlaPivotSetShowDetail',
    'VlaPivotSetSubtotals',
    'VlaPivotSort',
    'VlaSendMail'
)

$raiseRe = 'RaiseRuntimeMsg|RaiseMsg|Err\.Raise'

function Test-IsComment([string]$line) {
    return ($line.Trim()).StartsWith("'")
}

# Parse a .bas file into procedures: name, visibility, start line, body.
function Get-Procedures([string[]]$lines) {
    $procs = @()
    $cur = $null
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $t = $lines[$i].Trim()
        if ($t -match '^(Public|Private)\s+(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)') {
            $cur = [pscustomobject]@{
                Name       = $Matches[3]
                Visibility = $Matches[1]
                Start      = $i + 1
                Body       = New-Object System.Collections.ArrayList
            }
        } elseif ($t -match '^End\s+(Sub|Function|Property)\b') {
            if ($null -ne $cur) { $procs += $cur; $cur = $null }
        } elseif ($null -ne $cur) {
            [void]$cur.Body.Add($lines[$i])
        }
    }
    return $procs
}

$runtimeLines = [IO.File]::ReadAllLines($runtimeFile)
$interpLines  = [IO.File]::ReadAllLines($interpFile)

# --- The inject boundary: the manifest's own horizon, so ours too. ---
$boundaryLine = 0
for ($i = 0; $i -lt $runtimeLines.Count; $i++) {
    if ($runtimeLines[$i] -match 'EN_RUNTIME INJECT BOUNDARY =') { $boundaryLine = $i + 1; break }
}
if ($boundaryLine -eq 0) {
    Write-Error "No 'EN_RUNTIME INJECT BOUNDARY' fence found in VLA_Runtime.bas - this script's governed-set rule depends on it; the module's shape may have changed."
}

$procs = Get-Procedures $runtimeLines
if ($procs.Count -eq 0) {
    Write-Error "No procedures parsed from VLA_Runtime.bas - this script's declaration pattern may be stale."
}

# --- Direct raise sites, per procedure ---
$direct = @{}
foreach ($p in $procs) {
    $hits = 0
    foreach ($ln in $p.Body) {
        if (-not (Test-IsComment $ln) -and $ln -match $raiseRe) { $hits++ }
    }
    $direct[$p.Name] = $hits
}

# --- Intra-module call graph (a name mentioned on a non-comment line) ---
$names = @($procs | ForEach-Object { $_.Name })
$calls = @{}
foreach ($p in $procs) {
    $edges = New-Object System.Collections.ArrayList
    $bodyText = ($p.Body | Where-Object { -not (Test-IsComment $_) }) -join "`n"
    foreach ($n in $names) {
        if ($n -eq $p.Name) { continue }
        if ($bodyText -match ('\b' + [regex]::Escape($n) + '\b')) { [void]$edges.Add($n) }
    }
    $calls[$p.Name] = $edges
}

# --- Transitive closure: can this procedure raise, directly or downstream? ---
$canRaise = @{}
foreach ($n in $names) { $canRaise[$n] = ($direct[$n] -gt 0) }
$changed = $true
while ($changed) {
    $changed = $false
    foreach ($n in $names) {
        if ($canRaise[$n]) { continue }
        foreach ($c in $calls[$n]) {
            if ($canRaise[$c]) { $canRaise[$n] = $true; $changed = $true; break }
        }
    }
}

# --- The governed set: public Vla*, above the inject boundary ---
$governed = @($procs | Where-Object {
    $_.Visibility -eq 'Public' -and $_.Name -like 'Vla*' -and $_.Start -lt $boundaryLine
})
$belowBoundary = @($procs | Where-Object {
    $_.Visibility -eq 'Public' -and $_.Name -like 'Vla*' -and $_.Start -gt $boundaryLine -and $canRaise[$_.Name]
})

# --- TryRuntimeHelper's native Case list ---
$inFn = $false
$nativeCases = New-Object System.Collections.ArrayList
foreach ($ln in $interpLines) {
    $t = $ln.Trim()
    if ($t -match '^(Public|Private)\s+Function\s+TryRuntimeHelper\b') { $inFn = $true; continue }
    if ($inFn) {
        if ($t -match '^End\s+Function\b') { break }
        if (-not (Test-IsComment $ln)) {
            $m = [regex]::Matches($t, 'Case\s+"([^"]+)"')
            foreach ($mm in $m) { [void]$nativeCases.Add($mm.Groups[1].Value.ToLower()) }
        }
    }
}
if (-not $inFn) {
    Write-Error "TryRuntimeHelper not found in VLA_Interpreter.bas - this script's Case-list parse depends on it."
}

# ---------------------------------------------------------------------
Write-Output '=== VLA_Runtime RAISE -> TryRuntimeHelper NATIVE DISPATCH (IN.15) ==='
Write-Output ("Runtime module:   {0} procedures ({1} public), inject boundary at line {2}" -f $procs.Count, @($procs | Where-Object { $_.Visibility -eq 'Public' }).Count, $boundaryLine)
Write-Output ("Governed set:     {0} public Vla* helpers above the boundary" -f $governed.Count)
Write-Output ("Native Cases:     {0} in TryRuntimeHelper" -f $nativeCases.Count)
Write-Output ''

$failures = New-Object System.Collections.ArrayList
$raising  = New-Object System.Collections.ArrayList

Write-Output '--- governed helpers that can raise ---'
foreach ($p in ($governed | Sort-Object Name)) {
    if (-not $canRaise[$p.Name]) { continue }
    [void]$raising.Add($p.Name)

    $how = 'direct'
    if ($direct[$p.Name] -eq 0) {
        $via = @($calls[$p.Name] | Where-Object { $canRaise[$_] }) -join ', '
        $how = "via $via"
    }

    $hasCase = $nativeCases -contains $p.Name.ToLower()
    if ($hasCase) {
        Write-Output ("  ok    {0,-24} line {1,5}  {2}" -f $p.Name, $p.Start, $how)
    } else {
        Write-Output ("  FAIL  {0,-24} line {1,5}  {2}  - no native Case; a refusal here breaks into the VBE" -f $p.Name, $p.Start, $how)
        [void]$failures.Add($p.Name)
    }
}
if ($raising.Count -eq 0) { Write-Output '  (none)' }

Write-Output ''
Write-Output '--- baseline drift (reported, not failed) ---'
$newRaisers  = @($raising  | Where-Object { $baseline -notcontains $_ })
$goneRaisers = @($baseline | Where-Object { $raising  -notcontains $_ })
if ($newRaisers.Count -eq 0 -and $goneRaisers.Count -eq 0) {
    Write-Output '  none - the raising set matches the recorded baseline'
} else {
    foreach ($n in $newRaisers)  { Write-Output "  NEW      $n now reaches a raise - add it to `$baseline once reviewed" }
    foreach ($n in $goneRaisers) { Write-Output "  DROPPED  $n no longer reaches a raise - its Case may be removable" }
}

Write-Output ''
Write-Output '--- native Cases for helpers that cannot raise (harmless; no longer required) ---'
$idleCases = @($nativeCases | Where-Object { $n = $_; -not ($raising | Where-Object { $_.ToLower() -eq $n }) })
if ($idleCases.Count -eq 0) { Write-Output '  (none)' } else { foreach ($c in $idleCases) { Write-Output "  $c" } }

Write-Output ''
Write-Output '--- below the inject boundary: build/deploy surface, not interpreter-reachable ---'
if ($belowBoundary.Count -eq 0) {
    Write-Output '  (none)'
} else {
    foreach ($p in ($belowBoundary | Sort-Object Name)) {
        Write-Output ("  {0,-24} line {1,5}  absent from VlaHelperManifest by construction" -f $p.Name, $p.Start)
    }
}

Write-Output ''
if ($failures.Count -gt 0) {
    Write-Output ("=== CHECK: FAILED - {0} helper(s) can raise with no native Case in TryRuntimeHelper ===" -f $failures.Count)
    Write-Output '    Each one refuses through Application.Run, which does not propagate'
    Write-Output "    Err.Raise to the caller's handler: the user gets a VBE break dialog"
    Write-Output '    instead of a Frazaro modal. Add a native Case per helper.'
    exit 1
}

Write-Output ("=== CHECK: clean - all {0} raising governed helper(s) have a native Case ===" -f $raising.Count)
exit 0
