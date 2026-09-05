<#
check_translate_purity.ps1 - PORT.2's mechanical check.

PORT.1 established, and confirmed by direct inspection rather than assumed,
that the English -> VLA -> VBA translate path touches no Excel host object at
all except through one seam (VlaTranspile's own prelude load, VLA.bas's
PreludeMacros - fixed at its source by VlaSetPreludeOverride/
VlaClearPreludeOverride, PORT.1's own change). That property is the entire
premise behind ever porting the engine to a non-VBA host (a browser page, a
Tauri/Electron shell - see docs/SUBSTRATE.md's watch-list and the VENTURE.md/
README.md discussion of a future web runtime): if the translate path stays
host-free, a port is a text-machinery port; if it silently grows a host
touch, a port inherits the host along with it, unnoticed until someone tries.

This script is that ratchet, in check_raise_ratchet.ps1's own exact shape:
scan a named, hand-maintained set of (module, function) pairs that together
make up "the translate path," count forbidden host-token hits per pair, and
fail if a pair's count rises above its held ceiling (implicitly 0 for any
pair not listed below).

SCANNED SET: hand-maintained below, not derived from a manifest (there is no
manifest of "which functions are the translate path" to derive from - this
list IS that manifest, in the same spirit VLA_Build.bas's mods array is the
manifest for "what ships"). Extraction relies on one fact that is guaranteed
by the VBA language itself, not merely usually true: Function/Sub definitions
never nest, so the first "End Function"/"End Sub" line after a scanned
function's own header line is unambiguously its own end - no brace-matching
or continuation-tracking needed, and no false-positive nesting is possible.

WHAT COUNTS AS A HIT: any non-comment line inside a scanned function's body
that matches one of the forbidden host-token patterns below. Comment-only
lines (first non-blank character `'`) are excluded, the same exclusion
check_raise_ratchet.ps1 already uses and for the identical reason - several
of these functions carry comments that MENTION ThisWorkbook/Application/etc.
without touching them.

CEILINGS: one exception is real and documented, not a bug to chase: VLA_
English.bas's EnsureInit registers two prelude phrase rules whose TEMPLATE
TEXT is the literal string "(msgbox {e})" - a VLA s-expression head name
that VlaTranspile later emits as a real VBA MsgBox call, not a live MsgBox
call inside EnsureInit itself. The regex below cannot distinguish "MsgBox
inside a quoted template string" from "MsgBox as a live statement" without a
real VBA string-literal parser this project deliberately doesn't carry for a
text-scan tool (the same trade-off check_raise_ratchet.ps1 already made for
comment-only lines) - so the two known hits are held at an explicit ceiling
instead, confirmed by hand on 2026-08-31 and re-confirmable by anyone who
reads the two AddPhraseRule lines this ceiling covers.

NOT wired into VlaSelfTest, same reasoning F.12/F.14/AS.8 give for their own
scripts: this is a version-close step a human runs, not a per-run gate.

Usage:  pwsh -File tools/check_translate_purity.ps1
Exit code: 0 if every scanned function is at or under its held ceiling; 1 if
any function's forbidden-token count has risen above it.
#>

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$srcDir   = Join-Path $repoRoot 'src'

# --- The translate path: hand-maintained (module, function) pairs ---
# Together, these compose the entire English -> VLA -> VBA chain a
# host-free caller (VLA_Browser.bas, and any future port) walks. A new
# function added to this chain belongs on this list the same day.
$scanTargets = @(
    # LX5.2: these five moved from VLA_English.bas to VLA_SentenceEngine.bas
    # (BETA_ROADMAP2.md's LX.5, phase 2 - the physical module split). Purity
    # is a property of the CODE, not of which file holds it, so the move
    # itself needed no re-audit - only this list's own Module field, or the
    # scanner would silently stop finding these functions at all.
    @{ Module = 'VLA_SentenceEngine'; Fn = 'EnglishToVla' }
    @{ Module = 'VLA_SentenceEngine'; Fn = 'EnglishToVba' }
    @{ Module = 'VLA_SentenceEngine'; Fn = 'EnglishResetGrammar' }
    @{ Module = 'VLA_SentenceEngine'; Fn = 'EnglishLoadVocabularyText' }
    @{ Module = 'VLA_SentenceEngine'; Fn = 'EnsureInit' }
    @{ Module = 'VLA';         Fn = 'VlaTranspile' }
    @{ Module = 'VLA_Browser';     Fn = 'EnglishTranslateTextToVla' }
    @{ Module = 'VLA_Browser';     Fn = 'EnglishTranslateTextToVba' }
)

# --- Forbidden host-token patterns (regex, case-insensitive) ---
$forbidden = @(
    'ActiveSheet', 'ActiveWorkbook', 'ThisWorkbook',
    'Application\.', 'Worksheets\(', 'Workbooks\(',
    '\.Cells\(', '\.Range\(', 'MsgBox',
    'Dir\$?\(', '^\s*Open\s', '^\s*Kill\s'
)
$forbiddenRegex = ($forbidden -join '|')

# --- Held ceilings: "Module::Function" -> allowed hit count ---
# A pair absent here has an implicit ceiling of 0.
$ceilings = @{
    'VLA_SentenceEngine::EnsureInit' = 2   # (msgbox {e}) x2, template text - see header
}

function Get-FunctionBody([string]$path, [string]$fnName) {
    $lines = Get-Content -LiteralPath $path
    $startPat = "^(Public|Private)\s+(Function|Sub)\s+$([regex]::Escape($fnName))\s*[( ]"
    $inBody = $false
    $body = New-Object System.Collections.Generic.List[string]
    foreach ($line in $lines) {
        if (-not $inBody) {
            if ($line -match $startPat) { $inBody = $true; $body.Add($line) }
            continue
        }
        $body.Add($line)
        if ($line -match '^End (Function|Sub)\b') { return $body }
    }
    if ($inBody) {
        Write-Error "Function '$fnName' in $path started but never found its own End Function/Sub - file may be truncated or the function name is ambiguous."
    }
    Write-Error "Function '$fnName' not found in $path - PORT.2's scan list may be stale (renamed or removed function), or the module's file name changed."
}

function Get-HitCount([System.Collections.Generic.List[string]]$body) {
    $n = 0
    foreach ($line in $body) {
        if ($line.TrimStart() -match "^'") { continue }   # comment-only line
        if ($line -match $forbiddenRegex) { $n++ }
    }
    return $n
}

$failed = New-Object System.Collections.Generic.List[string]

Write-Output '=== TRANSLATE-PATH PURITY RATCHET (PORT.1/PORT.2) ==='
Write-Output "Functions scanned: $($scanTargets.Count)"
Write-Output ''

foreach ($t in $scanTargets) {
    $path = Join-Path $srcDir "$($t.Module).bas"
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Error "Scanned module '$($t.Module)' has no .bas file at $path - PORT.2's scan list is stale."
    }
    $body = Get-FunctionBody -path $path -fnName $t.Fn
    $count = Get-HitCount -body $body
    $key = "$($t.Module)::$($t.Fn)"
    $ceiling = 0
    if ($ceilings.ContainsKey($key)) { $ceiling = $ceilings[$key] }

    $line = "  {0,-45} hits={1,3}  ceiling={2,3}" -f $key, $count, $ceiling
    if ($count -gt $ceiling) {
        Write-Output "$line  FAIL (+$($count - $ceiling))"
        $failed.Add($key)
    } elseif ($count -lt $ceiling) {
        Write-Output "$line  ok (room to tighten: -$($ceiling - $count))"
    } else {
        Write-Output "$line  ok"
    }
}

Write-Output ''
if ($failed.Count -eq 0) {
    Write-Output '=== CHECK: clean - the translate path touches no host object beyond its held, documented exceptions ==='
} else {
    Write-Output "=== CHECK: $($failed.Count) function(s) exceeded their ceiling - $($failed -join ', ') ==="
    Write-Output "A host-object token was added to a function on the translate path. If intentional and unavoidable, add or raise its ceiling in `$ceilings above and say why (PreludeMacros-style: guard it behind VlaSetPreludeOverride or an equivalent seam first, and prefer that over raising the ceiling). If it wasn't intentional, that is exactly the drift this ratchet exists to catch before a port silently inherits it."
}
exit ([Math]::Min($failed.Count, 1))
