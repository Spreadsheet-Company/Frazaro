<#
check_raise_ratchet.ps1 - F.14's mechanical check.

SD-2 ("no refusal ships as a raw string - every refusal goes through `Raise`
with a stable ID and named parameters") is enforced by `VLA_Messages.RaiseMsg`
(LX.2, the message catalogue - one `AddMsg` per id, `{slot}` templates) and,
above `VLA_Runtime.bas`'s inject boundary, by its self-contained twin
`RaiseRuntimeMsg` (LX2.1). This script is the ratchet that keeps every
shipped module at or under its held count of raw `Err.Raise` sites, so a
new refusal cannot land as an inline string without a deliberate, reviewed
ceiling bump. When this file was first written no wrapper existed at all
and the ceilings were the whole enforcement; today the ceilings are small
and mostly name re-raises (a `fail:` handler propagating an error it did
not originate) that have no id to carry.

SHIPPED SET: read directly from `VLA_Build.bas`'s own `mods` array (the same
move `check_backend_parity.ps1` uses for `VLA_HeadTable.bas`'s `AddRow`
rows), so this script's notion of "shipped" can never silently drift from
what the build actually exports. Every one of the shipped components is
scanned, not just the ones that carry a raw site today - a first raw site
landing in a currently-clean module is exactly the case most worth catching.

WHAT COUNTS AS RAW: any source line containing the text "Err.Raise" whose
first non-blank character is not `'` - i.e. real code, not a comment merely
mentioning the term. This is not hypothetical: `VLA_English.bas`,
`VLA_Interpreter.bas`, and `VLA_Runtime.bas` each have real comment lines
mentioning "Err.Raise" today, and a naive full-text scan would both
over-count them now and false-fail later if someone only edited a comment.

CEILINGS: hand-maintained below, one entry per module that currently has a
nonzero count. A module absent from the table has an implicit ceiling of 0.
Bumping a ceiling is a deliberate, reviewable one-line edit - the same shape
as check_id_registry.ps1's own hand-maintained $retired list - not friction
to be designed around; it is the point. A module whose live count has
dropped BELOW its held ceiling is reported, not failed, as room to tighten.

NOT wired into VlaSelfTest, same reasoning F.12 gave for
check_id_registry.ps1 and AS.8 gave for check_backend_parity.ps1: this is a
version-close step a human runs, not a per-run gate - VlaSelfTest does no
real multi-file disk I/O (VlaLintCheck/VlaGoldens/VlaWriteGoldens's own
shared norm) and a 13-file source scan is squarely that category.

Usage:  pwsh -File tools/check_raise_ratchet.ps1
Exit code: 0 if every shipped module is at or under its ceiling; 1 if any
module's raw Err.Raise count has risen above its held ceiling.
#>

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$srcDir   = Join-Path $repoRoot 'src'

# --- Shipped set: VLA_Build.bas's own mods array, not a hand-copied list ---
$buildFile = Join-Path $srcDir 'VLA_Build.bas'
$buildText = Get-Content -LiteralPath $buildFile -Raw
$modsMatch = [regex]::Match($buildText, 'mods\s*=\s*Array\(([^)]*)\)')
if (-not $modsMatch.Success) {
    Write-Error "No 'mods = Array(...)' line found in $buildFile - VLA_Build.bas's shipped-module list shape may have changed; update this script's `$modsMatch pattern."
}
$modNames = [regex]::Matches($modsMatch.Groups[1].Value, '"([^"]*)"') |
            ForEach-Object { $_.Groups[1].Value }
if ($modNames.Count -eq 0) {
    Write-Error "Parsed zero module names out of the mods array in $buildFile - regex mismatch, not an empty ship list."
}

function Get-ModuleFilePath([string]$name) {
    foreach ($ext in '.bas', '.cls', '.frm') {
        $p = Join-Path $srcDir "$name$ext"
        if (Test-Path -LiteralPath $p) { return $p }
    }
    Write-Error "Shipped module '$name' (from VLA_Build.bas's mods array) has no .bas/.cls/.frm file in $srcDir - export it, or the mods array is stale."
}

function Get-RawRaiseCount([string]$path) {
    $lines = Get-Content -LiteralPath $path
    $n = 0
    foreach ($line in $lines) {
        if ($line -notmatch 'Err\.Raise') { continue }
        if ($line.TrimStart() -match "^'") { continue }   # comment-only line
        $n++
    }
    return $n
}

# --- Held ceilings: hand-maintained, bumped deliberately, one line per bump ---
# Seeded 2026-08-27 from a real recount (F.14's own scoping pass), with
# comment-only mentions of "Err.Raise" excluded - see this script's own
# header for why that exclusion is necessary, not cosmetic.
# Lowered 2026-08-27 (LX.2, migration mostly complete): VLA_Loader
# 3->0, VLA_Lint 2->0, VLA_IDE 16->1 (its one remaining raw site,
# ~line 1972, is ReadWordFile's cleanup re-raise of an already-caught
# error - not an origination of new English text, stays raw by
# design), VLA_Interpreter 61->1 (same re-raise category, ~line 1199,
# ExecStmtTrapped), VLA_English 97->1 (same category again, ~line 7606,
# EnglishAuditText's cleanup), VLA 141->7 (the emitter's own more
# elaborate version of the same category - VlaTranspile's emitfail
# handler re-raises with an added location suffix across 4 branches
# rather than one bare line, plus two VlaReadForms-family fail:
# handlers and VlaExpandStepText's restoreBudget). VLA_Messages added
# at 4 - three are infrastructure RaiseMsg needs and cannot route
# through itself (duplicate-id in AddMsg, unknown-id and missing-slot
# in RaiseMsg/SlotValue); the fourth is RaiseMsg's own chokepoint call,
# the literal place a migrated refusal's Err.Number/Source/Description
# actually gets raised once the catalogue lookup and slot substitution
# succeed - this script counts it because it is textually still
# "Err.Raise", not because it is a raw refusal bypassing the id system;
# it is the opposite, the mechanism BY WHICH an id'd refusal is raised.
#
# VLA_Runtime raised BACK to 22, same session: 23->0 was live-caught
# crashing a fresh-workbook "Compile and Trace" ("Variable not
# defined", VLA_Messages unresolved) and reverted. This module has an
# EN_RUNTIME INJECT BOUNDARY (VLA_Runtime.bas's own header, and the
# marker itself ~line 1126): code above it ships verbatim into a
# user's workbook as Frazaro_EN_Runtime, with no other add-in module
# present - the same constraint that already forced a hand-duplicated
# VLA_Identity.Fold copy in this module (~line 406) before LX.2 ever
# started. 22 of 23 raw sites sit above the boundary and cannot call
# VLA_Messages; only 1 (runtime-helpers-unreadable, inside
# VlaInjectRuntime itself, add-in-side only) legitimately stays
# migrated. This ceiling will very likely never drop further without a
# structural change (a self-contained mini-catalogue duplicated above
# the boundary, mirroring Fold's own precedent) - not attempted here.
#
# LX2.1, later session: that structural change was built - a second,
# self-contained catalogue (RaiseRuntimeMsg/RuntimeCatalogue/
# RuntimeAddEntries/RuntimeSubstituteSlots, VLA_Runtime.bas, right
# after Fold) duplicated above the boundary, same shape as
# VLA_Messages.RaiseMsg but its own table and its own chokepoint, since
# nothing above the boundary may call VLA_Messages. All 22 sites now
# route through it. VLA_Runtime 22->4, the same 4-raw-site shape
# VLA_Messages already set the precedent for: RuntimeAddMsg's
# duplicate-id check, RaiseRuntimeMsg's unknown-id fallback and its own
# chokepoint call, and RuntimeSlotValue's missing-value fallback - none
# of these can route through the mechanism that just failed to build
# itself, same reasoning as VLA_Messages's own 4.
# LX5.2, later session: VLA_English.bas split into VLA_English.bas (twelve
# pure vocabulary functions, zero raw sites) and VLA_SentenceEngine.bas (the
# engine, including EnglishAuditText's own cleanup re-raise - the module's
# one raw site, unchanged in kind, just relocated). VLA_English's own entry
# below is gone rather than zeroed, matching this script's own "absent =
# implicit 0" convention.
# 0.5.1 pre-flight (2026-09-07): VLA 7->8. The eighth site is
# VlaProbeMacroForm's own fail: handler (P-PROBE round 2, 2026-09-01) - the
# third VlaReadForms-family cleanup re-raise (VlaReadForms,
# VlaReadFormsWithLines, VlaProbeMacroForm), same category as the other
# seven: a caught error propagated after VlaPopContext, no new English
# text, nothing an id could name. It landed after LX.2 set the 7 and this
# script was never re-run at a version close until tools/release.ps1
# started running it - 0.5.0 shipped past a red ratchet. The same
# pre-flight routed the four genuinely raw refusals that had accumulated
# in VLA_Relation (1) and VLA_Sql (3) through RaiseMsg instead of
# bumping their ceilings, so both stay at the implicit 0.
$ceilings = @{
    'VLA'                = 8
    'VLA_SentenceEngine' = 1
    'VLA_Interpreter'    = 1
    'VLA_Runtime'        = 4
    'VLA_IDE'            = 1
    'VLA_Messages'       = 4
}

$failed      = New-Object System.Collections.Generic.List[string]
$tightenable = New-Object System.Collections.Generic.List[string]

Write-Output '=== RAW Err.Raise RATCHET (SD-2 enforcement; F.14) ==='
Write-Output "Shipped modules scanned: $($modNames.Count) (from VLA_Build.bas's mods array)"
Write-Output ''

foreach ($name in $modNames) {
    $path    = Get-ModuleFilePath $name
    $count   = Get-RawRaiseCount $path
    $ceiling = 0
    if ($ceilings.ContainsKey($name)) { $ceiling = $ceilings[$name] }

    $line = "  {0,-18} raw={1,4}  ceiling={2,4}" -f $name, $count, $ceiling
    if ($count -gt $ceiling) {
        Write-Output "$line  FAIL (+$($count - $ceiling))"
        $failed.Add($name)
    } elseif ($count -lt $ceiling) {
        Write-Output "$line  ok (room to tighten: -$($ceiling - $count))"
        $tightenable.Add($name)
    } else {
        Write-Output "$line  ok"
    }
}

if ($tightenable.Count -gt 0) {
    Write-Output ''
    Write-Output "=== CEILINGS ABOVE ACTUAL (tighten in `$ceilings above when next touching this file) ==="
    $tightenable | ForEach-Object { Write-Output "  $_" }
}

Write-Output ''
if ($failed.Count -eq 0) {
    Write-Output '=== CHECK: clean - no shipped module raised its raw Err.Raise count above its held ceiling ==='
} else {
    Write-Output "=== CHECK: $($failed.Count) module(s) exceeded their ceiling - $($failed -join ', ') ==="
    Write-Output "A raw Err.Raise site was added. A refusal belongs in VLA_Messages.bas's catalogue (AddMsg an id, then VLA_Messages.RaiseMsg at the site - or RaiseRuntimeMsg above VLA_Runtime.bas's inject boundary). Only a bare re-raise of an already-caught error has no id to carry - if that is what this is, bump that module's number in `$ceilings above and say why in the comment and the commit message. If it wasn't intentional, that's the drift this ratchet exists to catch."
}
exit ([Math]::Min($failed.Count, 1))
