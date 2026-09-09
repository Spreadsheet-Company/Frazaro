<#
check_id_registry.ps1 - F.12's mechanical check.

Mechanizes SD-9 ("IDs are never reused across documents, and a retired ID is
never re-minted") the same way REBUILD.md's R11/Appendix D mechanizes it for
module names: a grep-shaped check, runnable in one call, no live Excel needed.

GOVERNED set: docs/BETA_ROADMAP.md + docs/ALPHA*_ROADMAP.md. These are "the
namespace" SD-9 and F.12 describe - the strategy file and its version ledgers,
the promotion path where SD-9's founding incident (Alpha 1's bare "F1" vs this
file's "F.1") actually happened. A new item ID is only minted here.

ADVISORY set: docs/REBUILD.md, docs/LESSONS.md, docs/AUDIT.md,
docs/PROJECT_BRIEF.md. These documents keep their own internal numbering for
their own purposes (REBUILD.md's R1-R11 lint rules and layer/plate steps,
etc.) and this script does not govern or rename them. But BETA_ROADMAP.md's own
prose cites REBUILD.md IDs inline (R7, R9, R10, S3.1, ...), so a token that is
free in the governed set can still collide in a reader's head with an advisory
one - this script reports that overlap so it can be a documented, understood
choice instead of a silent one, exactly as Appendix D did for module names.

A bulleted sub-line that CROSS-REFERENCES another item's id (this doc's own
"**TOKEN cross-reference:**" callout phrasing) is not a definition and is
excluded from Pass 1 accordingly - found live as a false positive (P-PROBE,
flagged "defined" 3 times when only one of the three was a real definition,
the other two callouts from LISTOPS-BUDGET/TABLESPEC) while scoping F.14's
own follow-ups. Excluded mentions print under "CROSS-REFERENCE MENTIONS" so
the exclusion stays visible rather than silent.

Usage:  pwsh -File tools/check_id_registry.ps1
Exit code: 0 always informational today (no governed collision has ever been
found); a real punctuation-variant collision in the governed set will still
print under "GOVERNED COLLISIONS" for a human to read at version-close.
#>

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$docsDir  = Join-Path $repoRoot 'docs'

$governedPaths = @(Get-ChildItem -Path $docsDir -Filter 'BETA_ROADMAP*.md') +
                  @(Get-ChildItem -Path $docsDir -Filter 'ALPHA*_ROADMAP.md') |
                  Sort-Object Name
$advisoryNames = @('REBUILD.md', 'LESSONS.md', 'AUDIT.md', 'PROJECT_BRIEF.md')
$advisoryPaths = $advisoryNames | ForEach-Object { Join-Path $docsDir $_ } |
                  Where-Object { Test-Path $_ }

# Three token shapes, tried in this order: SD-N; word-suffix (G-PIVOT,
# P-PROF, L-FILE-HELPERS); numeric, dot optional before the first digit
# group so F.12 and F1 are the same shape (IN.0.5, S3.1, G11r all included).
#
# The numeric arm's prefix bound was {1,4} until 2026-09-09, and that cap
# was a SILENT HOLE, found while minting PROLOG.10. Three whole families
# have prefixes longer than four letters - PROLOG (6), DATALOG (7) and
# SOLVE (5) - so none of PROLOG.1-.10, DATALOG.6 or SOLVE.1-.9 was ever
# governed: 24 ids over which SD-9's "never reused, never re-minted"
# guarantee simply did not apply, while the check reported clean. Nothing
# had gone wrong yet; the point is that nothing WOULD HAVE BEEN CAUGHT.
#
# Raised to {1,8} - DATALOG's seven plus one spare - rather than to
# [A-Z]+, deliberately: an unbounded prefix would start matching ordinary
# prose tokens of the same shape (SHA256, UTF8, ISO8601, RFC4180) if one
# ever appeared bolded at the head of a bullet, and an explicit reviewable
# bound is this file's own house style. Measured before landing: widening
# 4 -> 8 ADDS exactly those 24 ids and REMOVES nothing, with collisions
# and duplicate definitions both still zero.
#
# Only the NUMERIC arm moved. The word-suffix arm keeps {1,4} because
# every named form in use fits it (G-*, L-*, P-*, REPL-EVAL) and widening
# an arm with no demonstrated gap would trade a real hole for a
# speculative false-positive surface.
$tokenPattern = '(?<tok>SD-\d+|[A-Z]{1,4}-[A-Z][A-Z0-9-]*|[A-Z]{1,8}\.?\d+(?:\.\d+)*[a-z]?)'

function Get-Normalized([string]$tok) {
    ($tok.ToUpperInvariant() -replace '[.\-]', '')
}

# {1,8}, matching $tokenPattern's numeric arm above - all three bounds are
# one decision and must move together. Widening only the matcher left
# PROLOG/DATALOG/SOLVE governed for COLLISION purposes but still invisible
# to the high-water report: Get-LeadingNumber failed on "PROLOG.10" (it
# consumed at most PROL, then wanted a digit and found OG), so the token
# reported no series number and fell through to the named-forms bucket
# instead - which is exactly where DATALOG.6 was seen sitting, one edit
# short of the fix. Caught before landing; the lesson is that this cap was
# written down in three places, not one.
#
# [A-Z] never matches a digit, so a greedy {1,8} still stops correctly at
# the prefix boundary: PROLOG10 -> PROLOG, G11R -> G, DR2 -> DR, PORT3 ->
# PORT.
function Get-Prefix([string]$tok) {
    if ($tok -match '^([A-Z]{1,8})') { return $Matches[1] }
    return $tok
}

function Get-LeadingNumber([string]$tok) {
    if ($tok -match '^[A-Z]{1,8}\.?(\d+)') { return [int]$Matches[1] }
    return $null
}

# --- Pass 1: definition-site tokens in the governed set ------------------
# A "definition" is a bold token at the HEAD of a bullet line:
# "- <glyph?> **TOKEN". A token is at the head when nothing but bullet
# glyphs, other bold ids and punctuation precede it on the line - so a
# multi-item bullet ("- ⬜ **G-SORTFILTER**, ⬜ **G-TABS**, ⬜ **G-FORMULA**"
# defines all three, and "**U1 · U3 · U2 · U5**" is one bold span either
# way). A bold token that follows any prose - any lowercase letter once
# the ids themselves are stripped from the prefix - is a cross-reference
# to another item, reported under CROSS-REFERENCE MENTIONS rather than
# counted as a definition.
#
# Why (2026-09-07, 0.5.1's pre-flight): BETA_ROADMAP2.md's one-paragraph-
# per-item style bolds ids mid-sentence when an item points at another -
# "absorbed into **PF.4**", "🔒 **SD-10 governs:**", "**Not** the finding
# **P-PROBE** already closed", "**LE.1 is the read-only view of the
# identical table.**" - and every one of those read as a second definition
# of an id defined once, higher up the file: four false DUPLICATE
# DEFINITIONS, the same false-collision shape the earlier exclusion below
# was added for, now in a phrasing that exclusion cannot see. A first
# draft of this rule took only the line's FIRST bold span as the head;
# a before/after diff of this script's own output caught that it dropped
# G-TABS/G-FORMULA/G-TEXT from the multi-item line above, hence "no prose
# before it" rather than "first on the line".
#
# EXCLUSION, added after an earlier real false positive (P-PROBE, found
# while scoping F.14/LX.2's own follow-ups) and kept: a nested sub-bullet
# that POINTS AT another item's id rather than restating its own definition
# uses this doc's own recurring callout phrasing, "**TOKEN cross-reference:**"
# (BETA_ROADMAP.md's LISTOPS-BUDGET and TABLESPEC entries both cross-
# reference P-PROBE this way). That bolded token sits at the head of its
# own sub-bullet, so the head-position rule alone would still count it -
# excluded by the literal word immediately following the token, not by
# file/line - the phrasing is the signal, the same way "cross-reference"
# reads to a person skimming.
$defined   = New-Object System.Collections.Generic.List[object]
$crossRefs = New-Object System.Collections.Generic.List[object]
foreach ($f in $governedPaths) {
    $lineNo = 0
    foreach ($line in Get-Content -LiteralPath $f.FullName) {
        $lineNo++
        if ($line -notmatch '^\s*-\s') { continue }
        $ms = [regex]::Matches($line, "\*\*$tokenPattern")
        foreach ($m in $ms) {
            $tailStart = $m.Index + $m.Length
            $tail = $line.Substring($tailStart, [Math]::Min(20, $line.Length - $tailStart))
            # Head position: strip every bold id from the prefix, then any
            # lowercase letter left is prose, and this token follows it.
            $prefix = $line.Substring(0, $m.Index) -replace "\*\*$tokenPattern", ''
            $atHead = ($prefix -cnotmatch '[a-z]')
            if (-not $atHead -or $tail -match '^\s+cross-reference\b') {
                $crossRefs.Add([pscustomobject]@{
                    File = $f.Name
                    Line = $lineNo
                    Raw  = $m.Groups['tok'].Value
                })
                continue
            }
            $defined.Add([pscustomobject]@{
                File = $f.Name
                Line = $lineNo
                Raw  = $m.Groups['tok'].Value
                Norm = Get-Normalized $m.Groups['tok'].Value
            })
        }
    }
}

# --- Report: per-prefix high-water mark (numeric shapes only) ------------
Write-Output '=== HIGH-WATER MARK PER PREFIX (governed set) ==='
$numeric = $defined | Where-Object { (Get-LeadingNumber $_.Raw) -ne $null }
$numeric | Group-Object { Get-Prefix $_.Raw } | Sort-Object Name | ForEach-Object {
    $maxTok = $_.Group | Sort-Object { Get-LeadingNumber $_.Raw } | Select-Object -Last 1
    $maxN = Get-LeadingNumber $maxTok.Raw
    # "next free" mirrors the separator style of today's highest token (F.12
    # -> F.13, but DR2 -> DR3) - a hardcoded dot would misdescribe a bare family.
    $sep = if ($maxTok.Raw -match '\.') { '.' } else { '' }
    "{0,-4} highest = {1,-10} next free = {0}{2}{3}" -f $_.Name, $maxTok.Raw, $sep, ($maxN + 1)
}

$wordForm = $defined | Where-Object { (Get-LeadingNumber $_.Raw) -eq $null }
Write-Output ''
Write-Output '=== NAMED (word-suffix) FORMS IN USE (governed set) ==='
$wordForm | Sort-Object Raw -Unique | ForEach-Object { "  $($_.Raw)" }

# --- Report: governed collisions ------------------------------------------
# Same normalized token, more than one distinct raw spelling in the governed
# set - this is SD-9's exact failure shape (bare "F1" vs dotted "F.1").
Write-Output ''
Write-Output '=== GOVERNED COLLISIONS (punctuation-variant: SD-9 shape) ==='
$collisions = $defined | Group-Object Norm | Where-Object {
    ($_.Group.Raw | Sort-Object -Unique).Count -gt 1
}
if ($collisions.Count -eq 0) {
    Write-Output '  none found'
} else {
    foreach ($c in $collisions) {
        Write-Output "  $($c.Name):"
        $c.Group | ForEach-Object { "    $($_.Raw)  ($($_.File):$($_.Line))" }
    }
}

# --- Report: same raw token defined twice in one file --------------------
Write-Output ''
Write-Output '=== DUPLICATE DEFINITIONS (same raw token, same file, 2+ lines) ==='
$dups = $defined | Group-Object File, Raw | Where-Object { $_.Count -gt 1 }
if ($dups.Count -eq 0) {
    Write-Output '  none found'
} else {
    foreach ($d in $dups) {
        $d.Group | ForEach-Object { "  $($_.Raw) in $($_.File):$($_.Line)" }
    }
}

# --- Report: cross-reference mentions excluded from Pass 1 (informational) -
Write-Output ''
Write-Output '=== CROSS-REFERENCE MENTIONS (excluded from definitions above) ==='
if ($crossRefs.Count -eq 0) {
    Write-Output '  none found'
} else {
    $crossRefs | ForEach-Object { "  $($_.Raw) in $($_.File):$($_.Line)" }
}

# --- Report: mirrored across governed files (expected; informational) ----
Write-Output ''
Write-Output '=== MIRRORED ACROSS GOVERNED FILES (informational, not a collision) ==='
$mirrored = $defined | Group-Object Norm | Where-Object {
    ($_.Group.File | Sort-Object -Unique).Count -gt 1
}
foreach ($m in $mirrored) {
    $files = ($m.Group | ForEach-Object { "$($_.File):$($_.Line)" }) -join ', '
    "  $($m.Group[0].Raw)  -> $files"
}

# --- Report: advisory overlap with non-governed docs ----------------------
Write-Output ''
Write-Output '=== ADVISORY OVERLAP (governed token also appears in a non-governed doc) ==='
Write-Output '    (informational only - this script does not rename or govern these docs)'
$governedNorms = $defined.Norm | Sort-Object -Unique
foreach ($ap in $advisoryPaths) {
    $text = Get-Content -LiteralPath $ap -Raw
    $mentioned = [regex]::Matches($text, "\b$tokenPattern\b") |
                 ForEach-Object { Get-Normalized $_.Groups['tok'].Value } |
                 Sort-Object -Unique
    $hit = $governedNorms | Where-Object { $mentioned -contains $_ }
    if ($hit) {
        Write-Output "  $(Split-Path -Leaf $ap):"
        foreach ($h in $hit) {
            $rawsHere = ($defined | Where-Object Norm -eq $h | Select-Object -ExpandProperty Raw -Unique) -join '/'
            "    $rawsHere"
        }
    }
}

# Hand-maintained: retirement is a human judgment call (SD-9), so this list
# is not derived - but once entered, the exact retired SPELLING reappearing
# as a governed definition is exactly the failure SD-9 exists to prevent, so
# it IS checked. Matched on Raw, not Norm: F.1 is F1's legitimate successor,
# not a re-mint of it, and normalized matching would wrongly flag every use
# of F.1 forever. Only the retired spelling itself ("F1" bare) is forbidden.
$retired = @(
    [pscustomobject]@{ Raw = 'F1'; Note = "F1 (bare) - Alpha 1 interpreter mode; collided with F.1; SD-9's founding incident." }
)

Write-Output ''
Write-Output '=== RETIRED - NEVER RE-MINT ==='
$retiredHits = 0
foreach ($r in $retired) {
    Write-Output "  $($r.Note)"
    $reminted = $defined | Where-Object Raw -ceq $r.Raw
    foreach ($hit in $reminted) {
        Write-Output "    RE-MINTED: $($hit.Raw) in $($hit.File):$($hit.Line)"
        $retiredHits++
    }
}

# --- Summary / exit code ---------------------------------------------------
$issues = $collisions.Count + $dups.Count + $retiredHits
Write-Output ''
if ($issues -eq 0) {
    Write-Output "=== CHECK: clean - 0 governed collisions, 0 duplicate definitions, 0 retired IDs re-minted ==="
} else {
    Write-Output "=== CHECK: $issues issue(s) found - see GOVERNED COLLISIONS / DUPLICATE DEFINITIONS / RETIRED above ==="
}
exit ([Math]::Min($issues, 1))
