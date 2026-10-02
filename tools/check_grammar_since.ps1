<#
check_grammar_since.ps1 - CO.6's ratchet, landed with F.10, its first consumer.

docs/GRAMMAR_SINCE.md answers one question per form: WHICH RELEASE DID THIS
FIRST WORK IN. F.10's (requires-version "X") reads exactly that answer - so
the gate is only as trustworthy as the ledger's coverage. A form that
reaches a release with no row is a form an author cannot write a correct
requires-version against, and the ledger is append-only (SD-9's discipline
at form scale), so a date guessed later, after the fact, is frozen wrong
forever. This script makes "undated" fail at the check instead of surfacing
as a wrong answer months later.

CO.6 deliberately did NOT ship this script with the seed. Under SD-7 work
is not scheduled without something that needs it, and until F.10 nothing
read these dates - a ratchet guarding an unread file is guarding nothing.
The seed was taken then because it was the perishable half; this is the
half that waited for its consumer.

WHERE THE INVENTORY COMES FROM - and, just as important, where it does NOT.
Neither inventory is re-derived here. Both come from the sibling script
that already owns that parsing:

  - phrase rules -> check_rule_coverage.ps1 -ListRules, which reads the
    GENERATED artifact scripts/polyglotta/english_expanded.vla, never
    english.vla. CO.6 found both halves of why: source UNDERCOUNTS (three
    rules are generator-emitted and appear in no source file at all), and
    the artifact is unsafe for HISTORICAL dating (the v0.5.0 export was
    stale, so three rules sayable in 0.5.0 would have been frozen at
    0.5.1). Inventory from the artifact; dates from source. This script
    only needs the inventory - it never assigns a date, it only asks
    whether one exists.
  - core dispatch arms -> check_emitter_coverage.ps1 -ListArms, that
    script's own Get-CaseArmGroups over its own six dispatch functions.

Both switches exist so this file does not carry a second copy of subtle
parsing. That is not tidiness: the rule-pattern reader silently broke for
two weeks because the artifact went vertical and the extraction still read
only the head line. One copy, in the script that owns it.

WHAT THIS DOES NOT DO. It does not check that a date is CORRECT - nothing
can, after the fact, which is the whole reason the ledger is append-only
and was seeded against source history rather than the artifact. It only
checks that every form currently in the grammar has a row. It also does not
police rows with no matching inventory entry: a RETIRED form keeps its row
and gains an "until:" (the ledger's own rule 3), so an extra row is legal.
Extras are reported, never failed on - retirement itself is CO.1's item.

Host-independent: reads files and shells to two sibling scripts. No Excel,
and never wired into VlaSelfTest - house ratchet shape.

Usage:  pwsh -File tools/check_grammar_since.ps1
#>

param(
    [string]$LedgerPath
)

$ErrorActionPreference = 'Stop'

# The reviewable baseline, and the only number in this file. It is a
# CEILING on undated forms, not an expected inventory size: adding a form
# together with its row - which is the whole discipline - must keep this
# script green without anyone editing it. It is 0 and there is no honest
# reason for it ever to rise; a raise is a decision to ship a form nobody
# can write a requires-version against.
$AllowedUndated = 0

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $LedgerPath) { $LedgerPath = Join-Path $repoRoot (Join-Path 'docs' 'GRAMMAR_SINCE.md') }
if (-not (Test-Path -LiteralPath $LedgerPath)) {
    Write-Error "No since-ledger at $LedgerPath."
}

# --- The ledger's own rows ----------------------------------------------
# Snapshot layout (ID_REGISTRY.md's shape, which CO.6 followed): a
# "### Phrasebook rules" heading, then a fenced block of
# "<version>  <pattern>" rows; then "### Core dispatch arms" and a fenced
# block of "<version>  <function>   <arm>" rows. A comma-grouped arm is ONE
# row with its spellings joined by " | ", exactly as -ListArms prints it -
# the ledger dates a code path, not a spelling.
$ledgerRules = New-Object System.Collections.Generic.HashSet[string]
$ledgerArms  = New-Object System.Collections.Generic.HashSet[string]
# LX.14: the third section, function words and phrases, and the rule rows a
# phrase now carries ("<pattern>  carried-by: <phrase> | <phrase>").
$ledgerPhrases = New-Object System.Collections.Generic.HashSet[string]
$carried = [ordered]@{}
$section = ''
$inFence = $false
foreach ($line in (Get-Content -LiteralPath $LedgerPath)) {
    if ($line -match '^###\s+Phrasebook rules\s*$') { $section = 'rules'; continue }
    if ($line -match '^###\s+Core dispatch arms\s*$') { $section = 'arms'; continue }
    if ($line -match '^###\s+Function words and phrases\s*$') { $section = 'phrases'; continue }
    if ($line -match '^```') { $inFence = -not $inFence; continue }
    if (-not $inFence -or $section -eq '') { continue }

    if ($section -eq 'rules') {
        $m = [regex]::Match($line, '^(\d+\.\d+\.\d+)\s\s+(.+?)\s*$')
        if ($m.Success) {
            $pattern = $m.Groups[2].Value
            $cb = $pattern.IndexOf('  carried-by: ')
            if ($cb -ge 0) {
                $carriers = @($pattern.Substring($cb + 14) -split '\s\|\s' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
                $pattern = $pattern.Substring(0, $cb).Trim()
                $carried[$pattern] = $carriers
            }
            [void]$ledgerRules.Add($pattern)
        }
    } elseif ($section -eq 'phrases') {
        $m = [regex]::Match($line, '^(\d+\.\d+\.\d+)\s\s+(.+?)\s*$')
        if ($m.Success) { [void]$ledgerPhrases.Add($m.Groups[2].Value) }
    } else {
        # The function name is a single token; two-or-more spaces separate
        # it from the arm, which may itself contain single spaces and " | ".
        $m = [regex]::Match($line, '^(\d+\.\d+\.\d+)\s\s+(\S+)\s\s+(.+?)\s*$')
        if ($m.Success) {
            [void]$ledgerArms.Add(($m.Groups[2].Value + "`t" + $m.Groups[3].Value))
        }
    }
}

# --- The live inventory, from the scripts that own each parser -----------
function Invoke-Sibling([string]$scriptName, [string[]]$scriptArgs) {
    $path = Join-Path $PSScriptRoot $scriptName
    if (-not (Test-Path -LiteralPath $path)) { Write-Error "Missing sibling script: $path" }
    $out = & powershell -NoProfile -File $path @scriptArgs 2>&1
    if ($LASTEXITCODE -ne 0) {
        $joined = ($out | ForEach-Object { [string]$_ }) -join [Environment]::NewLine
        Write-Error "$scriptName exited $LASTEXITCODE - fix that script first; this one cannot take an inventory from a failed run.`n$joined"
    }
    return @($out | ForEach-Object { [string]$_ })
}

$liveRules = @(Invoke-Sibling 'check_rule_coverage.ps1' @('-ListRules') | Where-Object { $_.Trim() -ne '' })
$liveArms  = @(Invoke-Sibling 'check_emitter_coverage.ps1' @('-ListArms') | Where-Object { $_.Trim() -ne '' })
# LX.14: the function words and phrases - the engine's, and english.vla's.
$livePhrases = @(@(Invoke-Sibling 'check_engine_call_names.ps1' @('-ListFunctionWords')) + @(Invoke-Sibling 'check_rule_coverage.ps1' @('-ListPhrases')) | Where-Object { $_.Trim() -ne '' })

# --- Compare -------------------------------------------------------------
$liveRuleSet = New-Object System.Collections.Generic.HashSet[string]
$liveRules | ForEach-Object { [void]$liveRuleSet.Add($_) }
$liveArmSet = New-Object System.Collections.Generic.HashSet[string]
$liveArms | ForEach-Object { [void]$liveArmSet.Add($_) }
$livePhraseSet = New-Object System.Collections.Generic.HashSet[string]
$livePhrases | ForEach-Object { [void]$livePhraseSet.Add($_) }

$undatedRules   = @($liveRules | Where-Object { -not $ledgerRules.Contains($_) })
$undatedArms    = @($liveArms  | Where-Object { -not $ledgerArms.Contains($_) })
$undatedPhrases = @($livePhrases | Where-Object { -not $ledgerPhrases.Contains($_) })
$extraRules     = @($ledgerRules | Where-Object { -not $liveRuleSet.Contains($_) -and -not $carried.Contains($_) })
$extraArms      = @($ledgerArms  | Where-Object { -not $liveArmSet.Contains($_) })
$extraPhrases   = @($ledgerPhrases | Where-Object { -not $livePhraseSet.Contains($_) })
# A carried row must name phrases that are live (rule 3, LX.14).
$badCarried = New-Object System.Collections.Generic.List[string]
foreach ($k in $carried.Keys) {
    foreach ($c in $carried[$k]) {
        if (-not $livePhraseSet.Contains($c)) { $badCarried.Add("$k  carried-by: $c") }
    }
}

Write-Output '=== GRAMMAR SINCE-LEDGER COVERAGE (every live form must carry a date) ==='
Write-Output "Ledger:   $LedgerPath"
Write-Output "Live:     $($liveRules.Count) phrase rules, $($liveArms.Count) core dispatch arms, $($livePhrases.Count) function words and phrases"
Write-Output "Rows:     $($ledgerRules.Count) rule rows ($($carried.Count) carried by a phrase), $($ledgerArms.Count) arm rows, $($ledgerPhrases.Count) function-word and phrase rows"
Write-Output ''

if ($undatedRules.Count -gt 0) {
    Write-Output "--- Phrase rules with NO row in the ledger, $($undatedRules.Count) ---"
    $undatedRules | ForEach-Object { Write-Output "  $_" }
    Write-Output ''
}
if ($undatedArms.Count -gt 0) {
    Write-Output "--- Core dispatch arms with NO row in the ledger, $($undatedArms.Count) ---"
    $undatedArms | ForEach-Object { Write-Output ("  " + ($_ -replace "`t", "  ")) }
    Write-Output ''
}
if ($undatedPhrases.Count -gt 0) {
    Write-Output "--- Function words and phrases with NO row in the ledger, $($undatedPhrases.Count) ---"
    $undatedPhrases | ForEach-Object { Write-Output "  $_" }
    Write-Output ''
}

# Reported, never failed on - see WHAT THIS DOES NOT DO above.
if ($extraRules.Count -gt 0 -or $extraArms.Count -gt 0 -or $extraPhrases.Count -gt 0) {
    Write-Output "--- Rows with no live form, $($extraRules.Count) rule(s) + $($extraArms.Count) arm(s) + $($extraPhrases.Count) phrase(s) (legal for a RETIRED form, which keeps its row and gains an 'until:' - the ledger's own rule 3; not a failure) ---"
    $extraRules   | ForEach-Object { Write-Output "  rule: $_" }
    $extraArms    | ForEach-Object { Write-Output ("  arm:  " + ($_ -replace "`t", "  ")) }
    $extraPhrases | ForEach-Object { Write-Output "  phrase: $_" }
    Write-Output ''
}
# LX.14: a rule a phrase now carries - not retired; its phrases must be live.
if ($carried.Count -gt 0) {
    Write-Output "--- Rule rows a phrase now carries, $($carried.Count) (their sentences work as before, through a general rule and the phrase - the ledger's rule 3) ---"
    foreach ($k in $carried.Keys) { Write-Output "  $k  <-  $($carried[$k] -join ' | ')" }
    Write-Output ''
}
if ($badCarried.Count -gt 0) {
    Write-Output "=== CHECK FAILED: $($badCarried.Count) carried-by: name(s) no live function word or phrase ==="
    $badCarried | ForEach-Object { Write-Output "  $_" }
    Write-Output ''
    Write-Output "A carried-by: says a phrase carries the rule's sentences now. Name a phrase that exists, or, if the sentences stopped working, give the row an until: instead (rule 3)."
    exit 1
}

$undated = $undatedRules.Count + $undatedArms.Count + $undatedPhrases.Count
if ($undated -gt $AllowedUndated) {
    Write-Output "=== CHECK FAILED: $undated form(s) are in the grammar with no row in the since-ledger (ceiling $AllowedUndated) ==="
    Write-Output ''
    Write-Output "Add a row for each above, dated with the release it first WORKS in - for a form being added now that is the release currently being prepared, not the last one tagged. The ledger is append-only: a date entered wrong is frozen, and it makes F.10's (requires-version ...) reject builds that would have run the phrasebook fine."
    exit 1
}

Write-Output "=== CHECK: clean - every live phrase rule, dispatch arm, function word and phrase carries a since-date ==="
exit 0
