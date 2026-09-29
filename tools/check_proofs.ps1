<#
check_proofs.ps1 - METAPROOF.1's ratchet: the proof corpus is well-formed,
checked without Excel.

WHY THIS EXISTS. scripts/proofs/*.vla hold proofs as forms - DATALOG
programs beside the answers they must give - and TestDatalogProofs
(VLA_Tests_Query.bas) runs them inside TestDSLs. That run needs Excel, and
the owner at the keyboard. Much of what can go wrong with a proof, though,
is its SHAPE, and shape is text: a paren left open, two proofs sharing a
name, a row narrower than its header, a (refuses ...) naming a message id
nothing raises. Those are caught here, host-independently, before a live
run is spent finding them. It is the first thing proofs-as-data buys that
proofs-in-VBA never could: a second reader.

WHAT IS CHECKED, in every scripts/proofs/*.vla:
  1. It reads: every paren closed, every string closed, by a tokenizer that
     follows VLA.bas's own Tokenize - a ; comment runs to the end of its
     line, "..." takes \" and \\ as escapes, and every other run of
     characters up to a paren, a space, a ; or a " is one atom. The reader
     lives in tools\proofs_lib.ps1, which the exporter below reads with.
  2. Every top-level form is (test-<engine> "name" clause ...), for an
     engine this corpus has a runner for (today: datalog).
  3. Every proof's name is a quoted string, unique in its file (compared
     without case, as the runner's own name check compares them).
  4. Every proof has exactly one (program ...) clause and exactly one
     expectation - (rows ...), (rows-in-any-order ...), (answer ...) or
     (refuses ...) - and no other clause.
  5. (rows ...) and (rows-in-any-order ...): a header that is a list of
     plain cells or the word headless, then rows that are lists of plain
     cells, every row (and a list header) as wide as the first.
  6. (answer x): x is true or false.
  7. (refuses id): id is a message id one of src/VLA_Messages.bas's AddMsg
     rows registers.
  8. Every proof file is in the baseline below and holds at least its
     floor of proofs, so a truncated or emptied file fails here and not
     only in a live run; a new proof file is added to the baseline, with
     its floor, in the commit that adds it.
  9. METAPROOF.2: the clingo export is current. tools\proofs_lp.ps1 writes
     every answer proof as a clingo program in tools\clingo, with three
     controls; here the same translation is made again, in memory, and
     every tools\clingo\proof-*.lp must match it - none missing, none
     changed, none left over from a proof that is gone. A proof edited
     without a fresh export fails here, so the question clingo was last
     asked is always the one the corpus asks now. Each proof file also
     has a floor of proofs its export must carry (the second baseline
     below), so a translation that starts declining proofs it used to take
     fails too.

WHAT IT DOES NOT DO: run a proof, or run clingo. Whether an answer is RIGHT
is TestDatalogProofs' question, and - for every proof the translation can
state - clingo's, run by the owner by hand (tools\proofs_lp.ps1 says how).
This checks that a proof is well-formed enough to be asked, and that
clingo is being asked the current question.

House shape, per the standing convention for this project's static scans:
PowerShell 5.1, host-independent, a hardcoded and reviewable baseline,
never wired into VlaSelfTest.

Usage:  powershell -File tools\check_proofs.ps1 [-ProofsDir <dir>] [-ClingoDir <dir>]
Exit 0 clean, exit 1 with every failure listed.
#>
param(
    # Folders to read instead of scripts\proofs and tools\clingo. Their only
    # purpose is to make THIS script testable: pointed at a scratch folder
    # holding a broken copy, it must go red.
    [string]$ProofsDir = '',
    [string]$MessagesPath = '',
    [string]$ClingoDir = ''
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
if ($ProofsDir -eq '') { $ProofsDir = Join-Path $root 'scripts\proofs' }
if ($MessagesPath -eq '') { $MessagesPath = Join-Path $root 'src\VLA_Messages.bas' }
if ($ClingoDir -eq '') { $ClingoDir = Join-Path $PSScriptRoot 'clingo' }

# The reader (rule 1) and the clingo translation (rule 9).
. (Join-Path $PSScriptRoot 'proofs_lib.ps1')

# --- the baseline: every proof file, and the fewest proofs it may hold ---
# 2026-09-27, METAPROOF.1: datalog.vla, sixteen proofs - TestDatalogBoundArgument's
# eleven and five of TestDatalog's refusals, moved out of VBA.
$floors = [ordered]@{
    'datalog.vla' = 16
}
# --- rule 9's baseline: every proof file, and the fewest of its proofs its
# clingo export may carry (0 for an engine clingo cannot check) ---
# 2026-09-28, METAPROOF.2: datalog.vla, eleven - every answer proof. Its five
# refusals are DATALOG's own policy, which no other lineage raises.
$clingoFloors = [ordered]@{
    'datalog.vla' = 11
}
# The engines a (test-<engine> ...) head may name: one per runner that
# exists. A proof for any other would be read by nothing.
$proofEngines = @('datalog')
$expectHeads = @('rows', 'rows-in-any-order', 'answer', 'refuses')

$failures = New-Object System.Collections.Generic.List[string]

# --- the message ids VLA_Messages registers (rule 7) ----------------------
$knownIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
$msgSource = [IO.File]::ReadAllText($MessagesPath)
foreach ($mt in [regex]::Matches($msgSource, '(?m)^\s*AddMsg m, "([^"]+)"')) {
    [void]$knownIds.Add($mt.Groups[1].Value)
}
if ($knownIds.Count -lt 100) {
    Write-Error "Read only $($knownIds.Count) message ids from $MessagesPath - has AddMsg's shape changed? This check cannot vouch for (refuses ...) ids it cannot read."
}

# Rule 5: every cell a plain atom or string, every row as wide as the first.
function Test-ProofRows($expect, [string]$where, $fails) {
    if ($expect.Items.Count -lt 2) {
        $fails.Add("${where}: (rows ...) needs a header first - a list of names, or headless")
        return
    }
    $rowsToCheck = New-Object System.Collections.Generic.List[object]
    $hdr = $expect.Items[1]
    if ($hdr.Kind -eq 'list') {
        $rowsToCheck.Add($hdr)
    } elseif (-not ($hdr.Kind -eq 'atom' -and $hdr.Text.ToLowerInvariant() -eq 'headless')) {
        $fails.Add("${where}: a header is a list of names, or the word headless")
        return
    }
    for ($ri = 2; $ri -lt $expect.Items.Count; $ri++) {
        $rowNode = $expect.Items[$ri]
        if ($rowNode.Kind -ne 'list') {
            $fails.Add("${where}: every row is a list of cells (line $($rowNode.Line))")
            return
        }
        $rowsToCheck.Add($rowNode)
    }
    $wide = -1
    foreach ($rowNode in $rowsToCheck) {
        if ($wide -lt 0) { $wide = $rowNode.Items.Count }
        if ($rowNode.Items.Count -ne $wide) {
            $fails.Add("${where}: a row $($rowNode.Items.Count) cells wide beside one $wide wide (line $($rowNode.Line))")
            return
        }
        foreach ($cellNode in $rowNode.Items) {
            if ($cellNode.Kind -eq 'list') {
                $fails.Add("${where}: a list where a cell belongs (line $($cellNode.Line))")
                return
            }
        }
    }
    if ($wide -eq 0) { $fails.Add("${where}: a row with no cells - DATALOG has no zero-width answer") }
}

Write-Output '=== PROOF CORPUS (scripts/proofs, METAPROOF.1) ==='
Write-Output ("  message ids read from VLA_Messages.bas: {0}" -f $knownIds.Count)

# --- rule 8: the baseline, both ways --------------------------------------
$present = @()
if (Test-Path -LiteralPath $ProofsDir) {
    $present = @(Get-ChildItem -LiteralPath $ProofsDir -Filter '*.vla' -File | ForEach-Object { $_.Name } | Sort-Object)
}
foreach ($pf in $present) {
    if (-not $floors.Contains($pf)) {
        $failures.Add("scripts/proofs/$pf is not in this script's baseline - add it to `$floors, with the fewest proofs it may hold, in the commit that adds it")
    }
}
foreach ($pf in $floors.Keys) {
    if ($present -notcontains $pf) {
        $failures.Add("scripts/proofs/$pf is in the baseline but missing - the proofs it held are gone, and nothing else would say so")
    }
}

# The files that passed rules 1-8, for rule 9: the translation takes a
# proof's shape as given.
$clingoReady = New-Object System.Collections.Generic.List[object]

foreach ($pf in $present) {
    $label = "scripts/proofs/$pf"
    $src = [IO.File]::ReadAllText((Join-Path $ProofsDir $pf))
    try {
        $toks = Get-ProofTokens $src $label
        $forms = ConvertTo-ProofForms $toks $label
    } catch {
        $failures.Add($_.Exception.Message)
        Write-Output ("  FAIL  {0}: does not read" -f $label)
        continue
    }

    $names = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $before = $failures.Count
    $proofCount = 0
    foreach ($form in $forms) {
        $proofCount++
        $where = "$label line $($form.Line)"
        # rule 2
        $head = Get-ProofHeadWord $form
        if ($head -notmatch '^test-(.+)$' -or $proofEngines -cnotcontains $Matches[1]) {
            $failures.Add("${where}: a top-level form must be (test-<engine> ...) for one of: $($proofEngines -join ', ')")
            continue
        }
        # rule 3
        if ($form.Items.Count -lt 2 -or $form.Items[1].Kind -ne 'str') {
            $failures.Add("${where}: a proof's name is a ""quoted string"", right after its head")
            continue
        }
        $proofTitle = $form.Items[1].Text
        if (-not $names.Add($proofTitle)) {
            $failures.Add("${where}: a second proof named ""$proofTitle""")
        }
        # rule 4
        $programs = 0
        $expectNode = $null
        $expectCount = 0
        for ($ci = 2; $ci -lt $form.Items.Count; $ci++) {
            $clauseNode = $form.Items[$ci]
            $clauseWord = Get-ProofHeadWord $clauseNode
            if ($clauseWord -eq 'program') {
                $programs++
            } elseif ($expectHeads -ccontains $clauseWord) {
                $expectCount++
                $expectNode = $clauseNode
            } else {
                $failures.Add("${where}: (${clauseWord} ...) is not a clause a proof has (line $($clauseNode.Line))")
            }
        }
        if ($programs -ne 1) { $failures.Add("${where}: a proof has exactly one (program ...) clause, this one has $programs") }
        if ($expectCount -ne 1) {
            $failures.Add("${where}: a proof states exactly one expectation, this one states $expectCount")
            continue
        }
        # rules 5-7
        $expectWord = Get-ProofHeadWord $expectNode
        switch -CaseSensitive ($expectWord) {
            'rows' { Test-ProofRows $expectNode $where $failures }
            'rows-in-any-order' { Test-ProofRows $expectNode $where $failures }
            'answer' {
                if ($expectNode.Items.Count -ne 2 -or $expectNode.Items[1].Kind -ne 'atom' -or
                    @('true', 'false') -cnotcontains $expectNode.Items[1].Text.ToLowerInvariant()) {
                    $failures.Add("${where}: (answer ...) is (answer true) or (answer false)")
                }
            }
            'refuses' {
                if ($expectNode.Items.Count -ne 2 -or $expectNode.Items[1].Kind -ne 'atom') {
                    $failures.Add("${where}: (refuses ...) names exactly one message id")
                } elseif (-not $knownIds.Contains($expectNode.Items[1].Text)) {
                    $failures.Add("${where}: (refuses $($expectNode.Items[1].Text)) names no message id VLA_Messages.bas registers")
                }
            }
        }
    }
    # rule 8
    if ($floors.Contains($pf) -and $proofCount -lt $floors[$pf]) {
        $failures.Add("${label}: $proofCount proof(s), below its floor of $($floors[$pf]) - were proofs lost?")
    }
    $bad = $failures.Count - $before
    if ($bad -eq 0) {
        Write-Output ("  ok    {0}: {1} proof(s), floor {2}" -f $label, $proofCount, $floors[$pf])
        $clingoReady.Add([pscustomobject]@{ File = $pf; Label = $label; Forms = $forms })
    } else {
        Write-Output ("  FAIL  {0}: {1} proof(s), {2} problem(s)" -f $label, $proofCount, $bad)
    }
}

# --- rule 9: the clingo export is current (METAPROOF.2) -------------------
foreach ($pf in $present) {
    if (-not $clingoFloors.Contains($pf)) {
        $failures.Add("scripts/proofs/$pf is not in this script's clingo baseline - add it to `$clingoFloors, with the fewest proofs its export may carry, in the commit that adds it")
    }
}
$wantFiles = [ordered]@{}
$exportsMade = 0
foreach ($cr in $clingoReady) {
    try {
        $ex = ConvertTo-ClingoExport $cr.Forms ([IO.Path]::GetFileNameWithoutExtension($cr.File)) $cr.Label
    } catch {
        $failures.Add("$($cr.Label): its clingo export cannot be made - $($_.Exception.Message)")
        continue
    }
    $exportsMade++
    foreach ($fl in $ex.Files) { $wantFiles[$fl.Name] = $fl.Text }
    $clingoFloor = 0
    if ($clingoFloors.Contains($cr.File)) { $clingoFloor = $clingoFloors[$cr.File] }
    $mark = 'ok  '
    if ($ex.Exported -lt $clingoFloor) {
        $failures.Add("$($cr.Label): its clingo export carries $($ex.Exported) proof(s), below its floor of $clingoFloor - is the translation declining proofs it used to take?")
        $mark = 'FAIL'
    }
    Write-Output ("  {0}  clingo: {1}: {2} proof(s) exported and {3} control(s), floor {4}; {5} refusal(s) not exported" -f $mark, $cr.Label, $ex.Exported, ($ex.Files.Count - $ex.Exported), $clingoFloor, $ex.Refusals)
    foreach ($s in $ex.Skipped) { Write-Output ("          not exported: ""{0}"" - {1}" -f $s.Proof, $s.Reason) }
}
# The files themselves, only once every proof file is translated - a file
# that failed an earlier rule would make every one of its exports look stale.
if ($present.Count -gt 0 -and $exportsMade -eq $present.Count) {
    $onDisk = @()
    if (Test-Path -LiteralPath $ClingoDir) {
        $onDisk = @(Get-ChildItem -LiteralPath $ClingoDir -Filter 'proof-*.lp' -File | Where-Object { $_.Extension -eq '.lp' } | ForEach-Object { $_.Name } | Sort-Object)
    }
    $before = $failures.Count
    foreach ($nm in $wantFiles.Keys) {
        $lpPath = Join-Path $ClingoDir $nm
        if (-not (Test-Path -LiteralPath $lpPath)) {
            $failures.Add("tools/clingo/$nm is missing - run tools\proofs_lp.ps1, then clingo on it")
            continue
        }
        $have = [IO.File]::ReadAllText($lpPath) -replace "`r`n", "`n"
        if ($have -cne $wantFiles[$nm]) {
            $failures.Add("tools/clingo/$nm no longer matches its proof - run tools\proofs_lp.ps1, then clingo on it")
        }
    }
    foreach ($nm in $onDisk) {
        if (-not $wantFiles.Contains($nm)) {
            $failures.Add("tools/clingo/$nm is made by no proof now - run tools\proofs_lp.ps1, which removes it")
        }
    }
    $outOfStep = $failures.Count - $before
    if ($outOfStep -eq 0) {
        Write-Output ("  ok    clingo: all {0} file(s) in tools/clingo match the corpus" -f $wantFiles.Count)
    } else {
        Write-Output ("  FAIL  clingo: {0} file(s) in tools/clingo out of step with the corpus" -f $outOfStep)
    }
}

Write-Output ''
if ($failures.Count -eq 0) {
    Write-Output '=== CHECK: clean - every proof reads, is well-formed, names only real refusals, and its clingo export is current ==='
    exit 0
}
foreach ($f in $failures) { Write-Output "FAIL: $f" }
Write-Output ''
Write-Output "=== CHECK: FAILED - $($failures.Count) problem(s) ==="
exit 1
