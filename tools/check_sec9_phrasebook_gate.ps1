<#
check_sec9_phrasebook_gate.ps1 - SEC.9's mechanical pin.

SEC.9: grammar decides what every sentence MEANS, so whoever supplies the
phrasebook decides what a program does. Two loaders trusted the workbook's
own account of which grammar to use - `IdeVocabPath` preferred
`<host workbook dir>\...\<edition>.vla` over the add-in's own copy, and
`ReplayPersistedPhrasebooks` replayed absolute paths out of the workbook's
`VLA_LoadedPhrasebooks` property with no prompt. Observed live 2026-09-08,
benignly: a stale english.vla in a Downloads folder loaded ahead of the
add-in's own copy. `PhrasebookPathApproved` closes both.

WHY A STATIC SCAN AND NOT A TEST: the pure suite covers the whole string
DECISION - `VlaPhrasebookPathIsRemote` and `VlaPhrasebookPathIsUnderDir`,
in VLA_Tests.bas's TestSec9PhrasebookPaths - but it cannot reach the consent
record (a registry) or the dialog (a person). That leaves the CALL SITES
uncovered, and a call site is exactly where this item can silently die: drop
the gate from one loader and everything still works, on this machine, for
the person who already approved the file. Same reasoning
check_sec8_provenance_gate.ps1 and check_word_automation_security.ps1 gave
for their own items, and the same shape.

WHAT IS CHECKED, property 1 - both untrusted loaders call the gate. The two
procedures below must each contain a `PhrasebookPathApproved` call. They are
the only two places in the codebase where a phrasebook path is derived from
something the workbook controls rather than named by a person at a dialog.

WHAT IS CHECKED, property 2 - and this is the one that actually matters -
the gate PRECEDES the filesystem probe in the replay loop. A path out of the
workbook is an arbitrary attacker-chosen string, and `Dir$` on
`\\attacker\share\x.vla` hands this machine's Windows credentials to that
server before it returns anything. So in `ReplayPersistedPhrasebooks` the
`PhrasebookPathApproved` call must appear on an earlier line than any
`SafeFileExists`. An edit that "tidies up" by hoisting the existence check
to the top of the loop would reintroduce the credential leak while leaving
every visible behaviour, and every test, exactly as it was. Nothing else in
this project would notice. That is what this property exists for.

`IdeVocabPath` deliberately does NOT carry property 2, and the asymmetry is
the design rather than an exemption: its candidates live in the host
workbook's own directory, which Excel already opened the workbook from, so
probing it discloses nothing that opening the file has not already
disclosed.

WHAT IS CHECKED, property 3 - the pure predicates still exist and are still
Public. They are the whole security decision, and VlaSelfTest reaches them
only because they are Public; a well-meaning "these are only used inside
this module" pass would make them Private and silently unpin the item.

Usage:  powershell -File tools\check_sec9_phrasebook_gate.ps1
Exit 0 clean, exit 1 with every failure listed.
#>

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$idePath = Join-Path $root 'src\VLA_IDE.bas'
if (-not (Test-Path -LiteralPath $idePath)) { Write-Error "Missing $idePath" }
$lines = [System.IO.File]::ReadAllLines($idePath)

$failures = New-Object System.Collections.Generic.List[string]
Write-Output '=== SEC.9 PHRASEBOOK PATH GATE ==='
Write-Output ("Source: {0} ({1} lines)" -f $idePath, $lines.Count)
Write-Output ''

# --- procedure boundaries, the same way the SEC.8 scan finds them ----
function Get-ProcBody([string]$name) {
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match ("^\s*(Public|Private)?\s*(Sub|Function)\s+" + [regex]::Escape($name) + "\s*\(")) { $start = $i; break }
    }
    if ($start -lt 0) { return $null }
    for ($j = $start + 1; $j -lt $lines.Count; $j++) {
        if ($lines[$j] -match '^\s*End\s+(Sub|Function)\s*$') {
            return [pscustomobject]@{ Start = $start; End = $j }
        }
    }
    return $null
}

# A line that is code, not a comment. Every property below is about what
# the code DOES; a mention in a comment must never satisfy a check.
function Test-CodeLine([string]$line, [string]$needle) {
    $t = $line.TrimStart()
    if ($t.StartsWith("'")) { return $false }
    return $line.Contains($needle)
}

# ---- property 1: both untrusted loaders call the gate ---------------
Write-Output '--- the two workbook-controlled loaders are gated ---'
$gated = @('ReplayPersistedPhrasebooks', 'IdeVocabPath')
$bodies = @{}
foreach ($procName in $gated) {
    $b = Get-ProcBody $procName
    if ($null -eq $b) {
        $failures.Add("VLA_IDE.bas no longer defines $procName - this scan cannot check a procedure it cannot find. If it was renamed, rename it here too.")
        Write-Output ("  FAIL  {0,-30} not found" -f $procName)
        continue
    }
    $bodies[$procName] = $b
    $hit = $false
    for ($i = $b.Start; $i -le $b.End; $i++) {
        if (Test-CodeLine $lines[$i] 'PhrasebookPathApproved') { $hit = $true; break }
    }
    if ($hit) {
        Write-Output ("  ok    {0,-30} lines {1}-{2}, gated" -f $procName, ($b.Start + 1), ($b.End + 1))
    } else {
        $failures.Add("$procName derives a phrasebook path from something the workbook controls but never calls PhrasebookPathApproved - that is SEC.9 reopening.")
        Write-Output ("  FAIL  {0,-30} NO gate call" -f $procName)
    }
}

# ---- property 2: gate before probe, in the replay loop --------------
Write-Output ''
Write-Output '--- the replay gate runs BEFORE any filesystem probe (the credential leak) ---'
if ($bodies.ContainsKey('ReplayPersistedPhrasebooks')) {
    $b = $bodies['ReplayPersistedPhrasebooks']
    $firstGate = -1; $firstProbe = -1
    for ($i = $b.Start; $i -le $b.End; $i++) {
        if ($firstGate -lt 0 -and (Test-CodeLine $lines[$i] 'PhrasebookPathApproved')) { $firstGate = $i }
        if ($firstProbe -lt 0 -and (Test-CodeLine $lines[$i] 'SafeFileExists'))        { $firstProbe = $i }
    }
    if ($firstProbe -lt 0) {
        Write-Output ("  ok    no SafeFileExists in the loop at all (gate at line {0})" -f ($firstGate + 1))
    } elseif ($firstGate -ge 0 -and $firstGate -lt $firstProbe) {
        Write-Output ("  ok    gate line {0} precedes probe line {1}" -f ($firstGate + 1), ($firstProbe + 1))
    } else {
        $failures.Add("In ReplayPersistedPhrasebooks a SafeFileExists (line $($firstProbe + 1)) runs at or before PhrasebookPathApproved (line $(if ($firstGate -lt 0) { 'none' } else { $firstGate + 1 })). Dir\$ on a UNC path leaks this machine's Windows credentials to a server the WORKBOOK chose. The approval must come first.")
        Write-Output ("  FAIL  probe line {0} runs before the gate" -f ($firstProbe + 1))
    }
}

# ---- property 3: the pure predicates are still Public ---------------
Write-Output ''
Write-Output '--- the pure predicates are still present and Public (VlaSelfTest reaches them) ---'
foreach ($fn in @('VlaPhrasebookPathIsRemote', 'VlaPhrasebookPathIsUnderDir')) {
    $found = $false
    foreach ($l in $lines) {
        if ($l -match ("^\s*Public\s+Function\s+" + [regex]::Escape($fn) + "\s*\(")) { $found = $true; break }
    }
    if ($found) {
        Write-Output ("  ok    {0}" -f $fn)
    } else {
        $failures.Add("$fn is missing or no longer Public - it is the security decision itself, and VlaSelfTest's TestSec9PhrasebookPaths can only pin it while it is Public.")
        Write-Output ("  FAIL  {0} missing or not Public" -f $fn)
    }
}

# ---- property 4: non-loading callers never prompt -------------------
# A consent gate inside a path RESOLVER reaches callers that only want to
# know which file would be used. Two of them load no grammar at all -
# VlaIdeInfo builds a one-line diagnostics string, and the
# "ide-vocab-not-found" site is assembling the text of an error already
# being raised - and both must pass mayPrompt:=False. A status report
# that stops to ask a security question is a defect however good the
# question is. Found by reading the call sites after the first live
# pass; pinned here so it cannot come back.
Write-Output ''
Write-Output '--- callers that load no grammar pass mayPrompt:=False ---'
$nonLoading = @(
    @{ Proc = 'VlaIdeInfo';  Why = 'builds a diagnostics string' },
    @{ Proc = 'IdeLoadVocab'; Why = 'the ide-vocab-not-found message argument'; OnlyRaiseMsgLine = $true }
)
foreach ($nl in $nonLoading) {
    $b = Get-ProcBody $nl.Proc
    if ($null -eq $b) {
        $failures.Add("VLA_IDE.bas no longer defines $($nl.Proc) - this scan cannot check a procedure it cannot find.")
        Write-Output ("  FAIL  {0,-16} not found" -f $nl.Proc)
        continue
    }
    $bad = @()
    for ($i = $b.Start; $i -le $b.End; $i++) {
        if (-not (Test-CodeLine $lines[$i] 'IdeVocabPath(')) { continue }
        if ($nl.OnlyRaiseMsgLine -and -not $lines[$i].Contains('RaiseMsg')) { continue }
        if (-not $lines[$i].Contains('mayPrompt:=False')) { $bad += ($i + 1) }
    }
    if ($bad.Count -eq 0) {
        Write-Output ("  ok    {0,-16} ({1})" -f $nl.Proc, $nl.Why)
    } else {
        $failures.Add("$($nl.Proc) ($($nl.Why)) calls IdeVocabPath without mayPrompt:=False at line(s) $($bad -join ', ') - it loads no grammar, so it must never be able to raise a consent dialog.")
        Write-Output ("  FAIL  {0,-16} line(s) {1} may prompt" -f $nl.Proc, ($bad -join ', '))
    }
}

# ---- verdict --------------------------------------------------------
Write-Output ''
if ($failures.Count -eq 0) {
    Write-Output '=== CHECK: clean - both workbook-controlled loaders gated, approval precedes probe, predicates pinned ==='
    exit 0
} else {
    foreach ($f in $failures) { Write-Output "FAIL: $f" }
    Write-Output ''
    Write-Output "=== CHECK: FAILED - $($failures.Count) problem(s) ==="
    exit 1
}
