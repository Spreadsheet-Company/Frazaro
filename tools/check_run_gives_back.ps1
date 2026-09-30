<#
check_run_gives_back.ps1 - U.29's mechanical pin (the audit's C62).

U.29: a Run gives back Excel's settings as it found them. Before it, a Run
put back only screen updating: a program that turned calculation off, or put
words in the status bar, and then stopped left every open workbook in manual
calculation, or the words showing, for the rest of the Excel session, and
Undo Last Run - which puts back sheets - reached neither. VLA_IDE's
VlaIdeRecordExcel records the settings before a Run's first change, and
VlaIdeGiveBackExcel puts them back on every exit of RunProgram and
InterpretProgram (through GiveBackRun, which also takes away the program
sheet's protection when the Run made it).

WHY A STATIC SCAN AND NOT A TEST: C62 happened by drift. with-fast-excel set
calculation, and the phrasebook set the status bar, long before anything gave
either back, and nothing connects a phrasebook's (set! application.X ...) to
the Run's list. The next sentence that sets an Application property would
drift the same way, silently: the Run procedures are reached only by clicking
Run, which no suite does. The host suite pins the helpers themselves
(VLA_Tests_Host.TestU29GivesBackExcel); this pins what they must cover, and
that the Run procedures keep calling them.

WHAT IS CHECKED:
  A. Every Application property a phrasebook or the prelude sets - each
     (set! application.<member> ...) in scripts/prelude.vla and
     scripts/polyglotta/*.vla - is read in VlaIdeRecordExcel and written in
     VlaIdeGiveBackExcel (src/VLA_IDE.bas).
  B. HOW MANY times each Run procedure records and gives back, pinned below:
     RunProgram records once and gives back at three exits (the missing-main
     guard, the return from the program, the failure handler), and
     InterpretProgram records once and gives back at two (the return, the
     failure handler). An exit added without its give-back, or a give-back
     removed, moves a count.
  C. GiveBackRun calls VlaIdeGiveBackExcel.

WHAT IS NOT CHECKED: a compiled program's hand-written VLA can set any
Application property VBA can. The list is what Frazaro's own sentences and
macros reach; the interpreter reaches the same five properties, DynamicSet's
closed list (SEC.1), and from English only through a phrasebook macro. A
write to a property of an Application sub-object (application.x.y) is
listed for review and never failed on: it is not a setting of Excel's own.

Host-independent: reads files only. NOT wired into VlaSelfTest, the house
ratchet shape - a step a human and CI run, not a per-run gate.

Usage:  pwsh -File tools/check_run_gives_back.ps1 [-RepoRoot <dir>]
Exit code: 0 if clean; 1 if anything above fails.
#>

param(
    [string]$RepoRoot
)

$ErrorActionPreference = 'Stop'
if (-not $RepoRoot) { $RepoRoot = Split-Path -Parent $PSScriptRoot }

# --- Baseline, hand-maintained. ---
# 2026-09-30, U.29: the Application properties the phrasebooks and the
# prelude set. A new one is not a failure by itself - it fails only if the
# Run does not give it back - but it is reported, so that adding one stays
# a reviewed act. EnableEvents is given back too (Frazaro's own handlers
# need it) though no sentence sets it.
$knownMembers = @('calculation', 'cutcopymode', 'displayalerts', 'screenupdating', 'statusbar')

# How many times each Run procedure records, and gives back.
$pins = [ordered]@{
    'RunProgram'       = @{ Record = 1; GiveBack = 3 }
    'InterpretProgram' = @{ Record = 1; GiveBack = 2 }
}

$idePath     = Join-Path $RepoRoot (Join-Path 'src' 'VLA_IDE.bas')
$preludePath = Join-Path $RepoRoot (Join-Path 'scripts' 'prelude.vla')
$polyDir     = Join-Path $RepoRoot (Join-Path 'scripts' 'polyglotta')

$failed = New-Object System.Collections.Generic.List[string]

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

Write-Output '=== A RUN GIVES BACK EXCEL''S SETTINGS (U.29) ==='

if (-not (Test-Path -LiteralPath $idePath)) {
    Write-Output "  missing: $idePath"
    exit 1
}
$ideLines = [IO.File]::ReadAllText($idePath) -split "`r?`n"

# --- A. What the phrasebooks and the prelude set. ---
$vlaFiles = @()
if (Test-Path -LiteralPath $preludePath) { $vlaFiles += Get-Item -LiteralPath $preludePath }
if (Test-Path -LiteralPath $polyDir) {
    $vlaFiles += Get-ChildItem -LiteralPath $polyDir -File -Filter '*.vla' | Sort-Object Name
}
$setPattern = '\(set!\s+application\.([A-Za-z_][A-Za-z0-9_]*)((?:\.[A-Za-z_][A-Za-z0-9_]*)*)'
$members = New-Object System.Collections.Generic.SortedSet[string]
$deeper  = New-Object System.Collections.Generic.SortedSet[string]
foreach ($f in $vlaFiles) {
    foreach ($line in ([IO.File]::ReadAllText($f.FullName) -split "`r?`n")) {
        if ($line.TrimStart().StartsWith(';')) { continue }
        foreach ($m in [regex]::Matches($line, $setPattern, 'IgnoreCase')) {
            $member = $m.Groups[1].Value.ToLowerInvariant()
            if ($m.Groups[2].Value.Length -gt 0) {
                [void]$deeper.Add("application.$member$($m.Groups[2].Value.ToLowerInvariant()) ($($f.Name))")
            } else {
                [void]$members.Add($member)
            }
        }
    }
}
Write-Output "Phrasebook and prelude files scanned: $($vlaFiles.Count)"

$recordCode = Get-ProcCode $ideLines 'VlaIdeRecordExcel'
$giveCode   = Get-ProcCode $ideLines 'VlaIdeGiveBackExcel'
if ($null -eq $recordCode) { $failed.Add('VlaIdeRecordExcel not found in src/VLA_IDE.bas') }
if ($null -eq $giveCode)   { $failed.Add('VlaIdeGiveBackExcel not found in src/VLA_IDE.bas') }

$recorded = New-Object System.Collections.Generic.HashSet[string]
$givenBack = New-Object System.Collections.Generic.HashSet[string]
if ($null -ne $recordCode) {
    foreach ($c in $recordCode) {
        foreach ($m in [regex]::Matches($c, 'Application\.([A-Za-z]+)')) { [void]$recorded.Add($m.Groups[1].Value.ToLowerInvariant()) }
    }
}
if ($null -ne $giveCode) {
    foreach ($c in $giveCode) {
        foreach ($m in [regex]::Matches($c, 'Application\.([A-Za-z]+)\s*=')) { [void]$givenBack.Add($m.Groups[1].Value.ToLowerInvariant()) }
    }
}

Write-Output ''
Write-Output 'Application properties the phrasebooks and the prelude set:'
foreach ($member in $members) {
    $inRecord = $recorded.Contains($member)
    $inGive   = $givenBack.Contains($member)
    $known    = if ($knownMembers -contains $member) { 'known' } else { 'NEW' }
    if ($inRecord -and $inGive) {
        Write-Output ("  {0,-16} recorded, given back   ({1})" -f $member, $known)
    } else {
        $why = @()
        if (-not $inRecord) { $why += 'not recorded in VlaIdeRecordExcel' }
        if (-not $inGive)   { $why += 'not given back in VlaIdeGiveBackExcel' }
        Write-Output ("  {0,-16} {1}   ({2})  FAIL" -f $member, ($why -join ', '), $known)
        $failed.Add("application.$member - $($why -join ', ')")
    }
}

$missing = $knownMembers | Where-Object { -not $members.Contains($_) }
if ($missing) {
    Write-Output ''
    Write-Output '=== BASELINE MEMBERS NO PHRASEBOOK SETS ANY MORE (drop them from $knownMembers above, or the scan broke) ==='
    $missing | ForEach-Object { Write-Output "  $_" }
    $failed.Add("baseline members not found by the scan: $($missing -join ', ')")
}
$newMembers = $members | Where-Object { $knownMembers -notcontains $_ }
if ($newMembers) {
    Write-Output ''
    Write-Output '=== NEW APPLICATION PROPERTIES SET BY A PHRASEBOOK (add them to $knownMembers above once reviewed) ==='
    $newMembers | ForEach-Object { Write-Output "  $_" }
}
if ($deeper.Count -gt 0) {
    Write-Output ''
    Write-Output 'Writes to an Application sub-object, listed for review, not failed on:'
    $deeper | ForEach-Object { Write-Output "  $_" }
}

# --- B. The Run procedures keep calling the helpers, at every exit. ---
Write-Output ''
Write-Output 'Run procedures (records / gives back, pinned):'
foreach ($proc in $pins.Keys) {
    $code = Get-ProcCode $ideLines $proc
    if ($null -eq $code) {
        Write-Output "  $proc not found  FAIL"
        $failed.Add("$proc not found in src/VLA_IDE.bas")
        continue
    }
    $rec  = @($code | Where-Object { $_ -match '\bVlaIdeRecordExcel\b' }).Count
    $give = @($code | Where-Object { $_ -match '\bGiveBackRun\b' }).Count
    $want = $pins[$proc]
    $ok = ($rec -eq $want.Record) -and ($give -eq $want.GiveBack)
    $mark = if ($ok) { '' } else { '  FAIL' }
    Write-Output ("  {0,-16} {1} / {2}   (pinned {3} / {4}){5}" -f $proc, $rec, $give, $want.Record, $want.GiveBack, $mark)
    if (-not $ok) {
        $failed.Add("$proc records $rec time(s) and gives back at $give exit(s); pinned $($want.Record) and $($want.GiveBack)")
    }
}

# --- C. GiveBackRun gives the settings back. ---
$gbr = Get-ProcCode $ideLines 'GiveBackRun'
if ($null -eq $gbr) {
    $failed.Add('GiveBackRun not found in src/VLA_IDE.bas')
} elseif (@($gbr | Where-Object { $_ -match '\bVlaIdeGiveBackExcel\b' }).Count -ne 1) {
    $failed.Add('GiveBackRun does not call VlaIdeGiveBackExcel exactly once')
}

Write-Output ''
if ($failed.Count -eq 0) {
    Write-Output "=== CHECK: clean - all $($members.Count) Application properties the phrasebooks set are recorded and given back, at every pinned exit ==="
} else {
    Write-Output "=== CHECK: $($failed.Count) problem(s) ==="
    $failed | ForEach-Object { Write-Output "  $_" }
    Write-Output 'A Run must give back every Excel setting a sentence can change (U.29 in docs/BETA_REARVIEW.md): read it in VlaIdeRecordExcel, write it back in VlaIdeGiveBackExcel, and give back at every exit of RunProgram and InterpretProgram.'
}
exit ([Math]::Min($failed.Count, 1))
