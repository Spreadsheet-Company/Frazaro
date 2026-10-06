<#
check_data_exports.ps1 - the exported data files agree with their VBA sources.

WHY: PORT.5 gave the core two catalogues as data, each exported once from the
VBA that stays their source, and PORT.6 (slice 6a) two more, the English
engine's word tables and name lists:

    core/data/headtable.vla   from src/VLA_HeadTable.bas       (tools/export_headtable.ps1)
    core/data/messages.vla    from src/VLA_Messages.bas        (tools/export_messages.ps1)
    core/data/words.vla       from src/VLA_English.bas         (tools/export_words.ps1)
    core/data/names.vla       from src/VLA_SentenceEngine.bas  (tools/export_names.ps1)

Two copies of one list, and nothing mechanical to hold them together, is the
shape check_devrig_mods_parity.ps1 was written for after eight recurrences
of the same drift. Here the drift would be quieter still: a new refusal id
or a new head word lands in the VBA, the VBA tests pass, and the core goes on
refusing with yesterday's catalogue until a proof under the treaty's fifth
oracle fails on a machine without Excel. This check runs each exporter again
in memory and compares its text with the file on disk, so the drift fails
here, on every push, with the one-line fix named.

ALSO HELD: each export's entry count as a floor that never goes down (the
house style's hardcoded, reviewable number). An exporter whose pattern stops
matching a reshaped VBA line would quietly export fewer entries, and a file
regenerated from it would agree with it perfectly; the floor is what catches
that.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; exit 0 clean, exit 1 with every difference named.

Usage:  powershell -File tools\check_data_exports.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# --- the floors: raise them when an entry is added, never lower them --------
$floors = @{
    'core/data/headtable.vla' = 65    # one row per core form (VLA_HeadTable.bas, IN5.0: 65)
    'core/data/messages.vla'  = 587   # one entry per refusal id (VLA_Messages.bas, LX.14; 584 with PORT.7's ten; 587 with PORT.8's three)
    'core/data/words.vla'     = 113   # one entry per word-table row (VLA_English.bas, LX.14; PORT.6 2026-10-02)
    'core/data/names.vla'     = 218   # one entry per name-list row (VLA_SentenceEngine.bas; PORT.6 2026-10-02: 211; U.31, the seven reserved words: 218)
}
$exports = @(
    @{ Data = 'core/data/headtable.vla'; Script = 'tools/export_headtable.ps1'; Form = '^\(head ' },
    @{ Data = 'core/data/messages.vla';  Script = 'tools/export_messages.ps1';  Form = '^\(message ' },
    @{ Data = 'core/data/words.vla';     Script = 'tools/export_words.ps1';     Form = '^\([a-z-]+ ' },
    @{ Data = 'core/data/names.vla';     Script = 'tools/export_names.ps1';     Form = '^\([a-z-]+ ' }
)

function Get-NormalizedLines([string]$text) {
    $t = $text -replace "`r`n", "`n" -replace "`r", "`n"
    return $t.TrimEnd("`n") -split "`n"
}

$failed = New-Object System.Collections.Generic.List[string]
Write-Output '=== DATA EXPORTS (the VBA is the source; the .vla is its export) ==='
foreach ($e in $exports) {
    $dataPath = Join-Path $repoRoot $e.Data
    $scriptPath = Join-Path $repoRoot $e.Script
    if (-not (Test-Path $dataPath)) {
        Write-Output ("  {0,-24} MISSING  FAIL" -f $e.Data)
        $failed.Add("$($e.Data) is missing - run powershell -File $($e.Script)")
        continue
    }
    $want = Get-NormalizedLines (& $scriptPath -Print | Out-String)
    $have = Get-NormalizedLines ([System.IO.File]::ReadAllText($dataPath))
    $entries = @($have | Where-Object { $_ -match $e.Form }).Count
    $n = [Math]::Max($want.Count, $have.Count)
    $firstDiff = -1
    for ($i = 0; $i -lt $n; $i++) {
        $w = if ($i -lt $want.Count) { $want[$i] } else { '<end of export>' }
        $h = if ($i -lt $have.Count) { $have[$i] } else { '<end of file>' }
        if ($w -cne $h) { $firstDiff = $i; break }
    }
    if ($firstDiff -ge 0) {
        Write-Output ("  {0,-24} {1,4} entries  DRIFTED at line {2}  FAIL" -f $e.Data, $entries, ($firstDiff + 1))
        Write-Output ("      file:   {0}" -f $have[[Math]::Min($firstDiff, $have.Count - 1)])
        Write-Output ("      source: {0}" -f $want[[Math]::Min($firstDiff, $want.Count - 1)])
        $failed.Add("$($e.Data) has drifted from its VBA source - run powershell -File $($e.Script) and commit the result")
    } elseif ($entries -lt $floors[$e.Data]) {
        Write-Output ("  {0,-24} {1,4} entries  BELOW THE FLOOR of {2}  FAIL" -f $e.Data, $entries, $floors[$e.Data])
        $failed.Add("$($e.Data) holds $entries entries, below its floor of $($floors[$e.Data]) - an exporter pattern has stopped matching, or an entry was removed (then lower the floor deliberately, with the reason)")
    } else {
        Write-Output ("  {0,-24} {1,4} entries  agrees with its source (floor {2})" -f $e.Data, $entries, $floors[$e.Data])
    }
}

Write-Output ''
if ($failed.Count -eq 0) {
    Write-Output '=== CHECK: clean - every exported data file agrees with the VBA it was exported from ==='
    exit 0
}
Write-Output "=== CHECK: $($failed.Count) export problem(s) ==="
$failed | ForEach-Object { Write-Output "  $_" }
exit 1
