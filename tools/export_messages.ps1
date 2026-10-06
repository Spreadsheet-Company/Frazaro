<#
export_messages.ps1 - the message catalogue, exported to data (PORT.5).

WHY: every refusal in the reference implementation is `RaiseMsg "<id>", ...`
over the catalogue in src/VLA_Messages.bas (SD-2), and the treaty's fifth
oracle holds every implementation to the same ids in the same situations.
The core (core/) refuses with the same catalogue, read as data, never typed in
twice: this script writes core/data/messages.vla, one form per entry,

    (message <id> <error-number> "<Err.Source>" "<template>")

in the order the VBA registers them. The VBA stays the source until PORT.6
hands the catalogue to the data file; until then this export is re-run when
the catalogue changes, and tools/check_data_exports.ps1 fails when the two
have drifted.

WHAT IT READS: every `AddMsg m, <id>, <number>, <source>, <template>` line of
AddEntries, with VBA's own line continuations joined, string literals read
with their doubled quotes undoubled, `&` concatenations evaluated, vbCrLf
spelled as a line break, and a named error constant (VLA.VLA_ERR_...)
resolved from its declaration in src/VLA.bas (vbObjectError + n).

WHAT IT WRITES: a VLA file, CRLF like every .vla in the repository, UTF-8 without
a BOM, under a GENERATED stamp so Lint VLA refuses it by name. A template's
line break is written as a line break inside the string literal; the core
normalizes line endings when it loads the file, so the five templates that
hold one read the same on every platform.

House style: PowerShell 5.1, host-free, no Excel, no COM, no network.

Usage:  powershell -File tools\export_messages.ps1          (writes core\data\messages.vla)
        powershell -File tools\export_messages.ps1 -Print   (writes the text to stdout)
#>
param([switch]$Print)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$srcPath  = Join-Path $repoRoot 'src/VLA_Messages.bas'
$vlaPath  = Join-Path $repoRoot 'src/VLA.bas'
$outPath  = Join-Path $repoRoot 'core/data/messages.vla'

# --- named error constants a catalogue line may use as its number ----------
function Get-NamedConstants([string]$path) {
    $consts = @{ 'vbObjectError' = [long]-2147221504 }
    foreach ($line in Get-Content -LiteralPath $path) {
        if ($line -match '^\s*Public Const (VLA_ERR_\w+) As Long = ([^'']+)') {
            $name = $Matches[1]
            $value = [long]0
            foreach ($term in ($Matches[2] -split '\+')) {
                $t = $term.Trim()
                if ($t -match '^-?\d+$') { $value += [long]$t }
                elseif ($consts.ContainsKey($t)) { $value += $consts[$t] }
                else { throw "cannot evaluate '$t' in the declaration of $name" }
            }
            $consts[$name] = $value
        }
    }
    return $consts
}
$named = Get-NamedConstants $vlaPath

# --- the AddEntries body, continuations joined ------------------------------
function Get-LogicalStatements([string[]]$lines) {
    $stmts = New-Object System.Collections.Generic.List[string]
    $buf = ''
    foreach ($line in $lines) {
        if ($line -match '^(.*)_\s*$') { $buf += $Matches[1]; continue }
        $buf += $line
        $stmts.Add($buf)
        $buf = ''
    }
    if ($buf.Length -gt 0) { $stmts.Add($buf) }
    return $stmts
}

$all = Get-Content -LiteralPath $srcPath
$start = -1; $end = -1
for ($i = 0; $i -lt $all.Count; $i++) {
    if ($start -lt 0 -and $all[$i] -match '^Private Sub AddEntries\(') { $start = $i; continue }
    if ($start -ge 0 -and $all[$i] -match '^End Sub') { $end = $i; break }
}
if ($start -lt 0 -or $end -lt 0) { throw "AddEntries not found in $srcPath - has its shape changed?" }
$body = Get-LogicalStatements $all[($start + 1)..($end - 1)]

# --- a VBA argument list: string literals, names, numbers, & and , ----------
function Read-VbaArgs([string]$text) {
    # Returns a list of arguments; each argument is the evaluated value of
    # its `&`-joined parts (a string, or a number when every part is one).
    $values = New-Object System.Collections.Generic.List[object]
    $parts = New-Object System.Collections.Generic.List[object]
    $i = 0; $n = $text.Length
    $finish = {
        if ($parts.Count -eq 0) { throw "empty argument in: $text" }
        if ($parts.Count -eq 1 -and $parts[0] -is [long]) { $values.Add($parts[0]) }
        else { $values.Add(($parts | ForEach-Object { [string]$_ }) -join '') }
        $parts.Clear()
    }
    while ($i -lt $n) {
        $c = $text[$i]
        if ($c -eq ' ' -or $c -eq "`t") { $i++; continue }
        if ($c -eq '"') {
            $sb = New-Object System.Text.StringBuilder
            $i++
            while ($true) {
                if ($i -ge $n) { throw "unterminated string in: $text" }
                if ($text[$i] -eq '"') {
                    if ($i + 1 -lt $n -and $text[$i + 1] -eq '"') { [void]$sb.Append('"'); $i += 2; continue }
                    $i++; break
                }
                [void]$sb.Append($text[$i]); $i++
            }
            $parts.Add($sb.ToString()); continue
        }
        if ($c -eq '&') { $i++; continue }
        if ($c -eq ',') { & $finish; $i++; continue }
        if ($c -eq "'") { break }   # a trailing comment
        $j = $i
        while ($j -lt $n -and $text[$j] -notmatch '[\s,&"]') { $j++ }
        $word = $text.Substring($i, $j - $i); $i = $j
        if ($word -match '^-?\d+$') { $parts.Add([long]$word); continue }
        if ($word -eq 'vbCrLf') { $parts.Add("`r`n"); continue }
        $bare = $word -replace '^VLA\.', ''
        if ($named.ContainsKey($bare)) { $parts.Add([long]$named[$bare]); continue }
        throw "cannot evaluate '$word' in: $text"
    }
    & $finish
    return $values
}

function ConvertTo-VlaString([string]$s) {
    return '"' + ($s -replace '\\', '\\' -replace '"', '\"') + '"'
}

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('; GENERATED by tools/export_messages.ps1 from src/VLA_Messages.bas - do not hand-edit or lint; the VBA is the source (PORT.5).')
$lines.Add('; (message <id> <error-number> "<Err.Source>" "<template>"), in the order the VBA registers them.')
$count = 0
foreach ($stmt in $body) {
    if ($stmt -notmatch '^\s*AddMsg m,\s*(.*)$') { continue }
    $a = Read-VbaArgs $Matches[1]
    if ($a.Count -ne 4) { throw "expected 4 arguments, found $($a.Count) in: $stmt" }
    $id = [string]$a[0]
    if ($id -notmatch '^[a-z0-9?!-]+$') { throw "an id that is not one symbol: '$id'" }
    if ($a[1] -isnot [long]) { throw "the number of '$id' is not a number: $($a[1])" }
    $lines.Add(('(message {0} {1} {2} {3})' -f $id, $a[1], (ConvertTo-VlaString ([string]$a[2])), (ConvertTo-VlaString ([string]$a[3]))))
    $count++
}
if ($count -eq 0) { throw "no AddMsg line matched - the catalogue's shape has changed; update this script" }

$text = ($lines -join "`r`n") + "`r`n"
if ($Print) {
    # Through the pipeline, so a caller can capture it (check_data_exports.ps1 does).
    Write-Output $text
} else {
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($outPath, $text, $utf8)
    Write-Host "wrote $count messages -> $outPath"
}
