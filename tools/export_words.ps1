<#
export_words.ps1 - VLA_English.bas's word tables, exported to data (PORT.6, slice 6a).

WHY: the English engine holds nine word tables in code - the number and
ordinal words the tokenizer and the expression grammar rewrite, the noise and
dropped words, the operator words, the colour words, the slot descriptions a
refusal teaches with, and the stray-character hints. The core (core/) reads
the same tables as data, never typed in twice: this script writes
scripts/words.vla, one form per entry, in the order each VBA table lists
them:

    (expr-op "<symbol>" "<words>")          ExprOpWord
    (expr-op-word "<word>")                 IsExprOpWord
    (noise-word "<word>")                   IsNoiseWord
    (dropped-word "<word>")                 IsDroppedWord
    (stray-char-hint "<char>" "<hint>")     StrayCharHint, one per Case key
    (stray-char-default "<hint>")           StrayCharHint's Case Else
    (number-word "<word>" "<digits>")       NumberWord
    (ordinal-word "<word>" "<digits>")      OrdinalWord
    (color-word "<word>")                   IsColorWord
    (slot-desc "<category>" "<text>")       SlotDesc, one per Case key

The VBA stays the source (SD-18); this export is re-run when a table
changes, and tools/check_data_exports.ps1 fails when the two have drifted.

WHAT IT READS: each table's own procedure body, by its shape - `Case "k":
F = "v"` arms (keys may be several per arm, and a key may be ChrW$(n), which
is written as the character itself), `Case "a", "b"` lists that set the
function True, and `w = "a" Or w = "b"` chains. A table whose shape yields
no entry stops the export, so a reshaped procedure cannot export an empty
table quietly.

WHAT IT WRITES: a VLA file, CRLF like every .vla in scripts/, UTF-8 without
a BOM (two stray-character keys are curly punctuation), under a GENERATED
stamp so Lint VLA refuses it by name.

House style: PowerShell 5.1, host-free, no Excel, no COM, no network.

Usage:  powershell -File tools\export_words.ps1          (writes scripts\words.vla)
        powershell -File tools\export_words.ps1 -Print   (writes the text to stdout)
#>
param([switch]$Print)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$srcPath  = Join-Path $repoRoot 'src/VLA_English.bas'
$outPath  = Join-Path $repoRoot 'scripts/words.vla'

$all = [System.IO.File]::ReadAllText($srcPath) -split "`r?`n"

# --- one procedure's code lines (comment-only lines dropped) ----------------
function Get-ProcBody([string]$name) {
    $head = '^\s*(?:Public\s+|Private\s+)?(?:Function|Sub)\s+' + [regex]::Escape($name) + '\s*\('
    $start = -1
    for ($i = 0; $i -lt $all.Count; $i++) {
        if ($all[$i] -match $head) { $start = $i; break }
    }
    if ($start -lt 0) { throw "$name not found in $srcPath - has its shape changed?" }
    $body = New-Object System.Collections.Generic.List[string]
    for ($i = $start + 1; $i -lt $all.Count; $i++) {
        if ($all[$i] -match '^\s*End\s+(?:Function|Sub)\b') { return ,$body.ToArray() }
        if ($all[$i].TrimStart().StartsWith("'")) { continue }
        $body.Add($all[$i])
    }
    throw "$name has no End Function in $srcPath"
}

# --- the keys of one Case arm: "lit" (with "" undoubled) or ChrW$(n) --------
function Read-CaseKeys([string]$keyText) {
    $keys = New-Object System.Collections.Generic.List[string]
    $i = 0; $n = $keyText.Length
    while ($i -lt $n) {
        $c = $keyText[$i]
        if ($c -eq ' ' -or $c -eq ',' -or $c -eq "`t") { $i++; continue }
        if ($c -eq '"') {
            $sb = New-Object System.Text.StringBuilder
            $i++
            while ($true) {
                if ($i -ge $n) { throw "unterminated string in Case keys: $keyText" }
                if ($keyText[$i] -eq '"') {
                    if ($i + 1 -lt $n -and $keyText[$i + 1] -eq '"') { [void]$sb.Append('"'); $i += 2; continue }
                    $i++; break
                }
                [void]$sb.Append($keyText[$i]); $i++
            }
            $keys.Add($sb.ToString()); continue
        }
        $m = [regex]::Match($keyText.Substring($i), '^ChrW\$\((\d+)\)')
        if ($m.Success) {
            $keys.Add(([char][int]$m.Groups[1].Value).ToString())
            $i += $m.Length; continue
        }
        throw "a Case key this script cannot read: '$($keyText.Substring($i))' in: $keyText"
    }
    return ,$keys.ToArray()
}

function ConvertFrom-VbaLiteral([string]$s) { return ($s -replace '""', '"') }

function ConvertTo-VlaString([string]$s) {
    return '"' + ($s -replace '\\', '\\' -replace '"', '\"') + '"'
}

# `Case k1, k2: Name = "value"` arms, in order; Case Else separately.
function Read-KeyedCases([string[]]$body, [ref]$default) {
    $pairs = New-Object System.Collections.Generic.List[object]
    foreach ($line in $body) {
        $me = [regex]::Match($line, '^\s*Case\s+Else\s*:\s*[A-Za-z]+\s*=\s*"((?:[^"]|"")*)"\s*$')
        if ($me.Success) { $default.Value = ConvertFrom-VbaLiteral $me.Groups[1].Value; continue }
        $m = [regex]::Match($line, '^\s*Case\s+(.+?):\s*[A-Za-z]+\s*=\s*"((?:[^"]|"")*)"\s*$')
        if (-not $m.Success) { continue }
        $value = ConvertFrom-VbaLiteral $m.Groups[2].Value
        foreach ($k in (Read-CaseKeys $m.Groups[1].Value)) {
            $pairs.Add([pscustomobject]@{ Key = $k; Value = $value })
        }
    }
    return ,$pairs.ToArray()
}

# `Case "a", "b", "c"` lines (the arm sets the function True), in order.
function Read-ListCases([string[]]$body) {
    $words = New-Object System.Collections.Generic.List[string]
    foreach ($line in $body) {
        if ($line -notmatch '^\s*Case\s+"') { continue }
        if ($line -match ':\s*[A-Za-z]+\s*=') { continue }   # a keyed arm, not a list
        foreach ($k in (Read-CaseKeys $line.Trim().Substring(5))) { $words.Add($k) }
    }
    return ,$words.ToArray()
}

# `w = "a" Or w = "b"` chains, in order.
function Read-OrChain([string[]]$body, [string]$var) {
    $words = New-Object System.Collections.Generic.List[string]
    foreach ($line in $body) {
        foreach ($m in [regex]::Matches($line, [regex]::Escape($var) + '\s*=\s*"((?:[^"]|"")*)"')) {
            $words.Add((ConvertFrom-VbaLiteral $m.Groups[1].Value))
        }
    }
    return ,$words.ToArray()
}

function Assert-Some([object[]]$items, [string]$table) {
    if ($null -eq $items -or $items.Count -eq 0) { throw "$table exported no entry - its shape in $srcPath has changed; update this script" }
}

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('; GENERATED by tools/export_words.ps1 from src/VLA_English.bas - do not hand-edit or lint; the VBA is the source (PORT.6).')
$lines.Add('; One form per table entry, each table in its VBA order: (expr-op "<symbol>" "<words>"), (expr-op-word "<word>"),')
$lines.Add('; (noise-word "<word>"), (dropped-word "<word>"), (stray-char-hint "<char>" "<hint>"), (stray-char-default "<hint>"),')
$lines.Add('; (number-word "<word>" "<digits>"), (ordinal-word "<word>" "<digits>"), (color-word "<word>"), (slot-desc "<category>" "<text>").')
$count = 0
$noDefault = $null

$exprOps = Read-KeyedCases (Get-ProcBody 'ExprOpWord') ([ref]$noDefault)
Assert-Some $exprOps 'ExprOpWord'
foreach ($p in $exprOps) { $lines.Add(('(expr-op {0} {1})' -f (ConvertTo-VlaString $p.Key), (ConvertTo-VlaString $p.Value))); $count++ }

$opWords = Read-ListCases (Get-ProcBody 'IsExprOpWord')
Assert-Some $opWords 'IsExprOpWord'
foreach ($w in $opWords) { $lines.Add(('(expr-op-word {0})' -f (ConvertTo-VlaString $w))); $count++ }

$noise = Read-OrChain (Get-ProcBody 'IsNoiseWord') 'w'
Assert-Some $noise 'IsNoiseWord'
foreach ($w in $noise) { $lines.Add(('(noise-word {0})' -f (ConvertTo-VlaString $w))); $count++ }

$dropped = Read-OrChain (Get-ProcBody 'IsDroppedWord') 'w'
Assert-Some $dropped 'IsDroppedWord'
foreach ($w in $dropped) { $lines.Add(('(dropped-word {0})' -f (ConvertTo-VlaString $w))); $count++ }

$strayDefault = $null
$stray = Read-KeyedCases (Get-ProcBody 'StrayCharHint') ([ref]$strayDefault)
Assert-Some $stray 'StrayCharHint'
if ($null -eq $strayDefault) { throw "StrayCharHint has no Case Else - its shape in $srcPath has changed; update this script" }
foreach ($p in $stray) { $lines.Add(('(stray-char-hint {0} {1})' -f (ConvertTo-VlaString $p.Key), (ConvertTo-VlaString $p.Value))); $count++ }
$lines.Add(('(stray-char-default {0})' -f (ConvertTo-VlaString $strayDefault))); $count++

$numbers = Read-KeyedCases (Get-ProcBody 'NumberWord') ([ref]$noDefault)
Assert-Some $numbers 'NumberWord'
foreach ($p in $numbers) { $lines.Add(('(number-word {0} {1})' -f (ConvertTo-VlaString $p.Key), (ConvertTo-VlaString $p.Value))); $count++ }

$ordinals = Read-KeyedCases (Get-ProcBody 'OrdinalWord') ([ref]$noDefault)
Assert-Some $ordinals 'OrdinalWord'
foreach ($p in $ordinals) { $lines.Add(('(ordinal-word {0} {1})' -f (ConvertTo-VlaString $p.Key), (ConvertTo-VlaString $p.Value))); $count++ }

$colors = Read-ListCases (Get-ProcBody 'IsColorWord')
Assert-Some $colors 'IsColorWord'
foreach ($w in $colors) { $lines.Add(('(color-word {0})' -f (ConvertTo-VlaString $w))); $count++ }

$slots = Read-KeyedCases (Get-ProcBody 'SlotDesc') ([ref]$noDefault)
Assert-Some $slots 'SlotDesc'
foreach ($p in $slots) { $lines.Add(('(slot-desc {0} {1})' -f (ConvertTo-VlaString $p.Key), (ConvertTo-VlaString $p.Value))); $count++ }

$text = ($lines -join "`r`n") + "`r`n"
if ($Print) {
    # Through the pipeline, so a caller can capture it (check_data_exports.ps1 does).
    Write-Output $text
} else {
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($outPath, $text, $utf8)
    Write-Host "wrote $count entries -> $outPath"
}
