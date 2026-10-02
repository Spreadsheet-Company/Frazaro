<#
export_names.ps1 - VLA_SentenceEngine.bas's name lists, exported to data (PORT.6, slice 6a).

WHY: the sentence engine holds six name lists in code - VBA's reserved words
and the names the generated code calls (CheckName's two refusals), the
built-in function words and what they compile to, and three small word lists
of the conditions grammar and the function phrases. The core (core/) reads
the same lists as data, never typed in twice: this script writes
scripts/names.vla, one form per entry, in the order each VBA list gives them:

    (reserved-name "<word>")                    IsReservedName
    (engine-call-name "<word>")                 IsEngineCallName
    (function-word of|nullary "<word>" "<target>")   RegisterBuiltinFuncWords
    (function-word-display "<text>")            its display strings
    (conditions-grammar-word "<word>")          IsConditionsGrammarWord
    (set-verb "<word>")                         IsSetVerb
    (after-value-word "<word>")                 IsAfterValueWord's own words
                                                (IsExprOpWord's are in words.vla)

The VBA stays the source (SD-18); this export is re-run when a list changes,
and tools/check_data_exports.ps1 fails when the two have drifted.
tools/check_engine_call_names.ps1 goes on holding IsEngineCallName's list to
its four sources in the engine; this export carries the agreed list to the
core.

WHAT IT READS: each list's own procedure body, by its shape - a `list = " a
b c "` line, `AddFnEntry mFnOf|mFnNullary, "word", "target"` lines, a
`For Each w In Array(...)` of strings (continuations joined), `Case "a", "b"`
lists and `t = "a" Or t = "b"` chains. A list whose shape yields no entry
stops the export.

WHAT IT WRITES: a VLA file, CRLF like every .vla in scripts/, UTF-8 without
a BOM, under a GENERATED stamp so Lint VLA refuses it by name.

House style: PowerShell 5.1, host-free, no Excel, no COM, no network.

Usage:  powershell -File tools\export_names.ps1          (writes scripts\names.vla)
        powershell -File tools\export_names.ps1 -Print   (writes the text to stdout)
#>
param([switch]$Print)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$srcPath  = Join-Path $repoRoot 'src/VLA_SentenceEngine.bas'
$outPath  = Join-Path $repoRoot 'scripts/names.vla'

$all = [System.IO.File]::ReadAllText($srcPath) -split "`r?`n"

# --- one procedure's code lines, comment-only lines dropped, continuations joined
function Get-ProcBody([string]$name) {
    $head = '^\s*(?:Public\s+|Private\s+)?(?:Function|Sub)\s+' + [regex]::Escape($name) + '\s*\('
    $start = -1
    for ($i = 0; $i -lt $all.Count; $i++) {
        if ($all[$i] -match $head) { $start = $i; break }
    }
    if ($start -lt 0) { throw "$name not found in $srcPath - has its shape changed?" }
    $body = New-Object System.Collections.Generic.List[string]
    $buf = ''
    for ($i = $start + 1; $i -lt $all.Count; $i++) {
        $line = $all[$i]
        if ($line -match '^\s*End\s+(?:Function|Sub)\b') {
            if ($buf.Length -gt 0) { $body.Add($buf) }
            return ,$body.ToArray()
        }
        if ($line.TrimStart().StartsWith("'")) { continue }
        if ($line -match '^(.*)_\s*$') { $buf += $Matches[1]; continue }
        $body.Add($buf + $line)
        $buf = ''
    }
    throw "$name has no End Function in $srcPath"
}

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
        throw "a Case key this script cannot read: '$($keyText.Substring($i))' in: $keyText"
    }
    return ,$keys.ToArray()
}

function ConvertFrom-VbaLiteral([string]$s) { return ($s -replace '""', '"') }

function ConvertTo-VlaString([string]$s) {
    return '"' + ($s -replace '\\', '\\' -replace '"', '\"') + '"'
}

# `list = " a b c "`: the words, in order.
function Read-ListLine([string[]]$body, [string]$table) {
    foreach ($line in $body) {
        $m = [regex]::Match($line, '^\s*list\s*=\s*"([^"]*)"\s*$')
        if ($m.Success) {
            $words = @($m.Groups[1].Value -split '\s+' | Where-Object { $_.Length -gt 0 })
            if ($words.Count -eq 0) { throw "$table's list line is empty in $srcPath" }
            return ,$words
        }
    }
    throw "$table has no list = ""..."" line in $srcPath - has its shape changed?"
}

# `Case "a", "b", "c"` lines (the arm sets the function True), in order.
function Read-ListCases([string[]]$body) {
    $words = New-Object System.Collections.Generic.List[string]
    foreach ($line in $body) {
        if ($line -notmatch '^\s*Case\s+"') { continue }
        if ($line -match ':\s*[A-Za-z]+\s*=') { continue }
        foreach ($k in (Read-CaseKeys $line.Trim().Substring(5))) { $words.Add($k) }
    }
    return ,$words.ToArray()
}

# `t = "a" Or t = "b"` chains, in order.
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
$lines.Add('; GENERATED by tools/export_names.ps1 from src/VLA_SentenceEngine.bas - do not hand-edit or lint; the VBA is the source (PORT.6).')
$lines.Add('; One form per entry, each list in its VBA order: (reserved-name "<word>"), (engine-call-name "<word>"),')
$lines.Add('; (function-word of|nullary "<word>" "<target>"), (function-word-display "<text>"), (conditions-grammar-word "<word>"),')
$lines.Add('; (set-verb "<word>"), (after-value-word "<word>").')
$count = 0

$reserved = Read-ListLine (Get-ProcBody 'IsReservedName') 'IsReservedName'
foreach ($w in $reserved) { $lines.Add(('(reserved-name {0})' -f (ConvertTo-VlaString $w))); $count++ }

$engine = Read-ListLine (Get-ProcBody 'IsEngineCallName') 'IsEngineCallName'
foreach ($w in $engine) { $lines.Add(('(engine-call-name {0})' -f (ConvertTo-VlaString $w))); $count++ }

$rbf = Get-ProcBody 'RegisterBuiltinFuncWords'
$fnWords = New-Object System.Collections.Generic.List[string]
$display = New-Object System.Collections.Generic.List[string]
foreach ($line in $rbf) {
    $m = [regex]::Match($line, 'AddFnEntry\s+mFn(Of|Nullary)\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"')
    if ($m.Success) {
        $kind = if ($m.Groups[1].Value -eq 'Of') { 'of' } else { 'nullary' }
        $fnWords.Add(('(function-word {0} {1} {2})' -f $kind, (ConvertTo-VlaString $m.Groups[2].Value), (ConvertTo-VlaString $m.Groups[3].Value)))
        continue
    }
    $a = [regex]::Match($line, 'For Each\s+\w+\s+In\s+Array\((.*)\)\s*$')
    if ($a.Success) {
        foreach ($k in (Read-CaseKeys $a.Groups[1].Value)) { $display.Add(('(function-word-display {0})' -f (ConvertTo-VlaString $k))) }
    }
}
Assert-Some $fnWords.ToArray() 'RegisterBuiltinFuncWords (AddFnEntry)'
Assert-Some $display.ToArray() 'RegisterBuiltinFuncWords (display strings)'
foreach ($f in $fnWords) { $lines.Add($f); $count++ }
foreach ($d in $display) { $lines.Add($d); $count++ }

$cond = Read-ListCases (Get-ProcBody 'IsConditionsGrammarWord')
Assert-Some $cond 'IsConditionsGrammarWord'
foreach ($w in $cond) { $lines.Add(('(conditions-grammar-word {0})' -f (ConvertTo-VlaString $w))); $count++ }

$setVerbs = Read-OrChain (Get-ProcBody 'IsSetVerb') 't'
Assert-Some $setVerbs 'IsSetVerb'
foreach ($w in $setVerbs) { $lines.Add(('(set-verb {0})' -f (ConvertTo-VlaString $w))); $count++ }

$after = Read-ListCases (Get-ProcBody 'IsAfterValueWord')
Assert-Some $after 'IsAfterValueWord'
foreach ($w in $after) { $lines.Add(('(after-value-word {0})' -f (ConvertTo-VlaString $w))); $count++ }

$text = ($lines -join "`r`n") + "`r`n"
if ($Print) {
    # Through the pipeline, so a caller can capture it (check_data_exports.ps1 does).
    Write-Output $text
} else {
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($outPath, $text, $utf8)
    Write-Host "wrote $count entries -> $outPath"
}
