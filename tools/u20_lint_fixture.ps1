<#
u20_lint_fixture.ps1 - U.20's live-test fixtures for Lint VLA.

-Make builds one folder per live test under -Root (default C:\U20), each
holding only that test's own files, and records every file's SHA-256 and
write time. -Check reads the disk back and prints PASS/FAIL per test for
everything a FILE can show: bytes, hashes, write times, leftover
.lint-tmp copies. What a dialog SAYS is read by the person at the
keyboard; this script only checks what reached the disk.

Usage:
  powershell -ExecutionPolicy Bypass -File tools\u20_lint_fixture.ps1 -Make
  (run the numbered steps in Excel)
  powershell -ExecutionPolicy Bypass -File tools\u20_lint_fixture.ps1 -Check

-Make deletes -Root first, but only a folder this script made (it carries
u20.marker); any other existing folder is refused. Not wired into any
check or test run - a live-test aid, like sec13_word_fixture.md.

Non-ASCII fixture text is built from [char] codes so this file stays ASCII.
#>
param([switch]$Make, [switch]$Check, [string]$Root = 'C:\U20')
$ErrorActionPreference = 'Stop'

$repo      = Split-Path -Parent $PSScriptRoot
$expanded  = Join-Path $repo 'scripts\polyglotta\english_expanded.vla'
$utf8      = New-Object System.Text.UTF8Encoding($false)
$utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
$CRLF      = "`r`n"
$pound     = [string][char]0x00A3
$emDash    = [string][char]0x2014
$eAcute    = [string][char]0x00E9

$poundText = "; price in $pound, a dash $emDash and caf$eAcute" + $CRLF + "(debug-print `"caf$eAcute`")"
$ansiBytes = [byte[]](0x28, 0x61, 0x29, 0x0D, 0x0A, 0x3B, 0x20, 0xA3, 0x0D, 0x0A)
$commentText = "(begin" + $CRLF + "  ; keep me" + $CRLF + "  (stop))" + $CRLF
$fourBreaks = $CRLF + $CRLF + $CRLF + $CRLF

function Write-Text([string]$path, [string]$text) {
    [System.IO.File]::WriteAllBytes($path, $utf8.GetBytes($text))
}
function Hex([byte[]]$b) { return [BitConverter]::ToString($b) }
function Sha([string]$path) { return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash }
function Say([string]$name, [bool]$ok, [string]$detail) {
    if ($ok) { Write-Output "PASS  $name" } else { Write-Output "FAIL  $name - $detail" }
}

if ($Make) {
    if (Test-Path -LiteralPath $Root) {
        if (-not (Test-Path -LiteralPath (Join-Path $Root 'u20.marker'))) {
            throw "$Root exists and was not made by this script (no u20.marker) - pick another -Root."
        }
        Remove-Item -LiteralPath $Root -Recurse -Force
    }
    foreach ($d in 't1', 't2', 't3', 't4', 't5', 't6', 't7', 't8') {
        New-Item -ItemType Directory -Force -Path (Join-Path $Root $d) | Out-Null
    }
    Write-Text (Join-Path $Root 'u20.marker') 'made by tools\u20_lint_fixture.ps1'

    Write-Text (Join-Path $Root 't1\pound.vla') $poundText
    Write-Text (Join-Path $Root 't2\clean.vla') ($poundText + $fourBreaks)
    Copy-Item -LiteralPath $expanded -Destination (Join-Path $Root 't3\english_expanded.vla')
    [System.IO.File]::WriteAllBytes((Join-Path $Root 't4\ansi.vla'), $ansiBytes)
    Write-Text (Join-Path $Root 't5\comment.vla') $commentText

    Write-Text (Join-Path $Root 't6\messy.vla') '(stop)'
    Write-Text (Join-Path $Root 't6\clean.vla') ('(stop)' + $fourBreaks)
    Copy-Item -LiteralPath $expanded -Destination (Join-Path $Root 't6\english_expanded.vla')
    [System.IO.File]::WriteAllBytes((Join-Path $Root 't6\bad.vla'), $ansiBytes)
    Write-Text (Join-Path $Root 't6\comment.vla') $commentText

    Write-Text (Join-Path $Root 't7\pound.vla') $poundText
    Copy-Item -LiteralPath $expanded -Destination (Join-Path $Root 't7\english_expanded.vla')

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($f in Get-ChildItem -LiteralPath $Root -Recurse -File -Filter '*.vla') {
        $rel = $f.FullName.Substring($Root.Length).TrimStart('\')
        $lines.Add($rel + '|' + (Sha $f.FullName) + '|' + $f.LastWriteTimeUtc.Ticks)
    }
    [System.IO.File]::WriteAllLines((Join-Path $Root 'before.txt'), $lines.ToArray())
    Write-Output "Made $Root - now run the Excel steps, then -Check."
    exit 0
}

if (-not $Check) { throw 'Pass -Make or -Check.' }

$before = @{}
foreach ($line in [System.IO.File]::ReadAllLines((Join-Path $Root 'before.txt'))) {
    $parts = $line.Split('|')
    $before[$parts[0]] = @{ Hash = $parts[1]; Ticks = [long]$parts[2] }
}
function Unchanged([string]$rel) {
    $p = Join-Path $Root $rel
    if (-not (Test-Path -LiteralPath $p)) { return $false }
    return ((Sha $p) -eq $before[$rel].Hash)
}
function BytesAre([string]$rel, [string]$text) {
    $p = Join-Path $Root $rel
    if (-not (Test-Path -LiteralPath $p)) { return $false }
    return ((Hex ([System.IO.File]::ReadAllBytes($p))) -eq (Hex $utf8.GetBytes($text)))
}

Say 'Test 1: pound.vla is linted, UTF-8 intact, four trailing breaks' (BytesAre 't1\pound.vla' ($poundText + $fourBreaks)) 'bytes differ from the expected lint output'

$t2 = Join-Path $Root 't2\clean.vla'
Say 'Test 2: clean.vla is untouched (same bytes, same write time)' ((Unchanged 't2\clean.vla') -and ((Get-Item -LiteralPath $t2).LastWriteTimeUtc.Ticks -eq $before['t2\clean.vla'].Ticks)) 'the file was rewritten'

Say 'Test 3: the generated english_expanded.vla is untouched' (Unchanged 't3\english_expanded.vla') 'bytes changed'
Say 'Test 4: ansi.vla (invalid UTF-8) is untouched' (Unchanged 't4\ansi.vla') 'bytes changed'
Say 'Test 5: comment.vla (comment inside a form) is untouched' (Unchanged 't5\comment.vla') 'bytes changed'

Say 'Test 6: messy.vla was rewritten to house style' (BytesAre 't6\messy.vla' ('(stop)' + $fourBreaks)) 'bytes differ from the expected lint output'
$t6Others = @('t6\clean.vla', 't6\english_expanded.vla', 't6\bad.vla', 't6\comment.vla')
$t6Changed = @($t6Others | Where-Object { -not (Unchanged $_) })
Say 'Test 6: clean, generated, bad and comment files are untouched' ($t6Changed.Count -eq 0) ('changed: ' + ($t6Changed -join ', '))

Say 'Test 7: pound.vla and english_expanded.vla are both untouched' ((Unchanged 't7\pound.vla') -and (Unchanged 't7\english_expanded.vla')) 'a file changed'

$exp = Join-Path $Root 't8\exp.vla'
if (-not (Test-Path -LiteralPath $exp)) {
    Say 'Test 8: exported exp.vla' $false "$exp does not exist - was it saved there?"
} else {
    $b = [System.IO.File]::ReadAllBytes($exp)
    $strictOk = $true
    try { $null = $utf8Strict.GetString($b) } catch { $strictOk = $false }
    $head = [System.Text.Encoding]::ASCII.GetString($b, 0, [Math]::Min(11, $b.Length))
    $n = $b.Length
    $tailOk = ($n -ge 10) -and ((Hex $b[($n - 8)..($n - 1)]) -eq (Hex $utf8.GetBytes($fourBreaks))) -and -not ($b[$n - 10] -eq 0x0D -and $b[$n - 9] -eq 0x0A)
    Say 'Test 8: exp.vla is strict UTF-8' $strictOk 'not valid UTF-8'
    Say 'Test 8: exp.vla starts with the ; GENERATED stamp' ($head -eq '; GENERATED') ("starts with: $head")
    Say 'Test 8: exp.vla ends with exactly four line breaks, as the old writer left it' $tailOk 'wrong trailing line breaks'
}

$tmps = @(Get-ChildItem -LiteralPath $Root -Recurse -File -Filter '*.lint-tmp')
Say 'Every test: no .lint-tmp copy left behind' ($tmps.Count -eq 0) ('found: ' + (($tmps | ForEach-Object { $_.FullName }) -join ', '))
