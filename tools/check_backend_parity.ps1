<#
check_backend_parity.ps1 - AS.8's own coverage report.

"Which core forms have a pin under one backend and not the other" - SD-5's
enforcement, IN.3's generalization from one corpus file (instructions.txt) to
the whole grammar VLA_HeadTable.bas catalogs. Mirrors check_id_registry.ps1
in shape (a grep-style static scan, runnable without a live Excel host,
reporting rather than fixing) because AS.8 is the same kind of instrument
one level over: that script asks "does this ID collide"; this one asks
"does this form have proof under both backends."

WHAT COUNTS AS A "PIN": a logical VBA statement (line, or several joined
by trailing " _" continuations - VLA_Tests*.bas's real shape) that calls
a known emitter-path helper (AssertVla/TryTranspile/VlaTranspile - each
takes real VLA source text as an argument) or a known interpreter-path
helper (VlaInterpret/VlaEvalExpression, same contract) AND whose text
contains the form's own head symbol immediately after an open paren
("(if ", "(if)"), the actual S-expression head-position shape every real
use of a form takes. This is a heuristic, not a parser - same tolerance
this project's own DynamicGet/DynamicCall/DynamicSet already state for
themselves ("a heuristic, not a type system") - so a false negative is
possible if a form's own symbol is buried inside a helper macro this
script does not expand; a false positive is possible if a comment or an
unrelated string happens to contain the same head-position text. Good
enough for a coverage REPORT, the same way check_id_registry.ps1's own
token regex is good enough for a collision report without a real VLA
tokenizer.

FORMS SCANNED: read directly from VLA_HeadTable.bas's own AddRow calls
(the IN.1 catalog, 41 rows today) - not a separately hand-maintained
list, so this script and the catalog can never silently drift apart.
Export-only rows (raw/deflambda/lambda - IN.5's own adjudicated set) are
reported separately: a missing interpreter pin there is BY DESIGN, not a
gap, so folding them into the same "gap" list would misreport a settled
decision as an open one.

Usage:  pwsh -File tools/check_backend_parity.ps1
#>

$ErrorActionPreference = 'Stop'

$repoRoot   = Split-Path -Parent $PSScriptRoot
$srcDir     = Join-Path $repoRoot 'src'
$headTable  = Join-Path $srcDir 'VLA_HeadTable.bas'
$testFiles  = @('VLA_Tests.bas', 'VLA_Tests_Grammar.bas', 'VLA_Tests_Host.bas') |
              ForEach-Object { Join-Path $srcDir $_ } | Where-Object { Test-Path $_ }

# --- Pass 1: the canonical form list, read from the catalog, not hand-kept ----
# AddRow rows, "<symbol>", "<aliases>", "<arity>", "<interp>", "<vba>", "<formula>", <True|False>
$addRowPattern = 'AddRow\s+rows,\s*"((?:[^"])*)",\s*"((?:[^"])*)",\s*"((?:[^"])*)",\s*"((?:[^"])*)",\s*"((?:[^"])*)",\s*"((?:[^"])*)",\s*(True|False)'
$forms = New-Object System.Collections.Generic.List[object]
foreach ($line in Get-Content -LiteralPath $headTable) {
    $m = [regex]::Match($line, $addRowPattern)
    if ($m.Success) {
        $forms.Add([pscustomobject]@{
            Symbol     = $m.Groups[1].Value
            ExportOnly = ($m.Groups[7].Value -eq 'True')
        })
    }
}
if ($forms.Count -eq 0) {
    Write-Error "No rows parsed from $headTable - AddRow's own line shape may have changed; update `$addRowPattern."
}

# --- Pass 2: join VBA line-continuations into logical statements -------------
function Get-LogicalStatements([string]$path) {
    $lines = Get-Content -LiteralPath $path
    $stmts = New-Object System.Collections.Generic.List[string]
    $buf = ''
    foreach ($line in $lines) {
        $buf += $line
        if ($line -match '_\s*$') {
            $buf += "`n"
            continue
        }
        $stmts.Add($buf)
        $buf = ''
    }
    if ($buf.Length -gt 0) { $stmts.Add($buf) }
    return $stmts
}

$emitterCallNames = @('AssertVla\(', 'TryTranspile\(', 'VlaTranspile\(')
$interpCallNames  = @('VlaInterpret\(', 'VlaEvalExpression\(')
$emitterRe = ($emitterCallNames -join '|')
$interpRe  = ($interpCallNames  -join '|')

$emitterText = New-Object System.Text.StringBuilder
$interpText  = New-Object System.Text.StringBuilder
foreach ($f in $testFiles) {
    foreach ($stmt in (Get-LogicalStatements $f)) {
        if ($stmt -match $emitterRe) { [void]$emitterText.AppendLine($stmt) }
        if ($stmt -match $interpRe)  { [void]$interpText.AppendLine($stmt) }
    }
}
$emitterBlob = $emitterText.ToString()
$interpBlob  = $interpText.ToString()

# --- Pass 3: for each form, does its head-position text appear in each blob --
function Test-HeadPosition([string]$blob, [string]$symbol) {
    $esc = [regex]::Escape($symbol)
    return [regex]::IsMatch($blob, '\(' + $esc + '[\s\)]')
}

$rows = foreach ($f in $forms) {
    $inEmitter = Test-HeadPosition $emitterBlob $f.Symbol
    $inInterp  = Test-HeadPosition $interpBlob  $f.Symbol
    [pscustomobject]@{
        Symbol     = $f.Symbol
        ExportOnly = $f.ExportOnly
        Emitter    = $inEmitter
        Interp     = $inInterp
    }
}

Write-Output '=== BACKEND-PARITY PIN COVERAGE (heuristic, head-position text scan) ==='
Write-Output "Forms scanned: $($rows.Count) (from VLA_HeadTable.bas)"
Write-Output ''

$gapInterp = $rows | Where-Object { $_.Emitter -and -not $_.Interp -and -not $_.ExportOnly }
Write-Output "--- Emitter-only (no interpreter pin found), $($gapInterp.Count) form(s) ---"
$gapInterp | ForEach-Object { Write-Output "  $($_.Symbol)" }

Write-Output ''
$gapEmitter = $rows | Where-Object { $_.Interp -and -not $_.Emitter }
Write-Output "--- Interpreter-only (no emitter pin found), $($gapEmitter.Count) form(s) ---"
$gapEmitter | ForEach-Object { Write-Output "  $($_.Symbol)" }

Write-Output ''
$exportOnlyGap = $rows | Where-Object { $_.ExportOnly -and -not $_.Interp }
Write-Output "--- Export-only, no interpreter pin (expected by IN.5's own design, not a gap), $($exportOnlyGap.Count) form(s) ---"
$exportOnlyGap | ForEach-Object { Write-Output "  $($_.Symbol)" }

Write-Output ''
$neither = $rows | Where-Object { -not $_.Emitter -and -not $_.Interp -and -not $_.ExportOnly }
Write-Output "--- Neither backend: no pin found under either scan, $($neither.Count) form(s) - AS.1/AS.2's own territory, not this script's finding to explain ---"
$neither | ForEach-Object { Write-Output "  $($_.Symbol)" }

Write-Output ''
$both = $rows | Where-Object { $_.Emitter -and $_.Interp }
Write-Output "=== SUMMARY: $($both.Count)/$($rows.Count) forms have a pin under BOTH backends ==="
