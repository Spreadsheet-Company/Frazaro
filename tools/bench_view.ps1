<#
bench_view.ps1 - the renderer benchmark page of KERNEL.3 (the study before
the viewport, SD-29): tools/bench_view.template.html, filled with two real
view records of two synthetic programs, is tools/bench_view.html.

WHAT IT DOES: writes two synthetic programs to scratch files under $env:TEMP
with LF line endings. The sparse one is the program web/CALLOSUM.md section 10
generates with awk: n lines, an odd line i 'Put i into cell Ai.' and an even
line i 'Put formula "=A(i-1)*2" into cell Bi.', one cell a row. The dense one
fills every cell of R rows by C columns: in column c of row r, an odd column
'Put <r*c> into cell <col>r.' and an even column 'Put formula "=<col-1>r*2"
into cell <col>r.', so that a full-screen window is full of text. It runs the
door's view over each program's Output sheet with no window, so each record
holds the sheet's whole extent, and times each run with a Stopwatch around the
call (the time holds PowerShell's reading of the lines too, some 40 ms of a
second); holds each record to the extent it must have, (extent "Output"
"A1:B<n>") and (extent "Output" "A1:<col C><R>"), and refuses one holding
'</script', which would end the page's text block, as build_web.ps1 refuses a
phrasebook; holds the template to one {{VIEW_RECORD}}, one
{{VIEW_RECORD_DENSE}} and one {{DOOR_LINE}} and fills them, the door line read
off 'frazaro --version' and the two Stopwatches, such as
  frazaro 0.8.0 view: the sparse program, 10000 lines, 20006 record lines in 893 ms; the dense program, 600 rows x 52 columns, 31200 lines, 62406 record lines in 2710 ms (native debug door)
writes the page UTF-8 without a BOM, and prints what it wrote, each record's
line count and the door's times.

WHY: KERNEL.5 picks the viewport's renderer, canvas or virtualized DOM, by a
number and not by taste (SD-29), and the page measures both over the records a
program's build projects, in the owner's browser, where the viewport will run:
the sparse record at the protocol's two windows, and the dense record at full
screen down a ladder of cell sizes, so that the number also says how each
renderer scales with the text it draws (the amendment of 2026-10-08, after the
first fullscreen run showed a two-column record and a vsync-quantized clock
could not tell the two apart). The records are the door's own output, never
typed in by hand, so the page draws from what the viewport will draw from. The
filled page is a build artifact (.gitignore), as web/index.html is: the
template and this script are the source.

THE DOOR: -Impl names one; else target/debug/frazaro.exe, else
target/release/frazaro.exe; neither there fails naming both. The corpus
files are passed as the checks pass them (--prelude scripts/prelude.vla,
--phrasebook scripts/polyglotta/english.vla), so the records are of the
tree's own corpus and not of a door's built-in copy.

House style (tools/*.ps1): PowerShell 5.1, host-free, no Excel, no COM, no
network; exit 0 with the output named, exit 1 with the reason.

Nothing but this script's own assertions holds the template
(check_web_offline.ps1 scans web/ only): the placeholder counts, the records'
extents and the '</script' refusals below are the whole of it. A relative
-Out resolves against the repository root, as a relative -Impl does.

Smoke: a headless run of the built page in smoke mode (?smoke=1), judged by
the content of its <pre id="smoke"> block alone, since the page's own script
holds the words 'smoke ok' too, so never grep the whole dump; the numbers of
a headless run are virtual-clock artifacts, and only their presence and the
table's shape count. The command that works, its flags on one line:
  "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --headless=new --run-all-compositor-stages-before-draw --disable-frame-rate-limit --disable-gpu-vsync --virtual-time-budget=400000 --disable-gpu --no-first-run --no-default-browser-check --user-data-dir=<a scratch folder> --dump-dom "file:///C:/<repo>/tools/bench_view.html?smoke=1"
run through Start-Process with -RedirectStandardOutput for the dump,
-RedirectStandardError, -Wait, -PassThru and -NoNewWindow. Without
--disable-frame-rate-limit and --disable-gpu-vsync the headless run stalls at
its first animation frame, and the watchdog's line is all the block holds.

Usage:  powershell -File tools\bench_view.ps1
        powershell -File tools\bench_view.ps1 -Lines 1000 -DenseRows 300 -DenseCols 26
        powershell -File tools\bench_view.ps1 -Impl target\release\frazaro.exe
        powershell -File tools\bench_view.ps1 -Out C:\somewhere\bench_view.html
#>
param(
    [int]$Lines = 10000,
    [int]$DenseRows = 600,
    [int]$DenseCols = 52,
    [string]$Impl = '',
    [string]$Out = 'tools/bench_view.html'
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$template = Join-Path $PSScriptRoot 'bench_view.template.html'
$utf8 = New-Object System.Text.UTF8Encoding($false)

if ($Lines -lt 2) { Write-Output 'FAIL: -Lines must be at least 2: an odd line puts a value, an even line a formula'; exit 1 }
if ($DenseRows -lt 2 -or $DenseCols -lt 2) { Write-Output 'FAIL: -DenseRows and -DenseCols must be at least 2'; exit 1 }

# The door: -Impl, else the debug door, else the release door.
if ($Impl -eq '') {
    $debugDoor = Join-Path $root 'target\debug\frazaro.exe'
    $releaseDoor = Join-Path $root 'target\release\frazaro.exe'
    if (Test-Path $debugDoor) { $Impl = $debugDoor }
    elseif (Test-Path $releaseDoor) { $Impl = $releaseDoor }
    else { Write-Output ("FAIL: no door at {0} or at {1}: build one (cargo build -p frazaro, or --release) or name one with -Impl" -f $debugDoor, $releaseDoor); exit 1 }
} else {
    if (-not [System.IO.Path]::IsPathRooted($Impl)) { $Impl = Join-Path $root $Impl }
    if (-not (Test-Path $Impl)) { Write-Output ("FAIL: no door at {0}" -f $Impl); exit 1 }
}
if (-not [System.IO.Path]::IsPathRooted($Out)) { $Out = Join-Path $root $Out }
if (-not (Test-Path $template)) { Write-Output ("FAIL: no template at {0}" -f $template); exit 1 }
$prelude = Join-Path $root 'scripts\prelude.vla'
$phrasebook = Join-Path $root 'scripts\polyglotta\english.vla'
foreach ($p in @($prelude, $phrasebook)) { if (-not (Test-Path $p)) { Write-Output ("FAIL: no corpus file at {0}" -f $p); exit 1 } }

# The template: each placeholder exactly once.
$placeholders = @('{{VIEW_RECORD}}', '{{VIEW_RECORD_DENSE}}', '{{DOOR_LINE}}')
$html = [System.IO.File]::ReadAllText($template)
foreach ($ph in $placeholders) {
    $n = ([regex]::Matches($html, [regex]::Escape($ph))).Count
    if ($n -ne 1) { Write-Output ("FAIL: the template has {0} {1} time(s), not once" -f $ph, $n); exit 1 }
}

# A column's letters, 1 = A, 26 = Z, 27 = AA.
function Get-ColumnName([int]$c) {
    $s = ''
    while ($c -gt 0) { $r = ($c - 1) % 26; $s = [string][char](65 + $r) + $s; $c = [int](($c - 1 - $r) / 26) }
    return $s
}

# The sparse program, exactly the awk line's text, LF line endings.
$sb = New-Object System.Text.StringBuilder
for ($i = 1; $i -le $Lines; $i++) {
    if (($i % 2) -eq 1) { [void]$sb.Append('Put ').Append($i).Append(' into cell A').Append($i).Append(".`n") }
    else { [void]$sb.Append('Put formula "=A').Append($i - 1).Append('*2" into cell B').Append($i).Append(".`n") }
}
$program = Join-Path $env:TEMP ("frazaro_bench_{0}.txt" -f $Lines)
[System.IO.File]::WriteAllText($program, $sb.ToString(), $utf8)

# The dense program: every cell of R rows by C columns, values in the odd columns and formulas in the even.
$cols = @(1..$DenseCols | ForEach-Object { Get-ColumnName $_ })
$sd = New-Object System.Text.StringBuilder
for ($r = 1; $r -le $DenseRows; $r++) {
    for ($c = 1; $c -le $DenseCols; $c++) {
        if (($c % 2) -eq 1) { [void]$sd.Append('Put ').Append($r * $c).Append(' into cell ').Append($cols[$c - 1]).Append($r).Append(".`n") }
        else { [void]$sd.Append('Put formula "=').Append($cols[$c - 2]).Append($r).Append('*2" into cell ').Append($cols[$c - 1]).Append($r).Append(".`n") }
    }
}
$denseProgram = Join-Path $env:TEMP ("frazaro_bench_dense_{0}x{1}.txt" -f $DenseRows, $DenseCols)
[System.IO.File]::WriteAllText($denseProgram, $sd.ToString(), $utf8)

# The door's version: the second word of 'frazaro 0.8.0 (core abi 1)'.
$global:LASTEXITCODE = 0
$versionOut = & $Impl --version
if ($LASTEXITCODE -ne 0) { Write-Output ("FAIL: {0} --version exited {1}" -f $Impl, $LASTEXITCODE); exit 1 }
$versionLine = [string](@($versionOut)[0])
$vm = [regex]::Match($versionLine, '^frazaro (\S+)')
if (-not $vm.Success) { Write-Output ("FAIL: {0} --version printed '{1}', not 'frazaro <version> ...'" -f $Impl, $versionLine); exit 1 }
$version = $vm.Groups[1].Value

# The view of one program, timed around the call: the record's lines and the milliseconds, or a failure.
function Invoke-View([string]$path, [string]$expectedExtent) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $global:LASTEXITCODE = 0
    $printed = & $Impl view $path --sheet Output --prelude $prelude --phrasebook $phrasebook
    $code = $LASTEXITCODE
    $sw.Stop()
    if ($code -ne 0) { Write-Output ("FAIL: {0} view exited {1} on {2}" -f $Impl, $code, $path); exit 1 }
    $recordLines = @(@($printed) | ForEach-Object { [string]$_ })
    if ($recordLines.Count -eq 0) { Write-Output ("FAIL: {0} view printed nothing for {1}" -f $Impl, $path); exit 1 }
    if (-not ($recordLines -contains $expectedExtent)) { Write-Output ("FAIL: the record of {0} holds no line {1}" -f $path, $expectedExtent); exit 1 }
    $text = ($recordLines -join "`n") + "`n"
    if ($text -match '(?i)</script') { Write-Output ("FAIL: the record of {0} contains '</script', which would end its text block in the page" -f $path); exit 1 }
    return @{ Text = $text; Count = $recordLines.Count; Ms = [int][math]::Round($sw.Elapsed.TotalMilliseconds) }
}
# (Not $lines: a PowerShell name is read without case, and $Lines is the parameter.)
$sparseView = Invoke-View $program ('(extent "Output" "A1:B{0}")' -f $Lines)
$denseView = Invoke-View $denseProgram ('(extent "Output" "A1:{0}{1}")' -f $cols[$DenseCols - 1], $DenseRows)

# The door line, and the page.
$kind = if ($Impl -match '[\\/]release[\\/]') { 'native release door' } elseif ($Impl -match '[\\/]debug[\\/]') { 'native debug door' } else { 'native door' }
$doorLine = 'frazaro {0} view: the sparse program, {1} lines, {2} record lines in {3} ms; the dense program, {4} rows x {5} columns, {6} lines, {7} record lines in {8} ms ({9})' -f `
    $version, $Lines, $sparseView.Count, $sparseView.Ms, $DenseRows, $DenseCols, ($DenseRows * $DenseCols), $denseView.Count, $denseView.Ms, $kind
if ($doorLine -match '(?i)</script') { Write-Output "FAIL: the door line contains '</script'"; exit 1 }
$page = $html.Replace('{{VIEW_RECORD}}', $sparseView.Text).Replace('{{VIEW_RECORD_DENSE}}', $denseView.Text).Replace('{{DOOR_LINE}}', $doorLine)
foreach ($ph in $placeholders) {
    if ($page.Contains($ph)) { Write-Output ("FAIL: {0} is left in the page" -f $ph); exit 1 }
}
$outDir = Split-Path -Parent $Out
if ($outDir -ne '' -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
[System.IO.File]::WriteAllText($Out, $page, $utf8)
Write-Output ("OK: wrote {0} ({1:N0} bytes): the sparse record {2:N0} lines, extent A1:B{3}, in {4} ms; the dense record {5:N0} lines, extent A1:{6}{7}, in {8} ms ({9}); the programs at {10} and {11}" -f `
    $Out, (Get-Item $Out).Length, $sparseView.Count, $Lines, $sparseView.Ms, $denseView.Count, $cols[$DenseCols - 1], $DenseRows, $denseView.Ms, $Impl, $program, $denseProgram)
exit 0
