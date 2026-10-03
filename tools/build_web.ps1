<#
build_web.ps1 - the web door: web/index.template.html, filled, is web/index.html.

WHAT IT DOES: builds the wasm core (cargo build --release -p frazaro-core
--target wasm32-unknown-unknown; -NoBuild takes the one already built, as CI
does after its own build step), runs tools/check_core_imports.ps1 on it (a
page that could phone home is not written), then fills the three placeholders
of web/index.template.html - {{WASM_BASE64}} with the module as base64,
{{PRELUDE}} with scripts/prelude.vla and {{ENGLISH}} with
scripts/polyglotta/english.vla, each inlined as text, and every
{{BOOK:name}} with scripts/polyglotta/<name>.vla, the dialects the page's
language picker offers - and writes web/index.html: one file that runs from
file://, fetching nothing. The output
is a build artifact (.gitignore), as the add-in is: the template is the
source.

WHY A SCRIPT: the three texts change with the corpus and the core, and a
page built by hand would drift from both; this is VlaBuildAddin's role for
the web door, run before a release and by the core CI job, which keeps the
page as an artifact beside the wasm.

House style (tools/*.ps1): PowerShell 5.1, host-free, no Excel, no COM, no
network; exit 0 with the output named, exit 1 with the reason.

Usage:  powershell -File tools\build_web.ps1
        powershell -File tools\build_web.ps1 -NoBuild
        powershell -File tools\build_web.ps1 -Out C:\somewhere\index.html
#>
param(
    [switch]$NoBuild,
    [string]$Out = ''
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$wasm = Join-Path $root 'target/wasm32-unknown-unknown/release/frazaro_core.wasm'
$template = Join-Path $root 'web/index.template.html'
$preludePath = Join-Path $root 'scripts/prelude.vla'
$englishPath = Join-Path $root 'scripts/polyglotta/english.vla'
if ($Out -eq '') { $Out = Join-Path $root 'web/index.html' }

if (-not $NoBuild) {
    $cargo = 'cargo'
    if (-not (Get-Command cargo -ErrorAction SilentlyContinue)) {
        $cargo = Join-Path $env:USERPROFILE '.cargo\bin\cargo.exe'
    }
    & $cargo build --release -p frazaro-core --target wasm32-unknown-unknown
    if ($LASTEXITCODE -ne 0) { Write-Output 'FAIL: the wasm core did not build'; exit 1 }
}
if (-not (Test-Path $wasm)) { Write-Output ("FAIL: no wasm core at {0} (build it, or drop -NoBuild)" -f $wasm); exit 1 }

# The import section must be empty before the module goes into a page.
$shell = (Get-Process -Id $PID).Path
& $shell -NoProfile -File (Join-Path $PSScriptRoot 'check_core_imports.ps1') -Path $wasm
if ($LASTEXITCODE -ne 0) { Write-Output 'FAIL: the wasm core imports something; the page is not written'; exit 1 }

$html = [System.IO.File]::ReadAllText($template)
foreach ($ph in @('{{WASM_BASE64}}', '{{PRELUDE}}', '{{ENGLISH}}')) {
    $first = $html.IndexOf($ph)
    if ($first -lt 0) { Write-Output ("FAIL: the template has no {0}" -f $ph); exit 1 }
    if ($html.IndexOf($ph, $first + 1) -ge 0) { Write-Output ("FAIL: the template has {0} more than once" -f $ph); exit 1 }
}
$prelude = [System.IO.File]::ReadAllText($preludePath)
$english = [System.IO.File]::ReadAllText($englishPath)
foreach ($pair in @(@('prelude.vla', $prelude), @('english.vla', $english))) {
    if ($pair[1] -match '(?i)</script') { Write-Output ("FAIL: {0} contains '</script', which would end its text block in the page" -f $pair[0]); exit 1 }
}
$b64 = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($wasm))

# The dialects the language picker offers: each {{BOOK:name}} is
# scripts/polyglotta/<name>.vla, inlined as text, once.
$books = @{}
foreach ($m in [regex]::Matches($html, '\{\{BOOK:([a-z]+)\}\}')) {
    $name = $m.Groups[1].Value
    if ($books.ContainsKey($name)) { Write-Output ("FAIL: the template has {0} more than once" -f $m.Value); exit 1 }
    $bookPath = Join-Path $root ("scripts/polyglotta/" + $name + ".vla")
    if (-not (Test-Path $bookPath)) { Write-Output ("FAIL: no phrasebook for {0} at {1}" -f $m.Value, $bookPath); exit 1 }
    $bookText = [System.IO.File]::ReadAllText($bookPath)
    if ($bookText -match '(?i)</script') { Write-Output ("FAIL: {0}.vla contains '</script', which would end its text block in the page" -f $name); exit 1 }
    $books[$name] = $bookText
}
$html = $html.Replace('{{WASM_BASE64}}', $b64).Replace('{{PRELUDE}}', $prelude).Replace('{{ENGLISH}}', $english)
foreach ($name in $books.Keys) { $html = $html.Replace('{{BOOK:' + $name + '}}', $books[$name]) }
foreach ($ph in @('{{WASM_BASE64}}', '{{PRELUDE}}', '{{ENGLISH}}', '{{BOOK:')) {
    # The corpus texts may hold doubled braces of their own; only the names count.
    if ($html.Contains($ph)) { Write-Output ("FAIL: {0} is left in the page" -f $ph); exit 1 }
}
$outDir = Split-Path -Parent $Out
if ($outDir -ne '' -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
[System.IO.File]::WriteAllText($Out, $html, (New-Object System.Text.UTF8Encoding($false)))
Write-Output ("OK: wrote {0} ({1:N0} characters; the core {2:N0} bytes, english.vla {3:N0}, prelude.vla {4:N0}, {5} dialect(s))" -f $Out, $html.Length, (Get-Item $wasm).Length, $english.Length, $prelude.Length, $books.Count)
exit 0
