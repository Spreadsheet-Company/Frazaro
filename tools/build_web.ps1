<#
build_web.ps1 - the web door: web/index.template.html, filled from a distro,
is web/index.html.

WHAT IT DOES: builds the wasm core (cargo build --release -p frazaro-core
--target wasm32-unknown-unknown; -NoBuild takes the one already built, as CI
does after its own build step), runs tools/check_core_imports.ps1 on it (a
page that could phone home is not written), reads the distro's manifest
(distros/english/distro.vla unless -Distro names another folder; KERNEL.2,
2026-10-07), then fills the placeholders of web/index.template.html:
{{WASM_BASE64}} with the module as base64; {{PRELUDE}} and {{ENGLISH}} with
the distro's prelude and base phrasebook, each inlined as text; every
{{BOOK:name}} with the distro's dialect of that name, the language picker's
alternatives, and the two sets must agree exactly; {{TITLE}} (twice: the
page's title and its heading), {{TAGLINE}} and {{OPENS_WITH}} with the
distro's chrome, HTML-escaped; and {{PALETTE:lavender}}, {{PALETTE:whisper}}
and {{PALETTE:deep}} with the distro's three shades - and writes
web/index.html: one file that runs from file://, fetching nothing. The
output is a build artifact (.gitignore), as the add-in is: the template and
the distro are the source. .github/workflows/pages.yml runs this on each
release tag and hosts the file on GitHub Pages.

THE MANIFEST is read one directive a line, the convention
tools/check_distro.ps1 holds every manifest to, so this needs no form
parser: a line's quoted strings are its arguments, and a path is relative
to the distro's folder with forward slashes (distros/english/README.md has
the shape; core/src/distro.rs is the reader the command-line door uses).

WHY A SCRIPT: the texts change with the corpus and the core, and a page
built by hand would drift from both; this is VlaBuildAddin's role for the
web door, run before a release and by the core CI job, which keeps the page
as an artifact beside the wasm.

House style (tools/*.ps1): PowerShell 5.1, host-free, no Excel, no COM, no
network; exit 0 with the output named, exit 1 with the reason.

Usage:  powershell -File tools\build_web.ps1
        powershell -File tools\build_web.ps1 -NoBuild
        powershell -File tools\build_web.ps1 -Distro distros\english
        powershell -File tools\build_web.ps1 -Out C:\somewhere\index.html
        powershell -File tools\build_web.ps1 -NoBuild -Template <file> -Out <file>   # another template: a proof, or a page of one's own
#>
param(
    [switch]$NoBuild,
    [string]$Out = '',
    [string]$Distro = '',
    [string]$Template = ''
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$wasm = Join-Path $root 'target/wasm32-unknown-unknown/release/frazaro_core.wasm'
$template = if ($Template -ne '') { $Template } else { Join-Path $root 'web/index.template.html' }
if ($Out -eq '') { $Out = Join-Path $root 'web/index.html' }
if ($Distro -eq '') { $Distro = 'distros/english' }
if (-not [System.IO.Path]::IsPathRooted($Distro)) { $Distro = Join-Path $root $Distro }
$manifestPath = Join-Path $Distro 'distro.vla'
if (-not (Test-Path $manifestPath)) { Write-Output ("FAIL: no distro at {0}: a distro is a folder holding distro.vla (distros/english/README.md)" -f $Distro); exit 1 }

# The quoted strings of one manifest line, in order, \" and \\ read as the
# language reads them.
function Get-DirectiveStrings([string]$line) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($m in [regex]::Matches($line, '"((?:[^"\\]|\\.)*)"')) { $out.Add(($m.Groups[1].Value -replace '\\(.)', '$1')) }
    return ,$out
}

# A manifest's path, relative to its folder with forward slashes, as a full
# path with its .. segments collapsed.
function Resolve-DistroPath([string]$folder, [string]$rel) {
    return [System.IO.Path]::GetFullPath((Join-Path $folder ($rel -replace '/', '\')))
}

# The manifest, one directive a line, into a table of what the page needs.
function Read-Distro([string]$path) {
    $d = @{ Name = ''; Title = $null; Tagline = $null; OpensWith = $null; Palette = @{}; Prelude = $null; Base = $null; Dialects = New-Object System.Collections.Generic.List[object] }
    $seenDistro = 0
    foreach ($raw in [System.IO.File]::ReadAllLines($path)) {
        $t = $raw.Trim()
        if ($t -eq '' -or $t.StartsWith(';')) { continue }
        $head = [regex]::Match($t, '^\(([a-z-]+)\b').Groups[1].Value
        $s = Get-DirectiveStrings $t
        switch ($head) {
            'distro' { $seenDistro++; if ($s.Count -lt 1) { throw "$path`: (distro ...) names no distro" }; $d.Name = $s[0] }
            'title' { if ($null -ne $d.Title) { throw "$path`: (title ...) is given twice" }; $d.Title = $s[0] }
            'tagline' { if ($null -ne $d.Tagline) { throw "$path`: (tagline ...) is given twice" }; $d.Tagline = $s[0] }
            'opens-with' { if ($null -ne $d.OpensWith) { throw "$path`: (opens-with ...) is given twice" }; $d.OpensWith = $s[0] }
            'palette' {
                foreach ($pm in [regex]::Matches($t, '\((lavender|whisper|deep) "(#[0-9A-Fa-f]{6})"\)')) { $d.Palette[$pm.Groups[1].Value] = $pm.Groups[2].Value }
                foreach ($shade in @('lavender', 'whisper', 'deep')) { if (-not $d.Palette.ContainsKey($shade)) { throw "$path`: (palette ...) names no $shade shade as # and six hex digits" } }
            }
            'prelude' { if ($null -ne $d.Prelude) { throw "$path`: (prelude ...) is given twice" }; $d.Prelude = $s[0] }
            'phrasebook' { if ($null -ne $d.Base) { throw "$path`: (phrasebook ...) is given twice" }; if ($s.Count -ne 2) { throw "$path`: (phrasebook ...) wants a name and a path" }; $d.Base = @{ Name = $s[0]; Path = $s[1] } }
            'dialect' { if ($s.Count -ne 2) { throw "$path`: (dialect ...) wants a name and a path" }; $d.Dialects.Add(@{ Name = $s[0]; Path = $s[1] }) }
            default { }   # overlay, library, examples, addin, readme: not the page's
        }
    }
    if ($seenDistro -ne 1) { throw "$path`: it must hold one (distro ""name"" ...) form" }
    foreach ($need in @('Title', 'Tagline', 'OpensWith', 'Prelude', 'Base')) { if ($null -eq $d[$need]) { throw "$path`: it names no $need" } }
    return $d
}

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

try { $man = Read-Distro $manifestPath } catch { Write-Output ("FAIL: {0}" -f $_.Exception.Message); exit 1 }
$preludePath = Resolve-DistroPath $Distro $man.Prelude
$englishPath = Resolve-DistroPath $Distro $man.Base.Path
foreach ($p in @($preludePath, $englishPath)) { if (-not (Test-Path $p)) { Write-Output ("FAIL: the distro names {0}, which is not there" -f $p); exit 1 } }

$html = [System.IO.File]::ReadAllText($template)
# Each placeholder exactly as many times as the template has a place for it:
# the title twice, the page's <title> and its heading.
$placeholders = [ordered]@{ '{{WASM_BASE64}}' = 1; '{{PRELUDE}}' = 1; '{{ENGLISH}}' = 1; '{{TITLE}}' = 2; '{{TAGLINE}}' = 1; '{{OPENS_WITH}}' = 1; '{{PALETTE:lavender}}' = 1; '{{PALETTE:whisper}}' = 1; '{{PALETTE:deep}}' = 1 }
foreach ($ph in $placeholders.Keys) {
    $n = ([regex]::Matches($html, [regex]::Escape($ph))).Count
    if ($n -ne $placeholders[$ph]) { Write-Output ("FAIL: the template has {0} {1} time(s), not {2}" -f $ph, $n, $placeholders[$ph]); exit 1 }
}
$prelude = [System.IO.File]::ReadAllText($preludePath)
$english = [System.IO.File]::ReadAllText($englishPath)
foreach ($pair in @(@($man.Prelude, $prelude), @($man.Base.Path, $english))) {
    if ($pair[1] -match '(?i)</script') { Write-Output ("FAIL: {0} contains '</script', which would end its text block in the page" -f $pair[0]); exit 1 }
}
$b64 = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($wasm))

# The dialects the language picker offers: each {{BOOK:name}} is the
# distro's dialect of that name, inlined as text, once; a dialect the
# template has no place for, or a place the distro has no dialect for, fails.
$books = @{}
foreach ($m in [regex]::Matches($html, '\{\{BOOK:([a-z0-9-]+)\}\}')) {
    $name = $m.Groups[1].Value
    if ($books.ContainsKey($name)) { Write-Output ("FAIL: the template has {0} more than once" -f $m.Value); exit 1 }
    $dialect = $man.Dialects | Where-Object { $_.Name -eq $name } | Select-Object -First 1
    if ($null -eq $dialect) { Write-Output ("FAIL: the template has {0}, and the distro {1} names no dialect {2}" -f $m.Value, $man.Name, $name); exit 1 }
    $bookPath = Resolve-DistroPath $Distro $dialect.Path
    if (-not (Test-Path $bookPath)) { Write-Output ("FAIL: the distro names {0} for {1}, which is not there" -f $bookPath, $m.Value); exit 1 }
    $bookText = [System.IO.File]::ReadAllText($bookPath)
    if ($bookText -match '(?i)</script') { Write-Output ("FAIL: {0} contains '</script', which would end its text block in the page" -f $dialect.Path); exit 1 }
    $books[$name] = $bookText
}
foreach ($dialect in $man.Dialects) {
    if (-not $books.ContainsKey($dialect.Name)) { Write-Output ("FAIL: the distro {0} names the dialect {1}, and the template has no {{{{BOOK:{1}}}}} for it" -f $man.Name, $dialect.Name); exit 1 }
}
$esc = { param([string]$s) [System.Net.WebUtility]::HtmlEncode($s) }
$html = $html.Replace('{{WASM_BASE64}}', $b64).Replace('{{PRELUDE}}', $prelude).Replace('{{ENGLISH}}', $english)
$html = $html.Replace('{{TITLE}}', (& $esc $man.Title)).Replace('{{TAGLINE}}', (& $esc $man.Tagline)).Replace('{{OPENS_WITH}}', (& $esc $man.OpensWith))
foreach ($shade in @('lavender', 'whisper', 'deep')) { $html = $html.Replace('{{PALETTE:' + $shade + '}}', $man.Palette[$shade]) }
foreach ($name in $books.Keys) { $html = $html.Replace('{{BOOK:' + $name + '}}', $books[$name]) }
foreach ($ph in @('{{WASM_BASE64}}', '{{PRELUDE}}', '{{ENGLISH}}', '{{TITLE}}', '{{TAGLINE}}', '{{OPENS_WITH}}', '{{PALETTE:', '{{BOOK:')) {
    # The corpus texts may hold doubled braces of their own; only the names count.
    if ($html.Contains($ph)) { Write-Output ("FAIL: {0} is left in the page" -f $ph); exit 1 }
}
$outDir = Split-Path -Parent $Out
if ($outDir -ne '' -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
[System.IO.File]::WriteAllText($Out, $html, (New-Object System.Text.UTF8Encoding($false)))
Write-Output ("OK: wrote {0} ({1:N0} characters) from the {6} distro: the core {2:N0} bytes, {7} {3:N0}, the prelude {4:N0}, {5} dialect(s)" -f $Out, $html.Length, (Get-Item $wasm).Length, $english.Length, $prelude.Length, $books.Count, $man.Name, $man.Base.Name)
exit 0
