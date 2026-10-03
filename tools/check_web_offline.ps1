<#
check_web_offline.ps1 - the web page loads nothing, links nowhere, and is
filled from exactly three placeholders.

WHY: the web door (PORT.6, slice 6h; docs/HORIZON.md section 11.4) is one
HTML file that runs from disk, and its whole promise is the add-in's SD-13
in a browser: it reads nothing from the network and writes nothing anywhere.
The core it carries cannot phone home, since its import section is empty
(check_core_imports.ps1 reads that off the artifact); this check holds the
page around the core to the same doctrine, on every push, so that a script
tag, a stylesheet, a font, an image or a fetch cannot quietly make the page
need a network the day someone adds one for convenience.

WHAT IT HOLDS, on web/index.template.html (the source; web/index.html is a
build artifact): no external reference of any kind - script, link, img,
iframe, form, anchor, @import, url(), http(s)://, fetch, XMLHttpRequest,
WebSocket, EventSource, sendBeacon, dynamic import, importScripts; a
charset declaration; and each of the three placeholders tools/build_web.ps1
fills ({{WASM_BASE64}}, {{PRELUDE}}, {{ENGLISH}}) present exactly once, so a
built page is whole. When web/index.html exists beside the template, its
markup is held to the same list of tags (the base64 and the corpus texts
cannot spell a tag).

-Control proves the check on scratch copies of the template: the real file
passes; one with a script tag's src appended, one with a fetch call, and one
with a placeholder removed must each fail.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; exit 0 clean, exit 1 with every problem named.

Usage:  powershell -File tools\check_web_offline.ps1
        powershell -File tools\check_web_offline.ps1 -Control
#>
param(
    [switch]$Control
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$templatePath = Join-Path $repoRoot 'web/index.template.html'
$builtPath = Join-Path $repoRoot 'web/index.html'

# Markup that reaches out: held on the template and on a built page alike.
$tagTokens = @('<script src', '<script type="module" src', '<link ', '<img ', '<iframe', '<form', '<a href', '<video', '<audio', '<source', '<object', '<embed', '<meta http-equiv="refresh"')
# Code and style that reach out: held on the template (the built page carries
# the corpus texts, which may mention a URL in a comment).
$codeTokens = @('fetch(', 'XMLHttpRequest', 'WebSocket', 'EventSource', 'sendBeacon', 'import(', 'importScripts', '@import', 'url(', 'http://', 'https://')
$placeholders = @('{{WASM_BASE64}}', '{{PRELUDE}}', '{{ENGLISH}}')

function Test-Template([string]$path) {
    $problems = @()
    $text = [System.IO.File]::ReadAllText($path)
    $lower = $text.ToLowerInvariant()
    foreach ($t in ($tagTokens + $codeTokens)) {
        if ($lower.Contains($t.ToLowerInvariant())) { $problems += ("the page reaches out: '{0}'" -f $t) }
    }
    if (-not $lower.Contains('<meta charset="utf-8">')) { $problems += 'no <meta charset="utf-8">' }
    foreach ($ph in $placeholders) {
        $first = $text.IndexOf($ph)
        if ($first -lt 0) { $problems += ("placeholder missing: {0}" -f $ph); continue }
        if ($text.IndexOf($ph, $first + 1) -ge 0) { $problems += ("placeholder repeated: {0}" -f $ph) }
    }
    return $problems
}

function Test-Built([string]$path) {
    $problems = @()
    $lower = [System.IO.File]::ReadAllText($path).ToLowerInvariant()
    foreach ($t in $tagTokens) {
        if ($lower.Contains($t.ToLowerInvariant())) { $problems += ("the built page reaches out: '{0}'" -f $t) }
    }
    foreach ($ph in $placeholders) {
        # The corpus texts may hold doubled braces of their own; only the three names count.
        if ($lower.Contains($ph.ToLowerInvariant())) { $problems += ("the built page has {0} left in it" -f $ph) }
    }
    return $problems
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_weboff_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $text = [System.IO.File]::ReadAllText($templatePath)
        $real = @(Test-Template $templatePath)
        $a = Join-Path $tmp 'a.html'; [System.IO.File]::WriteAllText($a, ($text + "`n<script src=""x.js""></script>`n"), $utf8)
        $b = Join-Path $tmp 'b.html'; [System.IO.File]::WriteAllText($b, ($text + "`n<script>fetch('x')</script>`n"), $utf8)
        $c = Join-Path $tmp 'c.html'; [System.IO.File]::WriteAllText($c, $text.Replace('{{PRELUDE}}', ''), $utf8)
        $ra = @(Test-Template $a); $rb = @(Test-Template $b); $rc = @(Test-Template $c)
        Write-Output ("control: the real template has {0} problem(s); the script src {1}, the fetch {2}, the missing placeholder {3}" -f $real.Count, $ra.Count, $rb.Count, $rc.Count)
        if ($real.Count -eq 0 -and $ra.Count -ge 1 -and $rb.Count -ge 1 -and $rc.Count -ge 1) {
            Write-Output 'OK: the check passes the template and fails each mutant'
            exit 0
        }
        Write-Output 'FAIL: the control did not behave'
        foreach ($p in $real) { Write-Output ('  real: ' + $p) }
        exit 1
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}

if (-not (Test-Path $templatePath)) { Write-Output ("FAIL: no template at {0}" -f $templatePath); exit 1 }
$problems = @(Test-Template $templatePath)
$builtNote = 'no built page beside it (tools/build_web.ps1 writes one)'
if (Test-Path $builtPath) {
    $problems += @(Test-Built $builtPath)
    $builtNote = ("the built page ({0:N0} characters) holds too" -f (Get-Item $builtPath).Length)
}
if ($problems.Count -eq 0) {
    Write-Output ("OK: web/index.template.html loads nothing and links nowhere, with its three placeholders; {0}" -f $builtNote)
    exit 0
}
Write-Output ("FAIL: the web page has {0} problem(s):" -f $problems.Count)
foreach ($p in $problems) { Write-Output ('  ' + $p) }
exit 1
