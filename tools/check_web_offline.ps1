<#
check_web_offline.ps1 - the web page loads nothing, links to one place by the
person's click, and is filled from the english distro through nine
placeholders, each in its place.

WHY: the web door (PORT.6, slice 6h; docs/HORIZON.md section 11.4) is one
HTML file that runs from disk, and its whole promise is the add-in's SD-13
in a browser: it reads nothing from the network and writes nothing anywhere.
The core it carries cannot phone home, since its import section is empty
(check_core_imports.ps1 reads that off the artifact); this check holds the
page around the core to the same doctrine, on every push, so that a script
tag, a stylesheet, a font, an image or a fetch cannot quietly make the page
need a network the day someone adds one for convenience. Since 2026-10-05
the page is hosted on GitHub Pages (.github/workflows/pages.yml) and carries
one way back to the owner's site: a bare link at the very top, which the
browser follows only when the person clicks it, and nothing else (the owner
chose the bare address over a logo, so the page holds no image at all). That
one anchor is pinned here, whole, exactly as spelled and exactly once, and
the scan runs on the text with it removed, so a second link fails as any
other reaching out does.

WHAT IT HOLDS, on web/index.template.html (the source; web/index.html is a
build artifact): no external reference of any kind beyond the one allowed
anchor - script, link, img, iframe, form, anchor, @import, url(), http(s)://,
fetch, XMLHttpRequest, WebSocket, EventSource, sendBeacon, dynamic import,
importScripts; the allowed anchor present once; a charset declaration; each
of the nine placeholders tools/build_web.ps1 fills from the distro
(KERNEL.2, 2026-10-07) present exactly as many times as the template has a
place for it - {{WASM_BASE64}}, {{PRELUDE}}, {{ENGLISH}}, {{TAGLINE}},
{{OPENS_WITH}}, {{PALETTE:lavender}}, {{PALETTE:whisper}} and
{{PALETTE:deep}} once each, {{TITLE}} twice, the page's title and its
heading; and each {{BOOK:name}} of the language picker once, a dialect of
the english distro (distros/english/distro.vla), with no dialect of that
distro left without its place, so a built page is whole and the picker and
the distro cannot drift apart. When web/index.html exists beside the
template, its markup is held to the same list of tags (the base64 and the
corpus texts cannot spell a tag) and no placeholder is left in it.

-Control proves the check on scratch copies of the template: the real file
passes; one with a script tag's src appended, one with a fetch call, one
with a placeholder removed, one with a second anchor appended, one with an
image and one with a BOOK place for a dialect the distro does not have must
each fail.

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
# Each placeholder and how many places the template has for it (KERNEL.2):
# the title twice, the page's <title> and its heading; the rest once.
$placeholders = [ordered]@{ '{{WASM_BASE64}}' = 1; '{{PRELUDE}}' = 1; '{{ENGLISH}}' = 1; '{{TITLE}}' = 2; '{{TAGLINE}}' = 1; '{{OPENS_WITH}}' = 1; '{{PALETTE:lavender}}' = 1; '{{PALETTE:whisper}}' = 1; '{{PALETTE:deep}}' = 1 }
# The distro the builder reads by default; its dialects are the picker's.
$manifestPath = Join-Path $repoRoot 'distros/english/distro.vla'
function Get-Dialects() {
    $names = @()
    if (Test-Path $manifestPath) {
        foreach ($l in [System.IO.File]::ReadAllLines($manifestPath)) {
            $m = [regex]::Match($l, '^\s*\(dialect "([a-z0-9-]+)"')
            if ($m.Success) { $names += $m.Groups[1].Value }
        }
    }
    return ,$names
}
# The one thing the page may hold that would otherwise read as reaching out:
# the bare link to the owner's site, whole. It must appear exactly as spelled
# here, exactly once; the scan then runs on the text with it removed, so any
# other anchor or address fails.
$allowed = @('<a href="https://spreadsheet.company/">https://spreadsheet.company</a>')

function Remove-Allowed([string]$text) {
    $problems = @()
    foreach ($a in $allowed) {
        $first = $text.IndexOf($a)
        if ($first -lt 0) { $problems += ("the page lacks its one allowed '{0}'" -f $a); continue }
        if ($text.IndexOf($a, $first + 1) -ge 0) { $problems += ("the allowed '{0}' appears more than once" -f $a) }
        $text = $text.Replace($a, '')
    }
    return @{ Text = $text; Problems = $problems }
}

function Test-Template([string]$path) {
    $problems = @()
    $text = [System.IO.File]::ReadAllText($path)
    $scrubbed = Remove-Allowed $text
    $problems += $scrubbed.Problems
    $lower = $scrubbed.Text.ToLowerInvariant()
    foreach ($t in ($tagTokens + $codeTokens)) {
        if ($lower.Contains($t.ToLowerInvariant())) { $problems += ("the page reaches out: '{0}'" -f $t) }
    }
    if (-not $lower.Contains('<meta charset="utf-8">')) { $problems += 'no <meta charset="utf-8">' }
    foreach ($ph in $placeholders.Keys) {
        $n = ([regex]::Matches($text, [regex]::Escape($ph))).Count
        if ($n -eq 0) { $problems += ("placeholder missing: {0}" -f $ph); continue }
        if ($n -ne $placeholders[$ph]) { $problems += ("placeholder {0} appears {1} time(s), and the template has {2} place(s) for it" -f $ph, $n, $placeholders[$ph]) }
    }
    # The language picker's dialects: each {{BOOK:name}} once, each a dialect
    # of the english distro, and each dialect of the distro with its place.
    $dialects = Get-Dialects
    if ($dialects.Count -eq 0) { $problems += ("no dialects read from {0}" -f $manifestPath) }
    $seen = @{}
    foreach ($m in [regex]::Matches($text, '\{\{BOOK:([a-z0-9-]+)\}\}')) {
        $name = $m.Groups[1].Value
        if ($seen.ContainsKey($name)) { $problems += ("placeholder repeated: {0}" -f $m.Value) }
        $seen[$name] = $true
        if ($dialects -notcontains $name) { $problems += ("{0} has no dialect in the english distro" -f $m.Value) }
    }
    foreach ($name in $dialects) {
        if (-not $seen.ContainsKey($name)) { $problems += ("the english distro's dialect {0} has no {{{{BOOK:{0}}}}} place in the template" -f $name) }
    }
    return $problems
}

function Test-Built([string]$path) {
    $problems = @()
    $scrubbed = Remove-Allowed ([System.IO.File]::ReadAllText($path))
    $problems += $scrubbed.Problems
    $lower = $scrubbed.Text.ToLowerInvariant()
    foreach ($t in $tagTokens) {
        if ($lower.Contains($t.ToLowerInvariant())) { $problems += ("the built page reaches out: '{0}'" -f $t) }
    }
    foreach ($ph in (@($placeholders.Keys) + @('{{BOOK:', '{{PALETTE:'))) {
        # The corpus texts may hold doubled braces of their own; only the names count.
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
        $d = Join-Path $tmp 'd.html'; [System.IO.File]::WriteAllText($d, ($text + "`n<a href=""https://example.com/"">x</a>`n"), $utf8)
        $e = Join-Path $tmp 'e.html'; [System.IO.File]::WriteAllText($e, ($text + "`n<img src=""data:image/png;base64,AAAA"">`n"), $utf8)
        $f = Join-Path $tmp 'f.html'; [System.IO.File]::WriteAllText($f, $text.Replace('{{BOOK:pirate}}', '{{BOOK:klingon}}'), $utf8)
        $ra = @(Test-Template $a); $rb = @(Test-Template $b); $rc = @(Test-Template $c); $rd = @(Test-Template $d); $re = @(Test-Template $e); $rf = @(Test-Template $f)
        Write-Output ("control: the real template has {0} problem(s); the script src {1}, the fetch {2}, the missing placeholder {3}, the second anchor {4}, the image {5}, the stray dialect place {6}" -f $real.Count, $ra.Count, $rb.Count, $rc.Count, $rd.Count, $re.Count, $rf.Count)
        if ($real.Count -eq 0 -and $ra.Count -ge 1 -and $rb.Count -ge 1 -and $rc.Count -ge 1 -and $rd.Count -ge 1 -and $re.Count -ge 1 -and $rf.Count -ge 1) {
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
    Write-Output ("OK: web/index.template.html loads nothing and links to spreadsheet.company alone, by one bare link, with its nine placeholders each in its place and the picker's dialects the english distro's; {0}" -f $builtNote)
    exit 0
}
Write-Output ("FAIL: the web page has {0} problem(s):" -f $problems.Count)
foreach ($p in $problems) { Write-Output ('  ' + $p) }
exit 1
