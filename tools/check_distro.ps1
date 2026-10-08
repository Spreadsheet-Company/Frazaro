<#
check_distro.ps1 - every distro under distros/ is well formed, names files
that exist, and the three doors read it as one thing.

WHY: KERNEL.2 (2026-10-07) makes a distro folder the unit the build tools
take (web/CALLOSUM.md section 14.1): distros/<name>/distro.vla names the
prelude, the base phrasebook, an overlay, the dialects, the libraries, the
examples and the edition's chrome, by reference, and three readers take it -
core/src/distro.rs for `frazaro prove <folder>`, tools/build_web.ps1 for the
page, and src/VLA_Build.bas for the add-in. Three readers of one file is
the drift check_data_exports.ps1 and check_crate_data.ps1 were written
for, so this check pins, host-free and on every push:

  1. the folders under distros/ are exactly the ones pinned here, each with
     a distro.vla and a README.md that states the phrasebooks' terms;
  2. each manifest has the shape the two line-based readers (the page
     builder and the add-in builder) rely on: one (distro "name" ...) form,
     the opener alone on its line, one directive a line, every head one
     this version knows, the name the folder's own, the required directives
     once each (title, tagline, opens-with, palette with its three shades,
     prelude, phrasebook), at most one overlay, dialects and libraries each
     named once, every path written with forward slashes and relative to
     the folder, and every path pointing at something that exists;
  3. the english distro is what the doors ship: its prelude is
     scripts/prelude.vla and its base scripts/polyglotta/english.vla, the
     two files check_crate_data.ps1 holds the door's built-in copies to; its
     dialects are exactly the language picker's {{BOOK:name}} places in
     web/index.template.html; and their count is pinned;
  4. every edition VLA_Build.bas's VlaEditionNames builds has a distro
     folder, so VlaBuildAddin never names an edition by hand again.

-Control proves the check on a scratch copy of the files it reads: the copy
passes; a manifest naming a dialect whose file is missing, a manifest whose
name is not its folder's, a template with a BOOK place for a dialect the
distro lacks, a manifest with two (phrasebook ...) lines, and a
VlaEditionNames naming an edition with no folder must each fail.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no
COM, no network; a hardcoded, reviewable baseline; exit 0 clean, exit 1
with every problem named.

Usage:  powershell -File tools\check_distro.ps1
        powershell -File tools\check_distro.ps1 -Root <dir>   # a tree holding distros/, web/, src/, scripts/
        powershell -File tools\check_distro.ps1 -Control
#>
param(
    [string]$Root = '',
    [switch]$Control
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# ---- the baseline ----
# 2026-10-07, KERNEL.2: english (the doors' edition, seven dialects, no
# overlay) and espanol (the add-in's Spanish edition: english then the
# espanol overlay). A distro added raises the pin, with its reason.
$expectedDistros = @('english', 'espanol')
$expectedEnglishDialects = 7
$requiredOnce = @('title', 'tagline', 'opens-with', 'palette', 'prelude', 'phrasebook')
$knownHeads = @('title', 'tagline', 'opens-with', 'palette', 'prelude', 'phrasebook', 'overlay', 'dialect', 'library', 'examples', 'addin', 'readme')
$shades = @('lavender', 'whisper', 'deep')

# The quoted strings of one manifest line, in order, \" and \\ read as the
# language reads them.
function Get-DirectiveStrings([string]$line) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($m in [regex]::Matches($line, '"((?:[^"\\]|\\.)*)"')) { $out.Add(($m.Groups[1].Value -replace '\\(.)', '$1')) }
    return ,$out
}

function Resolve-DistroPath([string]$folder, [string]$rel) {
    return [System.IO.Path]::GetFullPath((Join-Path $folder ($rel -replace '/', '\')))
}

# A line's parentheses opened minus closed, outside its quoted strings: a
# directive nets 0, the opener nets 1, and the last directive nets -1 as it
# closes the form.
function Get-ParenNet([string]$line) {
    $bare = [regex]::Replace($line, '"(?:[^"\\]|\\.)*"', '')
    return ([regex]::Matches($bare, '\(')).Count - ([regex]::Matches($bare, '\)')).Count
}

# One manifest read as the line-based readers read it; problems added to
# $failures under $tag; a table of what was read, or $null when the shape
# is too broken to go on.
function Read-Manifest([string]$path, [string]$tag, $failures) {
    $d = @{ Name = ''; Prelude = $null; Base = $null; Overlay = $null; Dialects = @(); Paths = (New-Object System.Collections.Generic.List[object]); Counts = @{} }
    $opened = $false; $closed = $false
    foreach ($raw in [System.IO.File]::ReadAllLines($path)) {
        $t = $raw.Trim()
        if ($t -eq '' -or $t.StartsWith(';')) { continue }
        if (-not $opened) {
            $m = [regex]::Match($t, '^\(distro "([a-z0-9-]+)"$')
            if (-not $m.Success) { $failures.Add("${tag}: the first directive is not (distro ""name"") alone on its line, the name lowercase letters, digits and hyphens: $t"); return $null }
            $d.Name = $m.Groups[1].Value; $opened = $true; continue
        }
        if ($closed) { $failures.Add("${tag}: text after the form closed: $t"); continue }
        $hm = [regex]::Match($t, '^\(([a-z-]+)\b')
        if (-not $hm.Success) { $failures.Add("${tag}: not a directive: $t"); continue }
        $head = $hm.Groups[1].Value
        $net = Get-ParenNet $t
        if ($net -eq -1) { $closed = $true } elseif ($net -ne 0) { $failures.Add("${tag}: a directive opens and closes on its own line: $t") }
        if ($knownHeads -notcontains $head) { $failures.Add("${tag}: ($head ...) is not a directive this version knows"); continue }
        if ($d.Counts.ContainsKey($head)) { $d.Counts[$head]++ } else { $d.Counts[$head] = 1 }
        $s = Get-DirectiveStrings $t
        switch ($head) {
            { $_ -in @('title', 'tagline', 'opens-with') } { if ($s.Count -ne 1) { $failures.Add("${tag}: ($head ...) wants one quoted string") } }
            'palette' {
                foreach ($shade in $shades) {
                    if ($t -notmatch ('\(' + $shade + ' "#[0-9A-Fa-f]{6}"\)')) { $failures.Add("${tag}: (palette ...) names no ($shade ""#rrggbb"")") }
                }
            }
            { $_ -in @('prelude', 'examples', 'readme') } {
                if ($s.Count -ne 1) { $failures.Add("${tag}: ($head ...) wants one quoted path") }
                else { $d.Paths.Add(@{ Head = $head; Path = $s[0] }); if ($head -eq 'prelude') { $d.Prelude = $s[0] } }
            }
            { $_ -in @('phrasebook', 'overlay', 'dialect', 'library') } {
                if ($s.Count -ne 2) { $failures.Add("${tag}: ($head ...) wants a name and a path"); break }
                if ($s[0] -notmatch '^[a-z0-9-]+$') { $failures.Add("${tag}: ($head ""$($s[0])"" ...) is not a lowercase name") }
                $d.Paths.Add(@{ Head = $head; Path = $s[1] })
                if ($head -eq 'phrasebook') { $d.Base = $s[1] }
                if ($head -eq 'overlay') { $d.Overlay = $s[1] }
                if ($head -eq 'dialect') { if ($d.Dialects -contains $s[0]) { $failures.Add("${tag}: (dialect ""$($s[0])"" ...) is given twice") }; $d.Dialects += $s[0] }
            }
            'addin' { if ($s.Count -ne 1 -or $s[0] -eq '' -or $s[0].Contains('/') -or $s[0].Contains('\')) { $failures.Add("${tag}: (addin ...) wants one bare file name") } }
        }
    }
    if (-not $opened) { $failures.Add("${tag}: the manifest holds no (distro ...) form"); return $null }
    if (-not $closed) { $failures.Add("${tag}: the form never closes (the last directive ends with two parentheses)") }
    foreach ($h in $requiredOnce) {
        $n = if ($d.Counts.ContainsKey($h)) { $d.Counts[$h] } else { 0 }
        if ($n -ne 1) { $failures.Add("${tag}: it names ($h ...) $n time(s), not once") }
    }
    if ($d.Counts.ContainsKey('overlay') -and $d.Counts['overlay'] -gt 1) { $failures.Add("${tag}: (overlay ...) is given twice") }
    foreach ($p in $d.Paths) {
        if ($p.Path -eq '' -or $p.Path.Contains('\') -or $p.Path.StartsWith('/') -or $p.Path -match '^[A-Za-z]:') { $failures.Add("${tag}: ($($p.Head) ...) must name a path relative to the folder with forward slashes, not '$($p.Path)'") }
    }
    return $d
}

function Get-Failures([string]$root) {
    $failures = New-Object System.Collections.Generic.List[string]
    $distrosDir = Join-Path $root 'distros'
    if (-not (Test-Path $distrosDir)) { $failures.Add('no distros/ folder'); return ,$failures }
    $folders = @(Get-ChildItem -Directory $distrosDir | Sort-Object Name)
    $names = @($folders | ForEach-Object { $_.Name })
    if (($names -join ',') -ne ($expectedDistros -join ',')) { $failures.Add("distros/ holds [$($names -join ', ')], and the distros pinned here are [$($expectedDistros -join ', ')]; a distro added or removed changes the pin, with its reason") }
    $english = $null
    foreach ($f in $folders) {
        $tag = 'distros/' + $f.Name
        $manifest = Join-Path $f.FullName 'distro.vla'
        $readme = Join-Path $f.FullName 'README.md'
        if (-not (Test-Path $manifest)) { $failures.Add("${tag}: no distro.vla"); continue }
        if (-not (Test-Path $readme)) { $failures.Add("${tag}: no README.md") }
        elseif (-not ([System.IO.File]::ReadAllText($readme).Contains('MPL-2.0'))) { $failures.Add("${tag}: README.md does not state the phrasebooks' terms (MPL-2.0)") }
        $d = Read-Manifest $manifest $tag $failures
        if ($null -eq $d) { continue }
        if ($d.Name -ne $f.Name) { $failures.Add("${tag}: the manifest names the distro '$($d.Name)', and the folder is '$($f.Name)'") }
        foreach ($p in $d.Paths) {
            if ($p.Path -eq '' -or $p.Path.Contains('\')) { continue }
            $full = Resolve-DistroPath $f.FullName $p.Path
            if (-not (Test-Path $full)) { $failures.Add("${tag}: ($($p.Head) ...) names $($p.Path), and there is nothing at $full") }
        }
        if ($f.Name -eq 'english') { $english = $d; $english.Folder = $f.FullName }
    }
    if ($null -eq $english) {
        $failures.Add('no english distro to hold the doors to')
    } else {
        $wantPrelude = [System.IO.Path]::GetFullPath((Join-Path $root 'scripts/prelude.vla'))
        $wantBase = [System.IO.Path]::GetFullPath((Join-Path $root 'scripts/polyglotta/english.vla'))
        if ($null -ne $english.Prelude -and (Resolve-DistroPath $english.Folder $english.Prelude) -ne $wantPrelude) { $failures.Add("distros/english: the prelude is $($english.Prelude), not scripts/prelude.vla, the file check_crate_data.ps1 holds the door's built-in prelude to") }
        if ($null -ne $english.Base -and (Resolve-DistroPath $english.Folder $english.Base) -ne $wantBase) { $failures.Add("distros/english: the base phrasebook is $($english.Base), not scripts/polyglotta/english.vla, the file check_crate_data.ps1 holds the door's built-in english.vla to") }
        if ($english.Dialects.Count -ne $expectedEnglishDialects) { $failures.Add("distros/english: $($english.Dialects.Count) dialect(s), and $expectedEnglishDialects are pinned here; a dialect added or removed changes the pin and the picker") }
        $templatePath = Join-Path $root 'web/index.template.html'
        if (-not (Test-Path $templatePath)) { $failures.Add('no web/index.template.html to hold the picker to') }
        else {
            $places = @([regex]::Matches([System.IO.File]::ReadAllText($templatePath), '\{\{BOOK:([a-z0-9-]+)\}\}') | ForEach-Object { $_.Groups[1].Value })
            foreach ($p in $places) { if ($english.Dialects -notcontains $p) { $failures.Add("web/index.template.html: {{BOOK:$p}} has no dialect in the english distro") } }
            foreach ($dn in $english.Dialects) { if ($places -notcontains $dn) { $failures.Add("distros/english: the dialect $dn has no {{BOOK:$dn}} place in web/index.template.html") } }
        }
    }
    $buildPath = Join-Path $root 'src/VLA_Build.bas'
    if (-not (Test-Path $buildPath)) { $failures.Add('no src/VLA_Build.bas to read VlaEditionNames from') }
    else {
        $m = [regex]::Match([System.IO.File]::ReadAllText($buildPath), 'VlaEditionNames = Array\(([^)]*)\)')
        if (-not $m.Success) { $failures.Add('src/VLA_Build.bas: no VlaEditionNames = Array(...) line; the add-in builder names its editions there and reads each from distros/') }
        else {
            foreach ($em in [regex]::Matches($m.Groups[1].Value, '"([^"]+)"')) {
                $edition = $em.Groups[1].Value.ToLowerInvariant()
                if (-not (Test-Path (Join-Path $distrosDir ($edition + '/distro.vla')))) { $failures.Add("src/VLA_Build.bas: VlaEditionNames builds '$($em.Groups[1].Value)', and distros/$edition/distro.vla is not there") }
            }
        }
    }
    return ,$failures
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_distro_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        # A scratch tree holding what the check reads, copied once.
        foreach ($rel in @('distros', 'scripts/polyglotta', 'examples')) {
            $dst = Join-Path $tmp $rel
            New-Item -ItemType Directory -Path $dst -Force | Out-Null
            Copy-Item -Path (Join-Path (Join-Path $repoRoot $rel) '*') -Destination $dst -Recurse -Force
        }
        foreach ($rel in @('scripts/prelude.vla', 'web/index.template.html', 'src/VLA_Build.bas')) {
            $dst = Join-Path $tmp $rel
            New-Item -ItemType Directory -Path (Split-Path -Parent $dst) -Force | Out-Null
            Copy-Item -Path (Join-Path $repoRoot $rel) -Destination $dst -Force
        }
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $manifest = Join-Path $tmp 'distros/english/distro.vla'
        $template = Join-Path $tmp 'web/index.template.html'
        $build = Join-Path $tmp 'src/VLA_Build.bas'
        $m0 = [System.IO.File]::ReadAllText($manifest); $t0 = [System.IO.File]::ReadAllText($template); $b0 = [System.IO.File]::ReadAllText($build)
        $clean = Get-Failures $tmp
        $verdicts = New-Object System.Collections.Generic.List[string]
        $ok = ($clean.Count -eq 0)
        $verdicts.Add(("the clean copy: {0} problem(s)" -f $clean.Count))
        $mutants = @(
            @{ Name = 'a dialect whose file is missing'; Apply = { [System.IO.File]::WriteAllText($manifest, $m0.Replace('  (library "alien"', '  (dialect "klingon" "../../scripts/polyglotta/klingon.vla")' + "`r`n" + '  (library "alien"'), $utf8) } },
            @{ Name = 'a name that is not the folder''s'; Apply = { [System.IO.File]::WriteAllText($manifest, $m0.Replace('(distro "english"', '(distro "englsh"'), $utf8) } },
            @{ Name = 'a BOOK place for a dialect the distro lacks'; Apply = { [System.IO.File]::WriteAllText($template, $t0.Replace('{{BOOK:pirate}}', '{{BOOK:klingon}}'), $utf8) } },
            @{ Name = 'two (phrasebook ...) lines'; Apply = { [System.IO.File]::WriteAllText($manifest, $m0.Replace('  (library "alien"', '  (phrasebook "again" "../../scripts/polyglotta/english.vla")' + "`r`n" + '  (library "alien"'), $utf8) } },
            @{ Name = 'an edition with no folder'; Apply = { [System.IO.File]::WriteAllText($build, $b0.Replace('VlaEditionNames = Array("English", "Espanol")', 'VlaEditionNames = Array("English", "Espanol", "Klingon")'), $utf8) } }
        )
        foreach ($mu in $mutants) {
            & $mu.Apply
            $r = Get-Failures $tmp
            [System.IO.File]::WriteAllText($manifest, $m0, $utf8); [System.IO.File]::WriteAllText($template, $t0, $utf8); [System.IO.File]::WriteAllText($build, $b0, $utf8)
            if ($r.Count -ge 1) { $verdicts.Add(("{0}: fails, as it should ({1}: {2})" -f $mu.Name, $r.Count, $r[0])) } else { $ok = $false; $verdicts.Add(("{0}: PASSED but should fail" -f $mu.Name)) }
        }
        foreach ($v in $verdicts) { Write-Output ('control: ' + $v) }
        if ($ok) { Write-Output 'OK: control: the clean copy passes and each of the five mutants fails'; exit 0 }
        foreach ($p in $clean) { Write-Output ('  clean: ' + $p) }
        Write-Output 'FAIL: the control did not behave as the header says'
        exit 1
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}

$root = if ($Root -ne '') { (Resolve-Path $Root).Path } else { $repoRoot }
$failures = Get-Failures $root
if ($failures.Count -eq 0) {
    Write-Output ("OK: {0} distro(s) under distros/ ({1}), each a well-formed manifest naming files that exist; the english distro is the doors' (scripts/prelude.vla, scripts/polyglotta/english.vla, {2} dialects, the picker's places exactly), and every edition VlaBuildAddin builds has its folder" -f $expectedDistros.Count, ($expectedDistros -join ', '), $expectedEnglishDialects)
    exit 0
}
Write-Output ("FAIL: {0} problem(s) with the distros:" -f $failures.Count)
foreach ($p in $failures) { Write-Output ('  ' + $p) }
exit 1
