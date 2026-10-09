<#
check_kernel_boundary.ps1 - the kernel holds mechanism only: no English, no
sentence rule and no default enters vla-lang/src/ or core/src/ outside their
test modules (KERNEL.1, 2026-10-06; the design is web/CALLOSUM.md section
14; the kernel two crates since PORT.12, 2026-10-08).

WHY: the kernel is two crates in one dependency order, the language
(vla-lang) and the bridges over it (frazaro-core), with six seams -
sentences, paragraphs, engines, formats and hosts, projections, functions - and
everything a person reads or says is data the kernel reads (the tables under
vla-lang/data/ and core/data/, a phrasebook, a library) or an implementation
of a seam. The rule "the kernel grows a seam, never a feature" is cheap to
state and easy to break one literal at a time, so this check reads both
source trees and pins three things:

  1. the build-time includes outside a test module are exactly the five
     data tables, each crate's own: the head table and the language's half
     of the message catalogue under vla-lang/data/ (headtable, messages);
     the core's half, the word tables and the name lists under core/data/
     (messages, words, names). No prelude, no phrasebook and no corpus file
     is baked into either crate. The door carries its built-in pair under
     cli/data/, which is the door's business, not the kernel's;
  2. a slot-bearing literal - a string holding a {slot:type} with words
     beside it, which is what a sentence rule looks like - sits only where
     the reference itself writes such text in code: eleven in
     core/src/english/grammar.rs, ten of them the built-in rules that mirror
     the AddPhraseRule calls of src/VLA_SentenceEngine.bas (the eleventh
     built-in, "stop", carries no slot) and one the description of the slot
     kinds a refusal names (SD-18, the port follows the reference). A bare
     "{name:category}" with no words beside it is the slot syntax's own
     example and does not count. The count is pinned exactly, so the
     exception never grows, and a retirement lowers the pin deliberately,
     with its reason;
  3. an English sentence as a string literal - five or more words, ending
     in . : ? or ! - sits only in the pinned places: two in core/src/build.rs
     (the line `frazaro rebuild` prints, the door's chrome until the
     EDITION line moves it to data) and one in core/src/english/matcher.rs
     (the parse error's words, the reference's own, held to it by the
     refusal golden). Pinned exactly too. The language's tree holds none.

Comment lines are skipped, since a doc comment may quote a sentence, and the
test region is skipped as check_crate_package.ps1 skips it (the `#[cfg(test)]`
that heads a `mod` at a file's end), since fixtures quote the corpus on
purpose.

-Control proves the check on scratch copies of both source trees: the clean
copy passes; a prelude include, a sentence rule and an English sentence
planted outside a test module (in the language's form.rs) each fail; the
same sentence planted inside one passes; and one built-in rule removed from
grammar.rs fails, since the pin moved.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, a hardcoded and
reviewable baseline, the counts pinned. Exit 0 clean, exit 1 with every
problem named.

Usage:  powershell -File tools\check_kernel_boundary.ps1
        powershell -File tools\check_kernel_boundary.ps1 -Root <dir>   # a tree holding vla-lang/src and core/src
        powershell -File tools\check_kernel_boundary.ps1 -Control
#>
param(
    [string]$Root = '',
    [switch]$Control
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# ---- the baseline: each crate's own data, and the pinned exceptions ----
# The source trees, and the tables each bakes in, as the includes spell them
# (words.rs sits one folder deeper in the core).
$trees = @('vla-lang/src', 'core/src')
$dataIncludes = @{
    'vla-lang/src' = @('../data/headtable.vla', '../data/messages.vla')
    'core/src'     = @('../data/messages.vla', '../../data/words.vla', '../../data/names.vla')
}
# Slot-bearing literals outside tests, per file; any other file must hold none.
# grammar.rs: ten built-in rules with a slot, and the slot-kinds description a refusal names.
$ruleLiteralPins = @{ 'core/src/english/grammar.rs' = 11 }
# English-sentence literals outside tests, per file; any other file must hold none.
$prosePins = @{ 'core/src/build.rs' = 2; 'core/src/english/matcher.rs' = 1 }

function Test-Prose([string]$literal) {
    # Five or more words, the first beginning with a letter, ending in a sentence mark.
    if ($literal -notmatch '^[A-Za-z]') { return $false }
    if ($literal -notmatch '[.:?!]$') { return $false }
    $words = @($literal -split '\s+' | Where-Object { $_ -match "^[A-Za-z{']" })
    return ($words.Count -ge 5)
}

function Get-Failures([string]$root) {
    $failures = New-Object System.Collections.Generic.List[string]
    $root = $root.TrimEnd('\', '/')
    $ruleCounts = @{}
    $proseCounts = @{}
    $literalsSeen = 0

    foreach ($tree in $trees) {
        $srcDir = Join-Path $root ($tree -replace '/', '\')
        if (-not (Test-Path $srcDir)) { $failures.Add("no $tree under $root"); continue }
        $includes = New-Object System.Collections.Generic.List[string]
        $wanted = $dataIncludes[$tree]

        foreach ($f in Get-ChildItem -Path $srcDir -Filter '*.rs' -Recurse) {
            $lines = @(Get-Content $f.FullName)
            $rel = $f.FullName.Substring($root.Length + 1) -replace '\\', '/'
            $testStart = $lines.Count
            for ($i = 0; $i -lt $lines.Count - 1; $i++) {
                if ($lines[$i] -match '^#\[cfg\(test\)\]\s*$' -and $lines[$i + 1] -match '^mod\s+\w+') { $testStart = $i; break }
            }
            for ($i = 0; $i -lt $testStart; $i++) {
                $line = $lines[$i]
                if ($line.TrimStart() -match '^//') { continue }
                foreach ($mt in [regex]::Matches($line, 'include_(?:str|bytes)!\(\s*"([^"]+)"\s*\)')) {
                    $includes.Add($mt.Groups[1].Value)
                    if ($wanted -notcontains $mt.Groups[1].Value) {
                        $failures.Add("${rel}:$($i + 1): the kernel bakes in `"$($mt.Groups[1].Value)`", which is not one of this crate's data tables (a prelude, a phrasebook or a corpus file is the door's to carry, under cli/data/)")
                    }
                }
                foreach ($mt in [regex]::Matches($line, '"((?:[^"\\]|\\.)*)"')) {
                    $lit = $mt.Groups[1].Value
                    $literalsSeen++
                    if ($lit -match '\{[a-z]+:[a-z-]+\}') {
                        # Words beside the slot make it a rule's shape; a bare {name:category} is the syntax's own example.
                        $outside = $lit -replace '\{[^}]*\}', ''
                        if ($outside -notmatch '[A-Za-z]') { continue }
                        if ($ruleCounts.ContainsKey($rel)) { $ruleCounts[$rel]++ } else { $ruleCounts[$rel] = 1 }
                        if (-not $ruleLiteralPins.ContainsKey($rel)) {
                            $failures.Add("${rel}:$($i + 1): a sentence rule in the kernel, `"$lit`": a rule is a phrasebook's, loaded as data")
                        }
                        continue
                    }
                    if (Test-Prose $lit) {
                        if ($proseCounts.ContainsKey($rel)) { $proseCounts[$rel]++ } else { $proseCounts[$rel] = 1 }
                        if (-not $prosePins.ContainsKey($rel)) {
                            $failures.Add("${rel}:$($i + 1): an English sentence in the kernel, `"$lit`": a refusal goes through the catalogue, and chrome is the door's")
                        }
                    }
                }
            }
        }

        # 1. exactly this crate's tables, each present
        foreach ($want in $wanted) {
            if ($includes -notcontains $want) { $failures.Add("$tree no longer includes `"$want`": the data tables are each crate's own, and a table that moves lowers this list deliberately") }
        }
    }
    # 2 and 3. the pinned exceptions, exactly
    foreach ($k in $ruleLiteralPins.Keys) {
        $have = if ($ruleCounts.ContainsKey($k)) { $ruleCounts[$k] } else { 0 }
        if ($have -ne $ruleLiteralPins[$k]) { $failures.Add("${k}: $have slot-bearing literal(s) outside tests, pinned at $($ruleLiteralPins[$k]) (the reference's built-ins and its slot-kinds description; a change there changes the pin, with its reason)") }
    }
    foreach ($k in $prosePins.Keys) {
        $have = if ($proseCounts.ContainsKey($k)) { $proseCounts[$k] } else { 0 }
        if ($have -ne $prosePins[$k]) { $failures.Add("${k}: $have English-sentence literal(s) outside tests, pinned at $($prosePins[$k])") }
    }
    if ($literalsSeen -lt 100) { $failures.Add("only $literalsSeen string literal(s) read outside tests: the scan is not seeing the sources") }
    return ,$failures
}

function New-ScratchTree([string]$name) {
    $dir = Join-Path ([System.IO.Path]::GetTempPath()) "frazaro_kernel_boundary_$name"
    if (Test-Path $dir) { Remove-Item -Recurse -Force $dir }
    foreach ($tree in $trees) {
        $parent = Split-Path -Parent ($tree -replace '/', '\')
        New-Item -ItemType Directory -Path (Join-Path $dir $parent) -Force | Out-Null
        Copy-Item -Recurse (Join-Path $repoRoot ($tree -replace '/', '\')) (Join-Path $dir ($tree -replace '/', '\'))
    }
    return $dir
}

function Add-TopLine([string]$path, [string]$text) {
    $body = [System.IO.File]::ReadAllText($path)
    [System.IO.File]::WriteAllText($path, $text + "`n" + $body)
}

if ($Control) {
    $verdicts = New-Object System.Collections.Generic.List[string]
    $ok = $true

    $clean = New-ScratchTree 'clean'
    $f0 = Get-Failures $clean
    if ($f0.Count -ne 0) { $ok = $false; $verdicts.Add("clean copy failed: $($f0 -join ' | ')") } else { $verdicts.Add('clean copy passes') }
    Remove-Item -Recurse -Force $clean

    $mutants = @(
        @{ Name = 'a prelude include outside tests'; Line = 'pub const KB_MUTANT: &str = include_str!("../../scripts/prelude.vla");'; Inside = $false; WantFail = $true },
        @{ Name = 'a sentence rule outside tests';   Line = 'pub const KB_MUTANT: &str = "put {x:expr} into cell {r:cell}";'; Inside = $false; WantFail = $true },
        @{ Name = 'an English sentence outside tests'; Line = 'pub const KB_MUTANT: &str = "Nothing was written because the sheet is full.";'; Inside = $false; WantFail = $true },
        @{ Name = 'the same sentence inside a test module'; Line = 'const KB_MUTANT: &str = "Nothing was written because the sheet is full.";'; Inside = $true; WantFail = $false }
    )
    foreach ($m in $mutants) {
        $tree = New-ScratchTree 'mutant'
        $target = Join-Path $tree 'vla-lang\src\form.rs'
        if ($m.Inside) {
            $body = [System.IO.File]::ReadAllText($target)
            [System.IO.File]::WriteAllText($target, $body + "`n#[cfg(test)]`nmod kb_mutant {`n    $($m.Line)`n}`n")
        } else {
            Add-TopLine $target $m.Line
        }
        $fails = Get-Failures $tree
        $failed = ($fails.Count -gt 0)
        if ($failed -eq $m.WantFail) {
            $verdicts.Add("$($m.Name): " + $(if ($failed) { 'fails, as it should' } else { 'passes, as it should' }))
        } else {
            $ok = $false
            $verdicts.Add("$($m.Name): " + $(if ($failed) { "FAILED but should pass: $($fails -join ' | ')" } else { 'PASSED but should fail' }))
        }
        Remove-Item -Recurse -Force $tree
    }

    # one built-in rule removed: the pin moves
    $tree = New-ScratchTree 'mutant'
    $g = Join-Path $tree 'core\src\english\grammar.rs'
    $kept = @(Get-Content $g | Where-Object { $_ -notmatch '\("say \{e:expr\}", "\(msgbox \{e\}\)"\),' })
    [System.IO.File]::WriteAllLines($g, $kept)
    $fails = Get-Failures $tree
    if ($fails.Count -gt 0) { $verdicts.Add('a built-in rule removed: fails, as it should (the pin moved)') } else { $ok = $false; $verdicts.Add('a built-in rule removed: PASSED but should fail') }
    Remove-Item -Recurse -Force $tree

    # the language's table moved away: the include list is held per crate
    $tree = New-ScratchTree 'mutant'
    $h = Join-Path $tree 'vla-lang\src\headtable.rs'
    $body = [System.IO.File]::ReadAllText($h)
    [System.IO.File]::WriteAllText($h, ($body -replace 'include_str!\("\.\./data/headtable\.vla"\)', 'include_str!("../../core/data/headtable.vla")'))
    $fails = Get-Failures $tree
    if ($fails.Count -gt 0) { $verdicts.Add("the language's head table reached from another crate: fails, as it should") } else { $ok = $false; $verdicts.Add("the language's head table reached from another crate: PASSED but should fail") }
    Remove-Item -Recurse -Force $tree

    if ($ok) {
        Write-Host 'OK: control: the clean copy passes; three plants outside a test module fail, the plant inside one passes, a removed built-in rule moves the pin, and a table reached across the crate line fails'
        $verdicts | ForEach-Object { Write-Host "  - $_" }
        exit 0
    }
    Write-Host 'FAIL: control:'
    $verdicts | ForEach-Object { Write-Host "  - $_" }
    exit 1
}

$root = if ($Root -ne '') { (Resolve-Path $Root).Path } else { $repoRoot }
$failures = Get-Failures $root
if ($failures.Count -gt 0) {
    Write-Host "FAIL: $($failures.Count) problem(s)"
    $failures | ForEach-Object { Write-Host "  - $_" }
    exit 1
}
Write-Host "OK: the two crates bake in exactly their five data tables (vla-lang: the head table and its half of the catalogue; core: its half, the words and the names); slot-bearing literals outside tests only in grammar.rs ($($ruleLiteralPins['core/src/english/grammar.rs']): the reference's built-in rules and its slot-kinds description); English sentences outside tests only where pinned (build.rs $($prosePins['core/src/build.rs']), matcher.rs $($prosePins['core/src/english/matcher.rs']))"
exit 0
