<#
check_appearance_golden.ps1 - the appearance golden agrees with its fixture,
every case is two readings in the record's spelling or a refusal, and the
cases reach every property a format says this version.

WHY: KERNEL.50 (a format's appearance in the language) gives a format a
text's weight, slant, size and color, its alignment and its cell's four
edges, and Excel is the reference for each: scripts/appearance.txt holds the
cases, a format a case as the record spells it, and
scripts/appearance_golden.txt is what VlaWriteAppearanceGolden
(VLA_Tests.bas) writes of them in Excel: each case set on a cell through the
object model and read back (EXCEL), and the same case read back from the
cell frazaro-core's writer gave it in scripts/appearance_written.xlsx
(WRITTEN), each reading in the record's spelling with Excel's defaults left
out. The language's test, vla-lang/tests/appearance_golden.rs, holds each
case's print and each reading to the case. That test proves agreement on
what the fixture asks. It cannot see a property no case reaches, a golden
regenerated against a fixture that lost a case, or a refusal whose id the
catalogue lacks. This check holds those, host-free, on every push:

  1. the golden's headers are the fixture's, in order, and every fixture
     case is one line;
  2. every golden case is two records, EXCEL<TAB><a reading> or
     EXCEL<TAB>REFUSED <Excel's error number>, then WRITTEN<TAB><a reading>,
     or WRITTEN<TAB>REFUSED for a case whose header ends "!refused <words>"
     (a case the language refuses) and for no other;
  3. every reading is in the record's spelling: a fill, a number format and
     a wrap, then a font, an alignment and a border, each said only where it
     is not Excel's default (no not-bold, not-italic, 11, general, bottom or
     none) and no group left empty;
  4. grid-format-property, the refusal of a property this version does not
     take, is an id of the language's half of the catalogue
     (vla-lang/data/messages.vla);
  5. the properties the taken cases reach, each named below, all of them:
     each weight and slant said and unsaid, a size whole, of a half, of
     tenths and of hundredths, the least and the most, a font's color;
     every horizontal and vertical; every side and every one of the
     fourteen lines, an edge's color; a fill and a wrap; and the three
     groups in one format;
  6. scripts/appearance_written.xlsx is there, the workbook the WRITTEN
     readings read;
  7. the floors below: the cases and the cases refused here.

-Control proves the check on scratch copies: the real files must pass, and
each of these must fail: a golden with one WRITTEN tag misspelled, one with
its last case dropped, one reading a word the record has none for, one
reading a default said, one reading REFUSED where the language takes the
case, and a fixture whose slant-dash-dot case says thin.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; exit 0 clean, exit 1 with every problem named.

Usage:  powershell -File tools\check_appearance_golden.ps1
        powershell -File tools\check_appearance_golden.ps1 -Control
#>
param(
    [switch]$Control
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# --- the floors, never lowered ------------------------------------------------
# 2026-10-22, KERNEL.50: 46 cases, 2 refused here; the golden written as the
# prediction ahead of the reference's first run in Excel.
$caseFloor = 46
$refusedFloor = 2

# The properties rule 5 holds the taken cases to.
$properties = @('bold', 'not-bold', 'italic', 'not-italic', 'a whole size',
    'a size of a half', 'a size of tenths', 'a size of hundredths', 'the least size',
    'the most size', "a font's color", 'general', 'left', 'center', 'right', 'top',
    'middle', 'bottom', 'the left edge', 'the right edge', 'the top edge',
    'the bottom edge', 'none', 'thin', 'medium', 'thick', 'double', 'hair', 'dotted',
    'dashed', 'dash-dot', 'dash-dot-dot', 'medium-dashed', 'medium-dash-dot',
    'medium-dash-dot-dot', 'slant-dash-dot', "an edge's color", 'a fill', 'a wrap',
    'the three groups')

$fixturePath = Join-Path $repoRoot 'scripts/appearance.txt'
$goldenPath  = Join-Path $repoRoot 'scripts/appearance_golden.txt'
$catalogPath = Join-Path $repoRoot 'vla-lang/data/messages.vla'
$writtenPath = Join-Path $repoRoot 'scripts/appearance_written.xlsx'

# A reading in the record's spelling (rule 3), every group optional and in
# the record's order, each property in its own.
$rgb = '"[0-9A-F]{6}"'
$sizeText = '[1-9][0-9]{0,2}(?:\.[1-9]|\.[0-9]5)?'
$lineWords = '(?:medium-dash-dot-dot|medium-dash-dot|medium-dashed|slant-dash-dot|dash-dot-dot|dash-dot|dashed|dotted|double|medium|thick|thin|hair)'
$edgeText = { param($side) "(?: \($side $lineWords(?: $rgb)?\))?" }
$readingPattern = '^(?:none|' + $rgb + ') (?:general|text|"(?:[^"\\]|\\.)*") (?:wrap|nowrap)' +
    '(?: \(font(?: bold)?(?: italic)?(?: (?<size>' + $sizeText + '))?(?: ' + $rgb + ')?\))?' +
    '(?: \(align(?: (?:left|center|right))?(?: (?:top|middle))?\))?' +
    '(?: \(border' + (& $edgeText 'left') + (& $edgeText 'right') + (& $edgeText 'top') + (& $edgeText 'bottom') + '\))?$'

function Get-AppearanceLines([string]$path) {
    $text = [System.IO.File]::ReadAllText($path)
    $text = $text -replace "`r`n", "`n" -replace "`r", "`n"
    return ,@($text -split "`n")
}

function New-OrdinalSet {
    return ,(New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal))
}

# The cases of a fixture or a golden: each header with the lines under it,
# blank lines skipped, as VlaWriteAppearanceGolden reads them.
function Read-AppearanceCases([string]$path) {
    $cases = New-Object System.Collections.ArrayList
    $current = $null
    foreach ($ln in (Get-AppearanceLines $path)) {
        if ($ln.StartsWith('=== ')) {
            if ($null -ne $current) { [void]$cases.Add($current) }
            $current = @{ Header = $ln; Lines = (New-Object System.Collections.ArrayList) }
        } elseif ($null -ne $current -and $ln -ne '') {
            [void]$current.Lines.Add($ln)
        }
    }
    if ($null -ne $current) { [void]$cases.Add($current) }
    return ,$cases
}

# The problems of one reading's spelling (rule 3), empty where it is the
# record's.
function Test-Reading([string]$header, [string]$tag, [string]$reading) {
    $found = New-Object System.Collections.ArrayList
    $m = [regex]::Match($reading, $readingPattern)
    if (-not $m.Success) {
        [void]$found.Add(("'{0}': {1} reads '{2}', not in the record's spelling" -f $header, $tag, $reading))
        return ,$found
    }
    if ($reading -cmatch '\((font|align|border)\)') {
        [void]$found.Add(("'{0}': {1} reads an empty group: '{2}'" -f $header, $tag, $reading))
    }
    if ($m.Groups['size'].Success -and $m.Groups['size'].Value -ceq '11') {
        [void]$found.Add(("'{0}': {1} reads the default size, 11: '{2}'" -f $header, $tag, $reading))
    }
    return ,$found
}

# The properties one taken case reaches (rule 5), read from its tokens: the
# fill and the wrap, then each group's words by its head, and each edge's
# side, line and color.
function Get-CaseProperties([string]$line) {
    $found = New-OrdinalSet
    $tokens = @([regex]::Matches($line, '"[^"]*"|\(|\)|[^\s()"]+') | ForEach-Object { $_.Value })
    if ($tokens.Count -ge 3) {
        if ($tokens[0] -ne 'none') { [void]$found.Add('a fill') }
        if ($tokens[2] -ceq 'wrap') { [void]$found.Add('a wrap') }
    }
    $depth = 0
    $head = ''
    $groups = New-OrdinalSet
    $edgeItem = 0
    for ($i = 3; $i -lt $tokens.Count; $i++) {
        $t = $tokens[$i]
        if ($t -eq '(') { $depth++; if ($depth -eq 2) { $edgeItem = 0 }; continue }
        if ($t -eq ')') { $depth--; continue }
        if ($depth -eq 1 -and $tokens[$i - 1] -eq '(') {
            $head = $t
            [void]$groups.Add($t)
            continue
        }
        if ($depth -eq 1 -and $head -ceq 'font') {
            if ($t.StartsWith('"')) { [void]$found.Add("a font's color") }
            elseif ($t -cmatch '^\d') {
                if ($t -ceq '1') { [void]$found.Add('the least size') }
                if ($t -ceq '409.55') { [void]$found.Add('the most size') }
                if ($t -cmatch '^\d+$') { [void]$found.Add('a whole size') }
                elseif ($t -cmatch '\.5$') { [void]$found.Add('a size of a half') }
                elseif ($t -cmatch '\.\d$') { [void]$found.Add('a size of tenths') }
                elseif ($t -cmatch '\.\d\d$') { [void]$found.Add('a size of hundredths') }
            } else { [void]$found.Add($t) }
            continue
        }
        if ($depth -eq 1 -and $head -ceq 'align') { [void]$found.Add($t); continue }
        if ($depth -eq 2 -and $head -ceq 'border') {
            $edgeItem++
            switch ($edgeItem) {
                1 { [void]$found.Add("the $t edge") }
                2 { [void]$found.Add($t) }
                3 { [void]$found.Add("an edge's color") }
            }
        }
    }
    if ($groups.Contains('font') -and $groups.Contains('align') -and $groups.Contains('border')) {
        [void]$found.Add('the three groups')
    }
    return ,$found
}

# Every problem of one fixture/golden/catalogue set, with the counts the
# floors read.
function Test-AppearanceGolden([string]$fixture, [string]$golden, [string]$catalog) {
    $problems = New-Object System.Collections.ArrayList
    $fcases = Read-AppearanceCases $fixture
    $gcases = Read-AppearanceCases $golden

    # 4. the refusal's id
    $ids = New-OrdinalSet
    foreach ($ln in (Get-AppearanceLines $catalog)) {
        if ($ln -match '^\(message (\S+) ') { [void]$ids.Add($Matches[1]) }
    }
    if (-not $ids.Contains('grid-format-property')) { [void]$problems.Add("grid-format-property is not an id of $catalog") }

    # 6. the written workbook
    if (-not (Test-Path -LiteralPath $writtenPath)) {
        [void]$problems.Add("scripts/appearance_written.xlsx is missing: FRAZARO_BLESS=1 and frazaro-core's tests write it")
    }

    # 1. the headers
    if ($fcases.Count -ne $gcases.Count) {
        [void]$problems.Add(("the fixture has {0} cases and the golden {1}" -f $fcases.Count, $gcases.Count))
    }
    $n = [Math]::Min($fcases.Count, $gcases.Count)
    for ($i = 0; $i -lt $n; $i++) {
        if ($fcases[$i].Header -cne $gcases[$i].Header) {
            [void]$problems.Add(("case {0}: the fixture says '{1}', the golden '{2}'" -f ($i + 1), $fcases[$i].Header, $gcases[$i].Header))
        }
    }

    # 2, 3 and 5. the records, the readings and the properties
    $reached = New-OrdinalSet
    $refused = 0
    $excelRefused = 0
    $readings = 0
    for ($i = 0; $i -lt $n; $i++) {
        $f = $fcases[$i]
        $g = $gcases[$i]
        if ($f.Lines.Count -ne 1) { [void]$problems.Add(("'{0}': {1} lines after its header, where a case is one" -f $f.Header, $f.Lines.Count)); continue }
        $ours = $f.Header -match ' !refused '
        if ($ours) { $refused++ } else {
            foreach ($x in (Get-CaseProperties ([string]$f.Lines[0]))) { [void]$reached.Add($x) }
        }
        if ($g.Lines.Count -ne 2) { [void]$problems.Add(("'{0}': {1} records, where a case has EXCEL and WRITTEN" -f $g.Header, $g.Lines.Count)); continue }
        $excel = [string]$g.Lines[0]
        $written = [string]$g.Lines[1]
        if ($excel -cmatch '^EXCEL\tREFUSED \d+$') {
            $excelRefused++
        } elseif ($excel -cmatch '^EXCEL\t(.*)$') {
            foreach ($p in (Test-Reading $g.Header 'EXCEL' $Matches[1])) { [void]$problems.Add($p) }
            $readings++
        } else {
            [void]$problems.Add(("'{0}': a first record that is neither EXCEL<TAB><reading> nor EXCEL<TAB>REFUSED <number>: '{1}'" -f $g.Header, $excel))
        }
        if ($written -ceq "WRITTEN`tREFUSED") {
            if (-not $ours) { [void]$problems.Add(("'{0}': WRITTEN reads REFUSED, but the language takes the case" -f $g.Header)) }
        } elseif ($written -cmatch '^WRITTEN\t(.*)$') {
            if ($ours) { [void]$problems.Add(("'{0}': the language refuses the case, but WRITTEN reads a cell" -f $g.Header)) }
            foreach ($p in (Test-Reading $g.Header 'WRITTEN' $Matches[1])) { [void]$problems.Add($p) }
            $readings++
        } else {
            [void]$problems.Add(("'{0}': a second record that is neither WRITTEN<TAB><reading> nor WRITTEN<TAB>REFUSED: '{1}'" -f $g.Header, $written))
        }
    }
    foreach ($x in $properties) {
        if (-not $reached.Contains($x)) { [void]$problems.Add("no taken case reaches $x") }
    }

    # 7. the floors
    if ($gcases.Count -lt $caseFloor) { [void]$problems.Add(("{0} cases, below the floor of {1}" -f $gcases.Count, $caseFloor)) }
    if ($refused -lt $refusedFloor) { [void]$problems.Add(("{0} cases refused here, below the floor of {1}" -f $refused, $refusedFloor)) }

    return @{ Problems = $problems; Cases = $gcases.Count; Refused = $refused; ExcelRefused = $excelRefused; Readings = $readings }
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_appearance_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $real = Test-AppearanceGolden $fixturePath $goldenPath $catalogPath
        $lines = Get-AppearanceLines $goldenPath
        $firstWritten = -1
        $firstBold = -1
        $lastHeader = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($firstWritten -lt 0 -and $lines[$i].StartsWith("WRITTEN`t") -and $lines[$i] -cne "WRITTEN`tREFUSED") { $firstWritten = $i }
            if ($firstBold -lt 0 -and $lines[$i] -cmatch '^EXCEL\t.* \(font bold\)$') { $firstBold = $i }
            if ($lines[$i].StartsWith('=== ')) { $lastHeader = $i }
        }
        if ($firstWritten -lt 0 -or $firstBold -lt 0 -or $lastHeader -lt 1) {
            Write-Output 'FAIL: the control found no WRITTEN reading, no bold reading or no last header to mutate'
            exit 1
        }
        $mutants = New-Object System.Collections.ArrayList
        # (a) one WRITTEN tag misspelled
        $a = @($lines | ForEach-Object { $_ })
        $a[$firstWritten] = 'WRITEN' + $a[$firstWritten].Substring(7)
        [void]$mutants.Add(@('the misspelt tag', $a, $null))
        # (b) the last case dropped
        [void]$mutants.Add(@('the dropped case', @($lines[0..($lastHeader - 1)]), $null))
        # (c) a word the record has none for
        $c = @($lines | ForEach-Object { $_ })
        $c[$firstBold] = $c[$firstBold].Replace('(font bold)', '(font heavy)')
        [void]$mutants.Add(@('the word of none', $c, $null))
        # (d) a default said
        $d = @($lines | ForEach-Object { $_ })
        $d[$firstBold] = $d[$firstBold].Replace('(font bold)', '(font bold 11)')
        [void]$mutants.Add(@('the default said', $d, $null))
        # (e) REFUSED where the language takes the case
        $e = @($lines | ForEach-Object { $_ })
        $e[$firstWritten] = "WRITTEN`tREFUSED"
        [void]$mutants.Add(@('the refusal of a taken case', $e, $null))
        # (f) the fixture's slant-dash-dot case saying thin
        $fixText = [System.IO.File]::ReadAllText($fixturePath)
        if ($fixText.IndexOf('(left slant-dash-dot)') -lt 0) { Write-Output 'FAIL: the control found no slant-dash-dot case'; exit 1 }
        [void]$mutants.Add(@('the line not reached', $null, $fixText.Replace('(left slant-dash-dot)', '(left thin)')))
        $counts = New-Object System.Collections.ArrayList
        $allFail = $true
        $k = 0
        foreach ($mu in $mutants) {
            $k++
            $golden = $goldenPath
            $fixture = $fixturePath
            if ($null -ne $mu[1]) {
                $golden = Join-Path $tmp ("golden_$k.txt")
                [System.IO.File]::WriteAllText($golden, ($mu[1] -join "`n"), $utf8)
            }
            if ($null -ne $mu[2]) {
                $fixture = Join-Path $tmp ("fixture_$k.txt")
                [System.IO.File]::WriteAllText($fixture, $mu[2], $utf8)
            }
            $r = Test-AppearanceGolden $fixture $golden $catalogPath
            [void]$counts.Add(("{0} {1}" -f $mu[0], $r.Problems.Count))
            if ($r.Problems.Count -lt 1) { $allFail = $false }
        }
        Write-Output ("control: the real set has {0} problem(s); {1}" -f $real.Problems.Count, ($counts -join ', '))
        if ($real.Problems.Count -eq 0 -and $allFail) {
            Write-Output 'OK: the check passes the golden and fails each mutant'
            exit 0
        }
        Write-Output 'FAIL: the control did not behave'
        foreach ($p in $real.Problems) { Write-Output ('  real: ' + $p) }
        exit 1
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}

$r = Test-AppearanceGolden $fixturePath $goldenPath $catalogPath
if ($r.Problems.Count -eq 0) {
    $above = ($r.Cases -gt $caseFloor) -or ($r.Refused -gt $refusedFloor)
    $note = if ($above) { ' - above a floor: raise the floors in this file' } else { '' }
    Write-Output ("OK: {0} cases, {1} refused here, {2} readings in the record's spelling, {3} cases Excel's object model refused; every one of the {4} properties reached (floors {5}, {6}){7}" -f $r.Cases, $r.Refused, $r.Readings, $r.ExcelRefused, $properties.Count, $caseFloor, $refusedFloor, $note)
    exit 0
}
Write-Output ("FAIL: the appearance golden has {0} problem(s):" -f $r.Problems.Count)
foreach ($p in $r.Problems) { Write-Output ('  ' + $p) }
exit 1
