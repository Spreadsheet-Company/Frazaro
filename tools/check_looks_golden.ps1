<#
check_looks_golden.ps1 - the looks golden is Excel's reading of its two
workbooks in the record's spelling, every target of a kind this version
reads, and its cases reach every kind of look the reader's rules name.

WHY: KERNEL.57 (a file's looks as rows) reads a package's looks as the
record's rows, and Excel is the reference for each rule: frazaro-core's
test-only builder (core/src/reflect/looks_golden.rs) writes
scripts/reflect/looks.xlsx and looks_book.xlsx, a case a row, what neither
Excel's writer nor the core's writes, and VlaWriteLooksGolden
(VLA_Tests.bas) reads each case's targets through Excel's object model into
scripts/reflect/looks_golden.txt, saving each workbook beside it as Excel's
own. The core's test holds the reader's reading of each workbook, and of
Excel's saves, to the golden, target by target, and the golden's cases to
the builder's. That test proves agreement on what the golden says. It
cannot see a golden respelled by hand, a target of a kind nobody reads, a
look the cases stopped reaching, or a golden left the prediction. This
check holds those, host-free, on every push:

  1. the golden's first line names its source, Excel's (excel <version>
     <build>) or the prediction written before Excel's run, and the two
     workbooks it reads are there, with Excel's saves of them once it is
     Excel's;
  2. every other line is a case's header, "=== <file>!<sheet>!<name>" of
     one of the two workbooks, or a target of the case above it,
     <kind>:<place><TAB><reading>, its kind one this version reads;
  3. every reading is in its kind's spelling: a cell's look in the record's
     spelling, each property said only where it is not the workbook's
     default (no not-bold, not-italic, the default's size, general, bottom
     or none) and no group empty, then " !" and words of what the record
     has no word for, each a word this version names; a theme's slot and
     tint; a width or a height and shown or hidden; a merge's range; the
     rules of a conditional format and their fill; on or off; a standard
     width and height;
  4. the cases reach, each named below: the twelve theme slots, untinted
     and tinted; every index of the default palette and the two past it;
     every word of what the record has no word for; every kind of target;
  5. the floors below: the cases and the targets.

KERNEL.58's second golden, scripts/reflect/rules_golden.txt, is Excel's
reading of a conditional format, VlaWriteRulesGolden's of rules.xlsx,
rules_plain.xlsx and rules_x14.xlsx (core/src/reflect/rules_golden.rs), and
is held the same way: its source and its workbooks, with Excel's two saves of
each once it is Excel's; its headers and targets; a shown: target in the
looks' spelling, a text: target a quoted text, and a rule: target "none" or
where the rule applies, its kind, its operator, its priority, stop or go and
at most two quoted formulas; no target Excel refused, each a "!error"; every
sheet of the probes, every kind of rule and every operator reached; and its
floors. The core's test holds its cases to the builder's.

Its third, scripts/reflect/rules_written_golden.txt, is the same of the writer's
own workbook, VlaWriteRulesWrittenGolden's of rules_written.xlsx, with Excel's
save of it once it is Excel's, and is held by the same rules: its one sheet, the
kinds of rule the writer writes, every operator, a merge's target, each of its
readings a range, and a held target, each of its readings a row's words for a kind
the language does not compute, which no other golden reads, and its floors.

-Control proves the check on scratch copies: the real golden must pass,
and each of these must fail: one with a target's kind misspelled, one with
its last case dropped, one reading a word the record has none for, one
saying looks.xlsx's default size, one naming an unheld word of no such
name, and one whose source is no one's; and of the rules golden, one with a
target's kind misspelled, one with its last case dropped, one reading a word
the record has none for, one reading a rule of no such kind, one reading an
error, one whose source is no one's, and one with a merge's target, which it
has none of; and of the writer's, one with its last case dropped, one whose
source is no one's, one reading a merge that is no range, one with no merge's
target, one reading a held target of no kind, and one with no held target.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; exit 0 clean, exit 1 with every problem named.

Usage:  powershell -File tools\check_looks_golden.ps1
        powershell -File tools\check_looks_golden.ps1 -Control
#>
param(
    [switch]$Control
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# --- the floors, never lowered ------------------------------------------------
# 2026-10-22, KERNEL.57: 389 cases, 912 targets, the golden written as the
# prediction ahead of the reference's first run in Excel.
# 2026-10-22, the fold: 393 cases, 928 targets, four tints of a tint's
# quantum added for the second run.
$caseFloor = 393
$targetFloor = 928

$goldenPath = Join-Path $repoRoot 'scripts/reflect/looks_golden.txt'
$books = @('looks.xlsx', 'looks_book.xlsx')
$saves = @('looks_saved.xlsx', 'looks_book_saved.xlsx')
# Each workbook's default size, which no reading of its cells says.
$defaultSize = @{ 'looks.xlsx' = '11'; 'looks_book.xlsx' = '10' }

$kinds = @('cell', 'empty', 'blank', 'theme', 'col', 'row', 'merge', 'cf', 'grid', 'std')

# The rules golden's floors, never lowered.
# 2026-10-30, KERNEL.58's third slice: 195 cases, 998 targets, each rule a
# target's a line, the golden the reference's second run in Excel.
$rulesCaseFloor = 195
$rulesTargetFloor = 998

$rulesGoldenPath = Join-Path $repoRoot 'scripts/reflect/rules_golden.txt'
$rulesBooks = @('rules.xlsx', 'rules_plain.xlsx', 'rules_x14.xlsx')
$rulesSaves = @('rules_saved.xlsx', 'rules_plain_saved.xlsx', 'rules_x14_saved.xlsx',
    'rules_saved.ods', 'rules_plain_saved.ods', 'rules_x14_saved.ods')
$rulesKinds = @('shown', 'text', 'rule')
$rulesSheets = @('Truth', 'Cell', 'Refs', 'Bases', 'Edges', 'Order', 'Alike', 'Partial', 'Kinds', 'Names', 'Elsewhere', 'Rules')
$ruleKinds = @('cell', 'expression', 'color-scale', 'data-bar', 'top10', 'icon-set', 'unique-values', 'text', 'blanks',
    'time-period', 'above-average', 'no-blanks', 'errors', 'no-errors')
$ruleOperators = @('between', 'not-between', 'equal', 'not-equal', 'greater', 'less', 'greater-or-equal', 'less-or-equal')
$rulesDefaultSize = '11'
$rulesSpec = @{
    Name = 'rules'; Writer = 'VlaWriteRulesGolden'
    Books = $rulesBooks; Saves = $rulesSaves; Sheets = $rulesSheets; Kinds = $rulesKinds
    RuleKinds = $ruleKinds; Operators = $ruleOperators
    CaseFloor = $rulesCaseFloor; TargetFloor = $rulesTargetFloor
}

# The writer's own workbook (the item's fourth slice): every form it writes. Its
# floors, never lowered: 2026-11-02, 162 cases, 617 targets, each rule a target's a
# line, the merged ranges of the item's sixth slice and the kinds not computed of its
# seventh among them.
$writtenGoldenPath = Join-Path $repoRoot 'scripts/reflect/rules_written_golden.txt'
$writtenSpec = @{
    Name = 'written'; Writer = 'VlaWriteRulesWrittenGolden'
    Books = @('rules_written.xlsx'); Saves = @('rules_written_saved.xlsx'); Sheets = @('Written'); Kinds = ($rulesKinds + 'merge' + 'held')
    RuleKinds = @('cell', 'expression', 'color-scale', 'data-bar', 'icon-set', 'top10', 'above-average', 'unique-values'); Operators = $ruleOperators
    CaseFloor = 162; TargetFloor = 617
}
$words = @('superscript', 'subscript', 'outline', 'shadow',
    'runs', 'pattern', 'horizontal', 'vertical', 'indent', 'rotation', 'shrink', 'rtl', 'diagonal')

# A cell's look in the record's spelling (rule 3), every group optional and
# in the record's order.
$rgb = '"[0-9A-F]{6}"'
$sizeText = '[1-9][0-9]{0,2}(?:\.[1-9]|\.[0-9]5)?|size=\S+'
$lineWords = '(?:medium-dash-dot-dot|medium-dash-dot|medium-dashed|slant-dash-dot|dash-dot-dot|dash-dot|dashed|dotted|double|medium|thick|thin|hair|style=-?\d+ weight=-?\d+)'
$edgeText = { param($side) "(?: \($side $lineWords(?: $rgb)?\))?" }
$lookPattern = '^(?:none|' + $rgb + ') (?:general|text|"(?:[^"\\]|\\.)*") (?:wrap|nowrap)' +
    '(?: \(font(?: (?:bold|not-bold))?(?: (?:italic|not-italic))?(?: (?<size>' + $sizeText + '))?(?: ' + $rgb + ')?(?: (?:underline|double-underline|accounting-underline|double-accounting-underline))?(?: strike)?(?: \(family "(?:[^"\\]|\\.)*"\))?\))?' +
    '(?: \(align(?: (?:left|center|right))?(?: (?:top|middle))?\))?' +
    '(?: \(border' + (& $edgeText 'left') + (& $edgeText 'right') + (& $edgeText 'top') + (& $edgeText 'bottom') + '\))?$'
$wordPattern = '^(?:superscript|subscript|outline|shadow|runs|pattern -?\d+|horizontal -?\d+|vertical -?\d+|indent \d+|rotation -?\d+|shrink|rtl|diagonal)$'
$number = '-?\d+(?:\.\d+)?(?:E[-+]?\d+)?'
$readingPatterns = @{
    'theme' = "^(?:none|\d+ $number)$"
    'col'   = "^$number (?:shown|hidden)$"
    'row'   = "^$number (?:shown|hidden)$"
    'merge' = '^[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?$'
    'cf'    = "^\d+ (?:none|$rgb)$"
    'grid'  = '^(?:on|off)$'
    'std'   = "^$number $number$"
}

$quotedText = '"(?:[^"\\]|\\.)*"'
$sqrefText = '[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?(?:,[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?)*'
$rulePattern = "^$sqrefText (?<kind>[a-z0-9-]+) (?<operator>-|[a-z-]+) \d+ (?<stop>stop|go)(?: $quotedText){0,2}$"
$textPattern = "^$quotedText$"

function Get-GoldenLines([string]$path) {
    $text = [System.IO.File]::ReadAllText($path)
    $text = $text -replace "`r`n", "`n" -replace "`r", "`n"
    if ($text.EndsWith("`n")) { $text = $text.Substring(0, $text.Length - 1) }
    return ,@($text -split "`n")
}

function New-OrdinalSet {
    return ,(New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal))
}

# The words of what the record has no word for, after " !": each a word of
# the list and the value after it where it takes one.
function Split-Words([string]$tail) {
    $found = New-Object System.Collections.ArrayList
    $m = [regex]::Matches($tail, ' (pattern -?\d+|horizontal -?\d+|vertical -?\d+|indent \d+|rotation -?\d+|[a-z]+)')
    $at = 0
    foreach ($x in $m) {
        if ($x.Index -ne $at) { [void]$found.Add($null); return ,$found }
        [void]$found.Add($x.Groups[1].Value)
        $at = $x.Index + $x.Length
    }
    if ($at -ne $tail.Length) { [void]$found.Add($null) }
    return ,$found
}

# Every problem of one golden, with the counts the floors read.
function Test-LooksGolden([string]$golden) {
    $problems = New-Object System.Collections.ArrayList
    $lines = Get-GoldenLines $golden
    $reached = New-OrdinalSet
    $cases = 0
    $targets = 0

    # 1. the source and the workbooks
    $source = ''
    if ($lines.Count -gt 0 -and $lines[0] -cmatch '^source\t(excel \S+ \d+|prediction)$') {
        $source = $Matches[1]
    } else {
        [void]$problems.Add('the first line names no source: source<TAB>excel <version> <build>, or source<TAB>prediction')
    }
    foreach ($b in $books) {
        if (-not (Test-Path -LiteralPath (Join-Path $repoRoot "scripts/reflect/$b"))) {
            [void]$problems.Add("scripts/reflect/$b is missing: FRAZARO_BLESS=1 and frazaro-core's tests write it")
        }
    }
    if ($source.StartsWith('excel ')) {
        foreach ($b in $saves) {
            if (-not (Test-Path -LiteralPath (Join-Path $repoRoot "scripts/reflect/$b"))) {
                [void]$problems.Add("scripts/reflect/$b is missing: VlaWriteLooksGolden saves it beside the golden it writes")
            }
        }
    }

    # 2 and 3. the cases, their targets and the readings
    $file = $null
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $ln = $lines[$i]
        if ($ln.StartsWith('=== ')) {
            $parts = $ln.Substring(4).Split([char[]]@('!'), 3)
            if ($parts.Count -ne 3 -or ($books -notcontains $parts[0]) -or $parts[1] -eq '' -or $parts[2] -eq '') {
                [void]$problems.Add(("line {0}: a header that is no '=== <file>!<sheet>!<name>' of the two workbooks: '{1}'" -f ($i + 1), $ln))
                $file = $null
            } else {
                $file = $parts[0]
                $cases++
                if ($parts[2] -cmatch '^theme (\d+) ') {
                    $slot = $Matches[1]
                    if ($parts[2].Contains(' tint ')) { [void]$reached.Add("theme $slot tinted") } else { [void]$reached.Add("theme $slot") }
                }
                if ($file -ceq 'looks.xlsx' -and $parts[2] -cmatch '^indexed (\d+)$') { [void]$reached.Add("indexed $($Matches[1])") }
            }
            continue
        }
        if ($null -eq $file) {
            [void]$problems.Add(("line {0}: a target before any case: '{1}'" -f ($i + 1), $ln))
            continue
        }
        $tab = $ln.IndexOf("`t")
        if ($tab -lt 0) {
            [void]$problems.Add(("line {0}: neither a header nor <target><TAB><reading>: '{1}'" -f ($i + 1), $ln))
            continue
        }
        $target = $ln.Substring(0, $tab)
        $reading = $ln.Substring($tab + 1)
        $kind = $target.Split([char]':')[0]
        if ($kinds -notcontains $kind) {
            [void]$problems.Add(("line {0}: a target of no kind this version reads: '{1}'" -f ($i + 1), $target))
            continue
        }
        $targets++
        [void]$reached.Add("the $kind target")
        if ($kind -ceq 'cell' -or $kind -ceq 'empty' -or $kind -ceq 'blank') {
            $look = $reading
            $tail = ''
            $bang = $reading.IndexOf(' !')
            if ($bang -ge 0) { $look = $reading.Substring(0, $bang); $tail = $reading.Substring($bang + 2) }
            $m = [regex]::Match($look, $lookPattern)
            if (-not $m.Success) {
                [void]$problems.Add(("line {0}: {1} reads '{2}', not in the record's spelling" -f ($i + 1), $target, $look))
            } else {
                if ($look -cmatch '\((font|align|border)\)') { [void]$problems.Add(("line {0}: {1} reads an empty group: '{2}'" -f ($i + 1), $target, $look)) }
                if ($m.Groups['size'].Success -and $m.Groups['size'].Value -ceq $defaultSize[$file]) {
                    [void]$problems.Add(("line {0}: {1} reads {2}'s default size, {3}: '{4}'" -f ($i + 1), $target, $file, $defaultSize[$file], $look))
                }
            }
            if ($bang -ge 0) {
                $said = Split-Words $tail
                if ($said.Count -eq 0) { [void]$problems.Add(("line {0}: {1} reads ' !' and no word" -f ($i + 1), $target)) }
                foreach ($w in $said) {
                    if ($null -eq $w -or -not ($w -cmatch $wordPattern)) {
                        [void]$problems.Add(("line {0}: {1} reads a word of what the record cannot hold that is none this version names: '{2}'" -f ($i + 1), $target, $tail))
                        break
                    }
                    [void]$reached.Add("the word " + $w.Split([char]' ')[0])
                }
            }
        } elseif (-not ($reading -cmatch $readingPatterns[$kind])) {
            [void]$problems.Add(("line {0}: {1} reads '{2}', not a {3} target's spelling" -f ($i + 1), $target, $reading, $kind))
        }
    }

    # 4. what the cases reach
    $want = New-Object System.Collections.ArrayList
    for ($s = 0; $s -lt 12; $s++) { [void]$want.Add("theme $s"); [void]$want.Add("theme $s tinted") }
    for ($x = 0; $x -lt 66; $x++) { [void]$want.Add("indexed $x") }
    foreach ($w in $words) { [void]$want.Add("the word $w") }
    foreach ($k in $kinds) { [void]$want.Add("the $k target") }
    foreach ($x in $want) {
        if (-not $reached.Contains($x)) { [void]$problems.Add("no case reaches $x") }
    }

    # 5. the floors
    if ($cases -lt $caseFloor) { [void]$problems.Add(("{0} cases, below the floor of {1}" -f $cases, $caseFloor)) }
    if ($targets -lt $targetFloor) { [void]$problems.Add(("{0} targets, below the floor of {1}" -f $targets, $targetFloor)) }

    return @{ Problems = $problems; Cases = $cases; Targets = $targets; Source = $source; Reached = $want.Count }
}

# Every problem of the rules golden, with the counts the floors read.
function Test-RulesGolden([string]$golden, [hashtable]$spec) {
    $problems = New-Object System.Collections.ArrayList
    $lines = Get-GoldenLines $golden
    $reached = New-OrdinalSet
    $cases = 0
    $targets = 0

    # 1. the source and the workbooks
    $source = ''
    if ($lines.Count -gt 0 -and $lines[0] -cmatch '^source\t(excel \S+ \d+|prediction)$') {
        $source = $Matches[1]
    } else {
        [void]$problems.Add('the rules golden''s first line names no source: source<TAB>excel <version> <build>, or source<TAB>prediction')
    }
    foreach ($b in $spec.Books) {
        if (-not (Test-Path -LiteralPath (Join-Path $repoRoot "scripts/reflect/$b"))) {
            [void]$problems.Add("scripts/reflect/$b is missing: FRAZARO_BLESS=1 and frazaro-core's tests write it")
        }
    }
    if ($source.StartsWith('excel ')) {
        foreach ($b in $spec.Saves) {
            if (-not (Test-Path -LiteralPath (Join-Path $repoRoot "scripts/reflect/$b"))) {
                [void]$problems.Add("scripts/reflect/$b is missing: $($spec.Writer) saves it beside the golden it writes")
            }
        }
    }

    # 2 and 3. the cases, their targets and the readings
    $open = $false
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $ln = $lines[$i]
        if ($ln.StartsWith('=== ')) {
            $parts = $ln.Substring(4).Split([char[]]@('!'), 3)
            if ($parts.Count -ne 3 -or ($spec.Books -notcontains $parts[0]) -or $parts[1] -eq '' -or $parts[2] -eq '') {
                [void]$problems.Add(("rules line {0}: a header that is no '=== <file>!<sheet>!<name>' of its workbooks: '{1}'" -f ($i + 1), $ln))
                $open = $false
            } else {
                $open = $true
                $cases++
                [void]$reached.Add("the sheet $($parts[1])")
            }
            continue
        }
        if (-not $open) {
            [void]$problems.Add(("rules line {0}: a target before any case: '{1}'" -f ($i + 1), $ln))
            continue
        }
        $tab = $ln.IndexOf("`t")
        if ($tab -lt 0) {
            [void]$problems.Add(("rules line {0}: neither a header nor <target><TAB><reading>: '{1}'" -f ($i + 1), $ln))
            continue
        }
        $target = $ln.Substring(0, $tab)
        $reading = $ln.Substring($tab + 1)
        $kind = $target.Split([char]':')[0]
        if ($spec.Kinds -notcontains $kind) {
            [void]$problems.Add(("rules line {0}: a target of no kind this version reads: '{1}'" -f ($i + 1), $target))
            continue
        }
        $targets++
        [void]$reached.Add("the $kind target")
        if ($reading.StartsWith('!error')) {
            [void]$problems.Add(("rules line {0}: Excel refused {1}: '{2}'" -f ($i + 1), $target, $reading))
            continue
        }
        if ($kind -ceq 'shown') {
            $look = $reading
            $tail = ''
            $bang = $reading.IndexOf(' !')
            if ($bang -ge 0) { $look = $reading.Substring(0, $bang); $tail = $reading.Substring($bang + 2) }
            $m = [regex]::Match($look, $lookPattern)
            if (-not $m.Success) {
                [void]$problems.Add(("rules line {0}: {1} reads '{2}', not in the record's spelling" -f ($i + 1), $target, $look))
            } else {
                if ($look -cmatch '\((font|align|border)\)') { [void]$problems.Add(("rules line {0}: {1} reads an empty group: '{2}'" -f ($i + 1), $target, $look)) }
                if ($m.Groups['size'].Success -and $m.Groups['size'].Value -ceq $rulesDefaultSize) {
                    [void]$problems.Add(("rules line {0}: {1} reads the default size: '{2}'" -f ($i + 1), $target, $look))
                }
            }
            if ($bang -ge 0) {
                foreach ($w in (Split-Words $tail)) {
                    if ($null -eq $w -or -not ($w -cmatch $wordPattern)) {
                        [void]$problems.Add(("rules line {0}: {1} reads a word of what the record cannot hold that is none this version names: '{2}'" -f ($i + 1), $target, $tail))
                        break
                    }
                }
            }
        } elseif ($kind -ceq 'text') {
            if (-not ($reading -cmatch $textPattern)) {
                [void]$problems.Add(("rules line {0}: {1} reads '{2}', not a quoted text" -f ($i + 1), $target, $reading))
            }
        } elseif ($kind -ceq 'merge') {
            if (-not ($reading -cmatch '^[A-Z]+\d+(:[A-Z]+\d+)?$')) {
                [void]$problems.Add(("rules line {0}: {1} reads '{2}', not a range" -f ($i + 1), $target, $reading))
            }
        } elseif ($kind -ceq 'held') {
            if (-not ($reading -cmatch '^(none|\((color-scale|data-bar|icon-set|top|bottom|above-average|below-average|duplicate|unique)\b.*\))$')) {
                [void]$problems.Add(("rules line {0}: {1} reads '{2}', no row's words" -f ($i + 1), $target, $reading))
            }
        } elseif ($reading -cne 'none') {
            $m = [regex]::Match($reading, $rulePattern)
            if (-not $m.Success) {
                [void]$problems.Add(("rules line {0}: {1} reads '{2}', not a rule's spelling" -f ($i + 1), $target, $reading))
            } else {
                $word = $m.Groups['kind'].Value
                if ($spec.RuleKinds -cnotcontains $word) {
                    [void]$problems.Add(("rules line {0}: {1} reads a rule of no such kind: '{2}'" -f ($i + 1), $target, $word))
                }
                [void]$reached.Add("the rule kind $word")
                $operatorWord = $m.Groups['operator'].Value
                if ($operatorWord -cne '-') {
                    if ($spec.Operators -cnotcontains $operatorWord) {
                        [void]$problems.Add(("rules line {0}: {1} reads an operator of no such word: '{2}'" -f ($i + 1), $target, $operatorWord))
                    }
                    [void]$reached.Add("the operator $operatorWord")
                }
                [void]$reached.Add("a rule that says $($m.Groups['stop'].Value)")
            }
        }
    }

    # 4. what the cases reach
    $want = New-Object System.Collections.ArrayList
    foreach ($s in $spec.Sheets) { [void]$want.Add("the sheet $s") }
    foreach ($k in $spec.Kinds) { [void]$want.Add("the $k target") }
    foreach ($k in $spec.RuleKinds) { [void]$want.Add("the rule kind $k") }
    foreach ($o in $spec.Operators) { [void]$want.Add("the operator $o") }
    [void]$want.Add('a rule that says stop')
    [void]$want.Add('a rule that says go')
    foreach ($x in $want) {
        if (-not $reached.Contains($x)) { [void]$problems.Add("no rules case reaches $x") }
    }

    # 5. the floors
    if ($cases -lt $spec.CaseFloor) { [void]$problems.Add(("{0} {1} cases, below the floor of {2}" -f $cases, $spec.Name, $spec.CaseFloor)) }
    if ($targets -lt $spec.TargetFloor) { [void]$problems.Add(("{0} {1} targets, below the floor of {2}" -f $targets, $spec.Name, $spec.TargetFloor)) }

    return @{ Problems = $problems; Cases = $cases; Targets = $targets; Source = $source; Reached = $want.Count }
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_looks_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $real = Test-LooksGolden $goldenPath
        $lines = Get-GoldenLines $goldenPath
        $firstCell = -1
        $firstWords = -1
        $lastHeader = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($firstCell -lt 0 -and $lines[$i] -cmatch '^cell:[A-Z]+\d+\tnone general nowrap$') { $firstCell = $i }
            if ($firstWords -lt 0 -and $lines[$i] -cmatch '^cell:[A-Z]+\d+\t.* ! outline$') { $firstWords = $i }
            if ($lines[$i].StartsWith('=== ')) { $lastHeader = $i }
        }
        if ($firstCell -lt 0 -or $firstWords -lt 0 -or $lastHeader -lt 1) {
            Write-Output 'FAIL: the control found no plain cell, no outlined cell or no last header to mutate'
            exit 1
        }
        $mutants = New-Object System.Collections.ArrayList
        $copy = { @($lines | ForEach-Object { $_ }) }
        # (a) a target's kind misspelled
        $a = & $copy
        $a[$firstCell] = 'cel' + $a[$firstCell].Substring(4)
        [void]$mutants.Add(@('the misspelt kind', $a))
        # (b) the last case dropped
        [void]$mutants.Add(@('the dropped case', @($lines[0..($lastHeader - 1)])))
        # (c) a word the record has none for
        $c = & $copy
        $c[$firstCell] = $c[$firstCell].Replace('none general nowrap', 'none general nowrap (font heavy)')
        [void]$mutants.Add(@('the word of none', $c))
        # (d) looks.xlsx's default size said
        $d = & $copy
        $d[$firstCell] = $d[$firstCell].Replace('none general nowrap', 'none general nowrap (font 11)')
        [void]$mutants.Add(@('the default said', $d))
        # (e) an unheld word of no such name
        $e = & $copy
        $e[$firstWords] = $e[$firstWords].Replace(' ! outline', ' ! outlined')
        [void]$mutants.Add(@('the unheld word of none', $e))
        # (f) a source that is no one's
        $f = & $copy
        $f[0] = "source`tguess"
        [void]$mutants.Add(@('the source of none', $f))
        $counts = New-Object System.Collections.ArrayList
        $allFail = $true
        $k = 0
        foreach ($mu in $mutants) {
            $k++
            $golden = Join-Path $tmp ("golden_$k.txt")
            [System.IO.File]::WriteAllText($golden, ($mu[1] -join "`n"), $utf8)
            $r = Test-LooksGolden $golden
            [void]$counts.Add(("{0} {1}" -f $mu[0], $r.Problems.Count))
            if ($r.Problems.Count -lt 1) { $allFail = $false }
        }
        Write-Output ("control: the real golden has {0} problem(s); {1}" -f $real.Problems.Count, ($counts -join ', '))

        # The rules golden's mutants.
        $rulesReal = Test-RulesGolden $rulesGoldenPath $rulesSpec
        $lines = Get-GoldenLines $rulesGoldenPath
        $firstShown = -1
        $firstRule = -1
        $lastHeader = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($firstShown -lt 0 -and $lines[$i] -cmatch '^shown:[A-Z]+\d+\tnone general nowrap$') { $firstShown = $i }
            if ($firstRule -lt 0 -and $lines[$i] -cmatch '^rule:[A-Z]+\d+\t.* expression ') { $firstRule = $i }
            if ($lines[$i].StartsWith('=== ')) { $lastHeader = $i }
        }
        if ($firstShown -lt 0 -or $firstRule -lt 0 -or $lastHeader -lt 1) {
            Write-Output 'FAIL: the control found no plain cell, no rule or no last header to mutate'
            exit 1
        }
        $mutants = New-Object System.Collections.ArrayList
        $copy = { @($lines | ForEach-Object { $_ }) }
        $a = & $copy
        $a[$firstShown] = 'show' + $a[$firstShown].Substring(5)
        [void]$mutants.Add(@('the misspelt kind', $a))
        [void]$mutants.Add(@('the dropped case', @($lines[0..($lastHeader - 1)])))
        $c = & $copy
        $c[$firstShown] = $c[$firstShown].Replace('none general nowrap', 'none general nowrap (font heavy)')
        [void]$mutants.Add(@('the word of none', $c))
        $d = & $copy
        $d[$firstRule] = $d[$firstRule].Replace(' expression ', ' expressions ')
        [void]$mutants.Add(@('the rule of no kind', $d))
        $e = & $copy
        $e[$firstRule] = $e[$firstRule].Substring(0, $e[$firstRule].IndexOf("`t") + 1) + '!error 1004 no such member'
        [void]$mutants.Add(@('the refused target', $e))
        $f = & $copy
        $f[0] = "source`tguess"
        [void]$mutants.Add(@('the source of none', $f))
        $h = & $copy
        $h[$firstShown] = 'merge' + $h[$firstShown].Substring(5, $h[$firstShown].IndexOf("`t") - 5) + "`tC1:E1"
        [void]$mutants.Add(@('the merge no rules golden reads', $h))
        $rulesCounts = New-Object System.Collections.ArrayList
        $rulesAllFail = $true
        $k = 0
        foreach ($mu in $mutants) {
            $k++
            $golden = Join-Path $tmp ("rules_$k.txt")
            [System.IO.File]::WriteAllText($golden, ($mu[1] -join "`n"), $utf8)
            $r = Test-RulesGolden $golden $rulesSpec
            [void]$rulesCounts.Add(("{0} {1}" -f $mu[0], $r.Problems.Count))
            if ($r.Problems.Count -lt 1) { $rulesAllFail = $false }
        }
        Write-Output ("control: the real rules golden has {0} problem(s); {1}" -f $rulesReal.Problems.Count, ($rulesCounts -join ', '))

        # The writer's golden's mutants.
        $writtenReal = Test-RulesGolden $writtenGoldenPath $writtenSpec
        $lines = Get-GoldenLines $writtenGoldenPath
        $lastHeader = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i].StartsWith('=== ')) { $lastHeader = $i }
        }
        $g = @($lines | ForEach-Object { $_ })
        $g[0] = "source`tguess"
        $firstMerge = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($firstMerge -lt 0 -and $lines[$i].StartsWith('merge:')) { $firstMerge = $i }
        }
        if ($firstMerge -lt 0) {
            Write-Output 'FAIL: the control found no merged range to mutate'
            exit 1
        }
        $h = @($lines | ForEach-Object { $_ })
        $h[$firstMerge] = $h[$firstMerge].Substring(0, $h[$firstMerge].IndexOf("`t")) + "`tnope"
        $noMerge = @($lines | Where-Object { -not $_.StartsWith('merge:') })
        $firstHeld = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($firstHeld -lt 0 -and $lines[$i].StartsWith('held:')) { $firstHeld = $i }
        }
        if ($firstHeld -lt 0) {
            Write-Output 'FAIL: the control found no held target to mutate'
            exit 1
        }
        $j = @($lines | ForEach-Object { $_ })
        $j[$firstHeld] = $j[$firstHeld].Substring(0, $j[$firstHeld].IndexOf("`t")) + "`t(heat-map)"
        $noHeld = @($lines | Where-Object { -not $_.StartsWith('held:') })
        $writtenCounts = New-Object System.Collections.ArrayList
        $writtenAllFail = $true
        $k = 0
        foreach ($mu in @(@('the dropped case', @($lines[0..($lastHeader - 1)])), @('the source of none', $g), @('the range of none', $h), @('the merge of none', $noMerge), @('the words of none', $j), @('the held of none', $noHeld))) {
            $k++
            $golden = Join-Path $tmp ("written_$k.txt")
            [System.IO.File]::WriteAllText($golden, ($mu[1] -join "`n"), $utf8)
            $r = Test-RulesGolden $golden $writtenSpec
            [void]$writtenCounts.Add(("{0} {1}" -f $mu[0], $r.Problems.Count))
            if ($r.Problems.Count -lt 1) { $writtenAllFail = $false }
        }
        Write-Output ("control: the real written golden has {0} problem(s); {1}" -f $writtenReal.Problems.Count, ($writtenCounts -join ', '))
        if ($real.Problems.Count -eq 0 -and $allFail -and $rulesReal.Problems.Count -eq 0 -and $rulesAllFail -and $writtenReal.Problems.Count -eq 0 -and $writtenAllFail) {
            Write-Output 'OK: the check passes the three goldens and fails each mutant'
            exit 0
        }
        Write-Output 'FAIL: the control did not behave'
        foreach ($p in $real.Problems) { Write-Output ('  real: ' + $p) }
        foreach ($p in $rulesReal.Problems) { Write-Output ('  real rules: ' + $p) }
        foreach ($p in $writtenReal.Problems) { Write-Output ('  real written: ' + $p) }
        exit 1
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}

$r = Test-LooksGolden $goldenPath
$rules = Test-RulesGolden $rulesGoldenPath $rulesSpec
$written = Test-RulesGolden $writtenGoldenPath $writtenSpec
if ($r.Problems.Count -eq 0 -and $rules.Problems.Count -eq 0 -and $written.Problems.Count -eq 0) {
    $above = ($r.Cases -gt $caseFloor) -or ($r.Targets -gt $targetFloor)
    $note = ''
    if ($above) { $note = ' - above a floor: raise the floors in this file' }
    if ($r.Source -ceq 'prediction') { $note = $note + ' - the prediction, before the reference''s run in Excel' }
    if (($rules.Cases -gt $rulesCaseFloor) -or ($rules.Targets -gt $rulesTargetFloor)) { $note = $note + ' - above a rules floor: raise the floors in this file' }
    if ($rules.Source -ceq 'prediction') { $note = $note + ' - the rules golden is the prediction, before the reference''s run in Excel' }
    if (($written.Cases -gt $writtenSpec.CaseFloor) -or ($written.Targets -gt $writtenSpec.TargetFloor)) { $note = $note + ' - above a written floor: raise the floors in this file' }
    if ($written.Source -ceq 'prediction') { $note = $note + ' - the written golden is the prediction, before the reference''s run in Excel' }
    Write-Output ("OK: {0} cases, {1} targets in their kinds' spelling, from {2}; every one of the {3} looks and kinds reached (floors {4}, {5}); rules {6} cases, {7} targets, from {8}; every one of the {9} sheets, kinds and operators reached (floors {10}, {11}); the writer's {12} cases, {13} targets, from {14}; every one of its {15} reached (floors {16}, {17}){18}" -f $r.Cases, $r.Targets, $r.Source, $r.Reached, $caseFloor, $targetFloor, $rules.Cases, $rules.Targets, $rules.Source, $rules.Reached, $rulesCaseFloor, $rulesTargetFloor, $written.Cases, $written.Targets, $written.Source, $written.Reached, $writtenSpec.CaseFloor, $writtenSpec.TargetFloor, $note)
    exit 0
}
Write-Output ("FAIL: the looks golden has {0} problem(s), the rules golden {1}, the written golden {2}:" -f $r.Problems.Count, $rules.Problems.Count, $written.Problems.Count)
foreach ($p in $r.Problems) { Write-Output ('  ' + $p) }
foreach ($p in $rules.Problems) { Write-Output ('  ' + $p) }
foreach ($p in $written.Problems) { Write-Output ('  ' + $p) }
exit 1
