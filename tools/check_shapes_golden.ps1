<#
check_shapes_golden.ps1 - the door's census of each fixture is its golden
whole; each golden's line count never goes down; the census's order and its
totals hold in the golden itself.

WHY: KERNEL.30 prints the plan an engine's machine makes of a grid, its
census: `frazaro calc <file> --shapes` reads a workbook as `calc` reads it,
plans it as the machine plans one (vla-lang's Shapes::of, under Excel's
library) and prints a (plan ...) row of totals, a (readers ...) row for each
sheet whose changed cells are checked against readers, and a (shape ...) row
for each formula shape - its cells, the levels of the order it spans, its
runs and the path each of its reads takes - so that a model's cost is read
before anything computes it. These goldens are the door's and the crate's,
not an oracle of the treaty (conformance/README.md): the census is one
implementation's plan, and it changes on purpose whenever the plan does
(KERNEL.31 turned the reads of a sheet its plan keeps as an array, a range
of it or a run's moving cell, from `cells` and `lookup` into `array`,
`stride` and `slot`, the sums program's among them). Such a change
regenerates the goldens with the door, the diff showing what it moved, and
the floors move with it.

THE GOLDENS: scripts/recalc/shapes/<name>.vla, the door's own output saved
CRLF, over oracle 12's ten fixtures (the files check_recalc_golden.ps1
computes) and the five programs of scripts/recalc/bench/, which the door
builds first, as `frazaro build` builds any program, into target/shapes/;
the lookups' program is named lookups_bench, beside the lookups' fixture.

THE CACHE: a bench program's workbook and its census are kept under
target/shapes/ with the SHA-256 of the door and of the program they came
from. While both are byte for byte the ones held, the door, which reads no
clock and draws no chance, would print the same rows, and they are reused.
The first run after a door is built pays the build and the read: about a
minute with a debug door, most of it the 33 MB of Life's formulas.

ALSO HELD, read off each golden itself: its line count, as a floor; the
order, one plan row first with its nine counts in their order, then the
readers rows, then the shape rows by ascending number, every row's forms in
their order and its reads among the nine kinds in theirs; a readers row's
(most k), printed for a sheet whose readers are indexed (KERNEL.35), one
reader at least and the row's readers at most; the totals the
plan row states, summed from the shape rows (the shapes, their distinct
texts, the formula cells, the runs and their lanes, the cells alone, the
cells of the shapes computed every frame, and the cells on cycles, each
shape's cells beside its lanes and its cells alone); every (same k) naming
the first earlier row of the same text, and no other row repeating one; a
run at least four cells; a shape's levels holding its cells; and a row's
(unlowered <why>), the reason a shape's cells alone are left to the
evaluator with no program (KERNEL.44), only on a shape with cells alone and
with neither (missing ...) nor (every-frame ...) beside it.

WHAT IT RUNS: `<impl> calc <file> --shapes`, and for a bench program
`<impl> build <program> --prelude scripts/prelude.vla --phrasebook
scripts/polyglotta/english.vla --out target/shapes/<name>.xlsx --replace`
first, on the built door at target/debug/frazaro.exe (or target/release),
or on -Impl when given. Its stdout is compared to the golden after line
endings are normalized and trailing blank lines dropped, and the first
differing line is named with both texts. A tree with no door says SKIPPED
for that half and exits 0; the floors, the order and the totals are checked
either way.

-Control proves the check on fakes and copies in a scratch directory: a fake
door that prints each fixture's golden (every golden must pass, the cache
written and then reused); one that adds a cell to the first shape row it
prints (every golden must fail, naming its first shape row); and copies of a
golden broken one way each - a total, the order of two shape rows, a readers
row after a shape row, a (same k) pointing at a later row, the kinds out of
their order, a run under four cells, and an (unlowered ...) on a shape with
no cell alone - and a copy of the subset golden
whose most is above its readers (KERNEL.35), each of which the golden's own
checks must fail, the golden as it is passing.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no
COM, no network; exit 0 clean, exit 1 with every problem named. ASCII only,
since a BOM-less .ps1 is read as ANSI by PowerShell 5.1.

Usage:  powershell -File tools\check_shapes_golden.ps1
        powershell -File tools\check_shapes_golden.ps1 -Impl C:\path\to\frazaro.exe
        powershell -File tools\check_shapes_golden.ps1 -Control
#>
param(
    [string]$Impl = '',
    [switch]$Control
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# --- the goldens and their floors ---------------------------------------------
# 2026-10-12, KERNEL.30: the seven fixtures oracle 12 computes, and the four
# bench programs of KERNEL.20's sixth slice, built by the door. Floors are
# line counts: a plan row, a readers row a sheet read, a shape row a shape.
# Raise a floor when a regenerated golden is longer; lower it only with a
# regenerated golden that is shorter, and say why.
# 2026-10-18, KERNEL.8's first slice: the lookups' fixture, an eighth, and
# the lookups' bench program, a fifth, its golden named lookups_bench beside
# the fixture's; the subset's two goldens moved, not their floors (VLOOKUP
# computed, its shape's reads an array and its calls VLOOKUP's).
# 2026-10-19, KERNEL.8's first slice: the owner's Excel save of the lookups'
# fixture, a ninth, its census the build's but for the sheet's name as Excel
# writes it into the formulas (`table!`).
# 2026-10-19, KERNEL.8's second slice: the intersection's fixture, a tenth;
# no other golden moved, a range where one value is wanted met by `Ctx` at
# each cell of a run as before (the owner's ruling, the walk alone); 32
# lines, then 33 with `SINGLE` over a span of sheets after Excel's first
# save; Excel's second save of it, an eleventh, its census the build's but
# for the sheet's name as Excel writes it (`data!`) and `@` written bare at
# a legacy formula's top.
# 2026-10-19, KERNEL.8's third slice: the text functions' fixture, a
# twelfth, its text left to `Ctx` (`(unlowered text)`, a program holding
# numbers alone) and its Order sheet's ranks one run; the subset's two
# goldens moved, not their floors (CONCATENATE computed, its shape's calls).
# 2026-10-20, the third slice folded to Excel's save: no census moved; the
# owner's Excel save of the text functions' fixture, a thirteenth, its
# census the build's but for the sheet's name as Excel writes it (`data!`).
# 2026-10-20, KERNEL.8's fourth slice: the date functions' fixture, a
# fourteenth, its Calendar sheet's five fills five runs. Folded to Excel's
# first save, with 51 formulas after its 297 for Excel's second: 351 lines.
# Folded to its second, with 11 more for its third: 360 lines. Folded to
# its third, with 5 more for its fourth: 365 lines. The owner's fourth save
# of it, a fifteenth, its census the build's but for the sheet's name as
# Excel writes it into the formulas (`data!`).
# 2026-10-20, KERNEL.8's fifth slice: the aggregates by criteria's fixture,
# a sixteenth, every formula a cell alone; and the owner's second Excel
# save of it, a seventeenth.
# 2026-10-20, KERNEL.8's sixth slice: the arrays' fixture, an eighteenth,
# every formula a cell alone; eight censuses moved, not their counts: a
# dynamic formula's shape, which a program would read as a legacy one's,
# unlowered (`array`), and a shape whose parse holds an array computed at
# each cell by Ctx, its reads `each`.
# 2026-10-21, the arrays' fixture folded to Excel's first save, 15 rows
# added for its second: 212 lines. The owner's second Excel save of it, a
# nineteenth: 215 lines, the build's 210 formulas and the three legacy
# array formulas typed into K1:K3.
$goldens = @(
    @{ Name = 'saved';         Fixture = 'scripts/reflect/saved.xlsx';        Floor = 5 },
    @{ Name = 'fixture';       Fixture = 'scripts/reflect/fixture.xlsx';      Floor = 25 },
    @{ Name = 'build_fixture'; Fixture = 'scripts/build/fixture_golden.xlsx'; Floor = 5 },
    @{ Name = 'opendocument';  Fixture = 'scripts/reflect/opendocument.ods';  Floor = 26 },
    @{ Name = 'subset';        Fixture = 'scripts/recalc/subset.xlsx';        Floor = 106 },
    @{ Name = 'subset_saved';  Fixture = 'scripts/recalc/subset_saved.xlsx';  Floor = 106 },
    @{ Name = 'edges';         Fixture = 'scripts/recalc/edges.xlsx';         Floor = 17 },
    @{ Name = 'lookups';       Fixture = 'scripts/recalc/lookups.xlsx';       Floor = 165 },
    @{ Name = 'lookups_saved'; Fixture = 'scripts/recalc/lookups_saved.xlsx'; Floor = 165 },
    @{ Name = 'intersect';     Fixture = 'scripts/recalc/intersect.xlsx';     Floor = 33 },
    @{ Name = 'intersect_saved'; Fixture = 'scripts/recalc/intersect_saved.xlsx'; Floor = 33 },
    @{ Name = 'texts';         Fixture = 'scripts/recalc/texts.xlsx';         Floor = 194 },
    @{ Name = 'texts_saved';   Fixture = 'scripts/recalc/texts_saved.xlsx';   Floor = 194 },
    @{ Name = 'dates';         Fixture = 'scripts/recalc/dates.xlsx';         Floor = 365 },
    @{ Name = 'dates_saved';   Fixture = 'scripts/recalc/dates_saved.xlsx';   Floor = 365 },
    @{ Name = 'criteria';      Fixture = 'scripts/recalc/criteria.xlsx';      Floor = 230 },
    @{ Name = 'criteria_saved'; Fixture = 'scripts/recalc/criteria_saved.xlsx'; Floor = 230 },
    @{ Name = 'arrays';        Fixture = 'scripts/recalc/arrays.xlsx';        Floor = 212 },
    @{ Name = 'arrays_saved';  Fixture = 'scripts/recalc/arrays_saved.xlsx';  Floor = 215 },
    @{ Name = 'column';        Program = 'scripts/recalc/bench/column.txt';   Floor = 5 },
    @{ Name = 'chain';         Program = 'scripts/recalc/bench/chain.txt';    Floor = 5 },
    @{ Name = 'sums';          Program = 'scripts/recalc/bench/sums.txt';     Floor = 6 },
    @{ Name = 'life';          Program = 'scripts/recalc/bench/life.txt';     Floor = 4 },
    @{ Name = 'lookups_bench'; Program = 'scripts/recalc/bench/lookups.txt';  Floor = 9 }
)
$goldenDir = Join-Path $repoRoot 'scripts/recalc/shapes'
$cacheDir = Join-Path $repoRoot 'target/shapes'

# The kinds of read, in the order a row prints them; `stride` since a door's
# plan keeps its sheets as arrays where a machine's would (KERNEL.31); and
# `keyed` (KERNEL.8), a vector a run's cells search by the index they share.
$kinds = @('stride', 'slot', 'array', 'lookup', 'cells', 'once', 'start', 'each', 'keyed')

# A text as its lines: line endings normalized, trailing blank lines dropped.
function Get-Lines([string]$text) {
    $t = $text -replace "`r`n", "`n"
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($l in ($t -split "`n")) { $lines.Add($l) }
    while ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq '') { $lines.RemoveAt($lines.Count - 1) }
    return ,$lines
}

# A quoted string of the record, its escapes as WriteDatum writes them.
$q = '"((?:[^"\\]|\\.)*)"'
$planRow = '^\(plan \(shapes (\d+)\) \(texts (\d+)\) \(formulas (\d+)\) \(levels (\d+)\) \(runs (\d+)\) \(lanes (\d+)\) \(alone (\d+)\) \(every-frame (\d+)\) \(cycles (\d+)\)\)$'
$readersRow = '^\(readers ' + $q + ' (\d+)(?: \(most (\d+)\))?\)$'
$shapeRow = '^\(shape (\d+) ' + $q + ' ' + $q + ' ' + $q + ' \(cells (\d+)\) \(levels(?: (\d+) (\d+) (\d+) (\d+))?\) \(runs (\d+) (\d+) (\d+) (\d+)\) \(alone (\d+)\) \(reads((?: [a-z]+)*)\)' +
    '(?: \(calls((?: "(?:[^"\\]|\\.)*")+)\))?(?: \(missing((?: "(?:[^"\\]|\\.)*")+)\))?(?: \(sheets (\d+)\))?(?: \(same (\d+)\))?(?: \(every-frame (name|span|broken|unread)\))?' +
    '(?: \(unlowered (text|error|name|range|form|kinds|span|sheet|omitted|deep|array)\))?\)$'

# The order, the forms and the totals read off a golden's lines: problems,
# empty when they hold, and what was counted.
function Test-Census($lines) {
    $problems = New-Object System.Collections.Generic.List[string]
    $sum = @{ shapes = 0; texts = 0; formulas = 0; runs = 0; lanes = 0; alone = 0; every = 0; cycles = 0; deepest = -1; widestCount = 0 }
    if ($lines.Count -eq 0) { $problems.Add('an empty census'); return @{ Problems = $problems; Sum = $sum } }
    $plan = [regex]::Match($lines[0], $planRow)
    if (-not $plan.Success) { $problems.Add("line 1: not the plan row: $($lines[0])"); return @{ Problems = $problems; Sum = $sum } }
    $seenShape = $false
    $lastNumber = -1
    # Keyed exactly, as the door keys a shape by its text: two texts that
    # differ in a letter's case alone are two shapes (KERNEL.8, slice c,
    # `EXACT("a","A")` beside `EXACT("a","a")`), where `@{}` is caseless.
    $texts = New-Object System.Collections.Hashtable ([StringComparer]::Ordinal)
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $n = $i + 1
        $readers = [regex]::Match($line, $readersRow)
        if ($readers.Success) {
            if ($seenShape) { $problems.Add("line ${n}: a readers row after the shape rows began: $line") }
            # KERNEL.35: an indexed sheet's most a changed cell meets is one
            # reader at least and its readers at most.
            if ($readers.Groups[3].Success) {
                $most = [int]$readers.Groups[3].Value
                if ($most -lt 1 -or $most -gt [int]$readers.Groups[2].Value) { $problems.Add("line ${n}: a most that is not between one and its readers: $line") }
            }
            continue
        }
        $m = [regex]::Match($line, $shapeRow)
        if (-not $m.Success) { $problems.Add("line ${n}: not a row of the census: $line"); continue }
        $seenShape = $true
        $g = $m.Groups
        $number = [int]$g[1].Value
        if ($number -le $lastNumber) { $problems.Add("line ${n}: shape $number after shape $lastNumber") }
        $lastNumber = $number
        $text = $g[4].Value
        $cells = [int]$g[5].Value
        $runs = [int]$g[10].Value; $lanes = [int]$g[11].Value; $narrowest = [int]$g[12].Value; $widest = [int]$g[13].Value
        $alone = [int]$g[14].Value
        $inOrder = $lanes + $alone
        if ($inOrder -gt $cells) { $problems.Add("line ${n}: $lanes cells in runs and $alone alone, more than its $cells cells") }
        if ($runs -eq 0) {
            if ($lanes -ne 0 -or $narrowest -ne 0 -or $widest -ne 0) { $problems.Add("line ${n}: no run, yet lanes: $line") }
        } elseif ($narrowest -lt 4 -or $narrowest -gt $widest -or $lanes -lt $runs * $narrowest -or $lanes -gt $runs * $widest) {
            $problems.Add("line ${n}: runs that are not runs of at least four cells: $line")
        }
        if ($g[6].Success) {
            $first = [int]$g[6].Value; $last = [int]$g[7].Value; $count = [int]$g[8].Value; $wide = [int]$g[9].Value
            if ($first -gt $last -or $count -lt 1 -or $count -gt ($last - $first + 1) -or $wide -lt 1 -or $wide -gt $inOrder -or $inOrder -gt $count * $wide) {
                $problems.Add("line ${n}: levels that do not hold its cells: $line")
            }
            if ($last -gt $sum.deepest) { $sum.deepest = $last }
            if ($count -gt $sum.widestCount) { $sum.widestCount = $count }
        } elseif ($inOrder -ne 0) {
            $problems.Add("line ${n}: no level, yet cells in the order: $line")
        }
        $read = @($g[15].Value.Trim() -split ' ' | Where-Object { $_ -ne '' })
        $at = -1
        foreach ($k in $read) {
            $j = [array]::IndexOf($kinds, $k)
            if ($j -lt 0) { $problems.Add("line ${n}: not a kind of read: $k") }
            elseif ($j -le $at) { $problems.Add("line ${n}: the kinds out of their order: $line") }
            $at = $j
        }
        if ($g[19].Success) {
            $k = [int]$g[19].Value
            if (-not $texts.ContainsKey($text) -or $texts[$text] -ne $k) { $problems.Add("line ${n}: (same $k), not the first earlier row of its text") }
        } else {
            if ($texts.ContainsKey($text)) { $problems.Add("line ${n}: the text of shape $($texts[$text]) with no (same ...)") }
            $texts[$text] = $number
            $sum.texts++
        }
        $sum.shapes++
        $sum.formulas += $cells
        $sum.runs += $runs
        $sum.lanes += $lanes
        $sum.alone += $alone
        $sum.cycles += $cells - $inOrder
        if ($g[20].Success) { $sum.every += $cells }
        # KERNEL.44: a shape left to the evaluator says why only where it has
        # cells alone and neither a function the library lacks nor a reason to
        # be computed every frame says so already.
        if ($g[21].Success) {
            if ($alone -eq 0) { $problems.Add("line ${n}: (unlowered ...) on a shape with no cell alone: $line") }
            if ($g[17].Success -or $g[20].Success) { $problems.Add("line ${n}: (unlowered ...) beside (missing ...) or (every-frame ...): $line") }
        }
    }
    $p = $plan.Groups
    $stated = [ordered]@{ shapes = [int]$p[1].Value; texts = [int]$p[2].Value; formulas = [int]$p[3].Value; runs = [int]$p[5].Value; lanes = [int]$p[6].Value; alone = [int]$p[7].Value; every = [int]$p[8].Value; cycles = [int]$p[9].Value }
    foreach ($key in $stated.Keys) {
        if ($stated[$key] -ne $sum[$key]) { $problems.Add("the plan row states $key $($stated[$key]), the shape rows sum to $($sum[$key])") }
    }
    $levels = [int]$p[4].Value
    if ($sum.shapes -gt 0 -and ($levels -lt $sum.widestCount -or $levels -gt $sum.deepest + 1)) {
        $problems.Add("the plan row states $levels levels, outside what the shape rows hold ($($sum.widestCount) to $($sum.deepest + 1))")
    }
    return @{ Problems = $problems; Sum = $sum }
}

# `<impl> <args>` with the door's stderr set aside: its exit code and every
# line of its stdout, read as the UTF-8 the door writes whatever the
# console's code page (KERNEL.8, slice c: the texts' fixture holds text
# beyond ASCII, which PowerShell would otherwise read by the console's).
function Invoke-Door([string]$door, [string[]]$doorArgs) {
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $encoding = $null
    try { $encoding = [Console]::OutputEncoding; [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }
    try {
        $global:LASTEXITCODE = 0
        $lines = & $door @doorArgs 2>$null
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $old
        if ($null -ne $encoding) { try { [Console]::OutputEncoding = $encoding } catch { } }
    }
    if ($null -eq $code) { $code = 0 }
    $text = if ($null -eq $lines) { '' } else { (@($lines) | ForEach-Object { [string]$_ }) -join "`n" }
    return @{ ExitCode = [int]$code; Text = $text }
}

function Get-Sha([string]$path) {
    return (Get-FileHash -Algorithm SHA256 -LiteralPath $path).Hash
}

# The census a door prints of a fixture: a file's, run each time; a bench
# program's from the cache while the door and the program are the ones it
# was made from, else built and read again, and kept. Its problems, its
# text, and where it came from.
function Get-Census([string]$door, $g, [string]$cache) {
    $problems = New-Object System.Collections.Generic.List[string]
    if ($g.ContainsKey('Fixture')) {
        $fixture = Join-Path $repoRoot $g.Fixture
        if (-not (Test-Path $fixture)) { $problems.Add("$($g.Name): no fixture at $fixture"); return @{ Problems = $problems } }
        $r = Invoke-Door $door @('calc', $fixture, '--shapes')
        if ($r.ExitCode -ne 0) { $problems.Add("$($g.Name): calc --shapes exited $($r.ExitCode)"); return @{ Problems = $problems } }
        return @{ Problems = $problems; Text = $r.Text; From = 'the door' }
    }
    $program = Join-Path $repoRoot $g.Program
    if (-not (Test-Path $program)) { $problems.Add("$($g.Name): no program at $program"); return @{ Problems = $problems } }
    if (-not (Test-Path $cache)) { New-Item -ItemType Directory -Path $cache -Force | Out-Null }
    $stamp = "door $(Get-Sha $door) program $(Get-Sha $program)"
    $stampPath = Join-Path $cache "$($g.Name).stamp"
    $censusPath = Join-Path $cache "$($g.Name).census"
    if ((Test-Path $stampPath) -and (Test-Path $censusPath) -and ([System.IO.File]::ReadAllText($stampPath) -eq $stamp)) {
        return @{ Problems = $problems; Text = [System.IO.File]::ReadAllText($censusPath); From = 'the cache, door and program unchanged' }
    }
    if (Test-Path $stampPath) { Remove-Item -LiteralPath $stampPath }
    $book = Join-Path $cache "$($g.Name).xlsx"
    $r = Invoke-Door $door @('build', $program, '--prelude', (Join-Path $repoRoot 'scripts/prelude.vla'), '--phrasebook', (Join-Path $repoRoot 'scripts/polyglotta/english.vla'), '--out', $book, '--replace')
    if ($r.ExitCode -ne 0) { $problems.Add("$($g.Name): build exited $($r.ExitCode)"); return @{ Problems = $problems } }
    $r = Invoke-Door $door @('calc', $book, '--shapes')
    if ($r.ExitCode -ne 0) { $problems.Add("$($g.Name): calc --shapes exited $($r.ExitCode)"); return @{ Problems = $problems } }
    [System.IO.File]::WriteAllText($censusPath, $r.Text)
    [System.IO.File]::WriteAllText($stampPath, $stamp)
    return @{ Problems = $problems; Text = $r.Text; From = 'the door, built and read' }
}

# One door against one golden: problems, empty when its rows are the golden
# whole, and where the rows came from.
function Measure-Door([string]$door, $g, [string]$cache, [string]$dir) {
    $got = Get-Census $door $g $cache
    $problems = $got.Problems
    if ($problems.Count -gt 0) { return @{ Problems = $problems; From = '' } }
    $want = Get-Lines ([System.IO.File]::ReadAllText((Join-Path $dir "$($g.Name).vla")))
    $have = Get-Lines $got.Text
    $n = [Math]::Min($want.Count, $have.Count)
    for ($i = 0; $i -lt $n; $i++) {
        if ($want[$i] -ne $have[$i]) {
            $problems.Add("$($g.Name): differs at line $($i + 1) of $($want.Count): golden $($want[$i]) / door $($have[$i])")
            return @{ Problems = $problems; From = $got.From }
        }
    }
    if ($want.Count -ne $have.Count) { $problems.Add("$($g.Name): the door printed $($have.Count) lines, the golden holds $($want.Count); they agree to line $n") }
    return @{ Problems = $problems; From = $got.From }
}

# Every golden's floor, order and totals: problems, and lines to print.
function Test-Goldens([string]$dir) {
    $problems = New-Object System.Collections.Generic.List[string]
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($g in $goldens) {
        $path = Join-Path $dir "$($g.Name).vla"
        if (-not (Test-Path $path)) { $problems.Add("$($g.Name): no golden at $path"); continue }
        $rows = Get-Lines ([System.IO.File]::ReadAllText($path))
        if ($rows.Count -lt $g.Floor) {
            $problems.Add("$($g.Name): the golden holds $($rows.Count) lines, below its floor of $($g.Floor)")
            $lines.Add("  $($g.Name): $($rows.Count) lines, BELOW THE FLOOR of $($g.Floor)  FAIL")
        } elseif ($rows.Count -gt $g.Floor) {
            $lines.Add("  $($g.Name): $($rows.Count) lines, above the floor of $($g.Floor) - raise the floor in this file")
        } else {
            $lines.Add("  $($g.Name): $($g.Floor) lines, at its floor")
        }
        $t = Test-Census $rows
        foreach ($p in $t.Problems) { $problems.Add("$($g.Name): $p"); $lines.Add("  $($g.Name): $p  FAIL") }
        if ($t.Problems.Count -eq 0) {
            $s = $t.Sum
            $lines.Add("  $($g.Name): the order and the totals hold; $($s.shapes) shapes of $($s.texts) texts over $($s.formulas) cells, $($s.runs) runs")
        }
    }
    return @{ Problems = $problems; Lines = $lines }
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_shapes_golden_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    $failures = New-Object System.Collections.Generic.List[string]
    try {
        # The fakes answer `calc <file> --shapes` with the golden of the
        # file's base name, and `build ... --out <file>` by writing the file.
        $map = ($goldens | ForEach-Object { "'" + $_.Name + "' { '" + ((Join-Path $goldenDir ($_.Name + '.vla')) -replace "'", "''") + "' }" }) -join '; '
        $common = "param([Parameter(ValueFromRemainingArguments=`$true)][string[]]`$a)`n" +
            "if (`$a[0] -eq 'build') { `$o = `$a[[array]::IndexOf(`$a, '--out') + 1]; [System.IO.File]::WriteAllText(`$o, 'built'); exit 0 }`n" +
            "if (`$a[0] -ne 'calc' -or `$a[2] -ne '--shapes') { exit 2 }`n" +
            "`$base = [System.IO.Path]::GetFileNameWithoutExtension(`$a[1]); if (`$base -eq 'fixture_golden') { `$base = 'build_fixture' }`n" +
            "`$golden = switch (`$base) { $map }`n" +
            "`$lines = [System.IO.File]::ReadAllText(`$golden) -replace ""``r``n"", ""``n"" -split ""``n""`n"
        $clean = Join-Path $tmp 'clean.ps1'
        [System.IO.File]::WriteAllText($clean, $common + "`$lines | ForEach-Object { `$_ }`n")
        $mutant = Join-Path $tmp 'mutant.ps1'
        [System.IO.File]::WriteAllText($mutant, $common + "`$done = `$false; `$lines | ForEach-Object { if (-not `$done -and `$_ -match '^\(shape ') { `$done = `$true; `$_ -replace '\(cells (\d+)\)', '(cells 999999)' } else { `$_ } }`n")
        $cache = Join-Path $tmp 'cache'

        foreach ($g in $goldens) {
            $r = Measure-Door $clean $g $cache $goldenDir
            if ($r.Problems.Count -ne 0) { $failures.Add("clean fake: $($g.Name) failed: $($r.Problems -join '; ')") }
        }
        # A second pass reads the bench programs' rows from the cache.
        foreach ($g in ($goldens | Where-Object { $_.ContainsKey('Program') })) {
            $r = Measure-Door $clean $g $cache $goldenDir
            if ($r.From -ne 'the cache, door and program unchanged') { $failures.Add("clean fake: $($g.Name) was not read from the cache on the second pass: $($r.From)") }
        }
        $mutantCache = Join-Path $tmp 'mutant_cache'
        foreach ($g in $goldens) {
            $r = Measure-Door $mutant $g $mutantCache $goldenDir
            $named = @($r.Problems | Where-Object { $_ -match 'differs at line \d+ of' })
            if ($named.Count -eq 0) { $failures.Add("mutant fake: $($g.Name) passed, or failed without naming the line") }
        }
        $clean_ = Test-Goldens $goldenDir
        if ($clean_.Problems.Count -ne 0) { $failures.Add("the goldens as they are failed their own checks: $($clean_.Problems -join '; ')") }

        # Copies of the column golden, each broken one way.
        $columnLines = Get-Lines ([System.IO.File]::ReadAllText((Join-Path $goldenDir 'column.vla')))
        $broken = @(
            @{ Name = 'a total'; Edit = { param($l) $l[0] = $l[0] -replace '\(lanes 60000\)', '(lanes 59999)'; $l } },
            @{ Name = 'two shape rows swapped'; Edit = { param($l) $t = $l[2]; $l[2] = $l[3]; $l[3] = $t; $l } },
            @{ Name = 'a readers row after a shape row'; Edit = { param($l) $t = $l[1]; $l[1] = $l[2]; $l[2] = $t; $l } },
            @{ Name = 'a (same k) pointing at a later row'; Edit = { param($l) $l[2] = $l[2] -replace '\)$', ' (same 2))'; $l } },
            @{ Name = 'the kinds out of their order'; Edit = { param($l) $l[3] = $l[3] -replace '\(reads stride once\)', '(reads once stride)'; $l } },
            @{ Name = 'a run under four cells'; Edit = { param($l) $l[2] = $l[2] -replace '\(runs 1 20000 20000 20000\)', '(runs 1 20000 3 20000)'; $l } },
            @{ Name = 'an (unlowered ...) on a shape with no cell alone'; Edit = { param($l) $l[2] = $l[2] -replace '\)$', ' (unlowered text))'; $l } }
        )
        foreach ($b in $broken) {
            $copy = New-Object System.Collections.Generic.List[string]
            foreach ($l in $columnLines) { $copy.Add($l) }
            $edited = & $b.Edit $copy
            $applied = (($edited -join "`n") -ne ($columnLines -join "`n"))
            if (-not $applied) { $failures.Add("broken copy '$($b.Name)': the edit did not apply"); continue }
            $t = Test-Census $edited
            if ($t.Problems.Count -eq 0) { $failures.Add("broken copy '$($b.Name)': passed the golden's own checks") }
        }
        # KERNEL.35: a copy of the subset golden, whose sheets of 20 readers
        # and more are indexed, with a most above its sheet's readers.
        $subsetLines = Get-Lines ([System.IO.File]::ReadAllText((Join-Path $goldenDir 'subset.vla')))
        $copy = New-Object System.Collections.Generic.List[string]
        foreach ($l in $subsetLines) { $copy.Add($l) }
        $at = -1
        for ($i = 0; $i -lt $copy.Count; $i++) {
            if ($copy[$i] -match '^\(readers .* \(most \d+\)\)$') { $at = $i; break }
        }
        if ($at -lt 0) {
            $failures.Add('the subset golden holds no readers row with a most')
        } else {
            $copy[$at] = $copy[$at] -replace '\(most \d+\)\)$', '(most 9999))'
            $t = Test-Census $copy
            if ($t.Problems.Count -eq 0) { $failures.Add("broken copy 'a most above its readers': passed the golden's own checks") }
        }
    } finally {
        Remove-Item -Recurse -Force -LiteralPath $tmp -ErrorAction SilentlyContinue
    }
    if ($failures.Count -gt 0) {
        Write-Output 'check_shapes_golden -Control: FAIL'
        $failures | ForEach-Object { Write-Output "  $_" }
        exit 1
    }
    Write-Output "check_shapes_golden -Control: OK - the clean fake passes all $($goldens.Count) goldens, the second pass from the cache; the mutant fake fails each, naming its line; seven broken copies of the column golden and one of the subset golden each fail the golden's own checks"
    exit 0
}

$problems = New-Object System.Collections.Generic.List[string]
$report = New-Object System.Collections.Generic.List[string]
$own = Test-Goldens $goldenDir
foreach ($p in $own.Problems) { $problems.Add($p) }
foreach ($l in $own.Lines) { $report.Add($l) }

$door = $Impl
if ($door -eq '') {
    foreach ($c in @('target/debug/frazaro.exe', 'target/debug/frazaro', 'target/release/frazaro.exe', 'target/release/frazaro')) {
        $p = Join-Path $repoRoot $c
        if (Test-Path $p) { $door = $p; break }
    }
} elseif (-not (Test-Path $door)) {
    if (Test-Path "$door.exe") { $door = "$door.exe" } else { $problems.Add("no door at $door") ; $door = '' }
}
if ($door -eq '') {
    $report.Add('  the door: SKIPPED, no door built (cargo build -p frazaro)')
} else {
    foreach ($g in $goldens) {
        $r = Measure-Door $door $g $cacheDir $goldenDir
        foreach ($p in $r.Problems) { $problems.Add($p); $report.Add("  $p  FAIL") }
        if ($r.Problems.Count -eq 0) { $report.Add("  $($g.Name): the golden whole, from $($r.From)") }
    }
    # The census and the counts are two answers, never one: together they
    # are a usage error.
    $both = Invoke-Door $door @('calc', (Join-Path $repoRoot $goldens[0].Fixture), '--counts', '--shapes')
    if ($both.ExitCode -ne 2) {
        $problems.Add("calc with --counts and --shapes together exited $($both.ExitCode), not 2")
        $report.Add("  --counts with --shapes: exit $($both.ExitCode), not the usage error's 2  FAIL")
    } else {
        $report.Add('  --counts with --shapes: the usage error, exit 2')
    }
}

if ($problems.Count -gt 0) {
    Write-Output 'check_shapes_golden: FAIL'
    $report | ForEach-Object { Write-Output $_ }
    exit 1
}
Write-Output "check_shapes_golden: OK - $($goldens.Count) census goldens hold their floors, their order and their totals$(if ($door -ne '') { ', and the door prints each whole' } else { '' })"
$report | ForEach-Object { Write-Output $_ }
exit 0
