<#
check_study_fixture.ps1 - the study's fixtures and goldens hold the answer
key (KERNEL.3, the study before the viewport, SD-29): the four workbooks
exist and are tracked, each golden is CRLF and never shorter, the answer
key's rows are in the goldens whole and the omitted cells have no row, and
with a built door the door reproduces every golden whole.

WHY: KERNEL.3 runs a study before the viewport is designed (docs/PROTOCOL.md),
and its instrument is four workbooks under scripts/study/, written by
tools/build_study_fixture.ps1 as raw OOXML with no Office automation:
find_a.xlsx and find_b.xlsx, a small model of one product's year with five
planted defects each, one of each class of Powell, Baker and Lawson's field
audit, in two isomorphic variants (PROTOCOL.md section 3.2 is the answer
key); and build_a.xlsx and build_b.xlsx, the workbooks the build half asks
for, with no defect (section 3.3). Their relations, and the find half's
audit, as the door prints them, are the goldens beside them
(<name>_relations.vla, <name>_audit.vla; section 3.4): the door's own
output, saved CRLF, blessed by the owner opening each fixture in Excel and
reading the Sales sheet against the answer key. A study whose instrument
drifts from its answer key measures nothing, so this check holds the
goldens to the key cell by cell, and holds the door to the goldens, on
every push.

WHAT IT HOLDS, always, with or without a door:
  - tools/build_study_fixture.ps1 and the four workbooks exist, and
    .gitignore holds the line !scripts/study/*.xlsx, so the workbooks are
    tracked as the build and reflect fixtures are;
  - each of the six goldens exists, is CRLF (every LF preceded by a CR and
    no CR without its LF, counted by byte; no byte order mark), and holds
    at least its floor of lines (hardcoded below, dated);
  - find_a_audit.vla is exactly the three findings the audit names of
    find_a's plants, and find_b_audit.vla exactly the four of find_b's (the
    omitted G14 reported twice, from Sales!H14 and from Summary!B4, which is
    the treaty's empty-reference rule working as written), hardcoded below,
    in order, and nothing else;
  - the relations goldens hold the planted rows whole, exactly once each:
    the constant typed over D7 (find_a) and F10 (find_b) with no formula
    row behind it, the pasted G9 and D4, the short sums in D14, the logic
    error's first row (H2 =G2/F2, G2 =D2+F2), and no row at all for the
    omitted F14 and G14 - no cell row, no formula row, no refers row whose
    SOURCE is the cell (Summary's refers row names the omitted cell as its
    target, as it should); the build goldens hold the two totals. The
    number of planted-row pins per golden, and in all, is written here as
    a number and asserted: a scan that watches a list says how long it is.

WHAT IT RUNS, with a door: `<impl> reflect <workbook>` for each of the four
and `<impl> audit <workbook>` for each find workbook, the contract's own
commands (conformance/README.md, oracles 8 and 10), on the built door at
target/debug/frazaro.exe (or target/release), or on -Impl when given;
stdout is compared to the golden after line endings are normalized to LF
and trailing blank lines dropped, and the first differing line is named
with both texts; `<impl> audit` of build_a and build_b must print nothing
and exit 0. A tree with no door says SKIPPED for that half and exits 0, as
check_audit_golden.ps1 does, since CI always has it; the static half is
checked either way.

-StudyDir points the static pins and the door's workbooks at another
folder (the control's scratch copies); the builder and .gitignore are the
repository's either way.

-Control proves the check on fakes written to a scratch directory under the
temp folder: the study folder as it is (the static half must pass); a fake
door, a .ps1 the check runs as -Impl, that prints the golden beside the
workbook by the workbook's base name and nothing for a build audit (must
pass every row); a mutant door that changes one character of the first
line it prints, and prints one finding where the fake prints none (must
fail every golden row naming line 1, and both build audits); a copy of the
study folder with find_b_audit.vla's last line deleted (must fail the
exact-lines pin and the floor, and nothing else); and a copy with
find_a_relations.vla's G9 row changed from =D9-F8 to the =D9-F9 the column
expects (must fail that planted-row pin, and nothing else). One line per
case, PASS or FAIL; exit 1 if any case does not behave as it must.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no
COM, no network; exit 0 clean, exit 1 with every problem named. The door's
stderr is read under ErrorActionPreference Continue, as
check_audit_golden.ps1 does. ASCII only, since a BOM-less .ps1 is read as
ANSI by PowerShell 5.1.

Usage:  powershell -File tools\check_study_fixture.ps1
        powershell -File tools\check_study_fixture.ps1 -Impl C:\path\to\frazaro.exe
        powershell -File tools\check_study_fixture.ps1 -StudyDir C:\path\to\a\copy
        powershell -File tools\check_study_fixture.ps1 -Control
#>
param(
    [string]$Impl = '',
    [string]$StudyDir = '',
    [switch]$Control
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$utf8 = New-Object System.Text.UTF8Encoding($false)

# --- what must exist ---------------------------------------------------------
$builder = 'tools/build_study_fixture.ps1'
$defaultStudy = 'scripts/study'
$workbooks = @('find_a.xlsx', 'find_b.xlsx', 'build_a.xlsx', 'build_b.xlsx')
$ignoreLine = '!scripts/study/*.xlsx'

# --- the goldens and their floors: each line count, never lowered ------------
# 2026-10-08, KERNEL.3: find_a_relations 345 and find_b_relations 345 (the
# two variants are isomorphic: three sheet rows, three names, 129 cell rows,
# 80 formula rows and 130 refers rows each); build_a_relations 54 and
# build_b_relations 54 (one sheet row, 31 cells, 8 formulas, 14 refers);
# find_a_audit 3 (the typed-over D7, the inconsistent G9, Summary!B3's empty
# reference to F14); find_b_audit 4 (the typed-over F10, the inconsistent D4,
# and the omitted G14 reported twice, from Sales!H14 and from Summary!B4,
# which is the treaty's empty-reference rule working as written). Raise a
# floor when a regenerated golden is longer; lower it only with a
# regenerated golden that is shorter, and say why.
$goldens = @(
    @{ Name = 'find_a_relations';  Golden = 'find_a_relations.vla';  Floor = 345 },
    @{ Name = 'find_b_relations';  Golden = 'find_b_relations.vla';  Floor = 345 },
    @{ Name = 'build_a_relations'; Golden = 'build_a_relations.vla'; Floor = 54 },
    @{ Name = 'build_b_relations'; Golden = 'build_b_relations.vla'; Floor = 54 },
    @{ Name = 'find_a_audit';      Golden = 'find_a_audit.vla';      Floor = 3 },
    @{ Name = 'find_b_audit';      Golden = 'find_b_audit.vla';      Floor = 4 }
)

# --- the audit goldens, exactly ----------------------------------------------
# What `frazaro audit` names of each find workbook's five plants (PROTOCOL.md
# section 3.2, the last column): the typed-over constant, the pasted formula,
# and the omitted total through every cell that reads it; the short sum and
# the logic error are found only by reading. Each list is as long as its
# golden's floor, since the golden is exactly these lines.
$audits = @(
    @{ Name = 'find_a_audit'; Golden = 'find_a_audit.vla'; Lines = @(
        '(typed-over "Sales!D7" 6500 "=B7*C7")',
        '(inconsistent "Sales!G9" "=D9-F8" "=D9-F9")',
        '(empty-reference "Summary!B3" "Sales!F14")'
    ) },
    @{ Name = 'find_b_audit'; Golden = 'find_b_audit.vla'; Lines = @(
        '(typed-over "Sales!F10" 4000 "=B10*E10")',
        '(inconsistent "Sales!D4" "=B3*C4" "=B4*C4")',
        '(empty-reference "Sales!H14" "Sales!G14")',
        '(empty-reference "Summary!B4" "Sales!G14")'
    ) }
)

# --- the planted rows, in the relations goldens whole ------------------------
# Present: a line equal to the text, exactly once. Absent: no line matching
# the pattern; a refers pattern names the cell as a row's SOURCE, since the
# omitted cell is rightly the TARGET of Summary's refers row. Pins is the
# length of both lists together, written as a number and asserted, and
# $plantedPins the sum over the four goldens.
$planted = @(
    @{
        Name    = 'find_a_relations'
        Golden  = 'find_a_relations.vla'
        Pins    = 8
        Present = @(
            '(cell "Sales" "D7" 6500)',                # hard-coding: 6500 typed over =B7*C7
            '(formula "Sales" "G9" "=D9-F8")',         # copy and paste: August's margin on July's cost
            '(formula "Sales" "D14" "=SUM(D2:D12)")',  # reference: December left out
            '(formula "Sales" "H2" "=G2/F2")'          # logic: the margin over the cost, every row
        )
        Absent  = @(
            '^\(formula "Sales" "D7" ',                # the constant has no formula behind it
            '^\(cell "Sales" "F14" ',                  # omission: F14 was never written
            '^\(formula "Sales" "F14" ',
            '^\(refers "Sales!F14" '
        )
    },
    @{
        Name    = 'find_b_relations'
        Golden  = 'find_b_relations.vla'
        Pins    = 8
        Present = @(
            '(cell "Sales" "F10" 4000)',               # hard-coding: 4000 typed over =B10*E10
            '(formula "Sales" "D4" "=B3*C4")',         # copy and paste: March's revenue on February's units
            '(formula "Sales" "D14" "=SUM(D3:D13)")',  # reference: January left out
            '(formula "Sales" "G2" "=D2+F2")'          # logic: a sign flipped, every row
        )
        Absent  = @(
            '^\(formula "Sales" "F10" ',               # the constant has no formula behind it
            '^\(cell "Sales" "G14" ',                  # omission: G14 was never written
            '^\(formula "Sales" "G14" ',
            '^\(refers "Sales!G14" '
        )
    },
    @{
        Name    = 'build_a_relations'
        Golden  = 'build_a_relations.vla'
        Pins    = 2
        Present = @(
            '(cell "Plan" "B8" 890)',                  # the units total, Jan to Jun
            '(cell "Plan" "D8" 35600)'                 # the revenue total at a price of 40
        )
        Absent  = @()
    },
    @{
        Name    = 'build_b_relations'
        Golden  = 'build_b_relations.vla'
        Pins    = 2
        Present = @(
            '(cell "Plan" "B8" 900)',                  # the units total, Jul to Dec
            '(cell "Plan" "D8" 45000)'                 # the revenue total at a price of 50
        )
        Absent  = @()
    }
)
$plantedPins = 20

# --- what the door runs ------------------------------------------------------
# A row with no golden prints nothing: the build workbooks have no defect.
$doorRows = @(
    @{ Name = 'find_a reflect';  Command = 'reflect'; Fixture = 'find_a.xlsx';  Golden = 'find_a_relations.vla' },
    @{ Name = 'find_b reflect';  Command = 'reflect'; Fixture = 'find_b.xlsx';  Golden = 'find_b_relations.vla' },
    @{ Name = 'build_a reflect'; Command = 'reflect'; Fixture = 'build_a.xlsx'; Golden = 'build_a_relations.vla' },
    @{ Name = 'build_b reflect'; Command = 'reflect'; Fixture = 'build_b.xlsx'; Golden = 'build_b_relations.vla' },
    @{ Name = 'find_a audit';    Command = 'audit';   Fixture = 'find_a.xlsx';  Golden = 'find_a_audit.vla' },
    @{ Name = 'find_b audit';    Command = 'audit';   Fixture = 'find_b.xlsx';  Golden = 'find_b_audit.vla' },
    @{ Name = 'build_a audit';   Command = 'audit';   Fixture = 'build_a.xlsx'; Golden = '' },
    @{ Name = 'build_b audit';   Command = 'audit';   Fixture = 'build_b.xlsx'; Golden = '' }
)

# A text as its lines: line endings normalized, trailing blank lines dropped.
function Get-Lines([string]$text) {
    $t = $text -replace "`r`n", "`n"
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($l in ($t -split "`n")) { $lines.Add($l) }
    while ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq '') { $lines.RemoveAt($lines.Count - 1) }
    return ,$lines
}

# A file's line endings, counted by byte: its CRs, its LFs, its CRLF pairs,
# and whether a byte order mark opens it.
function Measure-Endings([string]$path) {
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $cr = 0; $lf = 0; $pairs = 0
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        if ($bytes[$i] -eq 13) {
            $cr++
            if ($i + 1 -lt $bytes.Length -and $bytes[$i + 1] -eq 10) { $pairs++ }
        } elseif ($bytes[$i] -eq 10) {
            $lf++
        }
    }
    $bom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    return @{ CR = $cr; LF = $lf; Pairs = $pairs; Bom = $bom }
}

# The static half over one study folder: a list of problems, empty when it
# holds, and the lines to print.
function Test-Static([string]$study) {
    $problems = New-Object System.Collections.Generic.List[string]
    $lines = New-Object System.Collections.Generic.List[string]

    # The builder, the workbooks, and the line that keeps them tracked.
    if (Test-Path (Join-Path $repoRoot $builder)) {
        $lines.Add("  the builder $builder exists")
    } else {
        $problems.Add("no builder at $builder"); $lines.Add("  no builder at $builder  FAIL")
    }
    $present = New-Object System.Collections.Generic.List[string]
    foreach ($w in $workbooks) {
        if (Test-Path (Join-Path $study $w)) { $present.Add($w) } else { $problems.Add("no workbook $w in $study"); $lines.Add("  no workbook $w in $study  FAIL") }
    }
    if ($present.Count -eq $workbooks.Count) { $lines.Add("  the four workbooks exist: $($workbooks -join ', ')") }
    $ignore = Join-Path $repoRoot '.gitignore'
    $tracked = $false
    if (Test-Path $ignore) {
        foreach ($l in (Get-Lines ([System.IO.File]::ReadAllText($ignore)))) { if ($l -ceq $ignoreLine) { $tracked = $true } }
    }
    if ($tracked) {
        $lines.Add("  .gitignore holds the line $ignoreLine, so the workbooks are tracked")
    } else {
        $problems.Add(".gitignore does not hold the line $ignoreLine, so the workbooks are not tracked"); $lines.Add("  .gitignore does not hold the line $ignoreLine  FAIL")
    }

    # Each golden: present, CRLF, at or above its floor.
    foreach ($g in $goldens) {
        $path = Join-Path $study $g.Golden
        if (-not (Test-Path $path)) { $problems.Add("$($g.Name): no golden at $path"); $lines.Add("  $($g.Name): no golden at $path  FAIL"); continue }
        $e = Measure-Endings $path
        if ($e.LF -eq 0 -or $e.LF -ne $e.Pairs -or $e.CR -ne $e.Pairs) {
            $problems.Add("$($g.Name): not CRLF ($($e.CR) CR, $($e.LF) LF, $($e.Pairs) CRLF pairs; every LF must follow a CR and every CR precede an LF)")
            $lines.Add("  $($g.Name): not CRLF ($($e.CR) CR, $($e.LF) LF, $($e.Pairs) CRLF pairs)  FAIL")
        } elseif ($e.Bom) {
            $problems.Add("$($g.Name): opens with a byte order mark, which the door never prints")
            $lines.Add("  $($g.Name): opens with a byte order mark  FAIL")
        } else {
            $lines.Add("  $($g.Name): CRLF ($($e.Pairs) line endings, no stray CR or LF, no byte order mark)")
        }
        $rows = Get-Lines ([System.IO.File]::ReadAllText($path))
        if ($rows.Count -lt $g.Floor) {
            $problems.Add("$($g.Name): the golden holds $($rows.Count) lines, below its floor of $($g.Floor)")
            $lines.Add("  $($g.Name): $($rows.Count) lines, BELOW THE FLOOR of $($g.Floor)  FAIL")
        } elseif ($rows.Count -gt $g.Floor) {
            $lines.Add("  $($g.Name): $($rows.Count) lines, above the floor of $($g.Floor) - raise the floor in this file")
        } else {
            $lines.Add("  $($g.Name): $($g.Floor) lines, at its floor")
        }
    }

    # The audit goldens: exactly the answer key's findings, in order.
    foreach ($a in $audits) {
        $path = Join-Path $study $a.Golden
        if (-not (Test-Path $path)) { continue }   # named above
        $floor = @($goldens | Where-Object { $_.Golden -ceq $a.Golden })[0].Floor
        if ($a.Lines.Count -ne $floor) {
            $problems.Add("$($a.Name): the script lists $($a.Lines.Count) finding(s) but sets the floor at $floor; an audit golden is exactly its findings, so the two are one number")
            $lines.Add("  $($a.Name): $($a.Lines.Count) finding(s) listed, the floor $floor  FAIL")
        }
        $rows = Get-Lines ([System.IO.File]::ReadAllText($path))
        $bad = ''
        $n = [Math]::Min($rows.Count, $a.Lines.Count)
        for ($i = 0; $i -lt $n; $i++) {
            if ($rows[$i] -cne $a.Lines[$i]) { $bad = "line $($i + 1) is $($rows[$i]), the answer key says $($a.Lines[$i])"; break }
        }
        if ($bad -eq '' -and $rows.Count -ne $a.Lines.Count) { $bad = "holds $($rows.Count) line(s), the answer key has $($a.Lines.Count)" }
        if ($bad -ne '') {
            $problems.Add("$($a.Name): $bad"); $lines.Add("  $($a.Name): $bad  FAIL")
        } else {
            $lines.Add("  $($a.Name): exactly the answer key's $($a.Lines.Count) findings, in order, and nothing else")
        }
    }

    # The planted rows: present exactly once, the omitted cells absent.
    $sum = 0
    foreach ($p in $planted) {
        $sum += $p.Pins
        $listed = $p.Present.Count + $p.Absent.Count
        if ($listed -ne $p.Pins) {
            $problems.Add("$($p.Name): $listed planted-row pin(s) listed, but the script says $($p.Pins)")
            $lines.Add("  $($p.Name): $listed planted-row pin(s) listed, but the script says $($p.Pins)  FAIL")
        }
        $path = Join-Path $study $p.Golden
        if (-not (Test-Path $path)) { continue }   # named above
        $rows = Get-Lines ([System.IO.File]::ReadAllText($path))
        $held = 0
        foreach ($want in $p.Present) {
            $hits = 0
            foreach ($r in $rows) { if ($r -ceq $want) { $hits++ } }
            if ($hits -eq 1) {
                $held++
            } else {
                $problems.Add("$($p.Name): the planted row $want appears $hits time(s), not once")
                $lines.Add("  $($p.Name): the planted row $want appears $hits time(s), not once  FAIL")
            }
        }
        foreach ($pat in $p.Absent) {
            $first = ''; $hits = 0
            foreach ($r in $rows) { if ([regex]::IsMatch($r, $pat)) { $hits++; if ($first -eq '') { $first = $r } } }
            if ($hits -eq 0) {
                $held++
            } else {
                $problems.Add("$($p.Name): $hits row(s) of the omitted cell match $pat, the first $first")
                $lines.Add("  $($p.Name): $hits row(s) of the omitted cell match $pat, the first $first  FAIL")
            }
        }
        $lines.Add("  $($p.Name): $held of $($p.Pins) planted-row pins hold")
    }
    if ($sum -ne $plantedPins) {
        $problems.Add("$sum planted-row pins over the four goldens, but the script says $plantedPins")
        $lines.Add("  $sum planted-row pins over the four goldens, but the script says $plantedPins  FAIL")
    } else {
        $lines.Add("  $plantedPins planted-row pins in all, as the script says")
    }
    return @{ Problems = $problems; Lines = $lines }
}

# `<impl> <args>` with the door's stderr set aside: its exit code and every
# line of its stdout.
function Invoke-Door([string]$impl, [string[]]$doorArgs) {
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $global:LASTEXITCODE = 0
        $lines = & $impl @doorArgs 2>$null
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $old
    }
    if ($null -eq $code) { $code = 0 }
    $text = if ($null -eq $lines) { '' } else { (@($lines) | ForEach-Object { [string]$_ }) -join "`n" }
    return @{ ExitCode = [int]$code; Text = $text }
}

# One implementation against one row: a list of problems, empty when its
# output is the golden whole, or nothing where the row has no golden.
function Measure-Impl([string]$impl, [string]$study, $row) {
    $problems = New-Object System.Collections.Generic.List[string]
    $fixture = Join-Path $study $row.Fixture
    if (-not (Test-Path $fixture)) { $problems.Add("$($row.Name): no workbook at $fixture"); return ,$problems }
    $r = Invoke-Door $impl @($row.Command, $fixture)
    if ($r.ExitCode -eq 3) { $problems.Add("$($row.Name): $($row.Command) not attempted (exit 3)"); return ,$problems }
    if ($r.ExitCode -ne 0) { $problems.Add("$($row.Name): $($row.Command) exited $($r.ExitCode)"); return ,$problems }
    $got = Get-Lines $r.Text
    if ($row.Golden -eq '') {
        if ($got.Count -ne 0) { $problems.Add("$($row.Name): the door printed $($got.Count) line(s) where a workbook with no defect has none: $($got[0])") }
        return ,$problems
    }
    $golden = Join-Path $study $row.Golden
    if (-not (Test-Path $golden)) { $problems.Add("$($row.Name): no golden at $golden"); return ,$problems }
    $want = Get-Lines ([System.IO.File]::ReadAllText($golden))
    $n = [Math]::Min($want.Count, $got.Count)
    for ($i = 0; $i -lt $n; $i++) {
        if ($want[$i] -cne $got[$i]) {
            $problems.Add("$($row.Name): differs at line $($i + 1) of $($want.Count): golden $($want[$i]) / door $($got[$i])")
            return ,$problems
        }
    }
    if ($want.Count -ne $got.Count) { $problems.Add("$($row.Name): the door printed $($got.Count) lines, the golden holds $($want.Count); they agree to line $n") }
    return ,$problems
}

# The fake door the control runs as -Impl: `reflect <workbook>` and `audit
# <workbook>` print the golden beside the workbook, by the workbook's base
# name, and an audit of a build workbook prints nothing. The mutant is the
# fake with its marked line swapped for one that changes one character of
# the first line it prints, and prints one finding where the fake prints
# none.
$fakeText = @'
param([Parameter(ValueFromRemainingArguments=$true)][string[]]$a)
$base = [System.IO.Path]::GetFileNameWithoutExtension($a[1])
$dir = Split-Path -Parent $a[1]
$which = switch ($a[0] + ' ' + $base) {
    'reflect find_a'  { 'find_a_relations.vla' }
    'reflect find_b'  { 'find_b_relations.vla' }
    'reflect build_a' { 'build_a_relations.vla' }
    'reflect build_b' { 'build_b_relations.vla' }
    'audit find_a'    { 'find_a_audit.vla' }
    'audit find_b'    { 'find_b_audit.vla' }
    'audit build_a'   { '' }
    'audit build_b'   { '' }
    default           { exit 3 }
}
$lines = @()
if ($which -ne '') { $lines = @([System.IO.File]::ReadAllText((Join-Path $dir $which)) -replace "`r`n", "`n" -split "`n") }
# MUTATION POINT
$lines | ForEach-Object { Write-Output $_ }
exit 0
'@
$mutation = 'if ($lines.Count -gt 0) { $lines[0] = $lines[0] -replace ''\)$'', ''X)'' } else { $lines = @(''(typed-over "Plan!D2" 0 "=B2*C2")'') }'

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_study_fixture_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $study0 = Join-Path $repoRoot $defaultStudy
        $verdicts = New-Object System.Collections.Generic.List[string]
        $ok = $true

        # The study folder as it is: the static half passes.
        $clean = Test-Static $study0
        if ($clean.Problems.Count -eq 0) {
            $verdicts.Add('the study folder as it is: PASS - 0 static problem(s)')
        } else {
            $ok = $false; $verdicts.Add("the study folder as it is: FAIL - $($clean.Problems.Count) static problem(s): $($clean.Problems[0])")
        }

        # The fake door passes every row; the mutant fails every row, each
        # golden row at line 1 and each build audit on the finding it adds.
        $fake = Join-Path $tmp 'fake.ps1'
        [System.IO.File]::WriteAllText($fake, $fakeText, $utf8)
        $mut = Join-Path $tmp 'mutant.ps1'
        $mutText = $fakeText.Replace('# MUTATION POINT', $mutation)
        if ($mutText -ceq $fakeText) { $ok = $false; $verdicts.Add('the mutant door: FAIL - the mutation did not apply (no marked line in the fake)') }
        [System.IO.File]::WriteAllText($mut, $mutText, $utf8)
        foreach ($row in $doorRows) {
            $r1 = Measure-Impl $fake $study0 $row
            if ($r1.Count -eq 0) {
                $verdicts.Add("the fake door, $($row.Name): PASS - 0 problem(s)")
            } else {
                $ok = $false; $verdicts.Add("the fake door, $($row.Name): FAIL - $($r1.Count) problem(s): $($r1[0])")
            }
            $r2 = Measure-Impl $mut $study0 $row
            $expect = if ($row.Golden -eq '') { '*printed 1 line(s) where a workbook with no defect has none*' } else { '*differs at line 1 of*' }
            if ($r2.Count -eq 1 -and $r2[0] -like $expect) {
                $verdicts.Add("the mutant door, $($row.Name): PASS - fails as it must: $($r2[0])")
            } else {
                $ok = $false; $verdicts.Add("the mutant door, $($row.Name): FAIL - $($r2.Count) problem(s), expected one like $expect" + $(if ($r2.Count -gt 0) { ": $($r2[0])" } else { '' }))
            }
        }

        # A copy of the study folder with find_b_audit.vla's last line deleted:
        # the floor and the exact-lines pin fail, and nothing else.
        $short = Join-Path $tmp 'short'
        New-Item -ItemType Directory -Path $short | Out-Null
        Copy-Item -Path (Join-Path $study0 '*') -Destination $short -Force
        $target = Join-Path $short 'find_b_audit.vla'
        $rows = Get-Lines ([System.IO.File]::ReadAllText($target))
        $before = $rows.Count
        $rows.RemoveAt($rows.Count - 1)
        [System.IO.File]::WriteAllText($target, (($rows -join "`r`n") + "`r`n"), $utf8)
        $after = (Get-Lines ([System.IO.File]::ReadAllText($target))).Count
        $r3 = Test-Static $short
        $floorHit = @($r3.Problems | Where-Object { $_ -like 'find_b_audit: the golden holds * lines, below its floor of *' }).Count
        $exactHit = @($r3.Problems | Where-Object { $_ -like 'find_b_audit: holds * line(s), the answer key has *' }).Count
        if ($after -eq $before - 1 -and $r3.Problems.Count -eq 2 -and $floorHit -eq 1 -and $exactHit -eq 1) {
            $verdicts.Add("find_b_audit.vla's last line deleted ($before to $after lines): PASS - fails the floor and the exact-lines pin, and nothing else: $($r3.Problems -join '; ')")
        } else {
            $ok = $false; $verdicts.Add("find_b_audit.vla's last line deleted ($before to $after lines): FAIL - $($r3.Problems.Count) problem(s), expected the floor and the exact-lines pin: $($r3.Problems -join '; ')")
        }

        # A copy with find_a_relations.vla's G9 row changed to the formula the
        # column expects: that planted-row pin fails, and nothing else.
        $pasted = Join-Path $tmp 'pasted'
        New-Item -ItemType Directory -Path $pasted | Out-Null
        Copy-Item -Path (Join-Path $study0 '*') -Destination $pasted -Force
        $target = Join-Path $pasted 'find_a_relations.vla'
        $t0 = [System.IO.File]::ReadAllText($target)
        $t1 = $t0.Replace('(formula "Sales" "G9" "=D9-F8")', '(formula "Sales" "G9" "=D9-F9")')
        [System.IO.File]::WriteAllText($target, $t1, $utf8)
        $applied = ($t1 -cne $t0)
        $r4 = Test-Static $pasted
        $pinHit = @($r4.Problems | Where-Object { $_ -like 'find_a_relations: the planted row (formula "Sales" "G9" "=D9-F8") appears 0 time(s), not once' }).Count
        if ($applied -and $r4.Problems.Count -eq 1 -and $pinHit -eq 1) {
            $verdicts.Add("find_a_relations.vla's G9 row changed to =D9-F9: PASS - fails the planted-row pin, and nothing else: $($r4.Problems[0])")
        } else {
            $ok = $false; $verdicts.Add("find_a_relations.vla's G9 row changed to =D9-F9 (applied: $applied): FAIL - $($r4.Problems.Count) problem(s), expected the planted-row pin alone: $($r4.Problems -join '; ')")
        }

        foreach ($v in $verdicts) { Write-Output ('control: ' + $v) }
        if ($ok) {
            Write-Output 'OK: control: the study folder passes, the fake door passes every row, the mutant door fails every row naming line 1 of each golden, and each mutant copy trips its pin and nothing else'
            exit 0
        }
        Write-Output 'FAIL: the control did not behave as the header says'
        exit 1
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}

$study = Join-Path $repoRoot $defaultStudy
if ($StudyDir -ne '') {
    if (-not (Test-Path $StudyDir)) { Write-Output "no study folder at $StudyDir"; exit 1 }
    $study = (Resolve-Path $StudyDir).Path
}

Write-Output '=== THE STUDY FIXTURES: THE GOLDENS HOLD THE ANSWER KEY, NEVER SHORTER, AND THE DOOR REPRODUCES THEM WHOLE (KERNEL.3, the study before the viewport) ==='
$failed = New-Object System.Collections.Generic.List[string]
$ts = Test-Static $study
$ts.Lines | ForEach-Object { Write-Output $_ }
foreach ($p in $ts.Problems) { $failed.Add($p) }

if ($Impl -eq '') {
    foreach ($candidate in @('target/debug/frazaro.exe', 'target/release/frazaro.exe', 'target/debug/frazaro', 'target/release/frazaro')) {
        $p = Join-Path $repoRoot $candidate
        if (Test-Path $p) { $Impl = $p; break }
    }
}
if ($Impl -eq '' -or -not (Test-Path $Impl)) {
    Write-Output '  SKIPPED: no built door at target/debug/frazaro.exe or target/release/frazaro.exe (cargo build --workspace writes it; CI always has it)'
} else {
    foreach ($row in $doorRows) {
        $r = Measure-Impl $Impl $study $row
        foreach ($p in $r) { $failed.Add($p); Write-Output "  $p  FAIL" }
        if ($r.Count -eq 0) {
            if ($row.Golden -eq '') {
                Write-Output "  $($row.Name): the door prints nothing, as a workbook with no defect has no finding"
            } else {
                Write-Output "  $($row.Name): the door's output is the golden, whole"
            }
        }
    }
}

Write-Output ''
if ($failed.Count -eq 0) {
    Write-Output '=== CHECK: clean - the study''s workbooks are tracked, every golden is CRLF at or above its floor and holds the answer key, and the door reproduces every golden whole ==='
    exit 0
}
Write-Output "=== CHECK: $($failed.Count) problem(s) ==="
$failed | ForEach-Object { Write-Output "  $_" }
exit 1
