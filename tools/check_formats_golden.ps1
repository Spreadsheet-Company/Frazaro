<#
check_formats_golden.ps1 - the formats golden agrees with its fixture, every
record is a code Excel took or refused and each value's text, the values
are the record's own spellings, and the codes reach every construct this
version draws.

WHY: KERNEL.47 (a number format in the record) draws a number's text under
a code of Excel's in the language (vla-lang/src/numfmt.rs), and Excel is
the reference for what each code draws: scripts/formats.txt holds the cases,
and scripts/formats_golden.txt is what VlaWriteFormatsGolden (VLA_Tests.bas)
writes of them in Excel, each code as Range.NumberFormat reads it back, or
Excel's refusal, and each value's Range.Text. The language's own test,
vla-lang/tests/formats_golden.rs, holds its texts to that file. That test
proves agreement on what the fixture asks. It cannot see a construct no case
reaches, a golden regenerated against a fixture that lost a case, a value
the golden echoes otherwise than the fixture writes it, a value spelled as
no record spells a number, or a refusal whose id the catalogue lacks. This
check holds those, host-free, on every push:

  1. the golden's headers are the fixture's, in order, and every fixture
     case has a code;
  2. every golden case is one record, CODE<TAB><the code read back> or
     REFUSED<TAB><Excel's error number>, then for a CODE one line a value
     of the fixture's, <the value as written><TAB><its text>, the value
     echoed exactly, and for a REFUSED none; a case whose header ends
     "!refused <words>" carries no value in the fixture;
  3. every value is a number in the record's spelling (digits, one point,
     a sign, no exponent) or a text after an apostrophe;
  4. grid-format-code, the refusal of a code this version does not draw, is
     an id of the language's half of the catalogue (vla-lang/data/messages.vla);
  5. the constructs the drawn cases reach, each named below, all of them:
     one to four sections and an empty one, 0, #, the point, the thousands,
     a thousandth, %, a quoted literal, an escaped character, a literal
     drawn unquoted, a literal outside ASCII, the text's @, and among the
     values a negative, zero and a text; and since KERNEL.8's third slice a
     year, a month's name, a weekday's name, an hour, an elapsed time, a
     twelve-hour clock and a second's decimals;
  6. the floors below: the cases, the values and the refused cases.

-Control proves the check on scratch copies: the real files must pass, and
each of these must fail: a golden with one record's tag misspelled, one with
its last case dropped, one echoing a value otherwise, a fixture with a value
spelled with an exponent, and a fixture with every backslash taken out of
its codes.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; exit 0 clean, exit 1 with every problem named.

Usage:  powershell -File tools\check_formats_golden.ps1
        powershell -File tools\check_formats_golden.ps1 -Control
#>
param(
    [switch]$Control
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# --- the floors, never lowered ------------------------------------------------
# 2026-10-17, KERNEL.47: 59 cases, 114 values, 13 refused cases; the golden
# written as the prediction ahead of the reference's first run in Excel.
# 2026-10-19, KERNEL.8's third slice: 68 cases, 139 values, 12 refused
# cases, the floor of refused cases lowered from 13 since a date's code is
# drawn now, its case and nine of dates and times beside it written as the
# prediction ahead of the reference's run.
$caseFloor = 68
$valueFloor = 139
$refusedFloor = 12

# The constructs rule 5 holds the drawn cases to.
$constructs = @('one section', 'two sections', 'three sections', 'four sections',
    'an empty section', 'the placeholder 0', 'the placeholder #', 'the point',
    'the thousands', 'a thousandth', 'a percent', 'a quoted literal',
    'an escaped character', 'a literal drawn unquoted', 'a literal outside ASCII',
    "the text's @", 'a negative value', 'a zero value', 'a text value',
    'a year', "a month's name", "a weekday's name", 'an hour', 'an elapsed time',
    'a twelve-hour clock', "a second's decimals")

$fixturePath = Join-Path $repoRoot 'scripts/formats.txt'
$goldenPath  = Join-Path $repoRoot 'scripts/formats_golden.txt'
$catalogPath = Join-Path $repoRoot 'vla-lang/data/messages.vla'

function Get-FormatsLines([string]$path) {
    $text = [System.IO.File]::ReadAllText($path)
    $text = $text -replace "`r`n", "`n" -replace "`r", "`n"
    return ,@($text -split "`n")
}

function New-OrdinalSet {
    return ,(New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal))
}

# The cases of a fixture or a golden: each header with the lines under it,
# blank lines skipped, as VlaWriteFormatsGolden reads them.
function Read-FormatsCases([string]$path) {
    $cases = New-Object System.Collections.ArrayList
    $current = $null
    foreach ($ln in (Get-FormatsLines $path)) {
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

# The constructs one code reaches, read outside its quotes: its sections,
# its placeholders and point, a comma before a placeholder (the thousands)
# or after the last (a thousandth), a percent, and its literals.
function Get-CodeConstructs([string]$code) {
    $found = New-OrdinalSet
    $sections = New-Object System.Collections.ArrayList
    $section = ''
    $chars = $code.ToCharArray()
    $i = 0
    while ($i -lt $chars.Length) {
        $c = $chars[$i]
        if ($c -eq '"') {
            $j = $code.IndexOf('"', $i + 1)
            if ($j -lt 0) { $j = $chars.Length }
            [void]$found.Add('a quoted literal')
            $section += 'L'
            $i = $j + 1
            continue
        }
        if ($c -eq '\') {
            [void]$found.Add('an escaped character')
            $section += 'L'
            $i += 2
            continue
        }
        if ($c -eq ';') {
            [void]$sections.Add($section)
            $section = ''
            $i++
            continue
        }
        if ([int]$c -gt 127) { [void]$found.Add('a literal outside ASCII'); $section += 'L' }
        elseif ('$-+():!^&''~{}<>= '.IndexOf($c) -ge 0) { [void]$found.Add('a literal drawn unquoted'); $section += 'L' }
        elseif ($c -eq '0') { [void]$found.Add('the placeholder 0'); $section += 'D' }
        elseif ($c -eq '#') { [void]$found.Add('the placeholder #'); $section += 'D' }
        elseif ($c -eq '.') { [void]$found.Add('the point'); $section += '.' }
        elseif ($c -eq '%') { [void]$found.Add('a percent'); $section += '%' }
        elseif ($c -eq '@') { [void]$found.Add("the text's @"); $section += '@' }
        elseif ($c -eq ',') { $section += ',' }
        $i++
    }
    [void]$sections.Add($section)
    switch ($sections.Count) {
        1 { [void]$found.Add('one section') }
        2 { [void]$found.Add('two sections') }
        3 { [void]$found.Add('three sections') }
        4 { [void]$found.Add('four sections') }
    }
    foreach ($s in $sections) {
        if ($s -eq '') { [void]$found.Add('an empty section') }
        for ($k = 0; $k -lt $s.Length; $k++) {
            if ($s[$k] -ne ',') { continue }
            $after = $s.Substring($k + 1)
            if ($after -match '^[^D]*$') { [void]$found.Add('a thousandth') }
            elseif ($after.StartsWith('D')) { [void]$found.Add('the thousands') }
        }
    }
    # A date's and a time's parts (KERNEL.8, slice c), read outside quotes.
    $bare = [regex]::Replace($code, '"[^"]*"|\\.', '')
    if ($bare -match '(?i)y') { [void]$found.Add('a year') }
    if ($bare -match '(?i)mmm') { [void]$found.Add("a month's name") }
    if ($bare -match '(?i)ddd') { [void]$found.Add("a weekday's name") }
    if ($bare -match '(?i)h') { [void]$found.Add('an hour') }
    if ($bare -match '(?i)\[[hms]+\]') { [void]$found.Add('an elapsed time') }
    if ($bare -match '(?i)am/pm|a/p') { [void]$found.Add('a twelve-hour clock') }
    if ($bare -match '(?i)s\.0') { [void]$found.Add("a second's decimals") }
    return ,$found
}

# Every problem of one fixture/golden/catalogue set, with the counts the
# floors read.
function Test-FormatsGolden([string]$fixture, [string]$golden, [string]$catalog) {
    $problems = New-Object System.Collections.ArrayList
    $fcases = Read-FormatsCases $fixture
    $gcases = Read-FormatsCases $golden

    # 4. the refusal's id
    $ids = New-OrdinalSet
    foreach ($ln in (Get-FormatsLines $catalog)) {
        if ($ln -match '^\(message (\S+) ') { [void]$ids.Add($Matches[1]) }
    }
    if (-not $ids.Contains('grid-format-code')) { [void]$problems.Add("grid-format-code is not an id of $catalog") }

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

    # 2, 3 and 5. the records, the values and the constructs
    $reached = New-OrdinalSet
    $values = 0
    $refused = 0
    for ($i = 0; $i -lt $n; $i++) {
        $f = $fcases[$i]
        $g = $gcases[$i]
        if ($f.Lines.Count -eq 0) { [void]$problems.Add(("'{0}': no code after its header" -f $f.Header)); continue }
        $ours = $f.Header -match ' !refused '
        if ($ours) {
            $refused++
            if ($f.Lines.Count -ne 1) { [void]$problems.Add(("'{0}': a case refused here carries a value" -f $f.Header)) }
        } else {
            foreach ($x in (Get-CodeConstructs ([string]$f.Lines[0]))) { [void]$reached.Add($x) }
        }
        $fvalues = @()
        if ($f.Lines.Count -gt 1) { $fvalues = @($f.Lines[1..($f.Lines.Count - 1)]) }
        foreach ($v in $fvalues) {
            if ($v -cmatch "^'") { [void]$reached.Add('a text value'); continue }
            if ($v -cnotmatch '^-?(\d+(\.\d+)?|\.\d+)$') {
                [void]$problems.Add(("'{0}': the value '{1}' is neither a number in the record's spelling nor a text after an apostrophe" -f $f.Header, $v))
                continue
            }
            if ($v.StartsWith('-')) { [void]$reached.Add('a negative value') }
            if ($v -cmatch '^-?0(\.0+)?$') { [void]$reached.Add('a zero value') }
        }
        if ($g.Lines.Count -eq 0) { [void]$problems.Add(("'{0}': no record" -f $g.Header)); continue }
        $first = [string]$g.Lines[0]
        if ($first -cmatch '^REFUSED\t\d+$') {
            if ($g.Lines.Count -ne 1) { [void]$problems.Add(("'{0}': a value after Excel's refusal" -f $g.Header)) }
            continue
        }
        if ($first -cnotmatch '^CODE\t') {
            [void]$problems.Add(("'{0}': a record that is neither CODE<TAB><code> nor REFUSED<TAB><number>: '{1}'" -f $g.Header, $first))
            continue
        }
        $gvalues = @()
        if ($g.Lines.Count -gt 1) { $gvalues = @($g.Lines[1..($g.Lines.Count - 1)]) }
        if ($gvalues.Count -ne $fvalues.Count) {
            [void]$problems.Add(("'{0}': the fixture has {1} values and the golden {2}" -f $g.Header, $fvalues.Count, $gvalues.Count))
        }
        $m = [Math]::Min($gvalues.Count, $fvalues.Count)
        for ($k = 0; $k -lt $m; $k++) {
            $tab = $gvalues[$k].IndexOf("`t")
            if ($tab -lt 0) { [void]$problems.Add(("'{0}': no tab in '{1}'" -f $g.Header, $gvalues[$k])); continue }
            if ($gvalues[$k].Substring(0, $tab) -cne $fvalues[$k]) {
                [void]$problems.Add(("'{0}': the golden echoes '{1}' where the fixture writes '{2}'" -f $g.Header, $gvalues[$k].Substring(0, $tab), $fvalues[$k]))
            }
            $values++
        }
    }
    foreach ($x in $constructs) {
        if (-not $reached.Contains($x)) { [void]$problems.Add("no drawn case reaches $x") }
    }

    # 6. the floors
    if ($gcases.Count -lt $caseFloor) { [void]$problems.Add(("{0} cases, below the floor of {1}" -f $gcases.Count, $caseFloor)) }
    if ($values -lt $valueFloor) { [void]$problems.Add(("{0} values, below the floor of {1}" -f $values, $valueFloor)) }
    if ($refused -lt $refusedFloor) { [void]$problems.Add(("{0} cases refused here, below the floor of {1}" -f $refused, $refusedFloor)) }

    return @{ Problems = $problems; Cases = $gcases.Count; Values = $values; Refused = $refused; Reached = $reached.Count }
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_formats_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $real = Test-FormatsGolden $fixturePath $goldenPath $catalogPath
        $lines = Get-FormatsLines $goldenPath
        $firstCode = -1
        $firstValue = -1
        $lastHeader = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($firstCode -lt 0 -and $lines[$i].StartsWith("CODE`t")) { $firstCode = $i }
            if ($firstCode -ge 0 -and $firstValue -lt 0 -and $i -gt $firstCode -and $lines[$i].IndexOf("`t") -gt 0 -and -not $lines[$i].StartsWith('=== ')) { $firstValue = $i }
            if ($lines[$i].StartsWith('=== ')) { $lastHeader = $i }
        }
        if ($firstCode -lt 0 -or $firstValue -lt 0 -or $lastHeader -lt 1) {
            Write-Output 'FAIL: the control found no CODE record, no value or no last header to mutate'
            exit 1
        }
        # (a) one record's tag misspelled
        $a = @($lines | ForEach-Object { $_ })
        $a[$firstCode] = 'CODD' + $a[$firstCode].Substring(4)
        $pa = Join-Path $tmp 'a.txt'
        [System.IO.File]::WriteAllText($pa, ($a -join "`n"), $utf8)
        # (b) the last case dropped
        $b = @($lines[0..($lastHeader - 1)])
        $pb = Join-Path $tmp 'b.txt'
        [System.IO.File]::WriteAllText($pb, ($b -join "`n"), $utf8)
        # (c) a value echoed otherwise
        $c = @($lines | ForEach-Object { $_ })
        $c[$firstValue] = '7' + $c[$firstValue]
        $pc = Join-Path $tmp 'c.txt'
        [System.IO.File]::WriteAllText($pc, ($c -join "`n"), $utf8)
        # (d) a fixture value spelled with an exponent
        $fixLines = Get-FormatsLines $fixturePath
        $d = @($fixLines | ForEach-Object { $_ })
        $at = [Array]::IndexOf($d, '10000000000000000')
        if ($at -lt 0) { Write-Output 'FAIL: the control found no value to respell'; exit 1 }
        $d[$at] = '1E+16'
        $pd = Join-Path $tmp 'd.txt'
        [System.IO.File]::WriteAllText($pd, ($d -join "`n"), $utf8)
        # (e) the fixture with every backslash taken out of its codes
        $fixText = [System.IO.File]::ReadAllText($fixturePath)
        if ($fixText.IndexOf('\') -lt 0) { Write-Output 'FAIL: the control found no backslash in the fixture'; exit 1 }
        $pe = Join-Path $tmp 'e.txt'
        [System.IO.File]::WriteAllText($pe, $fixText.Replace('\', ''), $utf8)
        $ra = Test-FormatsGolden $fixturePath $pa $catalogPath
        $rb = Test-FormatsGolden $fixturePath $pb $catalogPath
        $rc = Test-FormatsGolden $fixturePath $pc $catalogPath
        $rd = Test-FormatsGolden $pd $goldenPath $catalogPath
        $re = Test-FormatsGolden $pe $goldenPath $catalogPath
        Write-Output ("control: the real set has {0} problem(s); the misspelt tag {1}, the dropped case {2}, the value echoed otherwise {3}, the exponent {4}, the backslashes taken out {5}" -f $real.Problems.Count, $ra.Problems.Count, $rb.Problems.Count, $rc.Problems.Count, $rd.Problems.Count, $re.Problems.Count)
        if ($real.Problems.Count -eq 0 -and $ra.Problems.Count -ge 1 -and $rb.Problems.Count -ge 1 -and $rc.Problems.Count -ge 1 -and $rd.Problems.Count -ge 1 -and $re.Problems.Count -ge 1) {
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

$r = Test-FormatsGolden $fixturePath $goldenPath $catalogPath
if ($r.Problems.Count -eq 0) {
    $above = ($r.Cases -gt $caseFloor) -or ($r.Values -gt $valueFloor) -or ($r.Refused -gt $refusedFloor)
    $note = if ($above) { ' - above a floor: raise the floors in this file' } else { '' }
    Write-Output ("OK: {0} cases, {1} values, {2} cases refused here; every one of the {3} constructs reached (floors {4}, {5}, {6}){7}" -f $r.Cases, $r.Values, $r.Refused, $constructs.Count, $caseFloor, $valueFloor, $refusedFloor, $note)
    exit 0
}
Write-Output ("FAIL: the formats golden has {0} problem(s):" -f $r.Problems.Count)
foreach ($p in $r.Problems) { Write-Output ('  ' + $p) }
exit 1
