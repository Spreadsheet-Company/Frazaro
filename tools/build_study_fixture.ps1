<#
build_study_fixture.ps1 - the four workbooks of KERNEL.3's study (the study
before the viewport, SD-29), written deterministically into scripts/study/.

WHAT IT BUILDS: the two halves of the study's instrument (docs/PROTOCOL.md)
in one script, so that every number in them has one source. The 'find that'
half, find_a.xlsx and find_b.xlsx: a small model of one product's year with
five planted defects each, one of each class of Powell, Baker and Lawson's
field audit, in two isomorphic variants so that a participant who has
audited one has not seen the other's answers. The 'build this' half,
build_a.xlsx and build_b.xlsx: the workbook a participant is asked to
produce, with no defect.

  find_a and find_b, the sheets in tab order:
  Inputs   A1 Assumption, B1 Value; Unit price, Unit cost and Fixed cost
           per month in A2:B4 (find_a 40, 25, 1500; find_b 50, 30, 1200);
           the workbook names Price = Inputs!$B$2, UnitCost = Inputs!$B$3,
           FixedCost = Inputs!$B$4
  Sales    row 1 Month, Units, Price, Revenue, Unit cost, Cost, Margin,
           Margin %; rows 2 to 13 Jan to Dec, the units typed in B, and in
           every row C =Price, D =B*C, E =UnitCost, F =B*E, G =D-F, H =G/D;
           row 14 Total: B14 =SUM(B2:B13), D14, F14 and G14 the column
           sums, H14 =G14/D14; C14 and E14 absent
  Summary  A1 Measure, B1 Value; Total revenue =Sales!D14, Total cost
           =Sales!F14, Total margin =Sales!G14, Fixed costs for the year
           =FixedCost*12, Operating income =B2-B3-B5, in A2:B6

  The plants, the answer key of docs/PROTOCOL.md (section 3.2):
  find_a   logic: H =G/F in every row, the margin over the cost;
           reference: D14 =SUM(D2:D12), December left out;
           omission: F14 never written, so Summary!B3 reads an empty cell;
           hard-coding: D7 the constant 6500 where =B7*C7 gives 6800;
           copy and paste: G9 =D9-F8, August's margin takes July's cost
  find_b   logic: G =D+F in every row, a sign flipped;
           reference: D14 =SUM(D3:D13), January left out;
           omission: G14 never written, so Summary!B4 reads an empty cell;
           hard-coding: F10 the constant 4000 where =B10*E10 gives 4500;
           copy and paste: D4 =B3*C4, March's revenue takes February's units

  build_a and build_b, one sheet Plan: A1:D1 Month, Units, Price, Revenue;
  six months down the rows (build_a Jan to Jun with 120, 135, 150, 160,
  155, 170 at 40; build_b Jul to Dec with 130, 140, 150, 145, 160, 175 at
  50), the price a value in C, D =B*C; row 8 Total: B8 =SUM(B2:B7), D8
  =SUM(D2:D7), C8 absent.

THE CACHED VALUES are computed here, by this script's own arithmetic over
each formula's text as written (a small evaluator: the four operations,
SUM over a range, a defined name, a reference into another sheet), never
typed by hand, so that every formula carries the value Excel computes, the
effect of each plant included: a reference to an absent cell reads 0, SUM
adds the cells present and passes over text, a constant typed over a
formula is the constant, and a division by zero stops the script rather
than write a value Excel would not show. Doubles are written with
ToString('R') in the invariant culture and integers as integers; strings
go through sharedStrings.xml; a number sits in <v>; there is no style
beyond the default (cellXfs count 1); each sheet's <dimension> is its
cells' extent; docProps names the creator 'Frazaro study fixture'.

WHY A SCRIPT: a workbook made by hand could not be rebuilt byte for byte,
and these four are fixtures: their relations and the find half's audit, as
the door prints them, are the goldens beside them (scripts/study/
<name>_relations.vla, <name>_audit.vla), the answer key's cells quoted from
them, held by tools/check_study_fixture.ps1. The zip and document stamps
are fixed at 2026-10-08T00:00:00Z and the parts go in a fixed order, as
build_reflect_fixture.ps1 does; .NET's deflate output is what it is on the
machine that ran this, and the committed bytes are the fixtures, blessed by
the owner opening each in Excel, regenerated rarely and on purpose. Run it
under Windows PowerShell 5.1 only, as the usage line says: .NET Framework's
ToString('R') spells a double with Excel's own seventeen digits, where .NET
Core's (pwsh) prints the shortest round trip, so a regeneration under pwsh
would move five cached values in each find relations golden, not only the
zip bytes.

NO OFFICE AUTOMATION: raw OOXML, no Excel. House style (tools/*.ps1):
PowerShell 5.1, host-free, no network.

Usage:  powershell -File tools\build_study_fixture.ps1
#>
param()
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
$root = Split-Path -Parent $PSScriptRoot
$dir = Join-Path $root 'scripts\study'
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
$utf8 = New-Object System.Text.UTF8Encoding($false)
$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$stamp = New-Object DateTimeOffset(2026, 10, 8, 0, 0, 0, [TimeSpan]::Zero)
$date = '2026-10-08T00:00:00Z'
$head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`r`n"
$ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
$relBase = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
$ctBase = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'
$months = @('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')

# --- the model -----------------------------------------------------------------
# A cell is a hashtable: Kind 's' (a string, Text), 'n' (a number, Number)
# or 'f' (a formula, Formula without its '=', and Value once computed), with
# its Addr, Row and Col once placed. A sheet is an ordered table of cells by
# address; a book is its sheets in tab order and its defined names.
function New-Text([string]$text) { return @{ Kind = 's'; Text = $text } }
function New-Number([double]$number) { return @{ Kind = 'n'; Number = $number } }
function New-Formula([string]$formula) { return @{ Kind = 'f'; Formula = $formula; Value = $null; Busy = $false } }
function New-Book([string]$name) { return @{ Name = $name; Sheets = [ordered]@{}; Names = [ordered]@{} } }
function Add-Sheet($book, [string]$name) {
    $sheet = [ordered]@{}
    $book.Sheets[$name] = $sheet
    return $sheet
}

function Get-ColumnNumber([string]$letters) {
    $n = 0
    foreach ($ch in $letters.ToUpperInvariant().ToCharArray()) { $n = $n * 26 + ([int][char]$ch - 64) }
    return $n
}
function Get-ColumnLetters([int]$col) {
    $s = ''
    $n = $col
    while ($n -gt 0) {
        $r = ($n - 1) % 26
        $s = [string][char](65 + $r) + $s
        $n = [int][Math]::Floor(($n - 1) / 26)
    }
    return $s
}
# 'B2' or '$B$2' as its column number and row.
function Get-Address([string]$addr) {
    $m = [regex]::Match($addr, '^\$?([A-Za-z]{1,3})\$?([0-9]+)$')
    if (-not $m.Success) { throw "not a cell address: $addr" }
    return @{ Col = (Get-ColumnNumber $m.Groups[1].Value); Row = [int]$m.Groups[2].Value }
}
function Set-Cell($sheet, [string]$addr, $cell) {
    $a = Get-Address $addr
    $cell.Addr = $addr
    $cell.Row = $a.Row
    $cell.Col = $a.Col
    $sheet[$addr] = $cell
}
function Get-Sheet($book, [string]$name) {
    foreach ($k in $book.Sheets.Keys) {
        if ([string]::Equals($k, $name, [System.StringComparison]::OrdinalIgnoreCase)) { return $book.Sheets[$k] }
    }
    throw "no sheet named $name"
}

# --- the arithmetic: each formula's cached value, from its own text ------------
# The tokens of a formula after its '=': a number, a reference (with an
# optional sheet prefix), a name or function, and the marks ( ) , : + - * /.
$tokenRe = [regex]'\G(?:(?<ws>[ ]+)|(?<num>[0-9]+(?:\.[0-9]+)?)|(?<ref>(?:[A-Za-z_][A-Za-z0-9_]*!)?\$?[A-Z]{1,3}\$?[0-9]+(?![A-Za-z0-9_]))|(?<id>[A-Za-z_][A-Za-z0-9_.]*)|(?<op>[-+*/(),:]))'
function Get-FormulaTokens([string]$text) {
    $tokens = New-Object System.Collections.Generic.List[object]
    $at = 0
    while ($at -lt $text.Length) {
        $m = $tokenRe.Match($text, $at)
        if (-not $m.Success -or $m.Length -eq 0) { throw "the builder's arithmetic cannot read this formula at character $($at + 1): $text" }
        $at += $m.Length
        if ($m.Groups['ws'].Success) { continue }
        foreach ($kind in @('num', 'ref', 'id', 'op')) {
            if ($m.Groups[$kind].Success) { $tokens.Add(@{ Kind = $kind; Text = $m.Groups[$kind].Value }); break }
        }
    }
    return ,$tokens
}
# The parser's state: the tokens, the position, the book, and the sheet and
# cell the formula sits in (for relative sheet names and for messages).
function Get-Token($st) {
    if ($st.Pos -lt $st.Tokens.Count) { return $st.Tokens[$st.Pos] }
    return $null
}
function Read-Token($st) {
    $t = Get-Token $st
    if ($null -eq $t) { throw "the formula of $($st.Sheet)!$($st.Cell) ends early" }
    $st.Pos = $st.Pos + 1
    return $t
}
function Test-Mark($st, [string]$mark) {
    $t = Get-Token $st
    return ($null -ne $t -and $t.Kind -eq 'op' -and $t.Text -eq $mark)
}
function Read-Mark($st, [string]$mark) {
    if (-not (Test-Mark $st $mark)) { throw "expected '$mark' in the formula of $($st.Sheet)!$($st.Cell)" }
    $null = Read-Token $st
}
function Invoke-Expr($st) {
    $v = Invoke-Term $st
    while ((Test-Mark $st '+') -or (Test-Mark $st '-')) {
        $op = (Read-Token $st).Text
        $w = Invoke-Term $st
        if ($op -eq '+') { $v = $v + $w } else { $v = $v - $w }
    }
    return $v
}
function Invoke-Term($st) {
    $v = Invoke-Factor $st
    while ((Test-Mark $st '*') -or (Test-Mark $st '/')) {
        $op = (Read-Token $st).Text
        $w = Invoke-Factor $st
        if ($op -eq '*') {
            $v = $v * $w
        } else {
            if ($w -eq 0) { throw "division by zero at $($st.Sheet)!$($st.Cell): Excel would show #DIV/0!, which this builder does not write" }
            $v = $v / $w
        }
    }
    return $v
}
function Invoke-Factor($st) {
    $t = Read-Token $st
    if ($t.Kind -eq 'num') { return [double]::Parse($t.Text, $invariant) }
    if ($t.Kind -eq 'ref') {
        if (Test-Mark $st ':') { throw "a range outside SUM in the formula of $($st.Sheet)!$($st.Cell)" }
        return Get-ReferenceNumber $st $t.Text
    }
    if ($t.Kind -eq 'id') {
        if (Test-Mark $st '(') {
            $null = Read-Token $st
            return Invoke-Function $st $t.Text
        }
        return Get-NameNumber $st $t.Text
    }
    if ($t.Text -eq '(') {
        $v = Invoke-Expr $st
        Read-Mark $st ')'
        return $v
    }
    if ($t.Text -eq '-') { return -(Invoke-Factor $st) }
    throw "unexpected '$($t.Text)' in the formula of $($st.Sheet)!$($st.Cell)"
}
# SUM alone: each argument a range (the cells present, text passed over) or
# a value.
function Invoke-Function($st, [string]$name) {
    if ($name -ne 'SUM') { throw "the builder's arithmetic knows SUM alone, not $name, in the formula of $($st.Sheet)!$($st.Cell)" }
    $total = [double]0
    while ($true) {
        $t = Get-Token $st
        $isRange = $false
        if ($null -ne $t -and $t.Kind -eq 'ref' -and ($st.Pos + 1) -lt $st.Tokens.Count) {
            $colon = $st.Tokens[$st.Pos + 1]
            if ($colon.Kind -eq 'op' -and $colon.Text -eq ':') { $isRange = $true }
        }
        if ($isRange) {
            $null = Read-Token $st
            $null = Read-Token $st
            $u = Read-Token $st
            if ($u.Kind -ne 'ref') { throw "a range without its end in the formula of $($st.Sheet)!$($st.Cell)" }
            $total = $total + (Get-RangeSum $st $t.Text $u.Text)
        } else {
            $total = $total + (Invoke-Expr $st)
        }
        if (Test-Mark $st ',') {
            $null = Read-Token $st
            continue
        }
        Read-Mark $st ')'
        return $total
    }
}
# A reference's sheet (the formula's own without a prefix) and its address
# without dollar signs.
function Split-Reference($st, [string]$ref) {
    $sheetName = $st.Sheet
    $addr = $ref
    $bang = $ref.IndexOf('!')
    if ($bang -ge 0) {
        $sheetName = $ref.Substring(0, $bang)
        $addr = $ref.Substring($bang + 1)
    }
    $addr = $addr -replace '\$', ''
    return @{ Sheet = $sheetName; Addr = $addr }
}
# A single cell's number: 0 for a cell the sheet does not hold, as Excel
# reads an empty cell in arithmetic; a text cell stops the script, since
# Excel would show #VALUE!.
function Get-ReferenceNumber($st, [string]$ref) {
    $r = Split-Reference $st $ref
    $sheet = Get-Sheet $st.Book $r.Sheet
    if (-not $sheet.Contains($r.Addr)) { return [double]0 }
    $cell = $sheet[$r.Addr]
    if ($cell.Kind -eq 's') { throw "the formula of $($st.Sheet)!$($st.Cell) reads the text cell $($r.Sheet)!$($r.Addr): Excel would show #VALUE!" }
    return Get-CellNumber $st.Book $r.Sheet $cell
}
# A held cell's number: its constant, or its formula's value computed once.
function Get-CellNumber($book, [string]$sheetName, $cell) {
    if ($cell.Kind -eq 'n') { return $cell.Number }
    if ($null -ne $cell.Value) { return $cell.Value }
    if ($cell.Busy) { throw "a circular reference through $sheetName!$($cell.Addr)" }
    $cell.Busy = $true
    $cell.Value = Invoke-Formula $book $sheetName $cell.Addr $cell.Formula
    $cell.Busy = $false
    return $cell.Value
}
function Get-NameNumber($st, [string]$name) {
    foreach ($k in $st.Book.Names.Keys) {
        if ([string]::Equals($k, $name, [System.StringComparison]::OrdinalIgnoreCase)) {
            return Invoke-Formula $st.Book $st.Sheet ("the name " + $k) $st.Book.Names[$k]
        }
    }
    throw "no defined name $name for the formula of $($st.Sheet)!$($st.Cell)"
}
function Get-RangeSum($st, [string]$from, [string]$to) {
    $a = Split-Reference $st $from
    $b = Split-Reference $st $to
    if ($to.IndexOf('!') -ge 0 -and $b.Sheet -ne $a.Sheet) { throw "a range across sheets in the formula of $($st.Sheet)!$($st.Cell)" }
    $sheet = Get-Sheet $st.Book $a.Sheet
    $p = Get-Address $a.Addr
    $q = Get-Address $b.Addr
    $total = [double]0
    for ($row = [Math]::Min($p.Row, $q.Row); $row -le [Math]::Max($p.Row, $q.Row); $row++) {
        for ($col = [Math]::Min($p.Col, $q.Col); $col -le [Math]::Max($p.Col, $q.Col); $col++) {
            $addr = (Get-ColumnLetters $col) + [string]$row
            if (-not $sheet.Contains($addr)) { continue }
            $cell = $sheet[$addr]
            if ($cell.Kind -eq 's') { continue }
            $total = $total + (Get-CellNumber $st.Book $a.Sheet $cell)
        }
    }
    return $total
}
function Invoke-Formula($book, [string]$sheetName, [string]$addr, [string]$text) {
    $st = @{ Tokens = (Get-FormulaTokens $text); Pos = 0; Book = $book; Sheet = $sheetName; Cell = $addr }
    $v = Invoke-Expr $st
    if ($null -ne (Get-Token $st)) { throw "the formula of $sheetName!$addr has more after its end: $text" }
    return [double]$v
}
# Every formula of the book evaluated, in tab and document order.
function Resolve-Book($book) {
    foreach ($sheetName in @($book.Sheets.Keys)) {
        $sheet = $book.Sheets[$sheetName]
        foreach ($addr in @($sheet.Keys)) {
            $cell = $sheet[$addr]
            if ($cell.Kind -eq 'f') { $null = Get-CellNumber $book $sheetName $cell }
        }
    }
}

# --- the models ----------------------------------------------------------------
# A find model from its product's numbers, its per-row margin formulas (with
# {r} the row), its row-14 totals ($null for one never written) and its
# plants, each a cell set over what the rows wrote.
function New-FindBook($spec) {
    $book = New-Book $spec.Name
    $inputs = Add-Sheet $book 'Inputs'
    Set-Cell $inputs 'A1' (New-Text 'Assumption')
    Set-Cell $inputs 'B1' (New-Text 'Value')
    Set-Cell $inputs 'A2' (New-Text 'Unit price')
    Set-Cell $inputs 'B2' (New-Number $spec.Price)
    Set-Cell $inputs 'A3' (New-Text 'Unit cost')
    Set-Cell $inputs 'B3' (New-Number $spec.UnitCost)
    Set-Cell $inputs 'A4' (New-Text 'Fixed cost per month')
    Set-Cell $inputs 'B4' (New-Number $spec.Fixed)
    $book.Names['FixedCost'] = 'Inputs!$B$4'
    $book.Names['Price'] = 'Inputs!$B$2'
    $book.Names['UnitCost'] = 'Inputs!$B$3'
    $sales = Add-Sheet $book 'Sales'
    $heads = @('Month', 'Units', 'Price', 'Revenue', 'Unit cost', 'Cost', 'Margin', 'Margin %')
    for ($i = 0; $i -lt $heads.Count; $i++) { Set-Cell $sales ((Get-ColumnLetters ($i + 1)) + '1') (New-Text $heads[$i]) }
    for ($i = 0; $i -lt 12; $i++) {
        $r = [string]($i + 2)
        Set-Cell $sales ('A' + $r) (New-Text $months[$i])
        Set-Cell $sales ('B' + $r) (New-Number $spec.Units[$i])
        Set-Cell $sales ('C' + $r) (New-Formula 'Price')
        Set-Cell $sales ('D' + $r) (New-Formula ('B' + $r + '*C' + $r))
        Set-Cell $sales ('E' + $r) (New-Formula 'UnitCost')
        Set-Cell $sales ('F' + $r) (New-Formula ('B' + $r + '*E' + $r))
        Set-Cell $sales ('G' + $r) (New-Formula ($spec.Margin.Replace('{r}', $r)))
        Set-Cell $sales ('H' + $r) (New-Formula ($spec.MarginPct.Replace('{r}', $r)))
    }
    Set-Cell $sales 'A14' (New-Text 'Total')
    Set-Cell $sales 'B14' (New-Formula 'SUM(B2:B13)')
    if ($null -ne $spec.RevenueTotal) { Set-Cell $sales 'D14' (New-Formula $spec.RevenueTotal) }
    if ($null -ne $spec.CostTotal) { Set-Cell $sales 'F14' (New-Formula $spec.CostTotal) }
    if ($null -ne $spec.MarginTotal) { Set-Cell $sales 'G14' (New-Formula $spec.MarginTotal) }
    Set-Cell $sales 'H14' (New-Formula 'G14/D14')
    foreach ($p in $spec.Plants) { Set-Cell $sales $p.Cell $p.Is }
    $summary = Add-Sheet $book 'Summary'
    Set-Cell $summary 'A1' (New-Text 'Measure')
    Set-Cell $summary 'B1' (New-Text 'Value')
    Set-Cell $summary 'A2' (New-Text 'Total revenue')
    Set-Cell $summary 'B2' (New-Formula 'Sales!D14')
    Set-Cell $summary 'A3' (New-Text 'Total cost')
    Set-Cell $summary 'B3' (New-Formula 'Sales!F14')
    Set-Cell $summary 'A4' (New-Text 'Total margin')
    Set-Cell $summary 'B4' (New-Formula 'Sales!G14')
    Set-Cell $summary 'A5' (New-Text 'Fixed costs for the year')
    Set-Cell $summary 'B5' (New-Formula 'FixedCost*12')
    Set-Cell $summary 'A6' (New-Text 'Operating income')
    Set-Cell $summary 'B6' (New-Formula 'B2-B3-B5')
    return $book
}
# A build model: the six months from $firstMonth (0 = Jan), their units,
# one price typed in every row.
function New-BuildBook([string]$name, [int]$firstMonth, [double[]]$units, [double]$price) {
    $book = New-Book $name
    $plan = Add-Sheet $book 'Plan'
    $heads = @('Month', 'Units', 'Price', 'Revenue')
    for ($i = 0; $i -lt $heads.Count; $i++) { Set-Cell $plan ((Get-ColumnLetters ($i + 1)) + '1') (New-Text $heads[$i]) }
    for ($i = 0; $i -lt 6; $i++) {
        $r = [string]($i + 2)
        Set-Cell $plan ('A' + $r) (New-Text $months[$firstMonth + $i])
        Set-Cell $plan ('B' + $r) (New-Number $units[$i])
        Set-Cell $plan ('C' + $r) (New-Number $price)
        Set-Cell $plan ('D' + $r) (New-Formula ('B' + $r + '*C' + $r))
    }
    Set-Cell $plan 'A8' (New-Text 'Total')
    Set-Cell $plan 'B8' (New-Formula 'SUM(B2:B7)')
    Set-Cell $plan 'D8' (New-Formula 'SUM(D2:D7)')
    return $book
}

$findA = @{
    Name = 'find_a'; Price = 40; UnitCost = 25; Fixed = 1500
    Units = @(120, 135, 150, 160, 155, 170, 180, 175, 165, 190, 200, 210)
    Margin = 'D{r}-F{r}'
    MarginPct = 'G{r}/F{r}'          # the logic error: the margin over the cost, every row (=G/D is right)
    RevenueTotal = 'SUM(D2:D12)'     # the wrong reference: December left out
    CostTotal = $null                # the omission: the cost total never written
    MarginTotal = 'SUM(G2:G13)'
    Plants = @(
        @{ Cell = 'D7'; Is = (New-Number 6500) },     # the hard-coded constant over =B7*C7 (170*40 = 6800)
        @{ Cell = 'G9'; Is = (New-Formula 'D9-F8') }  # the copy-and-paste fault: August's margin takes July's cost
    )
}
$findB = @{
    Name = 'find_b'; Price = 50; UnitCost = 30; Fixed = 1200
    Units = @(90, 95, 110, 105, 120, 130, 125, 140, 150, 145, 160, 170)
    Margin = 'D{r}+F{r}'             # the logic error: a sign flipped, every row (=D-F is right)
    MarginPct = 'G{r}/D{r}'
    RevenueTotal = 'SUM(D3:D13)'     # the wrong reference: January left out
    CostTotal = 'SUM(F2:F13)'
    MarginTotal = $null              # the omission: the margin total never written
    Plants = @(
        @{ Cell = 'F10'; Is = (New-Number 4000) },    # the hard-coded constant over =B10*E10 (150*30 = 4500)
        @{ Cell = 'D4'; Is = (New-Formula 'B3*C4') }  # the copy-and-paste fault: March's revenue takes February's units
    )
}

# --- the package ---------------------------------------------------------------
function Format-Number([double]$v) {
    if ($v -eq [Math]::Floor($v) -and [Math]::Abs($v) -lt 1e15) { return ([int64]$v).ToString($invariant) }
    return $v.ToString('R', $invariant)
}
function Format-Xml([string]$s) {
    return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&apos;')
}
# A string's index in the shared table, added on first use; every use counted.
function Get-StringIndex($strings, [string]$text) {
    $strings.Uses = $strings.Uses + 1
    $i = $strings.List.IndexOf($text)
    if ($i -ge 0) { return $i }
    $strings.List.Add($text)
    return $strings.List.Count - 1
}
# One worksheet part: its cells in document order, rows grouped, the
# dimension their extent.
function Get-SheetXml($sheet, [bool]$first, $strings) {
    $cells = @($sheet.Values | Sort-Object -Property @{ Expression = { $_.Row } }, @{ Expression = { $_.Col } })
    if ($cells.Count -eq 0) { throw 'a sheet with no cell' }
    $minRow = $cells[0].Row; $maxRow = $cells[0].Row; $minCol = $cells[0].Col; $maxCol = $cells[0].Col
    foreach ($c in $cells) {
        if ($c.Row -lt $minRow) { $minRow = $c.Row }
        if ($c.Row -gt $maxRow) { $maxRow = $c.Row }
        if ($c.Col -lt $minCol) { $minCol = $c.Col }
        if ($c.Col -gt $maxCol) { $maxCol = $c.Col }
    }
    $dimension = (Get-ColumnLetters $minCol) + [string]$minRow
    if ($minRow -ne $maxRow -or $minCol -ne $maxCol) { $dimension = $dimension + ':' + (Get-ColumnLetters $maxCol) + [string]$maxRow }
    $view = if ($first) { '<sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView>' } else { '<sheetView workbookViewId="0"/>' }
    $sb = New-Object System.Text.StringBuilder
    $null = $sb.Append($head + '<worksheet ' + $ns + '>' + '<dimension ref="' + $dimension + '"/>' + '<sheetViews>' + $view + '</sheetViews>' + '<sheetFormatPr defaultRowHeight="15"/>' + '<sheetData>')
    $current = 0
    foreach ($c in $cells) {
        if ($c.Row -ne $current) {
            if ($current -ne 0) { $null = $sb.Append('</row>') }
            $current = $c.Row
            $null = $sb.Append('<row r="' + [string]$c.Row + '">')
        }
        if ($c.Kind -eq 's') {
            $null = $sb.Append('<c r="' + $c.Addr + '" t="s"><v>' + [string](Get-StringIndex $strings $c.Text) + '</v></c>')
        } elseif ($c.Kind -eq 'n') {
            $null = $sb.Append('<c r="' + $c.Addr + '"><v>' + (Format-Number $c.Number) + '</v></c>')
        } else {
            if ($null -eq $c.Value) { throw "the formula of $($c.Addr) was never computed" }
            $null = $sb.Append('<c r="' + $c.Addr + '"><f>' + (Format-Xml $c.Formula) + '</f><v>' + (Format-Number $c.Value) + '</v></c>')
        }
    }
    if ($current -ne 0) { $null = $sb.Append('</row>') }
    $null = $sb.Append('</sheetData>' + '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' + '</worksheet>')
    return $sb.ToString()
}
# The whole package, its parts in a fixed order, deflated, every entry
# stamped alike.
function Write-Xlsx($book, [string]$out) {
    $sheetNames = @($book.Sheets.Keys)
    $n = $sheetNames.Count
    $strings = @{ List = (New-Object System.Collections.Generic.List[string]); Uses = 0 }
    $sheetXml = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $n; $i++) { $sheetXml.Add((Get-SheetXml $book.Sheets[$sheetNames[$i]] ($i -eq 0) $strings)) }
    $parts = [ordered]@{}
    $parts['[Content_Types].xml'] = $head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="' + $ctBase + 'sheet.main+xml"/>' +
        ((1..$n | ForEach-Object { '<Override PartName="/xl/worksheets/sheet' + $_ + '.xml" ContentType="' + $ctBase + 'worksheet+xml"/>' }) -join '') +
        '<Override PartName="/xl/styles.xml" ContentType="' + $ctBase + 'styles+xml"/>' +
        '<Override PartName="/xl/sharedStrings.xml" ContentType="' + $ctBase + 'sharedStrings+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
        '</Types>'
    $parts['_rels/.rels'] = $head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="' + $relBase + 'officeDocument" Target="xl/workbook.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '<Relationship Id="rId3" Type="' + $relBase + 'extended-properties" Target="docProps/app.xml"/>' +
        '</Relationships>'
    $definedNames = ''
    if ($book.Names.Count -gt 0) {
        $definedNames = '<definedNames>' + (($book.Names.Keys | ForEach-Object { '<definedName name="' + $_ + '">' + (Format-Xml $book.Names[$_]) + '</definedName>' }) -join '') + '</definedNames>'
    }
    $parts['xl/workbook.xml'] = $head + '<workbook ' + $ns + '>' +
        '<workbookPr defaultThemeVersion="166925"/>' +
        '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12225" activeTab="0"/></bookViews>' +
        '<sheets>' +
        ((1..$n | ForEach-Object { '<sheet name="' + (Format-Xml $sheetNames[$_ - 1]) + '" sheetId="' + $_ + '" r:id="rId' + $_ + '"/>' }) -join '') +
        '</sheets>' +
        $definedNames +
        '<calcPr calcId="191029"/>' +
        '</workbook>'
    $parts['xl/_rels/workbook.xml.rels'] = $head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        ((1..$n | ForEach-Object { '<Relationship Id="rId' + $_ + '" Type="' + $relBase + 'worksheet" Target="worksheets/sheet' + $_ + '.xml"/>' }) -join '') +
        '<Relationship Id="rId' + ($n + 1) + '" Type="' + $relBase + 'styles" Target="styles.xml"/>' +
        '<Relationship Id="rId' + ($n + 2) + '" Type="' + $relBase + 'sharedStrings" Target="sharedStrings.xml"/>' +
        '</Relationships>'
    for ($i = 0; $i -lt $n; $i++) { $parts['xl/worksheets/sheet' + ($i + 1) + '.xml'] = $sheetXml[$i] }
    $parts['xl/sharedStrings.xml'] = $head + '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + $strings.Uses + '" uniqueCount="' + $strings.List.Count + '">' +
        (($strings.List | ForEach-Object { '<si><t>' + (Format-Xml $_) + '</t></si>' }) -join '') +
        '</sst>'
    $parts['xl/styles.xml'] = $head + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>' +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
        '<dxfs count="0"/>' +
        '</styleSheet>'
    $parts['docProps/core.xml'] = $head + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        '<dc:creator>Frazaro study fixture</dc:creator><cp:lastModifiedBy>Frazaro study fixture</cp:lastModifiedBy>' +
        '<dcterms:created xsi:type="dcterms:W3CDTF">' + $date + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + $date + '</dcterms:modified>' +
        '</cp:coreProperties>'
    $parts['docProps/app.xml'] = $head + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
        '<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>' +
        '<Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion>' +
        '</Properties>'
    if (Test-Path $out) { Remove-Item $out -Force }
    $fs = [System.IO.File]::Open($out, [System.IO.FileMode]::CreateNew)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($partName in $parts.Keys) {
                $e = $zip.CreateEntry($partName, [System.IO.Compression.CompressionLevel]::Optimal)
                $e.LastWriteTime = $stamp
                $bytes = $utf8.GetBytes([string]$parts[$partName])
                $s = $e.Open(); $s.Write($bytes, 0, $bytes.Length); $s.Dispose()
            }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
    Write-Output ("wrote {0} ({1} bytes, {2} parts, deflated)" -f $out, (Get-Item $out).Length, $parts.Count)
}

$books = @(
    (New-FindBook $findA),
    (New-FindBook $findB),
    (New-BuildBook 'build_a' 0 @(120, 135, 150, 160, 155, 170) 40),
    (New-BuildBook 'build_b' 6 @(130, 140, 150, 145, 160, 175) 50)
)
foreach ($book in $books) {
    Resolve-Book $book
    Write-Xlsx $book (Join-Path $dir ($book.Name + '.xlsx'))
}
