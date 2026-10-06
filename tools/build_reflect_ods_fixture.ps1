<#
build_reflect_ods_fixture.ps1 - the OpenDocument twin of the reader's
fixture, written deterministically: scripts/reflect/opendocument.ods
(PORT.8, slice 8e).

WHAT IT BUILDS: an OpenDocument spreadsheet holding what
scripts/reflect/fixture.xlsx holds, cell for cell, written in ODF's own
terms, so that `frazaro reflect` reads it as the same relations and
`frazaro diff scripts/reflect/fixture.xlsx scripts/reflect/opendocument.ods`
is the format gap itself and nothing else (fixture_opendocument_diff.vla,
the treaty's oracle 9): Secret is hidden, not very hidden, since ODF has
no very-hidden; Broken is absent, since a name whose text is #REF! has no
ODF spelling worth keeping; E1 sums the Table's column as a plain range,
since ODF has no structured references; A6 is a date value, not a serial;
Local is workbook-level, not Model's, for Excel's sake (below).
The rest reads the same: the sheets Model, Data, Scratch (hidden), Secret
(hidden), Q1 Data, It's and Review (slice 8d's column cases); the named
expressions Rate = 0.2, HiddenName = 3, Local and Range1, all workbook-level
(Local is sheet-local in the package; ODF has sheet-local names too, inside
the sheet's own element, but Excel's reader of .ods keeps none: the owner's
opening of 2026-10-05 showed =Local as #NAME? and a repair, so the twin's
Local is workbook-level and the pair's diff carries the scope it lost, two
rows; the reader itself reads a sheet-local name, held by its own test);
the database range Sales over Data.A1:B4, which the reader prints
as the Table; every formula in OpenFormula (`of:=[.B1]*2`, `[.B:.B]`,
`['Q1 Data'.B2]`, `[Model.B1:Scratch.B1]`,
`['file:///C:/Rates.xlsx'#$Sheet1.A1]` (an absolute address: Excel's reader
drops a link with a relative one as unreadable content, the owner's third
opening of 2026-10-05 showed, and Excel itself writes `file:///` addresses;
the reader prints the file's name alone, so the goldens do not care),
`COM.MICROSOFT.SEQUENCE(2)` with matrix spans, `OFFSET([.B1];1;0)`) with
the cached value Excel computes; a boolean, a date, an error (Calc's
calcext marker), a string with a span, a formatted blank, a repeated empty
cell. The zip's first entry is the mimetype, stored and without an extra
field, as the format requires (a reader checks for the text at byte 38),
then the manifest, content.xml, a minimal styles.xml and meta.xml.

THE CONTAINER IS WRITTEN BY HAND, every entry stored: .NET's ZipArchive
deflates even at CompressionLevel.NoCompression (method 8 with stored
deflate blocks), which put a five-byte block header where Excel looks for
the mimetype and made Excel offer a repair on the owner's first opening
(2026-10-05). A local header, a central directory and the end record, with
CRC-32 computed here; nothing else is needed, and the bytes are the same
on every machine.

WHY A SCRIPT: as for build_reflect_fixture.ps1: a fixture of the treaty's
oracles 8, 9 and 10 that rebuilds byte for byte on the machine that made
it, regenerated rarely and on purpose. The owner reads the file in Excel,
which opens .ods, by eye.

NO OFFICE AUTOMATION: raw ODF, no Excel, no Calc. House style
(tools/*.ps1): PowerShell 5.1, host-free, no network.

Usage:  powershell -File tools\build_reflect_ods_fixture.ps1
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dir = Join-Path $root 'scripts\reflect'
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
$out = Join-Path $dir 'opendocument.ods'
$utf8 = New-Object System.Text.UTF8Encoding($false)
# The entries' stamp, 2026-10-05 00:00, in the zip's DOS date and time.
$dosDate = ((2026 - 1980) -shl 9) -bor (10 -shl 5) -bor 5
$dosTime = 0
$head = '<?xml version="1.0" encoding="UTF-8"?>' + "`n"
$ns = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
      'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
      'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
      'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" ' +
      'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" ' +
      'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" ' +
      'office:version="1.2"'

# A cell: a value with its type, a formula, repeats, matrix spans, a style.
function Cell([string]$type, [string]$value, [string]$text, [string]$formula = '', [string]$extra = '') {
    $a = ''
    if ($formula -ne '') { $a += ' table:formula="' + $formula + '"' }
    if ($type -ne '') { $a += ' office:value-type="' + $type + '"' }
    switch ($type) {
        'float'      { $a += ' office:value="' + $value + '"' }
        'percentage' { $a += ' office:value="' + $value + '"' }
        'boolean'    { $a += ' office:boolean-value="' + $value + '"' }
        'date'       { $a += ' office:date-value="' + $value + '"' }
        default      { }
    }
    if ($extra -ne '') { $a += ' ' + $extra }
    if ($text -eq '') { return '<table:table-cell' + $a + '/>' }
    return '<table:table-cell' + $a + '><text:p>' + $text + '</text:p></table:table-cell>'
}
function Num([string]$v, [string]$formula = '', [string]$extra = '') { return (Cell 'float' $v $v $formula $extra) }
function Str([string]$s, [string]$formula = '') { return (Cell 'string' '' $s $formula) }
function Empty([int]$n = 1) { if ($n -eq 1) { return '<table:table-cell/>' } return '<table:table-cell table:number-columns-repeated="' + $n + '"/>' }
function Row([string[]]$cells) { return '<table:table-row>' + ($cells -join '') + '</table:table-row>' }
function Sheet([string]$name, [string]$style, [string]$rows, [string]$tail = '') {
    return '<table:table table:name="' + $name + '" table:style-name="' + $style + '">' +
           '<table:table-column table:number-columns-repeated="8" table:default-cell-style-name="Default"/>' +
           $rows + $tail + '</table:table>'
}

$matrix2 = 'table:number-matrix-columns-spanned="1" table:number-matrix-rows-spanned="2"'

$model = (Row @(
        (Str 'Revenue'), (Num '1200'), (Num '2400' 'of:=[.B1]*2'), (Num '2480' 'of:=SUM([.B:.B])'),
        (Num '60' 'of:=SUM([Data.B2:Data.B4])'), (Num '1' 'of:=COM.MICROSOFT.SEQUENCE(2)' $matrix2),
        (Num '3' 'of:=HiddenName'), '<table:table-cell table:style-name="ce1"/>')) +
    (Row @(
        (Str 'Cost'), (Num '800'), (Num '1600' 'of:=[.B2]*2'), (Num '1200' 'of:=INDIRECT(&quot;B1&quot;)'),
        (Num '99' 'of:=[''file:///C:/Rates.xlsx''#$Sheet1.A1]'), (Num '2'), (Num '800' 'of:=Local'),
        (Str 'Revenue!' 'of:=[.A1]&amp;&quot;!&quot;'))) +
    (Row @(
        (Str 'Margin'), (Num '400' 'of:=[.B1]-[.B2]'), (Num '800' 'of:=[.B3]*2'), (Num '800' 'of:=OFFSET([.B1];1;0)'),
        (Num '2400' 'of:=[.B1:.B2]*2' $matrix2))) +
    (Row @(
        '<table:table-cell office:value-type="string"><text:p>Rate <text:span text:style-name="T1">applies</text:span></text:p></table:table-cell>',
        (Num '80' 'of:=[.B3]*Rate'), (Empty), (Num '1205' 'of:=SUM([Model.B1:Scratch.B1])'), (Num '1600'))) +
    (Row @(
        (Cell 'boolean' 'true' 'TRUE'), (Empty 2), (Num '17' 'of:=[''Q1 Data''.B2]+[''It''''s''.A1]'),
        '<table:table-cell table:formula="of:=1/0" office:value-type="string" calcext:value-type="error"><text:p>#DIV/0!</text:p></table:table-cell>')) +
    (Row @(
        '<table:table-cell table:style-name="ce2" office:value-type="date" office:date-value="2025-09-30"><text:p>2025-09-30</text:p></table:table-cell>'))
$data = (Row @((Str 'Item'), (Str 'Amount'))) +
    (Row @((Str 'pens'), (Num '10'))) +
    (Row @((Str 'ink'), (Num '20'))) +
    (Row @((Str 'paper'), (Num '30')))
$scratch = (Row @((Empty), (Num '5')))
$secret = (Row @((Str 'do not show')))
$q1 = '<table:table-row><table:table-cell table:number-columns-repeated="2"/></table:table-row>' +
    (Row @((Empty), (Num '10')))
$its = (Row @((Num '7')))
$review = (Row @((Num '10'), (Num '20' 'of:=[.A1]*2'), (Num '11' 'of:=[.A1]+1'), (Num '0' 'of:=[.Z9]'))) +
    (Row @((Num '20'), (Num '40' 'of:=[.A2]*2'), (Num '21' 'of:=[.A2]+1'), (Num '0' 'of:=[Model.H1]'))) +
    (Row @((Num '30'), (Num '61'), (Num '32' 'of:=[.A3]+2'))) +
    (Row @((Num '40'), (Num '80' 'of:=[.A4]*2'), (Num '41' 'of:=[.A4]+1'))) +
    (Row @((Num '50'), (Num '100' 'of:=[.A5]*2'), (Num '51' 'of:=[.A5]+1')))

$content = $head + '<office:document-content ' + $ns + '>' +
    '<office:scripts/><office:font-face-decls/>' +
    '<office:automatic-styles>' +
    '<number:date-style style:name="N49"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>' +
    '<style:style style:name="ta1" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb"/></style:style>' +
    '<style:style style:name="ta2" style:family="table" style:master-page-name="Default"><style:table-properties table:display="false" style:writing-mode="lr-tb"/></style:style>' +
    '<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#f7f4fc"/></style:style>' +
    '<style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N49"/>' +
    '<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>' +
    '</office:automatic-styles>' +
    '<office:body><office:spreadsheet>' +
    (Sheet 'Model' 'ta1' $model) +
    (Sheet 'Data' 'ta1' $data) +
    (Sheet 'Scratch' 'ta2' $scratch) +
    (Sheet 'Secret' 'ta2' $secret) +
    (Sheet 'Q1 Data' 'ta1' $q1) +
    (Sheet 'It''s' 'ta1' $its) +
    (Sheet 'Review' 'ta1' $review) +
    '<table:named-expressions>' +
    '<table:named-expression table:name="HiddenName" table:base-cell-address="$Model.$A$1" table:expression="of:=3"/>' +
    '<table:named-range table:name="Local" table:base-cell-address="$Model.$A$1" table:cell-range-address="$Model.$B$2"/>' +
    '<table:named-range table:name="Range1" table:base-cell-address="$Data.$A$1" table:cell-range-address="$Data.$A$2:.$B$4"/>' +
    '<table:named-expression table:name="Rate" table:base-cell-address="$Model.$A$1" table:expression="of:=0.2"/>' +
    '</table:named-expressions>' +
    '<table:database-ranges>' +
    '<table:database-range table:name="Sales" table:target-range-address="Data.A1:Data.B4" table:display-filter-buttons="true"/>' +
    '</table:database-ranges>' +
    '</office:spreadsheet></office:body></office:document-content>'

$manifest = $head + '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">' +
    '<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>' +
    '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>' +
    '</manifest:manifest>'
$styles = $head + '<office:document-styles ' + $ns + '>' +
    '<office:font-face-decls/>' +
    '<office:styles><style:default-style style:family="table-cell"><style:table-cell-properties style:decimal-places="2"/><style:text-properties fo:font-size="11pt"/></style:default-style>' +
    '<style:style style:name="Default" style:family="table-cell"/></office:styles>' +
    '<office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties style:writing-mode="lr-tb"/></style:page-layout></office:automatic-styles>' +
    '<office:master-styles><style:master-page style:name="Default" style:page-layout-name="pm1"/></office:master-styles>' +
    '</office:document-styles>'
$meta = $head + '<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.2">' +
    '<office:meta><meta:generator>Frazaro fixture</meta:generator><dc:date>2026-10-05T00:00:00</dc:date><meta:creation-date>2026-10-05T00:00:00</meta:creation-date></office:meta>' +
    '</office:document-meta>'

$parts = [ordered]@{}
$parts['mimetype'] = 'application/vnd.oasis.opendocument.spreadsheet'
$parts['META-INF/manifest.xml'] = $manifest
$parts['content.xml'] = $content
$parts['styles.xml'] = $styles
$parts['meta.xml'] = $meta

# --- the container, a stored zip written by hand (see the header) ---------
$crcTable = New-Object 'System.UInt32[]' 256
for ($n = 0; $n -lt 256; $n++) {
    [long]$c = $n
    for ($k = 0; $k -lt 8; $k++) {
        if ($c -band 1) { $c = 0xEDB88320L -bxor ($c -shr 1) } else { $c = $c -shr 1 }
    }
    $crcTable[$n] = [uint32]($c -band 0xFFFFFFFFL)
}
function Get-Crc32([byte[]]$data) {
    [long]$c = 0xFFFFFFFFL
    foreach ($b in $data) {
        $c = ([long]$crcTable[(($c -bxor $b) -band 0xFF)]) -bxor ($c -shr 8)
    }
    return [long](($c -bxor 0xFFFFFFFFL) -band 0xFFFFFFFFL)
}
function Add-Le16([System.Collections.Generic.List[byte]]$list, [long]$v) {
    $list.Add([byte]($v -band 0xFF)); $list.Add([byte](($v -shr 8) -band 0xFF))
}
function Add-Le32([System.Collections.Generic.List[byte]]$list, [long]$v) {
    for ($i = 0; $i -lt 4; $i++) { $list.Add([byte](($v -shr (8 * $i)) -band 0xFF)) }
}
$file = New-Object System.Collections.Generic.List[byte]
$central = New-Object System.Collections.Generic.List[byte]
foreach ($name in $parts.Keys) {
    $data = $utf8.GetBytes([string]$parts[$name])
    $nameBytes = [System.Text.Encoding]::ASCII.GetBytes($name)
    $crc = Get-Crc32 $data
    $offset = $file.Count
    # The local header: stored, no flags, no extra field.
    Add-Le32 $file 0x04034b50; Add-Le16 $file 20; Add-Le16 $file 0; Add-Le16 $file 0
    Add-Le16 $file $dosTime; Add-Le16 $file $dosDate; Add-Le32 $file $crc
    Add-Le32 $file $data.Length; Add-Le32 $file $data.Length
    Add-Le16 $file $nameBytes.Length; Add-Le16 $file 0
    $file.AddRange([byte[]]$nameBytes); $file.AddRange([byte[]]$data)
    # Its central directory entry.
    Add-Le32 $central 0x02014b50; Add-Le16 $central 20; Add-Le16 $central 20; Add-Le16 $central 0; Add-Le16 $central 0
    Add-Le16 $central $dosTime; Add-Le16 $central $dosDate; Add-Le32 $central $crc
    Add-Le32 $central $data.Length; Add-Le32 $central $data.Length
    Add-Le16 $central $nameBytes.Length; Add-Le16 $central 0; Add-Le16 $central 0
    Add-Le16 $central 0; Add-Le16 $central 0; Add-Le32 $central 0; Add-Le32 $central $offset
    $central.AddRange([byte[]]$nameBytes)
}
$cdOffset = $file.Count
$file.AddRange($central)
Add-Le32 $file 0x06054b50; Add-Le16 $file 0; Add-Le16 $file 0
Add-Le16 $file $parts.Count; Add-Le16 $file $parts.Count
Add-Le32 $file $central.Count; Add-Le32 $file $cdOffset; Add-Le16 $file 0
if (Test-Path $out) { Remove-Item $out -Force }
[System.IO.File]::WriteAllBytes($out, $file.ToArray())
Write-Output ("wrote {0} ({1} bytes, {2} parts, stored)" -f $out, (Get-Item $out).Length, $parts.Count)
