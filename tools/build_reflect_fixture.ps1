<#
build_reflect_fixture.ps1 - the workbook the reader is held to, written
deterministically: scripts/reflect/fixture.xlsx (PORT.8, slice 8a).

WHAT IT BUILDS: a package of Excel's shape, deflated throughout, holding
every case the reader's walker has an arm for, so that one fixture and its
expected relations (scripts/reflect/fixture_relations.vla, the core's own
output, blessed by the owner opening this file in Excel and reading each
cell by eye) pin the whole walk:

  Model (visible)
    A1 Revenue        B1 1200   C1 =B1*2 (shared, C1:C3)   D1 =SUM(B:B) 2480
    A2 Cost           B2 800    C2 (shared child) 1600      D2 =INDIRECT("B1") 1200
    A3 Margin         B3 =B1-B2 400   C3 (shared child) 800  D3 =OFFSET(B1,1,0) 800
    A4 Rate applies (rich text)  B4 =B3*Rate 80              D4 =SUM(Model:Scratch!B1) 1205
    A5 TRUE                                                  D5 ='Q1 Data'!B2+'It''s'!A1 17
    A6 a date, the serial 45930 under a date format
    E1 =SUM(Sales[Amount]) 60    E2 =[1]Sheet1!A1 99 (a link to Rates.xlsx)
    E3 {=B1:B2*2} (a legacy array formula over E3:E4) 2400 / 1600
    E5 =1/0 #DIV/0!              F1 =SEQUENCE(2) (a dynamic array over F1:F2) 1 / 2
    G1 =HiddenName 3             G2 =Local 800                H1 a formatted blank
    H2 =A1&"!" "Revenue!"
  Data (visible): the Table Sales over A1:B4 (Item, Amount; pens 10, ink 20, paper 30)
  Scratch (hidden): B1 5
  Secret (very hidden): A1 do not show
  Q1 Data (visible): B2 10
  It's (visible): A1 7

and the defined names Rate = 0.2, HiddenName = 3 (hidden), Local =
Model!$B$2 (scoped to Model), Range1 = Data!$A$2:$B$4, Broken = Model!#REF!,
and Excel's own placeholder _xlfn.SINGLE = #NAME? (hidden), which the reader
counts apart and prints in no row. The cached values are the ones Excel
would compute, so that a `cell` row exists for every formula.

WHY A SCRIPT: a workbook made by hand could not be rebuilt byte for byte,
and this one is a fixture of the treaty's oracle 8 (conformance/README.md):
its relations, read by any implementation, are fixture_relations.vla. The
zip stamps are fixed and the parts go in a fixed order, as
build_model_fixture.ps1 does; .NET's deflate output is what it is on the
machine that ran this, and the committed bytes are the fixture, regenerated
rarely and on purpose.

NO OFFICE AUTOMATION: raw OOXML, no Excel. House style (tools/*.ps1):
PowerShell 5.1, host-free, no network.

Usage:  powershell -File tools\build_reflect_fixture.ps1
#>
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
$root = Split-Path -Parent $PSScriptRoot
$dir = Join-Path $root 'scripts\reflect'
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
$out = Join-Path $dir 'fixture.xlsx'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$stamp = New-Object DateTimeOffset(2026, 10, 3, 0, 0, 0, [TimeSpan]::Zero)
$head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`r`n"
$ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
$relBase = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
$ctBase = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'

$parts = [ordered]@{}
$parts['[Content_Types].xml'] = $head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="' + $ctBase + 'sheet.main+xml"/>' +
    ((1..6 | ForEach-Object { '<Override PartName="/xl/worksheets/sheet' + $_ + '.xml" ContentType="' + $ctBase + 'worksheet+xml"/>' }) -join '') +
    '<Override PartName="/xl/tables/table1.xml" ContentType="' + $ctBase + 'table+xml"/>' +
    '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="' + $ctBase + 'externalLink+xml"/>' +
    '<Override PartName="/xl/metadata.xml" ContentType="' + $ctBase + 'sheetMetadata+xml"/>' +
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
$parts['xl/workbook.xml'] = $head + '<workbook ' + $ns + '>' +
    '<workbookPr defaultThemeVersion="166925"/>' +
    '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12225" activeTab="0"/></bookViews>' +
    '<sheets>' +
    '<sheet name="Model" sheetId="1" r:id="rId1"/>' +
    '<sheet name="Data" sheetId="2" r:id="rId2"/>' +
    '<sheet name="Scratch" sheetId="3" state="hidden" r:id="rId3"/>' +
    '<sheet name="Secret" sheetId="4" state="veryHidden" r:id="rId4"/>' +
    '<sheet name="Q1 Data" sheetId="5" r:id="rId5"/>' +
    '<sheet name="It''s" sheetId="6" r:id="rId6"/>' +
    '</sheets>' +
    '<externalReferences><externalReference r:id="rId9"/></externalReferences>' +
    '<definedNames>' +
    '<definedName name="_xlfn.SINGLE" hidden="1">#NAME?</definedName>' +
    '<definedName name="Broken">Model!#REF!</definedName>' +
    '<definedName name="HiddenName" hidden="1">3</definedName>' +
    '<definedName name="Local" localSheetId="0">Model!$B$2</definedName>' +
    '<definedName name="Range1">Data!$A$2:$B$4</definedName>' +
    '<definedName name="Rate">0.2</definedName>' +
    '</definedNames>' +
    '<calcPr calcId="191029"/>' +
    '</workbook>'
$parts['xl/_rels/workbook.xml.rels'] = $head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    ((1..6 | ForEach-Object { '<Relationship Id="rId' + $_ + '" Type="' + $relBase + 'worksheet" Target="worksheets/sheet' + $_ + '.xml"/>' }) -join '') +
    '<Relationship Id="rId7" Type="' + $relBase + 'styles" Target="styles.xml"/>' +
    '<Relationship Id="rId8" Type="' + $relBase + 'sharedStrings" Target="sharedStrings.xml"/>' +
    '<Relationship Id="rId9" Type="' + $relBase + 'externalLink" Target="externalLinks/externalLink1.xml"/>' +
    '<Relationship Id="rId10" Type="' + $relBase + 'sheetMetadata" Target="metadata.xml"/>' +
    '</Relationships>'
$margins = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
$parts['xl/worksheets/sheet1.xml'] = $head + '<worksheet ' + $ns + '>' +
    '<dimension ref="A1:H6"/>' +
    '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="15"/>' +
    '<sheetData>' +
    '<row r="1">' +
      '<c r="A1" t="s"><v>0</v></c>' +
      '<c r="B1"><v>1200</v></c>' +
      '<c r="C1"><f t="shared" ref="C1:C3" si="0">B1*2</f><v>2400</v></c>' +
      '<c r="D1"><f>SUM(B:B)</f><v>2480</v></c>' +
      '<c r="E1"><f>SUM(Sales[Amount])</f><v>60</v></c>' +
      '<c r="F1" cm="1"><f t="array" ref="F1:F2">_xlfn.SEQUENCE(2)</f><v>1</v></c>' +
      '<c r="G1"><f>HiddenName</f><v>3</v></c>' +
      '<c r="H1" s="1"/>' +
    '</row>' +
    '<row r="2">' +
      '<c r="A2" t="s"><v>1</v></c>' +
      '<c r="B2"><v>800</v></c>' +
      '<c r="C2"><f t="shared" si="0"/><v>1600</v></c>' +
      '<c r="D2"><f>INDIRECT("B1")</f><v>1200</v></c>' +
      '<c r="E2"><f>[1]Sheet1!A1</f><v>99</v></c>' +
      '<c r="F2"><v>2</v></c>' +
      '<c r="G2"><f>Local</f><v>800</v></c>' +
      '<c r="H2" t="str"><f>A1&amp;"!"</f><v>Revenue!</v></c>' +
    '</row>' +
    '<row r="3">' +
      '<c r="A3" t="s"><v>2</v></c>' +
      '<c r="B3"><f>B1-B2</f><v>400</v></c>' +
      '<c r="C3"><f t="shared" si="0"/><v>800</v></c>' +
      '<c r="D3"><f>OFFSET(B1,1,0)</f><v>800</v></c>' +
      '<c r="E3"><f t="array" ref="E3:E4">B1:B2*2</f><v>2400</v></c>' +
    '</row>' +
    '<row r="4">' +
      '<c r="A4" t="s"><v>3</v></c>' +
      '<c r="B4"><f>B3*Rate</f><v>80</v></c>' +
      '<c r="D4"><f>SUM(Model:Scratch!B1)</f><v>1205</v></c>' +
      '<c r="E4"><v>1600</v></c>' +
    '</row>' +
    '<row r="5">' +
      '<c r="A5" t="b"><v>1</v></c>' +
      '<c r="D5"><f>''Q1 Data''!B2+''It''''s''!A1</f><v>17</v></c>' +
      '<c r="E5" t="e"><f>1/0</f><v>#DIV/0!</v></c>' +
    '</row>' +
    '<row r="6">' +
      '<c r="A6" s="2"><v>45930</v></c>' +
    '</row>' +
    '</sheetData>' + $margins + '</worksheet>'
$parts['xl/worksheets/sheet2.xml'] = $head + '<worksheet ' + $ns + '>' +
    '<dimension ref="A1:B4"/>' +
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="15"/>' +
    '<sheetData>' +
    '<row r="1"><c r="A1" t="s"><v>4</v></c><c r="B1" t="s"><v>5</v></c></row>' +
    '<row r="2"><c r="A2" t="s"><v>6</v></c><c r="B2"><v>10</v></c></row>' +
    '<row r="3"><c r="A3" t="s"><v>7</v></c><c r="B3"><v>20</v></c></row>' +
    '<row r="4"><c r="A4" t="s"><v>8</v></c><c r="B4"><v>30</v></c></row>' +
    '</sheetData>' + $margins +
    '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>' +
    '</worksheet>'
$parts['xl/worksheets/_rels/sheet2.xml.rels'] = $head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="' + $relBase + 'table" Target="../tables/table1.xml"/>' +
    '</Relationships>'
$parts['xl/worksheets/sheet3.xml'] = $head + '<worksheet ' + $ns + '>' +
    '<dimension ref="B1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
    '<sheetData><row r="1"><c r="B1"><v>5</v></c></row></sheetData>' + $margins + '</worksheet>'
$parts['xl/worksheets/sheet4.xml'] = $head + '<worksheet ' + $ns + '>' +
    '<dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
    '<sheetData><row r="1"><c r="A1" t="s"><v>9</v></c></row></sheetData>' + $margins + '</worksheet>'
$parts['xl/worksheets/sheet5.xml'] = $head + '<worksheet ' + $ns + '>' +
    '<dimension ref="B2"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
    '<sheetData><row r="2"><c r="B2"><v>10</v></c></row></sheetData>' + $margins + '</worksheet>'
$parts['xl/worksheets/sheet6.xml'] = $head + '<worksheet ' + $ns + '>' +
    '<dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
    '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>' + $margins + '</worksheet>'
$parts['xl/tables/table1.xml'] = $head + '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B4" totalsRowShown="0">' +
    '<autoFilter ref="A1:B4"/>' +
    '<tableColumns count="2"><tableColumn id="1" name="Item"/><tableColumn id="2" name="Amount"/></tableColumns>' +
    '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>' +
    '</table>'
$parts['xl/sharedStrings.xml'] = $head + '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="10" uniqueCount="10">' +
    '<si><t>Revenue</t></si><si><t>Cost</t></si><si><t>Margin</t></si>' +
    '<si><r><t xml:space="preserve">Rate </t></r><r><rPr><b/><sz val="11"/><rFont val="Calibri"/><family val="2"/></rPr><t>applies</t></r></si>' +
    '<si><t>Item</t></si><si><t>Amount</t></si><si><t>pens</t></si><si><t>ink</t></si><si><t>paper</t></si>' +
    '<si><t>do not show</t></si>' +
    '</sst>'
$parts['xl/styles.xml'] = $head + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="3">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '<dxfs count="0"/>' +
    '</styleSheet>'
$parts['xl/externalLinks/externalLink1.xml'] = $head + '<externalLink ' + $ns + '>' +
    '<externalBook r:id="rId1">' +
    '<sheetNames><sheetName val="Sheet1"/></sheetNames>' +
    '<sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>99</v></cell></row></sheetData></sheetDataSet>' +
    '</externalBook></externalLink>'
$parts['xl/externalLinks/_rels/externalLink1.xml.rels'] = $head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="' + $relBase + 'externalLinkPath" Target="Rates.xlsx" TargetMode="External"/>' +
    '</Relationships>'
$parts['xl/metadata.xml'] = $head + '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">' +
    '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
    '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>' +
    '</metadata>'
$parts['docProps/core.xml'] = $head + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '<dc:creator>Frazaro fixture</dc:creator><cp:lastModifiedBy>Frazaro fixture</cp:lastModifiedBy>' +
    '<dcterms:created xsi:type="dcterms:W3CDTF">2026-10-03T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-10-03T00:00:00Z</dcterms:modified>' +
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
        foreach ($name in $parts.Keys) {
            $e = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
            $e.LastWriteTime = $stamp
            $bytes = $utf8.GetBytes([string]$parts[$name])
            $s = $e.Open(); $s.Write($bytes, 0, $bytes.Length); $s.Dispose()
        }
    } finally { $zip.Dispose() }
} finally { $fs.Dispose() }
Write-Output ("wrote {0} ({1} bytes, {2} parts, deflated)" -f $out, (Get-Item $out).Length, $parts.Count)
