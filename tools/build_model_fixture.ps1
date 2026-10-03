<#
build_model_fixture.ps1 - the small workbook frazaro build --into adds to,
written deterministically: scripts/build/model.xlsx (PORT.7, slice 7d).

WHAT IT BUILDS: a workbook of the shape Excel saves and the writer must
leave untouched - deflated parts (.NET's DeflateStream, so the core's own
inflate decodes a real stream), a shared-string table, a styles part with
two cell formats, a calculation chain, a defined name, two sheets:

  Model   A1 Revenue  B1 1200
          A2 Cost     B2 800
          A3 Margin   B3 =B1-B2  (cached value 400)
  Notes   A1 Q3 close

and the defined name Rate = 0.2. The calcPr carries a calcId and no
fullCalcOnLoad, so the merge must add the attribute; Model's sheetView is
the selected tab, so the merge must not select a second one.

WHY A SCRIPT: a workbook made by hand could not be rebuilt byte for byte,
and this one is a fixture of the treaty's oracle 7 (conformance/README.md):
scripts/build/into.txt built into it is scripts/build/into_golden.xlsx,
byte for byte. The zip stamps are fixed and the parts go in a fixed order,
as tools/build_examples.ps1 does for the onboarding samples; .NET's deflate
output is what it is on the machine that ran this, and the committed bytes
are the fixture, regenerated rarely and on purpose.

NO OFFICE AUTOMATION: raw OOXML, no Excel. House style (tools/*.ps1):
PowerShell 5.1, host-free, no network.

Usage:  powershell -File tools\build_model_fixture.ps1
#>
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
$root = Split-Path -Parent $PSScriptRoot
$out = Join-Path $root 'scripts\build\model.xlsx'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$stamp = New-Object DateTimeOffset(2026, 9, 18, 0, 0, 0, [TimeSpan]::Zero)
$head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`r`n"

$parts = [ordered]@{}
$parts['[Content_Types].xml'] = $head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
    '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
    '</Types>'
$parts['_rels/.rels'] = $head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
    '</Relationships>'
$parts['xl/workbook.xml'] = $head + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<workbookPr defaultThemeVersion="166925"/>' +
    '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12225" activeTab="0"/></bookViews>' +
    '<sheets><sheet name="Model" sheetId="1" r:id="rId1"/><sheet name="Notes" sheetId="2" r:id="rId2"/></sheets>' +
    '<definedNames><definedName name="Rate">0.2</definedName></definedNames>' +
    '<calcPr calcId="191029"/>' +
    '</workbook>'
$parts['xl/_rels/workbook.xml.rels'] = $head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
    '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>' +
    '</Relationships>'
$parts['xl/worksheets/sheet1.xml'] = $head + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<dimension ref="A1:B3"/>' +
    '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B3" sqref="B3"/></sheetView></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="15"/>' +
    '<cols><col min="1" max="1" width="12.7109375" customWidth="1"/></cols>' +
    '<sheetData>' +
    '<row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1"><v>1200</v></c></row>' +
    '<row r="2"><c r="A2" s="1" t="s"><v>1</v></c><c r="B2"><v>800</v></c></row>' +
    '<row r="3"><c r="A3" s="1" t="s"><v>2</v></c><c r="B3"><f>B1-B2</f><v>400</v></c></row>' +
    '</sheetData>' +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    '</worksheet>'
$parts['xl/worksheets/sheet2.xml'] = $head + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<dimension ref="A1"/>' +
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="15"/>' +
    '<sheetData><row r="1"><c r="A1" t="s"><v>3</v></c></row></sheetData>' +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    '</worksheet>'
$parts['xl/sharedStrings.xml'] = $head + '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">' +
    '<si><t>Revenue</t></si><si><t>Cost</t></si><si><t>Margin</t></si><si><t>Q3 close</t></si></sst>'
$parts['xl/styles.xml'] = $head + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '<dxfs count="0"/>' +
    '</styleSheet>'
$parts['xl/calcChain.xml'] = $head + '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B3" i="1"/></calcChain>'
$parts['docProps/core.xml'] = $head + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '<dc:creator>Frazaro fixture</dc:creator><cp:lastModifiedBy>Frazaro fixture</cp:lastModifiedBy>' +
    '<dcterms:created xsi:type="dcterms:W3CDTF">2026-09-18T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-18T00:00:00Z</dcterms:modified>' +
    '</cp:coreProperties>'
$parts['docProps/app.xml'] = $head + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    '<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>' +
    '<HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant><vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>' +
    '<TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Model</vt:lpstr><vt:lpstr>Notes</vt:lpstr><vt:lpstr>Rate</vt:lpstr></vt:vector></TitlesOfParts>' +
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
