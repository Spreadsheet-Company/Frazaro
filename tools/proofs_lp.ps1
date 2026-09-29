<#
proofs_lp.ps1 - export the proof corpus's DATALOG proofs as clingo programs
(METAPROOF.2).

DEV-ONLY ORACLE, exactly as tools/optimize0_lp.ps1 and optimize01_lp.ps1
(settled 2026-09-19, the owner's call): this script only WRITES .lp text
files. It never runs clingo, nothing in Frazaro calls it, no clingo binary
is kept in this repository, and SD-13 is untouched: the owner runs clingo by
hand from PowerShell.

WHY. A proof in scripts/proofs/datalog.vla states the answer a DATALOG
program must give, and TestDatalogProofs checks that DATALOG gives it. But
the answers were written down by the same hands that wrote the engine, and
REBUILD.md's Whitworth caveat applies: two readings by one author converge
on consistency, not on truth. clingo, the Potassco answer-set solver, is a
reading of another lineage. When it derives the rows a proof expects, the
expectation is true of the program under the standard semantics, not only
consistent with DATALOG.

WHAT IT WRITES, into tools\clingo, where the files are kept and tracked so a
changed proof shows in the diff:
  - proof-<file>-<name>.lp for every answer proof: the program, the rows
    the proof expects, and a judge that shows vla_verdict(agrees) - or
    vla_verdict(differs), beside the rows that make the difference.
  - proof-<file>-control-*.lp, three CONTROLS: one proof with its
    expectation made wrong on purpose (a row left out, a row added, no row
    at all), each of which must show vla_verdict(caught).
A proof the translation cannot state faithfully is not exported, and the
run says which and why; tools\proofs_lib.ps1 has the translation's rules.
A file no proof makes any more is removed. tools\check_proofs.ps1 (rule 9)
fails whenever these files fall behind the corpus.

THEN, BY HAND - every file, from the repository root, as one line:

  Get-ChildItem tools\clingo\proof-*.lp | ForEach-Object { $v = "$((& clingo $_.FullName) -match 'vla_verdict' -replace '.*vla_verdict\((\w+)\).*', '$1')"; if (-not $v) { $v = 'NO-ANSWER' }; '{0,-9} {1}' -f $v, $_.BaseName }

Every line must say agrees - or caught, for a control. Anything else
(differs, missed, NO-ANSWER) is a finding: a wrong expectation, a wrong
engine, or a wrong translation, and which one is the question to answer.

  powershell -NoProfile -ExecutionPolicy Bypass -File tools\proofs_lp.ps1
  ... -List           say what would be written; write nothing
  ... -OutDir <dir>   default tools\clingo
  ... -ProofsDir <d>  default scripts\proofs
#>
param(
    [string]$ProofsDir = '',
    [string]$OutDir = '',
    [switch]$List
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
if ($ProofsDir -eq '') { $ProofsDir = Join-Path $root 'scripts\proofs' }
if ($OutDir -eq '') { $OutDir = Join-Path $PSScriptRoot 'clingo' }
. (Join-Path $PSScriptRoot 'proofs_lib.ps1')

$runLine = @'
Get-ChildItem tools\clingo\proof-*.lp | ForEach-Object { $v = "$((& clingo $_.FullName) -match 'vla_verdict' -replace '.*vla_verdict\((\w+)\).*', '$1')"; if (-not $v) { $v = 'NO-ANSWER' }; '{0,-9} {1}' -f $v, $_.BaseName }
'@

$proofFiles = @(Get-ChildItem -LiteralPath $ProofsDir -Filter '*.vla' -File | Where-Object { $_.Extension -eq '.vla' } | Sort-Object Name)
if ($proofFiles.Count -eq 0) { throw "no proof file in $ProofsDir" }

$all = New-Object System.Collections.Generic.List[object]
foreach ($pf in $proofFiles) {
    $label = 'scripts/proofs/' + $pf.Name
    $toks = Get-ProofTokens ([IO.File]::ReadAllText($pf.FullName)) $label
    $forms = ConvertTo-ProofForms $toks $label
    $ex = ConvertTo-ClingoExport $forms ([IO.Path]::GetFileNameWithoutExtension($pf.Name)) $label
    Write-Output ("{0}: {1} proof(s) exported and {2} control(s); {3} refusal(s) not exported - a refusal is DATALOG's own policy, and no other lineage raises its message ids" -f $label, $ex.Exported, ($ex.Files.Count - $ex.Exported), $ex.Refusals)
    foreach ($s in $ex.Skipped) { Write-Output ("    not exported: ""{0}"" - {1}" -f $s.Proof, $s.Reason) }
    foreach ($fl in $ex.Files) { $all.Add($fl) }
}

if ($List) {
    foreach ($fl in $all) { Write-Output "    $($fl.Name)" }
    exit 0
}

if (-not (Test-Path -LiteralPath $OutDir)) { throw "no folder $OutDir" }
$made = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$changed = 0
foreach ($fl in $all) {
    [void]$made.Add($fl.Name)
    $lpPath = Join-Path $OutDir $fl.Name
    $old = $null
    if (Test-Path -LiteralPath $lpPath) { $old = [IO.File]::ReadAllText($lpPath) -replace "`r`n", "`n" }
    if ($old -cne $fl.Text) {
        # LF and ASCII, as optimize01_lp.ps1 writes: one text on every machine.
        [IO.File]::WriteAllText($lpPath, $fl.Text, (New-Object System.Text.ASCIIEncoding))
        Write-Output "    wrote $($fl.Name)"
        $changed++
    }
}
foreach ($existing in @(Get-ChildItem -LiteralPath $OutDir -Filter 'proof-*.lp' -File | Where-Object { $_.Extension -eq '.lp' })) {
    if (-not $made.Contains($existing.Name)) {
        Remove-Item -LiteralPath $existing.FullName
        Write-Output "    removed $($existing.Name) - no proof makes it now"
        $changed++
    }
}
Write-Output ("{0} file(s) in {1}; {2} written or removed." -f $all.Count, $OutDir, $changed)
Write-Output ''
Write-Output 'Now clingo, by hand - every file, from the repository root, as one line:'
Write-Output ''
Write-Output ('  ' + $runLine.Trim())
Write-Output ''
Write-Output 'Every line must say agrees - or caught, for a control.'
