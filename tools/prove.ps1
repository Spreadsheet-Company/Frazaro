<#
prove.ps1 - the conformance harness: PORT.4's runner, slice 0 of
docs/HORIZON.md section 12. The treaty it enforces is conformance/README.md.

WHY: the VBA in src/ is the reference implementation (SD-18) and cannot run
outside Excel, so it testifies through the goldens it produced. Any other
implementation of the language is held to those goldens, the phrasebook proofs
and the engine proofs by this one runner, so that "the port matches" is a
number someone can quote and not a feeling.

THREE MODES:
  (default)   inventory: list every oracle with its count or size.
  -Impl <p>   score the implementation at <p> (an .exe or a .ps1; see the
              contract in conformance/README.md) against every oracle.
  -Control    prove the runner itself: a fake implementation that answers
              from the goldens must pass, and a mutant with one byte changed
              and one proof failed must fail. run_checks.ps1 -WithExtras runs
              this mode.

NOT NAMED check_*.ps1, deliberately: it takes arguments and has modes, so it
belongs with the other verifiers run_checks.ps1 lists by hand, not with the
plain pass/fail scans. House style otherwise: PowerShell 5.1, host-free, no
Excel, no COM, no network.

Usage:
  powershell -File tools\prove.ps1
  powershell -File tools\prove.ps1 -Impl C:\path\to\frazaro.exe
  powershell -File tools\prove.ps1 -Impl C:\path\to\wrapper.ps1 -ShowAll
  powershell -File tools\prove.ps1 -Control
Exit code: 0 when every attempted oracle passed (or, in inventory mode, when
every oracle file exists); 1 otherwise.
#>
param(
    [string]$Impl = '',
    [switch]$Control,
    [switch]$ShowAll,
    # A repository root other than this script's own, so the runner can be
    # pointed at a scratch tree.
    [string]$Root = ''
)

$ErrorActionPreference = 'Stop'
$root = if ($Root -ne '') { $Root } else { Split-Path -Parent $PSScriptRoot }

# ---- the treaty's normalization: LF everywhere, no trailing blank lines ----
function Get-NormalizedText([string]$text) {
    if ($null -eq $text) { return '' }
    $t = $text -replace "`r`n", "`n"
    $t = $t -replace "`r", "`n"
    return $t.TrimEnd("`n")
}
function Read-NormalizedFile([string]$path) {
    return Get-NormalizedText ([System.IO.File]::ReadAllText($path))
}
function Get-FormCount([string]$path, [string]$pattern) {
    return @(Select-String -Path $path -Pattern $pattern).Count
}

# ---- the oracles, as the treaty lists them ----
$phrasebookForm = '^\s*\(test\b'
$engineForm     = '^\s*\(test-'

function Get-Oracles([string]$root) {
    $list = New-Object System.Collections.Generic.List[object]
    $list.Add(@{ Kind = 'translate-vla'; Label = 'instructions.txt -> instructions_golden.vla'
                 Input = 'scripts/instructions.txt'; Golden = 'scripts/instructions_golden.vla'
                 Prelude = 'scripts/prelude.vla'; Phrasebook = 'scripts/polyglotta/english.vla' })
    $list.Add(@{ Kind = 'translate-vba'; Label = 'instructions.txt -> instructions_golden.vba'
                 Input = 'scripts/instructions.txt'; Golden = 'scripts/instructions_golden.vba'
                 Prelude = 'scripts/prelude.vla'; Phrasebook = 'scripts/polyglotta/english.vla' })
    $list.Add(@{ Kind = 'interpreter'; Label = 'interpreter_golden.txt (needs a workbook model: slice 6)'
                 Golden = 'scripts/interpreter_golden.txt' })
    $pb = Join-Path $root 'scripts/polyglotta'
    foreach ($f in Get-ChildItem -Path $pb -Filter '*.vla' | Sort-Object Name) {
        if ($f.Name -like '*_expanded*') { continue }   # an export, not a source
        $list.Add(@{ Kind = 'prove'; Label = ('polyglotta/' + $f.Name); File = ('scripts/polyglotta/' + $f.Name)
                     Pattern = $phrasebookForm })
    }
    $pr = Join-Path $root 'scripts/proofs'
    foreach ($f in Get-ChildItem -Path $pr -Filter '*.vla' | Sort-Object Name) {
        $list.Add(@{ Kind = 'prove'; Label = ('proofs/' + $f.Name); File = ('scripts/proofs/' + $f.Name)
                     Pattern = $engineForm })
    }
    return $list
}

# ---- running an implementation ----
function Invoke-Impl([string]$impl, [string[]]$cmdArgs) {
    $global:LASTEXITCODE = 0
    $lines = & $impl @cmdArgs
    $code = $LASTEXITCODE
    if ($null -eq $code) { $code = 0 }
    $text = if ($null -eq $lines) { '' } else { (@($lines) | ForEach-Object { [string]$_ }) -join "`n" }
    return @{ Stdout = $text; ExitCode = [int]$code }
}

function Measure-Oracles([string]$root, [string]$impl) {
    $results = New-Object System.Collections.Generic.List[object]
    foreach ($o in Get-Oracles $root) {
        $r = @{ Label = $o.Label; Kind = $o.Kind; Status = 'not attempted'; Detail = '' }
        switch ($o.Kind) {
            { $_ -in 'translate-vla', 'translate-vba' } {
                $golden = Join-Path $root $o.Golden
                $cmdArgs = @($o.Kind, (Join-Path $root $o.Input),
                             '--prelude', (Join-Path $root $o.Prelude),
                             '--phrasebook', (Join-Path $root $o.Phrasebook))
                $run = Invoke-Impl $impl $cmdArgs
                if ($run.ExitCode -eq 3) { break }
                $want = Read-NormalizedFile $golden
                $got  = Get-NormalizedText $run.Stdout
                if ($run.ExitCode -ne 0) { $r.Status = 'FAIL'; $r.Detail = "exit $($run.ExitCode)" }
                elseif ($got -eq $want) { $r.Status = 'PASS'; $r.Detail = "$($want.Length) chars matched" }
                else {
                    $n = [Math]::Min($got.Length, $want.Length); $at = $n
                    for ($i = 0; $i -lt $n; $i++) { if ($got[$i] -ne $want[$i]) { $at = $i; break } }
                    $r.Status = 'FAIL'; $r.Detail = "differs at char $at of $($want.Length)"
                }
            }
            'prove' {
                $file = Join-Path $root $o.File
                $n = Get-FormCount $file $o.Pattern
                $run = Invoke-Impl $impl @('prove', $file)
                if ($run.ExitCode -eq 3) { break }
                $last = (Get-NormalizedText $run.Stdout) -split "`n" | Select-Object -Last 1
                if ($run.ExitCode -eq 0 -and $last -match "^PASS (\d+)/(\d+)$" -and [int]$Matches[1] -eq $n -and [int]$Matches[2] -eq $n) {
                    $r.Status = 'PASS'; $r.Detail = "$n proof(s)"
                } else {
                    $r.Status = 'FAIL'; $r.Detail = "exit $($run.ExitCode), last line '$last', $n proof(s) expected"
                }
            }
            default { $r.Detail = 'inventoried, not scored' }
        }
        $results.Add($r)
    }
    return $results
}

function Write-Results($results, [switch]$all) {
    $fails = 0; $passes = 0; $skipped = 0
    foreach ($r in $results) {
        switch ($r.Status) { 'PASS' { $passes++ } 'FAIL' { $fails++ } default { $skipped++ } }
        if ($all -or $r.Status -eq 'FAIL') {
            Write-Host ("  {0,-13} {1,-60} {2}" -f $r.Status, $r.Label, $r.Detail)
        }
    }
    Write-Host ("prove: {0} passed, {1} failed, {2} not attempted" -f $passes, $fails, $skipped)
    return $fails
}

# ---- the control: a fake implementation and its mutant ----
function Write-FakeImpl([string]$path, [string]$root, [bool]$mutant) {
    $pb = $phrasebookForm; $en = $engineForm
    $body = @"
param([Parameter(ValueFromRemainingArguments=`$true)][string[]]`$a)
`$kind = `$a[0]
`$root = '$($root -replace "'", "''")'
`$mutant = `$$($mutant.ToString().ToLower())
switch (`$kind) {
    'translate-vla' { `$g = [System.IO.File]::ReadAllText((Join-Path `$root 'scripts/instructions_golden.vla')) }
    'translate-vba' { `$g = [System.IO.File]::ReadAllText((Join-Path `$root 'scripts/instructions_golden.vba')) }
    'prove' {
        `$pattern = if (`$a[1] -like '*polyglotta*') { '$pb' } else { '$en' }
        `$n = @(Select-String -Path `$a[1] -Pattern `$pattern).Count
        if (`$mutant) { Write-Output "FAIL 1/`$n"; exit 1 } else { Write-Output "PASS `$n/`$n"; exit 0 }
    }
    default { exit 3 }
}
if (`$mutant -and `$g.Length -gt 200) { `$g = `$g.Substring(0, 100) + 'X' + `$g.Substring(101) }
Write-Output `$g
exit 0
"@
    [System.IO.File]::WriteAllText($path, $body)
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_prove_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $fake = Join-Path $tmp 'fake.ps1'; $mut = Join-Path $tmp 'mutant.ps1'
        Write-FakeImpl $fake $root $false
        Write-FakeImpl $mut  $root $true
        Write-Host 'prove -Control: the fake implementation (answers from the goldens) must pass'
        $f1 = Write-Results (Measure-Oracles $root $fake) -all:$ShowAll
        Write-Host 'prove -Control: the mutant (one byte changed, one proof failed) must fail'
        $res2 = Measure-Oracles $root $mut
        $f2 = Write-Results $res2 -all:$ShowAll
        $attempted = @($res2 | Where-Object { $_.Status -ne 'not attempted' }).Count
        if ($f1 -eq 0 -and $f2 -eq $attempted -and $attempted -gt 0) {
            Write-Host "OK: control passed every attempted oracle; mutant failed all $attempted of them"
            exit 0
        }
        Write-Host "FAIL: control failures $f1 (want 0); mutant failures $f2 (want $attempted)"
        exit 1
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}

if ($Impl -ne '') {
    if (-not (Test-Path $Impl)) { Write-Host "FAIL: no implementation at $Impl"; exit 1 }
    $fails = Write-Results (Measure-Oracles $root $Impl) -all:$ShowAll
    if ($fails -gt 0) { exit 1 }
    exit 0
}

# ---- inventory ----
$missing = 0
Write-Host 'prove: the oracles (conformance/README.md)'
foreach ($o in Get-Oracles $root) {
    switch ($o.Kind) {
        'prove' {
            $p = Join-Path $root $o.File
            Write-Host ("  {0,-13} {1,-44} {2,6} proof form(s)" -f $o.Kind, $o.Label, (Get-FormCount $p $o.Pattern))
        }
        default {
            $p = Join-Path $root $o.Golden
            if (Test-Path $p) {
                Write-Host ("  {0,-13} {1,-44} {2,6} bytes" -f $o.Kind, $o.Label, (Get-Item $p).Length)
            } else { $missing++; Write-Host ("  {0,-13} {1,-44} MISSING" -f $o.Kind, $o.Label) }
        }
    }
}
if ($missing -gt 0) { Write-Host "FAIL: $missing oracle file(s) missing"; exit 1 }
Write-Host 'OK: every oracle file is present; score an implementation with -Impl, or prove this runner with -Control'
exit 0
