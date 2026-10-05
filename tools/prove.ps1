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

2026-10-01, PORT.5 (the treaty's amendment of that date): a fourth kind,
compile - the .vla golden, less its GENERATED stamp line, with prelude.vla,
to the .vba golden. The runner writes the stamp-less program to a file in a
scratch directory and names it in the result, so what was compiled is a file
someone can read; the control's fake answers it from the .vba golden and the
mutant changes one byte, as for the translate kinds.

2026-10-02, PORT.6 (the treaty's amendment of that date): oracle 1's .vla
golden begins with the same writer's stamp line, which the runner drops
before comparing, as it does for compile. Oracle 3's count, n, is the number
of test-success and test-fail forms the loader runs after a phrasebook's own
generators have expanded: read from the <name>_expanded.vla export beside the
source when there is one (its source-hash stamp checked against the source,
so a stale export fails here rather than counting), and from the source's
own top-level forms otherwise. The control's fake counts the same way.

2026-10-02, LX.15 (the treaty's third amendment of that date): a file under
scripts/polyglotta/ that holds no <lingua>-vla rule and no proof form is a
library of macros a program includes (alien.vla), not a phrasebook. It is
inventoried as a library and never scored; the control's attempted count
leaves it out.

2026-10-03, PORT.7 (the treaty's amendment of that date): a fifth kind,
build - scripts/build/fixture.txt, with prelude.vla and english.vla, to
scripts/build/fixture_golden.xlsx, a workbook, compared byte for byte with
no normalization, since a zip is bytes. The implementation writes the file
at --out in the runner's scratch directory, which the result names; the
control's fake copies the golden there and the mutant changes one byte of
it. The golden is the core's own, opened by the owner in Excel (the
amendment says why), not one the VBA reference produced.

2026-10-03, PORT.8 (the treaty's amendment of that date, slice 8a): a sixth
kind, reflect - a fixture workbook to its expected relations, one form a
line in a fixed order, compared as text the way the translate kinds compare
it. Four rows: the reader's own fixture (scripts/reflect/fixture.xlsx, which
tools/build_reflect_fixture.ps1 writes), the two build goldens read back,
and the model the into golden was built into; each golden is the core's own
output, blessed by the owner reading the fixture in Excel by eye. The
control's fake answers each from the golden beside its fixture and the
mutant changes one character.

2026-10-04, PORT.8 (the treaty's amendment of that date, slice 8c): a
seventh kind, diff - two fixture workbooks to their difference, one form a
line in a fixed order, compared as text the way the reflect kind compares
it. Three rows: the first build golden against its Excel-saved copy, the
model against the into golden built into it, and the reader's fixture
against the changed copy build_reflect_fixture.ps1 -Changed writes; the
changed copy is also the reflect kind's sixth row. The control's fake
answers each from the golden of the pair, found by the new file's name,
and the mutant changes one character.
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
# A file's bytes, read with sharing that tolerates a workbook open in Excel:
# the owner's live pass has the build golden open while the runner scores,
# and Excel's lock refuses a plain ReadAllBytes (2026-10-03).
function Read-BytesShared([string]$path) {
    $fs = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $buf = New-Object byte[] $fs.Length
        $got = 0
        while ($got -lt $buf.Length) {
            $n = $fs.Read($buf, $got, $buf.Length - $got)
            if ($n -le 0) { break }
            $got += $n
        }
        if ($got -ne $buf.Length) { throw "read $got of $($buf.Length) bytes from $path" }
        return ,$buf
    } finally {
        $fs.Dispose()
    }
}
function Get-FormCount([string]$path, [string]$pattern) {
    return @(Select-String -Path $path -Pattern $pattern).Count
}
# The compile oracle's input: the .vla golden without its first line, the
# GENERATED stamp VlaWriteGoldens adds after transpiling (the treaty's
# amendment of 2026-10-01 says why). Written in a scratch directory the
# caller owns, never beside the repository's own files, and returned as a
# path so the result can name what was compiled. Bytes are copied, not
# re-encoded: the file is ASCII and its line endings are left as found.
function Write-CompileInput([string]$golden, [string]$dir) {
    $bytes = [System.IO.File]::ReadAllBytes($golden)
    $i = [Array]::IndexOf($bytes, [byte]10)
    if ($i -lt 0) { throw "no line break in $golden" }
    $rest = New-Object byte[] ($bytes.Length - $i - 1)
    [Array]::Copy($bytes, $i + 1, $rest, 0, $rest.Length)
    $out = Join-Path $dir 'instructions_golden.unstamped.vla'
    [System.IO.File]::WriteAllBytes($out, $rest)
    return $out
}
# Oracle 1's .vla golden less its first line, the same stamp, normalized:
# what translate-vla must reproduce (the amendment of 2026-10-02).
function Read-GoldenLessStamp([string]$path) {
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $i = [Array]::IndexOf($bytes, [byte]10)
    if ($i -lt 0) { throw "no line break in $path" }
    $rest = New-Object byte[] ($bytes.Length - $i - 1)
    [Array]::Copy($bytes, $i + 1, $rest, 0, $rest.Length)
    return Get-NormalizedText ([System.Text.Encoding]::UTF8.GetString($rest))
}
# SHA-256 over a file's non-whitespace bytes, upper-case hex: the stamp an
# expanded phrasebook export carries (EnglishSourceHash; check_rule_coverage.ps1
# reads it the same way).
function Get-NonWhitespaceSha256([string]$path) {
    $raw = [System.IO.File]::ReadAllBytes($path)
    $packed = New-Object byte[] $raw.Length
    $k = 0
    foreach ($b in $raw) {
        if ($b -ne 9 -and $b -ne 10 -and $b -ne 13 -and $b -ne 32) { $packed[$k] = $b; $k++ }
    }
    $out = New-Object byte[] $k
    [Array]::Copy($packed, $out, $k)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $digest = $sha.ComputeHash($out) } finally { $sha.Dispose() }
    return @{ Hex = (([BitConverter]::ToString($digest)) -replace '-', ''); Count = [long]$k }
}
# Oracle 3's n (the amendment of 2026-10-02): the proof forms the loader runs
# after the phrasebook's own generators expand. With an export beside the
# source (<name>_expanded.vla, written by Export Expanded Phrasebook) the
# forms are counted there, once its source-hash stamp matches the source;
# otherwise the source's own top-level forms are counted.
function Get-ProofCount([string]$root, [string]$relFile, [string]$pattern) {
    $file = Join-Path $root $relFile
    $dir = Split-Path -Parent $file
    $base = [System.IO.Path]::GetFileNameWithoutExtension($file)
    $expanded = Join-Path $dir ($base + '_expanded.vla')
    if (-not (Test-Path $expanded)) {
        return @{ Count = (Get-FormCount $file $pattern); From = ''; Stale = '' }
    }
    $stamp = (Get-Content -LiteralPath $expanded -TotalCount 2)[1]
    $m = [regex]::Match($stamp, 'source-hash:\s*sha256:([0-9A-Fa-f]{64}) over (\d+) non-whitespace bytes')
    $have = Get-NonWhitespaceSha256 $file
    if (-not $m.Success -or $m.Groups[1].Value.ToUpperInvariant() -ne $have.Hex -or [long]$m.Groups[2].Value -ne $have.Count) {
        return @{ Count = 0; From = ($base + '_expanded.vla'); Stale = "$($base)_expanded.vla is stale for $relFile - re-export it (Export Expanded Phrasebook) before scoring" }
    }
    return @{ Count = (Get-FormCount $expanded $pattern); From = ($base + '_expanded.vla'); Stale = '' }
}

# ---- the oracles, as the treaty lists them ----
# A phrasebook proof is a test-success or test-fail form at the top level of
# its line (2026-10-02); an engine proof any (test-<engine> ...) form.
$phrasebookForm = '^\(test-(success|fail)\b'
$engineForm     = '^\s*\(test-'
# A phrasebook holds at least one rule or one proof; a file with neither is
# a library (2026-10-02, LX.15).
$ruleForm       = '^\([a-z]+-vla(-override)?\b'

function Get-Oracles([string]$root) {
    $list = New-Object System.Collections.Generic.List[object]
    $list.Add(@{ Kind = 'translate-vla'; Label = 'instructions.txt -> instructions_golden.vla'
                 Input = 'scripts/instructions.txt'; Golden = 'scripts/instructions_golden.vla'
                 Prelude = 'scripts/prelude.vla'; Phrasebook = 'scripts/polyglotta/english.vla'
                 Stamped = $true })   # the golden's first line is the writer's stamp
    $list.Add(@{ Kind = 'translate-vba'; Label = 'instructions.txt -> instructions_golden.vba'
                 Input = 'scripts/instructions.txt'; Golden = 'scripts/instructions_golden.vba'
                 Prelude = 'scripts/prelude.vla'; Phrasebook = 'scripts/polyglotta/english.vla' })
    $list.Add(@{ Kind = 'compile'; Label = 'instructions_golden.vla -> instructions_golden.vba'
                 Input = 'scripts/instructions_golden.vla'; Golden = 'scripts/instructions_golden.vba'
                 Prelude = 'scripts/prelude.vla' })
    $list.Add(@{ Kind = 'build'; Label = 'build/fixture.txt -> build/fixture_golden.xlsx'
                 Input = 'scripts/build/fixture.txt'; Golden = 'scripts/build/fixture_golden.xlsx'
                 Prelude = 'scripts/prelude.vla'; Phrasebook = 'scripts/polyglotta/english.vla'; Into = '' })
    # 2026-10-03 (slice 7d): the same kind, built into a workbook the fixture
    # script made; the golden is the model with the program's sheets added.
    $list.Add(@{ Kind = 'build'; Label = 'build/into.txt into model.xlsx -> build/into_golden.xlsx'
                 Input = 'scripts/build/into.txt'; Golden = 'scripts/build/into_golden.xlsx'
                 Prelude = 'scripts/prelude.vla'; Phrasebook = 'scripts/polyglotta/english.vla'
                 Into = 'scripts/build/model.xlsx' })
    # 2026-10-03 (PORT.8, slice 8a): the reflect kind, a workbook to its
    # relations as text; the fixture's golden, the two build goldens read
    # back, and the model the into golden was built into.
    $list.Add(@{ Kind = 'reflect'; Label = 'reflect/fixture.xlsx -> reflect/fixture_relations.vla'
                 Input = 'scripts/reflect/fixture.xlsx'; Golden = 'scripts/reflect/fixture_relations.vla' })
    $list.Add(@{ Kind = 'reflect'; Label = 'build/fixture_golden.xlsx -> reflect/build_fixture_relations.vla'
                 Input = 'scripts/build/fixture_golden.xlsx'; Golden = 'scripts/reflect/build_fixture_relations.vla' })
    $list.Add(@{ Kind = 'reflect'; Label = 'build/into_golden.xlsx -> reflect/build_into_relations.vla'
                 Input = 'scripts/build/into_golden.xlsx'; Golden = 'scripts/reflect/build_into_relations.vla' })
    $list.Add(@{ Kind = 'reflect'; Label = 'build/model.xlsx -> reflect/model_relations.vla'
                 Input = 'scripts/build/model.xlsx'; Golden = 'scripts/reflect/model_relations.vla' })
    # 2026-10-04 (the owner's live pass for 8a): the first build golden as
    # Excel 365 saved it, a host's cached values in every formula cell.
    $list.Add(@{ Kind = 'reflect'; Label = 'reflect/saved.xlsx -> reflect/saved_relations.vla'
                 Input = 'scripts/reflect/saved.xlsx'; Golden = 'scripts/reflect/saved_relations.vla' })
    # 2026-10-04 (PORT.8, slice 8c): the fixture's changed copy, and the diff
    # kind, two workbooks to their difference as text.
    $list.Add(@{ Kind = 'reflect'; Label = 'reflect/changed.xlsx -> reflect/changed_relations.vla'
                 Input = 'scripts/reflect/changed.xlsx'; Golden = 'scripts/reflect/changed_relations.vla' })
    $list.Add(@{ Kind = 'diff'; Label = 'fixture_golden.xlsx vs saved.xlsx -> build_fixture_saved_diff.vla'
                 Old = 'scripts/build/fixture_golden.xlsx'; New = 'scripts/reflect/saved.xlsx'; Golden = 'scripts/reflect/build_fixture_saved_diff.vla' })
    $list.Add(@{ Kind = 'diff'; Label = 'model.xlsx vs into_golden.xlsx -> model_into_diff.vla'
                 Old = 'scripts/build/model.xlsx'; New = 'scripts/build/into_golden.xlsx'; Golden = 'scripts/reflect/model_into_diff.vla' })
    $list.Add(@{ Kind = 'diff'; Label = 'fixture.xlsx vs changed.xlsx -> fixture_changed_diff.vla'
                 Old = 'scripts/reflect/fixture.xlsx'; New = 'scripts/reflect/changed.xlsx'; Golden = 'scripts/reflect/fixture_changed_diff.vla' })
    $list.Add(@{ Kind = 'interpreter'; Label = 'interpreter_golden.txt (needs a workbook model: slice 6)'
                 Golden = 'scripts/interpreter_golden.txt' })
    $pb = Join-Path $root 'scripts/polyglotta'
    foreach ($f in Get-ChildItem -Path $pb -Filter '*.vla' | Sort-Object Name) {
        if ($f.Name -like '*_expanded*') { continue }   # an export, not a source
        $rel = 'scripts/polyglotta/' + $f.Name
        if ((Get-FormCount $f.FullName $ruleForm) -eq 0 -and (Get-FormCount $f.FullName $phrasebookForm) -eq 0) {
            # LX.15: a library of macros, not a phrasebook; inventoried, never scored.
            $list.Add(@{ Kind = 'library'; Label = ('polyglotta/' + $f.Name); File = $rel })
            continue
        }
        $list.Add(@{ Kind = 'prove'; Label = ('polyglotta/' + $f.Name); File = $rel; Pattern = $phrasebookForm })
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

function Measure-Oracles([string]$root, [string]$impl, [string]$scratch) {
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
                $want = if ($o.Stamped) { Read-GoldenLessStamp $golden } else { Read-NormalizedFile $golden }
                $got  = Get-NormalizedText $run.Stdout
                if ($run.ExitCode -ne 0) { $r.Status = 'FAIL'; $r.Detail = "exit $($run.ExitCode)" }
                elseif ($got -eq $want) { $r.Status = 'PASS'; $r.Detail = "$($want.Length) chars matched" }
                else {
                    $n = [Math]::Min($got.Length, $want.Length); $at = $n
                    for ($i = 0; $i -lt $n; $i++) { if ($got[$i] -ne $want[$i]) { $at = $i; break } }
                    $r.Status = 'FAIL'; $r.Detail = "differs at char $at of $($want.Length)"
                }
            }
            'compile' {
                $golden = Join-Path $root $o.Golden
                $program = Write-CompileInput (Join-Path $root $o.Input) $scratch
                $run = Invoke-Impl $impl @('compile', $program, '--prelude', (Join-Path $root $o.Prelude))
                if ($run.ExitCode -eq 3) { break }
                $want = Read-NormalizedFile $golden
                $got  = Get-NormalizedText $run.Stdout
                if ($run.ExitCode -ne 0) { $r.Status = 'FAIL'; $r.Detail = "exit $($run.ExitCode) on $program" }
                elseif ($got -eq $want) { $r.Status = 'PASS'; $r.Detail = "$($want.Length) chars matched from $program" }
                else {
                    $n = [Math]::Min($got.Length, $want.Length); $at = $n
                    for ($i = 0; $i -lt $n; $i++) { if ($got[$i] -ne $want[$i]) { $at = $i; break } }
                    $r.Status = 'FAIL'; $r.Detail = "differs at char $at of $($want.Length), compiled from $program"
                }
            }
            'build' {
                # A workbook is bytes: no normalization, the first differing
                # byte named. The file is written in the scratch directory and
                # left there, so a FAIL can name what was built.
                $golden = Join-Path $root $o.Golden
                $built = Join-Path $scratch ([System.IO.Path]::GetFileNameWithoutExtension($o.Golden) + '.built.xlsx')
                if (Test-Path $built) { Remove-Item $built -Force }
                $buildArgs = @('build', (Join-Path $root $o.Input),
                               '--prelude', (Join-Path $root $o.Prelude),
                               '--phrasebook', (Join-Path $root $o.Phrasebook),
                               '--out', $built)
                if ($o.Into -ne '') { $buildArgs += @('--into', (Join-Path $root $o.Into)) }
                $run = Invoke-Impl $impl $buildArgs
                if ($run.ExitCode -eq 3) { break }
                if ($run.ExitCode -ne 0) { $r.Status = 'FAIL'; $r.Detail = "exit $($run.ExitCode)"; break }
                if (-not (Test-Path $built)) { $r.Status = 'FAIL'; $r.Detail = "exit 0 but nothing written at $built"; break }
                $want = Read-BytesShared $golden
                $got = Read-BytesShared $built
                $n = [Math]::Min($got.Length, $want.Length); $at = $n
                for ($i = 0; $i -lt $n; $i++) { if ($got[$i] -ne $want[$i]) { $at = $i; break } }
                if ($at -eq $n -and $got.Length -eq $want.Length) {
                    # 2026-10-03 (slice 7c): what was built must also verify as its own build.
                    $rb = Invoke-Impl $impl @('rebuild', $built,
                                              '--prelude', (Join-Path $root $o.Prelude),
                                              '--phrasebook', (Join-Path $root $o.Phrasebook))
                    if ($rb.ExitCode -eq 0) { $r.Status = 'PASS'; $r.Detail = "$($want.Length) bytes matched; rebuild says yes" }
                    else { $r.Status = 'FAIL'; $r.Detail = "$($want.Length) bytes matched, but rebuild exited $($rb.ExitCode): $(($rb.Stdout -split "`n") | Select-Object -Last 1)" }
                }
                else { $r.Status = 'FAIL'; $r.Detail = "differs at byte $at of $($want.Length), built at $built" }
            }
            'reflect' {
                # 2026-10-03 (PORT.8, slice 8a): the relations as text, normalized
                # as the translate kinds are; the golden carries no stamp.
                $golden = Join-Path $root $o.Golden
                $run = Invoke-Impl $impl @('reflect', (Join-Path $root $o.Input))
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
            'diff' {
                # 2026-10-04 (PORT.8, slice 8c): two workbooks to their difference
                # as text, normalized as the reflect kind is; the door exits 0
                # whether or not the files differ, and the golden carries no stamp.
                $golden = Join-Path $root $o.Golden
                $run = Invoke-Impl $impl @('diff', (Join-Path $root $o.Old), (Join-Path $root $o.New))
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
                $run = Invoke-Impl $impl @('prove', $file)
                if ($run.ExitCode -eq 3) { break }
                $pc = Get-ProofCount $root $o.File $o.Pattern
                if ($pc.Stale -ne '') { $r.Status = 'FAIL'; $r.Detail = $pc.Stale; break }
                $n = $pc.Count
                $last = (Get-NormalizedText $run.Stdout) -split "`n" | Select-Object -Last 1
                if ($run.ExitCode -eq 0 -and $last -match "^PASS (\d+)/(\d+)$" -and [int]$Matches[1] -eq $n -and [int]$Matches[2] -eq $n) {
                    $r.Status = 'PASS'; $r.Detail = "$n proof(s)"
                } else {
                    $r.Status = 'FAIL'; $r.Detail = "exit $($run.ExitCode), last line '$last', $n proof(s) expected"
                }
            }
            'library' { $r.Status = 'library'; $r.Detail = 'a library of macros, not a phrasebook: no rule, no proof; inventoried, not scored' }
            default { $r.Detail = 'inventoried, not scored' }
        }
        $results.Add($r)
    }
    return $results
}

function Write-Results($results, [switch]$all) {
    $fails = 0; $passes = 0; $skipped = 0; $libraries = 0
    foreach ($r in $results) {
        switch ($r.Status) { 'PASS' { $passes++ } 'FAIL' { $fails++ } 'library' { $libraries++ } default { $skipped++ } }
        if ($all -or $r.Status -eq 'FAIL') {
            Write-Host ("  {0,-13} {1,-60} {2}" -f $r.Status, $r.Label, $r.Detail)
        }
    }
    $line = "prove: {0} passed, {1} failed, {2} not attempted" -f $passes, $fails, $skipped
    if ($libraries -gt 0) { $line += ", {0} library" -f $libraries }
    Write-Host $line
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
    'translate-vla' {
        # The golden less its first line, the writer's stamp (2026-10-02).
        `$b = [System.IO.File]::ReadAllBytes((Join-Path `$root 'scripts/instructions_golden.vla'))
        `$i = [Array]::IndexOf(`$b, [byte]10)
        `$g = [System.Text.Encoding]::UTF8.GetString(`$b, `$i + 1, `$b.Length - `$i - 1)
    }
    'translate-vba' { `$g = [System.IO.File]::ReadAllText((Join-Path `$root 'scripts/instructions_golden.vba')) }
    'compile'       { `$g = [System.IO.File]::ReadAllText((Join-Path `$root 'scripts/instructions_golden.vba')) }
    'prove' {
        `$pattern = if (`$a[1] -like '*polyglotta*') { '$pb' } else { '$en' }
        # n as the treaty counts it: the expanded export beside the source, where there is one.
        `$exp = Join-Path (Split-Path -Parent `$a[1]) ([System.IO.Path]::GetFileNameWithoutExtension(`$a[1]) + '_expanded.vla')
        `$counted = if (Test-Path `$exp) { `$exp } else { `$a[1] }
        `$n = @(Select-String -Path `$counted -Pattern `$pattern).Count
        if (`$mutant) { Write-Output "FAIL 1/`$n"; exit 1 } else { Write-Output "PASS `$n/`$n"; exit 0 }
    }
    'build' {
        # The golden copied to --out (the second golden when --into is given,
        # 2026-10-03 slice 7d); the mutant with one byte changed.
        `$dst = `$a[[Array]::IndexOf(`$a, '--out') + 1]
        `$which = if (`$a -contains '--into') { 'scripts/build/into_golden.xlsx' } else { 'scripts/build/fixture_golden.xlsx' }
        `$fs = New-Object System.IO.FileStream((Join-Path `$root `$which), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        `$b = New-Object byte[] `$fs.Length
        [void]`$fs.Read(`$b, 0, `$b.Length)
        `$fs.Dispose()
        if (`$mutant) { `$b[100] = `$b[100] -bxor 1 }
        [System.IO.File]::WriteAllBytes(`$dst, `$b)
        exit 0
    }
    'rebuild' {
        # Yes when the file is one of the goldens byte for byte, no otherwise (2026-10-03, slice 7c).
        `$f = [System.IO.File]::ReadAllBytes(`$a[1])
        `$same = `$false
        foreach (`$which in @('scripts/build/fixture_golden.xlsx', 'scripts/build/into_golden.xlsx')) {
            `$p = Join-Path `$root `$which
            if (-not (Test-Path `$p)) { continue }
            `$fs = New-Object System.IO.FileStream(`$p, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            `$g = New-Object byte[] `$fs.Length
            [void]`$fs.Read(`$g, 0, `$g.Length)
            `$fs.Dispose()
            if (`$f.Length -ne `$g.Length) { continue }
            `$eq = `$true
            for (`$i = 0; `$i -lt `$f.Length; `$i++) { if (`$f[`$i] -ne `$g[`$i]) { `$eq = `$false; break } }
            if (`$eq) { `$same = `$true; break }
        }
        if (`$same) { Write-Output 'This workbook was built from these sentences by the fake: yes.'; exit 0 }
        Write-Output 'This workbook was built from these sentences by the fake: no.'; exit 1
    }
    'reflect' {
        # The golden beside the fixture, by the fixture's name (2026-10-03, PORT.8 slice 8a).
        `$base = [System.IO.Path]::GetFileNameWithoutExtension(`$a[1])
        `$which = switch (`$base) {
            'fixture'        { 'scripts/reflect/fixture_relations.vla' }
            'fixture_golden' { 'scripts/reflect/build_fixture_relations.vla' }
            'into_golden'    { 'scripts/reflect/build_into_relations.vla' }
            'model'          { 'scripts/reflect/model_relations.vla' }
            'saved'          { 'scripts/reflect/saved_relations.vla' }
            'changed'        { 'scripts/reflect/changed_relations.vla' }
            default          { exit 3 }
        }
        `$g = [System.IO.File]::ReadAllText((Join-Path `$root `$which))
    }
    'diff' {
        # The golden of the pair, by the new file's name (2026-10-04, PORT.8 slice 8c).
        `$base = [System.IO.Path]::GetFileNameWithoutExtension(`$a[2])
        `$which = switch (`$base) {
            'saved'       { 'scripts/reflect/build_fixture_saved_diff.vla' }
            'into_golden' { 'scripts/reflect/model_into_diff.vla' }
            'changed'     { 'scripts/reflect/fixture_changed_diff.vla' }
            default       { exit 3 }
        }
        `$g = [System.IO.File]::ReadAllText((Join-Path `$root `$which))
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
        $f1 = Write-Results (Measure-Oracles $root $fake $tmp) -all:$ShowAll
        Write-Host 'prove -Control: the mutant (one byte changed, one proof failed) must fail'
        $res2 = Measure-Oracles $root $mut $tmp
        $f2 = Write-Results $res2 -all:$ShowAll
        $attempted = @($res2 | Where-Object { $_.Status -in 'PASS', 'FAIL' }).Count
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
    # The compile oracle's stamp-less input is written here and left in
    # place, so a FAIL line can name a file that still exists afterwards.
    $scratch = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_prove_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $scratch | Out-Null
    $fails = Write-Results (Measure-Oracles $root $Impl $scratch) -all:$ShowAll
    if ($fails -gt 0) { exit 1 }
    exit 0
}

# ---- inventory ----
$missing = 0
Write-Host 'prove: the oracles (conformance/README.md)'
foreach ($o in Get-Oracles $root) {
    switch ($o.Kind) {
        'prove' {
            $pc = Get-ProofCount $root $o.File $o.Pattern
            if ($pc.Stale -ne '') {
                $missing++
                Write-Host ("  {0,-13} {1,-44} STALE: {2}" -f $o.Kind, $o.Label, $pc.Stale)
            } else {
                $note = if ($pc.From -ne '') { " (counted in $($pc.From))" } else { '' }
                Write-Host ("  {0,-13} {1,-44} {2,6} proof form(s){3}" -f $o.Kind, $o.Label, $pc.Count, $note)
            }
        }
        'library' {
            Write-Host ("  {0,-13} {1,-44} a library of macros, not a phrasebook; not scored" -f $o.Kind, $o.Label)
        }
        default {
            $p = Join-Path $root $o.Golden
            if (Test-Path $p) {
                Write-Host ("  {0,-13} {1,-44} {2,6} bytes" -f $o.Kind, $o.Label, (Get-Item $p).Length)
            } else { $missing++; Write-Host ("  {0,-13} {1,-44} MISSING" -f $o.Kind, $o.Label) }
        }
    }
}
if ($missing -gt 0) { Write-Host "FAIL: $missing oracle file(s) missing or stale"; exit 1 }
Write-Host 'OK: every oracle file is present; score an implementation with -Impl, or prove this runner with -Control'
exit 0
