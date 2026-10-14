<#
check_crate_data.ps1 - the files a crate carries as copies are their sources,
byte for byte.

WHY: `cargo install frazaro` builds the door from its crate directory and
nothing beside it, so the prelude and the English phrasebook the door carries
(cli/data/, embedded at build time, used when --prelude or --phrasebook is
not given) are COPIES of scripts/prelude.vla and
scripts/polyglotta/english.vla, the corpus files the add-in, the web page and
the treaty read. Two copies of one file with nothing mechanical to hold them
together is the drift check_data_exports.ps1 was written for: a prelude edit
lands in scripts/, every test passes, and the published door translates with
yesterday's prelude. This check compares each pair byte for byte and names
the one copy command that mends a drift. The copies are never edited by hand.

THE SECOND CRATE (VLA.1, 2026-10-10): vla-lang's own tests travel with its
package, so the nine files they read - the prelude, the corpus golden, the
refers fixture and golden, and five view goldens - are copies under
vla-lang/testdata/, mirroring scripts/. Until then the published crate's
suite did not compile outside this repository. A golden regenerated in
scripts/ (by the reference, or by the door for the view goldens) now arrives
here as a failure with a one-line fix, instead of changing the language's
test inputs silently.

THE LICENCE FILES (2026-10-14, at 0.9.3): the root's LICENSE (Apache-2.0's
text) and NOTICE sit beside the crates, not in them, so no published package
before 0.9.3 held either, and whoever redistributes a crate needs both:
Apache-2.0 section 4 asks them to give the License and the NOTICE's
attributions on. Each crate's directory carries the two as copies, and the
door's also the Mozilla Public License's text as LICENSE-MPL, for the
phrasebook it carries, whose header names MPL-2.0 and nothing more.

ALSO HELD, on a built door (-Impl, or target/debug/frazaro beside the tree,
as the other -Impl checks find it): the door with no flags translates the
corpus program exactly as it does with the corpus files named, so that the
embedding itself, not only the bytes on disk, is what the flags would give.
SKIPPED where no door was built, exit 0, since CI always has one.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; a hardcoded, reviewable baseline (the pair count). Exit 0 clean,
exit 1 with every drift named.

Usage:  powershell -File tools\check_crate_data.ps1
        powershell -File tools\check_crate_data.ps1 -Impl target\debug\frazaro.exe
        powershell -File tools\check_crate_data.ps1 -Root <dir>   # testability
#>
param(
    [string]$Impl = '',
    [string]$Root = ''
)

$ErrorActionPreference = 'Stop'
$root = if ($Root -ne '') { (Resolve-Path $Root).Path } else { Split-Path -Parent $PSScriptRoot }
$root = $root.TrimEnd('\', '/')

# ---- the baseline: the pairs, copy then source; raise the count with a pair ----
$pairs = @(
    @{ Copy = 'cli/data/prelude.vla'; Source = 'scripts/prelude.vla' },
    @{ Copy = 'cli/data/english.vla'; Source = 'scripts/polyglotta/english.vla' },
    # VLA.1: the language's test inputs, travelling with its package
    @{ Copy = 'vla-lang/testdata/prelude.vla'; Source = 'scripts/prelude.vla' },
    @{ Copy = 'vla-lang/testdata/instructions_golden.vla'; Source = 'scripts/instructions_golden.vla' },
    @{ Copy = 'vla-lang/testdata/refers.txt'; Source = 'scripts/refers.txt' },
    @{ Copy = 'vla-lang/testdata/refers_golden.txt'; Source = 'scripts/refers_golden.txt' },
    @{ Copy = 'vla-lang/testdata/view/fixture_frazaro.vla'; Source = 'scripts/view/fixture_frazaro.vla' },
    @{ Copy = 'vla-lang/testdata/view/fixture_output.vla'; Source = 'scripts/view/fixture_output.vla' },
    @{ Copy = 'vla-lang/testdata/view/fixture_data.vla'; Source = 'scripts/view/fixture_data.vla' },
    @{ Copy = 'vla-lang/testdata/view/into_output.vla'; Source = 'scripts/view/into_output.vla' },
    @{ Copy = 'vla-lang/testdata/view/into_checks.vla'; Source = 'scripts/view/into_checks.vla' },
    # VLA.3: the expand fixture and the reference's golden of it
    @{ Copy = 'vla-lang/testdata/expand.txt'; Source = 'scripts/expand.txt' },
    @{ Copy = 'vla-lang/testdata/expand_golden.txt'; Source = 'scripts/expand_golden.txt' },
    # 2026-10-14: the licence texts and the NOTICE, in every package
    @{ Copy = 'vla-lang/LICENSE'; Source = 'LICENSE' },
    @{ Copy = 'vla-lang/NOTICE'; Source = 'NOTICE' },
    @{ Copy = 'core/LICENSE'; Source = 'LICENSE' },
    @{ Copy = 'core/NOTICE'; Source = 'NOTICE' },
    @{ Copy = 'cli/LICENSE'; Source = 'LICENSE' },
    @{ Copy = 'cli/NOTICE'; Source = 'NOTICE' },
    @{ Copy = 'cli/LICENSE-MPL'; Source = 'LICENSES/MPL-2.0.txt' }
)
$expectedPairs = 20   # 2 for the door (2026-10-05), 9 for the language's tests (VLA.1, 2026-10-10), 2 for the expand golden (VLA.3), 7 for the licence files (2026-10-14)

$failures = New-Object System.Collections.Generic.List[string]
if ($pairs.Count -ne $expectedPairs) { $failures.Add("the table holds $($pairs.Count) pair(s), not the $expectedPairs pinned here") }

function Get-FirstDifference([byte[]]$a, [byte[]]$b) {
    $n = [Math]::Min($a.Length, $b.Length)
    for ($i = 0; $i -lt $n; $i++) { if ($a[$i] -ne $b[$i]) { return $i } }
    if ($a.Length -ne $b.Length) { return $n }
    return -1
}

# ---- 1. each copy is its source, byte for byte ----
$sha = [System.Security.Cryptography.SHA256]::Create()
foreach ($p in $pairs) {
    $copyPath = Join-Path $root $p.Copy
    $sourcePath = Join-Path $root $p.Source
    if (-not (Test-Path $sourcePath)) { $failures.Add("$($p.Source) is missing"); continue }
    if (-not (Test-Path $copyPath)) { $failures.Add("$($p.Copy) is missing - run: Copy-Item $($p.Source) $($p.Copy)"); continue }
    $a = [System.IO.File]::ReadAllBytes($copyPath)
    $b = [System.IO.File]::ReadAllBytes($sourcePath)
    if ([Convert]::ToBase64String($sha.ComputeHash($a)) -ne [Convert]::ToBase64String($sha.ComputeHash($b))) {
        $at = Get-FirstDifference $a $b
        $failures.Add("$($p.Copy) differs from $($p.Source) at byte $at ($($a.Length) and $($b.Length) bytes) - run: Copy-Item $($p.Source) $($p.Copy)")
    }
}

# ---- 2. a built door with no flags is the door with the corpus files named ----
function Invoke-Door([string]$exe, [string[]]$arguments) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $exe
    $psi.Arguments = (($arguments | ForEach-Object { '"' + $_ + '"' }) -join ' ')
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $proc = [System.Diagnostics.Process]::Start($psi)
    $out = New-Object System.IO.MemoryStream
    $proc.StandardOutput.BaseStream.CopyTo($out)
    $err = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()
    return @{ Code = $proc.ExitCode; Bytes = $out.ToArray(); Err = $err.Trim() }
}

if ($Impl -eq '') {
    foreach ($p in @('target/debug/frazaro.exe', 'target/release/frazaro.exe', 'target/debug/frazaro', 'target/release/frazaro')) {
        $candidate = Join-Path $root $p
        if (Test-Path $candidate) { $Impl = $candidate; break }
    }
}
$doorNote = 'no built door, so the embedding itself was not run (cargo build --workspace writes one; CI always has it)'
if ($Impl -ne '' -and (Test-Path $Impl)) {
    $program = Join-Path $root 'scripts/instructions.txt'
    $prelude = Join-Path $root 'scripts/prelude.vla'
    $english = Join-Path $root 'scripts/polyglotta/english.vla'
    $bare = Invoke-Door $Impl @('translate-vla', $program)
    $named = Invoke-Door $Impl @('translate-vla', $program, '--prelude', $prelude, '--phrasebook', $english)
    if ($bare.Code -ne 0 -or $named.Code -ne 0) {
        $failures.Add("the door exited $($bare.Code) with no flags ($($bare.Err)) and $($named.Code) with the corpus files named ($($named.Err)) on scripts/instructions.txt")
    } elseif ((Get-FirstDifference $bare.Bytes $named.Bytes) -ge 0) {
        $failures.Add("the door translates scripts/instructions.txt differently with no flags and with --prelude scripts/prelude.vla --phrasebook scripts/polyglotta/english.vla: its built-in pair is not the corpus pair (copy, then cargo build --workspace)")
    } else {
        $doorNote = "the built door translates the corpus program the same with no flags and with the corpus files named ($($bare.Bytes.Length) bytes)"
    }
}

if ($failures.Count -gt 0) {
    Write-Host "FAIL: $($failures.Count) problem(s)"
    $failures | ForEach-Object { Write-Host "  - $_" }
    exit 1
}
Write-Host "OK: $($pairs.Count) copies in the crates' directories are their sources byte for byte; $doorNote"
exit 0
