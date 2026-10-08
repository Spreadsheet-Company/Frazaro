<#
check_compile_prefix.ps1 - the core reproduces the compile golden; the
matched prefix never goes down.

WHY: PORT.5 ports the reader, the macroexpander and the emitters of VLA.bas
to the core (core/), and the treaty's oracle 1b (conformance/README.md,
amendment of 2026-10-01) is what holds the port to the reference:
scripts/instructions_golden.vla, less its GENERATED stamp line, compiled with
scripts/prelude.vla, is scripts/instructions_golden.vba byte for byte after
the treaty's normalization (LF, trailing blank lines dropped). While the port
was being built the plan was to hold the length of the prefix it matched as a
floor that never goes down; the prefix reached the end the same day the
emitters were written (2026-10-01: 301,861 characters, the whole golden), so
the floor below is the whole golden's length. It stays a hardcoded number in
the house style, raised by hand when the corpus grows the golden, so that a
compile that matches less than it did yesterday fails here, on every push,
with the first differing line printed from both sides - which tools/prove.ps1
does not print, and which is what a porter needs to see first.

WHAT IT RUNS: `<impl> compile <unstamped golden> --prelude scripts/prelude.vla`,
the contract's own command, on the built door at target/debug/frazaro.exe
(or target/release), or on -Impl when given. A tree with no binary says
SKIPPED and exits 0, as check_core_imports.ps1 does, since CI always has it.

-Control proves the check on two fake implementations written to a scratch
directory: one that answers with the golden itself (the whole golden must
match) and one that changes a byte near the start (the prefix must fall
below the floor and fail).

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; exit 0 clean, exit 1 with the difference named.

Usage:  powershell -File tools\check_compile_prefix.ps1
        powershell -File tools\check_compile_prefix.ps1 -Impl C:\path\to\frazaro.exe
        powershell -File tools\check_compile_prefix.ps1 -Control
#>
param(
    [string]$Impl = '',
    [switch]$Control
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# --- the floor: the matched prefix, in characters after normalization ------
# 2026-10-01, PORT.5: 301861 - the whole of instructions_golden.vba, the day
# the prefix reached the end. Raise it when a regenerated golden is longer;
# lower it only with a regenerated golden that is shorter, and say so.
# 2026-10-02, PORT.6 step 0: 291316 - the goldens regenerated without F.9's
# section markers (the treaty's amendment of that date, the owner's call):
# 153 comment lines gone and every vla:N tag after the first marker moved
# down; the whole of the regenerated .vba.
# 2026-10-05, SEC.15: 291496 - the golden's eighteen formula writes are
# calls to the runtime's VlaSetFormula (ten characters longer each), the
# whole of the regenerated .vba.
# 2026-10-07, L-SHEET-HELPERS: 291184 - LOWERED, with the regenerated golden:
# the corpus's four "Delete sheet" sentences compile to one call of the
# runtime's VlaDeleteSheet each where they were three lines (DisplayAlerts
# off, the delete, DisplayAlerts on); the whole of the regenerated .vba.
$floor = 291184

$goldenVla = Join-Path $repoRoot 'scripts/instructions_golden.vla'
$goldenVba = Join-Path $repoRoot 'scripts/instructions_golden.vba'
$prelude   = Join-Path $repoRoot 'scripts/prelude.vla'

function Get-NormalizedText([string]$text) {
    if ($null -eq $text) { return '' }
    $t = $text -replace "`r`n", "`n" -replace "`r", "`n"
    return $t.TrimEnd("`n")
}

# The stamp-less golden, as prove.ps1 writes it: the bytes after the first
# line feed, copied, not re-encoded.
function Write-Unstamped([string]$dir) {
    $bytes = [System.IO.File]::ReadAllBytes($goldenVla)
    $i = [Array]::IndexOf($bytes, [byte]10)
    $rest = New-Object byte[] ($bytes.Length - $i - 1)
    [Array]::Copy($bytes, $i + 1, $rest, 0, $rest.Length)
    $out = Join-Path $dir 'instructions_golden.unstamped.vla'
    [System.IO.File]::WriteAllBytes($out, $rest)
    return $out
}

function Measure-Prefix([string]$impl, [string]$dir) {
    $program = Write-Unstamped $dir
    $global:LASTEXITCODE = 0
    $lines = & $impl compile $program --prelude $prelude 2>$null
    $code = $LASTEXITCODE
    $got = if ($null -eq $lines) { '' } else { (@($lines) | ForEach-Object { [string]$_ }) -join "`n" }
    $got = Get-NormalizedText $got
    $want = Get-NormalizedText ([System.IO.File]::ReadAllText($goldenVba))
    $n = [Math]::Min($got.Length, $want.Length); $at = $n
    for ($i = 0; $i -lt $n; $i++) { if ($got[$i] -ne $want[$i]) { $at = $i; break } }
    $whole = ($code -eq 0) -and ($got -eq $want)
    return @{ ExitCode = $code; Prefix = $at; Length = $want.Length; Whole = $whole; Got = $got; Want = $want }
}

function Write-FirstDifference($m) {
    if ($m.Whole) { return }
    $at = $m.Prefix
    $line = ($m.Want.Substring(0, $at) -split "`n").Count
    $wantLine = ($m.Want.Substring(0, [Math]::Min($m.Want.Length, $at + 200)) -split "`n")[-1]
    $gotLine  = ($m.Got.Substring(0,  [Math]::Min($m.Got.Length,  $at + 200)) -split "`n")[-1]
    $wantFrom = $m.Want.Substring(0, $at)
    $lineStart = $wantFrom.LastIndexOf("`n") + 1
    $wantRest = if ($at -lt $m.Want.Length) { $m.Want.Substring($at) } else { '' }
    $gotRest  = if ($at -lt $m.Got.Length)  { $m.Got.Substring($at)  } else { '' }
    $wantEnd = $wantRest.IndexOf("`n"); if ($wantEnd -lt 0) { $wantEnd = $wantRest.Length }
    $gotEnd  = $gotRest.IndexOf("`n");  if ($gotEnd -lt 0)  { $gotEnd  = $gotRest.Length }
    Write-Output ("  first difference at char {0} of {1}, golden line {2}:" -f $at, $m.Length, $line)
    Write-Output ("    want: {0}{1}" -f $m.Want.Substring($lineStart, $at - $lineStart), $wantRest.Substring(0, $wantEnd))
    Write-Output ("    got:  {0}{1}" -f $m.Got.Substring($lineStart, [Math]::Min($at - $lineStart, [Math]::Max(0, $m.Got.Length - $lineStart))), $gotRest.Substring(0, $gotEnd))
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_prefix_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $g = $goldenVba -replace "'", "''"
        $fake = Join-Path $tmp 'fake.ps1'
        [System.IO.File]::WriteAllText($fake, "Write-Output ([System.IO.File]::ReadAllText('$g'))`nexit 0`n")
        $mut = Join-Path $tmp 'mutant.ps1'
        [System.IO.File]::WriteAllText($mut, "`$t = [System.IO.File]::ReadAllText('$g')`nWrite-Output (`$t.Substring(0, 100) + 'X' + `$t.Substring(101))`nexit 0`n")
        $m1 = Measure-Prefix $fake $tmp
        $m2 = Measure-Prefix $mut $tmp
        Write-Output ("control: the fake matched {0} of {1} (whole: {2}); the mutant matched {3}" -f $m1.Prefix, $m1.Length, $m1.Whole, $m2.Prefix)
        if ($m1.Whole -and $m1.Length -ge $floor -and -not $m2.Whole -and $m2.Prefix -lt $floor) {
            Write-Output 'OK: the check passes the golden and fails the mutant'
            exit 0
        }
        Write-Output 'FAIL: the control did not behave'
        exit 1
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}

if ($Impl -eq '') {
    foreach ($candidate in @('target/debug/frazaro.exe', 'target/release/frazaro.exe', 'target/debug/frazaro', 'target/release/frazaro')) {
        $p = Join-Path $repoRoot $candidate
        if (Test-Path $p) { $Impl = $p; break }
    }
}
if ($Impl -eq '' -or -not (Test-Path $Impl)) {
    Write-Output 'SKIPPED: no built door at target/debug/frazaro.exe or target/release/frazaro.exe (cargo build --workspace writes it; CI always has it)'
    exit 0
}

$scratch = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_prefix_' + [System.IO.Path]::GetRandomFileName())
New-Item -ItemType Directory -Path $scratch | Out-Null
try {
    $m = Measure-Prefix $Impl $scratch
} finally {
    Remove-Item -Recurse -Force $scratch -ErrorAction SilentlyContinue
}

if ($m.Whole) {
    if ($m.Length -gt $floor) {
        Write-Output ("OK: the whole golden matched, {0} characters, above the floor of {1} - raise the floor in this file" -f $m.Length, $floor)
    } else {
        Write-Output ("OK: the whole golden matched, {0} characters (floor {1})" -f $m.Length, $floor)
    }
    exit 0
}
if ($m.ExitCode -ne 0) {
    Write-Output ("FAIL: {0} exited {1} on the compile golden (a refusal, or no compile command)" -f $Impl, $m.ExitCode)
    Write-FirstDifference $m
    exit 1
}
if ($m.Prefix -lt $floor) {
    Write-Output ("FAIL: the matched prefix is {0} characters, below the floor of {1} ({2} in the golden)" -f $m.Prefix, $floor, $m.Length)
    Write-FirstDifference $m
    exit 1
}
Write-Output ("OK: the matched prefix is {0} characters, at or above the floor of {1}, but not the whole golden ({2}) - raise the floor as it grows" -f $m.Prefix, $floor, $m.Length)
Write-FirstDifference $m
exit 0
