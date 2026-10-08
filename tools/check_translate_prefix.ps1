<#
check_translate_prefix.ps1 - the core reproduces the translate golden; the
matched prefix never goes down.

WHY: PORT.6 ports VLA_SentenceEngine.bas to the core (core/src/english/), and
the treaty's oracle 1 (conformance/README.md) is what holds the port to the
reference: scripts/instructions.txt, translated with scripts/prelude.vla and
scripts/polyglotta/english.vla, is scripts/instructions_golden.vla less its
GENERATED stamp line, byte for byte after the treaty's normalization (LF,
trailing blank lines dropped). The length of the prefix the port matches is
held as a floor that never goes down, in the house style's hardcoded number
with its dated reason, raised by hand when the corpus grows the golden, so
that a translation matching less than it did yesterday fails here, on every
push, with the first differing line printed from both sides - which
tools/prove.ps1 does not print, and which is what a porter needs first.
check_compile_prefix.ps1 is this check's twin for oracle 1b.

WHAT IT RUNS: `<impl> translate-vla scripts/instructions.txt --prelude
scripts/prelude.vla --phrasebook scripts/polyglotta/english.vla`, the
contract's own command, on the built door at target/debug/frazaro.exe (or
target/release), or on -Impl when given. A tree with no binary says SKIPPED
and exits 0, as check_core_imports.ps1 does, since CI always has it.
-Control proves the check on two fake implementations written to a scratch
directory: one that answers with the golden itself (the whole golden must
match) and one that changes a byte near the start (the prefix must fall
below the floor and fail).

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; exit 0 clean, exit 1 with the difference named.

Usage:  powershell -File tools\check_translate_prefix.ps1
        powershell -File tools\check_translate_prefix.ps1 -Impl C:\path\to\frazaro.exe
        powershell -File tools\check_translate_prefix.ps1 -Control
#>
param(
    [string]$Impl = '',
    [switch]$Control
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# --- the floor: characters of the translate golden matched, never lowered ----
# 2026-10-02, PORT.6 slice 6e: the whole golden on the slice's first green
# run, 224,380 characters (the golden less its stamp, after the treaty's
# normalization: LF line endings, trailing blank lines dropped; with its
# CRLFs the same text is 229,156).
# 2026-10-07, L-SHEET-HELPERS: 225,531 - the eleven sheet macros joined the
# phrasebook and delete-sheet's body became one call, all carried into the
# golden's tail; the whole of the regenerated .vla.
$floor = 225531

$program   = Join-Path $repoRoot 'scripts/instructions.txt'
$goldenVla = Join-Path $repoRoot 'scripts/instructions_golden.vla'
$prelude   = Join-Path $repoRoot 'scripts/prelude.vla'
$book      = Join-Path $repoRoot 'scripts/polyglotta/english.vla'

function Get-NormalizedText([string]$text) {
    if ($null -eq $text) { return '' }
    $t = $text -replace "`r`n", "`n" -replace "`r", "`n"
    return $t.TrimEnd("`n")
}

# The golden less its first line, the writer's stamp (the treaty's amendment
# of 2026-10-02), normalized.
function Read-GoldenLessStamp() {
    $bytes = [System.IO.File]::ReadAllBytes($goldenVla)
    $i = [Array]::IndexOf($bytes, [byte]10)
    $text = [System.Text.Encoding]::UTF8.GetString($bytes, $i + 1, $bytes.Length - $i - 1)
    return Get-NormalizedText $text
}

function Measure-Prefix([string]$impl) {
    $global:LASTEXITCODE = 0
    $lines = & $impl translate-vla $program --prelude $prelude --phrasebook $book 2>$null
    $code = $LASTEXITCODE
    $got = if ($null -eq $lines) { '' } else { (@($lines) | ForEach-Object { [string]$_ }) -join "`n" }
    $got = Get-NormalizedText $got
    $want = Read-GoldenLessStamp
    $n = [Math]::Min($got.Length, $want.Length); $at = $n
    for ($i = 0; $i -lt $n; $i++) { if ($got[$i] -ne $want[$i]) { $at = $i; break } }
    $whole = ($code -eq 0) -and ($got -eq $want)
    return @{ ExitCode = $code; Prefix = $at; Length = $want.Length; Whole = $whole; Got = $got; Want = $want }
}

function Write-FirstDifference($m) {
    if ($m.Whole) { return }
    $at = $m.Prefix
    $line = ($m.Want.Substring(0, $at) -split "`n").Count
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
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_tprefix_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        # The fake answers with the golden less its stamp; the mutant changes
        # one byte near the start.
        $unstamped = Join-Path $tmp 'golden.unstamped.vla'
        [System.IO.File]::WriteAllText($unstamped, (Read-GoldenLessStamp), (New-Object System.Text.UTF8Encoding($false)))
        $u = $unstamped -replace "'", "''"
        $fake = Join-Path $tmp 'fake.ps1'
        [System.IO.File]::WriteAllText($fake, "Write-Output ([System.IO.File]::ReadAllText('$u'))`nexit 0`n")
        $mut = Join-Path $tmp 'mutant.ps1'
        [System.IO.File]::WriteAllText($mut, "`$t = [System.IO.File]::ReadAllText('$u')`nWrite-Output (`$t.Substring(0, 100) + 'X' + `$t.Substring(101))`nexit 0`n")
        $m1 = Measure-Prefix $fake
        $m2 = Measure-Prefix $mut
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

$m = Measure-Prefix $Impl
if ($m.Whole) {
    if ($m.Length -gt $floor) {
        Write-Output ("OK: the whole golden matched, {0} characters, above the floor of {1} - raise the floor in this file" -f $m.Length, $floor)
    } else {
        Write-Output ("OK: the whole golden matched, {0} characters (floor {1})" -f $m.Length, $floor)
    }
    exit 0
}
if ($m.ExitCode -ne 0) {
    Write-Output ("FAIL: {0} exited {1} on the translate golden (a refusal, or no translate-vla command)" -f $Impl, $m.ExitCode)
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
