<#
check_build_golden.ps1 - the core reproduces the build golden byte for byte;
the golden's length never goes down; every entry is stored and carries no clock.

WHY: PORT.7 writes workbooks from the core (core/src/sheet, core/src/build),
and the treaty's seventh oracle (conformance/README.md, the amendment of
2026-10-03) holds the writer to its own golden: scripts/build/fixture.txt,
built with scripts/prelude.vla and scripts/polyglotta/english.vla, is
scripts/build/fixture_golden.xlsx byte for byte. The VBA reference writes no
workbook file, so this golden is not the reference's: it is the core's own
output, blessed by the owner opening it in Excel, and what this check holds
is that the writer never drifts from what was blessed without a regenerated
golden committed beside the change. A zip is bytes, so there is no
normalization, and a difference is named by part and offset, which
tools/prove.ps1 does not print and which is what a porter needs to see first.

ALSO HELD, read off the golden itself: its length, as a floor that never
goes down (the house style's hardcoded, reviewable number); and the
determinism the writer promises - every entry STORED (method 0) and stamped
with the zip epoch, 1980-01-01 00:00:00 - so that a clock, a compressor or a
dependency reaching the writer fails here on every push, even if someone
regenerated the golden with it.

WHAT IT RUNS: `<impl> build <fixture> --prelude <prelude> --phrasebook
<english.vla> --out <scratch file>`, the contract's own command, on the built
door at target/debug/frazaro.exe (or target/release), or on -Impl when given.
A tree with no door says SKIPPED and exits 0 for that half, as
check_core_imports.ps1 does, since CI always has it; the floor and the stamps
are checked either way.

-Control proves the check on fakes written to a scratch directory: one that
copies the golden to --out (must pass); one that changes a byte inside the
worksheet part (must fail, naming that part and the offset); and a copy of
the golden with one entry's stamp moved off the epoch (the determinism pin
must fail it, and must pass the golden as it is).

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; exit 0 clean, exit 1 with every problem named. The door's stderr
is read under ErrorActionPreference Continue, as check_prove_floors.ps1 does,
since a native command's stderr line is a terminating error under Stop. The
golden is read with FileShare.ReadWrite (Read-BytesShared), because the
owner's live pass has it open in Excel while the checks run, and Excel holds
an open workbook with a lock that refuses a plain ReadAllBytes (found on the
first live pass, 2026-10-03).

Usage:  powershell -File tools\check_build_golden.ps1
        powershell -File tools\check_build_golden.ps1 -Impl C:\path\to\frazaro.exe
        powershell -File tools\check_build_golden.ps1 -Control
#>
param(
    [string]$Impl = '',
    [switch]$Control
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# --- the floor: the golden's length in bytes, never lowered -----------------
# 2026-10-03, PORT.7 slice 7a: 7544 - the Frazaro sheet alone, the eleven
# lines of the fixture. Raise it when a regenerated golden is longer; lower it
# only with a regenerated golden that is shorter, and say why.
# 2026-10-03, PORT.7 slice 7b: 9798 - the Output and data sheets the fixture's
# sentences write (values, a formula, a shared formula over three cells, a
# prefixed IFS) beside the Frazaro sheet; ten entries.
# 2026-10-03, PORT.7 slice 7c: 11390 - the stamp (the defined name
# Frazaro.Build in workbook.xml) and the IFS cell as a dynamic-array formula
# with its metadata part (xl/metadata.xml); eleven entries.
$floor = 11390

$fixture = Join-Path $repoRoot 'scripts/build/fixture.txt'
$golden  = Join-Path $repoRoot 'scripts/build/fixture_golden.xlsx'
$prelude = Join-Path $repoRoot 'scripts/prelude.vla'
$english = Join-Path $repoRoot 'scripts/polyglotta/english.vla'

# The writer's promises, as the zip format spells them: method 0 is stored;
# the MS-DOS stamp 0000/0021 is 00:00:00 on 1980-01-01.
$stored = 0
$epochTime = 0
$epochDate = 0x21

function Read-U16([byte[]]$b, [int]$at) {
    return ([int]$b[$at]) -bor (([int]$b[$at + 1]) -shl 8)
}
function Read-U32([byte[]]$b, [int]$at) {
    return ([long]$b[$at]) -bor (([long]$b[$at + 1]) -shl 8) -bor (([long]$b[$at + 2]) -shl 16) -bor (([long]$b[$at + 3]) -shl 24)
}

# A file's bytes, read with sharing that tolerates a workbook open in Excel.
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

# The central directory of an archive with no comment: each entry's name,
# method, stamp, size, and where its data starts, read from its own local
# header as a reader must. Throws when the bytes are not such an archive.
function Read-Entries([byte[]]$b) {
    if ($b.Length -lt 22) { throw 'shorter than an end-of-central-directory record' }
    $eocd = $b.Length - 22
    if ((Read-U32 $b $eocd) -ne 0x06054B50) { throw 'no end-of-central-directory record at the end (a comment, or not a zip)' }
    $count = Read-U16 $b ($eocd + 10)
    $pos = [int](Read-U32 $b ($eocd + 16))
    $list = New-Object System.Collections.Generic.List[object]
    for ($k = 0; $k -lt $count; $k++) {
        if ((Read-U32 $b $pos) -ne 0x02014B50) { throw "no central directory header at byte $pos" }
        $nameLen = Read-U16 $b ($pos + 28)
        $extraLen = Read-U16 $b ($pos + 30)
        $commentLen = Read-U16 $b ($pos + 32)
        $local = [int](Read-U32 $b ($pos + 42))
        if ((Read-U32 $b $local) -ne 0x04034B50) { throw "no local header at byte $local" }
        $dataStart = $local + 30 + (Read-U16 $b ($local + 26)) + (Read-U16 $b ($local + 28))
        $list.Add(@{
            Name = [System.Text.Encoding]::UTF8.GetString($b, $pos + 46, $nameLen)
            Method = (Read-U16 $b ($pos + 10)); Time = (Read-U16 $b ($pos + 12)); Date = (Read-U16 $b ($pos + 14))
            Size = [int](Read-U32 $b ($pos + 24)); DataStart = $dataStart
        })
        $pos += 46 + $nameLen + $extraLen + $commentLen
    }
    return ,$list
}

# The determinism pin: every entry stored, every stamp the epoch. A list of
# problems, empty when the archive keeps the promise.
function Test-NoClock([byte[]]$b) {
    $problems = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($e in (Read-Entries $b)) {
            if ($e.Method -ne $stored) { $problems.Add("$($e.Name) is not stored (method $($e.Method))") }
            if ($e.Time -ne $epochTime -or $e.Date -ne $epochDate) {
                $problems.Add(("{0} is stamped date {1:X4} time {2:X4}, not the epoch (0021 0000)" -f $e.Name, $e.Date, $e.Time))
            }
        }
    } catch {
        $problems.Add("not an archive this check can read: $($_.Exception.Message)")
    }
    return ,$problems
}

# Two archives compared: equal, or the first differing byte named by the
# golden's part it falls in and the offset within that part.
function Compare-Workbooks([byte[]]$want, [byte[]]$got) {
    $n = [Math]::Min($want.Length, $got.Length); $at = $n
    for ($i = 0; $i -lt $n; $i++) { if ($want[$i] -ne $got[$i]) { $at = $i; break } }
    if ($at -eq $n -and $want.Length -eq $got.Length) { return @{ Equal = $true; Detail = "$($want.Length) bytes" } }
    $where = 'a header'
    try {
        foreach ($e in (Read-Entries $want)) {
            if ($at -ge $e.DataStart -and $at -lt ($e.DataStart + $e.Size)) { $where = "$($e.Name) at offset $($at - $e.DataStart)"; break }
        }
    } catch { $where = "a header ($($_.Exception.Message))" }
    return @{ Equal = $false; Detail = "differs at byte $at of $($want.Length) ($where); the output has $($got.Length) bytes" }
}

# `<impl> build ...` into a file, its exit code; the door's stderr set aside.
function Invoke-Build([string]$impl, [string]$outFile) {
    if (Test-Path $outFile) { Remove-Item $outFile -Force }
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $global:LASTEXITCODE = 0
        & $impl build $fixture --prelude $prelude --phrasebook $english --out $outFile 2>$null | Out-Null
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $old
    }
    if ($null -eq $code) { $code = 0 }
    return [int]$code
}

# One implementation against the golden: a list of problems, empty when its
# build is the golden byte for byte.
function Measure-Impl([string]$impl, [string]$scratch) {
    $problems = New-Object System.Collections.Generic.List[string]
    $built = Join-Path $scratch 'fixture.built.xlsx'
    $code = Invoke-Build $impl $built
    if ($code -eq 3) { $problems.Add('build not attempted (exit 3), but the golden exists'); return ,$problems }
    if ($code -ne 0) { $problems.Add("build exited $code"); return ,$problems }
    if (-not (Test-Path $built)) { $problems.Add("build exited 0 but wrote nothing at $built"); return ,$problems }
    $cmp = Compare-Workbooks (Read-BytesShared $golden) (Read-BytesShared $built)
    if (-not $cmp.Equal) { $problems.Add("the build is not the golden: $($cmp.Detail)"); return ,$problems }
    # 2026-10-03 (slice 7c): the golden must verify as its own build.
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $global:LASTEXITCODE = 0
        $lines = & $impl rebuild $golden --prelude $prelude --phrasebook $english 2>$null
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $old
    }
    if ($null -eq $code) { $code = 0 }
    if ([int]$code -ne 0) {
        $last = if ($null -eq $lines) { '' } else { @($lines)[-1] }
        $problems.Add("rebuild of the golden exited $code ($last)")
    }
    return ,$problems
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_build_golden_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $bytes = Read-BytesShared $golden
        $g = $golden -replace "'", "''"
        # The fakes read the golden with the same sharing, for the same reason.
        # Both fakes answer rebuild (slice 7c): yes when the file is the golden.
        $copy = "param([Parameter(ValueFromRemainingArguments=`$true)][string[]]`$a)`n" +
                "`$fs = New-Object System.IO.FileStream('$g', [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)`n" +
                "`$b = New-Object byte[] `$fs.Length`n[void]`$fs.Read(`$b, 0, `$b.Length)`n`$fs.Dispose()`n" +
                "if (`$a[0] -eq 'rebuild') {`n  `$f = [System.IO.File]::ReadAllBytes(`$a[1])`n  `$same = (`$f.Length -eq `$b.Length)`n" +
                "  if (`$same) { for (`$i = 0; `$i -lt `$f.Length; `$i++) { if (`$f[`$i] -ne `$b[`$i]) { `$same = `$false; break } } }`n" +
                "  if (`$same) { Write-Output 'yes'; exit 0 } else { Write-Output 'no'; exit 1 }`n}`n" +
                "`$dst = `$a[[Array]::IndexOf(`$a, '--out') + 1]`n"
        $fake = Join-Path $tmp 'fake.ps1'
        [System.IO.File]::WriteAllText($fake, $copy + "[System.IO.File]::WriteAllBytes(`$dst, `$b)`nexit 0`n")
        # The mutant changes the byte 40 in from the start of the worksheet
        # part's data. (A foreach, not a pipeline: Read-Entries returns one
        # List, which a pipeline would hand on whole.)
        $sheet = $null
        foreach ($e in (Read-Entries $bytes)) { if ($e.Name -like 'xl/worksheets/*') { $sheet = $e; break } }
        if ($null -eq $sheet) { Write-Output 'FAIL: the golden has no worksheet part'; exit 1 }
        $flipAt = $sheet.DataStart + 40
        $mut = Join-Path $tmp 'mutant.ps1'
        [System.IO.File]::WriteAllText($mut, $copy + "`$b[$flipAt] = `$b[$flipAt] -bxor 1`n[System.IO.File]::WriteAllBytes(`$dst, `$b)`nexit 0`n")
        $r1 = Measure-Impl $fake $tmp
        $r2 = Measure-Impl $mut $tmp
        # The determinism pin: the golden as it is, and a copy whose first
        # central-directory record says 00:01 instead of 00:00.
        $r0 = Test-NoClock $bytes
        $clocked = [byte[]]$bytes.Clone()
        $cd = [int](Read-U32 $bytes ($bytes.Length - 22 + 16))
        $clocked[$cd + 12] = 0x20
        $r3 = Test-NoClock $clocked
        Write-Output ("control: the fake has {0} problem(s); the mutant {1}: {2}; the golden's stamps {3} problem(s); the clocked copy {4}: {5}" -f $r1.Count, $r2.Count, ($r2 -join '; '), $r0.Count, $r3.Count, ($r3 -join '; '))
        $named = ($r2.Count -eq 1) -and ($r2[0] -like "*$($sheet.Name) at offset 40*")
        if ($r1.Count -eq 0 -and $named -and $r0.Count -eq 0 -and $r3.Count -eq 1) {
            Write-Output 'OK: the check passes the golden, fails the mutant naming its part and offset, and catches a clock'
            exit 0
        }
        Write-Output 'FAIL: the control did not behave as the header says'
        exit 1
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}

Write-Output '=== THE BUILD GOLDEN: REPRODUCED BYTE FOR BYTE, NEVER SHORTER, NO CLOCK (PORT.7, oracle 7) ==='
$failed = New-Object System.Collections.Generic.List[string]
if (-not (Test-Path $golden)) { Write-Output "FAIL: no golden at $golden"; exit 1 }
$goldenBytes = Read-BytesShared $golden
if ($goldenBytes.Length -lt $floor) {
    $failed.Add("the golden is $($goldenBytes.Length) bytes, below its floor of $floor")
    Write-Output "  the golden is $($goldenBytes.Length) bytes, BELOW THE FLOOR of $floor  FAIL"
} elseif ($goldenBytes.Length -gt $floor) {
    Write-Output "  the golden is $($goldenBytes.Length) bytes, above the floor of $floor - raise the floor in this file"
} else {
    Write-Output "  the golden is $floor bytes, at its floor"
}
$clock = Test-NoClock $goldenBytes
foreach ($p in $clock) { $failed.Add($p); Write-Output "  $p  FAIL" }
if ($clock.Count -eq 0) {
    Write-Output ("  every entry stored and stamped 1980-01-01 00:00:00 ({0} entries)" -f (Read-Entries $goldenBytes).Count)
}

if ($Impl -eq '') {
    foreach ($candidate in @('target/debug/frazaro.exe', 'target/release/frazaro.exe', 'target/debug/frazaro', 'target/release/frazaro')) {
        $p = Join-Path $repoRoot $candidate
        if (Test-Path $p) { $Impl = $p; break }
    }
}
if ($Impl -eq '' -or -not (Test-Path $Impl)) {
    Write-Output '  SKIPPED: no built door at target/debug/frazaro.exe or target/release/frazaro.exe (cargo build --workspace writes it; CI always has it)'
} else {
    $scratch = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_build_golden_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $scratch | Out-Null
    try {
        $r = Measure-Impl $Impl $scratch
        foreach ($p in $r) { $failed.Add($p); Write-Output "  $p  FAIL" }
        if ($r.Count -eq 0) { Write-Output "  the door's build is the golden, byte for byte ($($goldenBytes.Length) bytes)" }
    } finally {
        Remove-Item -Recurse -Force $scratch -ErrorAction SilentlyContinue
    }
}

Write-Output ''
if ($failed.Count -eq 0) {
    Write-Output '=== CHECK: clean - the build golden is reproduced, at or above its floor, with no clock ==='
    exit 0
}
Write-Output "=== CHECK: $($failed.Count) problem(s) ==="
$failed | ForEach-Object { Write-Output "  $_" }
exit 1
