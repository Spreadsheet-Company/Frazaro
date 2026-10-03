<#
check_build_golden.ps1 - the core reproduces the build goldens byte for byte;
each golden's length never goes down; every entry carries no clock and is
stored unless it is the model's, copied as it was.

WHY: PORT.7 writes workbooks from the core (core/src/sheet, core/src/build),
and the treaty's seventh oracle (conformance/README.md, the amendments of
2026-10-03) holds the writer to its own goldens: scripts/build/fixture.txt,
built with scripts/prelude.vla and scripts/polyglotta/english.vla, is
scripts/build/fixture_golden.xlsx byte for byte; and scripts/build/into.txt,
built into scripts/build/model.xlsx (tools/build_model_fixture.ps1 writes
it), is scripts/build/into_golden.xlsx. The VBA reference writes no workbook
file, so these goldens are not the reference's: they are the core's own
output, blessed by the owner opening them in Excel, and what this check
holds is that the writer never drifts from what was blessed without a
regenerated golden committed beside the change. A zip is bytes, so there is
no normalization, and a difference is named by part and offset, which
tools/prove.ps1 does not print and which is what a porter needs to see first.

ALSO HELD, read off each golden itself: its length, as a floor that never
goes down (the house style's hardcoded, reviewable number); and the
determinism the writer promises - every entry stamped with the zip epoch,
1980-01-01 00:00:00, and every entry STORED (method 0) unless it is one of
the model's, carried with the model's own method, checksum and size - so
that a clock, a compressor or a dependency reaching the writer fails here on
every push, even if someone regenerated a golden with it. And rebuild: the
door's `rebuild` of each golden must say yes (slice 7c).

WHAT IT RUNS: `<impl> build <fixture> --prelude <prelude> --phrasebook
<english.vla> --out <scratch file> [--into <model>]`, the contract's own
command, on the built door at target/debug/frazaro.exe (or target/release),
or on -Impl when given; then `<impl> rebuild <golden> ...`. A tree with no
door says SKIPPED for that half and exits 0, as check_core_imports.ps1 does,
since CI always has it; the floors and the stamps are checked either way.

-Control proves the check on fakes written to a scratch directory: one that
copies the right golden to --out and answers rebuild by comparing to the
goldens (must pass both rows); one that changes a byte inside the first
Frazaro worksheet part of whatever it copies (must fail both rows, naming
the part and the offset); and a copy of the first golden with one entry's
stamp moved off the epoch (the determinism pin must fail it, and pass the
golden as it is).

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; exit 0 clean, exit 1 with every problem named. The door's stderr
is read under ErrorActionPreference Continue, as check_prove_floors.ps1 does,
since a native command's stderr line is a terminating error under Stop. The
goldens are read with FileShare.ReadWrite (Read-BytesShared), because the
owner's live pass has them open in Excel while the checks run, and Excel
holds an open workbook with a lock that refuses a plain ReadAllBytes (found
on the first live pass, 2026-10-03).

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

# --- the goldens and their floors: each length in bytes, never lowered ------
# 2026-10-03, PORT.7 slice 7a: fixture 7544 - the Frazaro sheet alone, the
# eleven lines of the fixture. Raise a floor when a regenerated golden is
# longer; lower it only with a regenerated golden that is shorter, and say why.
# 2026-10-03, slice 7b: fixture 9798 - the Output and data sheets the
# fixture's sentences write; ten entries.
# 2026-10-03, slice 7c: fixture 11390 - the stamp (the defined name
# Frazaro.Build) and the IFS cell as a dynamic-array formula with its
# metadata part; eleven entries.
# 2026-10-03, slice 7d: into 13697 - the fixture model's eleven parts as they
# were (the four edited ones stored) with this writer's metadata part and
# three sheets; fifteen entries.
$goldens = @(
    @{ Name = 'fixture'; Fixture = 'scripts/build/fixture.txt'; Golden = 'scripts/build/fixture_golden.xlsx'; Into = '';                        Floor = 11390 },
    @{ Name = 'into';    Fixture = 'scripts/build/into.txt';    Golden = 'scripts/build/into_golden.xlsx';    Into = 'scripts/build/model.xlsx'; Floor = 13697 }
)
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
# method, stamp, checksum, sizes, and where its data starts, read from its
# own local header as a reader must. Throws when the bytes are not such an
# archive.
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
            Crc = (Read-U32 $b ($pos + 16)); CompressedSize = [int](Read-U32 $b ($pos + 20)); Size = [int](Read-U32 $b ($pos + 24))
            DataStart = $dataStart
        })
        $pos += 46 + $nameLen + $extraLen + $commentLen
    }
    return ,$list
}

# The determinism pin: every stamp the epoch; every entry stored, unless
# the model (when there is one) holds an entry of that name with the same
# method, checksum and size, copied as it was. A list of problems, empty
# when the archive keeps the promise.
function Test-NoClock([byte[]]$b, $modelEntries) {
    $problems = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($e in (Read-Entries $b)) {
            if ($e.Time -ne $epochTime -or $e.Date -ne $epochDate) {
                $problems.Add(("{0} is stamped date {1:X4} time {2:X4}, not the epoch (0021 0000)" -f $e.Name, $e.Date, $e.Time))
            }
            if ($e.Method -eq $stored) { continue }
            $copied = $false
            if ($null -ne $modelEntries) {
                foreach ($m in $modelEntries) {
                    if ($m.Name -eq $e.Name -and $m.Method -eq $e.Method -and $m.Crc -eq $e.Crc -and $m.Size -eq $e.Size) { $copied = $true }
                }
            }
            if (-not $copied) { $problems.Add("$($e.Name) is not stored (method $($e.Method)) and is not one of the model's parts carried as it was") }
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
            if ($at -ge $e.DataStart -and $at -lt ($e.DataStart + $e.CompressedSize)) { $where = "$($e.Name) at offset $($at - $e.DataStart)"; break }
        }
    } catch { $where = "a header ($($_.Exception.Message))" }
    return @{ Equal = $false; Detail = "differs at byte $at of $($want.Length) ($where); the output has $($got.Length) bytes" }
}

# `<impl> <args>` with the door's stderr set aside: its exit code and its
# last line of output.
function Invoke-Door([string]$impl, [string[]]$doorArgs) {
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $global:LASTEXITCODE = 0
        $lines = & $impl @doorArgs 2>$null
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $old
    }
    if ($null -eq $code) { $code = 0 }
    $last = if ($null -eq $lines) { '' } else { [string](@($lines)[-1]) }
    return @{ ExitCode = [int]$code; Last = $last }
}

# One implementation against one golden: a list of problems, empty when its
# build is the golden byte for byte and its rebuild of the golden says yes.
function Measure-Impl([string]$impl, [string]$scratch, $g) {
    $problems = New-Object System.Collections.Generic.List[string]
    $golden = Join-Path $repoRoot $g.Golden
    $built = Join-Path $scratch ($g.Name + '.built.xlsx')
    if (Test-Path $built) { Remove-Item $built -Force }
    $buildArgs = @('build', (Join-Path $repoRoot $g.Fixture), '--prelude', $prelude, '--phrasebook', $english, '--out', $built)
    if ($g.Into -ne '') { $buildArgs += @('--into', (Join-Path $repoRoot $g.Into)) }
    $r = Invoke-Door $impl $buildArgs
    if ($r.ExitCode -eq 3) { $problems.Add("$($g.Name): build not attempted (exit 3), but the golden exists"); return ,$problems }
    if ($r.ExitCode -ne 0) { $problems.Add("$($g.Name): build exited $($r.ExitCode)"); return ,$problems }
    if (-not (Test-Path $built)) { $problems.Add("$($g.Name): build exited 0 but wrote nothing at $built"); return ,$problems }
    $cmp = Compare-Workbooks (Read-BytesShared $golden) (Read-BytesShared $built)
    if (-not $cmp.Equal) { $problems.Add("$($g.Name): the build is not the golden: $($cmp.Detail)"); return ,$problems }
    # 2026-10-03 (slice 7c): the golden must verify as its own build.
    $rb = Invoke-Door $impl @('rebuild', $golden, '--prelude', $prelude, '--phrasebook', $english)
    if ($rb.ExitCode -ne 0) { $problems.Add("$($g.Name): rebuild of the golden exited $($rb.ExitCode) ($($rb.Last))") }
    return ,$problems
}

# Every golden's floor and stamps: problems, and lines to print.
function Test-Goldens() {
    $problems = New-Object System.Collections.Generic.List[string]
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($g in $goldens) {
        $path = Join-Path $repoRoot $g.Golden
        if (-not (Test-Path $path)) { $problems.Add("$($g.Name): no golden at $path"); continue }
        $bytes = Read-BytesShared $path
        if ($bytes.Length -lt $g.Floor) {
            $problems.Add("$($g.Name): the golden is $($bytes.Length) bytes, below its floor of $($g.Floor)")
            $lines.Add("  $($g.Name): $($bytes.Length) bytes, BELOW THE FLOOR of $($g.Floor)  FAIL")
        } elseif ($bytes.Length -gt $g.Floor) {
            $lines.Add("  $($g.Name): $($bytes.Length) bytes, above the floor of $($g.Floor) - raise the floor in this file")
        } else {
            $lines.Add("  $($g.Name): $($g.Floor) bytes, at its floor")
        }
        $modelEntries = $null
        if ($g.Into -ne '') {
            $modelPath = Join-Path $repoRoot $g.Into
            if (-not (Test-Path $modelPath)) { $problems.Add("$($g.Name): no model at $modelPath"); continue }
            try { $modelEntries = Read-Entries (Read-BytesShared $modelPath) } catch { $problems.Add("$($g.Name): the model does not read: $($_.Exception.Message)"); continue }
        }
        $clock = Test-NoClock $bytes $modelEntries
        foreach ($p in $clock) { $problems.Add("$($g.Name): $p"); $lines.Add("  $($g.Name): $p  FAIL") }
        if ($clock.Count -eq 0) {
            $n = (Read-Entries $bytes).Count
            $what = if ($null -eq $modelEntries) { 'every entry stored' } else { "every entry stored or one of the model's as it was" }
            $lines.Add("  $($g.Name): $what, every stamp 1980-01-01 00:00:00 ($n entries)")
        }
    }
    return @{ Problems = $problems; Lines = $lines }
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_build_golden_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $g1 = ($goldens[0].Golden -replace '/', '\') ; $g2 = ($goldens[1].Golden -replace '/', '\')
        $p1 = (Join-Path $repoRoot $g1) -replace "'", "''"
        $p2 = (Join-Path $repoRoot $g2) -replace "'", "''"
        # Both fakes: the golden chosen by --into; rebuild compares the file to either golden.
        $common = "param([Parameter(ValueFromRemainingArguments=`$true)][string[]]`$a)`n" +
                  "function Read-Shared([string]`$p) { `$fs = New-Object System.IO.FileStream(`$p, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite); `$b = New-Object byte[] `$fs.Length; [void]`$fs.Read(`$b, 0, `$b.Length); `$fs.Dispose(); return ,`$b }`n" +
                  "function Same([byte[]]`$x, [byte[]]`$y) { if (`$x.Length -ne `$y.Length) { return `$false }; for (`$i = 0; `$i -lt `$x.Length; `$i++) { if (`$x[`$i] -ne `$y[`$i]) { return `$false } }; return `$true }`n" +
                  "if (`$a[0] -eq 'rebuild') { `$f = [System.IO.File]::ReadAllBytes(`$a[1]); if ((Same `$f (Read-Shared '$p1')) -or (Same `$f (Read-Shared '$p2'))) { Write-Output 'yes'; exit 0 } else { Write-Output 'no'; exit 1 } }`n" +
                  "`$b = if (`$a -contains '--into') { Read-Shared '$p2' } else { Read-Shared '$p1' }`n" +
                  "`$dst = `$a[[Array]::IndexOf(`$a, '--out') + 1]`n"
        $fake = Join-Path $tmp 'fake.ps1'
        [System.IO.File]::WriteAllText($fake, $common + "[System.IO.File]::WriteAllBytes(`$dst, `$b)`nexit 0`n")
        # The mutant changes the byte 40 in from the start of the first Frazaro
        # worksheet part's data, found in the golden it is about to copy.
        $mut = Join-Path $tmp 'mutant.ps1'
        $flip = "`$eocd = `$b.Length - 22; `$cnt = [int]`$b[`$eocd + 10] -bor ([int]`$b[`$eocd + 11] -shl 8); `$pos = [int]`$b[`$eocd + 16] -bor ([int]`$b[`$eocd + 17] -shl 8) -bor ([int]`$b[`$eocd + 18] -shl 16) -bor ([int]`$b[`$eocd + 19] -shl 24)`n" +
                "for (`$k = 0; `$k -lt `$cnt; `$k++) { `$nl = [int]`$b[`$pos + 28] -bor ([int]`$b[`$pos + 29] -shl 8); `$xl = [int]`$b[`$pos + 30] -bor ([int]`$b[`$pos + 31] -shl 8); `$cl = [int]`$b[`$pos + 32] -bor ([int]`$b[`$pos + 33] -shl 8); `$loc = [int]`$b[`$pos + 42] -bor ([int]`$b[`$pos + 43] -shl 8) -bor ([int]`$b[`$pos + 44] -shl 16) -bor ([int]`$b[`$pos + 45] -shl 24); `$nm = [System.Text.Encoding]::UTF8.GetString(`$b, `$pos + 46, `$nl); if (`$nm -like 'xl/worksheets/frazaro_1.xml') { `$ds = `$loc + 30 + ([int]`$b[`$loc + 26] -bor ([int]`$b[`$loc + 27] -shl 8)) + ([int]`$b[`$loc + 28] -bor ([int]`$b[`$loc + 29] -shl 8)); `$b[`$ds + 40] = `$b[`$ds + 40] -bxor 1; break }; `$pos += 46 + `$nl + `$xl + `$cl }`n"
        [System.IO.File]::WriteAllText($mut, $common + $flip + "[System.IO.File]::WriteAllBytes(`$dst, `$b)`nexit 0`n")
        $fakeProblems = 0; $mutantProblems = 0; $named = $true
        foreach ($g in $goldens) {
            $r1 = Measure-Impl $fake $tmp $g
            $r2 = Measure-Impl $mut $tmp $g
            $fakeProblems += $r1.Count; $mutantProblems += $r2.Count
            if ($r2.Count -ne 1 -or -not ($r2[0] -like '*xl/worksheets/frazaro_1.xml at offset 40*')) { $named = $false }
            Write-Output ("control, {0}: the fake has {1} problem(s); the mutant {2}: {3}" -f $g.Name, $r1.Count, $r2.Count, ($r2 -join '; '))
        }
        # The determinism pin: the first golden as it is, and a copy whose
        # first central-directory record says 00:01 instead of 00:00.
        $bytes = Read-BytesShared (Join-Path $repoRoot $goldens[0].Golden)
        $r0 = Test-NoClock $bytes $null
        $clocked = [byte[]]$bytes.Clone()
        $cd = [int](Read-U32 $bytes ($bytes.Length - 22 + 16))
        $clocked[$cd + 12] = 0x20
        $r3 = Test-NoClock $clocked $null
        Write-Output ("control: the first golden's stamps {0} problem(s); the clocked copy {1}: {2}" -f $r0.Count, $r3.Count, ($r3 -join '; '))
        if ($fakeProblems -eq 0 -and $named -and $mutantProblems -eq $goldens.Count -and $r0.Count -eq 0 -and $r3.Count -eq 1) {
            Write-Output 'OK: the check passes both goldens, fails the mutant on each naming its part and offset, and catches a clock'
            exit 0
        }
        Write-Output 'FAIL: the control did not behave as the header says'
        exit 1
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}

Write-Output '=== THE BUILD GOLDENS: REPRODUCED BYTE FOR BYTE, NEVER SHORTER, NO CLOCK (PORT.7, oracle 7) ==='
$failed = New-Object System.Collections.Generic.List[string]
$tg = Test-Goldens
$tg.Lines | ForEach-Object { Write-Output $_ }
foreach ($p in $tg.Problems) { $failed.Add($p) }

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
        foreach ($g in $goldens) {
            $r = Measure-Impl $Impl $scratch $g
            foreach ($p in $r) { $failed.Add($p); Write-Output "  $p  FAIL" }
            if ($r.Count -eq 0) { Write-Output "  $($g.Name): the door's build is the golden, byte for byte, and rebuild says yes" }
        }
    } finally {
        Remove-Item -Recurse -Force $scratch -ErrorAction SilentlyContinue
    }
}

Write-Output ''
if ($failed.Count -eq 0) {
    Write-Output '=== CHECK: clean - both build goldens are reproduced, at or above their floors, with no clock ==='
    exit 0
}
Write-Output "=== CHECK: $($failed.Count) problem(s) ==="
$failed | ForEach-Object { Write-Output "  $_" }
exit 1
