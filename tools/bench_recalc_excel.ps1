<#
bench_recalc_excel.ps1 - the recalculation floors' Excel half: the seven
fixtures the native runner built, opened in Excel and timed doing the work
the machine was timed doing, and the flag's three files read; printed as one
block to paste into tools/check_recalc_floors.ps1 beside the native half's.

WHY: KERNEL.20's sixth slice (docs/BETA_REARVIEW.md, its scoping). The item
is "speed, measured against Excel on the fixtures", and CALLOSUM section
12.6 calls the fastest spreadsheet a prediction until it is measured. The
native half (tools/bench_recalc.ps1) times the machine; this half times
Excel on the same files, byte for byte (each file's SHA-256 is printed and
the check holds the two halves to the same bytes), and sums the same ranges,
so that the two are known to have computed the same values.

THE OWNER RUNS IT: it starts Excel through COM, invisible, and no session
starts Office without the owner's word. Close your own workbooks first, or
leave them: this Excel is an instance of its own, and its settings are put
back before it quits. Life's frames may be copied through the clipboard,
so something you copied before the run may be replaced.

WHAT IT DOES:
  1. reads target/bench/ (run tools/bench_recalc.ps1 first: the runner
     writes the files there) and each fixture's "# bench:" comments from
     scripts/recalc/bench/, the same comments the runner reads;
  2. the flag: calcpr_built.xlsx, calcpr_calcid.xlsx and calcpr_none.xlsx,
     each in an Excel of its own, opened, Output!C2 read before anything is
     calculated, the workbook's Saved state read, closed unsaved. Each holds
     999 where the build folded 15: 999 read back is a value Excel kept
     from the file, 15 a value Excel computed on opening it;
  3. one Excel for the fixtures, its process opted out of Windows' power
     throttling (an invisible process is a background one, slowed by about
     40% a few seconds in, measured for Alonzo's CART.1). Each fixture opened
     read-only and timed (the load with fullCalcOnLoad's calculation), its
     sum read, calculation set to manual; then twice, with multi-threaded
     calculation as Excel ships it and with one thread on the performance
     cores: the change timed, Calculate after each write (an edit cell
     written with its other value and back, 21 times; or Life's frames 2 to
     31, the previous frame's values copied into Screen.last, timed apart
     from the frame, and the clock advanced; the copy is made by the first of
     four ways that Excel accepts and holds equal cell for cell: the range
     handed over, Excel's own copy and paste of values, which goes through
     the clipboard, and the array of values as it came and from 0), the sum read after; CalculateFull timed 11 times;
     CalculateFullRebuild timed 5 times; and a Calculate with nothing to
     compute, timed 21 times, the floor of what a COM call adds to each;
  4. Excel's multi-threaded setting put back as it was (Excel saves it in
     its options), every workbook closed unsaved, Excel quit;
  5. prints the block and writes it to target/bench_recalc_excel.txt, to
     copy from (a terminal that wraps a long line can drop a space).
It is an instrument and not a check: it measures, the owner pastes, and the
check holds what was pasted.

-Fake runs the whole script against a stand-in for Excel that computes
nothing, so that its control flow and its block are proven on a machine
where no Excel may be started; its numbers mean nothing and it says so.
-FakeRefuse <fixture> makes the stand-in refuse to open that fixture, as
Excel refused life.xlsx on the owner's first run (2026-10-10): the
fixture's two rows say so and the others are measured all the same.

House style (tools/*.ps1): PowerShell 5.1; ASCII; no network.

Usage:  powershell -File tools\bench_recalc_excel.ps1
        powershell -File tools\bench_recalc_excel.ps1 -Fake
#>
param(
    [switch]$Fake,
    [string]$FakeRefuse = ''
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$bench = Join-Path $repoRoot 'target/bench'
$fixtures = @('column', 'chain', 'sums', 'life', 'lookups', 'criteria', 'netlist')
$edits = 21
$wholeRuns = 11
$rebuildRuns = 5
$floorRuns = 21
$xlCalculationManual = -4135

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class RecalcExcelProcess {
    [StructLayout(LayoutKind.Sequential)]
    public struct State { public uint Version; public uint ControlMask; public uint StateMask; }
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool SetProcessInformation(IntPtr process, int infoClass, ref State info, uint size);
    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    // ProcessPowerThrottling is class 4; the execution-speed bit is 1; a state of 0 opts the process out.
    public static bool OptOut(IntPtr process) {
        var s = new State(); s.Version = 1; s.ControlMask = 1; s.StateMask = 0;
        return SetProcessInformation(process, 4, ref s, (uint)Marshal.SizeOf(typeof(State)));
    }
    public static uint ProcessOf(IntPtr window) { uint p; GetWindowThreadProcessId(window, out p); return p; }
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool GetSystemCpuSetInformation(IntPtr information, uint length, out uint returned, IntPtr process, uint flags);
    // Group 0's logical processors as pairs, its index then its efficiency class.
    public static int[] Classes() {
        uint length;
        GetSystemCpuSetInformation(IntPtr.Zero, 0, out length, IntPtr.Zero, 0);
        var pairs = new System.Collections.Generic.List<int>();
        if (length == 0) return pairs.ToArray();
        IntPtr buffer = Marshal.AllocHGlobal((int)length);
        try {
            if (!GetSystemCpuSetInformation(buffer, length, out length, IntPtr.Zero, 0)) return pairs.ToArray();
            int at = 0;
            while (at + 20 <= length) {
                int size = Marshal.ReadInt32(buffer, at);
                if (size <= 0) break;
                if (Marshal.ReadInt32(buffer, at + 4) == 0 && Marshal.ReadInt16(buffer, at + 12) == 0) {
                    pairs.Add(Marshal.ReadByte(buffer, at + 14));
                    pairs.Add(Marshal.ReadByte(buffer, at + 18));
                }
                at += size;
            }
        } finally { Marshal.FreeHGlobal(buffer); }
        return pairs.ToArray();
    }
}
"@

# The performance cores as an affinity mask; 0 on a machine of one kind of core.
function Get-PerformanceMask {
    try { $pairs = @([RecalcExcelProcess]::Classes()) } catch { return [long]0 }
    if ($pairs.Count -lt 2) { return [long]0 }
    $classes = @(for ($i = 1; $i -lt $pairs.Count; $i += 2) { $pairs[$i] })
    $top = ($classes | Measure-Object -Maximum).Maximum
    $mask = [long]0; $fast = 0
    for ($i = 0; $i -lt $pairs.Count; $i += 2) {
        if ($pairs[$i + 1] -eq $top -and $pairs[$i] -lt 64) { $mask = $mask -bor ([long]1 -shl $pairs[$i]); $fast++ }
    }
    if ($fast -eq $classes.Count) { return [long]0 }
    return $mask
}

# A fixture's "# bench:" comments, as the runner reads them.
function Read-Bench([string]$name) {
    $path = Join-Path $repoRoot "scripts/recalc/bench/$name.txt"
    $b = @{ Name = $name; Kind = ''; Sheet = ''; Cell = ''; Built = ''; Other = ''; Count = 0; Twin = ''; SumSheet = ''; SumArea = '' }
    foreach ($line in [System.IO.File]::ReadAllLines($path)) {
        $t = $line.Trim()
        if (-not $t.StartsWith('# bench:')) { continue }
        $w = @($t.Substring(8).Trim() -split '\s+')
        if ($w[0] -eq 'edit' -and $w.Count -eq 4) {
            $b.Kind = 'edit'; $b.Sheet, $b.Cell = $w[1] -split '!', 2; $b.Built = $w[2]; $b.Other = $w[3]
        } elseif ($w[0] -eq 'frames' -and $w.Count -eq 7 -and $w[3] -eq 'twin' -and $w[5] -eq 'seed') {
            $b.Kind = 'frames'; $b.Sheet, $b.Cell = $w[1] -split '!', 2; $b.Count = [int]$w[2]; $b.Twin = $w[4]
        } elseif ($w[0] -eq 'sum' -and $w.Count -eq 2) {
            $b.SumSheet, $b.SumArea = $w[1] -split '!', 2
        } else { throw "$name.txt: a bench comment this script does not read: $t" }
    }
    if ($b.Kind -eq '' -or $b.SumArea -eq '') { throw "$name.txt: its bench comments do not say what changes and what is summed" }
    return $b
}

function Get-Median([double[]]$v) {
    $s = @($v | Sort-Object)
    if ($s.Count -eq 0) { return 0.0 }
    if ($s.Count % 2 -eq 1) { return [double]$s[[int](($s.Count - 1) / 2)] }
    return ([double]$s[$s.Count / 2 - 1] + [double]$s[$s.Count / 2]) / 2.0
}

# One COM call timed, in milliseconds.
function Measure-Call([scriptblock]$call) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    & $call | Out-Null
    $sw.Stop()
    return $sw.Elapsed.TotalMilliseconds
}

function Get-Number([object]$v) { return [double]$v }

# The previous frame's values copied into the sheet Excel keeps them in, by
# the first way that works, each held to its source by Excel itself before it
# is kept: SUMPRODUCT over the two ranges finds no cell differing. Moving
# Excel's array of values through PowerShell failed on the owner's machine,
# as it came and from 0 ("Specified cast is not valid.", 2026-10-10 and
# 2026-10-11), so first Excel is handed the range itself and reads its
# values; then Excel copies the values and pastes them, which goes through
# the clipboard; then the arrays. The way kept, and the ways refused, are
# named on the fixture's rows.
$script:copyWay = ''
$script:copyRefused = New-Object System.Collections.Generic.List[string]
function Copy-Frame($from, $to, [string]$area) {
    $ways = if ($script:copyWay -ne '') { @($script:copyWay) } else { @('the range itself', 'paste values', 'the array as it came', 'the array from 0') }
    foreach ($way in $ways) {
        try {
            switch ($way) {
                'the range itself' { $to.Range($area).Value2 = $from.Range($area) }
                'paste values' {
                    [void]$from.Range($area).Copy()
                    [void]$to.Range($area).PasteSpecial(-4163)
                    # The copied range unmarked as VBA's CutCopyMode = False
                    # unmarks it, through IDispatch: the interop types the
                    # property as XlCutCopyMode, to which PowerShell converts
                    # no False (the owner's machine, 2026-10-11). A range left
                    # marked costs nothing a calculation reads.
                    try { [void][System.__ComObject].InvokeMember('CutCopyMode', [System.Reflection.BindingFlags]::SetProperty, $null, $script:xl, @($false)) } catch { }
                }
                'the array as it came' { $to.Range($area).Value2 = $from.Range($area).Value2 }
                'the array from 0' {
                    $v = $from.Range($area).Value2
                    $rows = $v.GetLength(0); $cols = $v.GetLength(1)
                    $r0 = $v.GetLowerBound(0); $c0 = $v.GetLowerBound(1)
                    $z = New-Object 'object[,]' $rows, $cols
                    for ($i = 0; $i -lt $rows; $i++) {
                        for ($j = 0; $j -lt $cols; $j++) { $z[$i, $j] = $v[($i + $r0), ($j + $c0)] }
                    }
                    $to.Range($area).Value2 = $z
                }
            }
            if ($script:copyWay -eq '') {
                $source = "'" + ($from.Name -replace "'", "''") + "'!" + $area
                $check = $to.Evaluate("SUMPRODUCT(--($area<>$source))")
                # An error Excel answers comes back as its code, a large negative number.
                if ($check -is [int] -and $check -lt 0) { throw "Excel's check of the copy answered the error code $check" }
                $differ = [double]$check
                if ($differ -ne 0) { throw "$differ cells differ from the source" }
                $script:copyWay = $way
            }
            return
        } catch {
            if ($script:copyWay -ne '') { throw }
            $script:copyRefused.Add(("{0}: {1}" -f $way, ((($_.Exception.Message -replace '[\r\n]+', ' ') -replace "'", '') -replace '\|', '/')))
        }
    }
    throw ('no way of copying a frame worked: ' + ($script:copyRefused -join ' | '))
}

# --- the stand-in for Excel (-Fake): the members this script calls, computing nothing ---
function New-FakeRange([hashtable]$cells, [string]$key) {
    $r = New-Object PSObject
    $r | Add-Member -MemberType NoteProperty -Name Key -Value $key
    $r | Add-Member -MemberType NoteProperty -Name Cells -Value $cells
    # A range of many cells answers an array from 1, as Excel does; it refuses
    # an array handed back, of either shape, as the owner's machine did, and a
    # range handed over, so that a run walks Copy-Frame down to its paste.
    $r | Add-Member -MemberType ScriptProperty -Name Value2 -Value {
        if ($this.Cells.ContainsKey($this.Key)) { return $this.Cells[$this.Key] }
        if ($this.Key.Contains(':')) { $a = [Array]::CreateInstance([object], [int[]](2, 2), [int[]](1, 1)); foreach ($i in 1, 2) { foreach ($j in 1, 2) { $a[$i, $j] = 1.0 } }; return ,$a }
        return 999
    } -SecondValue {
        param($v)
        if ($v -is [Array] -or ($v -is [PSObject] -and $null -ne $v.PSObject.Properties['Key'])) { throw (New-Object System.InvalidCastException 'Specified cast is not valid.') }
        $this.Cells[$this.Key] = $v
    }
    $r | Add-Member -MemberType ScriptMethod -Name Copy -Value { return $true }
    $r | Add-Member -MemberType ScriptMethod -Name PasteSpecial -Value { param($how) return $true }
    return $r
}
function New-FakeExcel {
    $cells = @{}
    $app = New-Object PSObject
    $app | Add-Member -MemberType NoteProperty -Name Hwnd -Value 0
    $app | Add-Member -MemberType NoteProperty -Name Version -Value '16.0 (stand-in)'
    $app | Add-Member -MemberType NoteProperty -Name Build -Value 0
    foreach ($p in @('Visible', 'DisplayAlerts', 'ScreenUpdating', 'EnableEvents', 'AskToUpdateLinks', 'CalculateBeforeSave', 'Calculation')) { $app | Add-Member -MemberType NoteProperty -Name $p -Value $null }
    # CutCopyMode refuses a value set through PowerShell, as the interop's
    # XlCutCopyMode refused False on the owner's machine (2026-10-11).
    $app | Add-Member -MemberType ScriptProperty -Name CutCopyMode -Value { return 0 } -SecondValue { param($v) throw (New-Object System.Management.Automation.PSInvalidCastException ('Cannot convert value "{0}" to type "Microsoft.Office.Interop.Excel.XlCutCopyMode".' -f $v)) }
    $mtc = New-Object PSObject -Property @{ Enabled = $true; ThreadMode = 0; ThreadCount = 4 }
    $app | Add-Member -MemberType NoteProperty -Name MultiThreadedCalculation -Value $mtc
    $wf = New-Object PSObject
    $wf | Add-Member -MemberType ScriptMethod -Name Sum -Value { param($range) return 12345 }
    $app | Add-Member -MemberType NoteProperty -Name WorksheetFunction -Value $wf
    foreach ($m in @('Calculate', 'CalculateFull', 'CalculateFullRebuild', 'Quit')) { $app | Add-Member -MemberType ScriptMethod -Name $m -Value { } }
    $books = New-Object PSObject
    $books | Add-Member -MemberType NoteProperty -Name Cells -Value $cells
    $books | Add-Member -MemberType ScriptMethod -Name Open -Value {
        param($path, $links, $readOnly)
        if (-not (Test-Path -LiteralPath $path)) { throw "the stand-in cannot open $path" }
        if ($script:FakeRefuse -ne '' -and [System.IO.Path]::GetFileNameWithoutExtension($path) -eq $script:FakeRefuse) { throw "Unable to get the Open property of the Workbooks class (the stand-in, refusing $script:FakeRefuse as asked)" }
        $c = $this.Cells
        $wb = New-Object PSObject
        $wb | Add-Member -MemberType NoteProperty -Name Saved -Value $true
        $wb | Add-Member -MemberType ScriptMethod -Name Close -Value { param($save) }
        $sheets = New-Object PSObject
        $sheets | Add-Member -MemberType NoteProperty -Name Cells -Value $c
        $sheets | Add-Member -MemberType ScriptMethod -Name Item -Value {
            param($name)
            $cc = $this.Cells
            $ws = New-Object PSObject
            $ws | Add-Member -MemberType NoteProperty -Name Name -Value $name
            $ws | Add-Member -MemberType NoteProperty -Name Cells -Value $cc
            $ws | Add-Member -MemberType ScriptMethod -Name Range -Value { param($a) return (New-FakeRange $this.Cells ($this.Name + '!' + $a)) }
            $ws | Add-Member -MemberType ScriptMethod -Name Evaluate -Value { param($formula) return 0 }
            return $ws
        }
        $wb | Add-Member -MemberType NoteProperty -Name Worksheets -Value $sheets
        return $wb
    }
    $app | Add-Member -MemberType NoteProperty -Name Workbooks -Value $books
    return $app
}

function New-Excel {
    if ($Fake) { return (New-FakeExcel) }
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false
    $xl.EnableEvents = $false
    $xl.AskToUpdateLinks = $false
    return $xl
}

# The Excel's process, opted out of power throttling; the process or $null.
function Get-ExcelProcess($xl) {
    if ($Fake -or -not $xl.Hwnd) { return $null }
    try {
        $id = [RecalcExcelProcess]::ProcessOf([IntPtr][long]$xl.Hwnd)
        if ($id -eq 0) { return $null }
        return [System.Diagnostics.Process]::GetProcessById([int]$id)
    } catch { return $null }
}

function Close-Excel($xl) {
    if ($null -eq $xl) { return }
    try { $xl.Quit() } catch { }
    if (-not $Fake) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

Write-Output '=== THE RECALCULATION FLOORS, EXCEL''S HALF (KERNEL.20) ==='
if ($Fake) { Write-Output '  -Fake: a stand-in for Excel that computes nothing; the numbers below mean nothing' }
$files = @{}
foreach ($f in $fixtures + @('calcpr_built', 'calcpr_calcid', 'calcpr_none')) {
    $p = Join-Path $bench "$f.xlsx"
    if (-not (Test-Path -LiteralPath $p)) { Write-Output "FAIL: $p is not built; run tools\bench_recalc.ps1 first"; exit 1 }
    $files[$f] = (Resolve-Path -LiteralPath $p).Path
}
$benches = @{}
foreach ($f in $fixtures) { $benches[$f] = Read-Bench $f }
$mask = Get-PerformanceMask

# --- the flag: each file in an Excel of its own, read before any calculation ---
$flagRows = New-Object System.Collections.Generic.List[string]
foreach ($f in @('calcpr_built', 'calcpr_calcid', 'calcpr_none')) {
    Write-Output ("  the flag: {0}.xlsx in an Excel of its own..." -f $f)
    $xl = $null
    try {
        $xl = New-Excel
        $wb = $xl.Workbooks.Open($files[$f], 0, $true)
        $c2 = Get-Number ($wb.Worksheets.Item('Output').Range('C2').Value2)
        $saved = [bool]$wb.Saved
        $wb.Close($false)
        $verdict = if ($c2 -eq 999) { 'kept the folded value' } elseif ($c2 -eq 15) { 'computed on opening' } else { 'neither' }
        $flagRows.Add(("    @{{ Half = 'excel-flag'; File = '{0}'; C2 = {1}; Saved = '{2}'; Verdict = '{3}' }}" -f $f, $c2, $saved, $verdict))
    } finally { Close-Excel $xl }
}

# --- the fixtures, in one Excel ---
$rows = New-Object System.Collections.Generic.List[string]
$xl = $null
$restore = $null
$optedOut = $false
$version = ''
$threadsAll = 0
try {
    $xl = New-Excel
    $proc = Get-ExcelProcess $xl
    if ($proc) { $optedOut = [RecalcExcelProcess]::OptOut($proc.Handle); $ownMask = [long]$proc.ProcessorAffinity }
    $version = ('{0} build {1}' -f $xl.Version, $xl.Build)
    $mtc = $xl.MultiThreadedCalculation
    $restore = @{ Enabled = [bool]$mtc.Enabled; ThreadMode = [int]$mtc.ThreadMode; ThreadCount = [int]$mtc.ThreadCount }
    foreach ($name in $fixtures) {
        $b = $benches[$name]
        $sha = (Get-FileHash -Algorithm SHA256 -LiteralPath $files[$name]).Hash
        $wb = $null
        $done = @{}
        $step = 'opening it'
        $script:copyWay = ''
        $script:copyRefused.Clear()
        # A fixture Excel cannot open or compute is a row that says so, and the
        # rest are still measured: one refusal lost three fixtures' rows once
        # (2026-10-10, Life's first run).
        try {
            Write-Output ("  {0}: opening..." -f $name)
            $script:wb = $null
            $open = Measure-Call { $script:wb = $xl.Workbooks.Open($files[$name], 0, $true) }
            $wb = $script:wb
            $step = 'setting calculation to manual'
            $xl.Calculation = $xlCalculationManual
            $xl.CalculateBeforeSave = $false
            $step = 'reading its sum as built'
            $sumSheet = $wb.Worksheets.Item($b.SumSheet)
            $sumBuilt = Get-Number ($xl.WorksheetFunction.Sum($sumSheet.Range($b.SumArea)))
            foreach ($mode in @('all', 'one')) {
                $step = "setting $mode thread(s)"
                if ($mode -eq 'all') {
                    $mtc.Enabled = $true; $mtc.ThreadMode = 0
                    if ($proc -and $ownMask) { $proc.ProcessorAffinity = [IntPtr]$ownMask }
                } else {
                    $mtc.Enabled = $false
                    if ($proc -and $mask -ne 0) { $proc.ProcessorAffinity = [IntPtr]$mask }
                }
                $threads = if ($mode -eq 'all') { [int]$mtc.ThreadCount } else { 1 }
                if ($mode -eq 'all') { $threadsAll = $threads }
                Write-Output ("  {0}: {1} thread(s)..." -f $name, $threads)
                $changes = New-Object System.Collections.Generic.List[double]
                $copies = New-Object System.Collections.Generic.List[double]
                $ws = $wb.Worksheets.Item($b.Sheet)
                if ($b.Kind -eq 'edit') {
                    for ($k = 0; $k -lt $edits; $k++) {
                        $step = "edit $($k + 1), $mode thread(s)"
                        $v = if ($k % 2 -eq 0) { [double]$b.Other } else { [double]$b.Built }
                        $ws.Range($b.Cell).Value2 = $v
                        $changes.Add((Measure-Call { $xl.Calculate() }))
                    }
                } else {
                    $step = "the first frame, $mode thread(s)"
                    $screen = $wb.Worksheets.Item($b.SumSheet)
                    $last = $wb.Worksheets.Item($b.Twin)
                    $ws.Range($b.Cell).Value2 = 1.0
                    [void](Measure-Call { $xl.Calculate() })
                    for ($f = 2; $f -le $b.Count; $f++) {
                        $step = "copying frame $($f - 1) into $($b.Twin), $mode thread(s)"
                        $copies.Add((Measure-Call { Copy-Frame $screen $last $b.SumArea }))
                        $step = "advancing the clock to frame $f, $mode thread(s)"
                        $ws.Range($b.Cell).Value2 = [double]$f
                        $step = "calculating frame $f, $mode thread(s)"
                        $changes.Add((Measure-Call { $xl.Calculate() }))
                    }
                }
                $step = "reading its sum edited, $mode thread(s)"
                $sumEdited = Get-Number ($xl.WorksheetFunction.Sum($sumSheet.Range($b.SumArea)))
                $step = "CalculateFull, $mode thread(s)"
                $whole = New-Object System.Collections.Generic.List[double]
                for ($k = 0; $k -lt $wholeRuns; $k++) { $whole.Add((Measure-Call { $xl.CalculateFull() })) }
                $step = "CalculateFullRebuild, $mode thread(s)"
                $rebuild = New-Object System.Collections.Generic.List[double]
                for ($k = 0; $k -lt $rebuildRuns; $k++) { $rebuild.Add((Measure-Call { $xl.CalculateFullRebuild() })) }
                $step = "a Calculate with nothing to compute, $mode thread(s)"
                $floor = New-Object System.Collections.Generic.List[double]
                for ($k = 0; $k -lt $floorRuns; $k++) { $floor.Add((Measure-Call { $xl.Calculate() })) }
                if ($b.Kind -eq 'edit') {
                    # back to the built value, so the second pass starts where the first did
                    $ws.Range($b.Cell).Value2 = [double]$b.Built
                    [void](Measure-Call { $xl.Calculate() })
                }
                $step = "the row, $mode thread(s)"
                $copy = if ($copies.Count -gt 0) { Get-Median $copies.ToArray() } else { 0.0 }
                $copyField = ''
                if ($b.Kind -eq 'frames') {
                    $copyField = "; CopyWay = '$($script:copyWay)'"
                    if ($script:copyRefused.Count -gt 0) { $copyField += "; CopyRefused = '$($script:copyRefused -join ' | ')'" }
                }
                $rows.Add(("    @{{ Half = 'excel'; Fixture = '{0}'; Threads = '{1}'; ThreadCount = {2}; Sha256 = '{3}'; Open = {4:F1}; Whole = {5:F3}; Rebuild = {6:F3}; Edit = {7:F3}; Copy = {8:F1}; ComFloor = {9:F3}; SumBuilt = {10}; SumEdited = {11}{12} }}" -f $name, $mode, $threads, $sha, $open, (Get-Median $whole.ToArray()), (Get-Median $rebuild.ToArray()), (Get-Median $changes.ToArray()), $copy, (Get-Median $floor.ToArray()), $sumBuilt, $sumEdited, $copyField))
                $done[$mode] = $true
            }
            $step = 'closing it'
            $wb.Close($false)
            $wb = $null
        } catch {
            $why = ("$step`: " + (($_.Exception.Message -replace '[\r\n]+', ' ') -replace "'", '')).Trim()
            Write-Output ("  {0}: FAILED, {1}" -f $name, $why)
            foreach ($mode in @('all', 'one')) {
                if (-not $done.ContainsKey($mode)) { $rows.Add(("    @{{ Half = 'excel'; Fixture = '{0}'; Threads = '{1}'; Sha256 = '{2}'; Failed = '{3}' }}" -f $name, $mode, $sha, $why)) }
            }
            if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
        }
    }
} finally {
    if ($null -ne $xl -and $null -ne $restore) {
        try { $xl.MultiThreadedCalculation.Enabled = $restore.Enabled; $xl.MultiThreadedCalculation.ThreadMode = $restore.ThreadMode } catch { }
    }
    Close-Excel $xl
}

$cpu = ''
try { $cpu = (Get-CimInstance Win32_Processor | Select-Object -First 1).Name.Trim() } catch { $cpu = 'unknown processor' }
$os = ''
try { $os = (Get-CimInstance Win32_OperatingSystem).Caption + ' ' + [System.Environment]::OSVersion.Version.ToString() } catch { $os = 'unknown system' }
$fakeWord = if ($Fake) { ' | THE STAND-IN, NOT EXCEL' } else { '' }
$maskWord = if ($mask -ne 0) { ('one thread on the performance cores, mask 0x{0:X}' -f $mask) } else { 'one kind of core' }
$environment = ('{0} | {1} | Excel {2} | all threads {3}; {4} | power throttling opted out: {5}{6}' -f $cpu, $os, $version, $threadsAll, $maskWord, $(if ($optedOut) { 'yes' } else { 'no' }), $fakeWord)
$block = New-Object System.Collections.Generic.List[string]
$block.Add(('# Excel''s half, {0}: {1}' -f (Get-Date -Format 'yyyy-MM-dd'), $environment))
foreach ($r in $flagRows) { $block.Add($r) }
foreach ($r in $rows) { $block.Add($r) }
$text = ($block -join "`r`n") + "`r`n"
try { [System.IO.File]::WriteAllText((Join-Path $repoRoot 'target/bench_recalc_excel.txt'), $text, (New-Object System.Text.UTF8Encoding($false))) } catch { }
Write-Output ''
Write-Output 'Paste this block, and the native half''s, into the reply (also in target\bench_recalc_excel.txt):'
Write-Output $text
exit 0
