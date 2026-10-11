<#
bench_recalc.ps1 - the recalculation floors' native half: the machine timed
on the four fixtures of scripts/recalc/bench/ on the owner's machine, and
printed as one block to paste into tools/check_recalc_floors.ps1, beside
the block tools/bench_recalc_excel.ps1 prints for Excel.

WHY: KERNEL.20's sixth slice (docs/BETA_REARVIEW.md, its scoping). The
item's oracle is check_recalc_floors.ps1, "the timings as floors that never
go down", against Excel on the fixtures. The machine is the engine
KERNEL.20's slices made faster; this script runs core/examples/recalc.rs,
which builds each fixture through the core's own surface as frazaro build
builds it, writes it to target/bench/ for Excel's half to open, loads it
into the machine under Excel's library, holds every value before it prints
a number, and times a whole frame and the frame after the fixture's edit.

THE TRAPS IT IS BUILT AROUND, both measured for Alonzo's CART.1 on this
machine and its instrument tools/bench_frames.ps1, whose discipline this is:
Windows throttles a process it deems in the background (EcoQoS) about three
seconds into a run, by about 40%, so the runner's process is opted out of
power throttling as it starts; and a process may run on an efficiency core,
about a quarter slower with a clock that looks steady, so the runner is
started on the performance cores, the logical processors of the highest
efficiency class, by this script's own affinity set before the start and
put back after, and at high priority, as KERNEL.25's A/B rig ran, since
another process on the same cores moved a frame by a quarter (2026-10-10,
a peer session's headless Chrome beside the first runs). The runner also watches the clock itself, a fixed
computation timed beside every frame, and prints no rows when it moved
more than 10%.

WHAT IT DOES:
  1. builds the runner, unless -NoBuild: cargo build --release -p
     frazaro-core --example recalc;
  2. runs it, opted out and on the performance cores, and reads its four
     rows;
  3. prints the block: the date, the environment, the four rows, and the
     floor history as the check holds it with this run's entry added last,
     each floor the run's number less the check's margin of 15%, rounded
     down to three figures, to paste over the check's history whole; and
     writes the same block to target/bench_recalc.txt, to copy from.
It is an instrument and not a check: it measures, the owner pastes, and the
check holds what was pasted. No job runs it.

House style (tools/*.ps1): PowerShell 5.1 and pwsh alike; ASCII; no network.

Usage:  powershell -File tools\bench_recalc.ps1
        powershell -File tools\bench_recalc.ps1 -NoBuild
#>
param(
    [switch]$NoBuild
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$onWindows = [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT
$margin = 0.15
$fixtures = @('column', 'chain', 'sums', 'life')

function Find-Cargo {
    $c = Get-Command cargo -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    $home2 = if ($env:USERPROFILE) { $env:USERPROFILE } else { $env:HOME }
    foreach ($p in @((Join-Path $home2 '.cargo/bin/cargo.exe'), (Join-Path $home2 '.cargo/bin/cargo'))) { if (Test-Path -LiteralPath $p) { return $p } }
    return ''
}

if ($onWindows) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class RecalcPowerThrottling {
    [StructLayout(LayoutKind.Sequential)]
    public struct State { public uint Version; public uint ControlMask; public uint StateMask; }
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool SetProcessInformation(IntPtr process, int infoClass, ref State info, uint size);
    // ProcessPowerThrottling is class 4; the execution-speed bit is 1; a state of 0 opts the process out.
    public static bool OptOut(IntPtr process) {
        var s = new State(); s.Version = 1; s.ControlMask = 1; s.StateMask = 0;
        return SetProcessInformation(process, 4, ref s, (uint)Marshal.SizeOf(typeof(State)));
    }
}
public static class RecalcCores {
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool GetSystemCpuSetInformation(IntPtr information, uint length, out uint returned, IntPtr process, uint flags);
    // Group 0's logical processors as pairs, its index then its efficiency
    // class, a higher class a faster core (SYSTEM_CPU_SET_INFORMATION: Size,
    // Type at 4, Group at 12, LogicalProcessorIndex at 14, EfficiencyClass at 18).
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
}

# The performance cores: the logical processors of the highest efficiency
# class as an affinity mask; 0 where the machine has one kind of core.
function Get-PerformanceCores {
    $set = @{ Mask = [long]0; Fast = 0; All = 0; Class = 0 }
    if (-not $onWindows) { return $set }
    try { $pairs = @([RecalcCores]::Classes()) } catch { return $set }
    if ($pairs.Count -lt 2) { return $set }
    $classes = @(for ($i = 1; $i -lt $pairs.Count; $i += 2) { $pairs[$i] })
    $set.All = $classes.Count
    $set.Class = ($classes | Measure-Object -Maximum).Maximum
    for ($i = 0; $i -lt $pairs.Count; $i += 2) {
        if ($pairs[$i + 1] -eq $set.Class -and $pairs[$i] -lt 64) { $set.Mask = $set.Mask -bor ([long]1 -shl $pairs[$i]); $set.Fast++ }
    }
    if ($set.Fast -eq $set.All) { $set.Mask = [long]0 }
    return $set
}

# A number less the margin, rounded down to three significant figures.
function Get-Floor([double]$x) {
    $y = $x * (1.0 - $margin)
    if ($y -le 0) { return [long]0 }
    $p = [math]::Pow(10, [math]::Floor([math]::Log10($y)) - 2)
    return [long]([math]::Floor($y / $p) * $p)
}

# The floor history's entries as tools/check_recalc_floors.ps1 holds them, one a line.
function Get-HistoryEntries([string]$path) {
    $entries = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $path)) { return $entries.ToArray() }
    $inside = $false
    foreach ($line in [System.IO.File]::ReadAllLines($path)) {
        if (-not $inside) { if ($line -match '^\$floorHistory = @\(\s*$') { $inside = $true }; continue }
        if ($line -match '^\)\s*$') { break }
        $t = $line.Trim().TrimEnd(',').Trim()
        if ($t -ne '' -and -not $t.StartsWith('#')) { $entries.Add($t) }
    }
    return $entries.ToArray()
}

function Get-Field([string]$row, [string]$name) { $m = [regex]::Match($row, "\b$name = '?([^;']+)'?"); if ($m.Success) { return $m.Groups[1].Value.Trim() }; return '' }

$runnerExe = Join-Path $repoRoot ('target/release/examples/recalc' + $(if ($onWindows) { '.exe' } else { '' }))
Write-Output '=== THE RECALCULATION FLOORS, THE NATIVE HALF: THE MACHINE ON THE FIXTURES (KERNEL.20) ==='
if (-not $NoBuild) {
    $cargo = Find-Cargo
    if ($cargo -eq '') { Write-Output 'FAIL: no cargo on PATH or in ~/.cargo/bin; build the runner, then run with -NoBuild'; exit 1 }
    Write-Output '  building the runner...'
    Push-Location $repoRoot
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $out = & $cargo build --release -p frazaro-core --example recalc 2>&1 | Out-String } finally { $ErrorActionPreference = $old; Pop-Location }
    if ($LASTEXITCODE -ne 0) { Write-Output $out; Write-Output 'FAIL: the runner did not build'; exit 1 }
}
if (-not (Test-Path -LiteralPath $runnerExe)) { Write-Output "FAIL: $runnerExe is not built; run without -NoBuild"; exit 1 }

$cores = Get-PerformanceCores
$self = [System.Diagnostics.Process]::GetCurrentProcess()
$ownAffinity = $self.ProcessorAffinity
$coresWord = if ($cores.Mask -ne 0) { ', on the performance cores' } else { '' }
Write-Output ("  the runner, its process opted out of power throttling{0}..." -f $coresWord)
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $runnerExe
$psi.UseShellExecute = $false
$psi.RedirectStandardOutput = $true
$psi.WorkingDirectory = $repoRoot
if ($cores.Mask -ne 0) { $self.ProcessorAffinity = [IntPtr]$cores.Mask }
try { $proc = [System.Diagnostics.Process]::Start($psi) } finally { $self.ProcessorAffinity = $ownAffinity }
$optedOut = $true
if ($onWindows) { try { $optedOut = [RecalcPowerThrottling]::OptOut($proc.Handle) } catch { $optedOut = $false } }
$high = $true
try { $proc.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::High } catch { $high = $false }
$pinned = $true
if ($cores.Mask -ne 0) { try { $pinned = ([long]$proc.ProcessorAffinity -eq $cores.Mask) } catch { $pinned = $false } }
$text = $proc.StandardOutput.ReadToEnd()
$proc.WaitForExit()
$rows = @(($text -split "`r?`n") | Where-Object { $_ -match "^\s+@\{ Half = 'native'" } | ForEach-Object { $_.Trim() })
if ($rows.Count -ne $fixtures.Count) {
    Write-Output $text
    Write-Output ("FAIL: the runner printed {0} row(s), not {1} (exit {2}); its words are above" -f $rows.Count, $fixtures.Count, $proc.ExitCode)
    exit 1
}

# The environment line, as the check reads it.
$cpu = 'unknown processor'
try { $cpu = (Get-CimInstance Win32_Processor | Select-Object -First 1).Name.Trim() } catch { }
$logical = [System.Environment]::ProcessorCount
$os = 'unknown system'
try { $os = (Get-CimInstance Win32_OperatingSystem).Caption.Trim() + ' ' + [System.Environment]::OSVersion.Version.ToString() } catch { }
$rustc = ''
try { $rc = Get-Command rustc -ErrorAction SilentlyContinue; if (-not $rc) { $rc = Get-Item (Join-Path $env:USERPROFILE '.cargo/bin/rustc.exe') -ErrorAction SilentlyContinue }; if ($rc) { $rustc = (& $(if ($rc.Source) { $rc.Source } else { $rc.FullName }) -V) } } catch { }
$version = ''
$m = [regex]::Match([System.IO.File]::ReadAllText((Join-Path $repoRoot 'Cargo.toml')), '(?ms)^\[workspace\.package\].*?^version\s*=\s*"([^"]+)"')
if ($m.Success) { $version = $m.Groups[1].Value }
$coreLine = if ($cores.Mask -ne 0) { ('on the performance cores, {0} of {1} logical processors (efficiency class {2}): the runner {3}' -f $cores.Fast, $cores.All, $cores.Class, $(if ($pinned) { 'yes' } else { 'no' })) } else { 'one kind of core' }
$environment = ('{0}, {1} logical processors | {2} | {3} | frazaro-core {4} | power throttling opted out: the runner {5} | high priority: the runner {6} | {7}' -f $cpu, $logical, $os, $rustc, $version, $(if ($optedOut) { 'yes' } else { 'no' }), $(if ($high) { 'yes' } else { 'no' }), $coreLine)

# The floor history with this run's entry added last.
$today = Get-Date -Format 'yyyy-MM-dd'
$floors = New-Object System.Collections.Generic.List[string]
foreach ($r in $rows) {
    $f = Get-Field $r 'Fixture'
    $floors.Add(("'{0} whole' = {1}" -f $f, (Get-Floor ([double](Get-Field $r 'WholePerSecond')))))
    $floors.Add(("'{0} edit' = {1}" -f $f, (Get-Floor ([double](Get-Field $r 'EditsPerSecond')))))
}
$entry = ("@{{ Date = '{0}'; Reason = ''; Floors = @{{ {1} }} }}" -f $today, ($floors -join '; '))
$history = @(Get-HistoryEntries (Join-Path $repoRoot 'tools/check_recalc_floors.ps1')) + @($entry)

$block = New-Object System.Collections.Generic.List[string]
$block.Add(("`$baselineDate = '{0}'" -f $today))
$block.Add(("`$baselineNative = '{0}'" -f ($environment -replace "'", "''")))
$block.Add('    # the native rows:')
foreach ($r in $rows) { $block.Add('    ' + $r) }
$block.Add('# the floor history, with this run''s entry last:')
$block.Add('$floorHistory = @(')
for ($i = 0; $i -lt $history.Count; $i++) { $block.Add('    ' + $history[$i] + $(if ($i -lt $history.Count - 1) { ',' } else { '' })) }
$block.Add(')')
$out = ($block -join "`r`n") + "`r`n"
try { [System.IO.File]::WriteAllText((Join-Path $repoRoot 'target/bench_recalc.txt'), $out, (New-Object System.Text.UTF8Encoding($false))) } catch { }
Write-Output ''
Write-Output ($text.TrimEnd())
Write-Output ''
Write-Output 'Paste this block, and Excel''s half''s, into the reply (also in target\bench_recalc.txt):'
Write-Output $out
exit 0
