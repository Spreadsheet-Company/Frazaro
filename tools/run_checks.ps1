<#
run_checks.ps1 - runs every tools\check_*.ps1 and reports one total.

WHY: there are 30 of them now. Running them one at a time is a page of
scrollback in which a single red line is easy to walk past, and "they all
seemed to pass" is not a result anybody can quote in a commit message. This
prints one line per check and one total, and its own exit code is the answer.

IT IS NOT NAMED check_*.ps1, deliberately: this file would otherwise match
its own glob and run itself forever. It also skips itself by name as a belt
to that brace, in case it is ever copied under a different name.

THE FLOOR IS A CHECK OF ITS OWN. A check that is deleted, renamed out of the
glob, or moved to another folder does not fail - it simply stops running, and
nothing notices. So the count discovered is compared against a hardcoded
floor below, in the same hand-maintained, reviewable shape every other scan
in this folder uses. Lowering it is a deliberate, reviewed act.

WHAT IT DOES NOT DO: it is not wired into VlaSelfTest, the same reasoning
every tools\check_*.ps1 gives - a step a human runs, not a per-run gate. Nor
does it run the other verifiers in this folder that are not named check_*
(optimize01_corpus.ps1, datalog14_proof.ps1 and friends): those take
arguments, have their own modes, and pretending they are interchangeable
with a plain pass/fail scan would hide what each one actually asserts.
-WithExtras runs the hardcoded list of them below, each with the arguments
it needs, and says so separately.

SIDE BY SIDE (AS.10, 2026-10-21): the checks run at once, as many as the
machine has logical processors (-Throttle to choose), each still in a
powershell.exe of its own, so a check sees what it saw when they ran one
after another. The report is the same: a line a check, in the order of
their names, each printed once it and every check before it are done, then
the total, which gives how many ran at once and the wall time. One after
another they took 133 s on the owner's machine that day, side by side the
longest check. -Throttle 1 runs them one after another, as before. -Control
proves the gathering: fake checks, one slow and first by name, one failing,
one writing to stderr, run three at a time against a folder of their own,
must report in the order of their names, the failure with its tail and the
exit code 1, and a floor not met must fail.

Usage:
  powershell -File tools\run_checks.ps1
  powershell -File tools\run_checks.ps1 -Filter prolog      # only matching names
  powershell -File tools\run_checks.ps1 -ShowAll            # print passing output too
  powershell -File tools\run_checks.ps1 -WithExtras         # also the named verifiers
  powershell -File tools\run_checks.ps1 -Throttle 1         # one at a time
  powershell -File tools\run_checks.ps1 -Control            # the gathering proved

Exit code: 0 if every check passed AND at least the floor were found;
1 otherwise.
#>

param(
    [string]$Filter = '',
    [switch]$ShowAll,
    [switch]$WithExtras,
    [int]$Floor = 0,
    # A folder to scan instead of this script's own. Its only purpose is to
    # make THIS script testable: a runner that cannot be pointed at a broken
    # tree cannot be shown to go red, and an unverified runner reporting
    # "ALL GREEN" is worse than no runner.
    [string]$ToolsDir = '',
    # How many scripts run at once (AS.10): every logical processor unless
    # given; 1 runs them one after another.
    [int]$Throttle = 0,
    [switch]$Control
)

$ErrorActionPreference = 'Stop'

$toolsDir = if ($ToolsDir -ne '') { $ToolsDir } else { $PSScriptRoot }
$selfName = Split-Path -Leaf $PSCommandPath

# --- The floor, hand-maintained. ---
# 2026-09-20, DATALOG.14: 27, the 26 that stood at 0.6.2 plus
# check_datalog_per_tuple_alloc.ps1.
# 2026-09-24, OPTIMIZE.3 slice 2: 28, plus
# check_optimize_search_discipline.ps1.
# 2026-09-27, METAPROOF.1: 30 - the floor had fallen one behind the 29
# already here, and check_proofs.ps1 (the proof corpus, read without
# Excel) makes thirty.
# 2026-09-30, U.30: 32 - U.29's check_run_gives_back.ps1 made thirty-one
# without raising this, and check_engine_call_names.ps1 makes thirty-two.
# 2026-10-01, PORT.4: 34 - check_version_twin.ps1 (one corpus, one version)
# and check_core_imports.ps1 (the wasm core imports nothing).
# 2026-10-01, PORT.5: 36 - check_data_exports.ps1 (the head table and the
# message catalogue, exported to data, agree with the VBA they came from)
# and check_compile_prefix.ps1 (the core reproduces the compile golden; the
# matched prefix never goes down).
# 2026-10-02, PORT.6: 37 - check_prove_floors.ps1 (the core's passing
# phrasebook proofs per file never go down).
# 2026-10-02, PORT.6 slice 6e: 38 - check_translate_prefix.ps1 (the core
# reproduces the translate golden; the matched prefix never goes down).
# 2026-10-02, PORT.6 slice 6f: 39 - check_refusal_golden.ps1 (the refusal
# golden agrees with its fixture, every case refuses, and the distinct
# refusal ids it reaches never go down).
# 2026-10-02, PORT.6 slice 6h: 40 - check_web_offline.ps1 (the web page
# loads nothing, links nowhere, and is filled from exactly three
# placeholders).
# 2026-10-03, PORT.7 (slice 7a): 41 - check_build_golden.ps1 (the writer's
# golden reproduced byte for byte, its length a floor, no clock in the zip).
# 2026-10-03, PORT.8 (slice 8a): 42 - check_reflect_golden.ps1 (the reader's
# relations reproduced whole for every fixture, each golden's line count a
# floor, the fixed order read off the golden itself).
# 2026-10-04, AXM.7: 43 - check_refers_golden.ps1 (the refers golden agrees
# with its fixture, every case names its cell and ends in its R1C1 record,
# and the distinct reference kinds it reaches never go down).
# 2026-10-05, the crates prepared for crates.io: 46 - check_crate_package.ps1
# (what cargo package would refuse, read without cargo: no build-time
# include leaves its crate, and every manifest names what crates.io shows).
# 2026-10-05, the door's built-in prelude and english.vla: 47 -
# check_crate_data.ps1 (the copies under cli/data/ are the corpus files byte
# for byte, and a built door with no flags translates as it does with them).
# 2026-10-05, SEC.15: 48 - check_egress_golden.ps1 (the egress golden agrees
# with its fixture, every record is REFUSED with a listed name or WRITTEN,
# the VBA's list and the core's are one, and every listed name is reached).
# 2026-10-06, KERNEL.1: 49 - check_kernel_boundary.ps1 (the kernel bakes in
# exactly its four data tables, and no sentence rule or English sentence
# enters core/src/ outside a test module beyond the pinned exceptions).
# 2026-10-07, KERNEL.4: 50 - check_view_golden.ps1 (the view record of each
# fixture window reproduced whole, each golden's line count a floor, the
# fixed order read off the golden itself).
# 2026-10-07, KERNEL.2: 51 - check_distro.ps1 (every distro under distros/
# well formed and naming files that exist, the english distro the doors'
# own, and the three readers of a manifest held to one shape).
# 2026-10-08, KERNEL.7: 52 - check_recalc_golden.ps1 (the recalc goldens
# reproduced whole, every Excel-saved value agreeing, the study's workbooks
# agreeing with their builder).
# 2026-10-09, KERNEL.22: 53 - check_wasm_exports.ps1 (the core's module and the
# language's export exactly their lists, the language's C surface only under
# its feature).
# 2026-10-09, KERNEL.3: 54 - check_study_fixture.ps1 (the study's answer key,
# written 2026-10-08 and committed after KERNEL.22).
# 2026-10-10, VLA.3: 55 - check_expand_golden.ps1 (the expand golden agrees
# with its fixture, the reference's lists of the expand-time primitives are
# the port's, and every primitive is reached).
# 2026-10-10, KERNEL.20's sixth slice: 56 - check_recalc_floors.ps1 (the
# machine's recalculation of four fixtures measured beside Excel's of the
# same files, the floors a history that never goes down, the values held).
# 2026-10-12, KERNEL.30: 57 - check_shapes_golden.ps1 (the door's census of
# oracle 12's fixtures and of the four bench programs is each golden whole,
# its order and its totals holding in the golden itself).
# 2026-10-17, KERNEL.47: 58 - check_formats_golden.ps1 (the formats golden,
# Excel's text of each value under each code, its shape held to its fixture
# and every construct the language draws reached).
$expectedAtLeast = 58
if ($Floor -gt 0) { $expectedAtLeast = $Floor }

# --- The other verifiers, each with the arguments it needs. ---
# Hand-maintained for the same reason: a list that discovers itself would
# quietly stop covering something the day a file was renamed.
$extras = @(
    @{ Script = 'optimize01_corpus.ps1';   Args = @();              What = 'the OPTIMIZE corpus against its keys' },
    @{ Script = 'datalog14_proof.ps1';     Args = @();              What = 'DATALOG.14 specialisation, 15 hand-derived programs' },
    @{ Script = 'datalog14_proof.ps1';     Args = @('-Mutations');  What = 'DATALOG.14 conditions, mutation control' },
    @{ Script = 'datalog14_model.ps1';     Args = @('-Control');    What = 'DATALOG.14 cost model against the measured ladder' },
    @{ Script = 'optimize3_model.ps1';     Args = @('-Control');    What = 'OPTIMIZE.3 grounding model against the pre-flight ladder' },
    @{ Script = 'check_optimize_search_discipline.ps1'; Args = @('-Control'); What = 'OPTIMIZE.3 search discipline, mutation control' },
    @{ Script = 'optimize3_search_twin.ps1'; Args = @('-Control'); What = 'OPTIMIZE.3 search pins against their host-independent twin' },
    @{ Script = 'prove.ps1';               Args = @('-Control');    What = 'PORT.4 conformance runner: the fake implementation passes, the mutant fails' },
    @{ Script = 'check_core_imports.ps1';  Args = @('-Control');    What = 'PORT.4 import reader: an empty module counts 0, a module importing a.b counts 1' },
    @{ Script = 'check_wasm_exports.ps1';  Args = @('-Control');    What = 'KERNEL.22 export reader: a module reads as its names, a list one short or one long is reported' },
    @{ Script = 'check_crate_package.ps1'; Args = @('-Control');    What = 'VLA.1 package reader: the clean crates pass; a test include leaving vla-lang, an integration test reaching out, a core include leaving its crate and a respelled include each fail' },
    @{ Script = 'check_expand_golden.ps1'; Args = @('-Control');    What = 'VLA.3 expand golden: the real set passes; a misspelt tag, a dropped case, an unknown id, a broken reference Select and a renamed primitive each fail' },
    @{ Script = 'check_recalc_floors.ps1'; Args = @('-Control');    What = 'KERNEL.20 recalculation floors: a clean baseline passes; its mutants, pasted histories and fake runners each fail; the runners found, run side by side and kept (AS.10)' },
    @{ Script = 'check_shapes_golden.ps1'; Args = @('-Control');    What = 'KERNEL.30 census goldens: a clean fake door passes, the second pass from its cache; a mutant door and six broken goldens each fail' },
    @{ Script = 'check_formats_golden.ps1'; Args = @('-Control');   What = 'KERNEL.47 formats golden: the real set passes; a misspelt tag, a dropped case, a value echoed otherwise, an exponent and the backslashes taken out each fail' },
    @{ Script = 'run_checks.ps1';          Args = @('-Control');    What = 'AS.10 this runner: fake checks side by side and one at a time report in the order of their names; a failure and a floor not met each fail' }
)

# One script in a powershell.exe of its own, its output merged in the order
# it printed it. The block runs in a runspace of the pool below (AS.10), so
# that scripts run at once while each still runs alone in its own process.
# The runspace's $ErrorActionPreference is 'Continue', not this script's
# 'Stop', and must stay so: a CHECK writing to stderr is this script's
# ordinary business, not an error in it. In Windows PowerShell 5.1, `2>&1` on
# a native executable wraps each stderr line in a NativeCommandError, which
# under 'Stop' TERMINATES the runner - so a check that failed loudly would
# kill the run instead of being reported as FAIL, and the total would never
# print. Found by trying to make this script go red, which is the whole
# reason for trying.
$runOne = {
    param([string]$Path, [string[]]$Arguments)
    $ErrorActionPreference = 'Continue'
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $out = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Path @Arguments 2>&1
    $code = $LASTEXITCODE
    $sw.Stop()
    [pscustomobject]@{ Code = $code; Output = @($out | ForEach-Object { "$_" }); Seconds = $sw.Elapsed.TotalSeconds }
}

# ------------------------------------------------------------------ -Control
# Fake checks in a folder of their own, gathered by this script run again
# over it: the order of the report, a failure's tail, the exit code, the
# floor, a filter, and side by side against one at a time, read off the times
# each fake writes as it starts and ends.
if ($Control) {
    $failures = New-Object System.Collections.Generic.List[string]
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('run_checks_control_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $stamp = { param([string]$who, [string]$body) "Add-Content -LiteralPath (Join-Path `$PSScriptRoot '$who.time') ('start ' + [DateTime]::UtcNow.Ticks)`r`n$body`r`nAdd-Content -LiteralPath (Join-Path `$PSScriptRoot '$who.time') ('end ' + [DateTime]::UtcNow.Ticks)`r`nexit `$code`r`n" }
        $fakes = [ordered]@{
            'check_a_slow'   = "Start-Sleep -Milliseconds 2500`r`nWrite-Output 'a: slow and first by name'`r`n`$code = 0"
            'check_b_fails'  = "Write-Output 'b: one line'`r`nWrite-Output 'b: the reason it fails'`r`n`$code = 1"
            'check_c_stderr' = "[Console]::Error.WriteLine('c: a line on stderr')`r`nWrite-Output 'c: passes'`r`n`$code = 0"
            'check_d_quick'  = "Write-Output 'd: quick'`r`n`$code = 0"
        }
        foreach ($k in $fakes.Keys) { [System.IO.File]::WriteAllText((Join-Path $tmp "$k.ps1"), (& $stamp $k $fakes[$k]), $utf8) }
        $times = {
            $t = @{}
            foreach ($k in $fakes.Keys) {
                $p = Join-Path $tmp "$k.time"
                if (Test-Path -LiteralPath $p) {
                    $l = @(Get-Content -LiteralPath $p)
                    $t[$k] = @{ Start = [long](($l | Where-Object { $_ -like 'start *' } | Select-Object -First 1) -replace '^start ', ''); End = [long](($l | Where-Object { $_ -like 'end *' } | Select-Object -First 1) -replace '^end ', '') }
                    Remove-Item -LiteralPath $p
                }
            }
            $t
        }
        $gather = {
            param([string[]]$More)
            $prev = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            try { $o = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -ToolsDir $tmp @More 2>&1 | ForEach-Object { "$_" }; $c = $LASTEXITCODE } finally { $ErrorActionPreference = $prev }
            @{ Lines = @($o); Code = $c }
        }
        # Three at a time: the report in the order of the names, though the
        # first by name ends last; the failure's tail; stderr kept; exit 1.
        $r = & $gather @('-Throttle', '3', '-Floor', '4', '-ShowAll')
        $t = & $times
        $order = @($r.Lines | Where-Object { $_ -match '^  (ok  |FAIL)  check_' } | ForEach-Object { ($_ -split '\s+')[2] })
        if (($order -join ',') -ne 'check_a_slow.ps1,check_b_fails.ps1,check_c_stderr.ps1,check_d_quick.ps1') { $failures.Add("three at a time: the report's order was $($order -join ', ')") }
        $tags = @($r.Lines | Where-Object { $_ -match '^  (ok  |FAIL)  check_' } | ForEach-Object { $_.Substring(2, 4).Trim() })
        if (($tags -join ',') -ne 'ok,FAIL,ok,ok') { $failures.Add("three at a time: the verdicts were $($tags -join ', ')") }
        if ($r.Code -ne 1) { $failures.Add("three at a time, one failing: exit $($r.Code), not 1") }
        $at = [array]::IndexOf($r.Lines, '--- check_b_fails.ps1 ---')
        if ($at -lt 0 -or $r.Lines[$at + 2] -ne '  b: the reason it fails') { $failures.Add('three at a time: the failure''s tail was not printed under its name') }
        if (-not ($r.Lines -contains '  c: a line on stderr')) { $failures.Add('three at a time: the line on stderr was not kept in its check''s output') }
        if (-not ($r.Lines | Where-Object { $_ -match '^=== 4 checks: 3 passed, 1 failed, 3 at a time in ' })) { $failures.Add('three at a time: no total naming 4 checks, 3 at a time') }
        if ($t.Count -ne 4) { $failures.Add("three at a time: $($t.Count) fakes wrote their times, not 4") }
        elseif (-not ($t['check_d_quick'].Start -lt $t['check_a_slow'].End -and $t['check_a_slow'].Start -lt $t['check_d_quick'].End)) { $failures.Add('three at a time: the quick fake did not run while the slow one ran') }
        else { Write-Output '  control: three at a time: the slow fake ran beside the quick one, the report in the order of the names, the failure''s tail under its name, stderr kept, exit 1' }
        # One at a time: none overlaps the next.
        $r = & $gather @('-Throttle', '1', '-Floor', '4')
        $t = & $times
        $keys = @($fakes.Keys)
        $overlap = $false
        for ($i = 1; $i -lt $keys.Count; $i++) { if ($t.ContainsKey($keys[$i - 1]) -and $t.ContainsKey($keys[$i]) -and $t[$keys[$i]].Start -lt $t[$keys[$i - 1]].End) { $overlap = $true } }
        if ($t.Count -ne 4 -or $overlap) { $failures.Add("one at a time: $($t.Count) fakes wrote their times, and one began before the one before it ended: $overlap") }
        elseif ($r.Code -ne 1) { $failures.Add("one at a time, one failing: exit $($r.Code), not 1") }
        else { Write-Output '  control: one at a time: each fake began after the one before it ended, exit 1' }
        # A floor not met.
        $r = & $gather @('-Floor', '5')
        if ($r.Code -ne 1 -or -not ($r.Lines | Where-Object { $_ -like 'FLOOR: 4 check(s) found, but 5 were expected*' })) { $failures.Add("a floor of 5 over 4 checks: exit $($r.Code), and no FLOOR line") }
        else { Write-Output '  control: a floor of 5 over 4 checks fails by name' }
        $null = & $times
        # A filter.
        $r = & $gather @('-Filter', 'quick')
        if ($r.Code -ne 0 -or -not ($r.Lines | Where-Object { $_ -like "=== 1 check(s) matching 'quick': 1 passed, 0 failed*" })) { $failures.Add("a filter: exit $($r.Code), and no total of the one matching check") }
        else { Write-Output '  control: a filter runs the one check it matches, exit 0' }
    } finally { Remove-Item -Recurse -Force -LiteralPath $tmp -ErrorAction SilentlyContinue }
    if ($failures.Count -gt 0) { foreach ($f in $failures) { Write-Output "FAIL: $f" }; exit 1 }
    Write-Output 'OK: run_checks -Control: side by side and one at a time, the report in the order of the names, a failure and a floor each fail'
    exit 0
}

$scripts = @(Get-ChildItem -Path (Join-Path $toolsDir 'check_*.ps1') -File |
             Where-Object { $_.Name -ne $selfName } |
             Sort-Object Name)

if ($Filter -ne '') {
    $scripts = @($scripts | Where-Object { $_.Name -like "*$Filter*" })
}

# Every script to run, the checks and then, with -WithExtras, the other
# verifiers, each started in the pool in that order and reported in it.
$work = New-Object System.Collections.Generic.List[object]
foreach ($s in $scripts) {
    $work.Add([pscustomobject]@{ Extra = $false; Name = $s.Name; Path = $s.FullName; Arguments = @(); What = ''; Missing = $false })
}
if ($WithExtras) {
    foreach ($e in $extras) {
        $path = Join-Path $toolsDir $e.Script
        $work.Add([pscustomobject]@{ Extra = $true; Name = ($e.Script + ' ' + ($e.Args -join ' ')).Trim(); Path = $path; Arguments = @($e.Args); What = $e.What; Missing = -not (Test-Path -LiteralPath $path) })
    }
}

$atOnce = if ($Throttle -gt 0) { $Throttle } else { [Math]::Max(1, [Environment]::ProcessorCount) }
$wall = [System.Diagnostics.Stopwatch]::StartNew()
$pool = [runspacefactory]::CreateRunspacePool(1, $atOnce)
$pool.Open()
$tasks = New-Object System.Collections.Generic.List[object]
foreach ($w in $work) {
    if ($w.Missing) { $tasks.Add([pscustomobject]@{ Work = $w; Shell = $null; Handle = $null }); continue }
    $shell = [powershell]::Create()
    $shell.RunspacePool = $pool
    [void]$shell.AddScript($runOne.ToString()).AddArgument($w.Path).AddArgument([string[]]$w.Arguments)
    $tasks.Add([pscustomobject]@{ Work = $w; Shell = $shell; Handle = $shell.BeginInvoke() })
}

Write-Output '=== ALL STATIC CHECKS ==='
Write-Output ''

$results = New-Object System.Collections.Generic.List[object]
$extraResults = New-Object System.Collections.Generic.List[object]
$headed = $false
foreach ($t in $tasks) {
    $w = $t.Work
    if ($w.Extra -and -not $headed) {
        Write-Output ''
        Write-Output '=== THE OTHER VERIFIERS ==='
        Write-Output ''
        $headed = $true
    }
    if ($w.Missing) {
        $extraResults.Add([pscustomobject]@{ Name = $w.Name; Code = 1; Output = @("not found at $($w.Path)"); Seconds = 0 })
        '  {0}  {1,-44} {2}' -f 'FAIL', $w.Name, 'NOT FOUND' | Write-Output
        continue
    }
    # The next in the order, waited for while the rest run on.
    [void]$t.Handle.AsyncWaitHandle.WaitOne()
    try {
        $got = $t.Shell.EndInvoke($t.Handle)
        $r = if ($got.Count -gt 0) { $got[$got.Count - 1] } else { [pscustomobject]@{ Code = 1; Output = @($t.Shell.Streams.Error | ForEach-Object { "$_" }); Seconds = 0 } }
    } catch {
        $r = [pscustomobject]@{ Code = 1; Output = @("the runner could not run it: $($_.Exception.Message)"); Seconds = 0 }
    }
    $t.Shell.Dispose()
    $tag = if ($r.Code -eq 0) { 'ok  ' } else { 'FAIL' }
    if ($w.Extra) {
        $extraResults.Add([pscustomobject]@{ Name = $w.Name; Code = $r.Code; Output = $r.Output; Seconds = $r.Seconds })
        '  {0}  {1,-44} {2,6:N1}s   {3}' -f $tag, $w.Name, $r.Seconds, $w.What | Write-Output
    } else {
        $results.Add([pscustomobject]@{ Name = $w.Name; Code = $r.Code; Output = $r.Output; Seconds = $r.Seconds })
        '  {0}  {1,-44} {2,6:N1}s' -f $tag, $w.Name, $r.Seconds | Write-Output
    }
}
$pool.Close()
$pool.Dispose()
$wall.Stop()

$passed = @($results | Where-Object { $_.Code -eq 0 })
$failed = @($results | Where-Object { $_.Code -ne 0 })
$extraFailed = @($extraResults | Where-Object { $_.Code -ne 0 })

# Anything that failed prints its own tail, so the answer does not need a
# second run to be actionable; in GitHub's runner, the whole of it.
$onActions = $env:GITHUB_ACTIONS -eq 'true'
$tailLength = if ($onActions) { [int]::MaxValue } else { 20 }
foreach ($f in ($failed + $extraFailed)) {
    Write-Output ''
    Write-Output ("--- $($f.Name) ---")
    $tail = @($f.Output) | Select-Object -Last $tailLength
    foreach ($ln in $tail) { Write-Output "  $ln" }
}

if ($ShowAll) {
    foreach ($p in $passed) {
        Write-Output ''
        if ($onActions) { Write-Output ("::group::$($p.Name)") }
        Write-Output ("--- $($p.Name) ---")
        foreach ($ln in @($p.Output)) { Write-Output "  $ln" }
        if ($onActions) { Write-Output '::endgroup::' }
    }
}

Write-Output ''
$total = $results.Count
$short = ($Filter -eq '') -and ($total -lt $expectedAtLeast)
$howRun = if ($atOnce -eq 1) { 'one at a time' } else { "$atOnce at a time" }

if ($Filter -ne '') {
    Write-Output ("=== {0} check(s) matching '{1}': {2} passed, {3} failed, {4} in {5:N1} s ===" -f $total, $Filter, $passed.Count, $failed.Count, $howRun, $wall.Elapsed.TotalSeconds)
    Write-Output '    (a filtered run does not test the floor - run it unfiltered for that)'
} else {
    Write-Output ("=== {0} checks: {1} passed, {2} failed, {3} in {4:N1} s ===" -f $total, $passed.Count, $failed.Count, $howRun, $wall.Elapsed.TotalSeconds)
}
if ($WithExtras) {
    Write-Output ("=== {0} other verifier(s): {1} passed, {2} failed ===" -f $extraResults.Count, ($extraResults.Count - $extraFailed.Count), $extraFailed.Count)
}

if ($short) {
    Write-Output ''
    Write-Output ("FLOOR: {0} check(s) found, but {1} were expected. A check has been deleted," -f $total, $expectedAtLeast)
    Write-Output '  renamed out of the check_*.ps1 glob, or moved. That is not a failure any of'
    Write-Output '  them can report, which is why this floor exists. If the reduction is'
    Write-Output '  deliberate, lower $expectedAtLeast in this file and say why.'
}

if ($failed.Count -gt 0 -or $extraFailed.Count -gt 0 -or $short) { exit 1 }
Write-Output ''
Write-Output 'ALL GREEN.'
exit 0
