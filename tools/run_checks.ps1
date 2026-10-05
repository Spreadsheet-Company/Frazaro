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

Usage:
  powershell -File tools\run_checks.ps1
  powershell -File tools\run_checks.ps1 -Filter prolog      # only matching names
  powershell -File tools\run_checks.ps1 -ShowAll            # print passing output too
  powershell -File tools\run_checks.ps1 -WithExtras         # also the named verifiers

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
    [string]$ToolsDir = ''
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
$expectedAtLeast = 44
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
    @{ Script = 'check_core_imports.ps1';  Args = @('-Control');    What = 'PORT.4 import reader: an empty module counts 0, a module importing a.b counts 1' }
)

function Invoke-OneScript {
    param([string]$Path, [string[]]$Arguments)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    # $ErrorActionPreference is 'Stop' for this script's own mistakes, but a
    # CHECK writing to stderr is this script's ordinary business, not an
    # error in it. In Windows PowerShell 5.1, `2>&1` on a native executable
    # wraps each stderr line in a NativeCommandError, which under 'Stop'
    # TERMINATES the runner - so a check that failed loudly would kill the
    # run instead of being reported as FAIL, and the total would never
    # print. Found by trying to make this script go red, which is the whole
    # reason for trying.
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Path @Arguments 2>&1
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $prev
    }
    $sw.Stop()
    return [pscustomobject]@{
        Code    = $code
        Output  = $out
        Seconds = $sw.Elapsed.TotalSeconds
    }
}

$scripts = @(Get-ChildItem -Path (Join-Path $toolsDir 'check_*.ps1') -File |
             Where-Object { $_.Name -ne $selfName } |
             Sort-Object Name)

if ($Filter -ne '') {
    $scripts = @($scripts | Where-Object { $_.Name -like "*$Filter*" })
}

Write-Output '=== ALL STATIC CHECKS ==='
Write-Output ''

$results = New-Object System.Collections.Generic.List[object]
foreach ($s in $scripts) {
    $r = Invoke-OneScript -Path $s.FullName -Arguments @()
    $results.Add([pscustomobject]@{
        Name = $s.Name; Code = $r.Code; Output = $r.Output; Seconds = $r.Seconds
    })
    $tag = if ($r.Code -eq 0) { 'ok  ' } else { 'FAIL' }
    '  {0}  {1,-44} {2,6:N1}s' -f $tag, $s.Name, $r.Seconds | Write-Output
}

$passed = @($results | Where-Object { $_.Code -eq 0 })
$failed = @($results | Where-Object { $_.Code -ne 0 })

$extraResults = New-Object System.Collections.Generic.List[object]
if ($WithExtras) {
    Write-Output ''
    Write-Output '=== THE OTHER VERIFIERS ==='
    Write-Output ''
    foreach ($e in $extras) {
        $path = Join-Path $toolsDir $e.Script
        if (-not (Test-Path -LiteralPath $path)) {
            $extraResults.Add([pscustomobject]@{
                Name = $e.Script; Code = 1; Output = @("not found at $path"); Seconds = 0
            })
            '  {0}  {1,-44} {2}' -f 'FAIL', ($e.Script + ' ' + ($e.Args -join ' ')), 'NOT FOUND' | Write-Output
            continue
        }
        $r = Invoke-OneScript -Path $path -Arguments $e.Args
        $extraResults.Add([pscustomobject]@{
            Name = ($e.Script + ' ' + ($e.Args -join ' ')).Trim()
            Code = $r.Code; Output = $r.Output; Seconds = $r.Seconds
        })
        $tag = if ($r.Code -eq 0) { 'ok  ' } else { 'FAIL' }
        '  {0}  {1,-44} {2,6:N1}s   {3}' -f $tag, ($e.Script + ' ' + ($e.Args -join ' ')).Trim(), $r.Seconds, $e.What | Write-Output
    }
}

$extraFailed = @($extraResults | Where-Object { $_.Code -ne 0 })

# Anything that failed prints its own tail, so the answer does not need a
# second run to be actionable.
foreach ($f in ($failed + $extraFailed)) {
    Write-Output ''
    Write-Output ("--- $($f.Name) ---")
    $tail = @($f.Output) | Select-Object -Last 20
    foreach ($ln in $tail) { Write-Output "  $ln" }
}

if ($ShowAll) {
    foreach ($p in $passed) {
        Write-Output ''
        Write-Output ("--- $($p.Name) ---")
        foreach ($ln in @($p.Output)) { Write-Output "  $ln" }
    }
}

Write-Output ''
$total = $results.Count
$short = ($Filter -eq '') -and ($total -lt $expectedAtLeast)

if ($Filter -ne '') {
    Write-Output ("=== {0} check(s) matching '{1}': {2} passed, {3} failed ===" -f $total, $Filter, $passed.Count, $failed.Count)
    Write-Output '    (a filtered run does not test the floor - run it unfiltered for that)'
} else {
    Write-Output ("=== {0} checks: {1} passed, {2} failed ===" -f $total, $passed.Count, $failed.Count)
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
