<#
check_recalc_floors.ps1 - KERNEL.20's oracle: the machine's recalculation
of seven fixtures measured on the owner's machine beside Excel's of the same
files, held to floors that never go down; and, on every push where the
runner is built, the fixtures' values held.

WHY: KERNEL.20 (docs/BETA_ROADMAP.md: "Oracle: check_recalc_floors.ps1, the
timings as floors that never go down, in the house's measure-first form";
web/CALLOSUM.md section 12.6 calls the fastest spreadsheet a prediction
until it is measured). Its sixth slice built the measurement
(docs/BETA_REARVIEW.md, its scoping): core/examples/recalc.rs builds each
program of scripts/recalc/bench/ as frazaro build builds it and times the
machine computing it, under Excel's library, after holding every value;
tools/bench_recalc.ps1 runs it on the owner's machine and
tools/bench_recalc_excel.ps1 times Excel on the same files. A timing is a
measurement on the owner's machine, so the floors are the owner's run of
the two, pasted below.

Since KERNEL.49's first slice (docs/BETA_REARVIEW.md, the item's scoping,
decision 12 and its addendum) the native half has a second runner,
core/examples/memory.rs: every bench program, and two fixtures of rows the
machine loads as they stand (scripts/recalc/bench/superposition.vla and
life-x16.vla, Alonzo's Superposition and Life on the size ladder's sixteenth
rung), each loaded into the machine under a counting allocator and stepped
twice: the bytes it holds after its load and after two frames, and the most
it held between; and its load timed. Bytes asked of the allocator are the
same on every run of one build and on every 64-bit machine, so they are held
exactly, as ceilings that never rise, here and on every push. Since KERNEL.67
a third fixture of rows adds to a twin's: superposition-styled.vla, one style
row over Superposition's Field, loaded after superposition.vla's rows, and
held within a constant of its twin's bytes, so that a style row over a range
holds its rectangle and never a format a cell. Since KERNEL.72 the runner prints
each fixture's save after its frames, every sheet's fills, and loads it again,
so that a grid's save stays small and loading it costs about what the grid
holds: its text's bytes, the most its printing holds over the machine, and a
load from it, held after the load and at the peak.

WHAT IT HOLDS, from the baseline pasted below:
  - a date and two environment lines, each saying, as its instrument
    writes it, that it ran opted out of Windows' power throttling (an
    invisible or background process is slowed by about 40% a few seconds
    in), the native half on the performance cores and at high priority, and
    Excel's half by Excel and not the instrument's stand-in;
  - seven native rows, one a fixture, each measured in at least three
    rounds, its clock within 10% across them and from the run's start, its
    per-second numbers agreeing with its times within half a percent; a
    typed fixture's row (KERNEL.56, the netlist's) its type as well, a value
    written into a fresh cell and the frames to the grid's stillness, and
    Excel's rows of it their type too; a rewritten fixture's row (KERNEL.69, the
    netlist's) its rewrite as well, a formula written over and back, its write
    and its frames to the grid's stillness apart, and Excel's rows theirs;
  - fourteen Excel rows, one a fixture and a mode (all threads, as Excel
    ships, and one): each fixture's file the same bytes in both halves, by
    its SHA-256, and Excel's sums of the program's range, as built and as
    edited, equal to the machine's in both modes, so the two computed the
    same values;
  - three flag rows, Excel's reading of the build golden's fixture written
    with fullCalcOnLoad, with Excel 365's calcId, and with neither, one of
    the three verdicts each;
  - the fixtures' programs, each one's SHA-256 over LF bytes equal to the
    one measured, so a program edited without a new measurement fails by
    name;
  - the floors, a dated history, each fixture's whole frames in formulas a
    second and its edits a second, a typed fixture's types a second
    (KERNEL.56), and a rewritten fixture's frames and writes a second
    (KERNEL.69): each entry's floor at least the entry
    before it unless the entry gives its reason, a fixture's floors from the
    entry it joined on and every fixture's in the latest; each native row's numbers
    at least the latest floors. The floor is the run's number less a margin
    of 15%: measured over three rounds, a fixture's medians moved up to 11%
    between processes (chain's whole frame, 3.87 to 4.31 ms over three runs,
    2026-10-10, the slice's build), and Alonzo's 10% would not hold it;
  - ten memory rows, the memory runner's, one a fixture: its bytes held
    after its load, after two frames and at the peak, the peak at least the
    bytes held, its source's SHA-256 over LF bytes (the program's, or the
    rows' own file) the one pinned here; each load timed in at least three
    rounds, its clock within 10%, its formulas loaded a second agreeing
    with its time; and the three fixtures of rows as they stand, each one's
    SHA-256 over LF bytes the one measured;
  - a fixture of rows that adds to a twin's within 4,096 bytes of its twin's
    after its load, after two frames and at the peak (KERNEL.67), and through
    the round trip, in its save, printing it, after a load from it and at that
    load's peak (KERNEL.72), in the baseline's rows and in the memory runner's
    check mode alike;
  - what a save costs (KERNEL.72), in the baseline's rows and in the memory
    runner's check mode alike: a load from a save holds no more than the
    machine that printed it; a save's printing peaks at no more than four times
    its text; a fixture of rows, a machine's own fills, saves in at most 4 bytes
    a formula cell, and a load from its save peaks within 4,096 bytes of what it
    holds; a program's, a record a cell until the machine takes them, at no more
    than four times. Each bound the rule of the item's title and no pinned
    number: the ceilings hold those;
  - the ceilings, a dated history beside the floors', each fixture's bytes
    held and at the peak, and since KERNEL.72 its save's and a load from it:
    each entry's ceiling at most the entry before it
    unless the entry gives its reason, a fixture's ceilings from the entry
    it joined on and every fixture's in the latest; each memory row's bytes
    at most the latest ceilings. A ceiling is the run's bytes exactly, with
    no margin, since the count is exact;
  - three load floors among the floors, the fixtures of rows' formulas loaded
    a second, the run's number less the margin, held as a frame's are;
  - both histories appended to and never pasted over: every entry the last
    commit's copy of this file holds is still here, in its order, at the
    head (SKIPPED where git, the commit or the file in it is not there to
    ask), since a paste over an entry would erase it and could lower a
    floor, or raise a ceiling, with no reason written (Alonzo's frame floors
    lost an entry so);
  - the bar: Excel's numbers printed beside the machine's, with the ratio;
    a measurement, not gated, since Excel's numbers move with its version.
An empty baseline fails: a floor with no number is no floor.

WITH A BUILT RUNNER (target/debug/examples/recalc or
target/release/examples/recalc, the newer, or -Runner): its check
mode, every value held (each fixture's first frame to the values its build
folded, its edited state to Calc, every frame of Life to a Life written by
hand, every gate of the netlist to its adder by hand, and the netlist's
program its generator's), and each fixture's sums equal to the baseline's,
so the values measured are the values computed today. A file's bytes may
move with the stamp (a phrasebook or the prelude edited) while its values
do not; that is printed and not failed. SKIPPED without a runner: the ratchets job builds
none. WITH A BUILT MEMORY RUNNER (target/debug/examples/memory or
target/release/examples/memory, the newer, or -MemoryRunner): its check
mode, every fixture's bytes, exact, at most its
latest ceilings, and its formulas the baseline's; bytes under a ceiling
printed, so that a new entry may lower it. SKIPPED without one.

THE RUNNERS SINCE AS.10 (2026-10-21): `cargo test --workspace` builds both
in the dev profile, optimized since that day, and their check modes print
there what the release profile's print, byte for byte. A runner found, not
named, and older than a file it is built from (the two crates' sources,
data and examples, and the manifests) is SKIPPED by name, since it would
check a tree that is no longer here. The two run side by side. Each one's
output is kept under target/recalc_floors/ with the SHA-256 of the runner
and of every file it reads (scripts/recalc/bench/, the prelude, english.vla
and the build golden's fixture), and reused while both are byte for byte
the ones it came from: the runner reads no clock and draws no chance, so it
would print the same lines, which are judged against this baseline again
every time. Only output that passed is kept. Delete the folder to run them
again.

-Control proves the judgment: a clean baseline passes, and so does one
with a floor lowered for a written reason, and one with a ceiling raised for
a written reason; an empty baseline, a native row
missing, an Excel row missing, a file not the same bytes in both halves,
sums that differ, a clock that moved, too few rounds, per-second numbers
that disagree with the times, a native half not opted out, Excel's half
from the stand-in, a floor lowered with no reason, a row under its floor, a
program edited, a flag row with no verdict and a fixture Excel could not
open must each fail, and so must a memory row missing, a memory row over its
ceiling, a peak under the bytes held, a ceiling raised with no reason, a
fixture's ceilings left out after it joined, a load under its floor, a load
a second that disagrees with its time, a fixture of rows edited, and a
fixture over its twin's bytes by more than the bound, while one at the bound
passes; the histories,
read from text as the commit's copy is: an entry added after the committed
ones passes, and so does a history unchanged, while an entry pasted over a
committed one and a committed entry edited in place must each fail; then
fake runners: a clean one passes, and one whose values differ, one whose
sums differ and one that leaves out a fixture must each fail; and fake
memory runners: a clean one passes, and one over a ceiling, one whose
formulas differ, one that leaves out a fixture, one that failed and one
whose fixture holds more than its twin by more than the bound must each
fail. Then the runners' finding, running and keeping (AS.10): of two
built, the newer is found; a runner older than a source is named stale and
one newer is not; two fake runners started together run at once, read off
the times each writes; a clean output kept is read back from the cache
with the runner not run again, while a runner of other bytes, inputs of
other bytes and an output that failed each run it again.

House style (tools/check_*.ps1): PowerShell 5.1 and pwsh alike, host-free,
no network; a hardcoded, reviewable baseline; exit 0 clean, exit 1 with
every problem named.

Usage:  powershell -File tools\check_recalc_floors.ps1
        powershell -File tools\check_recalc_floors.ps1 -Runner target\release\examples\recalc.exe
        powershell -File tools\check_recalc_floors.ps1 -MemoryRunner target\release\examples\memory.exe
        powershell -File tools\check_recalc_floors.ps1 -Control
#>
param(
    [string]$Root = '',
    [string]$Runner = '',
    [string]$MemoryRunner = '',
    [switch]$Control
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if ($Root -eq '') { $Root = $repoRoot }
$onWindows = [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT

# --- the baseline, hand-maintained: the owner's run of the two instruments ---
# Paste the native half's block (tools/bench_recalc.ps1) over $baselineDate,
# $baselineNative and its seven rows and ten memory rows, Excel's half's
# (tools/bench_recalc_excel.ps1) over $baselineExcel and its seventeen rows, and
# the two histories the native half prints over $floorHistory and
# $ceilingHistory, whole: each is this file's history with the run's entry
# added last. A later run slower than a floor, or heavier than a ceiling, is
# a finding to record in docs/BETA_REARVIEW.md, never a number to paste over
# a better one.
$baselineDate = '2026-11-01'
$baselineNative = 'Intel(R) Core(TM) Ultra 7 265KF, 20 logical processors | Microsoft Windows 11 Home 10.0.26200.0 | rustc 1.99.0 (b940084d7 2026-09-28) | frazaro-core 0.12.0 | power throttling opted out: the runner yes | high priority: the runner yes | on the performance cores, 8 of 20 logical processors (efficiency class 1): the runner yes'
$baselineExcel = 'Excel''s half, 2026-11-01: Intel(R) Core(TM) Ultra 7 265KF | Microsoft Windows 11 Home 10.0.26200.0 | Excel 16.0 build 20527 | all threads 20; one thread on the performance cores, mask 0xC03C3 | power throttling opted out: yes'
$baseline = @(
    # the native rows, of the run of 2026-11-01, the files Excel's half opened. Marked for a later pass,
    # after 0.13.0's bump moves the files' bytes again: both instruments run, both blocks pasted whole:
    @{ Half = 'native'; Fixture = 'column'; Formulas = 60000; Sha256 = 'A2D8C4EB45CD88AC49C9147A46133B5DE0D9CF63070744C2B5BA1CC2C106FF30'; Rounds = 3; Load = 214.5; First = 1.30; Whole = 0.422; Edit = 1.098; Calc = 218.4; WholePerSecond = 142112743; EditsPerSecond = 910.7; SumBuilt = 1100080000; SumEdited = 1425100000; ClockMs = 0.178; ClockSpread = 5.8; ClockDrift = -4.6 }
    @{ Half = 'native'; Fixture = 'chain'; Formulas = 39999; Sha256 = 'F76AB27B97D13C6D377E64C3B4C349B9A6ACC3622E7E377ED006BEA56FBE53D1'; Rounds = 3; Load = 133.7; First = 1.14; Whole = 0.554; Edit = 1.316; Calc = 118.3; WholePerSecond = 72213396; EditsPerSecond = 759.6; SumBuilt = 1200039974; SumEdited = 1200059975; ClockMs = 0.178; ClockSpread = 5.9; ClockDrift = -4.7 }
    @{ Half = 'native'; Fixture = 'sums'; Formulas = 28014; Sha256 = '97BF496D5BC647D5E0E1121C9E6EF37F9C6A50CDCB431460981FB06232778D68'; Rounds = 3; Load = 103.7; First = 1.19; Whole = 0.771; Edit = 1.267; Calc = 143.6; WholePerSecond = 36315789; EditsPerSecond = 789.5; SumBuilt = 14800794; SumEdited = 14900794; ClockMs = 0.178; ClockSpread = 6.1; ClockDrift = -4.6 }
    @{ Half = 'native'; Fixture = 'life'; Formulas = 62964; Sha256 = '909C55816387F98469B83C6B063B26510BAA60BF7CA4FF5BA5B2CBF1785C13C1'; Rounds = 3; Load = 1326.0; First = 3.50; Whole = 1.811; Edit = 2.321; Calc = 2884.6; WholePerSecond = 34767532; EditsPerSecond = 430.8; SumBuilt = 23683; SumEdited = 9603; ClockMs = 0.187; ClockSpread = 6.0; ClockDrift = 0.1 }
    @{ Half = 'native'; Fixture = 'lookups'; Formulas = 140000; Sha256 = '3A1D792702414449471A9ACF60E11D6E17B1F06716373A7B73ADD0FC519F08D1'; Rounds = 3; Load = 636.6; First = 11.40; Whole = 7.298; Edit = 9.236; Calc = 704.1; WholePerSecond = 19183601; EditsPerSecond = 108.3; SumBuilt = 51995036; SumEdited = 52005012; ClockMs = 0.187; ClockSpread = 5.8; ClockDrift = -0.1 }
    @{ Half = 'native'; Fixture = 'criteria'; Formulas = 61370; Sha256 = 'A16458EA427961F4F92A34B661E091715385692B8C6651844DD9C3E42F120AF6'; Rounds = 3; Load = 181.6; First = 7.43; Whole = 1.240; Edit = 3.944; Calc = 260.9; WholePerSecond = 49483954; EditsPerSecond = 253.6; SumBuilt = 12010000; SumEdited = 12030000; ClockMs = 0.178; ClockSpread = 6.2; ClockDrift = -4.5 }
    @{ Half = 'native'; Fixture = 'netlist'; Formulas = 20480; Sha256 = '1D707F4433CE055C8E7103BFFFFAAA8C4AB7F6502969CF6F45DC7AA74B902805'; Rounds = 3; Load = 279.0; First = 3.67; Whole = 1.621; Edit = 4.989; Calc = 75.3; WholePerSecond = 12638075; EditsPerSecond = 200.4; Type = 0.003; TypesPerSecond = 322580.6; Rewrite = 0.006; RewritesPerSecond = 175438.6; RewriteWrite = 30.746; RewriteWritesPerSecond = 32.5; SumBuilt = 8192; SumEdited = 12288; ClockMs = 0.180; ClockSpread = 6.2; ClockDrift = -3.4 }
    # the memory rows (KERNEL.49), of the same run, their bytes the memory runner's check at KERNEL.72:
    @{ Half = 'memory'; Fixture = 'column'; Formulas = 60000; Sha256 = '751885133C249C25C8C4BB3854877A72E6054B2E0945BB2C1931C3FF05D55E2E'; Rounds = 3; Load = 272.7; LoadPerSecond = 220009; LoadHeld = 10438452; Held = 10463956; Peak = 23110916; Save = 3261708; SavePeak = 9553164; SavedHeld = 10437445; SavedPeak = 23602317; ClockMs = 0.182; ClockSpread = 0.8; ClockDrift = 0.2 }
    @{ Half = 'memory'; Fixture = 'chain'; Formulas = 39999; Sha256 = '72FA6E5DB80DF71796473E21B89432C598AD8719C473D8955B24B796417A8A5A'; Rounds = 3; Load = 177.2; LoadPerSecond = 225701; LoadHeld = 7031446; Held = 7064278; Peak = 20219356; Save = 1963032; SavePeak = 5855654; SavedHeld = 7030439; SavedPeak = 20546575; ClockMs = 0.182; ClockSpread = 0.2; ClockDrift = 0.2 }
    @{ Half = 'memory'; Fixture = 'sums'; Formulas = 28014; Sha256 = 'A339D809DE2493BBA3A129FC3FC4A5AEC458A5A6F1B5FB590DB04D2CCB3F7348'; Rounds = 3; Load = 140.4; LoadPerSecond = 199512; LoadHeld = 4604155; Held = 4636987; Peak = 11858469; Save = 1823353; SavePeak = 4969081; SavedHeld = 4603148; SavedPeak = 12086670; ClockMs = 0.182; ClockSpread = 0.5; ClockDrift = 0.2 }
    @{ Half = 'memory'; Fixture = 'life'; Formulas = 62964; Sha256 = 'C2B7AC34005AB88CFE2F4CE98FAC3C704CCF26E8AEED90BE221BF800D63F1716'; Rounds = 3; Load = 1627.9; LoadPerSecond = 38679; LoadHeld = 40489701; Held = 40624645; Peak = 53634545; Save = 32623420; SavePeak = 74173244; SavedHeld = 40488692; SavedPeak = 54175816; ClockMs = 0.182; ClockSpread = 0.2; ClockDrift = 0.2 }
    @{ Half = 'memory'; Fixture = 'lookups'; Formulas = 140000; Sha256 = '5D25CCE0460063145619F43B64ED898BA6A073F5F83003D53B92B8D1C8C1619A'; Rounds = 3; Load = 722.9; LoadPerSecond = 193651; LoadHeld = 25957942; Held = 29487646; Peak = 74421446; Save = 8901488; SavePeak = 21484400; SavedHeld = 25956933; SavedPeak = 75556497; ClockMs = 0.178; ClockSpread = 5.3; ClockDrift = -1.9 }
    @{ Half = 'memory'; Fixture = 'criteria'; Formulas = 61370; Sha256 = 'B80A36C00F9445793B061A2CCAA2138F7B0FC407469A87EC6259F1A878C48D63'; Rounds = 3; Load = 220.7; LoadPerSecond = 278097; LoadHeld = 10143314; Held = 12209962; Peak = 23344415; Save = 3372714; SavePeak = 9664170; SavedHeld = 10142305; SavedPeak = 23852392; ClockMs = 0.182; ClockSpread = 0.2; ClockDrift = 0.3 }
    @{ Half = 'memory'; Fixture = 'netlist'; Formulas = 20480; Sha256 = 'CD618419C5220021E55F9DDEE66699DDCBE9E384E2FFCC150A2C86FC56741C26'; Rounds = 3; Load = 342.6; LoadPerSecond = 59770; LoadHeld = 53186266; Held = 53186298; Peak = 63014223; Save = 2530377; SavePeak = 6105357; SavedHeld = 53185253; SavedPeak = 60515836; ClockMs = 0.182; ClockSpread = 1.7; ClockDrift = 0.2 }
    @{ Half = 'memory'; Fixture = 'superposition'; Formulas = 798731; Sha256 = '6A3AE552F614AC0CD5582D69ED0992E804032831798E92C0445EEFA478551B98'; Rounds = 3; Load = 12.3; LoadPerSecond = 64823117; LoadHeld = 10583312; Held = 10803600; Peak = 10804040; Save = 2372117; SavePeak = 4995415; SavedHeld = 10583312; SavedPeak = 10583376; ClockMs = 0.182; ClockSpread = 0.7; ClockDrift = 0.1 }
    @{ Half = 'memory'; Fixture = 'life-x16'; Formulas = 1019844; Sha256 = '7366BE8061601852C8447E246CA98E58D11DF699A965B659BB48826975ED001D'; Rounds = 3; Load = 14.4; LoadPerSecond = 70742424; LoadHeld = 18840036; Held = 18974980; Peak = 18974980; Save = 2051565; SavePeak = 4189031; SavedHeld = 18840036; SavedPeak = 18840100; ClockMs = 0.182; ClockSpread = 0.1; ClockDrift = 0.1 }
    @{ Half = 'memory'; Fixture = 'superposition-styled'; Formulas = 798731; Sha256 = '67097ADEFC12CFDB5A0C99F564A239AAF811FD7E6B0005CE1DED53AB411024B3'; Rounds = 3; Load = 11.7; LoadPerSecond = 67988679; LoadHeld = 10583832; Held = 10804120; Peak = 10804560; Save = 2372233; SavePeak = 4995647; SavedHeld = 10583832; SavedPeak = 10583896; ClockMs = 0.182; ClockSpread = 1.9; ClockDrift = 0.2 }
    # Excel's half, its flag first:
    @{ Half = 'excel-flag'; File = 'calcpr_built'; C2 = 15; Saved = 'True'; Verdict = 'computed on opening' }
    @{ Half = 'excel-flag'; File = 'calcpr_calcid'; C2 = 15; Saved = 'True'; Verdict = 'computed on opening' }
    @{ Half = 'excel-flag'; File = 'calcpr_none'; C2 = 15; Saved = 'True'; Verdict = 'computed on opening' }
    @{ Half = 'excel'; Fixture = 'column'; Threads = 'all'; ThreadCount = 20; Sha256 = 'A2D8C4EB45CD88AC49C9147A46133B5DE0D9CF63070744C2B5BA1CC2C106FF30'; Open = 153.4; Whole = 1.479; Rebuild = 11.471; Edit = 1.382; Copy = 0.0; ComFloor = 0.253; SumBuilt = 1100080000; SumEdited = 1425100000 }
    @{ Half = 'excel'; Fixture = 'column'; Threads = 'one'; ThreadCount = 1; Sha256 = 'A2D8C4EB45CD88AC49C9147A46133B5DE0D9CF63070744C2B5BA1CC2C106FF30'; Open = 153.4; Whole = 2.883; Rebuild = 11.904; Edit = 2.249; Copy = 0.0; ComFloor = 0.255; SumBuilt = 1100080000; SumEdited = 1425100000 }
    @{ Half = 'excel'; Fixture = 'chain'; Threads = 'all'; ThreadCount = 20; Sha256 = 'F76AB27B97D13C6D377E64C3B4C349B9A6ACC3622E7E377ED006BEA56FBE53D1'; Open = 95.8; Whole = 2.389; Rebuild = 5.846; Edit = 1.023; Copy = 0.0; ComFloor = 0.239; SumBuilt = 1200039974; SumEdited = 1200059975 }
    @{ Half = 'excel'; Fixture = 'chain'; Threads = 'one'; ThreadCount = 1; Sha256 = 'F76AB27B97D13C6D377E64C3B4C349B9A6ACC3622E7E377ED006BEA56FBE53D1'; Open = 95.8; Whole = 2.713; Rebuild = 6.793; Edit = 0.774; Copy = 0.0; ComFloor = 0.237; SumBuilt = 1200039974; SumEdited = 1200059975 }
    @{ Half = 'excel'; Fixture = 'sums'; Threads = 'all'; ThreadCount = 20; Sha256 = '97BF496D5BC647D5E0E1121C9E6EF37F9C6A50CDCB431460981FB06232778D68'; Open = 77.0; Whole = 1.607; Rebuild = 6.372; Edit = 1.340; Copy = 0.0; ComFloor = 0.260; SumBuilt = 14800794; SumEdited = 14900794 }
    @{ Half = 'excel'; Fixture = 'sums'; Threads = 'one'; ThreadCount = 1; Sha256 = '97BF496D5BC647D5E0E1121C9E6EF37F9C6A50CDCB431460981FB06232778D68'; Open = 77.0; Whole = 4.298; Rebuild = 9.267; Edit = 4.272; Copy = 0.0; ComFloor = 0.243; SumBuilt = 14800794; SumEdited = 14900794 }
    @{ Half = 'excel'; Fixture = 'life'; Threads = 'all'; ThreadCount = 20; Sha256 = '909C55816387F98469B83C6B063B26510BAA60BF7CA4FF5BA5B2CBF1785C13C1'; Open = 2354.5; Whole = 3.274; Rebuild = 1133.207; Edit = 3.386; Copy = 65.5; ComFloor = 0.241; SumBuilt = 23683; SumEdited = 9603; CopyWay = 'paste values'; CopyRefused = 'the range itself: Specified cast is not valid.' }
    @{ Half = 'excel'; Fixture = 'life'; Threads = 'one'; ThreadCount = 1; Sha256 = '909C55816387F98469B83C6B063B26510BAA60BF7CA4FF5BA5B2CBF1785C13C1'; Open = 2354.5; Whole = 28.778; Rebuild = 1153.159; Edit = 29.439; Copy = 64.9; ComFloor = 0.224; SumBuilt = 23683; SumEdited = 9603; CopyWay = 'paste values'; CopyRefused = 'the range itself: Specified cast is not valid.' }
    @{ Half = 'excel'; Fixture = 'lookups'; Threads = 'all'; ThreadCount = 20; Sha256 = '3A1D792702414449471A9ACF60E11D6E17B1F06716373A7B73ADD0FC519F08D1'; Open = 390.1; Whole = 77.261; Rebuild = 106.797; Edit = 75.815; Copy = 0.0; ComFloor = 0.228; SumBuilt = 51995036; SumEdited = 52005012 }
    @{ Half = 'excel'; Fixture = 'lookups'; Threads = 'one'; ThreadCount = 1; Sha256 = '3A1D792702414449471A9ACF60E11D6E17B1F06716373A7B73ADD0FC519F08D1'; Open = 390.1; Whole = 722.756; Rebuild = 737.383; Edit = 722.093; Copy = 0.0; ComFloor = 0.198; SumBuilt = 51995036; SumEdited = 52005012 }
    @{ Half = 'excel'; Fixture = 'criteria'; Threads = 'all'; ThreadCount = 20; Sha256 = 'A16458EA427961F4F92A34B661E091715385692B8C6651844DD9C3E42F120AF6'; Open = 132.1; Whole = 11.860; Rebuild = 16.848; Edit = 9.652; Copy = 0.0; ComFloor = 0.178; SumBuilt = 12010000; SumEdited = 12030000 }
    @{ Half = 'excel'; Fixture = 'criteria'; Threads = 'one'; ThreadCount = 1; Sha256 = 'A16458EA427961F4F92A34B661E091715385692B8C6651844DD9C3E42F120AF6'; Open = 132.1; Whole = 87.536; Rebuild = 96.730; Edit = 81.828; Copy = 0.0; ComFloor = 0.246; SumBuilt = 12010000; SumEdited = 12030000 }
    @{ Half = 'excel'; Fixture = 'netlist'; Threads = 'all'; ThreadCount = 20; Sha256 = '1D707F4433CE055C8E7103BFFFFAAA8C4AB7F6502969CF6F45DC7AA74B902805'; Open = 222.0; Whole = 4.205; Rebuild = 12.042; Edit = 3.865; Type = 0.548; Rewrite = 1.913; RewriteWrite = 0.592; Copy = 0.0; ComFloor = 0.241; SumBuilt = 8192; SumEdited = 12288 }
    @{ Half = 'excel'; Fixture = 'netlist'; Threads = 'one'; ThreadCount = 1; Sha256 = '1D707F4433CE055C8E7103BFFFFAAA8C4AB7F6502969CF6F45DC7AA74B902805'; Open = 222.0; Whole = 1.542; Rebuild = 9.746; Edit = 1.294; Type = 0.555; Rewrite = 0.347; RewriteWrite = 0.498; Copy = 0.0; ComFloor = 0.241; SumBuilt = 8192; SumEdited = 12288 }
)

# --- the floors, a dated history that never goes down ---
# One entry a measurement, one a line inside the parentheses, a comma after
# the entry before it: @{ Date = 'yyyy-mm-dd'; Reason = ''; Floors = @{
# 'column whole' = n; 'column edit' = n; ... } }, whole frames in formulas a
# second and edits a second. An entry lower than the one before it in any
# floor needs its Reason.
$floorHistory = @(
    @{ Date = '2026-10-11'; Reason = ''; Floors = @{ 'column whole' = 47700000; 'column edit' = 353; 'chain whole' = 10600000; 'chain edit' = 318; 'sums whole' = 5560000; 'sums edit' = 176; 'life whole' = 15400000; 'life edit' = 314 } },
    @{ Date = '2026-10-12'; Reason = 'KERNEL.28, at the owner''s ruling: column edit and chain edit, rows the item does not move, lowered to this run less the margin; the unchanged code (5607910) took 2.83 to 3.23 and 2.92 to 3.58 ms an edit in nine runs on 2026-10-12, against the 2.41 and 2.67 of the run of 2026-10-11 the floors were set from (docs/BETA_REARVIEW.md, KERNEL.28)'; Floors = @{ 'column whole' = 52500000; 'column edit' = 272; 'chain whole' = 11900000; 'chain edit' = 233; 'sums whole' = 7040000; 'sums edit' = 223; 'life whole' = 15400000; 'life edit' = 314 } },
    @{ Date = '2026-10-16'; Reason = ''; Floors = @{ 'column whole' = 55200000; 'column edit' = 478; 'chain whole' = 14000000; 'chain edit' = 248; 'sums whole' = 7080000; 'sums edit' = 241; 'life whole' = 28500000; 'life edit' = 368 } },
    @{ Date = '2026-10-16'; Reason = ''; Floors = @{ 'column whole' = 56800000; 'column edit' = 478; 'chain whole' = 63000000; 'chain edit' = 609; 'sums whole' = 7210000; 'sums edit' = 247; 'life whole' = 28500000; 'life edit' = 380 } },
    @{ Date = '2026-10-17'; Reason = ''; Floors = @{ 'column whole' = 108000000; 'column edit' = 583; 'chain whole' = 63000000; 'chain edit' = 638; 'sums whole' = 31000000; 'sums edit' = 568; 'life whole' = 30500000; 'life edit' = 410 } },
    @{ Date = '2026-10-19'; Reason = 'KERNEL.8''s first slice, at the owner''s ruling: column whole and column edit, rows the slice does not move (its program reads no lookup), lowered to this run less the margin; the base the slice was built on (1cc86c9) took 0.572, 0.529, 0.604 and 0.527 ms a whole frame and 1.746, 1.756, 1.821 and 1.595 ms an edit through the same instrument on 2026-10-18, under the floors of 2026-10-17 in five of its eight numbers, which were set from one run of 0.468 and 1.456 ms (docs/BETA_REARVIEW.md, KERNEL.8)'; Floors = @{ 'column whole' = 87800000; 'column edit' = 490; 'chain whole' = 63000000; 'chain edit' = 638; 'sums whole' = 31000000; 'sums edit' = 568; 'life whole' = 30500000; 'life edit' = 410; 'lookups whole' = 12400000; 'lookups edit' = 72 } },
    @{ Date = '2026-10-19'; Reason = ''; Floors = @{ 'column whole' = 87800000; 'column edit' = 490; 'chain whole' = 63000000; 'chain edit' = 638; 'sums whole' = 31000000; 'sums edit' = 568; 'life whole' = 30500000; 'life edit' = 410; 'lookups whole' = 12400000; 'lookups edit' = 72 } },
    @{ Date = '2026-10-20'; Reason = 'KERNEL.8''s fifth slice, at the owner''s ruling: sums whole and life edit, rows the slice does not move, lowered to this run less the margin; the base the slice was built on (3d06abc) took 0.898 to 0.959 ms a whole frame of sums and 2.293 to 2.479 ms an edit of Life in eight rounds through the same instrument on 2026-10-20, under the floors of 2026-10-17 in seven and in two of the eight, floors set from one run (docs/BETA_REARVIEW.md, KERNEL.8)'; Floors = @{ 'column whole' = 87800000; 'column edit' = 490; 'chain whole' = 63000000; 'chain edit' = 638; 'sums whole' = 25600000; 'sums edit' = 568; 'life whole' = 30500000; 'life edit' = 345; 'lookups whole' = 12400000; 'lookups edit' = 72; 'criteria whole' = 4250000; 'criteria edit' = 106 } },
    @{ Date = '2026-10-21'; Reason = ''; Floors = @{ 'column whole' = 87800000; 'column edit' = 490; 'chain whole' = 63000000; 'chain edit' = 638; 'sums whole' = 25700000; 'sums edit' = 568; 'life whole' = 30500000; 'life edit' = 356; 'lookups whole' = 12400000; 'lookups edit' = 72; 'criteria whole' = 4250000; 'criteria edit' = 109; 'superposition load' = 1100000; 'life-x16 load' = 1340000 } },
    @{ Date = '2026-10-21'; Reason = ''; Floors = @{ 'column whole' = 87800000; 'column edit' = 490; 'chain whole' = 63000000; 'chain edit' = 638; 'sums whole' = 25700000; 'sums edit' = 568; 'life whole' = 30500000; 'life edit' = 356; 'lookups whole' = 12400000; 'lookups edit' = 72; 'criteria whole' = 4250000; 'criteria edit' = 111; 'netlist whole' = 2500000; 'netlist edit' = 12; 'superposition load' = 1100000; 'life-x16 load' = 1340000 } },
    @{ Date = '2026-10-21'; Reason = ''; Floors = @{ 'column whole' = 87800000; 'column edit' = 537; 'chain whole' = 63000000; 'chain edit' = 638; 'sums whole' = 25700000; 'sums edit' = 568; 'life whole' = 30500000; 'life edit' = 356; 'lookups whole' = 12400000; 'lookups edit' = 72; 'criteria whole' = 4340000; 'criteria edit' = 111; 'netlist whole' = 2650000; 'netlist edit' = 114; 'superposition load' = 1100000; 'life-x16 load' = 1370000 } },
    @{ Date = '2026-10-22'; Reason = ''; Floors = @{ 'column whole' = 100000000; 'column edit' = 772; 'chain whole' = 63000000; 'chain edit' = 664; 'sums whole' = 28600000; 'sums edit' = 670; 'life whole' = 30600000; 'life edit' = 364; 'lookups whole' = 12500000; 'lookups edit' = 80; 'criteria whole' = 4960000; 'criteria edit' = 151; 'netlist whole' = 4180000; 'netlist edit' = 139; 'superposition load' = 1150000; 'life-x16 load' = 1370000 } },
    @{ Date = '2026-10-22'; Reason = ''; Floors = @{ 'column whole' = 101000000; 'column edit' = 778; 'chain whole' = 63000000; 'chain edit' = 667; 'sums whole' = 29100000; 'sums edit' = 684; 'life whole' = 30800000; 'life edit' = 364; 'lookups whole' = 12600000; 'lookups edit' = 81; 'criteria whole' = 4960000; 'criteria edit' = 152; 'netlist whole' = 4180000; 'netlist edit' = 142; 'netlist type' = 122; 'superposition load' = 1160000; 'life-x16 load' = 1370000 } },
    @{ Date = '2026-10-22'; Reason = ''; Floors = @{ 'column whole' = 103000000; 'column edit' = 778; 'chain whole' = 63000000; 'chain edit' = 667; 'sums whole' = 29100000; 'sums edit' = 684; 'life whole' = 30800000; 'life edit' = 364; 'lookups whole' = 12600000; 'lookups edit' = 81; 'criteria whole' = 4960000; 'criteria edit' = 152; 'netlist whole' = 4250000; 'netlist edit' = 142; 'netlist type' = 242000; 'superposition load' = 1160000; 'life-x16 load' = 1370000 } },
    @{ Date = '2026-10-22'; Reason = ''; Floors = @{ 'column whole' = 103000000; 'column edit' = 778; 'chain whole' = 63000000; 'chain edit' = 667; 'sums whole' = 29500000; 'sums edit' = 684; 'life whole' = 30800000; 'life edit' = 364; 'lookups whole' = 14000000; 'lookups edit' = 90; 'criteria whole' = 9400000; 'criteria edit' = 218; 'netlist whole' = 4270000; 'netlist edit' = 155; 'netlist type' = 274000; 'superposition load' = 1670000; 'life-x16 load' = 2660000 } },
    @{ Date = '2026-10-23'; Reason = ''; Floors = @{ 'column whole' = 103000000; 'column edit' = 778; 'chain whole' = 63000000; 'chain edit' = 667; 'sums whole' = 29500000; 'sums edit' = 684; 'life whole' = 30800000; 'life edit' = 367; 'lookups whole' = 14000000; 'lookups edit' = 90; 'criteria whole' = 9570000; 'criteria edit' = 218; 'netlist whole' = 4310000; 'netlist edit' = 155; 'netlist type' = 303000; 'superposition load' = 38100000; 'life-x16 load' = 42200000 } },
    @{ Date = '2026-10-30'; Reason = ''; Floors = @{ 'column whole' = 103000000; 'column edit' = 778; 'chain whole' = 63000000; 'chain edit' = 667; 'sums whole' = 29500000; 'sums edit' = 684; 'life whole' = 30800000; 'life edit' = 367; 'lookups whole' = 14000000; 'lookups edit' = 90; 'criteria whole' = 9570000; 'criteria edit' = 218; 'netlist whole' = 4310000; 'netlist edit' = 155; 'netlist type' = 303000; 'superposition load' = 52000000; 'life-x16 load' = 56100000; 'superposition-styled load' = 52700000 } },
    @{ Date = '2026-11-01'; Reason = 'KERNEL.69, the rewrite''s two floors joining, and the owner''s ruling of 2026-11-01 on the floor ''netlist type'', KERNEL.67''s third finding: 303000 was set from one run of 2.8 microseconds, and the same code measured 3.4 to 4.9 after it. This run''s 3.1 clears it by 6%, and the floor is set as every floor is, 15% under the run; no other floor moved down (docs/BETA_REARVIEW.md, KERNEL.69)'; Floors = @{ 'column whole' = 120000000; 'column edit' = 778; 'chain whole' = 63000000; 'chain edit' = 667; 'sums whole' = 30800000; 'sums edit' = 684; 'life whole' = 30800000; 'life edit' = 367; 'lookups whole' = 16300000; 'lookups edit' = 92; 'criteria whole' = 42000000; 'criteria edit' = 218; 'netlist whole' = 10700000; 'netlist edit' = 170; 'netlist type' = 274000; 'netlist rewrite' = 149000; 'netlist rewrite write' = 28; 'superposition load' = 55000000; 'life-x16 load' = 60100000; 'superposition-styled load' = 57700000 } }
)

# --- the ceilings, a dated history that never rises (KERNEL.49) ---
# One entry a measurement, one a line inside the parentheses, a comma after
# the entry before it: @{ Date = 'yyyy-mm-dd'; Reason = ''; Ceilings = @{
# 'superposition held' = n; 'superposition peak' = n; ... } }, each fixture's
# bytes held after its load and two frames, and its bytes at the peak, the
# memory runner's exactly, and since KERNEL.72 its save's four, 'save', 'save
# peak', 'saved held' and 'saved peak'. An entry higher than the one before it
# in any ceiling needs its Reason.
$ceilingHistory = @(
    @{ Date = '2026-10-21'; Reason = ''; Ceilings = @{ 'column held' = 29102875; 'column peak' = 62909797; 'chain held' = 15909993; 'chain peak' = 37447954; 'sums held' = 11888295; 'sums peak' = 31769224; 'life held' = 55883911; 'life peak' = 283714012; 'lookups held' = 77776017; 'lookups peak' = 155549356; 'criteria held' = 28977401; 'criteria peak' = 51926531; 'superposition held' = 320715071; 'superposition peak' = 338673647; 'life-x16 held' = 405270455; 'life-x16 peak' = 480767927 } },
    @{ Date = '2026-10-21'; Reason = ''; Ceilings = @{ 'column held' = 29102875; 'column peak' = 62909797; 'chain held' = 15909993; 'chain peak' = 37447954; 'sums held' = 11888295; 'sums peak' = 31769224; 'life held' = 55883911; 'life peak' = 283714012; 'lookups held' = 77776017; 'lookups peak' = 155549356; 'criteria held' = 28977401; 'criteria peak' = 51926531; 'netlist held' = 69154762; 'netlist peak' = 96352009; 'superposition held' = 320715071; 'superposition peak' = 338673647; 'life-x16 held' = 405270455; 'life-x16 peak' = 480767927 } },
    @{ Date = '2026-10-21'; Reason = 'KERNEL.54''s second slice, at the owner''s word: every program''s bytes held, and four programs'' peaks, raised by what the slice keeps, each shape''s stretches of the order (8 bytes each and 4 a shape), the frame''s lists given room at the plan so that no frame allocates (4 bytes a shape made whole, 4 a shape with parts, 4 a word of the marks) and the netlist''s 992 pieces; held, the column 3,808 bytes more, the chain 2,556, the sums 1,828, Life 3,984, the lookups 8,868, the criteria 3,952, the netlist 393,908 (0.57%), the board 50,664 and Life x16 63,792 (0.016%); at the peak the criteria 192 more, the netlist, the board and Life x16 as much as they hold more, the others the same (docs/BETA_REARVIEW.md, KERNEL.54)'; Ceilings = @{ 'column held' = 29106683; 'column peak' = 62909797; 'chain held' = 15912549; 'chain peak' = 37447954; 'sums held' = 11890123; 'sums peak' = 31769224; 'life held' = 55887895; 'life peak' = 283714012; 'lookups held' = 77784885; 'lookups peak' = 155549356; 'criteria held' = 28981353; 'criteria peak' = 51926723; 'netlist held' = 69548670; 'netlist peak' = 96745917; 'superposition held' = 320765735; 'superposition peak' = 338724311; 'life-x16 held' = 405334247; 'life-x16 peak' = 480831719 } },
    @{ Date = '2026-10-21'; Reason = 'KERNEL.50''s first slice, at the owner''s word: every program''s bytes held, and the criteria''s, the netlist''s, the board''s and Life x16''s peaks, raised by a format''s appearance, 32 bytes more a slot of a grid''s style table (a Style 56 bytes where it was 24, its font, alignment and four edges each said or not); held, every program and Life x16 32 bytes more, a table of one slot, and the board 512, its 13 formats in 16 slots; at the peak the criteria, the netlist, the board and Life x16 as much as they hold more, the others the same (docs/BETA_REARVIEW.md, KERNEL.50)'; Ceilings = @{ 'column held' = 29106715; 'column peak' = 62909797; 'chain held' = 15912581; 'chain peak' = 37447954; 'sums held' = 11890155; 'sums peak' = 31769224; 'life held' = 55887927; 'life peak' = 283714012; 'lookups held' = 77784917; 'lookups peak' = 155549356; 'criteria held' = 28981385; 'criteria peak' = 51926755; 'netlist held' = 69548702; 'netlist peak' = 96745949; 'superposition held' = 320766247; 'superposition peak' = 338724823; 'life-x16 held' = 405334279; 'life-x16 peak' = 480831751 } },
    @{ Date = '2026-10-22'; Reason = ''; Ceilings = @{ 'column held' = 17220986; 'column peak' = 62909797; 'chain held' = 10674075; 'chain peak' = 37447954; 'sums held' = 7144895; 'sums peak' = 31769224; 'life held' = 47510443; 'life peak' = 283714012; 'lookups held' = 45870596; 'lookups peak' = 155549356; 'criteria held' = 22549734; 'criteria peak' = 51926639; 'netlist held' = 62156189; 'netlist peak' = 90269686; 'superposition held' = 188069711; 'superposition peak' = 331719923; 'life-x16 held' = 269575847; 'life-x16 peak' = 430814741 } },
    @{ Date = '2026-10-22'; Reason = 'KERNEL.56''s first slice, at the owner''s word: the netlist''s bytes held and its peak raised by its program, two lines longer (the bench''s type row and a line of its comment), each line a text cell of the Frazaro sheet a build writes; held 1,222 bytes more and at the peak 2,026, the other programs the same (docs/BETA_REARVIEW.md, KERNEL.56)'; Ceilings = @{ 'column held' = 17220986; 'column peak' = 62909797; 'chain held' = 10674075; 'chain peak' = 37447954; 'sums held' = 7144895; 'sums peak' = 31769224; 'life held' = 47510443; 'life peak' = 283714012; 'lookups held' = 45870596; 'lookups peak' = 155549356; 'criteria held' = 22549734; 'criteria peak' = 51926639; 'netlist held' = 62157411; 'netlist peak' = 90271712; 'superposition held' = 188069711; 'superposition peak' = 331719923; 'life-x16 held' = 269575847; 'life-x16 peak' = 430814741 } },
    @{ Date = '2026-10-22'; Reason = 'KERNEL.56''s second slice, at the owner''s word: every program''s bytes held, and four programs'' peaks, raised by what the slice keeps, each rows'' sheet''s rectangle of the twin cells a frame gained (20 bytes a sheet), the list of sheets whose extent moved, given room at the first frame''s end that grew a twin (32 bytes, the board''s 128), and the lookups'' four readers of an extent (32 bytes); held, a program of two sheets 72 bytes more, Life 92, the lookups 104, Life x16 152 and the board 448; at the peak the criteria, the netlist, the board and Life x16 24 more, the others the same (docs/BETA_REARVIEW.md, KERNEL.56)'; Ceilings = @{ 'column held' = 17221058; 'column peak' = 62909797; 'chain held' = 10674147; 'chain peak' = 37447954; 'sums held' = 7144967; 'sums peak' = 31769224; 'life held' = 47510535; 'life peak' = 283714012; 'lookups held' = 45870700; 'lookups peak' = 155549356; 'criteria held' = 22549806; 'criteria peak' = 51926663; 'netlist held' = 62157483; 'netlist peak' = 90271736; 'superposition held' = 188070159; 'superposition peak' = 331719947; 'life-x16 held' = 269575999; 'life-x16 peak' = 430814765 } },
    @{ Date = '2026-10-22'; Reason = ''; Ceilings = @{ 'column held' = 13697921; 'column peak' = 62909797; 'chain held' = 8325463; 'chain peak' = 37447954; 'sums held' = 5499027; 'sums peak' = 31769224; 'life held' = 43811180; 'life peak' = 283714012; 'lookups held' = 37652887; 'lookups peak' = 155549356; 'criteria held' = 18944123; 'criteria peak' = 50294479; 'netlist held' = 53238277; 'netlist peak' = 86925008; 'superposition held' = 76952156; 'superposition peak' = 194408556; 'life-x16 held' = 129893031; 'life-x16 peak' = 238717605 } },
    @{ Date = '2026-10-23'; Reason = ''; Ceilings = @{ 'column held' = 10463980; 'column peak' = 62909797; 'chain held' = 7064302; 'chain peak' = 37447954; 'sums held' = 4637011; 'sums peak' = 31769224; 'life held' = 40624685; 'life peak' = 283714012; 'lookups held' = 29487670; 'lookups peak' = 155549356; 'criteria held' = 12209986; 'criteria peak' = 41898740; 'netlist held' = 53186144; 'netlist peak' = 85038276; 'superposition held' = 10803728; 'superposition peak' = 10804168; 'life-x16 held' = 18975068; 'life-x16 peak' = 18975068 } },
    @{ Date = '2026-10-30'; Reason = 'KERNEL.67, at the owner''s word: a style row over a range held as its rectangle, a page''s formats a word where they were three, boxed and made at a sheet''s first format; held, every sheet and its twin 16 bytes fewer, a program of two sheets 64, Life 96, Life x16 192 and the board 512; at the peak the board and Life x16 as much, the criteria and the netlist 112 fewer, and the column, the chain, the sums, Life and the lookups the same, their peaks coming before the draft holds a sheet; the third fixture of rows joins 520 bytes over its twin at every point (docs/BETA_REARVIEW.md, KERNEL.67)'; Ceilings = @{ 'column held' = 10463916; 'column peak' = 62909797; 'chain held' = 7064238; 'chain peak' = 37447954; 'sums held' = 4636947; 'sums peak' = 31769224; 'life held' = 40624589; 'life peak' = 283714012; 'lookups held' = 29487606; 'lookups peak' = 155549356; 'criteria held' = 12209922; 'criteria peak' = 41898628; 'netlist held' = 53186080; 'netlist peak' = 85038164; 'superposition held' = 10803216; 'superposition peak' = 10803656; 'life-x16 held' = 18974876; 'life-x16 peak' = 18974876; 'superposition-styled held' = 10803736; 'superposition-styled peak' = 10804176 } },
    @{ Date = '2026-10-30'; Reason = 'KERNEL.58''s second slice, at the owner''s word: a format''s family a pointer in its table''s slot, a Style 64 bytes where 56, so every fixture''s bytes held and at the peak up by 8 a slot, Superposition''s 16 slots 128 and a fixture of one format 8; no frame or plan moves. The memory runner''s bytes, exact on every 64-bit machine; the rows'' timings stay the run''s of 2026-10-30.'; Ceilings = @{ 'column held' = 10463924; 'column peak' = 62909797; 'chain held' = 7064246; 'chain peak' = 37447954; 'sums held' = 4636955; 'sums peak' = 31769224; 'life held' = 40624597; 'life peak' = 283714012; 'lookups held' = 29487614; 'lookups peak' = 155549356; 'criteria held' = 12209930; 'criteria peak' = 41898644; 'netlist held' = 53186088; 'netlist peak' = 85038180; 'superposition held' = 10803344; 'superposition peak' = 10803784; 'life-x16 held' = 18974884; 'life-x16 peak' = 18974884; 'superposition-styled held' = 10803864; 'superposition-styled peak' = 10804304 } },
    @{ Date = '2026-11-01'; Reason = 'KERNEL.69, at the owner''s word: the netlist''s two ceilings alone, 178 bytes held and 848 at the peak, by the two comment lines its program gains, 130 characters carried with it into what the runner builds. Measured: the program before them over this code holds 53186088 and peaks at 85038180, the ceilings before, so that the code adds nothing, and no other fixture moved (docs/BETA_REARVIEW.md, KERNEL.69)'; Ceilings = @{ 'column held' = 10463924; 'column peak' = 62909797; 'chain held' = 7064246; 'chain peak' = 37447954; 'sums held' = 4636955; 'sums peak' = 31769224; 'life held' = 40624597; 'life peak' = 283714012; 'lookups held' = 29487614; 'lookups peak' = 155549356; 'criteria held' = 12209930; 'criteria peak' = 41898644; 'netlist held' = 53186266; 'netlist peak' = 85039028; 'superposition held' = 10803344; 'superposition peak' = 10803784; 'life-x16 held' = 18974884; 'life-x16 peak' = 18974884; 'superposition-styled held' = 10803864; 'superposition-styled peak' = 10804304 } },
    @{ Date = '2026-11-01'; Reason = 'KERNEL.58''s third slice, at the owner''s word: a sheet''s conditional formats a word, boxed and made at its first, so every fixture''s bytes held up by 8 a sheet, twins among them, Superposition''s 32 sheets 256 and a program of two sheets 32, and the peaks of criteria and netlist, a load''s, up by 248, each draft sheet holding two maps more; a grid holding no rule moves no plan and no frame. The memory runner''s own bytes, exact on any 64-bit machine.'; Ceilings = @{ 'column held' = 10463956; 'column peak' = 62909797; 'chain held' = 7064278; 'chain peak' = 37447954; 'sums held' = 4636987; 'sums peak' = 31769224; 'life held' = 40624645; 'life peak' = 283714012; 'lookups held' = 29487646; 'lookups peak' = 155549356; 'criteria held' = 12209962; 'criteria peak' = 41898892; 'netlist held' = 53186298; 'netlist peak' = 85039276; 'superposition held' = 10803600; 'superposition peak' = 10804040; 'life-x16 held' = 18974980; 'life-x16 peak' = 18974980; 'superposition-styled held' = 10804120; 'superposition-styled peak' = 10804560 } },
    @{ Date = '2026-11-01'; Reason = 'KERNEL.72, at the owner''s word: the save''s shape. A load reads its rows one form at a time and holds none past its row, so that the seven programs'' peaks fall, the read by forms gone from them, 283.7 MB to 53.6 for Life, 155.5 to 74.4 for the lookups, 62.9 to 23.1 for the column; no fixture''s bytes held move, and the fixtures of rows, whose peaks were the machine''s, stand. Each fixture''s save joins, its text''s bytes, the most its printing holds over the machine, and a load from it held and at the peak: Superposition''s 29.7 MB and 526.6 at the peak over `c1c03a1` are 2.4 and 10.6, and the styled board''s 36.7 and 573.5 stand within 4,096 bytes of its twin. The memory runner''s bytes, exact on every 64-bit machine; the rows'' timings stay the run''s of 2026-11-01.'; Ceilings = @{ 'column held' = 10463956; 'column peak' = 23110916; 'column save' = 3261708; 'column save peak' = 9553164; 'column saved held' = 10437445; 'column saved peak' = 23602317; 'chain held' = 7064278; 'chain peak' = 20219356; 'chain save' = 1963032; 'chain save peak' = 5855654; 'chain saved held' = 7030439; 'chain saved peak' = 20546575; 'sums held' = 4636987; 'sums peak' = 11858469; 'sums save' = 1823353; 'sums save peak' = 4969081; 'sums saved held' = 4603148; 'sums saved peak' = 12086670; 'life held' = 40624645; 'life peak' = 53634545; 'life save' = 32623420; 'life save peak' = 74173244; 'life saved held' = 40488692; 'life saved peak' = 54175816; 'lookups held' = 29487646; 'lookups peak' = 74421446; 'lookups save' = 8901488; 'lookups save peak' = 21484400; 'lookups saved held' = 25956933; 'lookups saved peak' = 75556497; 'criteria held' = 12209962; 'criteria peak' = 23344415; 'criteria save' = 3372714; 'criteria save peak' = 9664170; 'criteria saved held' = 10142305; 'criteria saved peak' = 23852392; 'netlist held' = 53186298; 'netlist peak' = 63014223; 'netlist save' = 2530377; 'netlist save peak' = 6105357; 'netlist saved held' = 53185253; 'netlist saved peak' = 60515836; 'superposition held' = 10803600; 'superposition peak' = 10804040; 'superposition save' = 2372117; 'superposition save peak' = 4995415; 'superposition saved held' = 10583312; 'superposition saved peak' = 10583376; 'life-x16 held' = 18974980; 'life-x16 peak' = 18974980; 'life-x16 save' = 2051565; 'life-x16 save peak' = 4189031; 'life-x16 saved held' = 18840036; 'life-x16 saved peak' = 18840100; 'superposition-styled held' = 10804120; 'superposition-styled peak' = 10804560; 'superposition-styled save' = 2372233; 'superposition-styled save peak' = 4995647; 'superposition-styled saved held' = 10583832; 'superposition-styled saved peak' = 10583896 } }
)

# --- the programs measured, each one's SHA-256 over LF bytes ---
# 2026-10-10, KERNEL.20's sixth slice: the four programs as the slice wrote
# them, scripts/recalc/bench/<name>.txt; 2026-10-18, KERNEL.8's first slice:
# the fifth, the lookups; 2026-10-20, its fifth slice: the sixth, the
# criteria; 2026-10-21, KERNEL.54's first slice: the seventh, the netlist,
# written by core/examples/recalc.rs --netlist and held to it by the runner's
# check mode; 2026-11-01, KERNEL.69: the netlist again, its rewrite's two comment
# lines.
$programs = [ordered]@{
    'column' = '751885133C249C25C8C4BB3854877A72E6054B2E0945BB2C1931C3FF05D55E2E'
    'chain'  = '72FA6E5DB80DF71796473E21B89432C598AD8719C473D8955B24B796417A8A5A'
    'sums'   = 'A339D809DE2493BBA3A129FC3FC4A5AEC458A5A6F1B5FB590DB04D2CCB3F7348'
    'life'   = 'C2B7AC34005AB88CFE2F4CE98FAC3C704CCF26E8AEED90BE221BF800D63F1716'
    'lookups' = '5D25CCE0460063145619F43B64ED898BA6A073F5F83003D53B92B8D1C8C1619A'
    'criteria' = 'B80A36C00F9445793B061A2CCAA2138F7B0FC407469A87EC6259F1A878C48D63'
    'netlist' = 'CD618419C5220021E55F9DDEE66699DDCBE9E384E2FFCC150A2C86FC56741C26'
}
# 2026-10-21, KERNEL.49's first slice: the two fixtures of rows the memory
# runner loads as they stand, scripts/recalc/bench/<name>.vla; 2026-10-30,
# KERNEL.67: the third, which the runner loads after its twin's rows, its own
# file pinned here and its twin's by its twin's line.
$rowsFixtures = [ordered]@{
    'superposition' = '6A3AE552F614AC0CD5582D69ED0992E804032831798E92C0445EEFA478551B98'
    'life-x16'      = '7366BE8061601852C8447E246CA98E58D11DF699A965B659BB48826975ED001D'
    'superposition-styled' = '67097ADEFC12CFDB5A0C99F564A239AAF811FD7E6B0005CE1DED53AB411024B3'
}
# A fixture of rows that adds to a twin's, held within a constant of it after
# its load, after two frames and at the peak (KERNEL.67): one style row over
# Superposition's Field holds its rectangle and never a format a cell, 520
# bytes at its first measurement where a format a cell held 6.9 MB more and
# 15.2 MB more at the peak. The bound is a page, not the measured bytes: the
# ceilings hold those exactly.
$twins = [ordered]@{ 'superposition-styled' = 'superposition' }
$twinBound = 4096
$twinPoints = [ordered]@{ LoadHeld = 'after its load'; Held = 'after two frames'; Peak = 'at the peak'; Save = 'in its save'; SavePeak = 'printing its save'; SavedHeld = 'after a load from its save'; SavedPeak = 'at the peak of a load from its save' }
# What a save costs (KERNEL.72), each the rule of the item's title and no pinned
# number, the ceilings holding those. Measured 2026-11-01, where it stood before
# the item: a fixture of rows' save 2.01 to 2.97 bytes a formula cell (27.3 to
# 45.9); printing a save 2.04 to 2.98 times its text (1.85 to 3.92); a program's
# load from its save 1.14 to 2.92 times what it holds (1.56 to 6.79); a fixture of
# rows' load from its save 64 bytes over what it holds (516 MB over Superposition's).
$saveBytesAFormula = 4
$printingTimes = 4
$savedPeakTimes = 4
# The ceilings of a memory row: each ceiling's word, the row's field, and how a
# line reads it.
$ceilingFields = [ordered]@{
    'held' = @('Held', 'bytes held')
    'peak' = @('Peak', 'bytes at the peak')
    'save' = @('Save', 'bytes in its save')
    'save peak' = @('SavePeak', 'bytes over the machine, printing its save')
    'saved held' = @('SavedHeld', 'bytes held after a load from its save')
    'saved peak' = @('SavedPeak', 'bytes at the peak of a load from its save')
}

# --- the rules ---
$fixtures = @('column', 'chain', 'sums', 'life', 'lookups', 'criteria', 'netlist')
# The fixtures whose programs name a type (KERNEL.56), each measured typing a
# value into a fresh cell beside its edit, in both halves.
$typed = @('netlist')
# The fixtures whose programs name a rewrite (KERNEL.69), each measured writing
# a cell's formula over and back beside its edit, in both halves.
$rewritten = @('netlist')
# The memory runner's fixtures, every program and the fixtures of rows, each
# weighed; the fixtures of rows' loads floored (KERNEL.49).
$weighed = @($fixtures) + @($rowsFixtures.Keys)
$loadFloored = @($rowsFixtures.Keys)
$modes = @('all', 'one')
$flagFiles = @('calcpr_built', 'calcpr_calcid', 'calcpr_none')
$verdicts = @('kept the folded value', 'computed on opening', 'neither')
$minRounds = 3
$clockMax = 10.0
$agreeWithin = 0.005

function Get-LfDigest([string]$path) {
    $text = [System.IO.File]::ReadAllText($path) -replace "`r`n", "`n"
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($text)) } finally { $sha.Dispose() }
    return (($hash | ForEach-Object { $_.ToString('X2') }) -join '')
}

function Test-Near([double]$a, [double]$b) {
    if ($b -eq 0) { return ($a -eq 0) }
    return ([math]::Abs($a / $b - 1.0) -le $agreeWithin)
}

# The whole judgment of a baseline, its history and the programs as they
# stand; the problems, one a line.
function Test-Floors($date, $native, $excel, $rows, $history, $measured, $now) {
    $problems = New-Object System.Collections.Generic.List[string]
    $rows = @($rows | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0 -or $date -eq '') {
        $problems.Add('the baseline is empty: run tools\bench_recalc.ps1 and tools\bench_recalc_excel.ps1 on the owner''s machine and paste both blocks')
        return $problems.ToArray()
    }
    # The environments.
    if ($native -notmatch 'power throttling opted out: the runner yes') { $problems.Add('the native half did not run opted out of Windows'' power throttling: its environment line does not say so') }
    if ($native -notmatch 'high priority: the runner yes') { $problems.Add('the native half did not run at high priority: its environment line does not say so') }
    if ($native -notmatch 'on the performance cores, .*: the runner yes' -and $native -notmatch 'one kind of core') { $problems.Add('the native half did not run on the performance cores: its environment line does not say so') }
    if ($excel -notmatch 'power throttling opted out: yes') { $problems.Add('Excel''s half did not run opted out of Windows'' power throttling: its environment line does not say so') }
    if ($excel -match 'STAND-IN' -or $excel -eq '') { $problems.Add('Excel''s half is the instrument''s stand-in, or missing: no Excel was measured') }
    # The native rows.
    $nat = @{}
    foreach ($r in @($rows | Where-Object { $_.Half -eq 'native' })) {
        if ($nat.ContainsKey($r.Fixture)) { $problems.Add("native $($r.Fixture): two rows"); continue }
        $nat[$r.Fixture] = $r
    }
    foreach ($f in $fixtures) {
        if (-not $nat.ContainsKey($f)) { $problems.Add("native $($f): no row"); continue }
        $r = $nat[$f]
        if ([int]$r.Rounds -lt $minRounds) { $problems.Add("native $($f): measured in $($r.Rounds) round(s), fewer than $minRounds") }
        if ([math]::Abs([double]$r.ClockSpread) -gt $clockMax -or [math]::Abs([double]$r.ClockDrift) -gt $clockMax) { $problems.Add("native $($f): the clock moved, spread $($r.ClockSpread)% and drift $($r.ClockDrift)%, more than $clockMax%") }
        if (-not (Test-Near ([double]$r.WholePerSecond) ([double]$r.Formulas / ([double]$r.Whole / 1000.0)))) { $problems.Add("native $($f): $($r.WholePerSecond) formulas a second is not $($r.Formulas) formulas in $($r.Whole) ms") }
        if (-not (Test-Near ([double]$r.EditsPerSecond) (1000.0 / [double]$r.Edit))) { $problems.Add("native $($f): $($r.EditsPerSecond) edits a second is not one in $($r.Edit) ms") }
        if ($typed -contains $f) {
            # A type prints in milliseconds to three places, so its time and its
            # types a second agree within half that last place beside the
            # tolerance (KERNEL.56): a type of 3.5 us prints as 0.004.
            $typeOff = 0.0
            if ($r.ContainsKey('Type') -and $r.ContainsKey('TypesPerSecond')) { $typeOff = [math]::Abs(1000.0 / [double]$r.TypesPerSecond - [double]$r.Type) }
            if (-not $r.ContainsKey('Type') -or -not $r.ContainsKey('TypesPerSecond')) { $problems.Add("native $($f): its program names a type and its row measured none") }
            elseif ($typeOff -gt 0.0005 + [double]$r.Type * $agreeWithin) { $problems.Add("native $($f): $($r.TypesPerSecond) types a second is not one in $($r.Type) ms") }
        }
        if ($rewritten -contains $f) {
            # As a type's: milliseconds to three places, its frames microseconds.
            if (-not $r.ContainsKey('Rewrite') -or -not $r.ContainsKey('RewritesPerSecond') -or -not $r.ContainsKey('RewriteWrite') -or -not $r.ContainsKey('RewriteWritesPerSecond')) { $problems.Add("native $($f): its program names a rewrite and its row measured none") }
            else {
                $rewriteOff = [math]::Abs(1000.0 / [double]$r.RewritesPerSecond - [double]$r.Rewrite)
                if ($rewriteOff -gt 0.0005 + [double]$r.Rewrite * $agreeWithin) { $problems.Add("native $($f): $($r.RewritesPerSecond) rewrites a second is not one in $($r.Rewrite) ms") }
                $writeOff = [math]::Abs(1000.0 / [double]$r.RewriteWritesPerSecond - [double]$r.RewriteWrite)
                if ($writeOff -gt 0.0005 + [double]$r.RewriteWrite * $agreeWithin) { $problems.Add("native $($f): $($r.RewriteWritesPerSecond) rewrite writes a second is not one in $($r.RewriteWrite) ms") }
            }
        }
    }
    # Excel's rows: the same bytes, the same sums.
    $xl = @{}
    foreach ($r in @($rows | Where-Object { $_.Half -eq 'excel' })) {
        $k = "$($r.Fixture) $($r.Threads)"
        if ($xl.ContainsKey($k)) { $problems.Add("excel $($k): two rows"); continue }
        $xl[$k] = $r
    }
    foreach ($f in $fixtures) {
        foreach ($mode in $modes) {
            $k = "$f $mode"
            if (-not $xl.ContainsKey($k)) { $problems.Add("excel $($k): no row"); continue }
            $r = $xl[$k]
            if ($r.ContainsKey('Failed')) { $problems.Add("excel $($k): Excel could not measure it: $($r.Failed)"); continue }
            if ($typed -contains $f -and -not $r.ContainsKey('Type')) { $problems.Add("excel $($k): its program names a type and Excel's row measured none") }
            if ($rewritten -contains $f -and (-not $r.ContainsKey('Rewrite') -or -not $r.ContainsKey('RewriteWrite'))) { $problems.Add("excel $($k): its program names a rewrite and Excel's row measured none") }
            if (-not $nat.ContainsKey($f)) { continue }
            $n = $nat[$f]
            if ($r.Sha256 -ne $n.Sha256) { $problems.Add("excel $($k): measured $($r.Sha256), not the native half's file $($n.Sha256)") }
            if ([double]$r.SumBuilt -ne [double]$n.SumBuilt) { $problems.Add("excel $($k): summed $($r.SumBuilt) as built where the machine summed $($n.SumBuilt)") }
            if ([double]$r.SumEdited -ne [double]$n.SumEdited) { $problems.Add("excel $($k): summed $($r.SumEdited) edited where the machine summed $($n.SumEdited)") }
            if ($mode -eq 'one' -and [int]$r.ThreadCount -ne 1) { $problems.Add("excel $($k): $($r.ThreadCount) threads, not one") }
        }
    }
    # The flag.
    $flags = @($rows | Where-Object { $_.Half -eq 'excel-flag' })
    foreach ($file in $flagFiles) {
        $match = @($flags | Where-Object { $_.File -eq $file })
        if ($match.Count -ne 1) { $problems.Add("the flag: $($match.Count) row(s) for $file, not one"); continue }
        if ($verdicts -notcontains $match[0].Verdict) { $problems.Add("the flag: $file's verdict '$($match[0].Verdict)' is none of the three") }
    }
    # The programs.
    foreach ($f in $fixtures) {
        if (-not $now.Contains($f)) { $problems.Add("the program $($f): not found"); continue }
        if ($now[$f] -ne $measured[$f]) { $problems.Add("the program $($f): its SHA-256 over LF is $($now[$f]), not the measured $($measured[$f]); measure it again") }
    }
    # The floors.
    $history = @($history | Where-Object { $null -ne $_ })
    if ($history.Count -eq 0) { $problems.Add('the floor history is empty'); return $problems.ToArray() }
    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($f in $fixtures) { $keys.Add("$f whole"); $keys.Add("$f edit") }
    # A typed fixture's types a second (KERNEL.56).
    foreach ($f in $typed) { $keys.Add("$f type") }
    # A rewritten fixture's frames and writes a second (KERNEL.69).
    foreach ($f in $rewritten) { $keys.Add("$f rewrite"); $keys.Add("$f rewrite write") }
    # The fixtures of rows' loads, formulas loaded a second (KERNEL.49).
    foreach ($f in $loadFloored) { $keys.Add("$f load") }
    # A fixture's floors from the entry it joined the history on, every later
    # entry holding them and the latest every fixture's (KERNEL.8: the fifth
    # joined after five entries held four, which no paste may edit).
    $joinedAt = @{}
    foreach ($k in $keys) {
        for ($j = 0; $j -lt $history.Count; $j++) {
            if ($null -ne $history[$j].Floors -and $history[$j].Floors.ContainsKey($k)) { $joinedAt[$k] = $j; break }
        }
    }
    for ($i = 0; $i -lt $history.Count; $i++) {
        $e = $history[$i]
        foreach ($k in $keys) {
            $before = $i -lt $history.Count - 1 -and (-not $joinedAt.ContainsKey($k) -or $i -lt $joinedAt[$k])
            if ($before) { continue }
            if ($null -eq $e.Floors -or -not $e.Floors.ContainsKey($k)) { $problems.Add("the floor history's entry of $($e.Date): no floor '$k'"); continue }
            if ($i -gt 0 -and $history[$i - 1].Floors.ContainsKey($k) -and [double]$e.Floors[$k] -lt [double]$history[$i - 1].Floors[$k] -and [string]::IsNullOrWhiteSpace($e.Reason)) {
                $problems.Add("the floor history's entry of $($e.Date): '$k' is $($e.Floors[$k]), under the $($history[$i - 1].Floors[$k]) before it, with no reason written")
            }
        }
    }
    $latest = $history[$history.Count - 1]
    foreach ($f in $fixtures) {
        if (-not $nat.ContainsKey($f)) { continue }
        $n = $nat[$f]
        if ($latest.Floors.ContainsKey("$f whole") -and [double]$n.WholePerSecond -lt [double]$latest.Floors["$f whole"]) { $problems.Add("native $($f): $($n.WholePerSecond) formulas a second, under the floor of $($latest.Floors["$f whole"])") }
        if ($latest.Floors.ContainsKey("$f edit") -and [double]$n.EditsPerSecond -lt [double]$latest.Floors["$f edit"]) { $problems.Add("native $($f): $($n.EditsPerSecond) edits a second, under the floor of $($latest.Floors["$f edit"])") }
        if ($latest.Floors.ContainsKey("$f type") -and $n.ContainsKey('TypesPerSecond') -and [double]$n.TypesPerSecond -lt [double]$latest.Floors["$f type"]) { $problems.Add("native $($f): $($n.TypesPerSecond) types a second, under the floor of $($latest.Floors["$f type"])") }
        if ($latest.Floors.ContainsKey("$f rewrite") -and $n.ContainsKey('RewritesPerSecond') -and [double]$n.RewritesPerSecond -lt [double]$latest.Floors["$f rewrite"]) { $problems.Add("native $($f): $($n.RewritesPerSecond) rewrites a second, under the floor of $($latest.Floors["$f rewrite"])") }
        if ($latest.Floors.ContainsKey("$f rewrite write") -and $n.ContainsKey('RewriteWritesPerSecond') -and [double]$n.RewriteWritesPerSecond -lt [double]$latest.Floors["$f rewrite write"]) { $problems.Add("native $($f): $($n.RewriteWritesPerSecond) rewrite writes a second, under the floor of $($latest.Floors["$f rewrite write"])") }
    }
    foreach ($f in $loadFloored) {
        $m = @($rows | Where-Object { $_.Half -eq 'memory' -and $_.Fixture -eq $f })
        if ($m.Count -ne 1) { continue }
        if ($latest.Floors.ContainsKey("$f load") -and [double]$m[0].LoadPerSecond -lt [double]$latest.Floors["$f load"]) { $problems.Add("memory $($f): $($m[0].LoadPerSecond) formulas loaded a second, under the floor of $($latest.Floors["$f load"])") }
    }
    return $problems.ToArray()
}

# The memory half (KERNEL.49): its rows, the sources they weighed, and the
# ceilings; the problems, one a line. The load floors are Test-Floors'.
# What a save costs (KERNEL.72), of a memory row or a line of the memory
# runner's check mode: each rule's words, where it does not hold.
function Get-SaveProblems([string]$who, [string]$f, $s) {
    $out = New-Object System.Collections.Generic.List[string]
    if ([long]$s.Save -le 0 -or [long]$s.SavedHeld -le 0) { $out.Add("$who $($f): no save weighed") }
    if ([long]$s.SavedPeak -lt [long]$s.SavedHeld) { $out.Add("$who $($f): a load from its save peaks at $($s.SavedPeak) bytes, under the $($s.SavedHeld) it holds") }
    if ([long]$s.SavedHeld -gt [long]$s.Held) { $out.Add("$who $($f): a load from its save holds $($s.SavedHeld) bytes, more than the $($s.Held) the machine that printed it held") }
    if ([long]$s.SavePeak -gt $printingTimes * [long]$s.Save) { $out.Add("$who $($f): printing its save of $($s.Save) bytes held $($s.SavePeak) over the machine, more than $printingTimes times its text") }
    if ($rowsFixtures.Contains($f)) {
        if ([long]$s.Save -gt $saveBytesAFormula * [long]$s.Formulas) { $out.Add("$who $($f): its save of $($s.Save) bytes is more than $saveBytesAFormula bytes a formula cell of its $($s.Formulas)") }
        if ([long]$s.SavedPeak -gt [long]$s.SavedHeld + $twinBound) { $out.Add("$who $($f): a load from its save peaks at $($s.SavedPeak) bytes, more than $twinBound over the $($s.SavedHeld) it holds") }
    }
    elseif ([long]$s.SavedPeak -gt $savedPeakTimes * [long]$s.SavedHeld) { $out.Add("$who $($f): a load from its save peaks at $($s.SavedPeak) bytes, more than $savedPeakTimes times the $($s.SavedHeld) it holds") }
    return $out.ToArray()
}

function Test-Memory($rows, $ceilings, $measured, $rowsMeasured, $rowsNow) {
    $problems = New-Object System.Collections.Generic.List[string]
    $rows = @($rows | Where-Object { $null -ne $_ })
    $mem = @{}
    foreach ($r in @($rows | Where-Object { $_.Half -eq 'memory' })) {
        if ($mem.ContainsKey($r.Fixture)) { $problems.Add("memory $($r.Fixture): two rows"); continue }
        $mem[$r.Fixture] = $r
    }
    foreach ($f in $weighed) {
        if (-not $mem.ContainsKey($f)) { $problems.Add("memory $($f): no row"); continue }
        $r = $mem[$f]
        if ([long]$r.Formulas -le 0) { $problems.Add("memory $($f): no formula weighed") }
        if ([long]$r.LoadHeld -le 0 -or [long]$r.Held -le 0) { $problems.Add("memory $($f): no bytes held") }
        if ([long]$r.Peak -lt [long]$r.Held -or [long]$r.Peak -lt [long]$r.LoadHeld) { $problems.Add("memory $($f): a peak of $($r.Peak) bytes, under the $($r.Held) held after two frames or the $($r.LoadHeld) after the load") }
        if ([int]$r.Rounds -lt $minRounds) { $problems.Add("memory $($f): its load timed in $($r.Rounds) round(s), fewer than $minRounds") }
        if ([math]::Abs([double]$r.ClockSpread) -gt $clockMax -or [math]::Abs([double]$r.ClockDrift) -gt $clockMax) { $problems.Add("memory $($f): the clock moved, spread $($r.ClockSpread)% and drift $($r.ClockDrift)%, more than $clockMax%") }
        if (-not (Test-Near ([double]$r.LoadPerSecond) ([double]$r.Formulas / ([double]$r.Load / 1000.0)))) { $problems.Add("memory $($f): $($r.LoadPerSecond) formulas loaded a second is not $($r.Formulas) formulas in $($r.Load) ms") }
        $pinned = if ($rowsMeasured.Contains($f)) { $rowsMeasured[$f] } elseif ($measured.Contains($f)) { $measured[$f] } else { '' }
        if ($r.Sha256 -ne $pinned) { $problems.Add("memory $($f): weighed $($r.Sha256), not the source pinned here, $pinned; measure it again") }
        foreach ($p in @(Get-SaveProblems 'memory' $f $r)) { $problems.Add($p) }
    }
    # A fixture that adds to a twin's, within the bound of it (KERNEL.67).
    foreach ($f in $twins.Keys) {
        $t = $twins[$f]
        if (-not $mem.ContainsKey($f) -or -not $mem.ContainsKey($t)) { continue }
        foreach ($point in $twinPoints.Keys) {
            $more = [long]$mem[$f][$point] - [long]$mem[$t][$point]
            if ($more -gt $twinBound) { $problems.Add("memory $($f): $more bytes more than its twin $t $($twinPoints[$point]), over the bound of $twinBound") }
        }
    }
    # The fixtures of rows as they stand.
    foreach ($f in $rowsMeasured.Keys) {
        if (-not $rowsNow.Contains($f)) { $problems.Add("the fixture $($f): not found"); continue }
        if ($rowsNow[$f] -ne $rowsMeasured[$f]) { $problems.Add("the fixture $($f): its SHA-256 over LF is $($rowsNow[$f]), not the measured $($rowsMeasured[$f]); measure it again") }
    }
    # The ceilings: a fixture's from the entry it joined on, never raised
    # with no reason written, every fixture's in the latest.
    $ceilings = @($ceilings | Where-Object { $null -ne $_ })
    if ($ceilings.Count -eq 0) { $problems.Add('the ceiling history is empty'); return $problems.ToArray() }
    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($f in $weighed) { foreach ($what in $ceilingFields.Keys) { $keys.Add("$f $what") } }
    $joinedAt = @{}
    foreach ($k in $keys) {
        for ($j = 0; $j -lt $ceilings.Count; $j++) {
            if ($null -ne $ceilings[$j].Ceilings -and $ceilings[$j].Ceilings.ContainsKey($k)) { $joinedAt[$k] = $j; break }
        }
    }
    for ($i = 0; $i -lt $ceilings.Count; $i++) {
        $e = $ceilings[$i]
        foreach ($k in $keys) {
            $before = $i -lt $ceilings.Count - 1 -and (-not $joinedAt.ContainsKey($k) -or $i -lt $joinedAt[$k])
            if ($before) { continue }
            if ($null -eq $e.Ceilings -or -not $e.Ceilings.ContainsKey($k)) { $problems.Add("the ceiling history's entry of $($e.Date): no ceiling '$k'"); continue }
            if ($i -gt 0 -and $ceilings[$i - 1].Ceilings.ContainsKey($k) -and [double]$e.Ceilings[$k] -gt [double]$ceilings[$i - 1].Ceilings[$k] -and [string]::IsNullOrWhiteSpace($e.Reason)) {
                $problems.Add("the ceiling history's entry of $($e.Date): '$k' is $($e.Ceilings[$k]), over the $($ceilings[$i - 1].Ceilings[$k]) before it, with no reason written")
            }
        }
    }
    $latest = $ceilings[$ceilings.Count - 1]
    foreach ($f in $weighed) {
        if (-not $mem.ContainsKey($f)) { continue }
        $r = $mem[$f]
        foreach ($what in $ceilingFields.Keys) {
            $field = $ceilingFields[$what]
            if ($latest.Ceilings.ContainsKey("$f $what") -and [long]$r[$field[0]] -gt [long]$latest.Ceilings["$f $what"]) { $problems.Add("memory $($f): $($r[$field[0]]) $($field[1]), over the ceiling of $($latest.Ceilings["$f $what"])") }
        }
    }
    return $problems.ToArray()
}

# A history's entries as text, one a line, from a copy of this file: the
# floors', or by its name another's (KERNEL.49: the ceilings').
function Get-HistoryLines([string]$text, [string]$name = 'floorHistory') {
    $entries = New-Object System.Collections.Generic.List[string]
    $inside = $false
    $opening = '^\$' + $name + ' = @\(\s*$'
    foreach ($line in ($text -split "`r?`n")) {
        if (-not $inside) { if ($line -match $opening) { $inside = $true }; continue }
        if ($line -match '^\)\s*$') { break }
        $t = $line.Trim().TrimEnd(',').Trim()
        if ($t -ne '' -and -not $t.StartsWith('#')) { $entries.Add($t) }
    }
    return $entries.ToArray()
}

# Every committed entry still here, in its order, at the head.
function Test-Appended($committed, $current) {
    $problems = New-Object System.Collections.Generic.List[string]
    $committed = @($committed | Where-Object { $null -ne $_ }); $current = @($current | Where-Object { $null -ne $_ })
    for ($i = 0; $i -lt $committed.Count; $i++) {
        if ($i -ge $current.Count -or $current[$i] -ne $committed[$i]) {
            $problems.Add("the floor history's committed entry $($i + 1) is not where the last commit had it: an entry was pasted over or edited; put it back and add the new entry after it")
        }
    }
    return $problems.ToArray()
}

# A runner's check mode judged against the baseline's sums.
function Test-Runner([string[]]$lines, $rows) {
    $problems = New-Object System.Collections.Generic.List[string]
    $info = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($l in $lines) {
        if ($l -match '^\(recalc-differs ') { $problems.Add("the runner: $l") }
        if ($l -match '^\(recalc-flag-failed ') { $problems.Add("the runner: $l") }
        $m = [regex]::Match($l, '^\(recalc-check "([a-z]+)" \(sha256 "([0-9A-F]{64})"\) \(formulas (\d+)\) \(held (\d+)\) \(sum-built ([-0-9.e+]+)\) \(sum-edited ([-0-9.e+]+)\)\)$')
        if ($m.Success) { $seen[$m.Groups[1].Value] = @{ Sha256 = $m.Groups[2].Value; Held = [long]$m.Groups[4].Value; SumBuilt = [double]$m.Groups[5].Value; SumEdited = [double]$m.Groups[6].Value } }
    }
    if (-not ($lines -match ('^\(recalc-check-done ' + $fixtures.Count + '\)$'))) { $problems.Add("the runner did not finish its check mode over the $($fixtures.Count) fixtures") }
    foreach ($f in $fixtures) {
        if (-not $seen.ContainsKey($f)) { $problems.Add("the runner: no check row for $f"); continue }
        $s = $seen[$f]
        if ($s.Held -le 0) { $problems.Add("the runner: $f held no value") }
        $n = @($rows | Where-Object { $null -ne $_ -and $_.Half -eq 'native' -and $_.Fixture -eq $f })
        if ($n.Count -ne 1) { continue }
        if ([double]$n[0].SumBuilt -ne $s.SumBuilt -or [double]$n[0].SumEdited -ne $s.SumEdited) { $problems.Add("the runner: $f sums $($s.SumBuilt) and $($s.SumEdited) today where the baseline measured $($n[0].SumBuilt) and $($n[0].SumEdited)") }
        if ($n[0].Sha256 -ne $s.Sha256) { $info.Add("  $($f): today's file is $($s.Sha256.Substring(0, 12))..., the measured one $($n[0].Sha256.Substring(0, 12))...; its values are the measured ones (the stamp or the writer moved, not the computation)") }
    }
    return @{ Problems = $problems; Info = $info }
}

# The memory runner's check mode judged against the latest ceilings and the
# baseline's formulas (KERNEL.49): exact bytes, so a heavier build fails by
# name on every push, and a lighter one is printed for a new entry to lower
# its ceiling.
function Test-MemoryRunner([string[]]$lines, $rows, $ceilings) {
    $problems = New-Object System.Collections.Generic.List[string]
    $info = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($l in $lines) {
        if ($l -match '^\(memory-failed ') { $problems.Add("the memory runner: $l") }
        $m = [regex]::Match($l, '^\(memory-check "([a-z0-9-]+)" \(sha256 "([0-9A-F]{64})"\) \(formulas (\d+)\) \(load-held (\d+)\) \(held (\d+)\) \(peak (\d+)\) \(save (\d+)\) \(save-peak (\d+)\) \(saved-held (\d+)\) \(saved-peak (\d+)\)\)$')
        if ($m.Success) { $seen[$m.Groups[1].Value] = @{ Sha256 = $m.Groups[2].Value; Formulas = [long]$m.Groups[3].Value; LoadHeld = [long]$m.Groups[4].Value; Held = [long]$m.Groups[5].Value; Peak = [long]$m.Groups[6].Value; Save = [long]$m.Groups[7].Value; SavePeak = [long]$m.Groups[8].Value; SavedHeld = [long]$m.Groups[9].Value; SavedPeak = [long]$m.Groups[10].Value } }
    }
    if (-not ($lines -match ('^\(memory-check-done ' + $weighed.Count + '\)$'))) { $problems.Add("the memory runner did not finish its check mode over the $($weighed.Count) fixtures") }
    $ceilings = @($ceilings | Where-Object { $null -ne $_ })
    $latest = if ($ceilings.Count -gt 0) { $ceilings[$ceilings.Count - 1].Ceilings } else { @{} }
    foreach ($f in $weighed) {
        if (-not $seen.ContainsKey($f)) { $problems.Add("the memory runner: no check row for $f"); continue }
        $s = $seen[$f]
        $n = @($rows | Where-Object { $null -ne $_ -and $_.Half -eq 'memory' -and $_.Fixture -eq $f })
        if ($n.Count -eq 1 -and [long]$n[0].Formulas -ne $s.Formulas) { $problems.Add("the memory runner: $f loads $($s.Formulas) formulas today where the baseline weighed $($n[0].Formulas)") }
        if ($n.Count -eq 1 -and $n[0].Sha256 -ne $s.Sha256) { $problems.Add("the memory runner: $f weighs $($s.Sha256) today, not the measured $($n[0].Sha256)") }
        foreach ($what in $ceilingFields.Keys) {
            $k = "$f $what"
            if (-not $latest.ContainsKey($k)) { continue }
            $field = $ceilingFields[$what]
            $bytes = $s[$field[0]]
            $ceiling = [long]$latest[$k]
            if ($bytes -gt $ceiling) { $problems.Add("the memory runner: $f, $bytes $($field[1]) today, over its ceiling of $ceiling") }
            elseif ($bytes -lt $ceiling) { $info.Add(("  {0}: {1} {2} today, {3} under its ceiling of {4}: a new entry may lower it" -f $f, $bytes, $field[1], ($ceiling - $bytes), $ceiling)) }
        }
        foreach ($p in @(Get-SaveProblems 'the memory runner:' $f $s)) { $problems.Add($p) }
    }
    # A fixture that adds to a twin's, within the bound of it today (KERNEL.67).
    foreach ($f in $twins.Keys) {
        $t = $twins[$f]
        if (-not $seen.ContainsKey($f) -or -not $seen.ContainsKey($t)) { continue }
        foreach ($point in $twinPoints.Keys) {
            $more = $seen[$f][$point] - $seen[$t][$point]
            if ($more -gt $twinBound) { $problems.Add("the memory runner: $f holds $more bytes more than its twin $t $($twinPoints[$point]) today, over the bound of $twinBound") }
        }
    }
    return @{ Problems = $problems; Info = $info }
}

# A runner's check mode started, its output read as it comes and its stderr
# let go; a .ps1 is a fake runner of the control, run by the shell. Started
# first and waited for after, so that the two runners run side by side (AS.10).
function Start-Runner([string]$exe) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    if ($exe.EndsWith('.ps1')) {
        $psi.FileName = if (Get-Command powershell -ErrorAction SilentlyContinue) { 'powershell' } else { 'pwsh' }
        $psi.Arguments = ('-NoProfile -ExecutionPolicy Bypass -File "{0}" --check' -f $exe)
    } else {
        $psi.FileName = $exe
        $psi.Arguments = '--check'
    }
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $p = [System.Diagnostics.Process]::Start($psi)
    return @{ Process = $p; Out = $p.StandardOutput.ReadToEndAsync(); Err = $p.StandardError.ReadToEndAsync() }
}

function Wait-Runner($started) {
    $started.Process.WaitForExit()
    $text = $started.Out.Result
    $null = $started.Err.Result
    $started.Process.Dispose()
    $lines = @($text -split "`r?`n")
    if ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq '') { $lines = @($lines | Select-Object -First ($lines.Count - 1)) }
    return $lines
}

function Invoke-Runner([string]$exe) {
    return @(Wait-Runner (Start-Runner $exe))
}

# The newest built runner of a name under a root, the dev profile's or the
# release profile's, or '' (AS.10).
function Find-Runner([string]$root, [string]$name) {
    $suffix = if ($onWindows) { '.exe' } else { '' }
    $found = @(@('target/debug/examples', 'target/release/examples') | ForEach-Object { Join-Path $root "$_/$name$suffix" } | Where-Object { Test-Path -LiteralPath $_ } | ForEach-Object { Get-Item -LiteralPath $_ } | Sort-Object LastWriteTimeUtc -Descending)
    if ($found.Count -eq 0) { return '' }
    return $found[0].FullName
}

# The files a runner is built from, and the files it reads as it runs.
$runnerSources = @('vla-lang/src', 'vla-lang/data', 'vla-lang/Cargo.toml', 'core/src', 'core/data', 'core/examples', 'core/Cargo.toml', 'Cargo.toml', 'Cargo.lock')
$runnerInputs = @('scripts/recalc/bench', 'scripts/prelude.vla', 'scripts/polyglotta/english.vla', 'scripts/build/fixture.txt')

# The newest of a runner's sources when it is newer than the runner, its
# path from the root; else ''.
function Get-StaleSource([string]$exe, [string]$root, [string[]]$sources) {
    $built = (Get-Item -LiteralPath $exe).LastWriteTimeUtc
    $full = (Resolve-Path -LiteralPath $root).Path
    $newest = $null
    foreach ($s in $sources) {
        $p = Join-Path $full $s
        if (-not (Test-Path -LiteralPath $p)) { continue }
        foreach ($f in @(Get-ChildItem -LiteralPath $p -Recurse -File -ErrorAction SilentlyContinue)) {
            if ($null -eq $newest -or $f.LastWriteTimeUtc -gt $newest.LastWriteTimeUtc) { $newest = $f }
        }
    }
    if ($null -ne $newest -and $newest.LastWriteTimeUtc -gt $built) { return $newest.FullName.Substring($full.Length).TrimStart('\', '/').Replace('\', '/') }
    return ''
}

# One SHA-256 over every file a runner reads, each by its path and its own.
function Get-InputsDigest([string]$root, [string[]]$inputs) {
    $full = (Resolve-Path -LiteralPath $root).Path
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($s in $inputs) {
        $p = Join-Path $full $s
        if (-not (Test-Path -LiteralPath $p)) { $lines.Add("$s absent"); continue }
        foreach ($f in @(Get-ChildItem -LiteralPath $p -Recurse -File | Sort-Object FullName)) {
            $lines.Add(('{0} {1}' -f $f.FullName.Substring($full.Length).Replace('\', '/'), (Get-FileHash -Algorithm SHA256 -LiteralPath $f.FullName).Hash))
        }
    }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return ([System.BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes(($lines -join "`n")))) -replace '-', '') } finally { $sha.Dispose() }
}

# A runner's output kept (AS.10): its stamp, the runner's SHA-256 and its
# inputs', and the lines when the cache holds them under that stamp.
function Get-RunnerKept([string]$exe, [string]$kind, [string]$cache, [string]$inputs) {
    $stamp = "runner $((Get-FileHash -Algorithm SHA256 -LiteralPath $exe).Hash) inputs $inputs"
    $stampPath = Join-Path $cache "$kind.stamp"
    $outPath = Join-Path $cache "$kind.check"
    if ((Test-Path -LiteralPath $stampPath) -and (Test-Path -LiteralPath $outPath) -and ([System.IO.File]::ReadAllText($stampPath) -eq $stamp)) {
        return @{ Stamp = $stamp; Lines = @([System.IO.File]::ReadAllLines($outPath)) }
    }
    return @{ Stamp = $stamp; Lines = $null }
}

# Output that passed, kept: the lines first and the stamp after, each
# written beside its place and moved in, so that a stamp is never read over
# lines it did not come with.
function Save-RunnerKept([string]$kind, [string]$cache, [string]$stamp, [string[]]$lines) {
    if (-not (Test-Path -LiteralPath $cache)) { New-Item -ItemType Directory -Path $cache -Force | Out-Null }
    $outPath = Join-Path $cache "$kind.check"
    $stampPath = Join-Path $cache "$kind.stamp"
    [System.IO.File]::WriteAllLines("$outPath.$PID.tmp", $lines)
    Move-Item -LiteralPath "$outPath.$PID.tmp" -Destination $outPath -Force
    [System.IO.File]::WriteAllText("$stampPath.$PID.tmp", $stamp)
    Move-Item -LiteralPath "$stampPath.$PID.tmp" -Destination $stampPath -Force
}

# ------------------------------------------------------------------ -Control
if ($Control) {
    $failures = New-Object System.Collections.Generic.List[string]
    $okEnvNative = 'cpu | os | rustc | frazaro-core 0.8.1 | power throttling opted out: the runner yes | high priority: the runner yes | on the performance cores, 8 of 20 logical processors (efficiency class 1): the runner yes'
    $okEnvExcel = 'cpu | os | Excel 16.0 build 19127 | all threads 20; one thread on the performance cores, mask 0xC03C3 | power throttling opted out: yes'
    function New-Clean {
        $rows = New-Object System.Collections.Generic.List[object]
        $sums = @{ 'column' = @(1100080000, 1425100000); 'chain' = @(1200039974, 1200059975); 'sums' = @(14800794, 14900794); 'life' = @(23683, 9603); 'lookups' = @(51995036, 52005012); 'criteria' = @(12010000, 12030000); 'netlist' = @(8192, 12288) }
        $sha = @{ 'column' = ('A' * 64); 'chain' = ('B' * 64); 'sums' = ('C' * 64); 'life' = ('D' * 64); 'lookups' = ('F' * 64); 'criteria' = ('1' * 64); 'netlist' = ('2' * 64) }
        foreach ($f in $fixtures) {
            $native = @{ Half = 'native'; Fixture = $f; Formulas = 60000; Sha256 = $sha[$f]; Rounds = 3; Load = 200.0; First = 15.0; Whole = 3.0; Edit = 5.0; Calc = 200.0; WholePerSecond = 20000000; EditsPerSecond = 200.0; SumBuilt = $sums[$f][0]; SumEdited = $sums[$f][1]; ClockMs = 0.182; ClockSpread = 1.0; ClockDrift = 0.5 }
            if ($typed -contains $f) { $native.Type = 20.0; $native.TypesPerSecond = 50.0 }
            if ($rewritten -contains $f) { $native.Rewrite = 0.008; $native.RewritesPerSecond = 125000.0; $native.RewriteWrite = 40.0; $native.RewriteWritesPerSecond = 25.0 }
            $rows.Add($native)
            foreach ($mode in $modes) {
                $excelRow = @{ Half = 'excel'; Fixture = $f; Threads = $mode; ThreadCount = $(if ($mode -eq 'one') { 1 } else { 20 }); Sha256 = $sha[$f]; Open = 900.0; Whole = 4.0; Rebuild = 9.0; Edit = 3.0; Copy = 0.0; ComFloor = 0.05; SumBuilt = $sums[$f][0]; SumEdited = $sums[$f][1] }
                if ($typed -contains $f) { $excelRow.Type = 0.3 }
                if ($rewritten -contains $f) { $excelRow.Rewrite = 0.4; $excelRow.RewriteWrite = 0.5 }
                $rows.Add($excelRow)
            }
        }
        foreach ($file in $flagFiles) { $rows.Add(@{ Half = 'excel-flag'; File = $file; C2 = 15; Saved = 'False'; Verdict = 'computed on opening' }) }
        $floors = @{}
        foreach ($f in $fixtures) { $floors["$f whole"] = 18000000; $floors["$f edit"] = 180 }
        foreach ($f in $typed) { $floors["$f type"] = 45 }
        foreach ($f in $rewritten) { $floors["$f rewrite"] = 100000; $floors["$f rewrite write"] = 20 }
        foreach ($f in $loadFloored) { $floors["$f load"] = 250000 }
        $measured = [ordered]@{}; $now = [ordered]@{}
        foreach ($k in $programs.Keys) { $measured[$k] = $programs[$k]; $now[$k] = $programs[$k] }
        $rowsMeasured = [ordered]@{}; $rowsNow = [ordered]@{}
        foreach ($k in $rowsFixtures.Keys) { $rowsMeasured[$k] = $rowsFixtures[$k]; $rowsNow[$k] = $rowsFixtures[$k] }
        $ceilings = @{}
        foreach ($f in $weighed) {
            $source = if ($rowsFixtures.Contains($f)) { $rowsFixtures[$f] } else { $programs[$f] }
            $savedPeak = if ($rowsFixtures.Contains($f)) { 20000064 } else { 30000000 }
            $rows.Add(@{ Half = 'memory'; Fixture = $f; Formulas = 60000; Sha256 = $source; Rounds = 3; Load = 200.0; LoadPerSecond = 300000; LoadHeld = 20000000; Held = 24000000; Peak = 30000000; Save = 200000; SavePeak = 400000; SavedHeld = 20000000; SavedPeak = $savedPeak; ClockMs = 0.182; ClockSpread = 1.0; ClockDrift = 0.5 })
            $ceilings["$f held"] = 24000000; $ceilings["$f peak"] = 30000000; $ceilings["$f save"] = 200000; $ceilings["$f save peak"] = 400000; $ceilings["$f saved held"] = 20000000; $ceilings["$f saved peak"] = $savedPeak
        }
        return @{ Rows = $rows; History = @(@{ Date = '2026-10-10'; Reason = ''; Floors = $floors }); Ceilings = @(@{ Date = '2026-10-10'; Reason = ''; Ceilings = $ceilings }); Native = $okEnvNative; Excel = $okEnvExcel; Programs = $measured; Now = $now; RowsMeasured = $rowsMeasured; RowsNow = $rowsNow }
    }
    function Get-Problems($c) { return @(Test-Floors '2026-10-10' $c.Native $c.Excel $c.Rows $c.History $c.Programs $c.Now) + @(Test-Memory $c.Rows $c.Ceilings $c.Programs $c.RowsMeasured $c.RowsNow) }
    # A case that fails by the rule meant, by a word of its line (KERNEL.72).
    function Expect-Says([string]$what, $c, [string]$word) {
        $p = @(Get-Problems $c)
        if (-not @($p | Where-Object { $_ -like "*$word*" })) { $failures.Add("control: $what should say '$word' and says: $($p -join ' / ')") }
        else { Write-Output ("  control: {0}: says so" -f $what) }
    }
    function Expect([string]$what, $c, [bool]$clean) {
        $p = @(Get-Problems $c)
        if ($clean -and $p.Count -ne 0) { $failures.Add("control: $what should pass and fails: $($p -join ' / ')") }
        elseif (-not $clean -and $p.Count -eq 0) { $failures.Add("control: $what should fail and passes") }
        else { Write-Output ("  control: {0}: {1}" -f $what, $(if ($clean) { 'passes' } else { "fails ($($p.Count))" })) }
    }
    Expect 'a clean baseline' (New-Clean) $true
    $c = New-Clean; $prev = @{}; foreach ($k in $c.History[0].Floors.Keys) { $prev[$k] = $c.History[0].Floors[$k] + 1000 }; $c.History = @(@{ Date = '2026-10-09'; Reason = ''; Floors = $prev }, @{ Date = '2026-10-10'; Reason = 'a correctness fix that costs speed'; Floors = $c.History[0].Floors }); Expect 'a floor lowered for a written reason' $c $true
    $c = New-Clean; $c.Rows = @(); Expect 'an empty baseline' $c $false
    $c = New-Clean; $c.Rows = @($c.Rows | Where-Object { -not ($_.Half -eq 'native' -and $_.Fixture -eq 'life') }); Expect 'a native row missing' $c $false
    $c = New-Clean; $c.Rows = @($c.Rows | Where-Object { -not ($_.Half -eq 'excel' -and $_.Fixture -eq 'chain' -and $_.Threads -eq 'one') }); Expect 'an Excel row missing' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'excel' -and $_.Fixture -eq 'sums' -and $_.Threads -eq 'all' })[0].Sha256 = ('E' * 64); Expect 'a file not the same bytes in both halves' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'excel' -and $_.Fixture -eq 'life' -and $_.Threads -eq 'one' })[0].SumEdited = 9604; Expect 'sums that differ' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq 'column' })[0].ClockSpread = 12.0; Expect 'a clock that moved' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq 'chain' })[0].Rounds = 1; Expect 'too few rounds' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq 'sums' })[0].WholePerSecond = 21000000; Expect 'formulas a second that disagree with the time' $c $false
    $c = New-Clean; $c.Native = $c.Native -replace 'the runner yes \| high', 'the runner no | high'; Expect 'a native half not opted out' $c $false
    $c = New-Clean; $c.Excel = $c.Excel + ' | THE STAND-IN, NOT EXCEL'; Expect 'Excel''s half from the stand-in' $c $false
    $c = New-Clean; $prev = @{}; foreach ($k in $c.History[0].Floors.Keys) { $prev[$k] = $c.History[0].Floors[$k] + 1000 }; $c.History = @(@{ Date = '2026-10-09'; Reason = ''; Floors = $prev }, $c.History[0]); Expect 'a floor lowered with no reason' $c $false
    $c = New-Clean; $prev = @{}; foreach ($k in $c.History[0].Floors.Keys) { if ($k -notlike 'lookups *') { $prev[$k] = $c.History[0].Floors[$k] } }; $c.History = @(@{ Date = '2026-10-09'; Reason = ''; Floors = $prev }, $c.History[0]); Expect 'a fixture joining the history at a later entry' $c $true
    $c = New-Clean; $later = @{}; foreach ($k in $c.History[0].Floors.Keys) { if ($k -notlike 'lookups *') { $later[$k] = $c.History[0].Floors[$k] } }; $c.History = @($c.History[0], @{ Date = '2026-10-11'; Reason = ''; Floors = $later }); Expect 'a fixture''s floors left out after it joined' $c $false
    $c = New-Clean; $c.History[0].Floors['life edit'] = 250; Expect 'a row under its floor' $c $false
    # The type (KERNEL.56).
    $c = New-Clean; $n = @($c.Rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq 'netlist' })[0]; $n.Remove('Type'); $n.Remove('TypesPerSecond'); Expect 'a typed fixture''s row with no type' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq 'netlist' })[0].TypesPerSecond = 60.0; Expect 'types a second that disagree with the time' $c $false
    $c = New-Clean; $n = @($c.Rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq 'netlist' })[0]; $n.Type = 0.004; $n.TypesPerSecond = 285714.3; Expect 'a type of microseconds at three places' $c $true
    $c = New-Clean; $c.History[0].Floors['netlist type'] = 60; Expect 'a type under its floor' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'excel' -and $_.Fixture -eq 'netlist' -and $_.Threads -eq 'one' })[0].Remove('Type'); Expect 'Excel''s row of a typed fixture with no type' $c $false
    $c = New-Clean; $prev = @{}; foreach ($k in $c.History[0].Floors.Keys) { if ($k -ne 'netlist type') { $prev[$k] = $c.History[0].Floors[$k] } }; $c.History = @(@{ Date = '2026-10-09'; Reason = ''; Floors = $prev }, $c.History[0]); Expect 'a type joining the history at a later entry' $c $true
    # The rewrite (KERNEL.69).
    $c = New-Clean; $n = @($c.Rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq 'netlist' })[0]; $n.Remove('Rewrite'); Expect 'a rewritten fixture''s row with no rewrite' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq 'netlist' })[0].RewritesPerSecond = 50000.0; Expect 'rewrites a second that disagree with the time' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq 'netlist' })[0].RewriteWritesPerSecond = 60.0; Expect 'rewrite writes a second that disagree with the time' $c $false
    $c = New-Clean; $c.History[0].Floors['netlist rewrite'] = 200000; Expect 'a rewrite under its floor' $c $false
    $c = New-Clean; $c.History[0].Floors['netlist rewrite write'] = 30; Expect 'a rewrite''s write under its floor' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'excel' -and $_.Fixture -eq 'netlist' -and $_.Threads -eq 'all' })[0].Remove('RewriteWrite'); Expect 'Excel''s row of a rewritten fixture with no rewrite' $c $false
    $c = New-Clean; $prev = @{}; foreach ($k in $c.History[0].Floors.Keys) { if ($k -notlike 'netlist rewrite*') { $prev[$k] = $c.History[0].Floors[$k] } }; $c.History = @(@{ Date = '2026-10-09'; Reason = ''; Floors = $prev }, $c.History[0]); Expect 'the rewrite joining the history at its entry' $c $true
    $c = New-Clean; $c.Now['chain'] = ('0' * 64); Expect 'a program edited' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'excel-flag' })[1].Verdict = 'maybe'; Expect 'a flag row with no verdict' $c $false
    $c = New-Clean; $c.Rows = @($c.Rows | Where-Object { -not ($_.Half -eq 'excel' -and $_.Fixture -eq 'life') }) + @(@{ Half = 'excel'; Fixture = 'life'; Threads = 'all'; Sha256 = ('D' * 64); Failed = 'Unable to get the Open property of the Workbooks class' }, @{ Half = 'excel'; Fixture = 'life'; Threads = 'one'; Sha256 = ('D' * 64); Failed = 'Unable to get the Open property of the Workbooks class' }); Expect 'a fixture Excel could not open' $c $false
    # The memory half (KERNEL.49).
    $c = New-Clean; $prev = @{}; foreach ($k in $c.Ceilings[0].Ceilings.Keys) { $prev[$k] = $c.Ceilings[0].Ceilings[$k] - 1000 }; $c.Ceilings = @(@{ Date = '2026-10-09'; Reason = ''; Ceilings = $prev }, @{ Date = '2026-10-10'; Reason = 'a frame that keeps a store a twin needs'; Ceilings = $c.Ceilings[0].Ceilings }); Expect 'a ceiling raised for a written reason' $c $true
    $c = New-Clean; $prev = @{}; foreach ($k in $c.Ceilings[0].Ceilings.Keys) { if ($k -notlike 'life-x16 *') { $prev[$k] = $c.Ceilings[0].Ceilings[$k] } }; $c.Ceilings = @(@{ Date = '2026-10-09'; Reason = ''; Ceilings = $prev }, $c.Ceilings[0]); Expect 'a fixture joining the ceilings at a later entry' $c $true
    $c = New-Clean; $c.Rows = @($c.Rows | Where-Object { -not ($_.Half -eq 'memory' -and $_.Fixture -eq 'superposition') }); Expect 'a memory row missing' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'memory' -and $_.Fixture -eq 'superposition' })[0].Held = 24000001; Expect 'a memory row over its ceiling' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'memory' -and $_.Fixture -eq 'chain' })[0].Peak = 23999999; Expect 'a peak under the bytes held' $c $false
    $c = New-Clean; $prev = @{}; foreach ($k in $c.Ceilings[0].Ceilings.Keys) { $prev[$k] = $c.Ceilings[0].Ceilings[$k] - 1000 }; $c.Ceilings = @(@{ Date = '2026-10-09'; Reason = ''; Ceilings = $prev }, $c.Ceilings[0]); Expect 'a ceiling raised with no reason' $c $false
    $c = New-Clean; $later = @{}; foreach ($k in $c.Ceilings[0].Ceilings.Keys) { if ($k -notlike 'superposition *') { $later[$k] = $c.Ceilings[0].Ceilings[$k] } }; $c.Ceilings = @($c.Ceilings[0], @{ Date = '2026-10-11'; Reason = ''; Ceilings = $later }); Expect 'a fixture''s ceilings left out after it joined' $c $false
    $c = New-Clean; $c.History[0].Floors['life-x16 load'] = 400000; Expect 'a load under its floor' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'memory' -and $_.Fixture -eq 'superposition' })[0].LoadPerSecond = 310000; Expect 'a load a second that disagrees with its time' $c $false
    $c = New-Clean; $c.RowsNow['superposition'] = ('0' * 64); Expect 'a fixture of rows edited' $c $false
    $c = New-Clean; @($c.Rows | Where-Object { $_.Half -eq 'memory' -and $_.Fixture -eq 'sums' })[0].Sha256 = ('E' * 64); Expect 'a memory row weighing another source' $c $false
    # A fixture against its twin (KERNEL.67), through the round trip (KERNEL.72):
    # over the bound at each point, by the rule itself, and at it.
    foreach ($point in $twinPoints.Keys) {
        $c = New-Clean; $t = @($c.Rows | Where-Object { $_.Half -eq 'memory' -and $_.Fixture -eq 'superposition' })[0]; $t[$point] = [long]$t[$point] - $twinBound - 1; Expect-Says "a fixture over its twin's bytes $($twinPoints[$point]) by more than the bound" $c "more than its twin superposition $($twinPoints[$point])"
    }
    $c = New-Clean; $t = @($c.Rows | Where-Object { $_.Half -eq 'memory' -and $_.Fixture -eq 'superposition' })[0]; $t.Peak = [long]$t.Peak - $twinBound; Expect 'a fixture at its twin''s bound' $c $true
    $c = New-Clean; $t = @($c.Rows | Where-Object { $_.Half -eq 'memory' -and $_.Fixture -eq 'superposition' })[0]; $t.Save = [long]$t.Save - $twinBound; Expect 'a fixture''s save at its twin''s bound' $c $true
    # What a save costs (KERNEL.72): each rule, by its own words, and at its bound.
    function Get-MemoryRow($c, [string]$f) { return @($c.Rows | Where-Object { $_.Half -eq 'memory' -and $_.Fixture -eq $f })[0] }
    $c = New-Clean; (Get-MemoryRow $c 'sums').SavedHeld = 24000001; Expect-Says 'a load from a save holding more than the machine that printed it' $c 'more than the 24000000 the machine'
    $c = New-Clean; (Get-MemoryRow $c 'sums').SavedHeld = 24000000; $c.Ceilings[0].Ceilings['sums saved held'] = 24000000; Expect 'a load from a save holding what the machine held' $c $true
    $c = New-Clean; (Get-MemoryRow $c 'chain').SavePeak = 800001; Expect-Says 'a save''s printing over its bound' $c 'more than 4 times its text'
    $c = New-Clean; (Get-MemoryRow $c 'chain').SavePeak = 800000; $c.Ceilings[0].Ceilings['chain save peak'] = 800000; Expect 'a save''s printing at its bound' $c $true
    $c = New-Clean; (Get-MemoryRow $c 'life-x16').Save = 240001; Expect-Says 'a fixture of rows'' save over its bytes a formula cell' $c 'more than 4 bytes a formula cell'
    $c = New-Clean; (Get-MemoryRow $c 'life-x16').Save = 240000; $c.Ceilings[0].Ceilings['life-x16 save'] = 240000; Expect 'a fixture of rows'' save at its bytes a formula cell' $c $true
    $c = New-Clean; (Get-MemoryRow $c 'life-x16').SavedPeak = 20000000 + $twinBound + 1; Expect-Says 'a fixture of rows'' load from its save over what it holds' $c "more than $twinBound over the 20000000 it holds"
    $c = New-Clean; (Get-MemoryRow $c 'life-x16').SavedPeak = 20000000 + $twinBound; $c.Ceilings[0].Ceilings['life-x16 saved peak'] = 20000000 + $twinBound; Expect 'a fixture of rows'' load from its save at its bound' $c $true
    $c = New-Clean; (Get-MemoryRow $c 'lookups').SavedPeak = 80000001; Expect-Says 'a program''s load from its save over its bound' $c 'more than 4 times the 20000000 it holds'
    $c = New-Clean; (Get-MemoryRow $c 'lookups').SavedPeak = 80000000; $c.Ceilings[0].Ceilings['lookups saved peak'] = 80000000; Expect 'a program''s load from its save at its bound' $c $true
    $c = New-Clean; (Get-MemoryRow $c 'netlist').SavedPeak = 19999999; Expect-Says 'a load from a save peaking under what it holds' $c 'under the 20000000 it holds'
    $c = New-Clean; (Get-MemoryRow $c 'netlist').Save = 0; Expect-Says 'a fixture with no save weighed' $c 'no save weighed'
    foreach ($what in @('save', 'save peak', 'saved held', 'saved peak')) {
        $c = New-Clean; $field = $ceilingFields[$what]; $r = Get-MemoryRow $c 'criteria'; $r[$field[0]] = [long]$r[$field[0]] + 1; Expect-Says "a row over its ceiling, $what" $c "$($field[1]), over the ceiling"
        $c = New-Clean; $c.Ceilings[0].Ceilings.Remove("criteria $what"); Expect-Says "the history with no ceiling '$what'" $c "no ceiling 'criteria $what'"
    }
    # The history, from text.
    $committed = Get-HistoryLines "`$floorHistory = @(`r`n    @{ Date = '2026-10-10'; Reason = ''; Floors = @{ 'column whole' = 1 } },`r`n    @{ Date = '2026-10-11'; Reason = ''; Floors = @{ 'column whole' = 2 } }`r`n)"
    $cases = @(
        @{ What = 'an entry added after the committed ones'; Clean = $true; Text = "`$floorHistory = @(`n    @{ Date = '2026-10-10'; Reason = ''; Floors = @{ 'column whole' = 1 } },`n    @{ Date = '2026-10-11'; Reason = ''; Floors = @{ 'column whole' = 2 } },`n    @{ Date = '2026-10-12'; Reason = ''; Floors = @{ 'column whole' = 3 } }`n)" },
        @{ What = 'a history unchanged'; Clean = $true; Text = "`$floorHistory = @(`n    @{ Date = '2026-10-10'; Reason = ''; Floors = @{ 'column whole' = 1 } },`n    @{ Date = '2026-10-11'; Reason = ''; Floors = @{ 'column whole' = 2 } }`n)" },
        @{ What = 'an entry pasted over a committed one'; Clean = $false; Text = "`$floorHistory = @(`n    @{ Date = '2026-10-11'; Reason = ''; Floors = @{ 'column whole' = 2 } },`n    @{ Date = '2026-10-12'; Reason = ''; Floors = @{ 'column whole' = 3 } }`n)" },
        @{ What = 'a committed entry edited in place'; Clean = $false; Text = "`$floorHistory = @(`n    @{ Date = '2026-10-10'; Reason = ''; Floors = @{ 'column whole' = 1 } },`n    @{ Date = '2026-10-11'; Reason = ''; Floors = @{ 'column whole' = 1 } }`n)" }
    )
    foreach ($case in $cases) {
        $p = @(Test-Appended $committed (Get-HistoryLines $case.Text))
        if ($case.Clean -and $p.Count -ne 0) { $failures.Add("control: the history, $($case.What), should pass and fails") }
        elseif (-not $case.Clean -and $p.Count -eq 0) { $failures.Add("control: the history, $($case.What), should fail and passes") }
        else { Write-Output ("  control: the history, {0}: {1}" -f $case.What, $(if ($case.Clean) { 'passes' } else { 'fails' })) }
    }
    # The ceilings' history, from text, by its own name (KERNEL.49).
    $committed = Get-HistoryLines "`$ceilingHistory = @(`r`n    @{ Date = '2026-10-21'; Reason = ''; Ceilings = @{ 'superposition held' = 2 } }`r`n)" 'ceilingHistory'
    $cases = @(
        @{ What = 'an entry added after the committed one'; Clean = $true; Text = "`$ceilingHistory = @(`n    @{ Date = '2026-10-21'; Reason = ''; Ceilings = @{ 'superposition held' = 2 } },`n    @{ Date = '2026-10-22'; Reason = ''; Ceilings = @{ 'superposition held' = 1 } }`n)" },
        @{ What = 'an entry pasted over the committed one'; Clean = $false; Text = "`$ceilingHistory = @(`n    @{ Date = '2026-10-22'; Reason = ''; Ceilings = @{ 'superposition held' = 1 } }`n)" }
    )
    foreach ($case in $cases) {
        $p = @(Test-Appended $committed (Get-HistoryLines $case.Text 'ceilingHistory'))
        if ($case.Clean -and $p.Count -ne 0) { $failures.Add("control: the ceilings' history, $($case.What), should pass and fails") }
        elseif (-not $case.Clean -and $p.Count -eq 0) { $failures.Add("control: the ceilings' history, $($case.What), should fail and passes") }
        else { Write-Output ("  control: the ceilings' history, {0}: {1}" -f $case.What, $(if ($case.Clean) { 'passes' } else { 'fails' })) }
    }
    # Fake runners: scripts printing a check mode.
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('recalc_floors_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $clean = New-Clean
        $good = New-Object System.Collections.Generic.List[string]
        foreach ($f in $fixtures) {
            $n = @($clean.Rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq $f })[0]
            $good.Add(('(recalc-check "{0}" (sha256 "{1}") (formulas 60000) (held 120000) (sum-built {2}) (sum-edited {3}))' -f $f, $n.Sha256, $n.SumBuilt, $n.SumEdited))
        }
        $good.Add('(recalc-check-flags 3)'); $good.Add('(recalc-check-done ' + $fixtures.Count + ')')
        $fakes = @(
            @{ What = 'a clean runner'; Clean = $true; Lines = $good.ToArray() },
            @{ What = 'a runner whose values differ'; Clean = $false; Lines = @('(recalc-differs "chain: Output!R5C2: the machine holds 1, Calc 2")') },
            @{ What = 'a runner whose sums differ'; Clean = $false; Lines = @($good.ToArray() | ForEach-Object { $_ -replace '\(sum-edited 9603\)', '(sum-edited 9600)' }) },
            @{ What = 'a runner that leaves out a fixture'; Clean = $false; Lines = @($good.ToArray() | Where-Object { $_ -notmatch '"sums"' }) }
        )
        $k = 0
        foreach ($fake in $fakes) {
            $k++
            $script = Join-Path $tmp "fake$k.ps1"
            $body = ($fake.Lines | ForEach-Object { "Write-Output '" + ($_ -replace "'", "''") + "'" }) -join "`r`n"
            [System.IO.File]::WriteAllText($script, $body + "`r`n", (New-Object System.Text.UTF8Encoding($false)))
            $judged = Test-Runner (Invoke-Runner $script) $clean.Rows
            $n = $judged.Problems.Count
            if ($fake.Clean -and $n -ne 0) { $failures.Add("control: $($fake.What) should pass and fails: $($judged.Problems -join ' / ')") }
            elseif (-not $fake.Clean -and $n -eq 0) { $failures.Add("control: $($fake.What) should fail and passes") }
            else { Write-Output ("  control: {0}: {1}" -f $fake.What, $(if ($fake.Clean) { 'passes' } else { "fails ($n)" })) }
        }
        # Fake memory runners (KERNEL.49).
        $weighs = New-Object System.Collections.Generic.List[string]
        foreach ($f in $weighed) {
            $n = @($clean.Rows | Where-Object { $_.Half -eq 'memory' -and $_.Fixture -eq $f })[0]
            $weighs.Add(('(memory-check "{0}" (sha256 "{1}") (formulas 60000) (load-held 20000000) (held 24000000) (peak 30000000) (save 200000) (save-peak 400000) (saved-held 20000000) (saved-peak {2}))' -f $f, $n.Sha256, $n.SavedPeak))
        }
        $weighs.Add('(memory-check-done ' + $weighed.Count + ')')
        $memoryFakes = @(
            @{ What = 'a clean memory runner'; Clean = $true; Lines = $weighs.ToArray() },
            @{ What = 'a memory runner over a ceiling'; Clean = $false; Lines = @($weighs.ToArray() | ForEach-Object { if ($_ -match '"life-x16"') { $_ -replace '\(peak 30000000\)', '(peak 30000001)' } else { $_ } }) },
            @{ What = 'a memory runner whose formulas differ'; Clean = $false; Lines = @($weighs.ToArray() | ForEach-Object { if ($_ -match '"column"') { $_ -replace '\(formulas 60000\)', '(formulas 59999)' } else { $_ } }) },
            @{ What = 'a memory runner that leaves out a fixture'; Clean = $false; Lines = @($weighs.ToArray() | Where-Object { $_ -notmatch '"superposition"' }) },
            @{ What = 'a memory runner that failed'; Clean = $false; Lines = @('(memory-failed "superposition: the machine refused line 3")') },
            @{ What = 'a memory runner whose fixture holds more than its twin by more than the bound'; Clean = $false; Lines = @($weighs.ToArray() | ForEach-Object { if ($_ -match '"superposition"') { $_ -replace '\(peak 30000000\)', ('(peak ' + (30000000 - $twinBound - 1) + ')') } else { $_ } }) },
            @{ What = 'a memory runner whose fixture''s save is more than its twin''s by more than the bound'; Clean = $false; Lines = @($weighs.ToArray() | ForEach-Object { if ($_ -match '"superposition"') { $_ -replace '\(save 200000\)', ('(save ' + (200000 - $twinBound - 1) + ')') } else { $_ } }) },
            @{ What = 'a memory runner over a save''s ceiling'; Clean = $false; Lines = @($weighs.ToArray() | ForEach-Object { if ($_ -match '"column"') { $_ -replace '\(save 200000\)', '(save 200001)' } else { $_ } }) },
            @{ What = 'a memory runner whose load from a save holds more than the machine that printed it'; Clean = $false; Lines = @($weighs.ToArray() | ForEach-Object { if ($_ -match '"chain"') { $_ -replace '\(saved-held 20000000\)', '(saved-held 24000001)' } else { $_ } }) },
            @{ What = 'a memory runner whose fixture of rows reads back over what it holds'; Clean = $false; Lines = @($weighs.ToArray() | ForEach-Object { if ($_ -match '"life-x16"') { $_ -replace '\(saved-peak 20000064\)', ('(saved-peak ' + (20000000 + $twinBound + 1) + ')') } else { $_ } }) },
            @{ What = 'a memory runner that weighs no save'; Clean = $false; Lines = @($weighs.ToArray() | ForEach-Object { $_ -replace ' \(save 200000\).*\)$', ')' }) }
        )
        foreach ($fake in $memoryFakes) {
            $k++
            $script = Join-Path $tmp "fake$k.ps1"
            $body = ($fake.Lines | ForEach-Object { "Write-Output '" + ($_ -replace "'", "''") + "'" }) -join "`r`n"
            [System.IO.File]::WriteAllText($script, $body + "`r`n", (New-Object System.Text.UTF8Encoding($false)))
            $judged = Test-MemoryRunner (Invoke-Runner $script) $clean.Rows $clean.Ceilings
            $n = $judged.Problems.Count
            if ($fake.Clean -and $n -ne 0) { $failures.Add("control: $($fake.What) should pass and fails: $($judged.Problems -join ' / ')") }
            elseif (-not $fake.Clean -and $n -eq 0) { $failures.Add("control: $($fake.What) should fail and passes") }
            else { Write-Output ("  control: {0}: {1}" -f $fake.What, $(if ($fake.Clean) { 'passes' } else { "fails ($n)" })) }
        }
        # The runners found, run side by side and kept (AS.10).
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $place = Join-Path $tmp 'root'
        $suffix = if ($onWindows) { '.exe' } else { '' }
        foreach ($d in @('target/debug/examples', 'target/release/examples', 'src')) { New-Item -ItemType Directory -Path (Join-Path $place $d) -Force | Out-Null }
        $devExe = Join-Path $place "target/debug/examples/recalc$suffix"
        $releaseExe = Join-Path $place "target/release/examples/recalc$suffix"
        $source = Join-Path $place 'src/lib.rs'
        foreach ($f in @($devExe, $releaseExe, $source)) { [System.IO.File]::WriteAllText($f, 'x') }
        $then = [DateTime]::UtcNow.AddHours(-3)
        (Get-Item -LiteralPath $releaseExe).LastWriteTimeUtc = $then.AddHours(1)
        (Get-Item -LiteralPath $devExe).LastWriteTimeUtc = $then.AddHours(2)
        (Get-Item -LiteralPath $source).LastWriteTimeUtc = $then
        $found = Find-Runner $place 'recalc'
        if ($found -ne (Get-Item -LiteralPath $devExe).FullName) { $failures.Add("control: of two runners built, the newer was not found: $found") }
        else { Write-Output '  control: of two runners built, the newer is found' }
        (Get-Item -LiteralPath $releaseExe).LastWriteTimeUtc = $then.AddHours(2.5)
        $found = Find-Runner $place 'recalc'
        if ($found -ne (Get-Item -LiteralPath $releaseExe).FullName) { $failures.Add("control: of two runners built, the newer, the release one, was not found: $found") }
        $fresh = Get-StaleSource $devExe $place @('src')
        (Get-Item -LiteralPath $source).LastWriteTimeUtc = $then.AddHours(2.25)
        $stale = Get-StaleSource $devExe $place @('src')
        if ($fresh -ne '') { $failures.Add("control: a runner newer than its source was named stale: $fresh") }
        elseif ($stale -ne 'src/lib.rs') { $failures.Add("control: a runner older than its source was not named stale: '$stale'") }
        else { Write-Output '  control: a runner older than a source is named stale, and one newer is not' }
        foreach ($who in @('one', 'two')) {
            $body = "Add-Content -LiteralPath (Join-Path `$PSScriptRoot '$who.time') ('start ' + [DateTime]::UtcNow.Ticks)`r`nStart-Sleep -Milliseconds 1500`r`nAdd-Content -LiteralPath (Join-Path `$PSScriptRoot '$who.time') ('end ' + [DateTime]::UtcNow.Ticks)`r`nWrite-Output '($who done)'`r`n"
            [System.IO.File]::WriteAllText((Join-Path $tmp "side_$who.ps1"), $body, $utf8)
        }
        $first = Start-Runner (Join-Path $tmp 'side_one.ps1')
        $second = Start-Runner (Join-Path $tmp 'side_two.ps1')
        $linesOne = @(Wait-Runner $first)
        $linesTwo = @(Wait-Runner $second)
        $timesOne = @(Get-Content -LiteralPath (Join-Path $tmp 'one.time'))
        $timesTwo = @(Get-Content -LiteralPath (Join-Path $tmp 'two.time'))
        $startOne = [long]($timesOne[0] -replace '^start ', ''); $endOne = [long]($timesOne[1] -replace '^end ', '')
        $startTwo = [long]($timesTwo[0] -replace '^start ', ''); $endTwo = [long]($timesTwo[1] -replace '^end ', '')
        if (-not ($startOne -lt $endTwo -and $startTwo -lt $endOne)) { $failures.Add('control: two runners started together did not run at once') }
        elseif ($linesOne.Count -ne 1 -or $linesOne[0] -ne '(one done)' -or $linesTwo.Count -ne 1 -or $linesTwo[0] -ne '(two done)') { $failures.Add("control: the two runners' lines were not each their own: $($linesOne -join ' ') / $($linesTwo -join ' ')") }
        else { Write-Output '  control: two runners started together run at once, each with its own lines' }
        # Kept: a fake runner that counts its runs in a file beside it.
        $cache = Join-Path $tmp 'kept'
        $counter = Join-Path $tmp 'runs.txt'
        $counting = Join-Path $tmp 'counting.ps1'
        $failing = Join-Path $tmp 'failing.ps1'
        $count = "Add-Content -LiteralPath (Join-Path `$PSScriptRoot 'runs.txt') 'ran'`r`n"
        [System.IO.File]::WriteAllText($counting, $count + (($good.ToArray() | ForEach-Object { "Write-Output '" + ($_ -replace "'", "''") + "'" }) -join "`r`n") + "`r`n", $utf8)
        [System.IO.File]::WriteAllText($failing, $count + "Write-Output '(recalc-differs ""chain: Output!R5C2: the machine holds 1, Calc 2"")'`r`n", $utf8)
        $judgeKept = {
            param([string]$exe, [string]$inputs)
            $k = Get-RunnerKept $exe 'recalc' $cache $inputs
            if ($null -ne $k.Lines) { return @{ From = 'kept'; Problems = @((Test-Runner $k.Lines $clean.Rows).Problems) } }
            $lines = @(Invoke-Runner $exe)
            $j = Test-Runner $lines $clean.Rows
            if ($j.Problems.Count -eq 0) { Save-RunnerKept 'recalc' $cache $k.Stamp $lines }
            return @{ From = 'run'; Problems = @($j.Problems) }
        }
        $runs = { if (Test-Path -LiteralPath $counter) { @(Get-Content -LiteralPath $counter).Count } else { 0 } }
        $steps = @(
            @{ What = 'a clean output, its first pass'; Exe = $counting; Inputs = 'X'; From = 'run'; Runs = 1; Clean = $true },
            @{ What = 'a clean output kept, read back'; Exe = $counting; Inputs = 'X'; From = 'kept'; Runs = 1; Clean = $true },
            @{ What = 'inputs of other bytes'; Exe = $counting; Inputs = 'Y'; From = 'run'; Runs = 2; Clean = $true },
            @{ What = 'a runner of other bytes'; Exe = $counting; Inputs = 'Y'; From = 'run'; Runs = 3; Clean = $true; Edit = $true },
            @{ What = 'an output that failed, its first pass'; Exe = $failing; Inputs = 'Y'; From = 'run'; Runs = 4; Clean = $false },
            @{ What = 'an output that failed, not kept'; Exe = $failing; Inputs = 'Y'; From = 'run'; Runs = 5; Clean = $false }
        )
        foreach ($step in $steps) {
            if ($step.ContainsKey('Edit')) { [System.IO.File]::AppendAllText($counting, "# another runner`r`n", $utf8) }
            $got = & $judgeKept $step.Exe $step.Inputs
            $n = & $runs
            $isClean = $got.Problems.Count -eq 0
            if ($got.From -ne $step.From -or $n -ne $step.Runs -or $isClean -ne $step.Clean) { $failures.Add("control: kept, $($step.What): $($got.From) after $n runs, clean $isClean, where $($step.From) after $($step.Runs), clean $($step.Clean)") }
            else { Write-Output ("  control: kept, {0}: {1}, {2} run(s) in all" -f $step.What, $(if ($got.From -eq 'kept') { 'read from the cache, not run' } else { 'run' }), $n) }
        }
    } finally { Remove-Item -Recurse -Force -LiteralPath $tmp -ErrorAction SilentlyContinue }
    if ($failures.Count -gt 0) { foreach ($f in $failures) { Write-Output "FAIL: $f" }; exit 1 }
    Write-Output 'OK: check_recalc_floors -Control: every clean case passes and every mutant fails'
    exit 0
}

# ------------------------------------------------------------------ the check
$problems = New-Object System.Collections.Generic.List[string]
$now = [ordered]@{}
foreach ($f in $fixtures) {
    $path = Join-Path $Root "scripts/recalc/bench/$f.txt"
    if (Test-Path -LiteralPath $path) { $now[$f] = Get-LfDigest $path }
}
foreach ($p in @(Test-Floors $baselineDate $baselineNative $baselineExcel $baseline $floorHistory $programs $now)) { $problems.Add($p) }
# The memory half (KERNEL.49): its rows, its fixtures of rows as they stand,
# and the ceilings.
$rowsNow = [ordered]@{}
foreach ($f in $rowsFixtures.Keys) {
    $path = Join-Path $Root "scripts/recalc/bench/$f.vla"
    if (Test-Path -LiteralPath $path) { $rowsNow[$f] = Get-LfDigest $path }
}
foreach ($p in @(Test-Memory $baseline $ceilingHistory $programs $rowsFixtures $rowsNow)) { $problems.Add($p) }

# The histories against the last commit's copy of this file.
$appended = 'SKIPPED (no git)'
$git = Get-Command git -ErrorAction SilentlyContinue
if ($git) {
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $committedText = & git -C $Root show 'HEAD:tools/check_recalc_floors.ps1' 2>$null | Out-String; $gitExit = $LASTEXITCODE } finally { $ErrorActionPreference = $old }
    if ($gitExit -eq 0 -and $committedText) {
        $self = [System.IO.File]::ReadAllText((Join-Path $Root 'tools/check_recalc_floors.ps1'))
        foreach ($p in @(Test-Appended (Get-HistoryLines $committedText) (Get-HistoryLines $self))) { $problems.Add($p) }
        foreach ($p in @(Test-Appended (Get-HistoryLines $committedText 'ceilingHistory') (Get-HistoryLines $self 'ceilingHistory'))) { $problems.Add("the ceilings' history: $p") }
        $appended = 'held against the last commit''s copy'
    } else { $appended = 'SKIPPED (not in the last commit yet)' }
}

# The two runners' check modes (the memory runner's since KERNEL.49): each
# named or found, a found one older than its sources skipped by name, its
# output from the cache while it and its inputs are unchanged, and the rest
# run side by side (AS.10).
$cacheDir = Join-Path $Root 'target/recalc_floors'
$inputsDigest = $null
$halves = @(
    @{ Kind = 'recalc'; Given = $Runner; Who = 'the runner' },
    @{ Kind = 'memory'; Given = $MemoryRunner; Who = 'the memory runner' }
)
foreach ($h in $halves) {
    $h.Exe = if ($h.Given -ne '') { $h.Given } else { Find-Runner $Root $h.Kind }
    $h.Skip = ''
    if ($h.Exe -eq '' -or -not (Test-Path -LiteralPath $h.Exe)) { $h.Skip = 'none built'; continue }
    $h.Exe = (Resolve-Path -LiteralPath $h.Exe).Path
    if ($h.Given -eq '') {
        $stale = Get-StaleSource $h.Exe $Root $runnerSources
        if ($stale -ne '') { $h.Skip = "$($h.Exe) was built before $stale last changed: build it again, cargo test --workspace or cargo build -p frazaro-core --example $($h.Kind)"; continue }
    }
    if ($null -eq $inputsDigest) { $inputsDigest = Get-InputsDigest $Root $runnerInputs }
    $kept = Get-RunnerKept $h.Exe $h.Kind $cacheDir $inputsDigest
    $h.Stamp = $kept.Stamp
    if ($null -ne $kept.Lines) { $h.Lines = $kept.Lines; $h.From = 'from the cache, the runner and its inputs unchanged' }
    else { $h.Started = Start-Runner $h.Exe; $h.From = 'run' }
}
foreach ($h in $halves) { if ($h.ContainsKey('Started')) { $h.Lines = @(Wait-Runner $h.Started) } }

$info = @()
$recalcHalf = $halves[0]
if ($recalcHalf.Skip -eq '') {
    $judged = Test-Runner $recalcHalf.Lines $baseline
    foreach ($p in $judged.Problems) { $problems.Add($p) }
    $info = @($judged.Info)
    if ($judged.Problems.Count -eq 0 -and $recalcHalf.From -eq 'run') { Save-RunnerKept $recalcHalf.Kind $cacheDir $recalcHalf.Stamp $recalcHalf.Lines }
    $runnerWord = "the runner's check mode, $($recalcHalf.From): every value held, the sums the measured ones"
} else { $runnerWord = "SKIPPED the runner ($($recalcHalf.Skip))" }

$memoryHalf = $halves[1]
if ($memoryHalf.Skip -eq '') {
    $judged = Test-MemoryRunner $memoryHalf.Lines $baseline $ceilingHistory
    foreach ($p in $judged.Problems) { $problems.Add($p) }
    $info = @($info) + @($judged.Info)
    if ($judged.Problems.Count -eq 0 -and $memoryHalf.From -eq 'run') { Save-RunnerKept $memoryHalf.Kind $cacheDir $memoryHalf.Stamp $memoryHalf.Lines }
    $memoryWord = "the memory runner's check mode, $($memoryHalf.From): every fixture's bytes at most its ceilings"
} else { $memoryWord = "SKIPPED the memory runner ($($memoryHalf.Skip))" }

# The bar: Excel beside the machine.
$rows = @($baseline | Where-Object { $null -ne $_ })
if ($rows.Count -gt 0) {
    Write-Output ("check_recalc_floors: the baseline of {0}, the machine against Excel on the same files:" -f $baselineDate)
    foreach ($f in $fixtures) {
        $n = @($rows | Where-Object { $_.Half -eq 'native' -and $_.Fixture -eq $f })
        $a = @($rows | Where-Object { $_.Half -eq 'excel' -and $_.Fixture -eq $f -and $_.Threads -eq 'all' })
        $o = @($rows | Where-Object { $_.Half -eq 'excel' -and $_.Fixture -eq $f -and $_.Threads -eq 'one' })
        if ($n.Count -ne 1 -or $a.Count -ne 1 -or $o.Count -ne 1) { continue }
        $ratio = { param($x, $y) if ([double]$y -le 0) { 'n/a' } elseif ([double]$x -le [double]$y) { '{0:F1}x faster' -f ([double]$y / [double]$x) } else { '{0:F1}x slower' -f ([double]$x / [double]$y) } }
        Write-Output ("  {0}: whole {1} ms, Excel {2} ms on {3} threads ({4}), {5} ms on one ({6}); the edit {7} ms, Excel {8} and {9} ms ({10}, {11})" -f $f, $n[0].Whole, $a[0].Whole, $a[0].ThreadCount, (& $ratio $n[0].Whole $a[0].Whole), $o[0].Whole, (& $ratio $n[0].Whole $o[0].Whole), $n[0].Edit, $a[0].Edit, $o[0].Edit, (& $ratio $n[0].Edit $a[0].Edit), (& $ratio $n[0].Edit $o[0].Edit))
    }
    foreach ($r in @($rows | Where-Object { $_.Half -eq 'excel-flag' })) { Write-Output ("  the flag: {0}, Output!C2 read {1} before any calculation: {2}; saved {3}" -f $r.File, $r.C2, $r.Verdict, $r.Saved) }
    # The memory half, a formula's bytes (KERNEL.49).
    foreach ($r in @($rows | Where-Object { $_.Half -eq 'memory' })) {
        $each = { param($b) '{0:F1}' -f ([double]$b / [double]$r.Formulas) }
        Write-Output ("  {0}: {1} formulas, {2} bytes a formula held after the load, {3} after two frames, {4} at the peak; the load {5} ms; its save {6} bytes a formula, a load from it {7} held and {8} at the peak" -f $r.Fixture, $r.Formulas, (& $each $r.LoadHeld), (& $each $r.Held), (& $each $r.Peak), $r.Load, (& $each $r.Save), (& $each $r.SavedHeld), (& $each $r.SavedPeak))
    }
}
foreach ($i in $info) { Write-Output $i }
if ($problems.Count -gt 0) {
    foreach ($p in $problems) { Write-Output "FAIL: $p" }
    exit 1
}
Write-Output ("OK: check_recalc_floors: {2} fixtures measured against Excel and {3} weighed, the floors and the ceilings held, the histories {0}; {1}; {4}" -f $appended, $runnerWord, $fixtures.Count, $weighed.Count, $memoryWord)
exit 0
