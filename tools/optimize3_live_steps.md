# OPTIMIZE.3 - live steps

*Slice 1 of five: the integer grounder. Nothing a user types reaches it
yet, so there is no worksheet step in this slice - the proof is the suite
and one timing harness. Four steps. Move this file to `archive/` once
OPTIMIZE.3 lands whole; the later slices' steps are appended below as they
are built.*

**Slice 1 passed live 2026-09-24:** all four steps; every count and every
`O3G|` line as expected, and the numbers are in the roadmap. Step 2's info
line scrolled out of the Immediate window's 200-line buffer, so a later
slice's steps ask you to paste only lines printed at the END of a run.

**Before you start.** `VlaDevReload`, then `Debug > Compile VBAProject`. No
module is new: `VLA_Datalog.bas`, `VLA_Relation.bas`, `VLA_Messages.bas` and
`VLA_Tests_Query.bas` changed, and nothing was added to or removed from
either `mods` array. `VLA_Optimize.bas` did not change in this slice.

**`VlaSelfTest` is 1302 now, not the 1215 the brief quoted.** Two peers'
items landed while this was built - CLI.3 (+45, to 1260) and CLI.4 (+42, to
1302, its own commit's record) - and nothing `VlaSelfTest` reaches changed
here except `VLA_Messages.bas`, which gained one entry and is read by no
catalogue loop. If a later peer lands pins before you run this, theirs is
the number to expect.

Predicted before the pass, so you can hold me to it:

- `TestDSLs` **2124 passed, 0 failed**, counted from the source: 1881 at
  DATALOG.14, plus **241** in the new `TestDatalogGroundRules` (47
  hand-derived pins, one per program of the parity table's 192, and 2
  floors), plus **2** in `TestOptimizeParity`, whose table grew from 190
  programs to 192. If the total differs but nothing FAILS, my arithmetic is
  wrong and the code is not;
- `VlaSelfTestHost` **152/152** and `VerifyReports` **242/242** on both
  backends, unmoved; `VlaSelfTest` **1302/1302**, unmoved (see above);
- **the one I am least sure of is step 3's timing**, because no integer
  grounder has ever been timed on this machine. The row counts, the
  agreement and the path in step 3 are exact; the seconds are estimates.

---

## 1. It compiles, and the versions moved

In the VBA editor: `Debug > Compile VBAProject`. Then in the Immediate
window:

```
?VLA_Datalog.VLA_DATALOG_VERSION & " " & VLA_Relation.VLA_RELATION_VERSION & " " & VLA_Messages.VLA_MESSAGES_VERSION & " " & VLA_Tests_Query.VLA_TESTS_QUERY_VERSION & " " & VLA_Optimize.VLA_OPTIMIZE_VERSION
```

**Expected:** compiles with no error, and the line reads
`OPTIMIZE.3 OPTIMIZE.3 OPTIMIZE.3 OPTIMIZE.3 OPTIMIZE.2` - the last one
unmoved, because `VLA_Optimize.bas` is untouched until slice 2.

**If the compile stops** on a line naming `VlaSymbols`: that is a
`Public Type` declared in `VLA_Relation.bas` and passed ByRef between
standard modules - the first `Public Type` in this project. Paste the
highlighted line and the error back; it is the one construct in this slice
I could not compile before handing it over.

## 2. The suites

Immediate window, one at a time:

```
?VLA_Tests_Query.TestDSLs
?VlaSelfTest
?VlaSelfTestHost
```

and then `VerifyReports` on both backends as you usually run it.

**Expected:** `TestDSLs` **2124 passed, 0 failed**, run to the end. Among
its lines there is one info line - not an assertion - reading:

```
  info: ground parity compared N rules, M of them in integers
```

**Paste that line back.** I counted about 67 rules and about 50 in integers
by hand from the parity table; the two floor pins require at least 40 and
25, so they pass anywhere near my count, and the real numbers go into the
roadmap. `VlaSelfTest` 1302/1302, `VlaSelfTestHost`
152/152, `VerifyReports` 242/242 on both backends.

If anything FAILS, paste the FAILED TESTS block and stop.

## 3. The grounder, timed against DATALOG's own evaluator

In the VBA editor: `File > Import File...` and choose
`tools\VLA_DiagO3.bas` (it is not in either `mods` array and never ships).
`Debug > Compile VBAProject`. **Close every other workbook first** - nothing
here recalculates, but anything else live competes for the machine.

Then, one at a time, in the Immediate window:

```
O3Ground 6
O3Ground 8
O3Ground 10
```

Each prints a header, one `O3G|` line per rung, and a done line. Each rung
builds the fixture in memory (no worksheet is written), times DATALOG's own
evaluator (`DatalogRunForms`, the fixpoint) and the integer grounder
(`DatalogGroundRules`) on the SAME rule over the SAME relations, and checks
three things before it reports a time: both made the row count the
fixture's closed form says; they agree on every row, in order; and the
integer grounder really took the integer path. The counter behind `path` is
reset by each call to `DatalogGroundRules` and nothing else moves it, so it
reads the rung it is printed beside.

**Expected, exactly, on every line:** verdict `ok`, `same rows, same order`
`yes`, path `integer`, and these row counts in both the `datalog rows` and
`integer rows` columns:

| step | rungs (people x weeks, k) | rows, in rung order |
|---|---|---|
| 6 | 10x1, 20x4, 50x4 | 200 · 1,660 · 4,150 |
| 8 | 10x1, 14x1, 20x1, 10x4, 14x4, 20x4 (k = 2) | 2,520 · 7,644 · 23,940 · 10,080 · 30,576 · 95,760 |
| 10 | 5x1 k2, 10x1 k2, 5x1 k3, 10x1 k3 | 6,650 · 13,300 · 29,925 · 59,850 |

**The seconds are estimates, and it is the ratio that matters.** For
reference, the pre-flight timed DATALOG on these rungs through a worksheet
formula, which also paid for reading the Table and spilling the answer:
step 8 at 20x4 **9.039s**, step 10 at 10x1 k3 **8.063s**, step 6 at 50x4
**0.188s**. This harness times DATALOG as a VBA call, so expect its DATALOG
column at or a little under those. My estimate for the integer column at
the two big rungs is **0.5 to 1.5 s**, most of it the comparisons (each
still calls the shared `CompareValues`, which keeps the two evaluators'
answers identical by construction); anything under 3 s is the order of
magnitude the fork was decided on. Step 6's small rungs may read 0.000 in
the integer column - under the timer's floor, and not a failure.

**Paste every `O3G|` line back.** They are the grounding half of the
before-and-after this item records against OPTIMIZE.0's ladder.

A rung reading `WRONG`, `DISAGREE` or `FELL BACK` is a finding, and the step
stops there by design: paste the line and stop.

## 4. Remove the harness

In the VBA editor, right-click `VLA_DiagO3` > `Remove VLA_DiagO3...` > **No**
(do not export). It is a diagnostic and nothing references it.
