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

---

# Slice 2 - the first search

*The choice forms run: a program that chooses is grounded into atoms,
counters and clauses, and a search decides the rows in the order your
Tables list them, answering the first schedule that breaks no rule. Eleven
steps, 5 to 15. Every expected result below is also a pin in `TestDSLs`,
and every decision and dead-end count was checked before the pass against
a line-for-line transliteration of the search.*

**Slice 2 passed live 2026-09-25:** all eleven steps; every count and every
sentence as expected, and the `O3S|` lines and step 14's N (3,411,712) are
in the roadmap. Step 13's times were not whole 1/128 s ticks, as slice 1's
were: VBA's `Timer` is a `Single` of seconds since midnight, so its step is
set by the hour - 1/128 s after 18:12, far finer at 00:38 - not the machine.

**The one thing I am least sure of is step 5's compile.** The new code
passes an undeclared-variable lint and every static check, but it adds two
private types to `VLA_Optimize.bas` (one holding the other, and arrays of
objects), and nothing like that has compiled in this project before.

## 5. Reload, compile, and the versions

`VLA_DevRig.bas`'s own module list changed (a new module,
`VLA_OptimizeSearch`), and the rig cannot reload itself. So first, in the
VBA editor: right-click `VLA_DevRig` > `Remove VLA_DevRig...` > **No**,
then `File > Import File...` and choose `src\VLA_DevRig.bas`. Then, in the
Immediate window, `VlaDevReload`, and then `Debug > Compile VBAProject`.

Then:

```
?VLA_Optimize.VLA_OPTIMIZE_VERSION & " " & VLA_OptimizeSearch.VLA_OPTIMIZE_SEARCH_VERSION & " " & VLA_Messages.VLA_MESSAGES_VERSION & " " & VLA_Tests_Query.VLA_TESTS_QUERY_VERSION
```

**Expected:** compiles with no error, and the line reads
`OPTIMIZE.3 OPTIMIZE.3 OPTIMIZE.3 OPTIMIZE.3`. (The third is
`VLA_Messages.bas`'s, a file CLI.5's work in progress also edits: if that
peer has moved it to `CLI.5`, that is theirs and not a failure.)

**If the compile stops**, paste the highlighted line and the error back and
stop - it is the step I could not do before handing this over.

## 6. The suites

```
?VLA_Tests_Query.TestDSLs
?VlaSelfTest
?VlaSelfTestHost
```

and `VerifyReports` on both backends as you usually run it.

**Expected:** `TestDSLs` **2235 passed, 0 failed** - 2124 at slice 1, plus
**53** in the new `TestOptimizeSearch`, **56** in the new
`TestOptimizeChoice`, and **2** in `TestOptimizeForms` (78 -> 80).
`VlaSelfTestHost` **152/152** and `VerifyReports` **242/242** on both
backends. `VlaSelfTest` **1302/1302** when I last saw it - CLI.5's work in
progress shares this tree, and if it has landed pins, its number is the
one to expect. Nothing to paste from the middle of the run this time: only
the totals.

If anything FAILS, paste the FAILED TESTS block and stop.

## 7. The toy, in a cell

A fresh sheet. Paste this text into **A1**, as a value (it is text, not a
formula - and too long for a string inside a formula, which Excel caps at
255 characters):

```
(fact (person p1)) (fact (person p2)) (fact (person p3)) (fact (person p4)) (fact (person p5)) (fact (shift s1)) (fact (shift s2)) (fact (shift s3)) (fact (shift s4)) (fact (shift s5)) (fact (shift s6)) (fact (shift s7)) (fact (next s1 s2)) (fact (next s2 s3)) (fact (next s3 s4)) (fact (next s4 s5)) (fact (next s5 s6)) (fact (next s6 s7)) (rule (elig S P) (shift S) (person P)) (choose-exactly 2 (assign S P) (elig S P) (per (shift S))) (forbid (assign S P) (assign T P) (next S T)) (query assign)
```

In **C1**: `=OPTIMIZE(A1)`. In **F1**: `=OPTIMIZE_STATUS(A1)`.

**Expected:** C1:D15 spills a header row `S` `P`, then fourteen rows:
`s1 p1`, `s1 p2`, `s2 p3`, `s2 p4`, `s3 p1`, `s3 p2`, `s4 p3`, `s4 p4`,
`s5 p1`, `s5 p2`, `s6 p3`, `s6 p4`, `s7 p1`, `s7 p2` - five people, two on
each of seven shifts, nobody on two in a row, the first shift filled
first. F1 reads exactly:

> proven best: every rule holds, and nothing is being minimized or
> maximized, so no schedule is better than this one - it is the first that
> breaks no rule when the rows are decided in the order your Tables list
> them (14 decisions, 0 dead ends).

## 8. A count read from a Table column, over real Tables

A fresh sheet. In A1:B3 type the headers `Shift`, `Need` and the rows
`s1` `1`, `s2` `2`; select A1:B3, press **Ctrl+T** (my table has headers),
and with a cell of it selected, in the Immediate window:

```
Selection.ListObject.Name = "Shifts": ?Selection.ListObject.Name
```

It must read back `Shifts`. In D1:D4 type the header `Name` and the rows
`x`, `y`, `z`; select D1:D4, **Ctrl+T**, and:

```
Selection.ListObject.Name = "People": ?Selection.ListObject.Name
```

It must read back `People`. Paste this into **F1** as a value:

```
(rule (elig S P) (shifts S N) (people P)) (choose-exactly N (assign S P) (elig S P) (per (shifts S N))) (query assign)
```

In **H1**: `=OPTIMIZE(F1, Shifts, People)`.

**Expected:** H1:I4 spills `S` `P`, then `s1 x`, `s2 x`, `s2 y` - s1 needs
one and gets the first person, s2 needs two and gets the first two: the
killer case's own choice line, its count taken from each shift's own row.

## 9. The effort runs out

A fresh sheet. Paste into **A1** as a value - the toy of step 7 with
`(effort 3)` in it:

```
(fact (person p1)) (fact (person p2)) (fact (person p3)) (fact (person p4)) (fact (person p5)) (fact (shift s1)) (fact (shift s2)) (fact (shift s3)) (fact (shift s4)) (fact (shift s5)) (fact (shift s6)) (fact (shift s7)) (fact (next s1 s2)) (fact (next s2 s3)) (fact (next s3 s4)) (fact (next s4 s5)) (fact (next s5 s6)) (fact (next s6 s7)) (rule (elig S P) (shift S) (person P)) (choose-exactly 2 (assign S P) (elig S P) (per (shift S))) (forbid (assign S P) (assign T P) (next S T)) (effort 3) (query assign)
```

In **C1**: `=OPTIMIZE(A1)`. In **F1**: `=OPTIMIZE_STATUS(A1)`.

**Expected:** C1:D1 is the header row `S` `P` with nothing under it. F1
reads exactly:

> no schedule found within the budget; there may be one: (effort 3) allows
> 3 units of work - a decision or a dead end is one each - and all of them
> went on 3 decisions and 0 dead ends. More effort may find one: (effort
> thorough), or a larger number.

## 10. No schedule, by counting

A fresh sheet. Paste into **A1** as a value (section 17's
`optimize-seat-overflow`):

```
(fact (guest g1)) (fact (guest g2)) (fact (guest g3)) (fact (guest g4)) (fact (guest g5)) (fact (guest g6)) (fact (guest g7)) (fact (table ta)) (fact (table tb)) (rule (seat G T) (guest G) (table T)) (choose-exactly 1 (sit G T) (seat G T) (per (guest G))) (choose-at-most 3 (sit G T) (seat G T) (per (table T))) (query sit)
```

In **C1**: `=OPTIMIZE(A1)`. In **F1**: `=OPTIMIZE_STATUS(A1)`.

**Expected:** C1:D1 is `G` `T` with nothing under it. F1 reads exactly:

> no schedule satisfies every rule: there are 7 rows of 'guest' to place,
> and 'table' has 2 rows holding at most 3 each, which is 6.

This is OPTIMIZE.2's counting check, built then and reached by nothing
until now.

## 11. No schedule, and the rules that collide

A fresh sheet. Paste into **A1** as a value (`optimize-quote-conflict`: the
customer wants the carbon frame and the child seat, and the seat is not
rated for the frame):

```
(fact (catalogue alloy-frame frame)) (fact (catalogue carbon-frame frame)) (fact (catalogue standard-wheels wheels)) (fact (catalogue carbon-wheels wheels)) (fact (catalogue child-seat extra)) (fact (catalogue lights extra)) (fact (catalogue rack extra)) (fact (kind frame)) (fact (kind wheels)) (rule (extra I) (catalogue I extra)) (choose-exactly 1 (in I) (catalogue I K) (per (kind K))) (choose-any (in I) (extra I)) (require (in carbon-frame) (in carbon-wheels)) (forbid (in child-seat) (in carbon-frame)) (require (in child-seat)) (require (in carbon-frame)) (query in)
```

In **C1**: `=OPTIMIZE_STATUS(A1)`.

**Expected:** C1 reads exactly:

> no schedule satisfies every rule: these cannot all hold, before anything
> is chosen - check 2, (forbid (in child-seat) (in carbon-frame)); check 3,
> (require (in child-seat)); check 4, (require (in carbon-frame)).

Check 1 - carbon wheels need the carbon frame - is not named, because it
plays no part: the three that are named are the whole contradiction.

## 12. A requirement no choice can meet, listed

A fresh sheet. Paste into **A1** as a value:

```
(fact (item a)) (choose-any (pick I) (item I)) (require (pick b)) (query pick)
```

In **C1**: `=OPTIMIZE_VIOLATIONS(A1)`. In **G1**: `=OPTIMIZE_STATUS(A1)`.

**Expected:** C1:E2 spills `Check` `Rule` `Where`, then `1`,
`(require (pick b))`, `(no names - this rule is about particular rows)`. G1
reads exactly:

> no schedule satisfies every rule: check 1, (require (pick b)), is broken
> by 1 row. OPTIMIZE_VIOLATIONS lists them.

(`b` is in no pool, so nothing can be chosen to meet it.)

## 13. The search, timed

In the VBA editor: `File > Import File...` and choose
`tools\VLA_DiagO3.bas` again (slice 1's harness, with slice 2's added).
`Debug > Compile VBAProject`. Close every other workbook. Then:

```
O3Search
```

**Expected, exactly, on every `O3S|` line:** verdict `ok`, and these
counts:

| run | people | k | atoms | decisions | dead ends | outcome |
|---|---|---|---|---|---|---|
| search alone | 14 | 8 | 28 | 923 | 924 | none |
| search alone | 16 | 9 | 32 | 3431 | 3432 | none |
| search alone | 18 | 10 | 36 | 12869 | 12870 | none |
| search alone | 20 | 11 | 40 | 48619 | 48620 | none |
| family through OPTIMIZE | 16 | 9 | 32 | 3431 | 3432 | none |
| toy scaled, through OPTIMIZE | 20 | 4 | 420 | 84 | 0 | found |

The family is n people, two shifts, exactly k on each and nobody on both:
impossible, but not by counting, so the search has to try every way. The
seconds and the time per unit of work are this item's first measurement of
the search, and I have no prediction worth holding me to: anywhere from
about 5 to 50 microseconds a unit would not surprise me. **Paste every
`O3S|` line back** - they say what the provisional effort levels (5,000,
50,000 and 500,000 units) really are, in seconds, on this machine.

## 14. The seconds guard

**Excel will be busy for about ten seconds in this step.** A fresh sheet.
Paste into **A1** as a value - the timed family at 28 people and 15 a
shift, with an effort far past anything it could spend:

```
(fact (person p1)) (fact (person p2)) (fact (person p3)) (fact (person p4)) (fact (person p5)) (fact (person p6)) (fact (person p7)) (fact (person p8)) (fact (person p9)) (fact (person p10)) (fact (person p11)) (fact (person p12)) (fact (person p13)) (fact (person p14)) (fact (person p15)) (fact (person p16)) (fact (person p17)) (fact (person p18)) (fact (person p19)) (fact (person p20)) (fact (person p21)) (fact (person p22)) (fact (person p23)) (fact (person p24)) (fact (person p25)) (fact (person p26)) (fact (person p27)) (fact (person p28)) (fact (shift a)) (fact (shift b)) (rule (elig S P) (shift S) (person P)) (choose-exactly 15 (on S P) (elig S P) (per (shift S))) (forbid (on a P) (on b P)) (effort 2000000000) (query on)
```

In **C1**: `=OPTIMIZE_STATUS(A1)`.

**Expected:** after about ten seconds, C1 reads:

> no schedule found within the budget; there may be one: the search was
> stopped after 10 seconds by the guard that keeps a formula from holding
> Excel, having done N of the 2000000000 units of work (effort 2000000000)
> allows. Unlike the effort, where this stops depends on how fast the
> machine is.

where **N** is this machine's own - ten seconds' worth of work - and is
the one number in the sentence that is not fixed. The seconds read `10`:
the guard is read every 256 units of work and the timer steps by at most 1/128 s,
so it fires within a few hundredths of a second of ten, which rounds to
10. Then in **E1**: `=OPTIMIZE(A1)` - it answers at once, from the memo,
with the header row `S` `P` and nothing under it. **Paste C1's text back**:
N is a second measurement of the rate.

## 15. Remove the harness

In the VBA editor, right-click `VLA_DiagO3` > `Remove VLA_DiagO3...` > **No**.
