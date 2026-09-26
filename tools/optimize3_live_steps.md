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

---

# Slice 3 - the ceilings

*A formula now lays out at most 100,000 rows of OPTIMIZE's own grounding,
and at most 50,000 in any one step; a program past either is refused,
naming the rule or the choice form and the number, and the refusal is
memoized. Each step is counted exactly before a row of it is made. The
rule over chosen rows is joined smallest part first, which is what lets
the reference roster's "never two in a row" fit at all. Plus one
separate fix, first: a message's dash. Ten steps, 16 to 25. Every
number below was checked before the pass against a transliteration of
the grounder and, where it searches, of the search.*

**Slice 3 passed live 2026-09-25, in two rounds:** every count and every
sentence as expected. Round 2's `TestDSLs` read 2264/0 again, and its
`O3C|` lines are in the roadmap: grounding at the ceiling 0.05-0.16 s, a
whole run at it 0.40 s.

**Round 1, 2026-09-25:**
- Steps 17 to 23 passed as expected.
- Step 16 ran on a workbook already reloaded with the fix, so it showed the
  hyphen. The old dash was never seen.
- Step 24's counts were exact, but two of its rungs took 16.6 s and 48.5 s.
  Three hash tables had gone quadratic at the ceiling, one of them slice 2's
  atom table. The fix is in, and round 2 (steps 26 to 29) re-times it.

**Two steps build their program with a formula.** The programs are
hundreds of facts long, so A1 builds the text with `TEXTJOIN` and
`SEQUENCE` rather than you pasting it. A1 shows the program text; that
is expected.

## 16. The dash, before you reload

Do this **before** step 17 - it looks at the code the workbook has now.
In the Immediate window:

```
VLA_Messages.RaiseMsg "english-unexpected-token-block", "tok", "done"
```

**Expected:** VBA's own error dialog, `Run-time error '5':`, reading
`Unexpected 'done' â€” is there a stray 'Done.' or a missing block?` -
three stray characters, `â€”`, where a dash belongs. Click **End**. That
message's source held a real em dash, stored as three bytes of UTF-8,
and VBA reads a module as Windows text, one character per byte. If the
dialog shows a proper dash instead, tell me: the release note says it
did not.

## 17. Reload, compile, and the versions

No module is new in this slice, and `VLA_DevRig.bas` did not change, so
`VlaDevReload` alone, then `Debug > Compile VBAProject`. Then:

```
?VLA_Datalog.VLA_DATALOG_VERSION & " " & VLA_Optimize.VLA_OPTIMIZE_VERSION & " " & VLA_Messages.VLA_MESSAGES_VERSION & " " & VLA_Tests_Query.VLA_TESTS_QUERY_VERSION
```

**Expected:** compiles with no error, and
`OPTIMIZE.3 OPTIMIZE.3 OPTIMIZE.3 OPTIMIZE.3` - unchanged since slice 2,
so here the compile is the check. (`VLA_Messages.bas` is CLI.5's file
too: if that peer has moved its version, that is theirs.) If the compile
stops, paste the highlighted line and the error and stop.

## 18. The suites

```
?VLA_Tests_Query.TestDSLs
?VlaSelfTest
?VlaSelfTestHost
```

and `VerifyReports` on both backends.

**Expected:** `TestDSLs` **2264 passed, 0 failed** - 2235 at slice 2,
plus **20** in the new `TestDatalogGroundCeilings`, **8** in the new
`TestOptimizeCeilings`, and **1** in `TestOptimizeChoice`. `VlaSelfTest`
**1347/1347** when I last saw it (CLI.5's pins, not this slice's);
`VlaSelfTestHost` **152/152**; `VerifyReports` **242/242** on both.
Only the totals to paste.

## 19. The dash, after

The same line as step 16:

```
VLA_Messages.RaiseMsg "english-unexpected-token-block", "tok", "done"
```

**Expected:** the dialog reads
`Unexpected 'done' - is there a stray 'Done.' or a missing block?` - a
plain hyphen, like every other message. Click **End**.

## 20. A rule that pairs too many

A fresh sheet. In **A1**, this formula (it builds a program of 300 items,
any of them picked, and a rule over two picks that share no name):

```
=TEXTJOIN(" ",TRUE,"(fact (item i"&SEQUENCE(300)&"))")&" (choose-any (pick X) (item X)) (forbid (pick X) (pick Y)) (query pick)"
```

In **C1**: `=OPTIMIZE_STATUS(A1)`.

**Expected:** C1 reads exactly:

> #OPTIMIZE! (forbid (pick X) (pick Y)) pairs more rows than a formula
> lays out in one step: the 300 rows of 'pick' share no name with the 300
> before it, so every one pairs with every one, making 90,000 rows, and a
> formula lays out at most 50,000. A condition that ties its rows together
> pairs fewer, and so does a smaller pool.

It answers at once: the 90,000 are counted, never made.

## 21. A choice with too many rows to choose from

A fresh sheet. In **A1**:

```
=TEXTJOIN(" ",TRUE,"(fact (item i"&SEQUENCE(250)&"))")&" "&TEXTJOIN(" ",TRUE,"(fact (slot t"&SEQUENCE(201)&"))")&" (choose-at-most 250 (pick X S) (item X) (per (slot S))) (query pick)"
```

In **C1**: `=OPTIMIZE_STATUS(A1)`.

**Expected:** C1 reads exactly:

> #OPTIMIZE! (choose-at-most 250 (pick X S) (item X) (per (slot S))) has
> more rows to choose from than a formula lays out in one step: the 201
> rows of 'slot' share no name with the 250 before it, so every one pairs
> with every one, making 50,250 rows, and a formula lays out at most
> 50,000. Choose from a smaller pool - a rule that keeps only the rows
> that could really be chosen.

## 22. A program past the total

A fresh sheet. In **A1** - step 21's formula with **200** slots instead
of 201:

```
=TEXTJOIN(" ",TRUE,"(fact (item i"&SEQUENCE(250)&"))")&" "&TEXTJOIN(" ",TRUE,"(fact (slot t"&SEQUENCE(200)&"))")&" (choose-at-most 250 (pick X S) (item X) (per (slot S))) (query pick)"
```

In **C1**: `=OPTIMIZE_STATUS(A1)`.

**Expected:** C1 reads exactly:

> #OPTIMIZE! this program lays out more rows than a formula may: with
> (choose-at-most 250 (pick X S) (item X) (per (slot S))), its choices and
> the rules over them reach 100,650 rows, and a formula lays out at most
> 100,000 in all. Smaller pools, or rules that pair fewer rows, lay out
> less.

Its one step of 50,000 is AT the step ceiling, not past it, so it is
made; the 50,000 members after it take the total past 100,000. So unlike
steps 20 and 21 it is refused only after laying out 100,650 rows, and
may not answer at once - step 24 times what that costs.

## 23. The reference roster's own shape, answered

A fresh sheet. 50 people, 84 shifts - four weeks of three shifts a day,
the reference roster's size - exactly two a shift, never two in a row.
In **A1**:

```
=TEXTJOIN(" ",TRUE,"(fact (person p"&SEQUENCE(50)&"))")&" "&TEXTJOIN(" ",TRUE,"(fact (shift s"&SEQUENCE(84)&"))")&" "&TEXTJOIN(" ",TRUE,"(fact (next s"&SEQUENCE(83)&" s"&SEQUENCE(83)+1&"))")&" (rule (elig S P) (shift S) (person P)) (choose-exactly 2 (assign S P) (elig S P) (per (shift S))) (forbid (assign S P) (assign T P) (next S T)) (query assign)"
```

In **C1**: `=OPTIMIZE(A1)`. In **F1**: `=OPTIMIZE_STATUS(A1)`.

**Expected:** C1 spills the header `S` `P` and 168 rows, from `s1 p1`,
`s1 p2`, `s2 p3`, `s2 p4`, `s3 p1`, `s3 p2` down to `s84 p3`, `s84 p4`
in C168:D169. F1 reads exactly:

> proven best: every rule holds, and nothing is being minimized or
> maximized, so no schedule is better than this one - it is the first that
> breaks no rule when the rows are decided in the order your Tables and
> facts list them (168 decisions, 0 dead ends).

- "Tables and facts" is the reworded status: slice 2's step 7 said "your
  Tables" of a program that had none.
- Before this slice, the forbid in its own written order would have paired
  every shift of each person with every other: 352,800 rows in one step,
  and this sheet would be refused. Joined smallest part first, the whole
  grounding lays out 25,301 rows.

## 24. The ceilings, timed

In the VBA editor, `File > Import File...` and choose
`tools\VLA_DiagO3.bas` again (slice 3's rungs are added to it).
`Debug > Compile VBAProject`. Close every other workbook. Then:

```
O3Ceiling
```

**Expected, exactly, on every `O3C|` line:** verdict `ok` and these
counts:

| run | rows laid out | largest step | step refused |
|---|---|---|---|
| member rule, 250 items x 198 slots | 99,250 | 49,500 | - |
| never two in a row, planned, 60 x 553 | 99,912 | 33,120 | - |
| never two in a row, written, 60 x 553 | 33,180 | 33,180 | 18,348,540 |
| OPTIMIZE, 250 items x 198 slots | 99,646 | 49,500 | - (and 49,500 decisions) |

The first two are the integer grounder alone at a formula's ceilings, one
per pass; the third is the rule this slice plans, in its own written
order, refused - its time is the first step plus counting 18 million
rows, which it never makes; the last is a whole run at the ceiling, its
search and its answer included. The seconds are this slice's own first
measurement of what the ceilings cost on the integer grounder. My guess is
about a tenth to a fifth of a second for each of the first two, and under
a second for the whole run, but I am not holding you to it. **Paste every
`O3C|` line back.**

## 25. Remove the harness

In the VBA editor, right-click `VLA_DiagO3` > `Remove VLA_DiagO3...` > **No**.

---

## Slice 3, round 2 - after the hash fix

*Step 24's first run was exact on every count, but its planned rung took
16.6 s and its whole run 48.5 s for rows the member rung laid out in
0.034 s. Three hash tables probed in line on a hash that packed a two-part
key into neighbouring slots, so every new key walked a run that grew with
the table:*
- *the integer grounder's key counter;*
- *its row sets;*
- *slice 2's atom table.*

*A simulation of the rungs' own keys counted 246 million and 1.5 billion
probes. All eleven hash sites now use one step that mixes each id, which
the same simulation puts at a random hash's probe rate. No answer can
change: those tables keep their rows in insertion order and only answer
whether a key is there. Four steps.*

## 26. Reload and compile

`VLA_Datalog.bas` and `VLA_Optimize.bas` changed; no module was added. So
`VlaDevReload`, then `Debug > Compile VBAProject`. It must compile.

## 27. The suite again

```
?VLA_Tests_Query.TestDSLs
```

**Expected:** **2264 passed, 0 failed**, the same as step 18. The hash
decides where a key sits, never which rows come out or in what order, and
`TestDatalogGroundRules` grounds every rule of the parity table both ways
again.

## 28. The ceilings, timed again

If you removed `VLA_DiagO3` at step 25, import `tools\VLA_DiagO3.bas`
again (it has not changed) and compile. Close every other workbook. Then:

```
O3Ceiling
```

**Expected:** exactly step 24's counts on all four lines, each `ok`. The
seconds are the point:
- **Rungs 2 and 4 should fall a long way from 16.566 and 48.487.**
- **Rung 2** should come in near rung 1's few hundredths of a second, with
  two counting passes on top.
- **Rung 4** also searches 49,500 decisions (about 0.12 s at slice 2's rate)
  and builds a 49,500-row answer, so my guess is under two seconds. I have
  not measured it.
- Rungs 1 and 3 may move a little either way: the new hash costs a few more
  operations a key and saves nothing there.

**Paste every `O3C|` line back.**

## 29. Remove the harness

In the VBA editor, right-click `VLA_DiagO3` > `Remove VLA_DiagO3...` > **No**.

---

# Slice 4 - the ladder

*OPTIMIZE.0's reference roster, searched for the first time. It has 50
people over 4 weeks of 3 shifts a day, a fifth of the people on every
shift, and one day's leave a person a week. Every night needs a senior,
nobody works more than five shifts a week, and nobody works two in a row.
The ladder runs it at 5, 10, 20 and 50 people over 1 and 4 weeks, written
as an OPTIMIZE program over its four Tables. Every run goes through the
whole engine: DATALOG's pass, the grounding, the search and the answer.
The roster has no objective and no kept schedule yet (OPTIMIZE.6 and .7),
so each answer is the first schedule in the Tables' own order.*

*Every count below was computed before the pass: the search translated
line for line into another language, over the roster built from its
definition. That translation reproduces every count measured live so far:
slice 2's four family rungs and its scaled toy, step 23's roster, and the
toy. Every schedule the harness gets back is checked against the roster's
own definition by the harness itself, not by OPTIMIZE. No `src` module
changed, so there are no suites to run this round. Two steps.*

**Round 1 passed live 2026-09-25:** all eleven lines `ok`, every count as
predicted. The reference roster took 0.745 s a run, against my guess of
0.3 s. A unit of search work took 3.70 us at `normal` and 4.35 us at
`thorough`, against my guess of 5, so `thorough`'s 500,000 units took
2.92 s in all: past a formula's 2 s.

## 30. The ladder

In the VBA editor, `File > Import File...` and choose
`tools\VLA_DiagO3.bas` (slice 4's rungs are added to it).
`Debug > Compile VBAProject`. Close every other workbook. Then:

```
O3Ladder
```

It takes under a minute, and Excel does not respond until it is done: each
rung runs again and again for two seconds, so the small ones are timed well
clear of `Timer`'s steps.

**Expected, exactly, on every `O3L|` line:** verdict `ok` and these counts.

| rung | a shift | atoms | counters | clauses | rows laid out | largest step | decisions | dead ends | answer |
|---|---|---|---|---|---|---|---|---|---|
| 5 x 1 | 1 | 90 | 33 | 80 | 889 | 90 | 0 | 0 | none |
| 5 x 4 | 1 | 360 | 132 | 335 | 3,604 | 360 | 0 | 0 | none |
| 10 x 1 | 2 | 180 | 38 | 162 | 1,707 | 180 | 35 | 2 | a schedule |
| 10 x 4 | 2 | 720 | 152 | 672 | 6,906 | 720 | 140 | 8 | a schedule |
| 20 x 1 | 4 | 360 | 48 | 325 | 3,358 | 360 | 80 | 2 | a schedule |
| 20 x 4 | 4 | 1,440 | 192 | 1,345 | 13,579 | 1,440 | 320 | 8 | a schedule |
| 50 x 1 | 10 | 900 | 78 | 814 | 8,277 | 900 | 200 | 0 | a schedule |
| 50 x 4 | 10 | 3,600 | 312 | 3,364 | 33,456 | 3,600 | 800 | 0 | a schedule |
| 50 x 4 at 11, (effort 1) | 11 | 3,600 | 312 | 3,364 | 33,456 | 3,600 | 1 | 0 | none within the effort |
| 50 x 4 at 11, (effort normal) | 11 | 3,600 | 312 | 3,364 | 33,456 | 3,600 | 25,084 | 24,916 | none within the effort |
| 50 x 4 at 11, (effort thorough) | 11 | 3,600 | 312 | 3,364 | 33,456 | 3,600 | 250,084 | 249,916 | none within the effort |

What the rows mean:
- **5 people:** no schedule, found by counting before any search. Their one
  senior, P4, is on leave on day 3, so nobody can take that night. The
  harness also checks the status reads exactly: "no schedule satisfies
  every rule: night S = 9 needs 1 row from 'senior-free', and only 0 rows
  can fill it."
- **10 to 50 people:** a schedule each time, and the harness checks every
  row of it. The full reference roster, 50 x 4, takes **800 decisions and
  no dead end**, and lays out 33,456 rows, a third of a formula's ceiling.
- **The last three rows are the calibration.** They use the same roster
  with 11 a shift instead of 10. The search never gets out of week 1:
  - its first dead end comes 175 decisions in, on day 6's early shift;
  - after that, backtracking one decision at a time never reaches further
    back than the 147th decision, on day 5's late shift;
  - so every effort level runs out, and each run's search is exactly its
    effort's units of work.

  `(effort 1)` times everything but the search. The last `O3L|` line
  divides what is left by the units: that is a unit of search work on a
  real roster, which is what the effort levels get set from.

The seconds are this slice's own measurement. My guesses, which I am not
holding you to: about 0.3 s a run at 50 x 4, and a unit of search work
around 5 us. That would put `thorough` near 2.5 s on this roster, past the
2 s a formula is allowed, which is the question the calibration is here to
answer. If the thorough line says `GUARD` instead of `ok`, a unit costs
more than 20 us here, and that is the answer too.

**Paste every `O3L|` line back.**

## 31. Remove the harness

In the VBA editor, right-click `VLA_DiagO3` > `Remove VLA_DiagO3...` > **No**.

---

## Slice 4, round 2 - the effort levels, set

**Round 2 passed live 2026-09-25:**
- **Suites and reload.** The reload showed ` 2500  25000  250000`.
  `TestDSLs` gave **2265/0**, as predicted. `VlaSelfTest` 1388/1388 and
  `VlaSelfTestHost` 194/194 moved by the other session's pins.
  `VerifyReports` gave 267/267 on both backends, moved by that session's
  new sentences and their checks, none of this slice's.
- **The ladder.** Every `O3L|` line `ok` and every count exact, including
  the fixed budget: 125,087 and 124,913. `thorough` took **1.871 s** a
  run, where it took 2.92 s, and a unit of search work took 4.47 to
  4.48 us.
- **Where the reference roster's run goes:** 0.495 s is DATALOG's own
  pass, 0.095 s the memo key, and 0.170 s OPTIMIZE's own work.

*Your call, 2026-09-25: halve all three. `quick` is now 2,500 units,
`normal` 25,000 and `thorough` 250,000. At round 1's 4.35 us a unit,
`thorough` is about 1.1 s of search on the reference roster, and with
the 0.75 s everything else takes there, 1.8 s in all: inside a
formula's 2 s. Each level is still a tenth of the next.*

*Setting the new levels turned up a defect from slice 2. A search could
do one unit more than its budget: when the budget's last unit went on a
decision that ran straight into a dead end, the dead end was counted as
well. The status then said "allows 250000 units of work ... and all of
them went on 125087 decisions and 124914 dead ends", which adds up to
250,001. Round 1 never showed it, because both of its budgets happened
to end on a dead end, not a decision. The search now stops at such a
dead end without taking it, so it never does more work than it was
given. Slice 2's own pin had the overshoot written into it, and it is
re-pointed.*

*Another session's work is in the tree as well: `VLA_Runtime`,
`VLA_Interpreter`, `VLA_Tests`, `VLA_Tests_Host` and the English
phrasebook. `VlaDevReload` loads it along with this slice, so
`VlaSelfTest` and `VlaSelfTestHost` will count its pins. This slice adds
none there. If a compile stops in one of those modules, that is the
other session's work in progress. Four steps.*

## 32. Reload and compile

`VLA_Optimize.bas`, `VLA_OptimizeSearch.bas` and `VLA_Tests_Query.bas`
changed; no module was added. `VlaDevReload`, then
`Debug > Compile VBAProject`. Then:

```
?VLA_Optimize.VLA_OPTIMIZE_WORK_QUICK; VLA_Optimize.VLA_OPTIMIZE_WORK_NORMAL; VLA_Optimize.VLA_OPTIMIZE_WORK_THOROUGH
```

**Expected:** ` 2500  25000  250000`. That shows the reload took the new
levels. If the compile stops, paste the highlighted line and the error,
then stop.

## 33. The suites

```
?VLA_Tests_Query.TestDSLs
?VlaSelfTest
?VlaSelfTestHost
```

and `VerifyReports` on both backends.

**Expected:** `TestDSLs` **2265 passed, 0 failed**. That is 2264 plus
**1**: the budget is now pinned at its edge from both sides. With one
unit it stops at the decision, and with two the dead end is the second
unit. Two more pins changed what they hold: the three levels, and slice
2's "the budget stops it" pin, which now reads `budget d1 c0` instead of
`budget d1 c1`. `VlaSelfTest` and `VlaSelfTestHost` will show whatever
the other session's pins make them (1358 and 152 before it).
`VerifyReports` **242/242** on both. Only the totals to paste.

## 34. The ladder again

`File > Import File...` and choose `tools\VLA_DiagO3.bas` again (round
2's counts and a new last line are in it). `Debug > Compile
VBAProject`. Close every other workbook. Then:

```
O3Ladder
```

**Expected, exactly:** the eight rungs as in step 30, every count the
same, since none of them spends more than 800 units of the new
`normal`'s 25,000. The calibration now runs the new levels:

| rung | decisions | dead ends | answer |
|---|---|---|---|
| 50 x 4 at 11, (effort 1) | 1 | 0 | none within the effort |
| 50 x 4 at 11, (effort normal) | 12,583 | 12,417 | none within the effort |
| 50 x 4 at 11, (effort thorough) | 125,087 | 124,913 | none within the effort |

125,087 and 124,913 make exactly 250,000. Without the fix the last row
would read 124,914 dead ends. Every line should say `ok`.

The seconds are the point:
- **The thorough line should come in near 1.8 s**, where round 1 took
  2.92 s: 0.75 s plus 250,000 units at about 4.35 us each.
- **The unit rates should repeat round 1's**, about 3.7 to 4.4 us.
- **The last line is new: where the reference roster's own run goes.**
  It splits the run into the memo key, DATALOG's own pass over the six
  rules, and the rest, which is OPTIMIZE's own work. Round 1's 0.75 s
  was 2.5 times my guess, so I don't know the split. My guesses are
  about 0.01 s for the key, 0.2 to 0.3 s for DATALOG's pass, and the
  remainder for OPTIMIZE.

**Paste every `O3L|` line back**, with the four suite totals from step 33.

## 35. Remove the harness

In the VBA editor, right-click `VLA_DiagO3` > `Remove VLA_DiagO3...` > **No**.
