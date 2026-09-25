# Pre-flight before OPTIMIZE.3

Two measurements, neither a build. Run them before a session breaks ground on
`OPTIMIZE.3`, and hand that session the numbers.

Move this file to `archive/` once `OPTIMIZE.3` starts.

## Why these two

**`OPTIMIZE.3` sets its own ceilings, on a basis that is now stale.** Its
entry says, in its own words:

> *"The measured basis, **while grounding goes through `DATALOG`'s
> relations**: linear under about 50,000 rows in the largest relation,
> superlinear past about 90,000, a memory risk past 500,000."*

That was measured 2026-09-19. Grounding does go through `DATALOG`'s
relations — `OPTIMIZE.2` runs every constraint through `DatalogRunForms` as
a rule — and `DATALOG.14` (2026-09-24) made that path roughly an order of
magnitude cheaper. A ceiling set on the old numbers refuses programs the
engine can now ground. The entry already says "re-measure before raising
them"; it just expected the change to come from integer arrays rather than
from underneath.

**And the 100,000-row claim is currently a projection.** `DATALOG.14`'s
entry and `RELEASES.md` 0.6.3 both say a 100,000-row scan works out at about
3.2s, extrapolated from the 1,000 and 10,000 rungs. It is flagged as a
projection in both places and should not be repeated anywhere a user reads
it until it has been run once.

---

## Step 1 — the 100,000-row rung — **DONE 2026-09-24**

**Result: 3.676s first, 3.555s after the change, 49,001 rows, `ok`.** Inside
the 10s limit, and 10.3× better than the 38s `DATALOG.13` projected. The
roadmap and `RELEASES.md` 0.6.3 are corrected from projection to measurement.

**It also found something step 2 should expect.** The scan is NOT perfectly
flat: marginal cost is 0.0291 ms a row from 3,000 to 10,000 and **0.0374 from
10,000 to 100,000, a 29% rise**. `O0DictCost` measured a 57% rise in per-key
dictionary cost over exactly that range, which accounts for about a fifth of
it by arithmetic and no more; the rest is unexplained and was not guessed at.
**So there is already one measurement saying the knee near 50,000–90,000 rows
is still there** — which is what step 2 predicts it will find, from a
different shape.

The instructions are kept below for the record.

### (as written, before it was run)

Do this one first: it is short, and it needs no new module.

**Before you start:** close every other workbook. `Application.Calculate`
recalculates volatile formulas everywhere, and anything else live lands in
these timings. Sheet `D12T1` and Table `Bills12x1` must exist — if they do
not, run the `Interpret` block from `archive/datalog12_live_steps.md` step 1
first.

**Type:** `D12Ladder 1, , 10000`

That re-establishes the 10,000 rung as a baseline in this session
(**expect ~0.33s, 4,901 rows, `ok`**). Then:

**Type:** `D12Ladder 1, , 100000`

**Expect:** the ladder runs 100 / 300 / 1,000 / 3,000 / 100,000. The guard
projects the top rung from the one before at about 3.6s, well under its 15s
limit, so it should run rather than be skipped.

| rows | expect |
|---|---|
| 100,000 | first answer **near 3.2s**, answer rows **49,001**, verdict `ok` |

49,001 is 49,000 big bills per 100,000 rows plus the one the ladder's own
cell change makes big. The Table write itself is not timed — only
`Application.Calculate` is — so a few seconds of Excel building the fixture
is expected and is not the measurement.

**If it comes in materially above 3.2s**, the scan is not as flat at
100,000 rows as it was between 1,000 and 10,000, and that is the finding:
send the line back. **If it refuses or errors**, that is a bigger finding —
no `DATALOG` answer has ever been asked for at this size.

**Then:** correct the 100,000-row sentence in `docs/BETA_ROADMAP1.md`
(`DATALOG.14`'s scan block) and in `docs/RELEASES.md` 0.6.3 from a
projection to a measurement, or to whatever it actually turns out to be.

---

## Step 2 — the OPTIMIZE.0 grounding ladder, re-run

**Before you start:**

1. Import `archive/VLA_DiagO0.bas` into `VLA.xlsm` and run Debug > Compile.
2. **Delete the sheets `O0T1`, `O0T4`, `O0T5`, `O0T6`, `O0T8`, `O0T10` and
   `O0T12` if they exist.** Each step refuses to reuse its own sheet, so a
   left-over sheet from 2026-09-19 stops the re-run rather than corrupting
   it. Results also append to `O0Results`, so the old rows stay — compare
   against the timestamps, not against row order.
3. Keep `archive/optimize0_live_steps.md` open. It carries the expected
   **ground rows** for every rung, and those must not move: this is a
   re-timing, not a re-derivation. A changed row count is a defect, not a
   speed-up.

**The two controls first, because they must NOT move.** If either does, the
machine changed and the rest of the run means nothing.

| type | what it is | expect |
|---|---|---|
| `O0Model` | the paper model, no engine at all | 104 `O0M\|` lines, exactly **9** ending `never run: memory` |
| `O0Ladder 12` | a bare UDF returning an array, no engine | every rung `ok` |

### Both controls RUN 2026-09-24

**`O0Model`: green, exactly.** 104 lines, exactly 9 `never run: memory`, and
on the nine rungs the steps doc names (step 8 at 50×1 k=2, 50×4 k=2 and 50×4
k=10; step 9 at 20×4, 50×1 and 50×4; step 10 at 20×1 k=3, 5×1 k=5 and 50×4
k=5). The model is pure arithmetic with no engine in it, so this says the
harness is intact and nothing drifted.

**`O0Ladder 12`: all six rungs `ok` — but it could NOT serve as a control,
and that is a finding about the record rather than about the code.** Its
September timings were never written down. The live-steps doc says "each must
read `ok`" and "send back every `O0|` line", and only the summary reached the
roadmap; `OPTIMIZE.0`'s entry carries no seconds for this step. A control
whose baseline is not in version control is not a control. **Today's numbers
are recorded below as the baseline for next time:**

| rows × cols | cells | seconds | µs per cell |
|---|---|---|---|
| 1,000 × 4 | 4,000 | 0.004 | (at the timer's floor) |
| 10,000 × 4 | 40,000 | 0.004 | (at the timer's floor) |
| 100,000 × 4 | 400,000 | 0.063 | 0.158 |
| 100,000 × 8 | 800,000 | 0.105 | 0.131 |
| 430,000 × 4 | 1,720,000 | 0.266 | 0.155 |
| 430,000 × 8 | 3,440,000 | 0.477 | 0.139 |

**And on its own terms it says something new and useful.** `DATALOG.12`'s
pass 4 measured handing cells back to Excel as **free** — a UDF returning
9,000 cells took 0.000s — and that has been quoted ever since. At 3.44
million cells it is **not** free: 0.477s, at a flat 0.13–0.16 µs a cell
across a 430-fold range. Cheap, linear, and no longer negligible. `OPTIMIZE.3`
should know it, because a ground program can produce an answer that large and
the spill is then a real fraction of the budget rather than a rounding error.

**Then the native forms — these set "grounding is under a second".**

| type | what it grounds | 2026-09-19 at 50×4 |
|---|---|---|
| `O0Ladder 1` | the pool, a Cartesian body | 4,200 rows, ~0.12s |
| `O0Ladder 4` | a native counter per shift | 84 rows, ~0.68s |
| `O0Ladder 5` | a native counter per person-week | 200 rows, ~0.65s |
| `O0Ladder 6` | never two in a row, via the `Next` pair relation | 4,150 rows, ~2.4s |

### The native forms RUN 2026-09-24

**Every ground-row count matched** — A1 105/210/420/420/840/1,050/1,680/4,200,
A4 21/21/84/21/84/21/84/84, A5 5/10/20/20/40/50/80/200, A6
100/200/415/400/830/1,000/1,660/4,150. So this is a pure re-timing, as
required: nothing was re-derived.

| shape | source rows | 2026-09-19 | 2026-09-24 | gain |
|---|---|---|---|---|
| A1, the pool (Cartesian body) | 218 | 0.12s | **0.074s** | 1.6× |
| A4, a counter per shift | 3,684 | 0.68s | **0.047s** | **14.5×** |
| A5, a counter per person-week | 3,800 | 0.65s | **0.047s** | **13.8×** |
| A6, never two in a row, via `Next` | 12,849 | 2.4s | **0.188s** | **12.8×** |

**This is `DATALOG.14`'s mechanism confirmed from a different harness, a
different shape family and a different fixture.** The filter fix removes one
`CreateObject` per SOURCE row a rule reads, at the 0.157 ms `DATALOG.12`
measured. Predict each saving as 0.157 ms × source rows and compare:

| shape | saving | 0.157 ms × source | ratio |
|---|---|---|---|
| A1 | 0.046s | 0.034s | 1.34 |
| A4 | 0.633s | 0.578s | 1.09 |
| A5 | 0.603s | 0.597s | 1.01 |
| A6 | 2.212s | 2.017s | 1.10 |

Three of the four within 10%, on savings spanning fiftyfold. **And it
explains the odd one out**: A1 gained only 1.6× because it READS 218 rows
while PRODUCING 12,902 — a Cartesian body's cost was always its output, and
the per-source-row constant was never what it was paying. The gain tracks
what a rule reads, not what it makes.

**What this means for `OPTIMIZE.3`, which is the whole point of the re-run.**
The `measured/model` column now reads **0.07–0.09 on the three source-heavy
shapes** and 0.42–0.45 on the pool. The paper model's own first coefficient
is `a = 0.16 ms per source row` — which is, to within noise, exactly the
constant `DATALOG.14` deleted. **So the model that sets `OPTIMIZE.3`'s
ceilings is calibrated on an engine that no longer exists, and is now about
tenfold pessimistic on any shape that reads more rows than it writes.** A
first re-fit off A4 and A6 gives roughly `a ≈ 0.010 ms` a source row and
`b ≈ 0.002 ms` a produced row; it reproduces A5 to 5% and does NOT reproduce
A1, so it is a starting point for `OPTIMIZE.3` to re-fit properly, not a
replacement model.

**Then the two that set the superlinear knee.**

| type | what it grounds | note |
|---|---|---|
| `O0Ladder 8` | at most 2 per shift, canonical triples | tops out at 95,760 rows at 20×4 |
| `O0Ladder 10` | at most k a week, (k+1)-subsets | rungs from 210 to 59,850 rows |

These two are where "linear under about 50,000 rows, superlinear past about
90,000" came from. **That knee is a `Scripting.Dictionary` property**
(`O0DictCost` measured 1.27 µs a key at 10,000 keys and 1.99 at 100,000),
and `DATALOG.14` did not touch dictionary growth — so the prediction is that
**the knee stays where it is and only the absolute times fall.** If the knee
MOVES, that is the interesting result and it changes what `OPTIMIZE.3` may
assume.

### RUN 2026-09-24 — and the prediction above was HALF WRONG

Every ground-row count matched again, and both guards fired on exactly the
rungs `O0Model` named. The knee half of the prediction held. **The "only the
absolute times fall" half did not, and this is the finding of the whole
re-run.**

**The ladder splits into two regimes that September's single model
conflated.**

*Regime one — shapes bound by the rows they READ* (A1, A4, A5, A6): **13–15×
faster**, `measured/model` 0.07–0.45. `DATALOG.14` deleted their dominant
term.

*Regime two — shapes bound by the rows they PRODUCE* (A8, A10: subsets,
triples, quadruples — **the exact shapes `OPTIMIZE.3` grounds**): **barely
moved.** Same rungs, same ground rows, September against today:

| rung | ground rows | source rows | Sept | today | gain |
|---|---|---|---|---|---|
| A8 20×4, k=2 | 95,760 | 10,080 | 10.9s | **9.039s** | 1.21× |
| A10 10×1, k=3 | 59,850 | 2,100 | 17.3s | **8.063s** | 2.15× |
| A10 1×1, k=5 | 54,264 | 441 | 17.6s | **18.719s** | **0.94× — slower** |

**Two of the three fit the mechanism exactly and the third does not.**
Predicting each saving as 0.157 ms × source rows, as the native forms
obeyed to within 10%: A8 predicted 1.58s against 1.86s actual (ratio 1.2),
and A10 k=5 predicted 0.07s against −1.1s actual — which is "no measurable
gain", correct, since that rung reads 441 rows and produces over a million.
**A10 k=3 is a 28× outlier against the same prediction**, and today's A10
series is internally smooth where it sits (3.461s → 8.063s → 18.719s as peak
rows go 139,650 → 279,300 → 427,329). The likeliest reading is that
September's 17.3s was the anomaly, not that this rung found a real 2.15×;
that cannot be re-tested, so it is named and not explained, and **the
conservative reading — combination forms gained about 1.2× and no more — is
the one to plan on**, because it is what two rungs of three say and what the
mechanism predicts.

**And above about 300,000 peak rows the paper model is now wrong in the
DANGEROUS direction.** `measured/model` rises monotonically with peak rows,
cleanly, in three separate series:

| series | measured/model against peak rows |
|---|---|
| A8, weeks = 1 | 1,050→0.40 · 3,087→0.41 · 9,450→0.51 · 26,754→0.65 · 79,800→0.77 |
| A8, weeks = 4 | 4,200→0.26 · 12,348→0.37 · 37,800→0.50 · 107,016→0.68 · **319,200→1.03** |
| A10 | 2,205→0.58 · 22,050→0.82 · 44,100→0.85 · 139,650→1.10 · 279,300→1.28 · **427,329→1.84** |

It crosses 1.0 near 300,000 peak rows and reaches **1.84**. So the knee is
still there — that much was predicted — but the consequence is the opposite
of the one this file expected. At the small sizes the model is ~10×
pessimistic and a ceiling drawn from it is too tight; **past ~300,000 peak
rows it is up to 1.84× OPTIMISTIC, and a ceiling drawn from it would be too
GENEROUS — it would admit a program that then takes twice as long as
promised.** For a ceiling whose job is to refuse before it searches, erring
generous is the failure mode that matters.

**Optional, for completeness:** `O0Ladder 2`, `3`, `7`, `9`, `11`. They add
coverage but decide no ceiling on their own.

**Send back:** every `O0|` line, as printed.

---

## What the numbers are for

**Both steps are DONE (2026-09-24). Hand the `OPTIMIZE.3` session these
four things.**

1. **The ceilings do not move uniformly, and for the shapes this item
   actually grounds they barely move at all.** Native forms got 13–15×
   cheaper; combination forms — subsets, triples, quadruples over a pool —
   got about 1.2×. A single "`DATALOG` is ten times faster now" adjustment
   applied across the board would be wrong, and wrong in the direction that
   admits programs it should refuse.
2. **The paper model must be re-fitted, and it is wrong in BOTH directions.**
   Its own first coefficient, `a = 0.16 ms per source row`, is to within
   noise the constant `DATALOG.14` deleted, so below ~100,000 peak rows the
   model is about tenfold pessimistic. Past ~300,000 peak rows it is up to
   1.84× **optimistic**. A first re-fit off A4 and A6 gives `a ≈ 0.010 ms`
   a source row and `b ≈ 0.002 ms` a produced row; it reproduces A5 to 5%,
   does not reproduce A1, and does not touch the superlinear term at all.
   **The superlinear term is the one that matters here and no one has fitted
   it.**
3. **The 2-second formula ceiling moved, but only for source-bound work.**
   It is a TIME ceiling, and grounding got ~10× cheaper on shapes that read
   more than they write — so far more of those now fit under it. It did not
   move for the combination forms.
4. **What `DATALOG.14` did NOT bank for it.** The per-predicate index was
   refused on arithmetic — its ceiling was 1.2–1.3× on *that ladder's*
   shapes, because the only atom with a bound argument filtered a delta the
   round had just built — and the magic-sets rewrite covers only the
   pass-through case. `OPTIMIZE.3`'s own join-order bullet ("most selective
   and already-bound first") is still its to build, and over a materialised
   pool it reuses, an index is a different proposition than it was there.
