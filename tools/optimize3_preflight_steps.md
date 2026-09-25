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

## Step 1 — the 100,000-row rung (about two minutes)

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
| `O0Ladder 12` | a bare UDF returning an array, no engine | every rung `ok`, times as 2026-09-19 |

**Then the native forms — these set "grounding is under a second".**

| type | what it grounds | 2026-09-19 at 50×4 |
|---|---|---|
| `O0Ladder 1` | the pool, a Cartesian body | 4,200 rows, ~0.12s |
| `O0Ladder 4` | a native counter per shift | 84 rows, ~0.68s |
| `O0Ladder 5` | a native counter per person-week | 200 rows, ~0.65s |
| `O0Ladder 6` | never two in a row, via the `Next` pair relation | 4,150 rows, ~2.4s |

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

**Optional, for completeness:** `O0Ladder 2`, `3`, `7`, `9`, `11`. They add
coverage but decide no ceiling on their own.

**Send back:** every `O0|` line, as printed.

---

## What the numbers are for

Hand the `OPTIMIZE.3` session three things:

1. **The re-timed ladder**, so its ground-size ceiling is set against the
   engine that exists rather than the one that did in September.
2. **A note that the 2-second formula ceiling is the one most likely to have
   moved.** It is a TIME ceiling over grounding that just got ~10× cheaper,
   so far more programs now fit under it than did when it was settled.
3. **What `DATALOG.14` did NOT bank for it.** The per-predicate index was
   refused on arithmetic — its ceiling was 1.2–1.3× on *that ladder's*
   shapes, because the only atom with a bound argument filtered a delta the
   round had just built — and the magic-sets rewrite covers only the
   pass-through case. `OPTIMIZE.3`'s own join-order bullet ("most selective
   and already-bound first") is still its to build, and over a materialised
   pool it reuses, an index is a different proposition than it was there.
