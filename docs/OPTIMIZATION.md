# Optimizing Spreadsheets

*Working notes toward a white paper on answer-set optimisation inside a
spreadsheet: what it costs, where it breaks, and what makes it tractable.
They record what we measured and what we predicted while building
Frazaro's fourth logic engine, `OPTIMIZE`. The other three are `DATALOG`,
`SQL` and `PROLOG`. The first two run as Excel worksheet functions today,
and `PROLOG` sits behind them as a query engine.*

*The premise. A great deal of the world's operational decision-making
already happens in Excel: rosters, seating plans, allocations, assignment
of reviewers to work. The constraints that decide those schedules usually
live in someone's head rather than in any system. Answer-set programming
(ASP) is the declarative technology built for exactly this class of
problem. You state the rules, and a solver finds the best arrangement that
breaks none of them. If a spreadsheet can host that honestly, at real sizes
and with answers anyone can audit, the most widely deployed end-user
programming environment becomes a runtime for it. These notes are the
evidence for whether it can.*

*House rules, the same as every other ledger in `docs/`:*

- *Append-only. An entry is never edited to look smarter than it was.*
- *Every figure is marked as a **prediction** or a **measurement**. A
  prediction is written before the measurement that tests it, and the gap
  is reported afterwards, whichever way it falls.*
- *Corrections are appended under the original, with their date.*

---

## Entry 1 — The shape of the search, on paper, before the engine exists *(2026-09-19)*

*Status: **predictions**, not measurements. They were written before any
optimisation code exists and before the harness that will test them has
run. Measured results will be appended below this entry, with their dates,
and the predictions will stay as written. The project's roadmap tracks
this work as `OPTIMIZE.0`.*

### Background, for a reader new to the project

`OPTIMIZE` will take a program of facts and rules, which are Excel Tables
and s-expression rules, plus three ingredients:

- a **choice**: "pick two eligible people for every shift";
- **constraints**: "nobody works two shifts in a row";
- an **objective**: "minimise overtime".

It returns the best schedule that breaks no rule, or it says in words that
none exists. Like every ASP system, it works in two phases:

- **Grounding** instantiates every rule over the actual data. This phase
  is polynomial, and the exponent depends on how a rule is written.
- **Search** decides which choices are true. This phase is exponential in
  the worst case, and its base depends on how much each decision forces.

The engine is required to be **deterministic**: the same inputs give the
same answer on every machine. So ties are broken by the Tables' own row
order, and search budgets are counted in units of work, never in seconds.

The running example is the **reference roster**: 50 people over 4 weeks,
with 3 shifts a day. That is 84 shifts and 2 to 4 people on each. It has
skills, vacations, "at most 5 shifts a week", "never two shifts in a row",
and an objective over overtime and changes from last month's schedule. The
**toy** is 5 people, 7 shifts, 2 per shift.

### 1. The paper model: how large the ground program is

A cardinality rule can be grounded in two ways.

- **As combinations**, the classical encoding. "At most 2 per shift"
  becomes one constraint for every group of 3 people who must not all be
  chosen.
- **As a native counter**: one constraint per shift, which counts its
  chosen atoms and propagates as a bound.

The table gives each shape the reference roster needs, both ways, at 10,
20 and 50 people over 1 and 4 weeks. Notation: P is people, W is weeks, and
S = 21W is shifts. Each cell is a count of ground constraints unless the
row says otherwise.

| shape | form | formula | 10×1w | 10×4w | 20×1w | 20×4w | 50×1w | 50×4w |
|---|---|---|---|---|---|---|---|---|
| choice atoms (the pool) | both | P·S | 210 | 840 | 420 | 1,680 | 1,050 | 4,200 |
| ″ with vacations pruned | both | 18·P·W | 180 | 720 | 360 | 1,440 | 900 | 3,600 |
| at most 2 per shift | combinations | S·C(P,3) | 2,520 | 10,080 | 23,940 | 95,760 | 411,600 | **1,646,400** |
| | native | S | 21 | 84 | 21 | 84 | 21 | **84** |
| at least 2 per shift | combinations | S·C(P,1) clauses of P−1 literals | 210 | 840 | 420 | 1,680 | 1,050 | 4,200 |
| | native | the same S counters | | | | | | |
| at most 5 a week | combinations | P·W·C(21,6) | 542,640 | 2,170,560 | 1,085,280 | 4,341,120 | 2,713,200 | **10,852,800** |
| | native | P·W | 10 | 40 | 20 | 80 | 50 | **200** |
| at most one shift a day | combinations | 7PW·C(3,2) | 210 | 840 | 420 | 1,680 | 1,050 | 4,200 |
| | native | 7PW | 70 | 280 | 140 | 560 | 350 | 1,400 |
| never two in a row | a pair relation (already binary) | P·(S−1) | 200 | 830 | 400 | 1,660 | 1,000 | **4,150** |
| | all pairs, filtered by arithmetic | P·S² join rows | 4,410 | 70,560 | 8,820 | 141,120 | 22,050 | 352,800 |
| a senior on every night shift | a clause or a counter | 7W | 7 | 28 | 7 | 28 | 7 | 28 |
| a shift that requires a skill | pruning | 0 constraints; the pool shrinks | | | | | | |
| vacations | pruned | 0 | | | | | | |
| | as unit constraints | 3PW | 30 | 120 | 60 | 240 | 150 | 600 |
| overtime above a contract of 4 | combinations ("5 or more of 21") | P·W·C(21,5) | 203,490 | 813,960 | 406,980 | 1,627,920 | 1,017,450 | 4,069,800 |
| | native | P·W counters + P·W terms | 10 | 40 | 20 | 80 | 50 | 200 |
| changes from the kept schedule | both | P·S terms | 210 | 840 | 420 | 1,680 | 1,050 | 4,200 |
| candidate worlds, exactly 2 per shift | | C(P,2)^S | 10^34.7 | 10^138.9 | 10^47.9 | 10^191.4 | 10^64.9 | 10^259.4 |

**What it says.** At the reference roster, the combination encodings come
to about **16.6 million** ground constraints. The native encodings come to
about **2,000 counters plus 4,150 pairs**. The difference is not a constant
factor. The combination count grows as C(P, k+1) in the bound k, while a
native counter's size does not depend on k at all. Cardinality written as
combinations is the single largest trap in the family. That is why
`OPTIMIZE` grounds every cardinality as a native counter and never expands
one.

**What grounding costs in the engine that exists today.** `OPTIMIZE` will
ground through `DATALOG`'s own evaluator, so the first question is what
`DATALOG` pays per row. Reading the code gives a two-cost model:

- **a ≈ 0.16 ms for every source row filtered.** The atom matcher creates
  one scripting Dictionary per row it tests. Creating one was measured at
  0.157 ms on the owner's machine.
- **b ≈ 0.01 ms for every row a join or a comparison produces** (band
  0.005 to 0.03). Join output rows are built as plain arrays, with no object
  creation. This figure is a prediction from component timings.
- **time ≈ a·F + b·J.** F is the number of source rows filtered and J the
  number of join and comparison rows, both counted across every round of
  the fixpoint. That includes the second round of semi-naive evaluation:
  each body position is re-run with the previous round's new rows, and the
  atoms ahead of that position are joined again before the empty delta is
  discovered.
- **Peak memory** is the largest single join output. Every group is joined
  at once, never one group at a time.

*A check against a real measurement:* the model gives 0.34 s for a
1,000-row scan that was measured at 0.383 s.

Predictions for the combination shapes:

| rung | exact ground rows | peak live rows | predicted time |
|---|---|---|---|
| at most 2 per shift, 10×1w | 2,520 | 9,450 | 0.4 s (0.3–0.8) |
| ″ 14×4w | 30,576 | 107,016 | 3.3 s (2.2–7.7) |
| ″ 20×4w | 95,760 | 319,200 | 7.8 s (4.7–20) |
| ″ **50×4w** | 1,646,400 | **5,145,000** | 95 s (49–276): **left to the model** |
| at most k a week, k=2, 20 people × 1w | 26,600 | 88,200 | 2.1 s (1.2–5.5) |
| ″ k=3, 10 × 1w | 59,850 | 279,300 | 5.7 s (3–16) |
| ″ **k=5, one person-week** | **54,264 = C(21,6)** | 427,329 | 9.6 s (4.9–29) |
| ″ k=5, 50×4w | 10,852,800 | about 85 million | about 32 min: **left to the model** |

At full size the native shapes are predicted to be cheap in today's
`DATALOG`: about 0.1 s for the pool, about 1 s for the counters (an
aggregate pays the per-row cost on every pool row), and about 2 s for the
pairs.

### 2. Where the first estimates were wrong

The plan this entry tests was drafted the day before. Checking its
arithmetic against the model turned up seven problems. They are recorded
here because each is a trap that a less careful implementation would ship.

1. **The reference roster was not a roster.**
   - Needs of 2 to 4 per shift add up to 63 slots a week, against a
     capacity of 250 (50 people × 5 shifts). That is a *tightness* (demand
     over capacity) of 0.25, or about 1.3 shifts per person-week: an on-call
     pool, not a roster.
   - A problem that loose measures nothing about search, because almost
     any assignment works.
   - The motivating sentence, "everyone works exactly 5 days", is
     **unsatisfiable** at that shape. It needs 250 slots a week, and 21
     shifts of at most 4 offer 84.
   - "5 days" and "5 shifts" are also different rules when a day has 3
     shifts.
   - Held at a realistic tightness of 0.84, the same roster needs 2, 4 and
     10 people per shift at 10, 20 and 50 people. "At most 10 per shift" as
     combinations is then C(50,11) = 3.7 × 10¹⁰ constraints per shift. At
     that point native counters are not a preference; nothing else exists.
2. **The first time estimates used the wrong unit.** Twenty minutes to
   ground "at most 2 per shift" and two hours for "at most 5 a week" came
   from pricing each *output* row at the cost of a *source* row. The model
   above says about 1.5 minutes and about 30 minutes. More importantly,
   **memory is the wall before time**: 5.1 million and about 85 million
   rows alive at once, which is gigabytes inside a spreadsheet process.
   Neither will be run. The measurement harness refuses any rung projected
   past half a million live rows.
3. **A refusal by shape would have refused the most basic roster.** The
   plan refused any rule whose body atoms share no variables, a "Cartesian
   body". But "anyone can work any shift" is `(shift S) (person P)`: a
   Cartesian product by nature, and the correct eligibility pool for most
   rosters. The refusal has to be by *projected size* against a ceiling.
   The shape is named in the message only when the size is over it.
4. **Canonical ordering saves more than claimed.** Imposing P1 < P2 on a
   symmetric pair halves it, as the plan said. But on a group of k+1 it
   divides by (k+1)!: by 6 for triples and by 720 for six-subsets.
5. **One question in the planned host probe cannot be asked.** Frazaro
   ships no custom ribbon, so there is nothing for a worksheet function to
   invalidate. The replacement question is whether a worksheet function can
   schedule a macro with `Application.OnTime`. That matters: one argument on
   record against a "keep searching?" dialog inside a formula is that "a
   worksheet function cannot write". A scheduled macro can. The design
   decision stands on three other reasons, and the record may need that
   sentence corrected.
6. **"A worksheet function has no DoEvents" is imprecise.** `DoEvents`
   compiles and runs inside one. What it *does* there, including whether
   Excel re-enters the function while it waits, is one of the probe's
   questions.
7. **The planned size ladder was too coarse for cubic shapes.** At 5, 10, 20
   and 50 people the guard would let only two points through, too few to
   test a growth curve. The ladder gains 7 and 14 people, and the weekly
   rule gets its own axis in the bound k.

### 3. The fixtures: every answer can be derived by hand

A measured ground size is only evidence if the expected size is known
independently. So every fixture is a closed form. There is no random
generator anywhere, seeded or otherwise.

- **People(Id, Name, Senior, Contract).** i = 1..P. Name is "P"&i. Senior
  when i mod 4 = 0. Contract 4 for odd i, 5 for even i.
- **Shifts(Id, Week, Day, Slot, Need, Skill).** s = 1..21W.
  Day = (s−1)\3 + 1. Slot is Early, Late or Night by (s−1) mod 3.
  Week = (Day−1)\7 + 1. Need is set by the reference roster's tightness.
  Night shifts require a senior.
- **Leave(Id, Day).** Person i is off on day d when (i + d) mod 7 = 0.
  That is exactly one day per person per week, so the vacation-pruned pool
  is exactly 18·P·W.
- **Next(From, To).** The pairs (s, s+1) for s = 1..S−1. "In a row" is a
  precomputed pair relation: the tractable shape, and the one the sentence
  grammar will be allowed to write.
- **Elig.** The pool, written directly by the generator. Each constraint
  measurement then pays only for its own rule body. (An earlier harness
  varied two things at once and produced a result that looked real and was
  an artefact.)
- **Kept.** Last month's schedule, by round-robin: shift s is staffed by
  persons ((2s−2) mod P) + 1 and ((2s−1) mod P) + 1. It never puts anyone
  on two shifts in a row once P ≥ 4, and never on more than 5 a week once
  P ≥ 9. Its clashes with Leave are counted exactly, which gives the first
  "these rows violate a rule" answer key.
- **The toy, solved by hand.** 5 people, 7 shifts, exactly 2 per shift,
  never two in a row. The first shift is any of C(5,2) = 10 pairs. Each
  later shift must take 2 of the 3 people who did not work the one before,
  which is 3 ways. So there are **10 × 3⁶ = 7,290 valid worlds** out of 10⁷
  candidates. An independent enumeration will confirm it.
- **Why a numeric Id.** `DATALOG` does not refuse `<` between names. It
  compares them as text, which puts "P10" before "P2" and reads a name like
  "1E3" as a number. A numeric Id makes the canonical order *equal* to the
  Tables' row order, which is also the engine's tie-break.

#### 3a. The reference fixture, in full, so it can be rebuilt from this page

*Added 2026-09-19, during `OPTIMIZE.2`. The list above was written before
the generator existed and left two things open — what `Need` is, and how
`Kept` staffs a shift that takes more than two people. Both are settled
below, from the generator as shipped. The generator itself is a diagnostic
module that is not kept in version control, so this section is the
authoritative statement of the fixture every later item is measured
against; `tools/optimize0_expected.ps1`, which is in version control,
re-derives every count from these same definitions.*

Two parameters: **P** people and **W** weeks. Everything else follows.

| | closed form |
|---|---|
| shifts | `S = 21·W` — 7 days × 3 slots a day |
| day of shift s | `(s−1)\3 + 1` |
| slot of shift s | `(s−1) mod 3` → 0 Early, 1 Late, 2 Night |
| week of shift s | `(day−1)\7 + 1` |
| senior | person i, when `i mod 4 = 0` |
| contract | 4 shifts for odd i, 5 for even i |
| on leave | person i on day d, when `(i + d) mod 7 = 0` — one day each a week |
| skill | Night shifts require `senior`; Early and Late take `any` |
| **Need, tight** | `max(1, P\5)` for every shift |
| **Need, loose** | `2 + slot` — Early 2, Late 3, Night 4 |
| Next(From, To) | `(s, s+1)` for s = 1..S−1 |

**Kept**, last month's roster, is a round-robin: shift *s* takes the next
`Need(s)` people in Id order after the shift before it, wrapping at P.
Person *j* of shift *s* (j = 0-based) is `((start(s) + j) mod P) + 1`,
where

- tight: `start(s) = Need·(s−1)`;
- loose: `start(s) = 9·((s−1)\3)`, plus 0, 2 or 5 for Early, Late, Night —
  9 being a whole day's demand under the loose needs.

Five tables are written, and their names carry the fixture's own size so
that several can live in one workbook: `PeopleO0f…`, `ShiftsO0f…`,
`LeaveO0f…`, `NextO0f…`, `KeptO0f…`, with columns `(Id, Name, Senior,
Contract)`, `(Id, Week, Day, Slot, Need, Skill)`, `(Person, Day)`,
`(From, To)` and `(Shift, Person, Week)`.

**Tightness** is the one ratio that matters, and it is weekly:
`weekly demand ÷ (P × 5)`, the most anyone may work in a week. A
tightness at or above 1 is impossible by counting alone, before any
search.

**The fixtures every later item is judged against.** Every figure here is
printed by `tools/optimize0_expected.ps1`, which derives them from the
definitions above rather than from the generator:

| fixture | need | weekly demand | capacity | tightness | seniors | a schedule exists |
|---|---|---|---|---|---|---|
| the reference roster, 50 × 4 | 10 a shift | 210 | 250 | **0.84** | 12 | not decided by hand |
| 10 × 1 loose | 2 / 3 / 4 | **63** | **50** | **1.26** | 2 | **NO, provably** — weekly demand over 5 per person |
| 5 × 1 tight | 1 a shift | 21 | 25 | 0.84 | 1 | **NO, provably** — too few seniors for 7 nights |
| the toy | 2 of 5, 7 shifts | — | — | — | — | yes: **7,290** worlds |

- The 10 × 1 loose fixture is impossible by one multiplication — 63 slots
  a week against a capacity of 50 — and is the case the reference solver
  could not prove in over three minutes under two encodings. It is
  `optimize-roster-loose` in the corpus.
- The 5 × 1 tight fixture is under its capacity and impossible anyway, for
  a reason counting the total misses: only **one** of its five people is
  senior (`4 mod 4 = 0`), and seven Nights a week each need one, so P4
  would have to work 7 shifts against a cap of 5. It is impossible
  *locally* as well — P4 is on leave on day 3 (`(4+3) mod 7 = 0`) and the
  Night of day 3 is shift 9 — and the solver found that local reason in
  0.003 s and the counting one never. It is `optimize-roster-senior`, and
  it is why the pre-checks test each counted resource separately rather
  than only the total.

**What the kept roster itself breaks**, at the reference size, counted
from its own rows — the first "these rows violate a rule" answer key this
project had, and the shape `OPTIMIZE.2`'s violations table now produces:

| | 50 × 4 |
|---|---|
| assignments on a leave day | **118** |
| two shifts in a row | 0 |
| person-weeks over five | 0 |
| Nights with no senior | 0 |
| person-weeks over contract | **20** |

The 118 leave clashes each cost at least a drop and an add, which is
where the hand floor of **236 changes** comes from. At 10 × 1 the same
count is 7 and the floor is 14, and the solver's proven optimum of 24
sits above it: a floor, not a prediction.

### 4. The measurement plan

A standalone Excel harness grows each fixture up a size ladder and times
Excel's own recalculation of a `=DATALOG(...)` formula that grounds one
rule body. Each step builds its own fresh sheet and its own Tables. Twelve
steps:

| steps | what is measured | sizes |
|---|---|---|
| 1–7 | the pool (Cartesian); the pool minus vacations; the pool through a skill join; a counter per shift; a counter per person-week; "in a row" through the pair relation; "in a row" by arithmetic | the full ladder, up to 50 × 4 weeks |
| 8–9 | "at most 2 per shift" as triples, with and without canonical order | as far as the guards allow |
| 10 | "at most k a week" as (k+1)-subsets, for k = 1, 2, 3, then 5 | k = 5 on a single person-week |
| 11 | overtime per person-week, against the kept schedule | full |
| 12 | a control: a bare worksheet function returning an array of the same shape | beside every rung |

The control exists because returning 9,000 cells was measured as free, and
nobody has measured 380,000.

**Guards:**
- Stop after any size that takes over 10 s.
- Skip any size projected past 15 s. The projection uses the cost model,
  with a and b re-fitted after every rung, rather than a guessed growth
  exponent. Steps 1–7 pin a, and steps 8–10 pin b.
- Skip any size predicted to hold more than 500,000 rows at once (about
  150 MB).
- Check every answer: the total, every group's tally (C(g,3) for a shift of
  g people, for example), and that every tuple is in canonical order. A
  wrong answer stops the ladder and is never timed.

An independent script re-derives every expected count by enumeration,
separately from the harness's closed forms.

**Left to the model, and why:**
- The 50-person rungs of the combination shapes. Their peaks, 1.3 to 5.1
  million live rows, are the memory wall from finding 2.
- "At most 5 a week" beyond one person-week. Groups are independent, so
  the reference figure is exactly the measured single group × 200. The
  k = 1 and k = 2 rungs test that linearity across the number of people.

*Measurements will be appended here.*

**Note, 2026-09-19, when the harness was built.** The harness prints its own
model beside every measurement. That model counts two things the table above
left out: reading each Table argument, and handing the answer's rows back. So
its predictions run 5–15% above the ones in this entry. For example, "at most 2
per shift" at 20 × 4 weeks is 8.7 s rather than 7.8 s, and one person-week of
"at most 5 a week" is 10.2 s rather than 9.6 s. The entry's figures stand as
written; the gap will be reported against both.
Three more facts were settled while building:
- the memory guard will never run exactly nine rungs;
- an independent enumeration of every rung agrees with the closed forms;
- the toy's 7,290 was confirmed by exhaustive search.

Two fixtures turned out to have no schedule, and that can be proved by hand:
- **5 people** have one senior. Seven night shifts each need a senior, and one
  person may work at most five a week.
- **10 people with the loose needs** must fill 63 slots a week, against a
  capacity of 50.

Those two are the first answer key for "no schedule exists".

### Measured, 2026-09-19

*Excel 16.0 64-bit on one Windows machine, the project's owner at the
keyboard. 101 ladder rungs ran, every answer was right on every rung, and
nothing crashed.*

**The central claim held.** Written as native counters, every rule of the
reference roster grounds in under a second in today's evaluator, and the
pair relation for "never two in a row" in 2.4 s. Written as combinations,
the same rules stop at toy sizes:

| rule | as combinations | as native counters | ratio |
|---|---|---|---|
| at most 2 per shift, 20 people × 4 weeks | 95,760 triples in 10.9 s | 84 counters in 0.28 s | about 39× |
| at most 3 a week, 10 people × 1 week | 59,850 quadruples in 17.3 s | 10 counters in 0.15 s | about 118× |
| at most 5 a week, one person-week | 54,264 six-subsets in 17.6 s | one counter in about 0.0035 s | about 5,000× |

At the reference roster the combination forms would be 1.6 million and 10.9
million rows. Neither was run: the harness refuses any rung predicted past
half a million live rows. Canonical ordering divided the rows by exactly 6
for triples, and the time by 2.3 to 4.0.

**The cost model: the cause held, some of the numbers did not.**
- The per-source-row cost was predicted at 0.16 ms and **measured at about
  0.17 ms**.
- The per-produced-row cost was predicted at 0.01 ms (band 0.005 to 0.03).
  It **measured about 0.0055 ms** for plain join rows, and **about 0.014 ms**
  for rows that pass through arithmetic or a comparison, which are keyed
  into a hash table.
- There is also a fixed cost of about 12 ms per recalculation, which the
  model left out.
- Inside the linear range, the model predicted 0.72 to 1.18 times the
  measured time.
- **Two of our refined predictions failed.** One was 2.6 times too
  optimistic: it priced keyed rows at the cost of plain join rows. The
  other was twice too pessimistic: it extrapolated a trend that turned out
  to depend on what had run before. Both are recorded as failures. The
  first was withdrawn before the steps it covered were run.

**What bends the curve is the host, not the algorithm.** A standalone test
of the scripting Dictionary, with no engine involved, measured the cost of
adding a key rising from 1.27 µs at 10,000 keys to 8.83 µs at 400,000. Every
rung whose largest intermediate relation passed about 90,000 rows ran 1.25
to 2.75 times slower than its linear model. That is measured support for a
design rule we had only argued for: once grounded, the search holds atoms as
integers in typed arrays, never as strings in a hash table.

**The host probe: what a worksheet formula may do while it runs.** Ten
questions, each asked directly:

- **Esc during a long formula** brings up the VBA debugger's "Code execution
  has been interrupted" dialog. Ending it leaves `#VALUE!` in the cell.
- **A formula cannot defend itself against Esc.** The setting that would
  turn Esc into a catchable error (`EnableCancelKey`) is *silently ignored*
  when a formula sets it: no error is raised, and it reads back unchanged.
  Set from a macro, it works exactly as documented. It reverts when the
  macro ends, so it cannot be armed in advance for a later calculation.
- **A dialog box does appear from inside a formula**, and its answer
  reaches the cell. But a full recalculation asks again. **Excel's Function
  Wizard evaluates the formula twice for every change** to an argument,
  including a paste with no keystrokes. And after the VBA project was reset,
  one unrelated sheet insertion re-ran every such formula in the workbook,
  asking all over again.
- **The status bar, and scheduling a macro with `OnTime`**, are both
  silently dropped when a formula asks for them. The same scheduling call
  from outside a formula works. So a formula cannot write to the workbook,
  even one step removed. **This corrects finding 5 of this entry**, which
  suspected that it could.
- **`DoEvents` inside a formula** keeps Excel responsive: the user can type
  while the formula runs, and Excel does not re-enter the formula. But an
  edit to the formula's own input, made during that time, left the cell
  **silently stale**. It was never recalculated, and F9 did not fix it.

Four requests were silently accepted and then ignored, dropped, or left the
cell wrong: the Esc setting, the status bar, the scheduled macro, and the
edit made during `DoEvents`. None raised an error. The read-backs either
reported success or showed the old value. The practical rule for anyone
building long computations into spreadsheet formulas: **code inside a
worksheet function cannot change anything outside the value it returns, and
cannot trust the host's answer to "did that work?"**

**The design decision it settled.**
- A formula performs a search only when the ground program is small enough
  to project under two seconds. Past that, it refuses by name and points to
  a command.
- Long searches, stopping cleanly on Esc, and the offer to keep searching
  belong to the command, where the host cooperates.
- The ceilings themselves are left to the first search item, which will
  measure them on its own integer-based grounder rather than inherit them
  from today's evaluator.

**The reference solver, as an oracle.** clingo 5.8.2, a mature answer-set
solver with conflict-driven learning, was run on the same fixtures as a
check that sits entirely outside our code.

- **The toy's 7,290 valid worlds** now agree three ways: by hand (10 × 3⁶),
  by exhaustive search in PowerShell, and by clingo's model count.
- **Finding a good schedule is easy; proving it's the best is where the
  time goes, even at 10 people.**
  - At 10 people and one week, clingo found the best schedule
    (overtime 0, 24 changes) at once, and spent 19.85 of its 21.7 seconds
    proving nothing better exists.
  - At 20 people it had no proof after five minutes.
  - At the reference roster it found a first schedule in 0.06 s, and after
    ten minutes still had no proof.
  - So for an engine inside a spreadsheet, "the best we found within the
    budget, not proven best" is not a failure mode. It is the ordinary
    answer at real sizes, and it has to be said in those words.
- **Some impossibilities are arithmetic, not search.** A roster whose shifts
  need 63 people a week, from 10 people who may work 5 each, is impossible by
  one multiplication. clingo could not prove it in over three minutes,
  under two encodings. This is the pigeonhole principle, whose proofs are
  exponential for the clause learning such solvers use. So an optimising
  engine should run counting pre-checks, demand against capacity for each
  counted resource, before it searches. It can then say why no schedule
  exists, in one sentence with the numbers.
- **Value order has to follow the objective's order.** At the reference
  roster, the objective is overtime first, then changes from last month's
  roster.
  - Hinting clingo to try last month's choices first reached the floor of
    236 changes almost at once. That floor was derived by hand, and the run
    proved it can be reached.
  - But it stalled at 35 overtime, where the plain search had reached 0.
  - The better schedule, by the stated order, is the plain search's (0,
    1,486). The true optimum lies at overtime 0 with between 236 and 1,486
    changes, and it is not yet known.
  - The rule we take from it: prefer what the first term of the objective
    prefers, break ties by the next, and only then by what was kept. We had
    argued that ordering; this is the first measurement of what happens
    without it.

---

## Entry 2 — Building the base case: what the corpus refused, and six defects a green suite could not see *(2026-09-19)*

*Status: written after the item shipped and was tested live, so unlike Entry
1 nothing here is a prediction. Entry 1 measured the shape of the search
before any engine existed. This entry records what building the first piece
of it actually cost. The project's roadmap tracks this work as `OPTIMIZE.1`.*

### Background, for a reader new to the project

`OPTIMIZE.1` is deliberately the least ambitious item in its family. A
program with no choice, no constraint and no objective *is* a Datalog
program, so `=OPTIMIZE(...)` must answer exactly what `=DATALOG(...)`
answers and nothing more: the same s-expression forms, handed to the same
evaluator, spilling the same table with the same header row.

What it settles is not behaviour but **vocabulary**. The words a user will
type for a choice, a constraint and an objective are fixed here, parsed
here, and refused by name here — before anything executes them — so that
nothing written today has to be rewritten when searching arrives. A spelling
is the one part of a program that cannot be changed later without breaking
somebody's workbook.

Two thirds of the work was therefore not code. What follows is the part that
was hard, and it was mostly not the part that looked hard.

### 1. The corpus refused the plan's own example

Before this item, the plan carried an illustrative choice form:
`(choose 2 (assign Shift Person) (eligible Person Shift))` — a count, the
rows chosen, the pool they come from. Three slots, and it reads well.

A corpus of twenty-one hand-solved questions had been written first, in
plain English, specifically so that spellings could be judged against
sentences rather than against taste. Checking the three-slot form against
all twenty-one found that it **cannot write the family's own killer case**.

| the sentence | what it needs | what three slots gave |
|---|---|---|
| "every shift gets exactly the people it needs" | the count comes from a *column of a table*, so it must be a variable | nowhere to bind one — unwritable |
| "every month, every person gets exactly one duty" | a group of *two* things at once | one implied group |
| "each extra is in the quote or not" | per-*row*, no group and no count | "at least zero" would mean the opposite of what it says |
| "every reviewer is senior" | a universally quantified implication | a direction nobody had fixed |
| "the least total weight of wishes broken" | the weight is a *table column* | a literal cost only |

Three shapes were added: a trailing optional `(per group-atoms...)` modifier
(which also holds the plural group), a fifth choice form `choose-any` for
the per-row case, and a cost term that accepts a variable bound by the body.
One shape was fixed rather than added: a requirement takes the same
head-then-body shape an ordinary rule does, which makes it an implication —
and that turns out to remove the need for an existential form altogether,
since "every night has a senior on it" is an ordinary derived rule doing the
existential, with a requirement over its head.

**No settled *word* changed.** Every word chosen the day before survived
contact with all twenty-one sentences. What had not been decided, and what
nobody had noticed was undecided, were the *shapes*. That is the return on
writing the corpus before the engine: the gap it found was invisible from
inside the design, and would otherwise have been found by the first user who
tried to write the one problem the engine exists for.

### 2. A file's text encoder is the wrong encoder for a key

The engine is allowed to remember an answer so that asking the same question
twice costs nothing, and it recognises a repeat question by a fingerprint of
the rules and of every cell of every table. The fingerprint must be
**injective**: two different inputs may never share one, or the memory hands
one workbook's answer to another workbook's question. That is a wrong
answer, not a slow one, so this precursor was built and pinned before any of
the engine's surface existed.

Two things already in the codebase were nearly right, and both were wrong in
instructive ways.

- The only text hash available converted text through the system's legacy
  code page, which is exact for ASCII and lossy for everything else. `Zoë`
  and `Zoe?` could land on one key.
- The UTF-8 encoder used for writing files encodes an unpaired surrogate as
  U+FFFD, the replacement character. For a **file** that is correct and
  careful: no reader accepts a raw surrogate sequence, so writing one would
  produce a file nothing can read. For a **key** it is a collision
  generator, because `"a" + U+D800` and `"a" + U+FFFD` become the same
  bytes.

The same fact — "this input cannot be represented, so substitute" — is right
in one context and a defect in the other. The fix is a sibling encoder
differing by exactly one absent branch, which keeps an unpaired surrogate as
its own three bytes; it is byte-identical to UTF-8 for every well-formed
string.

Encoding alone is not enough, because concatenation loses boundaries. Each
field is written as a type tag, a four-byte length, then its bytes, which is
what distinguishes:

| these | from these | by |
|---|---|---|
| the number `1` | the text `"1"` | the tag |
| a blank cell | an empty one | the tag |
| `TRUE` | the text `"TRUE"`, and the number `1` | the tag |
| `("ab", "c")` | `("a", "bc")` | the lengths |

An existing helper joined values with a control character for a different
purpose in the same codebase; it is not injective (a cell containing that
character forges a boundary) and was deliberately not reused. Numbers are
written as their exact eight bytes rather than as text, so no locale,
rounding or cell format enters — obtained without adding a second platform
call to a codebase that holds itself to one, by copying between two
same-sized user-defined types.

### 3. Making "the same engine" a measured claim rather than a description

The item's entire claim is that this is the existing engine under a new
name. A claim like that is easy to assert and easy to quietly break.

The proof is the existing engine's **whole test corpus**: all 178 distinct
Datalog programs the test module names, each run through both engines and
required to agree completely — the same relation, the same rows *in the same
order* (ties are broken by table order, so a reordering is a real
difference), the same header names, the same boolean answer — or to refuse
identically, with the same error number and the same words.

Two decisions made that affordable:

- **No fixtures.** A program whose tables its original test supplied refuses
  in *both* engines, identically, because its predicates are then undefined.
  That is still parity, so no program needs a fixture of its own. 128 of the
  178 answer outright; the rest agree on their refusal, to the word.
- **The comparison reuses the key from section 2.** Hashing each relation
  with the injective encoding makes one string comparison cover arity, row
  count, order and every cell's type at once, and it cannot raise on a cell
  the comparison did not anticipate. The precursor built for the memory did
  double duty as the instrument that measures the main claim.

A table of 178 programs is exactly the shape that silently stops being
complete: add a Datalog test a year from now and the parity suite still
passes, having quietly become "every program as of this item". So a static
check reads every call site in the test module and fails when one of their
programs is missing from the table. It earned itself the same day, refusing
a change until a newly added program joined the table.

### 4. Where a refusal belongs when two engines share one

If two engines share an evaluator, whose words does a refusal use? The
answer adopted is a split, and it needs stating because neither half is
obviously right:

- A refusal the **shared evaluator** raises keeps the original engine's
  wording, because that engine raised it. Only the cell's prefix names the
  function the user actually called.
- A refusal about something **only the new engine has** — its own forms, its
  own table arguments — is worded by the new engine, because a message
  reading "every DATALOG table argument…" inside a cell the user wrote
  `OPTIMIZE` in names the wrong function.

One consequence was not designed and is worth recording: because every one
of the new forms is refused *before* the shared evaluator is called, the
shared module needed **no change at all**. The new engine is purely
additive — which is the strongest available evidence that the base case
really is the old engine, since the old engine did not have to learn
anything.

### 5. Six defects, and what caught each

The item shipped with three failures found by the live pass and three found
before it. None was in what the engine answers. The table is the point of
this section: note the third column.

| what was wrong | what caught it | what a green suite showed |
|---|---|---|
| An assertion combining a shape guard with a conversion of the same value — it would raise, rather than report, on exactly the run where it failed | a static scan already in the repo, on the item's own first draft | nothing; the assertion passed |
| Two test expectations matching a message fragment that had been reworded after they were written | the live suite | the failure, correctly |
| A message template gained a slot; two of its **four** raise sites were updated | the live suite, via one test that happens to reach a third site | nothing for the fourth site |
| 180 generated lines inserted *above* a module's header line, by a script whose failed match did not stop it | a count printed for an unrelated reason reading double | **all 24 checks passed** |
| A hand-written source scan whose parameter regex stopped at the first close paren, so array parameters read as undeclared | running it against modules known to compile | its own "clean" run, on modules that had no array parameter |
| A new check reading only the last string literal of a template built by concatenation | its own first run over a known-good tree, reporting 14 problems | — |

Three lessons generalise past this project.

1. **A defect can be invisible to every mechanical check and to a green
   suite at once.** The module-header case is the clean example: the check
   that counts the very lines which had been duplicated read the right
   number, because it reads them from inside the function it expects them in
   and never looks at the top of the file. The bill for that one would have
   come due as an import failure on someone else's machine.
2. **A new check whose first run over a known-good tree is not clean is
   reporting its own defect.** Both harness bugs above were found that way,
   and one of them had already passed a control — against files that
   happened not to exercise the construct being parsed. Choose controls that
   exercise the construct, not merely files known to be good.
3. **A message reworded after its test was written needs that test reread**,
   and a fragment short enough to drift is worth lengthening. Both stale
   expectations were caused by improving a message.

### 6. And one defect that was not in the software at all

Two steps of the live pass failed with a correct refusal: a rule named a
relation nothing defined. The instruction said "set the table name to
`Chain`" without saying *where*, and a spreadsheet has two boxes that both
take a name and both commit on Enter — the table's own name box on the table
tools tab, and the name box above the columns, which creates an ordinary
defined name and leaves the table called `Table1`. Only the first is the
table's name, and it is the table's name a rule must use.

The second outcome is nearly indistinguishable from success: the name
resolves, so no reference error appears, and the formula bar colours it like
any other name. The refusal was right both times and told the reader nothing
they could act on, so the natural conclusion was that the spreadsheet was
misbehaving rather than the instructions.

Two things changed. The instructions now have the name set and read back in
one line, so neither box is involved. And the refusal — in all four engines
that share it — now ends by listing the names the program *does* define,
table arguments first, because those are what this failure is nearly always
about:

> `(query staff5)` names a predicate with no facts, no rule, and no matching
> table — check the spelling, or that the table argument's name matches.
> **The names this program does define, table arguments first, are: table1.**

The general form is worth keeping: **a fixture slip on a named input is
indistinguishable, in the refusal it produces, from a defect in the thing
being tested.** A refusal that names what it *did* find, and not only what it
wanted, collapses that ambiguity for nothing.

### What it cost, in numbers

| | before | after |
|---|---|---|
| automated assertions in the query suite | 1,443 | 1,768 |
| static scans | 22 | 25 |
| refusal catalogue entries / raise sites checked | — | 552 / 737 |
| Datalog programs re-run through the new engine | — | 178 |
| lines changed in the shared evaluator | — | 0 |

Of the 325 new assertions, 40 pin the key (against published digests
recomputed independently from the bytes the design names, so a shared bug in
the encoder and in its own test cannot cancel out), 78 pin the spellings —
including eighteen of the corpus's own English rule lines, written out in
the settled spellings and required to parse — and 179 are the parity proof.
The remaining 28 cover the memory, including the case that matters most:
wiping it between two questions must change no answer, only the time.

---

## Entry 3 — Constraints without a search: a seam not taken, and a crash nine items old *(2026-09-19)*

*Status: written after the item shipped and was tested live, so like Entry 2
nothing here is a prediction. Entry 1 measured the shape of a search before
any engine existed; Entry 2 recorded what building the base case cost. This
entry records the first item that makes the engine say something Datalog
cannot: that a schedule breaks a rule, and which rows break it. The
project's roadmap tracks this work as `OPTIMIZE.2`.*

### Background, for a reader new to the project

The engine answers questions written as s-expressions over spreadsheet
tables. Until this item it only ever **derived**: given facts and rules, it
computed everything that follows and spilled the result into the worksheet.
This item adds the first thing that **eliminates**. A constraint says "no
world in which this holds" (`forbid`) or "wherever this holds, that must
too" (`require`), and with no choice in the program there is exactly one
world, so a constraint is not a reason to try another — there is no other.
It is a check.

That makes this the smallest possible proof that a constraint is a different
kind of thing from a rule, and it is also the whole of the compliance case
the engine was scoped around: *does this month's assignment break any rule,
and which rows?* No search is involved, and none is needed.

### 1. The seam we did not build

The plan said a constraint's body should be instantiated over the world "by
the same join the evaluator runs for a rule body". The join is private, and
deliberately so: it is entangled with the evaluator's own atom
representation, its column bookkeeping, its aggregates, its built-ins and
its safety checks. Three ways to reach it were on the table, and all three
were bad.

**Exposing the join** would have promised that whole internal contract to
every future caller, in exchange for one capability.

**Moving the join into the shared substrate layer** looked like the
principled choice and turned out to be *structurally unavailable*, for a
reason worth recording. That layer is forbidden from raising user-facing
refusals — refusal wording belongs to each engine, so that a message inside
an `OPTIMIZE` cell never names a different function. The substrate's own
table-resolution helper returns a *reason code* rather than raising, for
exactly this reason. But the join raises about a dozen refusals directly. So
moving it means converting every one of those to a reason code and rewiring
four engines to re-word them. A layering rule adopted for a small reason —
who owns the words of an error — silently decided, years later, which
refactor was affordable.

**Duplicating the join** would have meant re-implementing column-name
desugaring, five kinds of body item and the safety checker, all of which
drift the moment the shared evaluator changes.

The fourth option is the one we took, and it is the finding: **do not borrow
the machinery — become the thing the machinery already processes.** A
constraint is rewritten into an ordinary rule whose head collects the rows
that break it.

```
  (forbid B1 B2)        ->  (rule (check-1 "1" V...) B1 B2)
  (require CONS B...)   ->  (rule (check-2 "2" V...) B... (not CONS))
  (require GROUND)      ->  (rule (check-3 "3") (not GROUND))
```

A violation is then exactly a non-empty derived relation. Negation,
aggregates, comparisons, column-name atoms, stratification and the safety
checker all arrive for free and cannot drift, because they are not being
reused — they are simply running, on a rule, as they always do.

The seam this needed is one function, and it is at the **program** level
rather than the algorithm level: an entry point that takes a program already
parsed into forms, rather than one that takes text. It promises only "a
program is a list of forms, and here is the entry that takes them". The
grounding machinery stayed private, and the shared evaluator changed in
three places totalling about forty lines.

The generalisable form: **the narrowest seam is at the layer where the data
already has a name**, not at the layer where the algorithm you want lives.
We wanted the join; what we actually needed was the ability to hand over a
program, which is a far smaller promise and one the module could already
almost make.

A variant we rejected is worth naming because it was tempting. The same
rewriting can be done with *no* change to the shared evaluator at all, by
re-serialising the rewritten program back into text and handing over the
text. We refused it: that puts a text writer between what a user wrote and
what the engine answers, where a writer bug becomes a **wrong answer**
rather than a refusal. A round trip through a serialiser is not free just
because it touches no files.

### 2. Where a mistake lands should decide where a feature goes

Entry 1 found that some impossibilities are arithmetic, not search: a roster
needing 63 person-shifts a week from ten people who may work five each is
impossible by one multiplication, and the reference solver could not prove
it in over three minutes under two encodings. So the engine should run
counting pre-checks — demand against capacity — before searching.

This item was supposed to land them, and the scoping run found they could
not run at all. **Every counting case takes its demand from a choice**, and
a choice is refused until the search item exists. Reaching them here would
mean inferring demand from the choice forms and the rule graph *without* the
grounder that the search item builds: the counted resource is often a
derived relation fed by two different choices, and a count can come from a
table column rather than a literal. That inference is a second, throwaway
implementation of the very machinery about to arrive — the same duplication
we had just rejected in section 1, relocated.

The argument that settled it, though, was not duplication. It was **where a
mistake lands**. Wired up, a slip in that arithmetic answers *"no schedule
satisfies every rule"* for a roster that has one — an engine named for
optimisation telling a manager their problem is impossible when it is not.
Left unreachable, the identical slip is a failing test.

So the three comparisons were built, and pinned against every counting case
in the corpus, and nothing calls them. When the search item lands it will
supply demand and capacity from its real grounding and call the same
functions unchanged. This is the same pattern the previous item used for its
five result sentences: **reserve and prove the part that is settled, and let
reachability arrive with the machinery that makes it honest.**

It bought something concrete beyond safety. Not one existing refusal message
had to be reworded, so not one existing test expectation had to be re-read —
which is precisely what cost the previous item's live pass two of its
fourteen steps.

### 3. A documented hazard is not a fixed hazard

The suite crashed on its first live run, inside the shared evaluator, with
"subscript out of range".

The cause: a rule whose body binds **nothing** — every argument a constant —
projects down to the empty tuple, and that path built it with a zero-length
`ReDim` of a one-based array. That idiom is unreliable in this host, and it
is *documented as unreliable in the very file whose own helper works around
it*, with a note recording that it was caught live once before.

It has been there since the evaluator's first version, through nine
subsequent items and some 1,800 automated assertions. Nothing had written
such a rule. Rules are written to derive things *about* data, so they
mention variables; a rule with no variables at all derives one fixed fact
and looks pointless.

What made one appear was a sentence. The corpus of manager-sized questions
contains a quoting rule — *"carbon wheels need the carbon frame"* — which
names two catalogue items and no variables whatsoever. Written as a
constraint, it is exactly the shape nobody had a reason to write.

Two findings, and the second is the useful one:

- **Coverage is a function of the shapes your corpus writes, not of the
  lines your tests touch.** Line coverage over that projector was total for
  nine items. The uncovered thing was an *input shape*, and no coverage tool
  measures those.
- **A hazard that is written down is not thereby handled.** The note
  describing this exact failure, and the helper that avoids it, sat in the
  same module. Documentation of a trap protects only the code that was
  written after someone read it.

Both shapes of the offending rule are now permanent cases in the
cross-engine comparison suite.

### 4. A test harness more adversarial than the product

The second defect was this item's own, and it shows a class of bug that
worksheet use cannot produce.

The rewritten check rules derive relations, and derived relations are
written into the dictionary of relations the caller supplies. Each worksheet
function builds a fresh dictionary per call, so nothing accumulates. **Every
programmatic test, by contrast, builds one dictionary and reuses it**, which
is the efficient and obvious thing to do — and which handed the second
program the first program's violations, still sitting under the same
generated name. A clean program then reported a violation from a rule it did
not contain.

That is the same wrong answer section 2 declined to risk, arriving through a
door nobody was watching: not from arithmetic, but from a side effect
crossing a call boundary.

The fix is that each run now works in its own copy of the dictionary, and
the cache key is taken over that copy — otherwise a cell's key would depend
on what ran before it, which is the determinism guarantee broken by the side
door. **The general shape: a function that writes into a structure its
caller owns has an unstated precondition — that the caller does not reuse it
— and the only callers that violate it may be your tests.**

### 5. An umbrella prefix needs sub-prefixes from the first generator, not the second

Generated identifiers in this project live under a reserved `vla-` prefix,
refused when a user writes one. The first draft of this item reserved that
whole prefix.

It was already occupied. A different subsystem — the one that compiles
English sentences into queries — emits `vla-ask-...` and `vla-not-...`
predicates into ordinary programs, and **two of them sit in the cross-engine
comparison table**. Reserving the umbrella would have refused a program the
other engine answers, and would eventually have made this engine refuse its
own sibling's output.

Narrowed to `vla-check-`, one sub-prefix per generator. The finding is a
small piece of API design that generalises: **a reserved namespace needs its
per-generator partition defined by the first generator, not negotiated by
the second**, because by then the first one's names are already in tests,
fixtures and other people's files.

A second, smaller hole in the same area: the reserved-name guard walked the
program's nested forms looking at each list's head, which is where a
predicate name appears — except in a bare `(query name)`, where the name is
a plain token and not a list head at all. `(query vla-check-1)` would
therefore have spilled the engine's own internal bookkeeping into a
worksheet. **An enumeration of "all the places X can appear" is worth
writing out explicitly; the one that gets missed is always the one with a
different shape.**

### 6. The live pass: three instruction defects to two engine defects

Fourteen numbered steps were run by hand in a live workbook. They found two
real defects (sections 3 and 4) and **three defects in the steps
themselves**, none of which was a defect in what the engine answers. That
ratio is worth recording, because the instructions are usually treated as
the trustworthy part of a verification pass.

All three were claims about the *host* rather than the engine:

- A step asked for `=COUNTIFS(Broken[[#All]],"<>")` over a named spilled
  range. That bracket syntax is a *structured reference*, which works only
  on a real spreadsheet Table; a named spill is referred to by its bare
  name. Excel rejected the formula **at entry**, with a dialog, rather than
  evaluating it to the `#NAME?` error the step had predicted as its fallback
  — and the step offering a fallback at all was the tell that neither half
  had been reasoned through.
- A step proved "three cells, one search" by clearing a counter, forcing a
  **full recalculation**, and expecting 1. It read 4. A full recalculation
  is workbook-wide, and by that step the workbook held nine earlier sheets
  of cells calling the same engine; with the counter just cleared, every
  distinct program in the workbook was a miss. The feature was working
  perfectly.
- The third was inherited from the previous item: a step that depends on a
  name the operator must set needs that name read back before the first
  formula, because a fixture slip on a named input is indistinguishable from
  an engine bug in the refusal it produces.

Two rules came out of it, both narrow:

- **Make a measurement's scope as narrow as the claim.** Recalculate one
  cell, not the workbook.
- **Ship a discriminator with every counter.** The 4-versus-1 question was
  settled in a single line by also printing the number of distinct cached
  keys: *four runs, four keys* means every run was a different question and
  no repeat ever missed — which is the very property the step was trying to
  demonstrate. Without that second number, a surprising count is an
  argument; with it, it is a diagnosis.

### 7. A scope finding for the next item

One measured result changes a future item's plan. Counting constraints —
"nobody holds more than two roles this month" — were expected to need the
aggregate work scheduled as a separate item. They do not, for *checking*: an
aggregate inside a constraint over a fixed world is an aggregate inside a
rule body, which the evaluator has done for several versions. It works
today, and is pinned doing so on the corpus's own segregation-of-duties
case, where the cap correctly does **not** fire at two and fires on six rows
at one.

What the aggregate item actually owns is narrower and harder than it looked:
propagating a counter over atoms that are still *being chosen*, during a
search, as a bound rather than an expansion. Splitting "check it" from
"propagate it" was not visible on paper and became obvious the moment a
constraint was a rule body.

### What it cost, in numbers

| | before | after |
|---|---|---|
| automated assertions in the query suite | 1,768 | 1,860 |
| static scans | 25 | 26 |
| Datalog programs re-run through the new engine | 178 | 180 |
| lines changed in the shared evaluator | 0 | ~40 |
| worksheet functions | 3 | 4 |
| refusal catalogue entries for this engine | 18 | 20 |

Of the 92 new assertions, 71 cover the checks themselves — including the
corpus's segregation-of-duties case end to end, both polarities satisfied
and broken, a requirement over data checked rather than granted, and six of
the corpus's own English rule lines now *checked* rather than merely parsed
— 19 pin the unreachable counting arithmetic against the corpus's own
numbers, and 2 are the two crash shapes from section 3.

The new static scan holds this engine's copy of the evaluator's fourteen
body-item keywords to the evaluator's own list. Two independent lists that
must agree, with nothing mechanical holding them together, is the actual
defect; the drift would show up as the engine building a rule it then
refuses, naming an internal predicate at a user who never wrote one. Its own
first run over a known-good tree was not clean — it was not joining
continued source lines, and the keyword list is exactly the kind that gets
split across two. **A new check whose first run over a known-good tree is
not clean is reporting its own defect**, which is Entry 2's lesson arriving
again in a different file.

---

## Entry 4 — The first search, on the reference roster *(2026-09-25)*

*Status: **predictions**, written before the ladder ran. The measurements
will be appended below them, with their dates, and the predictions will
stay as written. The project's roadmap tracks this work as `OPTIMIZE.3`,
whose fourth slice runs Entry 1's ladder through a search for the first
time.*

### The ladder, as run

Entry 1's reference roster (section 3a, with the tight needs), written as
an `OPTIMIZE` program over its four Tables, `people`, `shifts`, `leave`
and `next`:

    (rule (free S P) (shifts S W D Slot N K) (people P Nm Sr C) (not (leave P D)))
    (rule (week W) (shifts S W D Slot N K))
    (rule (person-week P W) (people P Nm Sr C) (week W))
    (rule (free-week S P W) (free S P) (shifts S W D Slot N K))
    (rule (night S) (shifts S W D Slot N "senior"))
    (rule (senior-free S P) (free S P) (night S) (people P Nm "yes" C))
    (choose-exactly N (assign S P) (free S P) (per (shifts S W D Slot N K)))
    (choose-at-most 5 (assign S P) (free-week S P W) (per (person-week P W)))
    (choose-at-least 1 (assign S P) (senior-free S P) (per (night S)))
    (forbid (assign S P) (assign T P) (next S T))
    (query assign)

That is everything but the objective and the kept schedule, which later
items build. With no objective, the answer is the first schedule when the
rows are decided in the Tables' own order: shift by shift, within a shift
person by person, each tried on before off.

### Predicted: the counts

These were computed before the ladder ran, by a line-for-line translation
of the search run over the roster built from its definition. The same
translation reproduces all seven counts measured live so far, so a count
below that comes out differently live is a defect in the engine or in this
table, not noise.

| people × weeks | a shift | atoms | counters | clauses | rows laid out | decisions | dead ends | answer |
|---|---|---|---|---|---|---|---|---|
| 5 × 1 | 1 | 90 | 33 | 80 | 889 | 0 | 0 | none, by counting |
| 5 × 4 | 1 | 360 | 132 | 335 | 3,604 | 0 | 0 | none, by counting |
| 10 × 1 | 2 | 180 | 38 | 162 | 1,707 | 35 | 2 | a schedule |
| 10 × 4 | 2 | 720 | 152 | 672 | 6,906 | 140 | 8 | a schedule |
| 20 × 1 | 4 | 360 | 48 | 325 | 3,358 | 80 | 2 | a schedule |
| 20 × 4 | 4 | 1,440 | 192 | 1,345 | 13,579 | 320 | 8 | a schedule |
| 50 × 1 | 10 | 900 | 78 | 814 | 8,277 | 200 | 0 | a schedule |
| **50 × 4** | **10** | **3,600** | **312** | **3,364** | **33,456** | **800** | **0** | **a schedule** |

What it says:
- **At the reference tightness, 0.84, the first schedule is nearly free.**
  There are 840 places to fill. The search makes 800 decisions and takes
  none of them back; the counters force the other 40 in. That agrees with
  clingo, which found a first schedule at this size in 0.06 s (Entry 1):
  finding a schedule is easy here, and proving one best is where the time
  goes.
- **The 5-person roster is refused before any search.** The reason is the
  local one clingo found in 0.003 s (section 3a): the only senior is on
  leave on day 3, so that night has no senior to take it.
- **The ground program is a third of a formula's ceiling**: 33,456 rows laid
  out, against the 100,000 set when this item was scoped.

### Predicted: one person more a shift

At 11 a shift, a tightness of 0.92 rather than 0.84, the search does not
finish:

| budget, in units of work | decisions | dead ends | answer |
|---|---|---|---|
| 50,000 (`normal`) | 25,084 | 24,916 | none within the effort |
| 500,000 (`thorough`) | 250,084 | 249,916 | none within the effort |
| 3,000,000 | 1,500,084 | 1,499,916 | none within the effort |

The same happens at 30, 35, 40 and 45 people, each one person a shift past
its tight need, over 1 week as over 4.

Where it stalls:
- The first dead end comes 175 decisions in, on day 6's early shift.
- For the next half-million units, backtracking never reaches further back
  than the 147th decision, on day 5's late shift.

The cause lies further back, in whatever week 1's first days used up.
Chronological backtracking undoes the most recent decision first, so it
tries every arrangement of the last day and a half and never reaches the
cause. This is Entry 1's trap B, "the same dead end entered twice", which
the roadmap gave to conflict learning and backjumping (`OPTIMIZE.9`), to be
"built when the ladder shows propagation alone stalling short of the
reference roster". The ladder shows it stalling one person a shift past it.
Whether a schedule exists at 11 a shift is not known.

### Predicted: the times

These are my guesses, written before the run and not held to:
- **A whole run at 50 × 4**, meaning DATALOG's pass, the grounding, 800
  decisions and the answer: about 0.3 s.
- **A unit of search work on this roster: about 5 µs.** That is twice
  the 2.2–2.9 µs slice 2 measured on a family whose counters held 14 to 28
  atoms; here a shift's counter holds 42 or 43.

At that rate, `thorough`'s 500,000 units would take about 2.5 s on this
roster, past the 2 s a formula may be projected at. The runs at 11 a shift
exist to answer exactly that: each spends its effort to the last unit, so
its time, less a run at `(effort 1)`, is the search alone. The effort
levels are set from what they measure.

### Measured, 2026-09-25

*Excel 16.0 on the owner's machine. The harness calls the engine
directly rather than through a worksheet formula, so reading the Tables
and spilling the answer are not in these times. Each rung ran again and
again for two seconds, and the mean is given. **Every count in the two
tables above came out exactly as predicted**, on all eleven lines, and
every schedule passed the harness's own check against the roster's
definition.*

| people × weeks | rows laid out | seconds a run |
|---|---|---|
| 5 × 1 | 889 | 0.090 |
| 5 × 4 | 3,604 | 0.176 |
| 10 × 1 | 1,707 | 0.110 |
| 10 × 4 | 6,906 | 0.245 |
| 20 × 1 | 3,358 | 0.142 |
| 20 × 4 | 13,579 | 0.385 |
| 50 × 1 | 8,277 | 0.243 |
| **50 × 4** | **33,456** | **0.745** |

| 50 × 4 at 11 a shift | units of work | seconds a run |
|---|---|---|
| `(effort 1)` | 1 | 0.747 |
| `normal`, then 50,000 | 50,000 | 0.932 |
| `thorough`, then 500,000 | 500,000 | 2.922 |

What it says, against the predictions:
- **A whole run at the reference roster took 0.745 s, two and a half
  times my guess.** The search is almost none of it: its 800 decisions
  come to about 3 ms. Across the eight rungs a run costs about 0.09 s,
  plus 20 µs for every row the grounding lays out. The grounder alone
  costs about 1 µs a row (slice 3's measurement, in the roadmap). So
  most of a run is not the grounding the ceilings count. The next
  measurement splits it.
- **A unit of search work costs 3.7 µs over the first 50,000 units, and
  4.35 µs over 500,000.** I guessed 5; slice 2 measured 2.2–2.9 µs on
  counters of 14 to 28 atoms. The rate rises as the search runs
  longer, most likely because backtracking reaches further back: the
  shallowest decision it returns to moves from the 152nd to the 147th
  between the two budgets.
- **So `thorough` took 2.92 s on the reference roster**, past the 2 s a
  formula may be projected at.

### Before and after

| | before this item | after |
|---|---|---|
| the reference roster's rules, grounded | one shape at a time, through a `DATALOG` formula: the pool less leave 0.16 s, a counter per shift 0.68 s, per person-week 0.69 s, the pair relation 2.44 s (2026-09-19); after `DATALOG.14`, 0.074, 0.047, 0.047 and 0.188 s (2026-09-24) | all of them at once, into 3,600 atoms, 312 counters and 3,364 clauses, then searched and answered: 0.745 s |
| a schedule for it | none: nothing could choose | the first in the Tables' own order, 800 decisions and no dead end |
| the 5-person roster | impossible by hand (section 3a) | refused by counting before any search, naming the night |
| the reference solver, for scale | clingo, a first schedule in 0.06 s | — |

The two sides are not measured the same way. The before column is a
worksheet formula per shape, Tables read and answer spilled. The after
column is one engine call for the whole program.

### What the measurement set

- **The effort levels, halved: `quick` 2,500 units, `normal` 25,000 and
  `thorough` 250,000** (the owner's call, 2026-09-25). `thorough` is the
  most search a formula's 2 s leaves room for on the reference roster:
  about 1.1 s at 4.35 µs a unit, and 1.8 s with everything else. Each
  level stays a tenth of the next. The alternative on the table was to
  give `normal` about a second, as the ceilings' own reasoning had
  assumed. That would have made `thorough` about 11 s, which a formula's
  10 s guard cuts short, so it would only make sense for a command.
- **The ceilings, kept at 100,000 rows in all and 50,000 in a step.**
  They were set on the old evaluator's cost of 7.6 µs a row. The grounder
  they count now costs 0.5 to 1.6 µs a row (slice 3). Raising them looked
  due. But by the
  slope above, a roster-shaped program at 100,000 rows would take about
  0.09 s + 100,000 × 20 µs ≈ 2.1 s in all: the formula's 2 s, reached
  by the part of the run the ceilings do not count. The reference roster
  uses a third of them.
- **A defect from slice 2, found by predicting the new levels.** A search
  could do one unit more than its budget. When the budget's last unit went
  on a decision that ran straight into a dead end, the dead end was counted
  too. The status then read "allows 250000 units of work … and all of them
  went on 125087 decisions and 124914 dead ends", which adds up to
  250,001. The translation found it at the new `thorough`, and at none of
  round 1's budgets, which both happened to end on a dead end. The search
  now stops at such a dead end without taking it. Slice 2's own pin had
  the overshoot written into it: a budget of one unit, answered with one
  decision and one dead end. It has been corrected, and the budget is now
  pinned at its edge from both sides.

### Predicted, before the second run

- The eight rungs: every count as before. None of them spends more than
  800 units of the new `normal`'s 25,000.
- At 11 a shift: `normal` (25,000 units) gives 12,583 decisions and
  12,417 dead ends, and `thorough` (250,000) gives 125,087 and 124,913.
- **`thorough` at about 1.8 s**, where it took 2.92 s.
- Where the reference roster's 0.75 s goes, my guesses: about 0.01 s for
  the memo key, a hash of the program and every Table; 0.2 to 0.3 s for
  `DATALOG`'s own pass over the six rules; and the rest, about 0.45 s,
  for `OPTIMIZE`'s own work around the grounding. I am least sure of the
  last, since round 1's 0.75 s was two and a half times what I expected.

### Measured, the second run, 2026-09-25

*The same harness, on the new levels, after the fix. **Every count came
out as predicted again**, the fixed budget's included: 125,087 decisions
and 124,913 dead ends at `thorough`, exactly 250,000 units. The eight
rungs ran within 0.02 s of the first run's times.*

- **`thorough` now takes 1.871 s at 50 × 4 and 11 a shift.** That is
  0.753 s for everything else and 1.118 s of search, inside a formula's
  2 s, as predicted (about 1.8 s).
- **A unit of search work cost 4.48 µs over the first 25,000 units, and
  4.47 µs over 250,000.** I predicted the first run's rates would repeat,
  "about 3.7 to 4.4 µs"; these sit just above that range.
- *Correction to the first run's reading, 2026-09-25.* "The rate rises as
  the search runs longer" does not survive. Here the first 25,000 units
  cost 4.48 µs, more than the 3.70 µs the first 50,000 cost in the first
  run. A short window's search time is a small difference between two
  larger times: 0.11 to 0.19 s, out of 0.75 to 0.93 s. So a wobble of a
  hundredth of a second in either one moves the rate by about a tenth.
  The two long windows agree to within 3%, at 4.35 and 4.47 µs. A unit of
  search work on this roster costs about 4.4 µs.

Where the reference roster's run goes, against my guesses:

| part of a run, 50 × 4 | measured | guessed |
|---|---|---|
| the memo key: a hash of the program and every Table | 0.095 s | 0.01 s |
| `DATALOG`'s own pass over the program's six rules | 0.495 s | 0.2–0.3 s |
| `OPTIMIZE`'s own work: the grounding the ceilings count, the counters and clauses, the problem, one decision and the answer | 0.170 s | about 0.45 s |
| in all, at `(effort 1)` | 0.760 s | 0.75 s |

The total held and the parts did not. What it says:
- **Two-thirds of a run is `DATALOG`'s pass over the user's own rules.**
  No ceiling counts that part; it was left out when the ceilings were
  set, since `DATALOG` has no ceiling of its own. The part the ceilings
  bound, `OPTIMIZE`'s own work with its grounding inside it, is about a
  fifth. So raising the ceilings would have loosened the smaller part of
  the cost, which is why they were kept.
- **The memo key costs a tenth of a second at this size, and every call
  pays it**, a memo hit included. The key is taken before the memo can be
  asked, so an unchanged roster recalculated costs 0.095 s rather than
  nothing. No item owns that yet.
- **`OPTIMIZE`'s own work is 0.17 s** for 33,456 rows laid out, 3,600
  atoms, 312 counters and 3,364 clauses: about 5 µs for each row laid
  out, everything included.

### What the ladder leaves for later items

- **Plain backtracking stalls one person a shift past the reference
  roster.** At 11 a shift the search never gets out of week 1, at any
  effort. This is the trigger the roadmap set for conflict learning and
  backjumping (`OPTIMIZE.9`), and it has now been pulled on a measured
  roster rather than a guessed one.
- **At real sizes, a formula's time is set by the part that is not
  search.** On the reference roster the search is 3 ms. The grounding,
  at slice 3's rate of 0.5 to 1.6 µs a row, is a few hundredths of a
  second. `DATALOG`'s pass over the user's rules takes half a second, and
  the memo key a tenth.
