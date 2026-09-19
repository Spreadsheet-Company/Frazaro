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
