# OPTIMIZE.0 - live steps: the shape of the search, measured

OPTIMIZE does not exist yet. These steps measure what it will pay first,
**grounding**, using the engine that does exist (`DATALOG`), and they probe what
Excel allows a long-running formula to do. Nothing here changes Frazaro.

- **Part A**: the ladder, 13 steps in the dev workbook. Each step is one
  Immediate-window line that builds its own fresh sheet and its own Tables.
- **Part B**: the reference fixtures, 4 steps in the dev workbook.
- **Part C**: the host probe, 10 steps in a **new blank workbook**.
- **Part D**: clingo, optional, from PowerShell. Nothing in Excel.

The predictions were written before any of this ran. They are in
`docs/OPTIMIZATION.md`, Entry 1. Every expected row count below was derived
twice: once by the harness's closed forms, and once by enumeration in
`tools/optimize0_expected.ps1`. The two agree on every rung.

## Before you start

1. **Save** the dev workbook (`VLA.xlsm`). Close every other workbook:
   `Application.Calculate` also recalculates their volatile formulas, and that
   would be noise in the timings.
2. Import `tools/VLA_DiagO0.bas` into `VLA.xlsm`, then run Debug > Compile.
   It should compile clean.
3. Open the Immediate window (Ctrl+G).
4. Each ladder step makes a sheet `O0T<n>`. To re-run a step, delete its
   sheet first; the step refuses to reuse one. Every result is also appended to
   the sheet `O0Results`, one row per rung.

**How to read an `O0|` line:** step, people, weeks, k, the ground rows, the first
answer's seconds, the control's seconds, the model's seconds (with its band),
**measured/model**, the model's peak rows, the verdict, and a note.

- `ok` means every row was checked against the fixture's definition.
- `not run` is a guard doing its job. It is not a failure.
- Anything else (`WRONG`, `REFUSED`, `VBA error`) stops the step. Send it back.

---

# Part A - the ladder (dev workbook)

## Step A0 - the model alone, touching nothing

**Type:** `O0Model`

**Expect:** 104 `O0M|` lines and no change to any sheet. Exactly **9** lines end
`never run: memory`:

- step 8: `50|1|2`, `50|4|2` and `50|4|10`;
- step 9: `20|4`, `50|1` and `50|4`;
- step 10: `20|1|3`, `5|1|5` and `50|4|5`.

Lines ending `skipped unless measured faster than the model` may also appear.

**Send back:** only whether the count of `never run: memory` lines is 9.

## Steps A1 to A12

For each step, type `O0Ladder <n>` with the step number. Every rung that runs
must read `ok`. Its `ground rows` must be the number given here. The rungs are
in order: people × weeks, then k where there is one.

**A1 - the pool: anyone, any shift (a Cartesian body).** `O0Ladder 1`
- Rows: 5×1 **105**, 10×1 **210**, 5×4 **420**, 20×1 **420**, 10×4 **840**,
  50×1 **1,050**, 20×4 **1,680**, 50×4 **4,200**.
- Model: all under 0.2 s.
- If this step reads `REFUSED`, DATALOG refuses a Cartesian body today. That
  would be a finding: send the text back.

**A2 - the pool minus vacations (`not`).** `O0Ladder 2`
- Rows: **90, 180, 360, 360, 720, 900, 1,440, 3,600**, which is 18 × people ×
  weeks.
- Model: under 0.25 s.

**A3 - the pool through a skill join (nights need a senior).** `O0Ladder 3`
- Rows: **77, 154, 308, 315, 616, 784, 1,260, 3,136**.
- Model: under 0.15 s.

**A4 - a native counter per shift (`count`).** `O0Ladder 4`
- Rows: **21, 21, 84, 21, 84, 21, 84, 84**, one per shift.
- Each shift's count is people minus those on leave that day.
- Model: 0.63 s at 50×4.

**A5 - a native counter per person-week.** `O0Ladder 5`
- Rows: **5, 10, 20, 20, 40, 50, 80, 200**, and every count is **18**.
- Model: 0.65 s at 50×4.

**A6 - never two in a row, through the pair relation `Next`.** `O0Ladder 6`
- Rows: **100, 200, 415, 400, 830, 1,000, 1,660, 4,150**.
- Model: 2.3 s at 50×4.

**A7 - never two in a row, by arithmetic over all pairs.** `O0Ladder 7`
- The same rows as A6. The same answer, reached through people × shifts² join
  rows.
- Model: 3.7 s at 20×4 and 9.3 s at 50×4, whose band reaches 24 s. So 50×4 may
  be skipped by the time guard.

**A8 - at most 2 per shift, as triples in canonical order.** `O0Ladder 8`
- Rows: 5×1 **210**, 7×1 **735**, 10×1 **2,520**, 5×4 **840**,
  14×1 **7,644**, 7×4 **2,940**, 20×1 **23,940**, 10×4 **10,080**,
  14×4 **30,576**, 20×4 **95,760**.
- Never run: 50×1 and 50×4 (the memory guard), and 50×4 with k = 10, the tight
  reference's "at most 10 per shift".
- Model: 8.7 s at 20×4, with a band of 5–23 s.

**A9 - at most 2 per shift, as triples WITHOUT canonical order.** `O0Ladder 9`
- Rows: **1,260, 4,410, 15,120, 5,040, 45,864, 17,640, 143,640, 60,480,
  183,456**. That is six times A8's rows, the (k+1)! of the symmetric trap.
- Never run: 20×4, 50×1 and 50×4 (memory).
- Expect the time guard to cut earlier than in A8.

**A10 - at most k a week, as (k+1)-subsets of a person's week.** `O0Ladder 10`
- Rows, as people × weeks, k:
  - 1×1, k=1: **210**
  - 5×1, k=1: **1,050**
  - 1×1, k=2: **1,330**
  - 10×1, k=1: **2,100**
  - 20×1, k=1: **4,200**
  - 5×1, k=2: **6,650**
  - 1×1, k=3: **5,985**
  - 10×1, k=2: **13,300**
  - 20×1, k=2: **26,600**
  - 5×1, k=3: **29,925**
  - 10×1, k=3: **59,850**
  - 1×1, k=5: **54,264**
- Never run: 20×1 k=3, 5×1 k=5 and 50×4 k=5 (memory).
- **The 1×1, k=5 rung is "at most 5 a week" measured directly on one
  person-week, C(21,6).** Its model is 10.2 s with a band of 5–30 s. If the
  time guard skips it, say so.

**A11 - overtime per person-week, against the kept schedule.** `O0Ladder 11`
- Rows: **1, 1, 2, 2, 4, 5, 8, 20**. Every row's overtime is **1**.
- Model: under 0.35 s.

**A12 - the control: a bare UDF returning an array, no engine.** `O0Ladder 12`
- Rows × columns: 1,000 × 4, 10,000 × 4, 100,000 × 4, 100,000 × 8,
  430,000 × 4, 430,000 × 8. Each must read `ok`.
- This step says whether handing back the largest answers above costs anything.
  It was measured free at 9,000 cells (`DATALOG.12`), and never at 3.4 million.

**Send back:** every `O0|` line from A1 to A12, as printed.

## Added after step A7 (2026-09-19)

A7's top rung ran at 1.76 times its model: 16.4 s at 50×4, where the rung
before it read 1.21. Its per-row cost rose by half at the one rung whose
relation held about 350,000 keyed rows. The two steps below separate the
suspected cause from everything else. **Re-import `tools/VLA_DiagO0.bas`
first**: remove the old module, then import the file again.

**A7b - Dictionary growth alone, with no engine.** `O0DictCost`
- It times the same work `RelTryAdd` does for each row: a four-column key,
  `Exists`, `Add`, and keeping the tuple. It does this at 10,000, 50,000,
  100,000, 200,000 and 400,000 keys, beside building the key alone and
  releasing the Dictionary afterwards.
- **If the "exists+add+keep" column stays flat, the Dictionary is cleared**,
  and the growth lies elsewhere.
- **If it rises between 100,000 and 400,000 keys, that is the cause**, measured.
- It stops after any size over 10 s. The whole run should take well under a
  minute.

**A rung the time guard skipped, run on purpose.** `O0Ladder <n>, <rung>`
- This runs only that rung (1 is the first in the step's list), without the
  time guard. The memory guard still applies.
- Delete the step's sheet first.
- The one likely to matter: A10's rung 12, one person-week of "at most 5 a
  week" (C(21,6) = 54,264 rows). If A10 skips it, run `O0Ladder 10, 12`. Save
  first: it may stall Excel for 15–20 s.

**For A8, A9 and A10: save first.** Each step's top rung that the memory guard
allows may now stall Excel for 15–20 s. The time guard lets one rung through
past its 10 s limit by design.

---

# Part B - the reference fixtures (dev workbook)

These are the Tables every later OPTIMIZE item will be measured against. Each
step writes a fresh sheet and prints `O0F|` lines.

## Step B1 - the reference roster (tight)

**Type:** `O0Fixture 50, 4`

**Expect:** sheet `O0F50x4`, and these values:

| line | value |
|---|---|
| need | 10 on every shift (tight) |
| weekly demand | 210 |
| weekly capacity at 5 each | 250 |
| tightness | 0.84 |
| seniors | 12 |
| leave rows | 200 |
| pool, full | 4200 |
| pool, minus leave | 3600 |
| kept rows | 840 |
| kept on leave | 118 |
| kept two in a row | 0 |
| kept over five a week | 0 |
| kept nights with no senior | 0 |
| kept person-weeks over contract | 20 |
| a schedule exists | not decided by hand - tools/optimize0_lp.ps1 exports this fixture for clingo |

## Step B2 - the ladder's own rung

**Type:** `O0Fixture 10, 1`

**Expect:**

| line | value |
|---|---|
| need | 2 on every shift (tight) |
| weekly demand | 42 |
| weekly capacity at 5 each | 50 |
| tightness | 0.84 |
| seniors | 2 |
| leave rows | 10 |
| pool, full | 210 |
| pool, minus leave | 180 |
| kept rows | 42 |
| kept on leave | 7 |
| kept two in a row | 0 |
| kept over five a week | 0 |
| kept nights with no senior | 5 |
| kept person-weeks over contract | 1 |
| a schedule exists | not decided by hand |

## Step B3 - a fixture with no schedule, by hand (too few seniors)

**Type:** `O0Fixture 5, 1`

**Expect:**
- need 1; weekly demand 21; capacity 25; tightness 0.84; seniors 1;
  leave rows 5.
- Kept: rows 21, on leave 3, two in a row 0, over five 0, nights with no
  senior 6, over contract 1.
- **"a schedule exists": NO, provably**: 7 nights a week each need a senior,
  and 1 senior may work at most 5 a week.

## Step B4 - a fixture with no schedule, by hand (too much demand)

**Type:** `O0Fixture 10, 1, True`

**Expect:**
- Sheet `O0F10x1L`.
- Need "Early 2, Late 3, Night 4 (loose)"; weekly demand 63; capacity 50;
  tightness 1.26; seniors 2; leave rows 10.
- Kept: rows 63, on leave 10, two in a row 0, over five 10, nights with no
  senior 1, over contract 10.
- **"a schedule exists": NO, provably**: the shifts need 63 people a week, and
  10 people may work at most 5 each.

**Send back:** the `O0F|` lines of each fixture, but only where a value differs
from these tables.

---

# Part C - the host probe (a NEW blank workbook)

**Before Part C:**

1. Save and close the dev workbook.
2. Open a new blank workbook, save it as `O0Probe.xlsm`, and import
   `tools/VLA_ProbeO0.bas`. Compile.
3. Type `O0ProbeInfo` in the Immediate window and send back its five lines.

**For each step `n` below:**

1. Type `O0ProbeSheet n`. It makes sheet `O0P<n>`, shows what to type, and
   selects B5.
2. **Type the formula into B5 yourself** and press Enter. A formula written by a
   macro would calculate inside that macro, which is a different question.

**If Excel stops responding:** a modal hidden behind the window is the likeliest
cause, so check Alt+Tab and the taskbar first. Every loop here ends by itself
within 30 seconds. If Excel is still frozen after 60 seconds, end it in Task
Manager, reopen it, and let Document Recovery run. **Then restart Excel cleanly
once more before continuing.** A forced kill has left bad session state before.

**The predictions below are my best guesses**, not known behaviour. Each step
says what each outcome would mean. Send back the cell's text and the `O0P|`
lines for every step.

**C1 - Esc against a looping UDF, EnableCancelKey untouched.**
- Type `=O0Spin(20)` and press Enter. After about 3 s, press Esc once.
- **Prediction (held weakly):** the loop is not interrupted. After about 20 s
  the cell reads `ran 20.0s, not interrupted (...)`.
- If instead it reads `interrupted at ~3s (error 18 ...)`, or you see a "Code
  execution has been interrupted" dialog, Esc *does* reach VBA inside a UDF.
- Then press F9 and note whether it spins again. It should not: the formula is
  not volatile.

**C2 - the same, with EnableCancelKey = xlErrorHandler.**
- Type `=O0Spin(20,2)` and press Esc at about 3 s.
- If the cell reads `interrupted at ~3s (error 18: ...)`, a UDF can catch Esc
  and hand back a clean result: the "budget stops, say so" shape.
- If it reads `ran 20.0s ... EnableCancelKey could not be set`, a UDF cannot
  change it.

**C3 - the same, with EnableCancelKey = xlDisabled.**
- Type `=O0Spin(20,0)` and press Esc at about 3 s.
- **Expect:** `ran 20.0s, not interrupted`. Esc is ignored.

**C4 - Esc against a looping MACRO: the command form's path.**
- In the Immediate window, type `O0SpinMacro 20` and press Esc at about 3 s.
  **Expect** the "Code execution has been interrupted" dialog. Click **End**.
- Then `O0SpinMacro 20, 2` with Esc. **Expect** no dialog, and the line
  `O0P|macro|end|interrupted|~3s|error 18: ...`.
- Then `O0SpinMacro 20, 0` with Esc. **Expect** Esc ignored:
  `...ran to the end|20.0s`.

**C5 - a MsgBox from inside a UDF.**
- Type `=O0Ask(B1)`. **Prediction:** a box "O0 probe: a MsgBox from inside a
  worksheet function. Continue?" appears. Click Yes. The cell reads
  `answered Yes (ask 1)`.
- Then press **F9**. **Expect** no box, since nothing changed.
- Then press **Ctrl+Alt+F9**. **Expect** the box again (ask 2): a full
  recalculation asks whoever is at the keyboard.

**C6 - the Function Wizard.**
- Type `=O0Calls(B1)`: the cell reads 1.
- Select the cell and click **fx**. Note how many new `O0P|calls|` lines
  appear.
- In the argument box, change `B1` to `B12` one keystroke at a time. Note the
  lines per keystroke, then Cancel.
- **Prediction:** the wizard evaluates the function once when it opens, and
  once after every keystroke. That is how many times a modal inside a UDF would
  ask, which is why modals are capped here.
- Then run `O0ProbeReset` and repeat with `=O0Ask(B1)`. Record how many boxes
  appear. The cap stops it at 3.

**C7 - a modal UserForm from inside a UDF.**
- In the VBA editor: Insert > UserForm, and set its (Name) to `O0Form`. Leave it
  empty.
- Back on the sheet, type `=O0ShowForm(B1)`.
- **Prediction:** the form shows. Close it with its X, and the cell reads
  `form shown and closed (show 1)`.
- If the cell reads `error ...`, a UDF cannot show a form. Record the error.

**C8 - the status bar, set from inside a UDF.**
- Type the formula the sheet shows: `=O0Status("O0 probe: a better roster may
  exist - search longer")`.
- **Prediction:** the status bar shows that text, and the cell reads
  `set; the status bar reads back: O0 probe: ...`.
- If it reads `error ...`, the non-blocking invitation cannot live in a
  formula.
- Run `O0StatusClear` afterwards.

**C9 - a macro scheduled from inside a UDF.**
- Type `=O0Later(B1)`. The cell reads `scheduled O0LaterMacro ...`.
- **Prediction:** a second later, C5 reads `written by a macro a UDF scheduled,
  at hh:mm:ss`.
- If so, a formula *can* cause a write, one step removed. The design record's
  "a UDF cannot write" needs correcting, although the no-modal decision stands
  on its other three reasons.

**C10 - DoEvents inside a UDF, and re-entrancy. Save first: this is the
riskiest step.**
- Type `=O0Pump(10,B1)`. While it runs (10 s):
  1. type `5` in **B1** and press Enter (B1 is the formula's own input);
  2. type `x` in **D9** and press Enter;
  3. press **Esc**.
- **What decides it:**
  - Whether your typing landed at all.
  - Whether any line reads `O0P|pump|enter|depth 2`. That would mean Excel
    re-entered the UDF while it was still running.
  - The final cell text, with its "deepest call seen".
- **Prediction (weak):** the typing is accepted, and depth stays at 1.

---

# Part D - clingo, the dev-only oracle (optional, PowerShell)

Settled 2026-09-19: clingo checks OPTIMIZE's answers at sizes hand-checking
cannot reach. It is installed by you and run by you. It is never shipped, and
nothing in Frazaro calls it. Compare only two things: satisfiable or not, and
the "Optimization:" line. Never the schedule itself.

Run these in PowerShell, from the repository folder, with clingo's environment
active. The exporter writes each `.lp` file to `$env:TEMP`. (PowerShell writes
the temp folder as `$env:TEMP`; `%TEMP%` is cmd.exe syntax and reaches clingo
as literal text.)

1. The toy:
   `powershell -NoProfile -ExecutionPolicy Bypass -File tools\optimize0_lp.ps1 -Toy`
   then `clingo "$env:TEMP\optimize0_toy.lp" 0`.
   **Expect:** `Models : 7290`, the same as the toy's hand count (10 × 3⁶).
2. `... optimize0_lp.ps1 -People 5 -Weeks 1`, then
   `clingo "$env:TEMP\optimize0_5x1.lp"`.
   **Expect:** `UNSATISFIABLE` (too few seniors, B3).
3. `... optimize0_lp.ps1 -People 10 -Weeks 1 -Loose`, then
   `clingo "$env:TEMP\optimize0_10x1L.lp"`.
   **Expect:** `UNSATISFIABLE` (too much demand, B4).
4. `... -People 10 -Weeks 1` (`optimize0_10x1.lp`), then `-People 20 -Weeks 1`
   (`optimize0_20x1.lp`), then `-People 50 -Weeks 4` (`optimize0_50x4.lp`,
   with `--time-limit=600` after the file name).
   **Record:** SATISFIABLE or not, the last "Optimization:" line, and whether
   it says `OPTIMUM FOUND`. These numbers are not known yet. They become the
   answer key for `OPTIMIZE.3` and `OPTIMIZE.6`.

5. **Value order (added 2026-09-19, after step 4c).** Step 4c ran for 600 s
   and ended at `0 1486`: 1,486 changes against a hand lower bound of 236.
   (The 118 kept assignments on a leave day each cost at least a drop plus an
   add.)
   `... optimize0_lp.ps1 -People 50 -Weeks 4 -KeptFirst`, then
   `clingo "$env:TEMP\optimize0_50x4.lp" --heuristic=Domain --time-limit=120`.
   The hint only changes which schedule clingo tries first, never which costs
   are optimal. **Prediction:** it gets well under 1,486 changes within 120 s,
   probably a few hundred, and the optimum is still unknown.

**Send back:** each run's result line and its Optimization line.
