# DATALOG.12 - live steps: DATALOG on a real-sized Table

A cost repro, then eight steps, one shape each. Each step is its own program
on its own fresh sheet, building its own small Tables and asking the
grammar's own question. Then one line in the Immediate window grows those
Tables up a ladder of sizes and times each size. The formula timed is the one
the grammar wrote.

The harness is `tools/VLA_Diag12.bas`; the cost repro is
`tools/VLA_Diag12b.bas`. Both are standalone and call nothing in Frazaro.

## What the first pass found (2026-09-18)

Step 1, the simplest shape there is, measured **about 2.5 ms per row**:
0.273s over 100 rows, 2.254s over 1,000, and 25.547s over 10,000, where
Excel stalled for 15-20 seconds and could not be switched to. The cost is
flat per row, so the problem is a large constant, not an algorithm that
degrades with size.

Two predictions held: a change to a cell nothing reads costs 0.000s (DATALOG
is not volatile), and a recalculation after one cell changes costs the same
as a first answer, 2.250s against 2.254s (there is no cache).

**So the ladder is now 100 / 300 / 1,000 / 3,000**, and a size projected past
15s is skipped. At 2.5 ms per row a 3,000-row scan is about 8s; the heavier
shapes will stop before that on their own.

## What the second pass found (2026-09-18)

The first suspect was wrong. `D12bCost` measured
`CreateObject("Scripting.Dictionary")` at 0.157 ms, so the two a row are
about 12% of the cost, and one row's whole simulated shape only 13%. Using a
dictionary (0.0016 ms) and building a key (0.0008 ms) are nothing. **87% of a
row is still unaccounted for.**

Step 2, the join, measured 0.508 · 1.297 · 4.063 · 11.867s — about 4 ms a
row, flat, the same constant as the scan with more of it per row. Every
answer was right.

The next suspect: every `VlaDictGet`/`VlaDictHas`/`VlaDictSet` opens with
`TypeName(d) = "Dictionary"`, and a row makes about a dozen of those calls.
Step 0 below now tests that two ways.

## Where this has got to (2026-09-18)

| run | state |
|---|---|
**The measurement is complete.** Everything below has been run, and the
results are recorded in `DATALOG.12` (BETA_ROADMAP1.md), in
`scripts/pareto_logic.txt` section 18 beside PROLOG's, and as a plain-words
limit in RELEASES.md.

| run | what it found |
|---|---|
| `D12bCost` sections 1-8 | the guard clause: `TypeName` on a Dictionary costs 0.145 ms, about 12 a row, 84% of a row |
| `D12bCost` sections 9-10 | both replacements cost 0.000 ms - `TypeOf d Is Collection` chosen for `DATALOG.13` |
| `D12Parts`, `D12Width` | returning cells to Excel is free; a row costs 1.4 ms plus 0.14 ms per output column |
| Steps 1-6 | scan 2.55, join 3.96, count 2.54, textjoin 2.33, `not` 2.71 ms a row; "otherwise" 13.5 |
| Steps 7 and 8 | closure 15-18 ms a row wide, 216 ms a row over one chain (100 links, 21.6s) |
| Step 5 re-run | the 13.8s recalc was noise: 7.813s first, 7.750s after the change |

The programs and ladders below are kept so the same measurement can be run
again after `DATALOG.13` is built - same sheets, same expected answers, so
before and after are the same measurement.

## Before you start

1. Save the workbook. A size can still hold Excel for a few seconds.
2. Import `tools/VLA_Diag12.bas` and `tools/VLA_Diag12b.bas` into the project
   you run from (VBE, File > Import File). Open the Immediate window
   (Ctrl+G).
3. If you re-run a step, delete its sheet first: Table names are
   workbook-wide. Delete any leftover `D12T<n> (2)` tab too.

**How a ladder stops on its own.** After a size over **10 s**, before a size
projected past **15 s**, and at a refusal, a wrong answer or a VBA error.
Each stop prints a `not run` line saying why. If Excel is still busy after
**2 minutes**, end it in Task Manager and send the last `D12|` line.

**What each size prints:**

```
D12|step|rows|first s|change s|unrelated s|answer rows|expected|verdict|note
```

`first s` is the answer recalculated after the Tables were rewritten;
`change s` after ONE Table cell changes; `unrelated s` after a cell nothing
reads changes. The `answer rows` and `expected` columns are read AFTER the
one-cell change, so they show the post-change count. The same rows are kept
on each step's sheet from column AD.

**Please also note, for any size over about 5 s:** did the title bar say
"(Not Responding)", and did Excel come back on its own?

---

## Step 0 - where the 2.5 ms per row goes (run both of these first)

The first writes no cells at all; the second writes four formulas to step 1's
sheet and clears them afterwards. Both run in seconds and neither should
stall Excel.

**Type:** `D12bCost`

Sections 1 to 5 are the ones you already ran, unchanged, so those numbers
should simply repeat. Sections 6 to 8 are new: `TypeName` on a
`Scripting.Dictionary`, the same on a `Collection` for contrast, and one
row's shape again with the dozen wrapper calls a row really makes, each
paying its own `TypeName` first. If section 8 lands near 2.5 ms where
section 5 reached an eighth of it, the wrappers are the answer.

**Then type:** `D12Parts 1000`

This one asks the engine instead of a model. It rewrites Bills12x1 at 1,000
rows, then times four formulas over that same Table, each adding one stage to
the one before: reading the Table and spilling it whole, then one rule, then
the comparison, then the question's narrowing rule (which is the whole step-1
program). The difference between two lines is what that stage costs per row.
It needs sheet D12T1 from step 1, writes to L2, P2, S2 and V2, and clears
them when it is done.

**Send back:** both outputs whole.

### Pass 5: two numbers the fix's fork needs

`D12bCost` has two new sections at the end, 9 and 10: `TypeOf d Is
Collection` asked of a Dictionary, and a Boolean read per call. They are the
two cheapest replacements for the guard clause that `DATALOG.13` is about, so
timing them decides that item's fork on numbers rather than on taste.
Sections 1 to 8 are unchanged and repeat as their own control.

**Type:** `D12bCost`

**Send back:** the output, or just sections 9 and 10 if the rest repeats.

### The width probe (pass 4, done)

`D12Parts` left one question open: its probes each varied the WORK and the
answer's WIDTH together, so they cannot say which one cost. This varies one
thing at a time.

**Type:** `D12Width 1000`, then `D12Width 3000`

Five formulas over the same Bills12x1 rows: three DATALOG questions doing
identical per-row work but answering 1, 2 and 3 columns, then a bare VBA UDF
(`D12Cells`) that only hands Excel an array of the same shape, with no
engine behind it at all. The difference between the DATALOG lines is what one
more column costs per row; the bare UDF says what that costs with no DATALOG
in the picture. Running it at both sizes separates the per-row cost from the
per-cell one.

It writes to AR2, AT2, AW2, BB2 and BD2 on sheet D12T1 and clears them
afterwards.

**Send back:** both outputs.

---

## Step 1 - a scan (`which bill is big`)

Already run at the old ladder. Re-running adds the 300 and 3,000 sizes and
costs about 10s in total: delete sheet D12T1 first, then Interpret:

```
Work on sheet D12T1.
Put "Bill" into cell A1.
Put "Vendor" into cell B1.
Put "Amount" into cell C1.
Put "B1" into cell A2.
Put "V1" into cell B2.
Put 7400 into cell C2.
Put "B2" into cell A3.
Put "V2" into cell B3.
Put 14800 into cell C3.
Put "B3" into cell A4.
Put "V3" into cell B4.
Put 2200 into cell C4.
Put "B4" into cell A5.
Put "V4" into cell B5.
Put 9600 into cell C5.
Turn A1:C5 into a table called Bills12x1.
Write in cell H2 that a bill is big if Bills12x1 lists the bill as Bill and the amount as Amount, and the amount is greater than 10000.
Show in cell J2 which bill is big by applying the rules in H2:H2 to the data tables Bills12x1.
```

**The cell:** J2 shows `Bill`, and J3 shows `B2`.

**Then type:** `D12Ladder 1`

**Expected answer rows:** 49 · 147 · 490 · 1,470, each one higher after the
change.

## Step 2 - a join (the README's `can-cover`)

**Already run (2026-09-18): 0.508 · 1.297 · 4.063 · 11.867s, every answer
right.** Skip it unless you want it again after a change; the program is kept
here for that.

Interpret:

```
Work on sheet D12T2.
Put "Name" into cell A1.
Put "Cert" into cell B1.
Put "Level" into cell C1.
Put "P1" into cell A2.
Put "C1" into cell B2.
Put 2 into cell C2.
Put "P11" into cell A3.
Put "C1" into cell B3.
Put 3 into cell C3.
Put "P3" into cell A4.
Put "C3" into cell B4.
Put 1 into cell C4.
Turn A1:C4 into a table called Staff12x2.
Put "Shift" into cell E1.
Put "Needs" into cell F1.
Put "MinLevel" into cell G1.
Put "S1" into cell E2.
Put "C1" into cell F2.
Put 2 into cell G2.
Put "S2" into cell E3.
Put "C3" into cell F3.
Put 1 into cell G3.
Turn E1:G3 into a table called Shifts12x2.
Put "Name" into cell I1.
Put "Shift" into cell J1.
Put "P1" into cell I2.
Put "S1" into cell J2.
Turn I1:J2 into a table called Leave12x2.
Write in cell L2 that a person can-cover a shift if Shifts12x2 lists the shift as Shift, the cert as Needs, and the min as MinLevel, and Staff12x2 lists the person as Name, the cert as Cert, and the level as Level, and the level is at least the min, and not Leave12x2 lists the person as Name and the shift as Shift.
Show in cell N2 who can-cover "S1" by applying the rules in L2:L2 to the data tables Staff12x2, Shifts12x2, and Leave12x2.
```

**The cell:** N2 shows `Who`, and N3 shows `P11`. P1 is on leave.

**Then type:** `D12Ladder 2`

**The sizes:** each size is the number of Staff rows. Shifts stays at 20 rows
and Leave has one row per 20 staff. The join makes about 2 pairs per Staff
row, so this should cost more per row than step 1.

**Expected answer rows:** 4 · 10 · 34 · 100, each one higher after the
change (P51's level goes to 3).

## Step 3 - a count per group

Interpret:

```
Work on sheet D12T3.
Put "Name" into cell A1.
Put "Dept" into cell B1.
Put "P1" into cell A2.
Put "D1" into cell B2.
Put "P2" into cell A3.
Put "D1" into cell B3.
Put "P3" into cell A4.
Put "D2" into cell B4.
Turn A1:B4 into a table called Staff12x3.
Put "Dept" into cell D1.
Put "D1" into cell D2.
Put "D2" into cell D3.
Put "D3" into cell D4.
Turn D1:D4 into a table called Depts12x3.
Write in cell F2 that a person works-in a dept if Staff12x3 lists the person as Name and the dept as Dept.
Write in cell F3 that a dept is listed if Depts12x3 lists the dept as Dept.
Show in cell H2 how many people works-in each dept that is listed by applying the rules in F2:F3 to the data tables Staff12x3, and Depts12x3.
```

**The cells:** H2:I2 show `Dept | People`. Below them: `D1 2`, `D2 1`,
`D3 0`.

**Then type:** `D12Ladder 3`

**The data:** Staff spread over depts D0 to D99, plus a listed D100 with
nobody in it.

**Expected answer rows:** 101 at every size. The harness also checks D100
reads 0, and 1 after P1 moves there.

## Step 4 - a textjoin per group

Interpret:

```
Work on sheet D12T4.
Put "Name" into cell A1.
Put "Dept" into cell B1.
Put "P1" into cell A2.
Put "D1" into cell B2.
Put "P2" into cell A3.
Put "D1" into cell B3.
Put "P3" into cell A4.
Put "D2" into cell B4.
Turn A1:B4 into a table called Staff12x4.
Put "Dept" into cell D1.
Put "D1" into cell D2.
Put "D2" into cell D3.
Put "D3" into cell D4.
Turn D1:D4 into a table called Depts12x4.
Write in cell F2 that a person works-in a dept if Staff12x4 lists the person as Name and the dept as Dept.
Write in cell F3 that a dept is listed if Depts12x4 lists the dept as Dept.
Show in cell H2 which people works-in each dept that is listed as one list by applying the rules in F2:F3 to the data tables Staff12x4, and Depts12x4.
```

**The cells:** H2:I2 show `Dept | People`. Below them: `D1 P1, P2`,
`D2 P3`, and `D3` with a blank beside it.

**Then type:** `D12Ladder 4`

**The data:** ten depts, D0 to D9, plus a listed D10 with nobody in it.

**Expected answer rows:** 11 at every size. The longest list runs 49 · 169 ·
590 · 1,990 characters.

**Named, and now out of reach:** the 32,767-character cell limit needs about
4,000 rows in one group, past this ladder. It stays predicted, not measured,
unless you want a one-off probe for it.

## Step 5 - `not` through a projection

Interpret:

```
Work on sheet D12T5.
Put "Name" into cell A1.
Put "Dept" into cell B1.
Put "P1" into cell A2.
Put "D1" into cell B2.
Put "P2" into cell A3.
Put "D2" into cell B3.
Put "P3" into cell A4.
Put "D3" into cell B4.
Turn A1:B4 into a table called Staff12x5.
Put "Name" into cell D1.
Put "Day" into cell E1.
Put "P3" into cell D2.
Put "Day1" into cell E2.
Turn D1:E2 into a table called Leave12x5.
Write in cell G2 that a person is available if Staff12x5 lists the person as Name, and not Leave12x5 lists the person as Name.
Show in cell I2 which person is available by applying the rules in G2:G2 to the data tables Staff12x5, and Leave12x5.
```

**The cell:** I2 shows `Person`, and below it `P1` and `P2`.

**Then type:** `D12Ladder 5`

**The data:** every third person is on leave. The rule's `not` reads a
generated projection that drops Day.

**Expected answer rows:** 67 · 200 · 667 · 2,000, each one higher after the
change.

## Step 6 - "otherwise" guards

Interpret:

```
Work on sheet D12T6.
Put "Customer" into cell A1.
Put "Spend" into cell B1.
Put "K1" into cell A2.
Put 3700 into cell B2.
Put "K2" into cell A3.
Put 7400 into cell B3.
Put "K3" into cell A4.
Put 11100 into cell B4.
Turn A1:B4 into a table called Customers12x6.
Write in cell D2 that a customer has-tier "Gold" if Customers12x6 lists the customer as Customer and the spend as Spend, and the spend is at least 10000, otherwise "Silver" if Customers12x6 lists the customer as Customer and the spend as Spend, and the spend is at least 5000, otherwise "Bronze" if Customers12x6 lists the customer as Customer.
Show in cell F2 who has-tier "Bronze" by applying the rules in D2:D2 to the data tables Customers12x6.
```

**The cell:** F2 shows `Who`, and F3 shows `K1`.

**Then type:** `D12Ladder 6`

**Expected answer rows:** 31 · 100 · 331 · 1,000, each one LOWER after the
change. Bronze is the branch behind two guards, so this is five rules over
the same Table.

## Step 7 - closure over a wide org

Interpret:

```
Work on sheet D12T7.
Put "Employee" into cell A1.
Put "Manager" into cell B1.
Put "E2" into cell A2.
Put "E1" into cell B2.
Put "E3" into cell A3.
Put "E1" into cell B3.
Put "E4" into cell A4.
Put "E2" into cell B4.
Turn A1:B4 into a table called Reports12x7.
Write in cell D2 that a person reports-to a boss if Reports12x7 lists the person as Employee and the boss as Manager.
Show in cell F2 who reports-to "E1" directly or not by applying the rules in D2:D2 to the data tables Reports12x7.
```

**The cell:** F2 shows `Who`, and below it `E2`, `E3` and `E4` in any order.
E4 is there only through E2.

**Then type:** `D12Ladder 7`

**The data:** ten reports each, and everyone reaches E1.

**Expected answer rows:** 100 · 300 · 1,000 · 3,000, each one lower after
the change.

**What DATALOG builds:** all 190 · 780 · 2,880 · 10,770 pairs, 2 · 3 · 3 · 4
links deep, before it narrows to E1.

## Step 8 - closure over one chain (the stress case; run it last)

Interpret:

```
Work on sheet D12T8.
Put "Employee" into cell A1.
Put "Manager" into cell B1.
Put "E2" into cell A2.
Put "E1" into cell B2.
Put "E3" into cell A3.
Put "E2" into cell B3.
Put "E4" into cell A4.
Put "E3" into cell B4.
Turn A1:B4 into a table called Chain12x8.
Write in cell D2 that a person reports-to a boss if Chain12x8 lists the person as Employee and the boss as Manager.
Show in cell F2 who reports-to "E1" directly or not by applying the rules in D2:D2 to the data tables Chain12x8.
```

**The cell:** F2 shows `Who`, and below it `E2`, `E3` and `E4`.

**Then type:** `D12Ladder 8`

**The ladder:** 100 · 250 · 500 · 1,000 links, its own, because the cost
here is quadratic: the harness assumes it at least quadruples when the chain
doubles, and will stop early if it grows faster.

**Expected answer rows:** 100 · 250 · 500 · 1,000, each one lower after the
change.

**What DATALOG builds:** 5,050 · 31,375 · 125,250 · 500,500 pairs, over
about as many rounds as there are links. If step 8 is the one that refuses
or runs out of memory, that is the finding, not a failure.

---

## What to send back

- `D12bCost`'s whole output (step 0).
- Every `D12|` line and each step's `formula:` line.
- For any size over about 5 s: "(Not Responding)" or not, and whether Excel
  came back.
- Anything that crashed Excel or raised a VBA error, with the step and size.
- Dev workbook or built add-in.

**Afterwards:** delete sheets D12T1 to D12T8 and remove both modules.
