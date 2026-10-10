# PROTOCOL - the study before the viewport

*`KERNEL.3`, the instrument `SD-29` requires before the viewport (`KERNEL.5`) is built: a fixture workbook with planted defects, a task sheet in two halves, the protocol that runs them on the existing page beside Excel against Excel alone, the measures, the decision rule written before the first run, and the result table, blank. This file is the instrument; the running is the owner's, with whoever they choose, and the numbers go into `web/CALLOSUM.md` as a dated measurement under the slot of 2026-10-08. Written 2026-10-08 from the sitting of the previous session, whose decisions it carries and does not reopen.*

*What this document is not. It is not a result: every number in section 9 is blank until a dated entry fills it, and a blank cell is a blank cell, never a zero. It is not a human-subjects study with a claim to generality: the first runner is a pilot of one, the person who built the tool, which section 5.6 says plainly. House rules as for every ledger in `docs/`: append-only; a correction is added and dated, never written over; where this file and the code disagree, the code is right; where this file and `web/CALLOSUM.md` disagree, the later dated entry governs. Marks follow `HORIZON.md` and `web/CALLOSUM.md`: an argument about the present is* **exhibit** *(a dated, public, checkable fact, its source in section 11),* **shelf** *(a claim a file in this repository already makes, cited) or* **recommendation** *(the instrument's choice, the owner's to change before the first run and never after it); every figure is a* **measurement**, *dated, with its method, or a* **prediction** *written before the measurement that tests it; a figure that is only arithmetic over stated inputs is* **derived**.

---

## 1. The question, and the decision it feeds

The viewport needs three decisions before its first line is written, and each is a decision about people or about a machine, never about taste. Canvas or DOM: whether the right pane draws its cells with the 2D context or with elements (`KERNEL.5`'s entry presumes a canvas, CALLOSUM §7 decision 5's recommendation; this study decides it). The formula bar's first line: which of value, formula and sentence (`SD-24`) a selected cell shows first. The pane's default for an author: whether the left pane is open when a program's author opens the page, given that `SD-28` already closes it for a reader. `SD-29` says a change to the interface is preceded by a study of construction and audit tasks, timed and scored, so that the numbers decide rather than confirm (*shelf*: `BETA_ROADMAP.md`'s register; `web/CALLOSUM.md` §13.2 item 10). This instrument has two halves because the three decisions do: the first is an engineering question a benchmark answers (section 7), the other two are human questions the study bears on (sections 4 and 5), and section 8 reads both into answers before anything has run.

## 2. Backed into from the literature

The house has no panel of analysts to ask, so the tasks and the questions are backed into from what is published about how people read spreadsheets and how spreadsheets go wrong, the way `scripts/pareto.txt` chose its sentence corpus to cover what real procedures are known to do, in its own header's words, rather than from a survey the project could not run (*shelf*; *recommendation*). The sources, and what each one lends:

- **Hendry and Green (1994)** lend the task shapes. Their study of discretionary spreadsheet users names three activities, creating, comprehending and explaining, and finds the second and third hard, the formulas being hard to understand (*exhibit*); Sarkar et al. (2018) put the same weakness as the computation hidden behind the values (*exhibit*). The comprehension questions of section 4.3 are in their shape, as `SD-29` already words them: choose a cell, explain its formula, find its inputs. The build half is their creating; the find half is their comprehending turned into an audit.
- **Panko and Halverson (1996) and Panko (1998, revised 2008)** lend the top of the defect taxonomy, mechanical, logic and omission errors, and the base rate the sitting already cites, that most audited spreadsheets hold at least one error (*exhibit*; *shelf*: `web/CALLOSUM.md` §12.2).
- **Powell, Baker and Lawson (2008a, 2009)** lend the classes the fixture plants. Their field audit of fifty operational spreadsheets sorts the errors found into six kinds: hard-coding, reference, logic, copy and paste, data input and omission (*exhibit*). Five are planted here, one each; data input is left out, since a wrong input cannot be told from a right one without the document it came from, and the study hands over no such document (*recommendation*). Their auditing protocol (2008b) is the lineage of the find half in one respect: the workbook is handed over cold and audited without its developer; the search by kind and the findings called aloud are this protocol's own.
- **Sarkar, Gordon, Peyton Jones and Toronto (2018)** lend the nearest prior instrument: Calculation View, a textual representation beside the grid, whose user study found that the second representation improved authoring and debugging (*exhibit*; *shelf*: `web/CALLOSUM.md` §13.1). The two conditions here, the page beside Excel against Excel alone, are that comparison's shape with Frazaro's page as the second representation.
- **Nardi and Miller (1990b)** lend the risk the study guards: spreadsheets won on a formula language matched to the users' tasks and a table that is at once the model's layout and its output, each cell a visual object tied to a small program, the values updating as the data change, and a pane that becomes primary trades that away (*exhibit*; *shelf*: `web/CALLOSUM.md` §13.1). That is why the human questions are answered provisionally and the measurement that would overturn each is named.
- **The thresholds** come from older ground. One frame at 60 Hz is 16.7 ms, so a renderer that takes longer than that to answer a scroll drops frames (*derived*). A response under about 100 ms feels immediate, Miller's bound of 1968, the perceptual-processing constant Card, Robertson and Mackinlay (1991) keep at 0.1 s, and the limit Nielsen (1993) keeps for a system that feels instantaneous (*exhibit*). Section 8.1 uses both.

## 3. The fixture

### 3.1 The model

A small monthly model in three sheets, written by `tools/build_study_fixture.ps1` as raw OOXML with every cached value computed by the script, deflated, in Excel's shape, with no Office automation, in the shape of `tools/build_reflect_fixture.ps1` (*shelf*). The sheets, in tab order:

- `Inputs`: three assumptions in `B2:B4`, unit price, unit cost and fixed cost per month, each under a defined name (`Price`, `UnitCost`, `FixedCost`).
- `Sales`: the months down the rows, `A2:A13`, and the measures across the columns: `B` units (values), `C` price (`=Price`), `D` revenue (`=B*C`), `E` unit cost (`=UnitCost`), `F` cost (`=B*E`), `G` margin (`=D-F`), `H` margin as a share of revenue (`=G/D`); a `Total` row 14 summing units, revenue, cost and margin, with `H14` the year's margin share.
- `Summary`: total revenue, total cost and total margin read off `Sales` row 14, the year's fixed costs (`=FixedCost*12`), and operating income (`=B2-B3-B5`).

Months go down the rows so that each measure is one filled column, which is the shape the audit's column walks read (*shelf*: `core/src/reflect/audit.rs`): a typed-over constant or an inconsistent formula is found between two neighbours that agree, and a cell at a column's end is never judged.

### 3.2 The five plants

One defect of each class, in each of two variants, `find_a` and `find_b`. The classes are Powell, Baker and Lawson's, with Panko's name beside each; the last column says which of `frazaro audit`'s six walks finds the plant, which is what makes the audit's share of the page's advantage in condition A bounded and measurable rather than assumed: the three plants it names are known before the run, and the two it does not name are found only by reading, in the page's relations or in Excel.

| Class | Source | `find_a` | `find_b` | The audit walk |
|---|---|---|---|---|
| hard-coded constant | Powell et al. "hard-coding"; Panko "mechanical"; the plug of `SINGULARITY.md` §8.3 | `Sales!D7` holds 6500 where the column's formula gives 6800 | `Sales!F10` holds 4000 where the column's formula gives 4500 | `typed-over` |
| copy-and-paste fault | Powell et al. "copy/paste"; Panko "mechanical" | `Sales!G9` is `=D9-F8`, August's margin on July's cost, among `=D8-F8` and `=D10-F10` | `Sales!D4` is `=B3*C4`, March's revenue on February's units, among `=B3*C3` and `=B5*C5` | `inconsistent` |
| wrong reference | Powell et al. "reference"; §8.3 "a summed range stops one row short" | `Sales!D14` is `=SUM(D2:D12)`, December left out | `Sales!D14` is `=SUM(D3:D13)`, January left out | none: a range is never expanded and a column's end is never judged |
| logic error | Powell et al. "logic"; Panko "logic"; §8.3 "a sign flipped" | `Sales!H2:H13` is `=G/F`, margin over cost, in every row | `Sales!G2:G13` is `=D+F`, a sign flipped, in every row | none: a column that agrees with itself is consistent |
| omission | Powell et al. "omission"; Panko "omission" | `Sales!F14`, the cost total, was never written; `Summary!B3` reads the empty cell | `Sales!G14`, the margin total, was never written; `Summary!B4` reads the empty cell | `empty-reference` |

So the audit names three of the five in each variant, and the other two are found only by reading. In `find_b` the omission is reported twice, once from `Sales!H14`, the year's margin share, which also reads the absent `G14`, and once from `Summary!B4`; that is the treaty's rule working as written (an empty reference is a single-cell reference in a formula, to a sheet the file has, whose cell has no row), so `find_b`'s audit golden holds four lines and `find_a`'s three, and a participant who calls the omission from either line has found it once. The cached values carry every plant's consequence (*derived*, held by `scripts/study/*_relations.vla`): in `find_a` the units total is 2010, so a correct model gives revenue 80,400, cost 50,250 and operating income 12,150, while the workbook as planted shows revenue 71,700, cost 0 and operating income 53,700; in `find_b` the units total is 1540, a correct model gives revenue 77,000, cost 46,200 and operating income 16,400, and the workbook shows 71,750, 45,700 and 11,650, with a margin total of 0. The two variants are isomorphic, the same sheets, layout and five classes with different numbers, cells and symptoms, so a participant who has audited one has not seen the other's answers (*recommendation*).

### 3.3 The build keys

`build_a.xlsx` and `build_b.xlsx` are the workbooks the build half asks for (section 4.1), written by the same script with no defect, so that a participant's file is scored against them by `frazaro diff` and needs no grader: one sheet `Plan`, four headed columns, six months, a price typed in each row, a revenue formula per row, and a `Total` row with two sums.

### 3.4 The files, and what is blessed

`scripts/study/` holds the four workbooks, each fixture's relations as the door prints them (`<name>_relations.vla`) and each find fixture's audit (`<name>_audit.vla`), the door's own output saved CRLF, held by `tools/check_study_fixture.ps1`: the audit of each find fixture is exactly its findings, three lines for `find_a` and four for `find_b`, the planted rows are in the relations goldens whole, the omitted cells have no row, and with a built door the door reproduces every golden whole. Regenerate with `powershell -File tools\build_study_fixture.ps1`, then the goldens with the door, then raise the check's floors. The owner blesses the fixtures by opening each in Excel: `find_a` and `find_b` with the `Sales` sheet read against the table of section 3.2, the planted cells as named and nothing else amiss; `build_a` and `build_b` with the `Plan` sheet read against section 4.1.

*Blessed 2026-10-09:* the owner opened the four workbooks in Excel and sent every sheet as a screenshot, and each plant shows its symptom as section 3.2 names it. In `find_a`, `D7` reads 6,500, `G9` 2,500, `D14` 71,700, the `H` column 0.6 save the two rows the plants touch, `F14` is empty and `Summary!B3` 0, and operating income is 53,700. In `find_b`, `D4` reads 4,750, `F10` 4,000, `D14` 71,750, the `G` column is revenue plus cost, `G14` is empty with `H14` and `Summary!B4` at 0, and operating income is 11,650. The keys' `Plan` sheets total 890 and 35,600, and 900 and 45,000, as section 4.1 builds them (*measurement*, the owner's). Two facts beside the blessing. Excel's own error checking marks the copy-and-paste plant of each variant, `find_a`'s `G9` and `find_b`'s `D4`, with its green triangle, so in condition B the host points at one plant of the five as the audit points at three in condition A; the comparison stands as designed, and the runner notes in condition B's row whether the triangle was the lead. And the cached values the owner read are the builder's arithmetic, which `KERNEL.7`'s recalculation reproduces for every formula of the four workbooks, 176 of 176 (*measurement*, `tools/check_recalc_golden.ps1`, 2026-10-09).

## 4. The task sheet

*What a participant receives, in both conditions, worded once so that the conditions differ in the tools and in nothing else. The runner reads each half aloud or hands it over as written.*

### 4.1 Build this

*Variant `a`.* On a sheet named `Plan`, put the headers `Month`, `Units`, `Price` and `Revenue` in `A1` to `D1`. Put the months `Jan` to `Jun` in `A2` to `A7`. Put the units 120, 135, 150, 160, 155 and 170 in `B2` to `B7`. Put the price, 40, in each of `C2` to `C7`. Put a formula in each of `D2` to `D7` that multiplies that row's units by its price. In row 8 put `Total` in `A8`, a formula in `B8` that adds the six units, and a formula in `D8` that adds the six revenues. Say "done" when you are done.

*Variant `b`.* The same, with the months `Jul` to `Dec`, the units 130, 140, 150, 145, 160 and 175, and the price 50.

In condition A the participant writes in the page and takes the workbook from *Download as .xlsx*; in condition B the participant works in a blank Excel workbook. In both, the file is then opened in Excel and saved, so that it carries Excel's cached values, as `build_<variant>_<condition>.xlsx` in the folder the runner works in, outside `scripts/study/`, so that `.gitignore`'s negation never offers it for tracking.

### 4.2 Find that

Open the workbook you are given (`find_a.xlsx` or `find_b.xlsx`). It has defects. Find as many as you can. Each time you find one, say the cell and what is wrong with it, and the runner writes it down with the clock. Say "done" when you are done, or the runner stops you at fifteen minutes. In condition A the page's Reflect pane holds the same workbook, picked from disk, with its relations and its audit on offer (the pane's third reading, a diff, needs an earlier copy the task does not hand over); in condition B Excel is alone, with everything Excel itself offers. The participant is not told how many defects there are (*recommendation*: the stopping is part of what is measured, as it is in a real audit).

### 4.3 The comprehension questions

Asked after the find half, on the same workbook, answered aloud, each timed from its asking:

1. *Choose a cell.* Which cell holds the year's operating income?
2. *Explain its formula.* In words, what does that cell's formula compute?
3. *Find its inputs.* Which cells does it read directly? And which cells on the `Sales` sheet feed it through those?

The keys. Q1: `Summary!B6`, in both variants. Q2: an answer is right when it names the three things and the operation, total revenue less total cost less the year's fixed costs; naming two of the three, or an addition, is wrong. Q3, the direct part: `Summary!B2`, `B3` and `B5`; the `Sales` part: `Sales!D14` and `Sales!F14` in both variants (in `find_a` `F14` is the empty cell, and saying so counts as naming it), with `Inputs!B4` through `FixedCost` named or not named without penalty. Q3's keys are the `refers` rows of the relations goldens, and Q1's and Q2's their `cell` and `formula` rows for `Summary!A6`, `B6`, `A2`, `A3` and `A5`, which is what makes them answers and not opinions (*shelf*: `scripts/study/find_a_relations.vla`, `find_b_relations.vla`).

## 5. The protocol

### 5.1 Materials

The four workbooks under `scripts/study/`; the page, `web/index.html`, built by `powershell -File tools\build_web.ps1` (which builds the wasm core first; `-NoBuild` takes one already built under `target\`), or the hosted page of the release on the project's GitHub Pages; Excel 365; a clock; section 9 open or printed; and a door for scoring, `target\debug\frazaro.exe` or the release's.

### 5.2 Conditions, order and assignment

Condition A is the page beside Excel; condition B is Excel alone. Every participant does both halves in both conditions, one variant per condition, so nobody sees a workbook twice. Order and assignment alternate down the participants: the first takes A with variant `a` then B with variant `b`; the second B with `a` then A with `b`; the third A with `b` then B with `a`; the fourth B with `b` then A with `a`; and so round. The pilot of one takes the first row (*recommendation*).

### 5.3 What the runner says

Once, before the clock: "You will build a small sheet from a description, then look for defects in a workbook you have not seen, then answer three questions about it. Twice, with different tools. There is no help once a half begins; say 'done' when you are done." Before condition A only, one sentence per pane of the page: the input rows and what each row shows beside it, the Workbook strip's download, and the Reflect pane's picker with its three readings. Nothing is demonstrated.

### 5.4 Stop rules

Fifteen minutes for each half; two minutes for each comprehension question. A half stopped by the clock is scored as it stands and marked `stopped` in the notes.

### 5.5 Scoring

- *Build.* `frazaro diff scripts\study\build_a.xlsx build_a_A.xlsx` (the variant's key, then the participant's Excel-saved file, run from the runner's folder). Every `changed` row on sheet `Plan` is one error, with two readings the runner makes by eye: the diff compares a formula by its text, so a correct formula in another spelling (`=C2*B2` for `=B2*C2`, six additions for a `SUM`) is not an error when its value matches, and is noted; and a file with no sheet named `Plan` prints a `sheet-removed` row and no cells, which is one error for the name, after which the sheet is renamed and the diff run again. The rows the page's own build adds, a `Frazaro` sheet in one file alone and its `Frazaro.Build` name, are not errors. Time is the seconds from the hand-over to "done". *Clarified 2026-10-09, before the first run:* an error is counted where it was typed, so a `changed` row whose formula text is the same on both sides and whose value alone differs is the consequence of an error elsewhere (one wrong unit moves its row's revenue and both totals) and is not counted. The diff matches sheet names without case, so the page's `plan` is the key's `Plan`. A formula with no cached value cannot be scored, which is why the file is opened in Excel, its editing enabled, and saved before the diff; a perfect answer from the page then prints only the page's own two rows, and one from Excel prints nothing.
- *Find.* Each of the five plants named by its cell, or by a cell its symptom shows in with the cause given (naming `Summary!B3` as "the cost total is missing" finds the omission), is one found, with the clock at the call. Anything else called is one false call. A symptom named without its cause (`Summary!B6` "looks too high") is neither. Found of five, false calls, the time to each found, and the half's own time to "done" or to the stop are recorded.
- *Comprehension.* Each question right or wrong by section 4.3's key, with its seconds.

### 5.6 The pilot of one

The first runner is the owner, who built the tool, the fixture and this protocol. A pilot of one validates the instrument: that the tasks can be run as written, that the keys score without a grader, that the clock and the notes capture what section 9 asks for, and that the benchmark's page prints what section 7 says. It yields the engineering numbers of section 7 honestly, since a renderer does not know who is watching. It does not yield honest human numbers: a person who knows where the five plants are cannot be timed finding them, and a person who wrote the grammar cannot stand for someone meeting it. The human rows of section 9 filled by the pilot are therefore marked `pilot` and are read as a rehearsal, and the human questions of section 8 keep their provisional answers until a participant who is not the tool's author has run (*recommendation*).

### 5.7 What the runner does not do

Hint, answer, touch the mouse or keyboard, read the page's audit aloud, or start it. The page is explained once (5.3) and then left to the participant.

### 5.8 Predictions, written before the first run

For the pilot of one (*prediction* throughout, written 2026-10-08): the build half takes within a third as long in A as in B, with no error in either, since the writer is fluent in both; in the find half the three plants the audit names are called inside the first minute of A, and the two reading plants take as long in A as in B; all three questions are answered right in both conditions, and Q3's `Sales` part faster in A, from the relations' `refers` rows. For a participant who is not the author: the build half is slower in A on a first meeting with the grammar and the errors are refusals rather than wrong cells; the find half finds more in A inside fifteen minutes, by the three the audit names; Q2 is answered right more often where the formula is in view in words.

## 6. The measures

Three, defined once. *Time*: seconds on the runner's clock, from the hand-over of a half or the asking of a question to "done" or the answer, with the clock at each found defect besides. *Errors*: for the build half, the count of `changed` rows on `Plan` under 5.5; for the find half, found of five and false calls. *Comprehension*: right or wrong for each of the three questions by the key, with its seconds. Nothing is rated, nothing is averaged across people, and a stopped half is a stopped half.

## 7. The benchmark

Canvas or DOM is an engineering measurement, not a human one (*recommendation*). `tools/bench_view.ps1` writes a synthetic program of ten thousand sentences, half values and half formulas, the generator of `web/CALLOSUM.md` §10 (*shelf*), runs `frazaro view` over its `Output` sheet for the whole extent, a real view record of twenty thousand lines, and fills `tools/bench_view.template.html` with it, writing `tools/bench_view.html`, a page that is not the product's and never will be. The page holds two throwaway renderers of the same window record, a virtualized DOM table and a canvas drawn with the 2D context, and times each at two windows: a normal one of 40 rows by 15 columns, and a full-screen one sized from the screen, at 20 px rows and 96 px columns.

For each renderer and window the page reports: *cells with text*, the non-empty cells of the window at the top of the sheet, since this record fills two of the window's columns and the load is the text drawn, not the window's area; *render ms*, the synchronous cost of the first draw; *first paint ms*, the first animation frame's timestamp after the render call, the render having started at the top of a frame's callback phase, so that every case is measured on one footing, the first vsync after the frame that painted it; a scroll run of 120 steps of three rows from the top, 360 rows in all, with the frame time of each step from the scroll to the next animation frame, reported as p50, p95 and max, as the count of frames over 16.7 ms, and as the count over one and a half times the display's own frame interval, which the page measures over thirty idle frames before the first case and prints in its environment line; two jumps, to the middle and the end, the larger reported; and *clears 16 ms*, yes when p95 is under 16.7 ms and max under 33.4 ms, the protocol's absolute bar at 60 Hz, which the interval column lets the owner read on a faster panel. The four cases run in the order of section 9.2's table, after one discarded warm-up of each renderer at a small window. A panel draws the same text in a DOM cell and a canvas cell side by side at the machine's device pixel ratio, prints the ratio, and says whether the canvas's backing store is its CSS size times that ratio, so HiDPI crispness is checked by eye and by number, not assumed. The numbers print in a textarea as a Markdown table, with an environment line (browser, ratio, screen, whether the run was in fullscreen, the door's line), to paste into section 9.2.

To run: `powershell -File tools\bench_view.ps1` (the debug door, or `-Impl` for the release's), open `tools\bench_view.html` in the browser, which runs once on opening, not in fullscreen, with both buttons grey until it is done; then press *Fullscreen*, which enters fullscreen and runs again by itself; wait for the table, press *Copy*. *Run* runs again. The full-screen row is honest only in fullscreen, since a browser culls painting outside its viewport; the page says so, and the environment line records it. A second browser is a second row of evidence, not a requirement.

*Amended 2026-10-08, after the owner's first fullscreen run.* The first harness read a step's time off the next animation frame's timestamp, which quantizes every case that keeps up to one frame interval (all four cases read 16.5 to 16.8 ms at every percentile) and cannot tell a margin from a miss; and the sparse record fills two of the window's columns, so the full-screen case drew 71 texts a frame. The page now measures a draw's cost on the main thread, from before the call to a task queued after the frame's rendering, so that the DOM's style, layout and paint recording count as the canvas's fill calls do and a draw that keeps up reads its true cost; a dropped frame is still read off the next animation frame, later than one and a half intervals. Beside the sparse record at the protocol's two windows, a dense record, every cell of 600 rows by 52 columns filled (`tools/bench_view.ps1`, `-DenseRows` and `-DenseCols`), is drawn at full screen down a ladder of four cell sizes, 20 by 96, 16 by 80, 12 by 64 and 10 by 48 px with 13, 11, 9 and 7 px text, so that the load climbs from the screen's default to the densest legible sheet and past it. Each case reports the first draw's cost; the step cost over 120 steps of three rows (p50 and p95); the page cost over 20 steps of a whole window, between the top and the middle of the sheet, so that every page step is a full rewrite (p95); the max; the synchronous part alone; the steps dropped of 140; and *holds the frame*, yes when none was dropped and both p95s are under one interval. The DOM renderer now recycles its rows and moves by a transform, as a real viewport does, so a three-row step rewrites three rows and a page step rewrites them all, which is the fair shape of each renderer, the canvas redrawing everything on every draw. Under the table the page fits each renderer's step cost and page cost against the cells with text over the dense cases, in milliseconds per thousand cells, and prints the largest load that held the frame, the capacity at which the fitted page cost reaches one interval, and the verdict section 8.1 reads off the table. Twelve cases run, the renderers alternating, after the warm-ups and the interval; the run takes under a minute.

## 8. The decision rule, written before anything runs

*Recommendation throughout. The rule is part of the deliverable so that the numbers decide; it is changed, if at all, before the first run and by a dated entry.*

### 8.1 Canvas or DOM

1. A renderer *clears the frame* when, at the full-screen window, in a real browser in fullscreen on the owner's machine, its scroll p95 is under 16.7 ms and its max under 33.4 ms.
2. If both clear it, the choice falls to the criteria the commandments leave, and they favour canvas: there is no in-cell editing to host (CALLOSUM §7 decision 5, the formula bar is the one editor); `SD-35` makes the program the accessible workbook, so the DOM's native accessibility is not what the grid must supply (`KERNEL.17` has the pane and the bar's three lines read by a screen reader and a cell's explanation as prose, and names no reading of the grid's cells themselves); and frozen headers, lit precedents and dependents, and refusal markers (`SD-25`) are drawn things that a canvas draws in one pass and a DOM composes from layers. Canvas is then chosen *only if* its text is as crisp as the DOM's by eye on the page's panel at the machine's ratio, with the backing-store line saying yes; blurred text is a DOM.
3. If one clears it and the other does not, the one that clears it, with the crispness check still applied to canvas.
4. If neither clears it, the window is the finding, not the renderer: `KERNEL.5` draws the normal window first, and the full-screen window becomes a measurement item of its own on the KERNEL line, not minted here.
5. The chosen renderer must also paint the normal window inside 100 ms; if it does not, step 2 is read again with the other.

*Prediction (2026-10-08):* both renderers clear the frame at the normal window; at the full-screen window the DOM's p95 is the one at risk; first paint is one frame interval for both at the normal window, since each draw fits inside the frame it starts in, and the DOM's full-screen first paint is the one that may take a second frame; both are crisp at a ratio of 1 and the canvas is crisp at 1.5 and 2 because it scales its backing store.

*Amended 2026-10-08, after the owner's first fullscreen run (Chrome 154 on a 2560 by 1440 CSS screen at a ratio of 1.5, 60 Hz): every case kept every frame, 0 of 120 steps over one and a half intervals, and the strict bar of step 1 tripped three of the four cases on vsync jitter at the interval itself, since the measurement could not resolve a cost below one frame; the prediction that the DOM's full-screen p95 was the one at risk did not hold at that load of 71 texts a frame. Steps 1 to 5 above are replaced by these, read off the amended table of section 7:*

1. A renderer *holds the frame* at a load when no step or page step of its case was dropped (none followed by an animation frame later than one and a half intervals) and both its step p95 and its page p95 are under one frame interval, on the owner's machine in fullscreen.
2. The loads are the ladder of section 7: the sparse record at the protocol's two windows, and the dense record at full screen at the four cell sizes; the densest legible load is the 12 by 64 px size.
3. If both renderers hold the frame at every load up to and including the densest legible one, the choice falls to the criteria the commandments leave, which favour canvas as the old step 2 says, crispness permitting.
4. If one renderer holds a load the other does not, the one that holds it, and the viewport caps its density where the chosen renderer holds; if neither holds the densest load, the larger capacity decides, and the viewport caps its density at that capacity.
5. The slope and the capacity are reported beside the choice so that the margin is known: a renderer whose capacity is under twice the densest legible load is one a slower machine will overrun, and the viewport then defers rows during a fast scroll, which both renderers already can, since both draw only the window.

*Prediction (2026-10-08, before the dense run):* the canvas's page cost rises at under 2 ms per thousand cells with text and the DOM's at over 3; both hold the frame at the sparse loads and at the dense 20 by 96 load (about 1,800 cells on that screen); the DOM is the first to drop, at the 12 by 64 or the 10 by 48 size; the canvas holds every size but perhaps the last, where it is near the interval; the DOM's step cost, three rows recycled, stays under 2 ms at every size, so a slow drag of the thumb is cheap for both and the page cost is what separates them.

### 8.2 Which of the formula bar's three lines shows first

*Provisional answer:* the formula, for an author and an auditor, as Excel's own bar shows it, since the find half and its questions are the formula's reading, and the literature's weakness of the spreadsheet model is that comprehending and explaining a sheet are hard, its formulas being hard to understand (Hendry and Green), the computation hidden behind the values (Sarkar et al.); the value first for a reader, whose pane is closed (`SD-28`) and whose cell shows it anyway; the sentence third, reached by one key. *The measurement that would overturn it:* on the viewport, once it exists, over a fixture built from sentences, a participant who is not the author answers "explain its formula" with the bar in each order, sentence first and formula first; if the sentence-first order gives the right explanation faster, the sentence goes first for the author too. This study's Q2 times are that measurement's baseline.

### 8.3 Whether the pane opens by default for an author

*Provisional answer:* open. The author's work is the sentences, and the nearest prior study (Sarkar et al. 2018) found the textual representation improved authoring; `SD-28` keeps it closed for a reader, and the explain act opens it. *The measurement that would overturn it:* the build half, on a participant who is not the author, in condition A against condition B on the isomorphic specs; if building by sentences costs more time *and* more errors than building by hand, the pane is not where an author's speed lives, and it opens on demand for authors too, with the grid first.

## 9. The result table

*Blank. Filled only by a dated entry appended below it, never by editing a cell; copies of the filled tables of 9.1 to 9.3 go to `web/CALLOSUM.md`'s slot of 2026-10-08 as a dated measurement, with the machine, the browser and the door named.*

### 9.1 The study

| participant | order | condition | variant | build time (s) | build errors | find time (s) | found of 5 | time to each found (s) | false calls | Q1 | Q2 | Q3 | notes |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 (pilot) | 1 | A | a | | | | | | | | | | pilot |
| 1 (pilot) | 2 | B | b | | | | | | | | | | pilot |

*Entry 2026-10-09: not run; the item closed without these rows, by section 10's entry of the same date.*

### 9.2 The benchmark

*The header amended 2026-10-08 with section 7's amendment; the earlier header was never filled.*

| renderer | record | cell px | window (rows x cols) | cells with text | first draw ms | step p50 ms | step p95 ms | page p95 ms | max ms | sync p50 ms | dropped of N | holds the frame |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| DOM (virtualized) | sparse | 20 x 96 | normal (40 x 15) | | | | | | | | | |
| canvas | sparse | 20 x 96 | normal (40 x 15) | | | | | | | | | |
| DOM (virtualized) | sparse | 20 x 96 | full-screen | | | | | | | | | |
| canvas | sparse | 20 x 96 | full-screen | | | | | | | | | |
| DOM (virtualized) | dense | 20 x 96 | full-screen | | | | | | | | | |
| canvas | dense | 20 x 96 | full-screen | | | | | | | | | |
| DOM (virtualized) | dense | 16 x 80 | full-screen | | | | | | | | | |
| canvas | dense | 16 x 80 | full-screen | | | | | | | | | |
| DOM (virtualized) | dense | 12 x 64 | full-screen | | | | | | | | | |
| canvas | dense | 12 x 64 | full-screen | | | | | | | | | |
| DOM (virtualized) | dense | 10 x 48 | full-screen | | | | | | | | | |
| canvas | dense | 10 x 48 | full-screen | | | | | | | | | |

Environment line, with the frame interval: (blank until pasted). The page's summary lines, the slope, the largest load held, the capacity and the verdict for each renderer: (blank). Crispness by eye at the machine's ratio: (blank).

*Entry 2026-10-08, the owner's run in fullscreen: Google Chrome 154.0.8037.93 on Windows, a 2560 by 1440 CSS screen at a ratio of 1.5, 60 Hz (a 16.7 ms interval), the native debug door 0.8.0; the sparse record 20,006 lines, the dense record 62,406 lines (600 rows by 52 columns). The twelve rows and the summary as the page printed them:*

| renderer | record | cell px | window (rows x cols) | cells with text | first draw ms | step p50 ms | step p95 ms | page p95 ms | max ms | sync p50 ms | dropped of N | holds the frame |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| DOM (virtualized) | sparse | 20 x 96 | normal (40 x 15) | 40 | 3.2 | 1.6 | 1.8 | 1.9 | 2.4 | 0.1 | 0 of 140 | yes |
| canvas | sparse | 20 x 96 | normal (40 x 15) | 40 | 1.4 | 0.4 | 0.9 | 0.7 | 1.0 | 0.2 | 0 of 140 | yes |
| DOM (virtualized) | sparse | 20 x 96 | full-screen (71 x 26) | 71 | 7.5 | 3.6 | 4.1 | 3.8 | 4.4 | 0.1 | 0 of 140 | yes |
| canvas | sparse | 20 x 96 | full-screen (71 x 26) | 71 | 1.7 | 0.6 | 1.1 | 1.4 | 1.4 | 0.4 | 0 of 140 | yes |
| DOM (virtualized) | dense | 20 x 96 | full-screen (71 x 26) | 1846 | 19.1 | 6.9 | 7.4 | 14.3 | 15.4 | 0.2 | 0 of 140 | yes |
| canvas | dense | 20 x 96 | full-screen (71 x 26) | 1846 | 6.6 | 3.0 | 3.7 | 5.5 | 5.5 | 1.5 | 0 of 140 | yes |
| DOM (virtualized) | dense | 16 x 80 | full-screen (89 x 31) | 2759 | 23.1 | 10.2 | 10.6 | 24.9 | 25.6 | 0.2 | 0 of 140 | no |
| canvas | dense | 16 x 80 | full-screen (89 x 31) | 2759 | 8.9 | 4.4 | 5.0 | 7.6 | 8.1 | 2.3 | 0 of 140 | yes |
| DOM (virtualized) | dense | 12 x 64 | full-screen (119 x 39) | 4641 | 43.3 | 17.3 | 17.7 | 39.5 | 41.9 | 0.2 | 8 of 140 | no |
| canvas | dense | 12 x 64 | full-screen (119 x 39) | 4641 | 16.3 | 7.5 | 8.2 | 13.1 | 13.6 | 3.9 | 0 of 140 | yes |
| DOM (virtualized) | dense | 10 x 48 | full-screen (143 x 52) | 7436 | 68.3 | 28.8 | 29.4 | 61.1 | 69.2 | 0.3 | 41 of 140 | no |
| canvas | dense | 10 x 48 | full-screen (143 x 52) | 7436 | 25.5 | 13.0 | 13.9 | 22.5 | 23.7 | 6.1 | 0 of 140 | no |

DOM (virtualized): step cost 3.92 ms per 1,000 cells (-0.6 ms fixed); page cost 8.19 ms per 1,000 cells (0.8 ms fixed); largest dense load that held the frame: 1,846 cells; capacity at one frame interval, by the page cost: about 1,942 cells.
canvas: step cost 1.79 ms per 1,000 cells (-0.5 ms fixed); page cost 3.07 ms per 1,000 cells (-0.6 ms fixed); largest dense load that held the frame: 4,641 cells; capacity at one frame interval, by the page cost: about 5,644 cells.
Verdict: neither holds the frame at 7,436 cells with text; the DOM held 1,846 and canvas 4,641; the larger capacity decides (step 4), and the viewport caps its density there.

*Read against section 8.1 as amended (*measurement*, the owner's): the DOM holds the frame at the sparse loads and at the dense default size, 1,846 cells, with 2.4 ms of its page cost to spare, and fails at 16 by 80, 2,759 cells, page p95 24.9 ms, and at every size after; canvas holds through 12 by 64, the densest legible size, 4,641 cells at a page p95 of 13.1 ms, and fails only at 10 by 48, 7,436 cells, 22.5 ms, with no frame dropped. Step 4 applies: canvas holds two loads the DOM does not, and its capacity is the larger. The slopes say why: a full rewrite rises at 8.2 ms per thousand cells with text for the DOM and 3.1 for canvas, a three-row step at 3.9 and 1.8; the DOM's step cost is the browser's and not the script's, whose own part is 0.3 ms at every size, so in Chrome a change inside a grid of elements costs in proportion to the elements on screen whatever the renderer recycles, where the canvas pays only for what it draws (*derived* from the sync and main columns; the mechanism is the sitting's reading). The same run not in fullscreen gave the same shape and the same decision. On a screen of more CSS pixels than this one the DOM fails at the default size, since 1,846 cells is already within 2.4 ms of its limit. Crispness at a ratio of 1.5: the page's backing-store line said yes; the eye's verdict is the owner's to add.*

*Entry 2026-10-09, the crispness clause of section 8.1 closed: the owner judged the canvas text crisp by eye at a ratio of 1.5 on 2026-10-08, during `KERNEL.5`'s hand test in Alonzo, on that page's HiDPI panel, the same text in a DOM cell and in the viewport's canvas side by side (*measurement*, the owner's; recorded in Alonzo's `tools/check_render_floors.ps1` header and its `REARVIEW.md`). Canvas stands by section 8.1 with nothing owed.*

### 9.3 The decisions

| decision | answer | by rule | date |
|---|---|---|---|
| canvas or DOM | | 8.1 | |
| the bar's first line | provisional: the formula | 8.2 | |
| the pane's default for an author | provisional: open | 8.3 | |

*Entry 2026-10-08:* canvas or DOM is decided: **canvas**, by section 8.1 step 4 as amended, on the entry of 9.2; written into `KERNEL.5`'s roadmap entry the same day. The bar's first line and the pane's default keep their provisional answers until the study's human half has run.

## 10. What "run" means

`KERNEL.3` stays open until the owner has: run both halves in both conditions and filled the two pilot rows of 9.1; run the benchmark in a real browser in fullscreen and pasted its table into 9.2; applied section 8 and filled 9.3 in a dated entry; copied 9.1 to 9.3 to `web/CALLOSUM.md`'s slot; and written `KERNEL.5`'s first shape, the renderer chosen, into its roadmap entry. The item closes on that entry, not on this file.

*Entry 2026-10-09, at the owner's word: `KERNEL.3` closes on what it delivered, the instrument and the one decision a measurement could take now.* The renderer question is answered, canvas, by the entry of 9.2, and `KERNEL.5` is built on it in Alonzo. The study's human half cannot change either of the other two decisions while its only participant is the person who built the tool: section 5.6 says so, and sections 8.2 and 8.3 keep their provisional answers whatever such a pilot measures, so a run now would test the instructions and decide nothing. The human half therefore runs the day the formula bar or the pane is next scoped, with a participant who is not the author, before either is built; `SD-29` requires a study then, and this instrument is that study, so no item is minted for it. Until then the formula bar shows the formula first and the pane opens by default for an author, provisionally, and the result tables of 9.1 stay blank.

## 11. Sources

- Hendry, D. G. and Green, T. R. G., "Creating, comprehending and explaining spreadsheets: a cognitive interpretation of what discretionary users think of the spreadsheet model", *International Journal of Human-Computer Studies* 40(6), 1994, pp. 1033-1065.
- Panko, R. R. and Halverson, R. P., "Spreadsheets on trial: a survey of research on spreadsheet risks", *Proceedings of the 29th Hawaii International Conference on System Sciences*, 1996.
- Panko, R. R., "What we know about spreadsheet errors", *Journal of End User Computing* 10(2), 1998, pp. 15-21; revised 2008.
- Powell, S. G., Baker, K. R. and Lawson, B., "A critical review of the literature on spreadsheet errors", *Decision Support Systems* 46(1), 2008a, pp. 128-138.
- Powell, S. G., Baker, K. R. and Lawson, B., "An auditing protocol for spreadsheet models", *Information & Management* 45(5), 2008b, pp. 312-320.
- Powell, S. G., Baker, K. R. and Lawson, B., "Errors in operational spreadsheets", *Journal of Organizational and End User Computing* 21(3), 2009, pp. 24-36.
- Sarkar, A., Gordon, A. D., Peyton Jones, S. and Toronto, N., "Calculation View: multiple-representation editing in spreadsheets", *IEEE Symposium on Visual Languages and Human-Centric Computing*, 2018.
- Nardi, B. A. and Miller, J. R., "An ethnographic study of distributed problem solving in spreadsheet development", *Proceedings of CSCW*, 1990a, pp. 197-208.
- Nardi, B. A. and Miller, J. R., "The spreadsheet interface: a basis for end user programming", in D. Diaper et al. (eds.), *Human-Computer Interaction: INTERACT '90*, North-Holland, 1990b, pp. 977-983.
- Miller, R. B., "Response time in man-computer conversational transactions", *AFIPS Fall Joint Computer Conference*, 1968. Card, S. K., Robertson, G. G. and Mackinlay, J. D., "The Information Visualizer, an information workspace", *CHI*, 1991. Nielsen, J., *Usability Engineering*, 1993, chapter 5.
- In this repository: `web/CALLOSUM.md` §7, §10, §13; `docs/SINGULARITY.md` §8.3 (the explanation schemas); `core/src/reflect/audit.rs` (the six walks); `conformance/README.md`, oracles 10 and 11; `scripts/pareto.txt`.
