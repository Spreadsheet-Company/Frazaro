# Frazaro

*A phrasebook for your workbook. You write the steps of a procedure as
English sentences. Frazaro checks every sentence, refuses the ones it
cannot read, in words, and runs the rest inside Excel.*

## What is a beta?

*With apologies to Kierkegaard, whose clown had the same trouble.*

> It happened one evening that a sentence broke in the beta-stage of an add-in.\
> Frazaro appeared & announced, in words: the gap, the row, and what to write instead.\
> The users took this for a feature and applauded.\
> Frazaro repeated itself, verbatim & verified, on the same row.\
> They applauded louder and asked, "When 1.0?"\
> I think that is precisely how the internet will come to an end:\
> amid general applause from users who believe they are shareholders.

Less lyrically: this is a known-unfinished program, released unfinished
on purpose, because the only way to learn which sentences real people
reach for is to let real people reach for them. The gaps are counted, in
order, in [docs/BETA_ROADMAP.md](docs/BETA_ROADMAP.md). Before
reporting a missing feature, check whether it is already there. If it
is, the complaint is heard and queued. If it is not, that is a genuinely
useful report and exactly why the beta exists.

---

## What is Frazaro?

Everyone who lives in spreadsheets has procedures: month-end checklists,
report formatting, data cleanup, the Friday reimbursement run. They can
describe each one in a breath and cannot automate it without learning
VBA. Frazaro is the missing middle: a small, checked English that Excel
translates into VBA macros.

Here is a real standard operating procedure, kept in Word (or.txt) the way
people already keep them. Two lines were added to it, and Frazaro reads
only what sits between them:

```text
Weekly Expense Reimbursement
Owner: Accounts Payable. Performed every Friday afternoon.

PROCEDURE
1. Open the week's expense report.
2. Add up every claim that has a receipt, and write the total beside the report.
3. Send the total to payroll before 3pm.

<Frazaro week="Week 39" (steps 1 and 2, automated by Accounts Payable)>
Work on sheet "Expenses".
Set receipted-total to sum of range D2:D26 where range E2:E26 matches "Yes".
Put "Reimbursement total, {week}" into cell H10.
Put receipted-total into cell I10.
Format cell I10 as dollars.
Make range H10:I10 bold.
Show "{week} reimbursement total: " joined with receipted-total.
</Frazaro>
```

That is not pseudocode. It is the first of the
[sample procedures](examples/), and it runs. The title, the purpose and
the manual step stay exactly as a person wrote them. `{week}` is filled
in from the tag, so running it for another week means changing `Week 39`
in one place.

Frazaro is not a prompt box. There is no model and no guessing. Every
sentence either matches a published sentence shape or is refused before
anything runs, and a sentence that matches means exactly one thing.

## Try it in two minutes

1. Download
   [Frazaro_English.xlam](https://github.com/Spreadsheet-Company/Frazaro/releases/latest/download/Frazaro_English.xlam)
   and open it. Excel asks once whether to trust the publisher; then a
   **Frazaro** tab appears on the ribbon. (Press **Register for
   Auto-Load** if you want it there every time; the [Install](#install)
   section has the longer story.)
2. Open [`examples/Frazaro Sample Data.xlsx`](examples/).
3. Press **Load Instructions** and pick
   `00 Weekly Expense Reimbursement.docx` (or the `.txt`, if this
   computer has no Word). The procedure appears on a new *Frazaro*
   sheet, one line per row.
4. Press **Validate Instructions**. Every line gets a green OK, or a
   note in plain words saying what to change. Nothing has run yet.
5. Press **Interpret and Run**, then look at the *Expenses* sheet.
6. Changed your mind? **Undo Last Run** puts the workbook back.

Eight more follow, easiest first: a sales cleanup in five sentences,
a weekly summary, an expense audit with a loop, receivables aging, pivot
tables, a month-end close, a purchase order, and a staffing policy you
can question. All of them run on the one sample workbook.
[`examples/README.md`](examples/README.md) walks through them.

---

## The idea

> The language in which a person thinks (source thought) should be the
> same language in which a person works (source code).

The last few years added a new way to automate a spreadsheet, the AI
prompt box, which accepts anything and is confidently wrong just often
enough to matter. Frazaro takes the opposite bet:

**A small, checked English is better than an unlimited, guessed one.**

- **Checked before anything runs.** Every sentence either parses
  against a published grammar or is refused, on its own row, in words.
  There is no "it did something, hopefully what you meant."
- **One sentence, one meaning.** Sentence shapes are tried in a fixed
  order and the first match wins, with no second reading. Because the
  shapes are finite, the loader can *prove* two of them do not overlap,
  and refuses a phrasebook where they do.
- **Refusals teach.** A refused sentence is told what was understood,
  where the reading stopped, and what to write instead, in the register
  a colleague would use. A real one, from the Datalog engine:

  > fact 'parent' uses 'X', which looks like a variable (it starts
  > with a capital letter) - facts must be fully specific; did you
  > mean to write a rule instead?

---

## What happens to your sentences

**1. Loading.** **Load Instructions** takes a Word document or a text
file. Word documents are opened in a hidden copy of Word with macros
switched off, and poured onto the *Frazaro* sheet one paragraph per row.
You can also type straight into the sheet, or press **New Frazaro** for
an empty one.

- `<Frazaro>` and `</Frazaro>`, each on a line of its own: only the
  lines between them are read. A document with no tag is read from top
  to bottom. A document can have several tagged sections.
- `<Frazaro week="Week 39">`: a value the section's sentences can use
  as `{week}`.
- `<Frazaro (a note for people) 0.7.0>`: anything in parentheses is a
  note; a bare version says the section needs that Frazaro or later.
  A bare `espanol` says it is written for the Spanish edition.
- A line starting with `#` is a note for people. Frazaro skips it.
- A line ending in a colon opens a block; the indented lines under it
  belong to it; a blank line, or `Done.`, closes it.

**2. Validating.** **Validate Instructions** reads each row word by
word, drops filler such as "the" and "please", and tries its sentence
shapes in order. Each shape has typed holes: a hole marked as a cell
must look like a cell, a hole marked as an expression is read by a small
built-in arithmetic grammar. A row that fits turns green. A row that
does not gets a note beside it in column C. Definitions are checked
against their uses, so a misspelled step name is caught here, not in a
crash later. Nothing has run.

**3. Translating.** Each accepted sentence is rewritten into VLA, a
plain-text middle language, by copying the sentence's pieces into the
shape's template. It is mechanical, like filling in a form. Nothing
appears in the VLA that the sentence did not say.

```text
English   Make cell A1 bold.
VLA       (make-bold (range "a1"))
VBA       range("a1").font.bold = True

English   Put today into cell D1.
VLA       (set! (range "d1") (date))
VBA       range("d1") = date()
```

**4. Running, one of two ways.** A snapshot of the workbook is taken
first, either way.

- **Interpret and Run**, the default. Frazaro walks the VLA one form at
  a time and performs each action directly on your workbook. No code is
  written into your file, and none of Excel's macro-trust settings are
  needed beyond running the add-in itself. If a sentence fails mid-run,
  the stop names the sentence and its row, and puts the sheets back.
- **Compile and Run.** Frazaro turns the VLA into a readable VBA module,
  every line tagged with the sentence it came from, and runs that. This
  path needs Excel's setting for trusting access to the VBA project, and
  refuses in words if the setting is off. It exists for the auditor who
  wants the artifact: **Show me the VBA** displays it, and
  **Translate File to VBA** writes it to a file.

**Interpret and Trace** and **Compile and Trace** do the same and show a
step-by-step transcript you can copy.

**5. Afterwards.** **Undo Last Run** restores the snapshot, including
sheets the run deleted, which Excel itself cannot do. A sentence Frazaro
could not read is logged on your machine; **Copy Diagnostic Report**
turns that log into a report you can paste into an email. That log is
the raw material for new sentences. **What can I say?** lists every
sentence shape currently loaded.

The two ways of running are held to agreement by a test corpus: the same
programs are compiled and interpreted, and the two results must match,
or a test fails. More on that [under the hood](#under-the-hood).

---

## What you can say

Every sentence below is one the grammar proves it accepts, taken from
the phrasebook's own proofs, the sample procedures, or the regression
corpus. References are bare, exactly as formulas write them: `cell B2`,
`range A1:C50`, `column C`, `sheet Data`. Quotes are for text, and for a
name with a space in it.

### Cells and values

`Put` places a value somewhere. `Set` gives a value a name you can use
in any later sentence. Names are one word, with hyphens if you like.

```text
Put 5 into cell B2.
Put today into cell D1.
Put formula "=B2*2" into cell B3.
Put 0 into range A1:B5.
Put 5 into cell B2 of sheet Data.
Put counter times 10 into column E row counter.
Set last-row to last filled row of column A.
Set remote to cell B2 of sheet Data.
Set difference to debits minus credits.
Set price to total rounded to 2 decimals.
Set msg to "total is " joined with total.
Set budget to 1,000,000.
Increase total by 5.
Grow total by 10%.
Copy range A1:D10 to range F1:I10.
Clear everything from A1:D50.
Remember range B2:B4 as results.
```

### Sums, counts and lookups

```text
Set grand to sum of range B2:B3.
Set average-order to average of revenues.
Set biggest to largest of results.
Set west-total to sum of range B2:B90 where range A2:A90 matches "West".
Set west-count to count of range A2:A90 matching "West".
Set price to lookup of part-code in range A2:C90 column 3.
Set found-row to row of "Widget" in column A.
```

### Making it look right

```text
Make range A1:C1 bold.
Set font size of cell A1 to 18.
Set fill-color of range A4:C4 to "#1F3A5F".
Set font-color of range A4:C4 to "#FFFFFF".
Format range B5:B10 as accounting in dollars.
Format range C5:C10 as percent with 1 decimal.
Format cell B2 as a long date.
Add a bottom border to range B5:D5.
Add a border around range B2:D4.
Band every other row of A5:C9 "#EEF3F8".
Make row 1 a header row.
Center range A1:D1.
Wrap text in range A1:A10.
Set width of column A to 18.
Fit all columns.
Freeze top row.
```

Colours can be named once and used everywhere:

```text
Define hot-pink as "#FF69B4".
Set font-color of cell A6 to hot-pink.
```

### Rows, columns, sheets and files

```text
Delete row 5.
Insert 3 rows at row 5.
Delete blank rows in A1:D50.
Hide column C.
Move column B before column D.
Remove duplicates from range A1:C50 by column 2.
Replace "N/A" with 0 in column C.
Work on sheet "Sales".
Add sheet called "Report".
Copy range A1:D10 to sheet Archive.
Protect sheet Data with password "abc".
Save this workbook.
Save a copy as backup-path.
Export this sheet as pdf report-path.
```

### Sorting, filtering, tables and pivots

```text
Sort range A1:C50 by column B descending then by column C ascending with a header row.
Filter range A1:D50 to show rows where column D is greater than 100.
Copy only the visible cells of range A1:D50 to F1.
Remove the filters.
Turn A1:D10 into a table called Sales.
Add a row to table Sales.
Make a pivot table from A1:G41 at J3 called RegionPivot.
Add rows of Region to pivot RegionPivot.
Add columns of Channel to pivot RegionPivot.
Add Revenue to pivot RegionPivot as a sum.
Show pivot RegionPivot in tabular form.
Sort Region in pivot RegionPivot descending by Revenue.
```

Where Excel offers two readings, the sentence names the choice:
`with a header row` or `without a header row`, `after any space` or
`after any non-letter`. A bare verb is not accepted when it could mean
either.

### Text

```text
Make range A2:A50 upper case.
Capitalize each word in range A2:A50 after any space.
Remove extra spaces from column A.
Remove non-printing characters from range A2:A50.
Set part to the text before "-" in code.
Set extension to the text after the last "." in file-name.
Set code to id padded on the left with "0" to 5 characters.
Set names to range A2:A9 as one list separated by "; ".
Set tidy to raw-name with extra spaces removed.
```

### Deciding

A one-line `If` ends with a comma and the sentence to run. A block `If`
ends with a colon, and a blank line separates the branches.

```text
If grand is at least 45, make cell C4 yellow.
If receipt is "No", put "Missing receipt" into column F row r.
If cell B2 is empty, show "blank".
If code starts with "AB", show "ab".

If grand is greater than 100:
  Set verdict to "huge".

Otherwise, if grand is greater than 40:
  Set verdict to "solid".

Otherwise:
  Set verdict to "small".

When region is "North":
  Set region-label to "cold".

When it is "South" or "East":
  Set region-label to "warm".

Otherwise:
  Set region-label to "unknown".
```

### Repeating

```text
Repeat 5 times:
  Increase total by counter.

Count r from 2 to last-row:
  Set amount to cell in column D row r.
  If amount is greater than 500, put "Needs manager approval" into column F row r.

Count stripe-row from 1 to f-last step 2:
  Make cell in column F row stripe-row bold.

While countdown is greater than 0:
  Log "countdown " joined with countdown.
  Decrease countdown by 1.

Repeat until fuel is 0:
  Decrease fuel by 1.
  Increase until-count by 1.

Count search-row from 1 to 10:
  If search-row is 3, stop the loop.
```

### Lists and lookups

```text
Create a list called found-items.
Repeat 4 times:
  If counter is greater than 2, append counter times 100 to found-items.

For each f in found-items, log "found " joined with f.
Set list-count to count of found-items.

Create a lookup called prices.
Store 100 at key "ax-7" in prices.
Store 250 under "bx-2" in prices.
Set ax-price to prices for "ax-7".
For each pair in prices:
  If value of pair is greater than 200, set pair-trace to pair-trace joined with key of pair.
```

### Steps of your own

Say a calculation once, then use it by name. A step can take
parameters, with defaults, and hand a value back.

```text
To landed-cost of amount:
  Give back amount times 1.05.

Set landed-total to landed-cost of order-total.

To get commission using sale of 1000 and rate of 5%:
  Give back sale times rate.

Set full-commission to commission using sale of 2000 and rate of 10%.
Set default-commission to get commission using sale of 600.

To tidy-up:
  Fit column A.
  Fit column B.
  Set font-color of cell A1 to hot-pink.

Tidy-up.
```

### Talking to the person running it

`Show` opens a message box. `Log` writes to a quiet log for whoever
maintains the procedure. `Ask` asks.

```text
Show total.
If total is greater than 5, show "big".
If unreconciled is greater than 0, show "The close summary is ready, but " joined with unreconciled joined with " accounts still need reconciling. The ledger stays unlocked until they are done".
Log "grand is " joined with grand.
Ask "Your name?" and put answer into who.
Put "working" in status bar.
```

### When something might go wrong

```text
Try:
  Go to sheet Nowhere-Land.
  Set rescue to "unreachable".

If that fails:
  Log "the problem was " joined with the problem.
```

Sample 05 uses this to delete last week's pivot tables, which do not
exist the first time it runs.

---

## Asking questions of your tables

The same engine ships **worksheet functions** for querying and logic
over your actual Excel Tables, spilled as dynamic arrays with headers.
No code is injected anywhere.

```text
=SQL("SELECT Name, Salary FROM staff WHERE Salary > 80000", Staff)
```

Real SQL text, a frozen SQLite-leaning subset. Everything outside it
refuses by name rather than guessing.

```text
=DATALOG("(fact (parent tom bob)) (fact (parent bob liz))
          (rule (ancestor X Y) (parent X Y))
          (rule (ancestor X Z) (parent X Y) (ancestor Y Z))
          (query ancestor)")
```

Recursive questions: org charts, bills of materials, anything shaped
like "who is under whom", which plain formulas cannot express. Atoms can
address table columns by header name, so `(staffing (name X) (salary
S))` reads the columns your table actually has.

```text
=PROLOG("(rule (can-cover Shift Who)
           (shifts (shift Shift) (needs Cert) (minlevel Min))
           (staff (name Who) (cert Cert) (level Level))
           (>= Level Min)
           (not (leave (name Who) (shift Shift))))
         (query (can-cover Shift Who))", Shifts, Staff, Leave)
```

Full unification with backtracking, so the question can be shaped like
the policy. Three ordinary tables become facts, and one rule says who
can cover a shift in the words a supervisor would use. The result spills
as a roster of every legal pairing. Negation, arithmetic, comparisons,
`findall`, `between`, cut and list operations are all in, and an
infinite rule is stopped by a step ceiling and refused by name.

**You do not have to write the parentheses.** The policy can be written
as sentences, which write the rules into cells for you, and the
questions as sentences too. This is sample 08:

```text
Write in cell M2 that a person can-cover a shift if Shifts lists the shift as Shift, the cert as Needs, and the min as MinLevel, and Staff lists the person as Name, the cert as Cert, and the level as Level, and the level is at least the min, and not Leave lists the person as Name and the shift as Shift.
Write in cell M3 that a shift is listed if Shifts lists the shift as Shift.
Write in cell M4 that a shift is covered if the person can-cover the shift.

Show in cell A4 who can-cover "Night" by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
Show in cell F4 which shift that is listed is not covered by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
Show in cell H4 whether every shift that is listed is covered by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
```

Every answer is live: change a table, or edit a rule in column M, and
the answers change with it. Simpler questions have simpler sentences:

```text
Show every Name in Staffing with Salary over 80000 as Rich in cell E2.
Show everyone who reports to "Alice" directly or not in Reports as team in cell E2.
```

**`=OPTIMIZE()`** is the newest engine and is partly built. It takes the
same rules plus *choices* (which rows to pick from a pool) and
*constraints* (`forbid`, `require`), and finds the first assignment that
breaks no rule, in the tables' own order. A violation names the
constraint and the rows. A formula that is too large for a cell can be
run again as a command, `Optimize cell C1.`, with no time limit. Counting
constraints and a best-answer objective are next; until they land the
engine refuses them by name. The design and the measurements behind it
are in [docs/OPTIMIZATION.md](docs/OPTIMIZATION.md).

---

## Phrasebooks: where the sentences come from

English is not the product. English is the **first phrasebook**.
(*Frazaro* is Esperanto for "phrasebook". The name is the roadmap.)

The grammar is data. A sentence shape is a **rule** in a phrasebook text
file: a pattern with typed holes, a template that says what VLA to
produce, and a proof. This is the rule behind `Make cell A1 bold.`, as
it appears in [english.vla](scripts/polyglotta/english.vla):

```text
(english-vla
    "make cell {r:cell} {d:bold|italic}"
    (make-{d} (range {r})))
(test-success
    "Make cell A1 bold."
    (make-bold (range "a1")))
```

`{r:cell}` is a hole that must look like a cell. `{d:bold|italic}`
accepts either word and carries the one it matched into the template.
Patterns can also hold optional words and alternative spellings, so one
rule covers `into` and `in`, or `center` and `centered`.

**A rule ships only with its proofs.** The `test-success` line says
exactly what the sentence must translate to; a `test-fail` line pins a
refusal and its wording. The loader runs every proof each time a
phrasebook loads, and a phrasebook that fails its own proofs does not
load. The loader also expands every alternative spelling of every rule
and refuses a set where two rules share a shape, unless the later one
says `override:` and that override replaces exactly one earlier rule.

**Phrasebooks layer.** The base phrasebook is embedded in the add-in. A
phrasebook of your own layers on top: load it with **Load Phrasebook**,
or place it beside your workbook, where it is picked up after a one-time
approval on that computer. Your file is yours to keep private or share
([PHRASEBOOK-TERMS.md](PHRASEBOOK-TERMS.md)). A rule marked `raw`
carries literal VBA and triggers a consent dialog naming the file before
it loads; declining refuses the whole phrasebook.

**Other languages ride the same machinery.** The Spanish edition is a
phrasebook layered on the English one, and the repository carries
demonstration phrasebooks for Esperanto, French, German, Danish, Latin,
pirate, and one alien whose every word is a glyph, as proof of the seam
rather than finished translations:

```text
Pon hoy en la celda D1.              espanol.vla
Metu hodiau en la chelon D1.         esperanto.vla
Pone hodie in cellula D1.            latin.vla
Mark today upon the cell D1, savvy.  pirate.vla
```

All of them produce the same VLA, so each inherits both ways of
running, Undo, tracing and the tests for free.

**How a new sentence is born.** Sentences come from real transcripts,
never from imagination: no grammar section is scheduled without a
sentence someone actually reached for, and the unknown-sentence log
behind **Copy Diagnostic Report** is where those arrive. Every verb is
auditioned against the person who will type it. A candidate is written
as a rule with its proof, the whole test corpus is regenerated and
diffed, and both ways of running are exercised live. Once shipped, a
spelling keeps its meaning forever: the release a sentence first worked
in is recorded in [docs/GRAMMAR_SINCE.md](docs/GRAMMAR_SINCE.md), and a
sentence is only ever retired through a refusal, never silently changed.

---

## Under the hood

*For developers. A business user can skip to the
[questions](#questions-people-ask).*

```text
   Word / text / the sheet
            |
   sentence engine ............ src/VLA_SentenceEngine.bas, src/VLA_English.bas
   first match wins, typed slots, phrasebook rules from scripts/polyglotta/*.vla
            |
           VLA ................. s-expression VBA, 1:1, the one statement of meaning
   prelude macros from scripts/prelude.vla, one head table for every core form
        /       \
  interpreter    compiler ....... src/VLA_Interpreter.bas | src/VLA.bas
  walks forms    emits VBA text, injects src/VLA_Runtime.bas and the module
  on the live    (needs VBA-project trust)
  workbook
        \       /
   the same golden corpus, both ways, must agree
```

**VLA** is "VBA wearing parentheses": every VBA statement has exactly
one spelling in VLA, and the project refuses on principle to make VLA
any nicer than VBA. No expression-if, no closures, no smoothing. That
refusal keeps it a spelling rather than a compiler one person must
maintain alone forever ([docs/LESSONS.md](docs/LESSONS.md), section I).
The standard library, [prelude.vla](scripts/prelude.vla), is written in
VLA as template macros: pure substitution, no gensym, inspectable with
one command.

**Why a middle language at all**, when the engine is itself VBA:

- Two runners need one statement of meaning. If English went straight
  to VBA text, the interpreter would need its own second reading of
  English, and two readings means two meanings.
- Templates stay one line. A rule produces a form, not VBA syntax with
  declarations, error handling, step numbering and quote escaping.
- Forms nest. A hole's value is a form spliced into a form, which
  composes; splicing text strings invites precedence and quoting
  accidents.
- Every language shares the middle, so a new phrasebook costs a text
  file, not a backend.
- VLA is text a reviewer can read, a linter can check and a test can
  compare byte for byte. A row of the *Frazaro* sheet that begins with
  `(` is accepted as VLA directly, and **Open CLI** gives you a console
  that takes either a sentence or a form.
- The ground may move. VBA is a backend, not the language (standing
  decision SD-1), so the interpreter is the default runtime, an Excel
  `LAMBDA` target already exists, and any future host must match the
  same goldens rather than invent its own reading (SD-18).

**What holds it together.** The regression corpus
[scripts/instructions.txt](scripts/instructions.txt) is translated and
compiled into golden files on every change; an empty diff is the
behaviour-preservation witness, and any change is a contract change that
must be read. The same corpus runs through both backends against a live
workbook and the resulting sheets are compared. Self-tests pin every bug
ever fixed, with its incident cited in place.

| | Today |
|---|---|
| English phrasebook | 220 rules, 192 macros, 404 proof sentences, 11 refusal proofs |
| Regression corpus | about 700 lines of program; 687 numbered steps and 645 sentence tags in the compiled VBA |
| Self-tests, 2026-09-26 | 1493 pure, 202 against a live host, 2299 across the query and logic engines, and 283 corpus checks passing on each backend |

**Standing decisions**, the ones a contributor meets first
([docs/BETA_REARVIEW.md](docs/BETA_REARVIEW.md) has the register):
VBA is a backend, not the language (SD-1); a shipped spelling keeps its
meaning (SD-4); every backend supports or explicitly refuses every core
form (SD-5); no grammar section without a real sentence that needs it
(SD-7); no outbound network call, ever (SD-13); the phrasebook's pattern
language is the only sentence grammar, first match wins, no backtracking
(SD-16); VBA remains the reference for any second host (SD-18).

**Where to read next.** [docs/README.md](docs/README.md) indexes the
shelf. The shortest path:

- [docs/TESTING.md](docs/TESTING.md), the six verification passes a
  change must survive, and how to build the lab bench from a clone.
- [docs/REBUILD.md](docs/REBUILD.md) and [docs/CUTS.md](docs/CUTS.md),
  the module topology, what VBA forces on it, and where the licence
  boundaries fall.
- [docs/INTRINSICS.md](docs/INTRINSICS.md), the VBA behaviours a port
  of the translator would have to reproduce exactly.
- [docs/THREAT_MODEL.md](docs/THREAT_MODEL.md), what a program may
  reach, what a phrasebook may reach, and who is trusted at each layer.
- [docs/LESSONS.md](docs/LESSONS.md) and
  [docs/TRENCHES.md](docs/TRENCHES.md), why each discipline exists,
  told as the bug that created it.
- [docs/IMMEDIATE.md](docs/IMMEDIATE.md), if you know VBA: the commands
  that list, explain and reload what Frazaro understands.
- [cli/README.md](cli/README.md), the command-line door: every `frazaro`
  command, with a tour on the repository's own files.
- [CONTRIBUTING.md](CONTRIBUTING.md), the DCO sign-off and what a patch
  needs.

---

## The command-line door and the web page

*For developers, and for anyone who cannot install an add-in. The add-in
is Frazaro's first implementation of the language; this is its second,
held to the same corpus by the treaty in
[conformance/README.md](conformance/README.md), so a sentence means the
same thing through every door and refuses in the same words.*

`frazaro` is one executable with no dependencies, built from this
repository with `cargo build --workspace`; the binary lands in
`target/debug/`. It reads the files it is given and writes text, or one
workbook, and nothing else: no network, no Excel, no registry.

| Command | What it does | On the Frazaro tab |
|---|---|---|
| `frazaro translate-vla program.txt --prelude prelude.vla --phrasebook english.vla` | the program's VLA, one form per sentence, or the refusal with its line | Validate Instructions; Translate File to VLA |
| `frazaro translate-vba …` (the same arguments) | the VBA the add-in would write into a module | Show me the VBA; Translate File to VBA |
| `frazaro build program.txt … --out program.xlsx [--into model.xlsx]` | a workbook from the sentences with nothing installed; `--into` adds its sheets to a workbook of yours, untouched | nothing runs here: what a sheet holds without running |
| `frazaro rebuild program.xlsx …` | reads the sentences back out of a built workbook, builds again, and says whether the bytes still match | the workbook proves its own build |
| `frazaro reflect model.xlsx [--counts] [--cone Sheet!A1]` | the workbook's sheets, names, Tables, cells, formulas and what each formula refers to, as relations, one per line, read from the file, `.xlsx` or `.ods`; `--counts` for counts and times alone; `--cone` for a cell's cone sized through names and Tables and across sheets, counts alone | nothing yet: `REFLECT` over the live workbook comes to the add-in with `AXM.8` |
| `frazaro diff old.xlsx new.xlsx [--counts]` | what changed between two workbooks, read from the files, `.xlsx` or `.ods`: the sheets in one alone, the names and Tables that differ, then cell by cell the old and the new side by side, a formula shown with the value the file holds for it; nothing when the two hold the same; `--counts` for counts alone | nothing yet: `DIFF` over the live workbook, its Undo snapshot and a saved copy comes to the add-in with `AXM.11` |
| `frazaro audit model.xlsx [--counts]` | where the workbook's risks are, read from the file, `.xlsx` or `.ods`: constants typed over a column of formulas, formulas inconsistent with their neighbours, names nothing refers to, references to empty cells, hidden sheets, links to other workbooks, one finding a line; nothing when there is nothing to report; `--counts` for the six counts alone | nothing yet: the audit questions asked in English of the live workbook come to the add-in with `AXM.10` |
| `frazaro load phrasebook.vla --prelude prelude.vla [--allow-raw]` | loads a phrasebook with every proof run and reports what it holds | Load Phrasebook |
| `frazaro prove phrasebook.vla` | every proof in a phrasebook: `PASS n/n` or `FAIL k/n` | the proofs Load Phrasebook runs |
| `frazaro compile program.vla --prelude prelude.vla` | the VBA of a program already written in VLA | Compile and Run, for a VLA file |

The reference is [cli/README.md](cli/README.md): every command with its
arguments, output, exit codes and refusals, and a ten-minute tour on the
repository's own files. The same core compiled to WebAssembly is the web
page, [web/README.md](web/README.md): one HTML file that runs from disk,
a sentence per row with its VLA beside it, and *Download as .xlsx* for
the workbook the rows build, with the file's digest beside the button.

---

## Questions people ask

**Is it AI?** No. There is no model and no statistics. Frazaro matches
sentence shapes in a fixed order, so it can refuse but cannot guess.

**How is it different from Copilot, Office Scripts or Python in
Excel?** Same input, same result, every time, on your machine, in words
the accountable person can read. Copilot is for a draft a person will
check; the other two are for programmers. One page, with what each does
better: [docs/COMPARISON.md](docs/COMPARISON.md).

**Does it send anything anywhere?** Never. No update checks, no
telemetry, no network calls of any kind, by standing decision rather
than current accident, and a check script scans the code for one.

**Is it a macro? Will IT allow it?** It is an Excel add-in file. The
default **Interpret and Run** path needs no VBA-project trust setting
and writes no code into your workbook. Only **Compile and Run** does,
and it refuses in words if the setting is off. Treat a phrasebook
someone sends you like a macro-enabled workbook from them.
[docs/IT_REVIEW.md](docs/IT_REVIEW.md) is the two-page summary for a
reviewer: what it installs, what it can reach, what it sends, how it is
removed.

**What if it does the wrong thing?** Every run snapshots first, and
**Undo Last Run** restores, including deleted sheets. A refused sentence
is Frazaro's intended failure; a silent wrong answer is a bug, and
[docs/SUPPORT.md](docs/SUPPORT.md) says where to send it.

**Mac? Web Excel?** Mac runs the interpreted path but cannot export
VBA; the object model does not exist there. Web Excel has no VBA at
all, so no.

**Can I write my own sentences?** Yes, in a phrasebook file of your
own, layered on the base. The [phrasebooks](#phrasebooks-where-the-sentences-come-from)
section shows a rule. If a sentence you expected is missing, press
**Copy Diagnostic Report** and send it: misunderstood sentences are
exactly how the language learns new ones.

**Who builds it, and how do you know it works?** One developer,
pair-programming with an AI assistant, under an unusually heavy testing
discipline precisely because of that arrangement: a self-test suite,
golden files, both backends run over the same programs, and build-time
phrasebook audits. The stories are in [docs/LESSONS.md](docs/LESSONS.md).

**What does "beta" mean here?** The beta ends when a named person
outside the project runs their own SOP, on their own machine, on a
Monday, with the owner unreachable. Until then, every missing sentence
someone reports is the point.

**Do I own what it produces?** Yes. See [License and status](#license-and-status).

---

## Today

*This section is dated, 2026-10-03 at `0.7.1`, and expects to be
rewritten as the project moves. The sections above it should barely
change; this one should.*

- **The second implementation, owner-verified 2026-10-03:** `frazaro`,
  one executable that translates, compiles and builds a workbook from
  sentences with nothing installed, and the web page built from the same
  core with its *Download as .xlsx* button; both held to the add-in's own
  corpus byte for byte ([cli/README.md](cli/README.md),
  [web/README.md](web/README.md)).

- **Works now, owner-verified:** the English → VLA → interpret or
  compile pipeline; the ribbon (Validate, Interpret, Compile, Trace,
  Undo, What can I say?); `<Frazaro>` sections in Word documents and
  text files; the nine sample procedures on one sample workbook; the
  English and Spanish editions and the Windows installer; `=SQL()`,
  `=DATALOG()` and `=PROLOG()`; policies and questions written as
  sentences; the first four of the eleven `OPTIMIZE` items; a CLI with
  history; the build that produces the add-ins and refuses a defective
  phrasebook.
- **Honest limits:** Windows desktop Excel is the platform. Mac runs
  the interpreted path and cannot export VBA. PDF intake was measured
  and deferred; use Word. Grammar coverage grows weekly and is not yet
  complete for any one profession's vocabulary; that is what the beta
  is for.
- **Pilots wanted.** One person, one real SOP, one workbook. If that
  could be you, the address in [docs/SUPPORT.md](docs/SUPPORT.md)
  reaches a person.

---

## Install

Two paths, both documented in [docs/DEPLOY.md](docs/DEPLOY.md):

- **Standalone add-in**, the universal default. It is an Office
  document, not a program, so it works wherever Excel does, including
  managed machines. Download the latest build directly:
  [Frazaro_English.xlam](https://github.com/Spreadsheet-Company/Frazaro/releases/latest/download/Frazaro_English.xlam)
  or
  [Frazaro_Espanol.xlam](https://github.com/Spreadsheet-Company/Frazaro/releases/latest/download/Frazaro_Espanol.xlam).
  Open it, and press **Register for Auto-Load** if you want the Frazaro
  tab there every time Excel starts. Every release is listed under
  [Releases](https://github.com/Spreadsheet-Company/Frazaro/releases),
  with its notes written before it shipped
  ([docs/RELEASES.md](docs/RELEASES.md)).
- **Installer** (`FrazaroSetup.exe`), a one-click setup with a normal
  Windows uninstall entry, for machines you control. English edition
  only:
  [FrazaroSetup.exe](https://github.com/Spreadsheet-Company/Frazaro/releases/latest/download/FrazaroSetup.exe).

Either way, removal is one honest in-product button, **Uninstall
Frazaro**. Frazaro does not update itself; when you come back for a
newer build, re-read the security section below.

**Reviewing Frazaro for your organization?**
[docs/IT_REVIEW.md](docs/IT_REVIEW.md) is the two-page summary: what it
installs and writes, what it can reach, what it sends over a network
(nothing of its own), how it updates, and how it is removed.

## The shelf

This repository's [docs/](docs/) folder is unusually complete:
strategy, standing decisions, campaign histories of every hard bug, and
a set of adversarial self-reviews (a consultant's audit, a premortem, a
devil's-advocate pass, a succession audit, a platform history)
commissioned against the project's own blind spots. If you want to
evaluate the engineering culture before the code, start with
[docs/LESSONS.md](docs/LESSONS.md); if you want the current plan,
[docs/BETA_ROADMAP.md](docs/BETA_ROADMAP.md); if you want the
arguments that are not yet decisions,
[docs/CONTEMPLATIONS.md](docs/CONTEMPLATIONS.md) and
[docs/SPITBALLS.md](docs/SPITBALLS.md).

## Known open security items

Frazaro's own threat model ([docs/THREAT_MODEL.md](docs/THREAT_MODEL.md))
is public, and so is the list of what it has not closed yet. Items are
listed here by roadmap ID so a downloader hears it from this page rather
than from the repository. Frazaro does not update itself and makes no
network call, so a copy you download today stays as it is until you come
back; check this section or the roadmap to see when each closes. The same
list, dated and set beside what Frazaro installs and can reach, is in
[docs/IT_REVIEW.md](docs/IT_REVIEW.md).

**Closed.** SEC.1 (`0.5.2`): a member reference the interpreter does not
recognize now refuses in words instead of falling through to VBA's own
late-bound dispatch. SEC.2: a phrasebook rule marked `raw` (literal VBA)
shows an explicit consent dialog, naming the phrasebook, before it loads
from disk — declining refuses the whole phrasebook, not just the
`raw`-bearing rules. SEC.13 (`0.5.3`): *Load Instructions* opens Word
documents with macros force-disabled, so a document that carries an
`AutoOpen` macro no longer runs it when Frazaro reads the text out.
SEC.8 (`0.5.3`): a workbook that Windows has marked as having come from
the internet can still be read and edited as usual, but a program in its
cells is now refused when it tries to reach OUTSIDE the workbook —
opening or saving files, printing, exporting a PDF, composing an email,
or password-protecting a sheet. To allow it, unblock the file
in Windows first (right-click the file, Properties, tick Unblock);
Frazaro deliberately has no button of its own for this, so that a
workbook arriving from outside cannot carry its own permission slip.

**Open from the original threat model:**

- **SEC.3 — generated code does not yet carry phrasebook provenance.**
  Emitted VBA says what it does, not which phrasebook layer introduced
  each line.
- **SEC.7 — verbs with real external effect are not yet permissioned.**
  A phrasebook you load can open or save workbooks to a path it names,
  export a sheet to PDF, and compose an Outlook email (it is displayed
  for you, never sent silently) — with no consent prompt.

**Open from the project's own code review of 2026-09-08.** The review found
ten items, SEC.8 through SEC.17. **Four are fixed**, all in `0.5.3` —
above: SEC.8, SEC.13, **SEC.11** (the fingerprint your consent is keyed to
was weak enough to forge, and is now SHA-256) and **SEC.9** (a grammar file
beside a workbook, or a path a workbook remembers, now needs this computer's
approval before it can override the built-in grammar). Four were assessed
and **accepted** rather than fixed,
with the mitigating control written down and a stated condition that reopens
each — see *Assessed and accepted* below. Two remain open and are listed
here, most-severe first. Both are audit findings read from the code rather
than exploits anyone has run. Each one's file, line, and fix is in
[docs/BETA_REARVIEW.md](docs/BETA_REARVIEW.md). In plain words:

- **SEC.10** — the "remember my consent for this workbook" record is
  stored inside the workbook, so a workbook someone sends you can arrive
  with consent already granted.
- **SEC.15** — formulas a program writes are not screened for functions
  that reach the network or the shell (`WEBSERVICE`, DDE).

**Assessed and accepted — deliberately not fixed, and why.** Each of these
needs a precondition an ordinary install does not meet. The full reasoning,
and the condition that would reopen each one, is in
[docs/BETA_REARVIEW.md](docs/BETA_REARVIEW.md); in short:

- **SEC.12** and **SEC.17** — both are on the *Compile* path, which refuses
  to run at all unless you have turned on Excel's *Trust access to the VBA
  project object model*. That setting is off in every Office install by
  default, and managed IT departments routinely disable it outright.
  Frazaro never asks you to turn it on, and *Interpret* — the ordinary way
  to run a program — does not touch it.
- **SEC.14** — a runaway program can hang Excel, and Ctrl+Break stops it.
  A deeply nested program can instead overflow VBA's stack and crash Excel,
  which Ctrl+Break cannot stop and which can lose unsaved work. Accepted
  because the worst outcome is a lost session rather than a compromise:
  nothing runs, nothing leaves the machine, nothing persists.
- **SEC.16** — replacing the grammar files beside the add-in requires
  already being able to run programs on your machine as you. It makes an
  existing compromise durable; it does not create one.

Until these close: **load phrasebooks only from people you would accept
a macro-enabled workbook from — and treat a workbook someone sent you
the same way before you press Interpret.** The phrasebooks embedded in
the downloads are audited at build time; a `.vla` file someone sends you
is not. Found something? [docs/SECURITY.md](docs/SECURITY.md) says where
to report it and what response to expect.

## License and status

Frazaro is open source. The engine is **Apache-2.0**; the phrasebooks
are **MPL-2.0** (file-scoped: a phrasebook you write in your own file is
yours, see [PHRASEBOOK-TERMS.md](PHRASEBOOK-TERMS.md)); the one module
Frazaro copies into your workbook is **0BSD**; and everything Frazaro
generates from your sentences is yours outright, see
[OUTPUT-EXCEPTION.md](OUTPUT-EXCEPTION.md). Full texts are in
[LICENSES/](LICENSES/), the per-file map is [REUSE.toml](REUSE.toml),
and `tools/check_spdx.ps1` keeps the two honest. The name is a
trademark; see [TRADEMARK.md](TRADEMARK.md). Contributions are welcome
under the DCO, no CLA; see [CONTRIBUTING.md](CONTRIBUTING.md).

**No warranty.** This is beta software, released unfinished on purpose
(see the top of this file). It is provided as-is, without warranty of
any kind, as the licences say in longer words. Take a backup before you
run a program you have not run before; Frazaro snapshots before every
Run and offers Undo, and that is a convenience, not a guarantee.

**What a program can do with your data.** Frazaro itself makes no
network connection, ever, and phones nothing home. A program *you* run
can do what the sentences say: read and write cells, sheets and files,
and, if a sentence asks for it, open a draft email in your own Outlook
for you to send (Frazaro never sends mail itself). A phrasebook can
define new sentences, and a phrasebook rule marked `raw` can contain
arbitrary VBA; load phrasebooks from people you trust, the same way you
would open a macro-enabled workbook from them.
[docs/THREAT_MODEL.md](docs/THREAT_MODEL.md) says this at length and
without flattery.

The project is looking for pilot users, one real SOP at a time, and for
skeptical readers of the docs above.
