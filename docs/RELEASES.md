# Releases

*Newest first. `tools/release.ps1 -Version X.Y.Z` publishes the section headed `## X.Y.Z` as that release's notes and refuses to run without one, so the notes are written before the release, never after. Cadence: a `0.5.N` patch at the end of each working day, a `0.N.0` minor at the end of each week; security and safety fixes ride the patches, larger features the minors. Each section carries a short *Known open security items* block: the standing advice, what closed in that release, and a pointer to the authoritative list. It does NOT re-enumerate every open item — that list lives in `docs/BETA_REARVIEW.md` (full, with dispositions) and `README.md` (plain words), which are edited once rather than copied into every release forever. Sections written before `0.5.3` keep their longer blocks as published; they are history, not a template.*

## 0.8.0

### What changed

- **A range's median and standard deviation, and an average over the rows
  that match, each in one sentence.**
  - `Set middle to median of range B2:B50.` gives the middle value, or the
    average of the middle two when there is an even number of them. `Put
    median of range B2:B50 into cell B54.` puts it in a cell.
  - `Set spread to standard deviation of range B2:B50 as a sample.` The
    sentence says which standard deviation it means. `as a sample` is
    Excel's STDEV.S, for rows that are some of the cases; `as the
    population` is STDEV.P, for rows that are all of them. The two give
    different numbers, so a sentence that names neither is refused before
    anything runs, and the message shows both. `Put standard deviation of
    range B2:B50 as the population into cell B55.` puts it in a cell.
  - `Set west-average to average of range B2:B90 where range A2:A90 matches
    "West".`, beside the existing sum. A matching row with an empty value is
    left out, as Excel's AVERAGEIF leaves it out, so this is not the sum
    divided by the count: West rows holding 10 and an empty cell average
    10, not 5.
  - `Put sum of range B2:B90 where range A2:A90 matches "West" into cell
    D1.` puts the sum over the rows that match straight into a cell, and
    `Put average of …` the average.
  - A range remembered by a name works too: after `Remember range G2:G41
    as revenues.`, write `Set middle to median of revenues.` or `Set spread
    to standard deviation of revenues as a sample.`, or `Put` either into a
    cell. A program with a step of its own called `median` goes on using
    it.
  - Like `Put sum` and `Put average`, these put the number in the cell, not
    a live formula.
  - When Excel has no answer, the run stops at that sentence, as a lookup
    that finds nothing does, and `Try:` catches it: a median of no numbers,
    a sample of fewer than two, no matching row. For now the message is
    Excel's own, and names the function that had no answer.
- **A formula filled down as far as the data goes, in one sentence.** `Put
  formula "=B2*C2" into cell D2 and fill down to the last filled row of
  column B.` puts the formula in D2 and fills it down to the last row with
  anything in column B, as Excel's Fill Down does: each row gets the
  formula adjusted, so D3 holds `=B3*C3`, and D2's format goes down with
  it. The sentence names the column that says how far, so `… and fill down
  to the last row.` is refused, and the message shows the whole sentence.
  D2 always gets its formula; when column B has nothing below row 2,
  nothing is filled.
- **Automatic calculation turned off and on, and a recalculate, each in one
  sentence.**
  - `Turn off automatic calculation.` and `Turn on automatic calculation.`
    switch Excel's own Automatic setting (Formulas, Calculation Options). A
    long program on a big workbook can turn it off while it writes, and on
    again at the end.
  - While it is off, a formula keeps its old value when a cell it reads
    changes. `Recalculate this sheet.` recalculates the sheet the program is
    on. `Recalculate all open workbooks.` is Excel's F9: it reaches every
    open workbook, because Excel has no way to recalculate just one.
    Turning automatic calculation back on recalculates at once.
  - `Recalculate the workbook.` is refused, and the message shows both
    sentences.
- **A run leaves Excel's settings as it found them.** When a run ends,
  finished or stopped, Excel's calculation, status bar, screen updating,
  alerts and events are back to what they were before it. A program that
  turned calculation off and then stopped used to leave every open workbook
  in manual calculation for the rest of the day, and words it put in the
  status bar stayed there. So a program cannot leave Excel in manual
  calculation: for that, use Excel's Formulas tab, or type the sentence in
  the CLI, where it lasts. If you protected your program's sheet
  yourself, a run now leaves that protection alone.
- **A stop or an Undo removes the sheet the run worked on, when the run
  made it.** On a workbook without the sheet a program works on (its first
  `Work on sheet`, or Output), a run that stopped used to leave that sheet
  behind, empty, and call it "put back as they were before the run". It is
  removed now, as every other sheet a run creates is, and the message says
  so. Undo Last Run does the same.
- **A name can no longer be a word the code Frazaro writes already uses.**
  The code Frazaro writes for your sentences calls some of Excel's words by
  name: `len`, `trim`, `left`, `right`, `round`, `month`, `year`, `day`,
  `range`, `rows`, `columns`, `cells` and a few more. A step, a detail or a
  value given one of those names took Excel's place. A step called `trim`
  changed every `is empty` in its program, and a step called `range`
  changed every cell it named, whether run or interpreted. A value called
  `month` stopped a run where the program also said `month of`, though
  Interpret carried on. Each is now refused before anything runs, and the
  message names the word. A hyphenated name always works: `month-number`,
  `my-range`. A program whose name of this kind never met the word it
  shadows ran fine and is refused now too; rename it the same way.
  `length`, `count` and the other words English reads with `of` are
  unchanged, and `int`, `fix`, `cstr` and the other conversions of Excel's
  programming language join its reserved words.
- **`month of`, `year of`, `day of`, `hour of`, `minute of` and `absolute
  of` work under Interpret.** They always worked in a program that runs.
  Interpret stopped at each one, naming it. `Set m to month of today.` now
  gives the same month on both.
- **A figure can go anywhere a value goes, in words.** `median of`,
  `standard deviation of … as a sample` (or `as the population`), `sum`,
  `average`, `largest` and `smallest of range …`, and `last filled row of
  column …` now work inside any sentence that takes a value, the way `sum
  of` always has: `If median of revenues is more than 100, show "high".`,
  `Put median of revenues plus 1 into cell H2.`, `Set gap to largest of
  range B2:B50 minus smallest of range B2:B50.`, `Show standard deviation
  of range B2:B50 as a sample.` Every sentence that worked before works
  the same way. A figure said halfway is refused where it stops, naming
  what it needed: `Set spread to standard deviation of revenues.` says it
  expected `as`, and shows the whole phrase, `standard deviation of ... as
  sample|population`. **What can I say?** lists them under `value`.
- **A step of your own may be named like a phrasebook word.** `To sum of
  amounts:` was refused, because the phrasebook already gives `sum of` a
  meaning. It is accepted now: inside that program, `sum of …` means the
  program's own step, and Check marks the step's row in yellow to say so,
  "OK. In this program, 'sum of ...' means its own sum, not the
  phrasebook's". Other programs keep the phrasebook's. This is so that a
  word the phrasebook gains in a later release can never break a program
  that already had a step of that name. The words Frazaro itself reads with
  `of` (`length of`, `count of`, `first of` and the rest) are still refused
  as a step's name.
- **In Spanish, a figure is a value too.** `Pon la suma de la region B2:B9
  en la celda B10.`, and `el promedio de`, `el mayor de` and `el menor de`
  the same way, over a region written out or over a value (`Si el
  promedio de ventas es 10, …`). The Spanish phrasebook had carried these
  four figures since it was written, with no sentence that could reach
  them.
- **What can I say? files each sentence under its first word.** The
  *Phrasebook* table has three columns now: Category, Template and
  Example. A sentence's category is the word it starts with, so filtering
  Category on `set` shows every `set` sentence. The shapes built into the
  language are filed the same way (`if`, `repeat`, `to`), and the words
  and phrases that fit anywhere a value goes are under `value`. The rows
  that only named a group are gone, and a long template or example wraps
  inside the table instead of running past its edge.
- **For phrasebook authors: a function word can be several words.**
  `(english-function "standard deviation of {x:value} as
  {k:sample|population}" ({k}-standard-deviation-of {x}))` declares a
  phrase: fixed words, then one value (`{x:value}`, or a reference -
  `{r:range}`, `{c:column}` or `{c:cell}`), then at most one closing
  clause of fixed words and one choice of words. The connector is the
  phrasebook's own, so `(espanol-function "la suma de" suma-de)` works as
  `"sum of"` always has. Loading refuses a phrase that would change a
  sentence that reads today (a later word that already follows a value,
  like `plus` or `is`), two phrases with the same words, and a phrase
  whose template drops a word the reader chose.
- **For contributors: a DATALOG proof can carry the tables its program
  reads.** None of this ships, and nothing in Frazaro calls it.
  - A proof in `scripts/proofs/datalog.vla` may hold a `(tables ...)`
    clause. Each table is written as an answer is written: a name, its
    header row (or `headless`, for a plain named range), then its rows,
    each cell read as a cell of an Excel Table is. So the 29 DATALOG tests
    that needed a table moved there too, as 26 proofs, and the six that
    accepted any error now name the refusal they expect. The file holds
    124 proofs.
  - The two tests that compare DATALOG with OPTIMIZE and with its own
    integer grounder hand them each proof's tables, so both now answer
    those programs, where before they only refused them alike.
  - `tools/check_proofs.ps1` checks every table without Excel. The clingo
    export passes over a proof with tables until it can state one
    faithfully, so clingo still reads the same 23 proofs.

### Known open security items

**Closed this release:** nothing. The new sentences read the cells they
name and write only the cell a `Put` names, or for a fill down that cell
and the ones below it in its column; a recalculate runs only the formulas
already in the open workbooks, as F9 does; none opens a file or makes a
network call. Standing advice unchanged. The full list of open
items is in `docs/BETA_REARVIEW.md`, in plain words in `README.md`.

## 0.7.1

### What changed

- **A formula into many cells, and a range's figures in one sentence.**
  - `Put formula "=B2*C2" into range D2:D50.` writes the formula into every
    cell of the range. Write it for the first cell; the cells below get it
    adjusted, the way Fill Down adjusts it, so D3 gets `=B3*C3`.
  - When the last row is known only as the program runs: `Set last-row to
    last filled row of column B.` then `Put formula "=B2*C2" into rows 2 to
    last-row of column D.` (`through` works too). If there are no data rows,
    nothing is written, so the header row stays as it is.
  - `Set top to largest of range B2:B50.` (or `smallest`), without
    remembering the range first.
  - `Put average of range B2:B50 into cell B51.` (or `largest`, `smallest`),
    beside the existing `Put sum of range …`. The cell gets the number, not
    a live formula; for one that recalculates, use `Put formula
    "=AVERAGE(B2:B50)" into cell B51.`
  - `Set blanks to count of empty cells in range A2:A50.` and `… count of
    filled cells …`. A cell whose formula shows nothing counts as both, as
    Excel's own COUNTBLANK and COUNTA count it.

- **`contains`, `does not contain` and `starts with` now give the right
  answer under Interpret and Run.** Before, when you pressed **Interpret and
  Run**, `If cell A1 contains "x", …` never ran its sentence, `… does not
  contain …` always did, and `… starts with …` almost never matched.
  **Compile and Run** was always right. The two now agree, and both still
  ignore capitals. If you ran a program that uses one of these conditions
  with Interpret and Run, run it again.
  - In the middle language: under Interpret, `(msgbox …)` now takes its
    buttons and title and gives back the button pressed, `(inputbox …)`
    takes its title and default, and a built-in given the wrong number of
    values is refused by name instead of quietly using some of them.

- **A word shaped like a cell is no longer taken as a name.** `Set A1 to 5.`
  used to keep 5 under a name `a1` and leave cell A1 empty, and `Put A1 plus
  B1 into cell C1.` read two names that were never set. Both are now refused
  before anything runs, and the message says what to write: `cell A1` for the
  cell (`Put 5 into cell A1.`), or a name that is not shaped like a cell
  (`q1-total`). A word is shaped like a cell when it is one to three letters
  and then digits: `A1`, `q1`, `fy24`. Sheets, tables and pivots may still be
  called `Q1`. If one of your programs used a name like that, rename it.

- **A refusal shows its first character again.** A message that begins by
  quoting a word, such as `'seek' is a reserved word …`, appeared in column C
  as `seek' is a reserved word …`: Excel hides an apostrophe at the start of
  a cell. Every message now shows exactly as written, and so does the log
  that **Copy Diagnostic Report** copies.

- **Replace changes values, and never a formula.** `Replace "N/A" with 0 in
  column C.` used to be Excel's own Find and Replace, which also rewrites the
  text inside formulas: replacing "A" with "B" in a column could quietly
  turn `=A1*2` into `=B1*2` and point it at a different cell. Now it changes
  only the values you typed, and formulas are left exactly as they were.
  `"N/A"` replaced by 0 still becomes the number 0. To change formulas, say
  so: `Replace "Sheet1" with "Data" in the formulas of range B2:B20.` (or `of
  column D`). Both work on a whole sheet too: `Replace "2025" with "2026" on
  this sheet.` or `… in the formulas on this sheet.`

- **Split a column.** `Split column C by "," as text.` keeps the first piece
  in C and puts the rest in D, E and so on. Every piece stays text, so
  "0042" keeps its zeros. `Split column C by "," reading numbers.` makes a
  plain number like 42 or -3.5 a number and keeps everything else as text.
  Neither ever turns "3/4" into a date, and the sentence must say which one
  you want. A split never writes over anything. If a cell the pieces need
  already holds something, nothing changes and the message names that cell,
  so you can clear it or insert columns first.

- **Find text in a range, and count it.** `Set r to the row of the first
  cell in range A1:D50 containing "total".` (or `the column of`, or `in
  column B`) gives a row or column number you can use with `column number …
  row …`, or 0 when nothing matches. `Set n to how many cells in range
  A2:A99 contain "late".` counts them. Both find text in any case, just as
  `If code contains "late"` does, and a `*` or `?` is just that character.
  To ask first: `If range A1:D50 contains "total", …` or `If column C does
  not contain "x", …`.

- **DATALOG and SQL group text exactly, the way they match it.** A DATALOG
  `count`, `sum` or `textjoin`, and a SQL `GROUP BY`, used to put "Bob" and
  "bob" in one group, although a join, `DISTINCT` and `=` always told them
  apart. So a count could say 2 where only one row belonged to Bob, and a
  `textjoin` could list the same value twice. Now each spelling is its own
  group, as in SQLite. If one of your questions groups text that differs
  only in capitals, its answer changes, so check it.

- **For contributors: DATALOG's tests are proofs now, and a second solver
  checks them.** None of this ships, and nothing in Frazaro calls it.
  - `scripts/proofs/datalog.vla`: 94 DATALOG tests, moved out of
    `VLA_Tests_Query.bas` and written as the programs they are, each beside
    the answer it must give. Every row is named, and every refusal is named
    by its message id, where many of the old tests accepted any error at
    all. `TestDSLs` runs them.
  - `tools/proofs_lp.ps1`: writes each proof it can translate as a program
    for clingo, the answer-set solver used by hand as a reference, so an
    answer is also checked by software written elsewhere. It never runs
    clingo and makes no network call. clingo agreed with all 23 proofs it
    can read today.
  - `tools/check_proofs.ps1` (new): checks the proof file without Excel, and
    fails when the clingo files fall behind it.

### Known open security items

**Closed this release: Replace could create a formula.** Excel's own Replace
enters every cell it changes again as if you had typed it. A value that ended
up starting with `=` became a live formula, which could reach outside the
workbook, the way a formula in a downloaded file can. Every other value
Frazaro writes had long been kept as text in that case. Replace now works the
same way: a changed value that would start with `=`, `+`, `-` or `@`, and is
not a plain number, is kept as text. Splitting a column follows the same
rule. Finding and counting only read. Standing advice unchanged. The full
list of open items is in `docs/BETA_REARVIEW.md`, in plain words in
`README.md`.

## 0.7.0

### What changed

- **Your SOP can stay exactly as it is. Put `<Frazaro>` on a line above the
  part Frazaro should run and `</Frazaro>` on a line below it, and Frazaro
  reads only that part.** The title, the purpose, the notes for people and
  the revision history stay as you wrote them, with no `#` needed anywhere. A
  document can have several tagged sections; one left open at the end runs
  to the end of the document; and a document with no tag is read from top to
  bottom, exactly as before. The opening tag can carry more.
  `week="Week 39"` is a value the section writes as `{week}`, filled in as
  the document loads, so running it for another week means changing it in
  one place. Anything in parentheses is a note for people. `0.7.0` says the
  section needs this Frazaro or later, and `espanol` that it is written for
  the Spanish phrasebook. A tag Frazaro cannot read is refused with its file
  and line before anything loads, and the program on your sheet stays as it
  was. If you load a whole SOP with no tag and its first check fails, the red
  row now says how to add one. The new sample **`00 Weekly Expense
  Reimbursement`** (`.docx` and `.txt`) is an ordinary SOP with two lines
  added, and `examples/README.md` now starts there.

- **Messages point you at the ribbon, never at VBA.** When Frazaro cannot
  read a sentence, it used to suggest printing a list in VBA's Immediate
  window, and it named a Known Sentences button the ribbon no longer has; it
  now says to press **What can I say?**. Seven other messages that sent you
  into VBA now say what happened in plain words, or name the button that
  helps. If you do know VBA, `docs/IMMEDIATE.md` lists the commands.

- **Questions that follow a chain — "who reports-to Alice directly or not" —
  are several times faster, and answer exactly what they answered before.**
  Two things were costing the time. The first is the sillier one: to work out
  which rows of a table a rule cares about, Frazaro was creating a small
  scratch object for *every single row*, even when the rule cared about all of
  them and the scratch object never rejected anything. It now works out what a
  rule is actually asking once, and where a rule asks for everything it simply
  uses the table as it stands instead of copying it row by row. The second is
  that a question about one person used to build every connection between
  every pair of people first, and only then narrow down to the person asked
  about; where it can be done without changing the answer, the person's name
  is now pushed into the search before it runs, so the connections nobody
  asked about are never built. Measured on the same org-chart questions as
  before, on the same machine: a chain of a hundred people **4.1 seconds to
  0.5**, and an organisation of three thousand **11.3 seconds to 0.3** — the
  same answers, thirty-six times sooner. Chains of 250, 500 and even 1,000
  people, which were previously too slow to be worth attempting at all, now
  finish in 1.4, 3.7 and 10.2 seconds.

- **Ordinary questions got faster too, which was not the point.** The scratch
  object above was created on every row of every table any rule read, so this
  was never only about org charts. A plain "which bill is big" over ten
  thousand rows went from **3.8 seconds to 0.3**. Counting and list-making
  questions gained the same way.

- **A hundred thousand rows, answered in under four seconds.** This is the
  one worth stopping on. The previous release could scan ten thousand rows in
  about four seconds, and put a hundred thousand at roughly thirty-eight —
  comfortably past the point where a spreadsheet feels broken. A hundred
  thousand rows now answers in **3.7 seconds**, with all 49,001 matching rows
  correct. That is measured, not projected: it is the size Frazaro's own
  scale testing was designed around a fortnight ago and could never actually
  reach. It is still the largest size anyone has asked a question at, so
  treat it as the edge of what is known rather than the edge of what works —
  but the edge moved by a factor of ten.

- **Nothing about any answer changed, and that is the part that was tested
  hardest.** The narrowing is only applied where it provably cannot be seen:
  Frazaro checks that every other part of the program already asks about the
  same person, that the relation being narrowed is not itself what the cell
  shows, and that the person's name is not used anywhere it would be read as a
  value rather than a name. If any of those is not true, it does the work the
  long way. Fifteen worked examples with their answers written out by hand
  were used to prove it before any of it went into Excel, and 190 test
  programs are required to give the same answer through two different engines.
  On the machine this was measured on, every one of the questions above was
  checked against its expected answer before it was timed, at every size.

- **The CLI remembers what you ran.** In the Frazaro CLI, **Ctrl+Up** brings
  back the command before, and the one before that, and **Ctrl+Down** walks
  forward again to whatever you were in the middle of typing, which is kept
  while you look back. A command is kept *before* it runs, so one that
  Frazaro refused is there to fix rather than to retype from memory. Type
  `history` to list the last twenty with their numbers (`history 50` lists
  fifty), and `!7` to put number 7 back in the box without running it. The
  history carries over from one Excel session to the next, in a file of
  your own - `history.txt`, in `%APPDATA%\Frazaro` - and never in your
  workbook: what one person typed at their console is not something a
  workbook they send on should carry. A command that mentions a password,
  a secret, a token, an API key or a credential is kept only until Excel
  closes, and is never written to that file. The **Clear History**
  button, or **Ctrl+Shift+Delete** from the keyboard, forgets them all,
  and the file with them, after asking, since that cannot be undone.

- **The CLI shows what came back.** Along the bottom of its window, under
  the box, the CLI now keeps a transcript: each command after a `~`, then
  whatever it printed, whatever Frazaro had to say about it - a refusal is
  written there in full, in its own words, instead of in a dialog you have
  to dismiss - the answer, when it worked one out, and how it ended: `OK`,
  `Failed` or `Translation failed`, with the time. `(* 6 7)` answers
  `= 42`. The last three answers are there to use in the next command, the
  way every Lisp listener has had them since the 1970s: `*` is the last,
  `**` the one before, `***` the one before that, so `(* * 2)` doubles the
  last answer. The transcript lasts until Excel closes and is never saved
  anywhere: it holds values from your own workbook. Type `clear` to empty
  it; the history is kept.

- **The CLI's window is simpler, and all of it works from the keyboard.**
  **Run** and **Clear History** sit above the box, each naming its keys:
  **Ctrl+Enter** runs, and **Ctrl+Shift+Delete** clears the history, after
  asking. **Esc** closes the window from anywhere in it, as its X does. The
  example the box opens with starts out selected, so the first thing you
  type replaces it, and its opening comment lines say which keys do what.
  A VLA command may now begin with blank lines or `;` comment lines, as
  that example does; before, either sent it to the English reader, which
  refused it.

- **`OPTIMIZE` makes choices.** Until now it could only check a schedule
  you already had. Now it builds one. Write who may work which shift, how
  many each shift needs, and the rules no schedule may break, and
  `OPTIMIZE` answers with the first schedule that breaks none. How many
  can be exactly so many, at least, at most, or between two numbers, each
  written in or read from a column of your own table, or left open, as
  for the extras a customer may or may not add. Several counts can cover
  the same rows, so "ten on every shift", "a senior on every night" and
  "at most five shifts a week" hold together. It fills your first shift
  first, and takes your rows in the order your tables list them.
  `OPTIMIZE_STATUS` says why this is the answer, and how much searching
  it took: with nothing to make as small or as large as possible, every
  schedule that breaks no rule is as good as another, and this is the
  first of them.
  - **When there is no schedule, it says so and says why.** That is
    either the rules that cannot all hold together, by number, leaving
    out any rule that plays no part, or the arithmetic that rules it out
    before anything is tried: "7 guests, and 2 tables of 3 seat at most
    6". A rule no choice could meet is listed row by row by
    `OPTIMIZE_VIOLATIONS`.
  - **How hard it may try is yours to set, in the rules themselves:**
    `(effort quick)`, `normal` or `thorough`, or a number. It is counted
    in work rather than seconds, so a workbook answers the same on a
    fast machine as on a slow one. If the effort runs out first, it says
    there may still be a schedule rather than claiming there is none. And
    a search inside a formula is stopped after ten seconds whatever its
    effort says, and the answer tells you that is what happened. The three
    levels are set on a real roster of 50 people over four weeks. There,
    `thorough` is about a second of searching, which with everything else
    stays under the two seconds a formula may take. `normal`, the default,
    is a tenth of that, and `quick` a tenth again.
  - **A four-week roster is answered in under a second.** It has 50
    people and 84 shifts of ten, with a day's leave for each person every
    week. There is a senior on every night, nobody works more than five
    shifts a week, and nobody works two in a row. `OPTIMIZE` finds the
    first schedule that breaks none of that in three-quarters of a second,
    without taking back a single choice. With one more person needed on
    every shift, it finds none within any of the three efforts, and says
    there may still be one.
  - **It lays out the possibilities much faster.** Before choosing, it
    has to spell out every possibility and every rule that could forbid
    one, and it now does that a new way, on numbers instead of on the
    text in your cells. A roster rule that took nine seconds (9.2) now
    takes about half a second (0.56). It is checked against `DATALOG`
    itself on every rule of 192 test programs, row for row, every time
    the tests run.
  - **A program too big for a cell is refused before the work is done,
    and it says which rule and how big.** A formula lays out at most
    100,000 rows of possibilities, and at most 50,000 in any one step.
    Each step is counted before a row of it is made, so a rule that
    would pair 18 million rows is refused having made none of them.
    Laying a program out right up to that limit was measured at under a
    fifth of a second. The refusal names the rule, the rows it would
    make and why - "the 300 rows of 'pick' share no name with the 300
    before it, so every one pairs with every one". A rule is worked
    through from its smallest part, whatever order you wrote it in. That
    is what lets "never two shifts in a row" fit at the size of a real
    four-week roster: written the obvious way round, it would pair
    352,800 rows in one step.
  - **What a cell refuses, a command runs.** Select the cell and choose
    Frazaro > Logic Engines > Optimize Selected Cell, or write
    `Optimize cell C1.` in the console or in a program you Interpret. The
    cell's own rules and Tables run again, with no ten-second limit and
    room for 500,000 rows in a step, and Esc stops it. The answer lands
    as values on a new sheet named after the cell, with its status above
    it, and nothing on your sheets is overwritten. A cell's size refusal
    now ends by pointing here, and a program still too big for a command
    is refused in the same words, saying "a command". A compiled program
    cannot run the command, since the OPTIMIZE engine does not travel
    with it, and Compile says so by name.
  - **Not yet:** "as few as possible" and "as many as possible"
    (`minimize`, `maximize`), preferences, staying close to last month's
    roster, a rule that reads what is being chosen, and a count of it
    inside a rule. Each is refused in words that say so, and the last two
    say what works today: a condition on chosen rows written as a rule no
    schedule may break, and a limit on how many written as another count
    over the same rows, as "at most five shifts a week" is.

- **One message's dash is a plain hyphen now, like every other message's.**
  The refusal of a stray `Done.` or `Otherwise` - "is there a stray 'Done.'
  or a missing block?" - was the only message whose words held a character
  outside plain ASCII, which Frazaro's code is not meant to contain.

- **Text can be tidied where it stands.** Five new sentences change the text
  in a range, or in a whole column, with no helper column and no formula:
  `Make range A2:A50 upper case.` (or `lower case`); `Capitalize each word in
  column B after any space.`, or `after any non-letter`, which is Excel's own
  PROPER - "o'neil" becomes "O'Neil" and, as in Excel, "don't" becomes
  "Don'T" - and the sentence says which, so nothing is guessed; `Remove extra
  spaces from column A.`, which is Excel's TRIM and also counts the
  non-breaking spaces that text pasted from the web is full of; and `Remove
  non-printing characters from range A2:A50.`, which is Excel's CLEAN, so
  tabs and line breaks inside a cell go. Only text changes. Numbers, dates,
  formulas, errors and blank cells are left exactly as they were, and text
  stays text: a code like " 00123 " keeps its zeros when trimmed instead of
  becoming the number 123, and "true" in capitals is still the word TRUE.

- **"Replace ... with ... in range" means the same thing every time.** Excel
  remembers the last settings its Find and Replace used - whether to match
  the whole cell, whether capitals matter - and quietly reuses them when a
  program does not say. So the same Replace sentence could change text
  anywhere in a cell one day and only whole cells the next: after someone
  used Ctrl+H, or even after an earlier sentence in the same program looked
  something up with `row of ... in column ...`. It now always matches
  anywhere in a cell, in any case, which is what it did for anyone who never
  changed those settings.

- **Pieces of text can be taken apart and put together in a variable.**
  `Set part to the text before "-" in code.` (or `after`, or `the text after
  the last "."` for a file's extension) finds its marker in any case, just as
  `If code contains "-"` does, and stops with a plain message naming the text
  when the marker is not there, rather than carrying on with a wrong value.
  `Set prefix to the first 3 characters of code.` counts every character, not
  only letters. `Set code to id padded on the left with "0" to 5 characters.`
  turns 42 into "00042", and never cuts a longer code short. `Set names to
  range A2:A9 as one list.` gives "a, b, c" (or `separated by "; "`), row by
  row, skipping blank cells, each cell's value rather than its display. And
  everything the new cell sentences do can be done to a value too: `Set tidy
  to name with extra spaces removed.`

- **A Run that stops puts its sheets back, and says where it stopped.** When
  a program stops part-way - a sentence refuses, or something goes wrong
  that no `Try:` catches - Frazaro now puts back what Undo Last Run would
  have: the Output sheet and every sheet the program names, as they were
  before the run, and removes any of those the run had created. Then one
  message says which line stopped it and quotes the sentence, gives the
  reason, and lists what was put back; the row is marked "Stopped here" on
  the program sheet. So a stop at step 50 of 100 no longer leaves the first
  49 done. What a program did outside those sheets - a file it saved, an
  email it drafted - stays as it is, and the message says so. This works
  the same whichever way the program runs. Before, Interpret's stop gave
  only the reason, with no line, and Compile and Run's labelled Frazaro's
  own refusals "Excel says".

### Known open security items

**Closed this release:** nothing. The query changes above are speed only: the
same questions, over the same tables, giving the same answers, with no new
capability and no new file or network access. `OPTIMIZE`'s choosing is new,
and with its command it is the one new thing a workbook can ask Frazaro to
do. A formula reads only the tables it names, writes nothing and calls
nothing outside Excel. Its search stops at the effort its own rules allow,
or after ten seconds of searching, whichever comes first. Laying out a
program's possibilities is bounded too: past 100,000 rows, or 50,000 in one
step, a formula refuses it before doing more. The command - the Optimize
Selected Cell button, or the sentence `Optimize cell C1.` - runs a cell's
own `OPTIMIZE` formula again: it evaluates only that formula's own
arguments, allows 500,000 rows in a step, has no ten-second limit and
stops on Esc. It writes one new sheet of values, never a formula, and
overwrites nothing; a refusal, or Esc, writes nothing. The part of a
program that is plain `DATALOG` - its facts and ordinary rules - is not
bounded yet, just as a `DATALOG` question is not.
The CLI's history is new. It is written to one file of
your own, `%APPDATA%\Frazaro\history.txt` - the first file Frazaro keeps in
your profile - and never into a workbook; nothing in a workbook can read it or
ask for it, and a command in it only ever comes back into the CLI's box, never
runs by itself. A command that mentions a password or another secret is not
written to it at all. The CLI's transcript is kept in memory only, while Excel
is open, and written nowhere. Nothing in any of this makes a network call.
The new text sentences change only cells that already hold text, in the range
or column the sentence names, and never write anything that becomes a formula:
text that would start with `=`, `+`, `-` or `@` is kept as text, the same
guard every value Frazaro writes already has. The sentences that work on a
value only read: a range as one list reads the cells the sentence names and
writes nothing.
A Run that stops puts back only the sheets its own snapshot copied before it
began, by the same steps Undo Last Run has always taken; it reaches no other
sheet, workbook or file.
Standing advice unchanged. The full list of open items is in `docs/BETA_REARVIEW.md`,
in plain words in `README.md`.

## 0.6.2

### What changed

- **Safety fix: loading a Word document could close one you had open, losing
  unsaved edits.** If Word was already running, Frazaro used to borrow it —
  and asking Word to open a file it *already had open* gives back the document
  you are working in, not a copy. Frazaro then closed it, and anything you had
  not saved went with it. That was most likely to happen in exactly the
  situation it is worst: editing a procedure in Word, then pressing **Reload
  Instructions**. Frazaro now starts its own copy of Word, hidden, for the
  moment it takes to read the file, and closes it afterwards. It never changes
  a setting in the Word you are using, never closes a document you have open,
  and never leaves one behind. Two consequences worth knowing: loading a Word
  document now takes a second or two longer - the status bar says what it is
  doing while Word works - and Frazaro reads the file **as saved**, so save in
  Word first, then Reload.

- **Frazaro now says why it cannot read a file, instead of showing a
  programmer's error or quietly loading nothing.** Five new messages, in plain
  words: Word is not installed or will not start; you are on a Mac, where
  Frazaro cannot drive Word (save as `.txt` instead); the document holds no
  text Frazaro can read — a scan or a page of screenshots is a picture of
  text, and there is no OCR; the file is protected by a password, which
  Frazaro never asks for (remove it in Word and save again); and the file is a
  PDF. Each one leaves the procedure already on the sheet untouched, which the
  message says — before this, a document with nothing readable in it would
  clear your program and leave a single empty row where it had been.

- **A password-protected Word document no longer freezes Excel.** Word asked
  for the password on a window nobody could see, so Excel simply stopped.
  Frazaro now tells Word not to ask, and refuses in words instead.

- **PDFs are refused by name, and the reason is not a shrug.** Frazaro was
  going to read PDFs through Word's own PDF conversion. Measuring that
  conversion on our own sample procedures stopped it: a PDF does not store
  lines and paragraphs, only marks on a page, so Word rebuilds them by
  measuring gaps — and it drops the blank line that ends an indented block.
  That silently changes what a procedure *does*, while every line still passes
  Validate: one of our samples came back with three steps swallowed into a
  loop that would have run them twenty times. It also joins steps that sit
  close together, about once a page. So Frazaro refuses a `.pdf` and tells you
  the way round it — open it in Word, save as `.docx` or `.txt`, load that —
  along with what to check, because that same conversion is the one you will
  be running by hand.

- **A question can read another formula's spilled answer as a table.**
  If a formula spills a table with a header row (a `FILTER`, a `VSTACK`,
  or another `DATALOG` question), give that spill a name in Name Manager,
  with `Refers to:` set to the spill's reference, for example
  `=Sheet1!$D$1#`. Then pass the name to `DATALOG`, `SQL` or `PROLOG`. The
  first row is read as column names, so a rule can say "Schedule lists the
  person as Name", and the name keeps working as the spill grows or
  shrinks. This is what the coming `OPTIMIZE` needs: one search, read by
  as many questions as you like. Excel Tables and ordinary named ranges
  work exactly as before, including a named range that covers only part of
  a spill. You are told in words, rather than getting a wrong answer, when:
  a spill has no name (the message gives the exact `Refers to:` text to
  type); two names point at one spill; or the first row looks like data (a
  number, a blank, TRUE/FALSE or an error), or has two columns with the
  same name. The last guards against a `SEQUENCE` or a `FILTER` over a
  Table's rows losing its first row as "headers". A reference such as
  `D1#` to a cell that is not spilling now says that it arrived as an
  error, not that it is "not a range". Tested live, fourteen checks each on
  a fresh sheet: one `DATALOG` answer read by a second question by its
  column names, the name following a spill as it grew, `SQL` and `PROLOG`
  reading the same spill, every refusal above appearing in the cell with
  its cell named, and a plain named range reading exactly as before. The
  new tests all pass, and every existing suite gave the same counts as
  before.

- **`OPTIMIZE` exists, and so far it answers exactly what `DATALOG`
  answers.** The new function is the first piece of the optimizer: for a
  program of facts, rules and one question it *is* `DATALOG`, the same
  engine under a different name, and it spills the same table with the same
  header row. That sounds like nothing and is the foundation for
  everything: the thousand-odd checks that already prove `DATALOG` right
  now also run through `OPTIMIZE` and must give identical answers, so
  nothing the optimizer adds later can quietly change what a question
  already answers. `=OPTIMISE(...)` works too, for anyone who spells it
  that way. A companion `=OPTIMIZE_STATUS(...)`, given the same arguments,
  says in words how good the answer is — today, always "proven best: this
  program makes no choices, so it has exactly one answer and nothing was
  searched", because with nothing to choose there is nothing to search.
  Asking the same question in two cells only works it out once.

  **The words you will write are settled now, before anything uses them.**
  You can already type a choice (`choose-exactly`, `choose-at-least`,
  `choose-at-most`, `choose-between`, or `choose-any` for "each one, in or
  out"), a hard rule (`require` and `forbid`), a preference with an
  optional weight (`prefer` and `avoid`), a goal (`minimize` and
  `maximize`, and the British spellings; the order you write them in is
  the order they matter), the schedule to stay close to
  (`fewest-changes-from`), and how hard to try (`effort quick`, `normal`,
  `thorough`, or a number). Each one is read, checked for shape, and then
  refused in a sentence that says plainly that this version makes no
  choices and points you at the answer it *can* give. So nothing you write
  today has to be rewritten when the searching arrives: the spelling is
  the same spelling. Write `choose` on its own and it tells you the five
  words that work. Write a shape wrong — a count that is not a number, a
  range the wrong way round, a word where a row belongs — and it says
  which, rather than guessing.

  Two details worth knowing now, because they decide what your other
  formulas can do. When there is no valid answer, the cell will spill its
  **header row with nothing under it** rather than a word like "None", so
  a `FILTER`, a `COUNTIFS` or a second question reading it sees zero rows
  and nothing breaks; the words live in the status cell, never in a row
  where they would be counted as data. And a genuinely unreadable program
  still shows readable `#OPTIMIZE!` text, where breaking those readers is
  the right thing to do.

  Tested with three hundred and twenty-one new automated checks, and then
  by hand in a live workbook: fourteen numbered steps, each on its own
  fresh sheet, covering the same program answered identically by
  `=DATALOG(...)` and `=OPTIMIZE(...)`, the British spelling, the status
  cell, every one of the six forms refused in its own words, the
  header-row-and-nothing-under-it answer read by a `ROWS()` beside it, one
  `OPTIMIZE` answer read back by a second `OPTIMIZE` question through a
  Name Manager name, and a table cell holding `Zoë`. The automated total
  includes every existing `DATALOG` test program re-run through
  `OPTIMIZE` and required to give an identical answer — that is what makes
  "the same engine" a measured claim rather than a description. Every
  other suite reports exactly the counts it did before.

- **`OPTIMIZE` now checks your rules against your data, and tells you
  which rows break which rule.** Write `(forbid ...)` or `(require ...)`
  beside your facts and rules and the cell no longer says "not yet" — it
  answers. Nothing is being searched for: with no choice to make there is
  exactly one way things are, and these are checks over it. So "does this
  month's close assignment break any rule, and which rows?" is a question
  you can now type.

  Three cells, given the same arguments, say three things about one
  answer, and the search behind them runs once. `=OPTIMIZE(...)` gives the
  rows — or, when a rule is broken, its **header row with nothing under
  it**, because rows that break a rule are not an answer and handing them
  back would look like one. `=OPTIMIZE_STATUS(...)` says "no schedule
  satisfies every rule", then names which checks are broken and by how
  many rows. And the new `=OPTIMIZE_VIOLATIONS(...)` spills the evidence
  as a table you can filter, count and ask further questions about:
  `Check`, `Rule` and `Where` — the rule exactly as you wrote it, and the
  rows that broke it, named. Those three columns are the same for every
  program, so a formula written over one of these tables keeps working
  over the next. With nothing broken it spills its header row and nothing
  under it, in the same shape, so nothing reading it breaks either.

  **A requirement over your data is checked, never quietly granted.**
  `(require (approved t2))` when `t2` is not in your approvals table
  reports that it is not — it does not add a row so that it is. For a
  compliance check that distinction is the whole point.

  Everything you could already write inside a rule works inside a check:
  "nobody reviews what they prepared", "every reviewer is senior",
  "nobody on the rota is untrained" (a `not`), "nobody holds more than two
  roles this month" (a `count`), "nothing over forty hours" (a
  comparison). Both directions, too — `require` says what must hold
  wherever something else does, `forbid` says what may never all hold at
  once, and neither makes you write the other one inside out.

  A rule of the form "every night has *some* senior on it" is still
  written as a rule that finds them followed by a `require` over it, and
  if you write it as a bare requirement the message says so and shows you
  the two lines. Nothing else about how a program is written changed, and
  every `DATALOG` question answers exactly what it did.

  Tested with ninety-two new automated checks, and then by hand in a live
  workbook: fourteen numbered steps, each on its own fresh sheet. Among
  them is the segregation-of-duties case this engine was scoped around,
  answered end to end: three close tasks, four staff, and the one row
  where the same person prepares and reviews and is not senior either —
  both broken rules named, on that row, with the workload cap correctly
  *not* firing. Also checked by hand: the evidence table named through
  Name Manager and then read back by a `DATALOG` question by its own
  column name; a check written with your Table's column names rather
  than positions; and the three other engines answering exactly as
  before.

- **Fixed: a rule that mentions no names at all no longer breaks the
  question.** A rule whose body is entirely fixed values — "carbon wheels
  need the carbon frame", naming two catalogue items and nothing else —
  used to stop the calculation with a Visual Basic error instead of
  answering. It affected `DATALOG`, `SQL`, `PROLOG` and `OPTIMIZE`
  equally, and it had been there since the first version of the
  question engine; nothing had written such a rule until a *check* needed
  to, because a rule that names nothing derives only one fixed fact and
  looks pointless until you are ruling something out. Both shapes of it
  are now permanent test cases.

- **When a question names something that does not exist, the message now
  tells you what does.** If a rule or a question names a relation nothing
  defines — almost always a misspelling, or a table whose name is not what
  you think — `DATALOG`, `SQL`, `PROLOG` and `OPTIMIZE` all told you to
  "check that the table argument's name matches" without saying what the
  names actually were. They now end with, for example:

  > The names this program does define, table arguments first, are:
  > parent, kid.

  Table arguments come first deliberately, because a table whose name is
  not what you expect is the usual cause. The commonest way to get there:
  select a range, press Ctrl+T, and then type the name into the **Name
  Box** above column A instead of the **Table Name** box on the Table
  Design tab. The first makes an ordinary defined name and leaves the table
  called `Table1`; only the second renames the table, and it is the
  table's own name that a rule must use. That looks identical from the
  formula bar, and it cost the author of this release two test steps before
  the message was improved to say so.

- **Non-English text in a table can no longer be confused with other
  text.** Behind the scenes, `OPTIMIZE` remembers an answer so that asking
  again costs nothing — and it recognises a repeat question by a
  fingerprint of your rules and of every cell of every table you passed.
  Making that fingerprint exact meant fixing something real: the only text
  fingerprint Frazaro had went through Windows' legacy code page, where
  `Zoë` and `Zoe?` could come out identical. They no longer can, for any
  text Excel can hold. Numbers are fingerprinted by their exact value and
  never by how they are displayed, so no locale or rounding gets in, and
  `1` and `"1"`, a blank cell and an empty one, are all told apart.
  Nothing you can see changes; what it buys is that a remembered answer is
  never the wrong answer. The memory is also allowed to vanish at any
  moment — editing the macro project empties it — and losing it costs a
  moment's recalculation and never changes a result.

- **The test questions for the coming `OPTIMIZE` function, answered before
  it exists.** Twenty small, real questions the optimizer will have to
  answer, from five kinds of work. Who prepares and who reviews each
  month-end task, so nobody checks their own work. A weekend roster where
  two people can't share a shift. A wedding seating plan with wishes of
  different weights. The best-margin bike quote under a price cap. Fitting
  exams into sittings so nobody sits two back to back. Each answer was
  worked out by hand first, then re-checked by a script that tries every
  possible arrangement, and again by clingo, an established answer-set
  solver run by hand outside Frazaro. All three agreed on every answer.
  Seven of the twenty have no valid answer, and four
  of those fail on simple arithmetic, such as more jobs than people to do
  them; the engine will be expected to say that in one sentence, before it
  searches. Nothing you can run changes: this is the yardstick the engine
  will be measured against.

### Known open security items

**Closed this release:** nothing. None of the changes above is a security
item: one reads a spilled range that Excel already showed you, one is a set
of test questions, and the new function answers what `DATALOG` already
answered. The exact-fingerprint fix is worth naming here even so, because
it touches the same SHA-256 the phrasebook-consent record is keyed on: that
record was never affected — it has always been fingerprinted from raw
bytes, not through the code page — and the fix adds a new way to
fingerprint text without changing the existing one. Nothing in any of this
makes a network call. Standing advice unchanged. The full list of open
items is in `docs/BETA_REARVIEW.md`, in plain words in `README.md`.

## 0.6.1

### What changed

- **Safety fix: Uninstall Frazaro could delete the wrong workbook.** In
  0.6.0, with another workbook open that carries Frazaro's own code (a
  developer's working copy, or anyone who imported the modules into their
  own file), pressing **Uninstall Frazaro** could run that workbook's copy
  of the command. It then removed *that* workbook from disk, permanently,
  bypassing the Recycle Bin. It happened to the project's own development
  workbook. Uninstall now refuses to touch anything but a Frazaro add-in
  file (`.xlam`). It checks twice, once when you press the button and again
  just before the file is removed, and the removal step refuses any other
  kind of file. The same dialog also told you to delete, by hand, any
  folder named `scripts` next to the add-in, whether or not it was
  Frazaro's. It now lists Frazaro's own files (`prelude.vla`, `english.vla`,
  `espanol.vla`) by name, and only when that folder holds nothing else.
  Tested live both ways: pressed with the development workbook active, it
  refuses and names that workbook; pressed from a blank workbook, it removes
  the add-in and leaves its `scripts` files where they were. If you use
  Uninstall Frazaro, update first.

  *Still open:* the cause is that Frazaro's ribbon buttons name their
  command without naming the add-in, so while another workbook with
  Frazaro's code is active, every button runs that workbook's copy. Only
  Uninstall can delete a file, and it is now guarded. The rest is `U.24` on
  the roadmap. If you keep Frazaro's modules in a workbook of your own,
  click into a different workbook before pressing a Frazaro button.

- **Questions about your rules answer five to ten times faster, with the
  same answers.** 0.6.0 measured how large a Table a question can read, and
  most of the time went on one small check repeated about thirteen times a
  row. It now costs almost nothing. On the same machine as before: a
  thousand rows answer in under half a second (they took over two), and ten
  thousand in about four seconds (they took twenty-five). A rule written
  with "otherwise" gained most and now costs about three and a half times
  an ordinary one, not five. "Directly or not" is faster too, but gained
  least, because most of its cost is working out every pair in the chain
  before narrowing to the person you asked about: a thousand people in a
  wide organization now take about three seconds, and a hundred in one
  reporting line about four, down from twenty. A hundred thousand rows are
  still out of reach. The generated code in your own workbook uses the
  same check, so it gets the saving the next time you compile.

- **A new release check: `tools/check_vladict_guard.ps1`.** The speed-up
  above is invisible to every test: putting the slow check back gives
  exactly the same answers, so every suite would stay green while every
  question got six times slower again. This check reads the text of the
  seven places the fast check now lives and fails the release if any of
  them asks the slow question again. It also refuses one shape by name: a
  test for "is this empty" joined on one line to "is this a list". VBA
  answers the second part even when the first has already said empty, and
  then raises an error. That shape shipped in this item's first build and
  was caught in the live test.
  Tested both ways: it fails that first build in all five places, fails a
  copy with the helper squeezed back onto one line, and fails the code as
  it was before the change.

- **The fourth engine has its name before a line of it exists: `OPTIMIZE`,
  not `SOLVE`.** It is the one that will produce a decision — who works
  which shift, who sits where — from choices, constraints and an objective
  over the same rules the other three read. Nothing of it is built, and the
  README says so. The name was settled now because it goes into your
  formulas the day the first piece ships, and a name that reads as a task
  is one a non-programmer can guess the meaning of, where `SOLVE` named no
  task at all. `OPTIMISE` will work too.

- **Eight sample procedures to start from, in a new `examples/` folder.**
  They are the kind of standard operating procedures people already keep in
  Word, written so Frazaro can run them, easiest first:
  01 Tidy the Sales Export (five sentences), 02 Weekly Sales Summary, 03
  Expense Report Audit, 04 Accounts Receivable Aging, 05 Sales Pivot by
  Region, 06 Month-End Close, 07 Inventory Reorder and 08 Shift Coverage.
  That spans sales, finance, accounting and operations, as `.txt` and
  `.docx` files. All of them run on one workbook, `Frazaro Sample
  Data.xlsx`, which opens on a Start Here sheet, so there is nothing to set
  up. Open it, press **Load Instructions**, pick
  `01 Tidy the Sales Export.txt`, then **Validate Instructions** and
  **Interpret and Run**. `examples/README.md` walks through the rest.

- **A practice procedure, written the way real ones are.** `Practice -
  Expense Reimbursement (as written).docx` has a title, a "Performed by"
  line and steps like "Send the total to payroll", all lines Frazaro cannot
  run. Validate it, then either put `#` in front of each flagged line, which
  keeps it as a note, or rewrite it as an instruction. That is the whole
  process for bringing your own SOP.

- **For painters: `examples/joy.txt`.** Bob Ross's fifteen paints and four
  bases, as colours you can name in any program. Load it and run it for a
  swatch sheet, or copy its `Define` lines to the top of your own procedure
  and write `Set fill-color of range A1:D1 to phthalo-blue.`

- **The README's first example is now a real procedure** (receivables
  aging, from sample 04) instead of a test program, with a pointer to
  `examples/` for anyone new.

- **For contributors: `tools/build_examples.ps1`.** Git cannot show a change
  inside a `.docx` or `.xlsx`, so every sample's text and every row of sample
  data live in this script, and the files are built from it. A rebuild with
  nothing changed is byte-for-byte identical. It needs no Word or Excel, and
  skips any file that is open instead of failing.

- **`OPTIMIZE` measured before it is built (`OPTIMIZE.0`).** Nothing you can
  run has changed. What changed is what the project knows about the engine
  that will build schedules. Its costs were measured on real-sized
  rosters, up to 50 people over four weeks, using the engines that exist.
  Three things came out of it that shape what you will get:
  - **A rule written as a count is thousands of times cheaper** than the same
    rule written as a list of forbidden combinations. "At most 5 shifts a
    week" took about 17 seconds as combinations for one person's week, and
    thousandths of a second as a count. So the sentences that will write
    these rules will only ever write the count.
  - **A long calculation inside a formula cannot be stopped cleanly.** Inside
    a worksheet formula, Excel silently ignores the setting that would let
    Esc stop it politely. It ignores status-bar messages and scheduled
    macros too. So long searches will be a command you run, one that can
    stop when you press Esc and keep the best schedule found so far. A
    formula will only search when the problem is small enough to answer in
    about two seconds.
  - **"No schedule is possible" is often plain arithmetic.** For example, the
    shifts need 63 people a week, and 10 people can work at most 5 each. So
    that is checked first, and said in those words, before any search.
  The full account, including the predictions that turned out wrong, is in
  `docs/OPTIMIZATION.md`.

- **Worth knowing today: Esc during a long Frazaro formula.** The probe above
  found that pressing Esc while a formula's code is running brings up
  Visual Basic's "Code execution has been interrupted" box. We expect the
  same from a long `DATALOG` or `PROLOG` question, though that is not yet
  tested in the installed add-in. If you see it, click **End**. The cell will
  show `#VALUE!` until you change one of the formula's inputs or press
  Ctrl+Alt+F9.

- **For contributors: the `OPTIMIZE.0` tools, all in `tools/`.** None of them
  ships, and none is called by Frazaro.
  - `VLA_DiagO0.bas`: the size ladder and the reference roster's Tables.
    Every answer is checked row by row.
  - `optimize0_expected.ps1`: re-derives every expected count by
    enumeration.
  - `VLA_ProbeO0.bas`: the ten host probes.
  - `optimize0_lp.ps1`: exports a roster for clingo, the answer-set solver
    used by hand as a reference. It never runs clingo and makes no network
    call.
  - `optimize0_live_steps.md`: the steps.

### Known open security items

**Closed this release:** the Uninstall Frazaro data-loss bug above, a
safety fix rather than a listed security item. Standing advice unchanged.
The full list of open items is in `docs/BETA_REARVIEW.md`, in plain words
in `README.md`.

## 0.6.0

### What changed

- **Write a policy in English, then ask questions about it.** This is the
  thing Frazaro exists for, and it starts here. You write the rule in a
  cell, in a sentence, and you ask about it in other cells. Nothing is
  hidden: the rule is text you can read and edit, and every answer
  recalculates when you change it or when the data changes.

  Put your data in Excel Tables — say `Staff` with columns Name, Cert and
  Level, `Shifts` with Shift, Needs and MinLevel, and `Leave` with Name
  and Shift. Then write the policy:

  - `Write in cell H2 that a person can-cover a shift if Shifts lists the
    shift as Shift, the cert as Needs, and the min as MinLevel, and Staff
    lists the person as Name, the cert as Cert, and the level as Level,
    and the level is at least the min, and not Leave lists the person as
    Name and the shift as Shift.`

  That is one rule, and it reads the way the policy is actually stated:
  the shift needs a certification at some level, the person holds that
  certification at some level, their level is at least what the shift
  asks, and they are not on leave for it. **The role nouns are the
  blanks** — "the person", "the shift", "the level" — and saying the same
  noun twice is how you say they must match. The "the" is optional:
  `Staff lists person as Name` means exactly the same thing.

  Then ask, in any cell:

  - `Show in cell E2 who can-cover "Night" by applying the rules in H2:H4
    to the data tables Staff, Shifts, and Leave.`
  - `Show in cell E3 what "Bob" can-cover by applying the rules in H2:H4
    to the data tables Staff, Shifts, and Leave.`
  - `Show in cell E4 whether "Bob" can-cover "Night" by applying the rules
    in H2:H4 to the data tables Staff, Shifts, and Leave.`

  "Who" and "what" spill a list under a header named for the question,
  with each answer listed once, however many ways the rule reaches it;
  "whether" answers TRUE or FALSE.

- **Ask for pairs, or name what you are asking for.** `Show in cell E5 who
  can-cover what by applying …` lists every person and shift together,
  under the headers Who and What. Say `which` and a noun to name a column
  yourself: `Show in cell E6 which person can-cover which shift by applying
  …` spills under Person and Shift. The headers always come from your
  question, never from how the rules happen to be written. A question that
  asks for the same noun twice, or a `whether` that also asks `who`, is
  refused with a note on what to write instead.

- **Facts, for the things that do not deserve a Table.** `Write in cell H7
  that "B3" is flagged.` puts one fact in a cell, and so does `Write in cell
  H8 that "Bob" manages "Carol".` Any rule in the range you ask over can use
  it — a bill an auditor flagged by hand, say, beside the rule that makes a
  big, unapproved bill a violation.

- **Rules about one thing.** `Write in cell H2 that a bill is big if Bills
  lists the bill as Bill and the amount as Amount, and the amount is greater
  than 10000.` Then ask `who is …`, `what is a violation`, `which bill is a
  violation`, or `whether "B4" is a violation`.

- **"Or" is a second rule with the same conclusion.** Write `a bill is a
  violation if the bill is flagged` in another cell, and a violation is now
  either kind. Ask over a range that holds every one of those cells: if a
  question's range holds one and leaves another out, Frazaro stops before
  anything runs and names the cell it left out, instead of quietly
  answering from part of the rule.

- **Quote the names you are asking about.** `"Night"` and `"Bob"` are in
  quotes on purpose. An unquoted word is lowercased on its way in, and
  your Table keeps its own capitals, so `night` would quietly match
  nothing. Quotes are how you say "this exact text". Numbers are the
  opposite: write `3`, not `"3"`. A number cell holds a number, and a quoted
  `"3"` would match it in one kind of question and not in the other, so it
  is refused.

- **Name your Tables as Excel names them.** List them after "to the data
  tables", separated by commas, with a comma before the "and". A name
  that could not stand in a formula as it is — one with a space, a comma
  or a quote in it — is refused before anything is written, rather than
  producing a formula that quietly means something else.

- **Conditions you can write today.** A Table row (`Staff lists the
  person as Name, …`), a relation between two roles (`the person holds
  the cert`), a set (`the bill is big`), a comparison (`is at least`, `is
  at most`, `is greater than`, `is less than`), and `not` in front of any
  of them. Join them with `, and`. A rule can use another rule's
  relation, so a longer policy can be written as several sentences in
  neighbouring cells and asked about as one — name the whole range, like
  `H2:H4`. A fixed value stands wherever a noun can: `Roster lists the
  person as Name and "Ops" as Dept`, `the person reports-to "Alice"`, `the
  amount is greater than 10000`. There is no `the dept is "Ops"`: write the
  value where the dept is read, and Frazaro tells you so if you forget.

- **It tells you when a blank is never filled in.** If you write `the
  level is at least the min` but nothing in the rule ever says where the
  level comes from, Frazaro stops and says so, naming the noun you used,
  instead of answering from a blank. This grammar's own words — `is`,
  `lists` and the rest — are refused as relation names for the same
  reason, and so is any name starting `vla-`: Frazaro writes a few helper
  rules of its own under that prefix, and a name of yours must never be
  mistaken for one of them. The same goes for what a rule concludes: `a
  person can-cover a shift if Staff lists the person as Name` never says
  which shift, so it would be true of every shift there is — it is refused,
  naming the shift.

- **A misspelled relation is named, in every kind of question.** `Show in
  cell E2 who can-drive "Night" …`, when your rules only ever say
  `can-cover`, used to show the header with nothing under it — which looks
  exactly like "nobody". It now says `'can-drive' is used in a rule, but
  nothing defines it`, the way a "whether" question already did. The same
  check covers every `=DATALOG(...)` formula: a rule that reads a relation
  no fact, no rule and no Table defines is refused by name before anything
  is worked out — including under `not`, `count` and `sum`, where the old
  silence was worse than empty: a misspelled `not` let everyone through,
  and a misspelled `count` answered 0. **Every rule in the text is
  checked, not only the ones your question uses**, so pass every Table the
  rules name, or keep rules that read different Tables in different ranges.
  A Table with no rows is fine: it is still a Table you passed.

- **`=DATALOG(...)` answers a yes-or-no question with TRUE or FALSE.** Write
  the fact you are asking about as the query, every value filled in:
  `=DATALOG("(rule (route X Y) (links X Y)) (rule (route X Y) (links X Z)
  (route Z Y)) (query (route ""A"" ""D""))", Links)` shows TRUE if A can
  reach D and FALSE if it cannot. It is a real TRUE or FALSE, so `=IF(...)`
  and conditional formatting read it. It finishes even when the data loops
  back on itself — A to B to C and back to A — where `=PROLOG(...)` stops
  with a refusal. The fact is checked exactly the way a rule checks it, so a
  number is written bare (`3`) and text must match the Table's capitals. To
  list answers instead, query a relation by name, as before. Three shapes
  are refused, each with a note on what to write instead: a query with a
  blank still to fill in, like `(query (route "A" Where))`; one that names
  Table columns, like `(query (links (from "A") (to "D")))`; and
  `(headless)` beside a TRUE-or-FALSE question.

- **`=DATALOG(...)` can also ask "is there nothing that…?"** Put `not`
  around the fact you are asking about, and leave a blank for anything
  that may be any value: `=DATALOG("… (query (not (route ""A"" X)))",
  Links)` is TRUE when A reaches nowhere and FALSE when it reaches
  somewhere. A blank written twice must be the same value, so `(query (not
  (link X X)))` asks whether nothing links to itself. This is how "does
  every shift have someone to cover it?" is asked: write a rule for a shift
  nobody covers, then ask whether there is no such shift. It is a real TRUE
  or FALSE, and it finishes over data that loops. A blank outside `not`,
  like `(query (route "A" X))`, is still refused with a note on how to list
  the answers instead, and a `not` with nothing, or with a bare name, under
  it now says what is wrong rather than naming something you did not write.

- **Say "directly or not" to follow a relation any number of steps.**
  `Write in cell H3 that a person is-under a boss if the person reports-to
  the boss directly or not.` means the person reports to the boss, or to
  someone who does, and so on up the chain. You can say it in a question
  too: `Show in cell E2 who reports-to "Alice" directly or not by applying
  the rules in H2:H2 to the data tables Reports.` lists everyone under
  Alice, while the same question without the words lists only the people
  who report to her directly. It finishes even where the data loops back on
  itself, like routes from A to B to C and back to A, and it follows a chain
  of any length. Frazaro writes the two rules that do this into the rule's
  cell for you, named after the relation (`vla-any-reports--to`).

- **Every "whether" question is now answered by DATALOG.** The answers are
  the same TRUE or FALSE, and a question over data that loops now finishes.
  One thing behaves differently: DATALOG checks every rule in the range, so
  if one rule names a relation nothing defines, a "whether" question is
  refused by name even when it never uses that rule — where it used to
  answer. Fix the spelling, or keep rules that do not belong together in
  different ranges.

- **Two kinds of rule are refused before anything runs.** A rule that asks
  its own relation first — `a person is-under a boss if the person is-under
  the middle, and the middle reports-to the boss` — would never finish, so
  it is refused with the fix: put that condition last. And a relation named
  after the Table it reads, like `a source feeds a target if Feeds lists …`,
  is that Table, so the rule only reads itself: it is refused, whether the
  two names meet in one sentence or in two cells. Give the relation a name
  of its own. "directly or not" after a set, a comparison or a Table row,
  or in what a rule concludes, is refused too, with a note on where it
  belongs.

- **Ask how many.** `Show in cell E2 how many people can-cover "Night" by
  applying the rules in H2:H4 to the data tables Staff, Shifts, and Leave.`
  puts one number in the cell, and counts each person once, however many
  ways the rules reach them. Put the value first to count the other side:
  `how many shifts "Bob" can-cover`. Add `directly or not` to count along a
  chain. When nobody matches, the cell shows 0.

- **Count for each member of a set, the empty ones included.** Write the set
  once, like `Write in cell H3 that a shift is listed if Shifts lists the
  shift as Shift.`, then `Show in cell E2 how many people can-cover each shift
  that is listed by applying …` spills Shift and People side by side, with a
  0 for a shift nobody can cover rather than leaving that shift out.

- **Ask what is missing, and whether anything is.** With a set for what counts
  and a set for what is covered, `Show in cell E2 which shift that is listed
  is not covered by applying …` lists the shifts nobody covers, and `Show in
  cell E3 whether every shift that is listed is covered by applying …`
  answers TRUE or FALSE — TRUE when there are no shifts at all. The set goes
  after `that is`, and a plural reads the same: `which shifts that are listed
  are not covered`.

- **Ask who alone.** `Show in cell E2 who alone can-cover "Night" by applying
  …` lists the one person who can, and nobody when two can or none can.
  `whether "Bob" alone can-cover "Night"` answers TRUE or FALSE. Because of
  this, `alone` can no longer name a relation or a set. "Name one person
  who can…" is not offered: which one it named would change with the order
  of your rows.

- **Test text in a rule.** `Write in cell H2 that a code is revenue if
  Accounts lists the code as Code, and the code starts with "GL-4".` keeps
  the codes that begin GL-4. `ends with` and `contains` work the same way,
  and `not` in front keeps the rest. Capitals must match exactly, and a
  number cell is read as its digits, so 4010 starts with "40". The code has
  to be found first, by a Table row or a relation: a text test only checks
  it. The other side may be a noun as well, like `the code starts with the
  prefix`. Because of this, `contains` can no longer name a relation — call
  it `has-part`, or another name.

- **Choose the first tier that fits, with "otherwise".** `Write in cell H2
  that a customer has-tier "Gold" if Customers lists the customer as
  Customer and the spend as Spend, and the spend is at least 10000,
  otherwise "Silver" if …, otherwise "Bronze" if Customers lists the
  customer as Customer.` gives each customer one tier, the first that fits,
  even when a customer has several rows. Each value after "otherwise" needs
  its own "if", and a value may be a noun its conditions find as well as a
  fixed one. Keep the whole choice in that one cell: writing `has-tier`
  again in another cell is refused, since its answers would stand beside
  the chosen one.

- **Put a list in one cell.** `Show in cell E2 who can-cover "Day" as one
  list by applying …` puts `Ann, Bob, Ed` in one cell, each name once, in
  the order the rows were found. `Show in cell E3 which people can-cover
  each shift that is listed as one list by applying …` spills each shift
  beside its list, with an empty cell for a shift nobody can cover. "As one
  list" follows a question that asks for one thing, or an "each"; after
  "whether", "how many", "alone" and the other shapes it is refused by name.

- **`=DATALOG(...)` and `=PROLOG(...)` can test text.** In a rule,
  `(text-starts-with Code "GL-4")` keeps the codes that start with GL-4, and
  `(text-ends-with Code "10")` and `(text-contains Name "an")` do what they
  say. Each keeps a row once, however often the text occurs, capitals must
  match exactly, and a number cell is read as its digits, so a code of 4010
  starts with 40. Put `not` round one to keep the rows it rejects. Both
  engines read the same rule the same way, so one rules range can feed
  either. These three names can no longer name a relation, a fact or a
  Table.

- **`=DATALOG(...)` can put a list in one cell.** `(textjoin People ", "
  (can-cover Person Shift))` in a rule joins everyone who can cover each
  shift into one cell — `Ann, Bob, Ed` — in the order the rows were found,
  each name once, and an empty cell for a shift nobody covers. The
  separator is whatever you write between the quotes. A list longer than a
  cell can hold is refused by name.

- **DATALOG says when you have written PROLOG.** `sub-atom`, `if`,
  `findall`, `=<`, a cut and the rest of PROLOG's own words used to be
  refused as a misspelling, as "nesting" or as "mixed keying". Now DATALOG
  names the word as PROLOG's and points at its own way of saying it — for
  `sub-atom`, the text tests above. A relation you named yourself, say
  `member`, is untouched.

- **Known limits, said plainly.** A relation's name is one word, so it is
  `can-cover`, with the hyphen, not `can cover`. Relations you declare in
  words, like `is submitted by`, are planned rather than guessed at, since a
  guess about where a relation's words end would have to be honoured
  forever. Questions are asked one at a time. A quoted name must match the
  data's capitals exactly: `"night"` finds nothing where the Table says
  `Night`, and says nothing about it. The range check sees only rules
  written by the same program: rules typed straight into cells, or written
  by another program, are not checked, and nor is a range on another sheet.
  A number kept as text in a column cannot be matched, since quoted numbers
  are refused. Adding up quantities along a chain ("how many spokes go into
  one bike") and putting things in order ("which course to take first")
  cannot be said yet. A list separates its values with a comma and a space,
  so a value that holds one reads as two. How large a Table a question can
  read has now been measured, on one machine: a few hundred rows answer in
  well under a second, a thousand rows take two to three seconds, and ten
  thousand take half a minute, during which Excel stops responding. A rule
  written with "otherwise" costs about five times an ordinary one, because it
  becomes five rules. "Directly or not" costs more again, because it works
  out every pair in the whole chain before narrowing to the person you asked
  about: a hundred people in one reporting line takes twenty seconds. Asking
  the same question again after changing one cell costs what it cost the
  first time, since nothing is remembered between answers — though changing a
  cell the question does not read costs nothing at all.

- **Lint VLA can no longer break a file it rewrites.** It used to save
  through the system's legacy code page, so a `£` in a comment came back as
  a byte that is not valid UTF-8. It now saves UTF-8, keeps a file's own
  byte-order mark and line breaks, and relinting a clean file changes
  nothing. It will not touch a generated file like `english_expanded.vla`,
  a file that is not UTF-8, or one whose reformatting would change more
  than spaces and line breaks (a `;` comment inside a form, which it would
  drop) — each is refused by name, with the line. Linting a folder shows
  what it will rewrite, and what it refused and why, before it changes
  anything. Export Expanded Phrasebook and Phrasebook Test Coverage save
  UTF-8 too.

- **Check Instructions clears every old mark.** If you deleted the text of
  a program's last sentence, or the whole program, its old OK or error mark
  stayed beside the empty row after the next Check. Now every mark from the
  previous Check is cleared first. And a program you have emptied gets the
  "Nothing to check" message instead of a "Subscript out of range" error in
  C1; Run, Interpret, Export and Show VBA each say there is nothing to do,
  too.

- **A new program tab starts in B1**, the first sentence cell, ready to type
  into, instead of in the hidden column A.

- **An empty rules text is named, not a raw error.** `=DATALOG("")` or
  `=PROLOG("")`, a formula reading a blank cell, or rules that hold only
  comments used to show "Subscript out of range". Each engine now says the
  text is empty and that a program needs at least a query.

- **A program with no sentences no longer fails.** A program of only
  comments, or a file with nothing in it, used to stop with "Subscript out
  of range" when checked, run or translated. It now translates to an empty
  program.

- **Translate to VLA marks its output, and keeps every character.** The
  `.vla` it writes now starts with a `; GENERATED` line naming the program
  it came from, so Lint VLA leaves it alone, and it is saved as UTF-8, so a
  sentence like `Put "café" into cell A1.` keeps its `é`.

- **Undo Last Run puts back the right sheet, and a Run it cannot undo no
  longer starts.** Before a Run changes anything, Frazaro saves a hidden
  copy of each sheet it can change, so Undo Last Run can put it back. When
  a program used a sheet by name (`Work on sheet Data.`, say), that copy
  was left visible as a tab like `Data (2)`, and Undo would have put the
  sheet back with the wrong contents, often blank. No error was shown.
  That is fixed. Separately, if saving a copy failed, the Run went ahead
  anyway, without Undo and without telling you. Now it stops before its
  first sentence, says which sheet could not be saved and why, and
  removes anything it had made for Undo. And a sentence that merely
  mentions a name no sheet can have, like `"see 'Q1/Q2'!A1"`, no longer
  leaves a stray `Sheet` tab each time it runs.

- **Undo Last Run holds up in three more cases.**
  - A Run that is refused because Frazaro could not save its copies no
    longer costs you the Undo you had: Undo Last Run still puts back the
    Run before it. Before, that Undo was already gone.
  - Changing only the capitals in a program's tab name (`Frazaro (sales)`
    to `Frazaro (Sales)`) no longer makes every later Run refuse, and Undo
    still finds that program's copies. Add Program now also refuses a name
    that differs from an existing program's only in capitals or
    punctuation (`q1a` beside `Q1-A`), as it already refused two names that
    shorten to the same thing.
  - With a chart sheet among the tabs, Undo could put back the wrong sheet
    or leave one renamed `VLAu_old`. It now puts back the right sheet, and
    if it cannot, it leaves that sheet as the Run left it.

- **A new release check: `tools/check_ptrsafe_declares.ps1`.** Frazaro calls
  into Windows in exactly one place — the four window calls behind the CLI's
  resizable frame — and a declaration like that is the one kind of VBA whose
  correctness depends on which Office compiles it. Written for 32-bit
  Office, it stops 64-bit Office from compiling the module at all, and
  64-bit is most installations today; written for 64-bit Office, it is
  refused by Office 2007 and older; and a window handle typed as a 32-bit
  number compiles on both, then quietly loses half its value on 64-bit. No
  compile and no test catches all three, because any one machine compiles
  only the branch written for it. This check reads the text instead: for
  nine Office versions it works out which of them compiles each
  declaration, and fails the release if any of them would refuse it or read
  it wrongly. It also holds that nothing Frazaro generates — no phrasebook
  template, no compiled module — writes such a declaration itself.
  Mutation-tested in both directions: twenty planted faults, each caught by
  the rule it was planted against, and three correct declarations that stay
  clean.

### Known open security items

**Closed this release:** none — 0.6.0 is a feature release. Four changes
touch the security surface without closing an item, listed so none is a
surprise:

- Sentences can now write `=DATALOG(...)` and `=PROLOG(...)` formulas into
  cells, and those formulas recalculate with the workbook. Both engines
  read only two things: the rules text in the cells a question names, and
  the Tables it names — no files, no network, nothing outside the
  workbook. The rules are ordinary cell text, so anyone who can edit those
  cells can change every answer that reads them; protect the sheet if that
  matters, as you would the inputs to any formula.
- The Table names a question lists are written into its formula, so each
  one is first held to Excel's own rule for a Table name. A quoted name
  carrying a comma, a quote or a bracket is refused by name, instead of
  becoming extra arguments in a formula you did not write.
- A quoted name inside a question is written into its formula with every
  quote mark doubled, so a quote inside the name stays part of the name
  rather than ending the formula's text early.
- Writing a formula is not new, and `SEC.15` below still applies:
  formulas a program writes are not screened for functions that reach the
  network. The formulas these sentences write contain only `DATALOG`,
  `PROLOG` and `TEXTJOIN`.

**Still open:** `SEC.3`, `SEC.7`, and two from the 2026-09-08 code
review — `SEC.10` and `SEC.15`. In plain words:
the remembered raw-VBA consent record still lives inside the workbook
(`SEC.10`);
formulas a program writes are not screened for functions that reach the
network (`SEC.15`); and effects like sending mail still run without a
permission prompt (`SEC.7`). `SEC.8` narrows that last one — it gates on
where the *workbook* came from — but does not close it: a phrasebook loaded
into a workbook of your own still reaches those verbs unprompted.

**Assessed and accepted, not fixed:** `SEC.12`, `SEC.14`, `SEC.16` and
`SEC.17`. Each needs a precondition an ordinary install does not meet —
mostly an Excel setting that ships off and that Frazaro never asks you to
turn on. The reasoning for each, and what would reopen it, is written down
rather than left implied.

The authoritative lists, kept current in one place instead of copied into
every release: [`README.md`](../README.md) in plain words, and
[`docs/BETA_REARVIEW.md`](BETA_REARVIEW.md) with the file, line, fix and
disposition for each.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from — and treat a workbook someone sent you the
same way before you press Interpret.** Frazaro makes no network call and
does not update itself; check the README's *Known open security items*
when you return for a newer build. Vulnerability reports:
`docs/SECURITY.md`. Everything else: `docs/SUPPORT.md`.

## 0.5.6

### What changed

- **Sorting, saying whether there is a header row.** Name the column by
  its letter, and say whether the range starts with a header row that
  should stay put:

  - `Sort range A1:C50 by column B with a header row.` — row 1 stays on
    top, the rest is sorted. `Sort range A2:C50 by column B without a
    header row.` sorts every row you named.
  - `descending` (or `ascending`) goes after the column: `Sort range A1:C50
    by column B descending with a header row.`
  - **Two columns:** `Sort range A1:C50 by column B then by column D with a
    header row.` — and to set their directions, name both: `… by column B
    descending then by column D ascending …`.
  - **The whole sheet:** `Sort this sheet by column B with a header row.`

  A sort that doesn't say "with" or "without" stops before it runs and asks.

- **Filtering.** Turn the filter buttons on, show only the rows you want,
  and put everything back:

  - `Add filters to range A1:D50.` puts the filter buttons on the header
    row. Saying it again leaves them on.
  - `Filter range A1:D50 to show rows where column C is "West".` — or
    `contains "we"`, or `is greater than 100`, or `is less than 100`.
    Rows that don't match are hidden, not deleted. A second condition on
    another column narrows the first: West rows *and* over 100.
  - `is 100` matches the number, however the cell shows it — `$100.00`
    included — and `"West"` matches West exactly, not Western. A `*` or
    `?` in what you type is just a character.
  - `Clear the filter conditions.` shows every row again and keeps the
    buttons. `Remove the filters.` takes the buttons away too.
  - `Copy only the visible cells of range A1:D50 to F1.` copies what a
    filter leaves showing — rows and columns that are hidden stay behind.

  Filters work one set per sheet, the way Excel's do: adding them to a
  second range while the first still has them stops and says so, and so
  does filtering an empty range. A filter sentence that stops leaves the
  sheet exactly as it was. In a sort or a filter, a column that isn't part
  of the range — `by column F` on `A1:C50` — stops and names the column.

- **Three sentences you may already use now have clearer twins.** They
  keep working exactly as before. `Sort range A1:C50 by column B1.` always
  keeps row 1 where it is, even when row 1 is data — the new sentences say
  `with a header row` or `without`. `Keep only rows of range A1:D50 where
  column 3 is "West".` counts columns from the start of the range and
  hides the other rows rather than deleting them — `Filter range … to show
  rows where column C is …` names the column by its letter and says it
  shows. `Show all rows.` clears filter conditions and does not unhide rows
  hidden with `Hide row` — `Clear the filter conditions.` says which it is.

- **Number formats, for a cell or a range.** Money, percentages, plain
  numbers, dates and times, each in one sentence:

  - **Money, with the currency named:** `Format range B2:B9 as dollars.` —
    or `euros`, or `pounds`. And Excel's accounting layout, with the sign
    at the left and negatives in brackets: `Format range B2:B9 as
    accounting in dollars.`
  - **Percentages and plain numbers:** `Format range C2:C9 as percent.`,
    `Format range D2:D9 as a number.`, and `… as a number with thousands
    separators.`
  - **Any of those with a set number of decimals:** `Format range B2:B9 as
    euros with 0 decimals.`, `Format cell C2 as percent with 1 decimal.`,
    `Format range D2:D9 as a number with 0 decimals and thousands
    separators.` (0 to 30 decimals.)
  - **Dates and times:** `Format range A2:A9 as a short date.` — or `a long
    date`, or `an ISO date` (2025-12-09) — and `Format range E2:E9 as a
    time.`
  - **Back to plain, or anything else:** `Format range A1:A9 as general.`
    clears the format; `Format range A1:A9 using "00000".` takes any Excel
    format code, written the way Excel's own Format Cells box writes it.
  - **Text, for what you type next:** `Format range A2:A9 as text for new
    entries.` Anything typed there afterwards stays exactly as typed —
    `007` keeps its zeros. A number already in the cell stays a number,
    which is what Excel's own Text format does, and the sentence says so.

- **Formats follow the person reading, and the currency follows the
  numbers.** The same workbook opened on computers set up for different
  countries shows dates in each country's own order and uses each
  country's own decimal point and thousands separator — so `as a short
  date` is 09/12/2025 on a UK machine and 12/9/2025 on a US one, both
  correct. The currency never changes that way: dollars stay dollars
  wherever the file is opened, because the currency is a fact about the
  numbers, not the computer. That is why the sentence names it.

  The older `Format cell A1 as currency.` and `… as date.` still work
  exactly as before, and it is worth knowing what "before" was: `as
  currency` has always shown *the currency of the computer opening the
  file* — a $ in the US, a £ in the UK, on the same figures. If your
  numbers are in one currency, say which: `as dollars`. `as date` keeps
  month, day, year in that order everywhere; `as a short date` follows
  the reader.

- **Formatting sentences now reach a whole range, not just one cell.**
  Five sentences only ever took a single cell — `Make cell A1 bold.` would
  refuse `A1:C3`. Each now has a range version, worded the same way with
  `range` in place of `cell`:

  - `Make range A1:C1 bold.` (and `italic`)
  - `Make range A1:C1 blue.` — any of the eight named colours
  - `Set font-color of range A1:C1 to "#FF0000".`
  - `Set fill-color of range A1:C1 to hot-pink.`
  - `Set font size of range A1:C1 to 14.`

  The cell versions are unchanged.

- **Twelve new ways to make a sheet look right.** Each new sentence
  below takes a cell or a range — `cell A1` or `range A1:C3`:

  - **Undo bold and italic:** `Make range A1:C1 not bold.`,
    `Make cell A1 not italic.`
  - **Underline and strike through:** `Underline range A1:C1.`,
    `Strike through cell A1.` — and the one sentence that takes marks off
    again: `Remove underline from …`, `Remove strikethrough from …`,
    `Remove borders from …`.
  - **Typeface:** `Set font of range A1:C1 to "Courier New".` Put a font
    name in quotes; a name with a space in it has to be.
  - **Vertical alignment:** `Align range A1:C1 to the top.` — or `middle`,
    or `bottom`. (Left, right and centre were already there.)
  - **Indent and rotate:** `Indent range A1:C1 by 2.` (0 to 15 levels),
    `Rotate text in range A1:C1 by 45 degrees.` (−90 to 90).
  - **Borders, three ways, each saying which lines it draws:**
    `Add a border around range B2:D4.` draws the outside box only;
    `Add a bottom border to range B5:D5.` draws one edge (`top`,
    `bottom`, `left` or `right`); `Add borders to every cell in range
    A1:C3.` draws every line, inside and out.
  - **Clear a fill:** `Clear fill-color of range A1:C1.`

- **Border colour, as part of drawing the border.** Say `colored` and a
  colour right after the word border: `Add a border colored red around
  range B2:D4.`, `Add a bottom border colored green to range B5:D5.`, `Add
  borders colored "#0000FF" to every cell in range A1:C3.` Any of the eight
  colour names works as it is — red, yellow, black, blue, cyan, green,
  magenta, white — and any other colour as a code in quotes, like
  `"#FF69B4"`; a name you have defined yourself, like `hot-pink`, goes in
  quotes here too. Each sentence colours only the lines it draws — the box
  around a range does not draw its inside lines.

- **Two sentences you may already use now have clearer twins.** `Add
  border to range A1:C3.` draws a line around *every cell*, though it reads
  like it draws one box around the range; `Clear color of cell A1.` clears
  the *fill* and leaves the text colour alone, though "color" could mean
  either. Both keep working exactly as before — nothing you have written
  changes. The new spellings say what they do: `Add borders to every cell
  in …` and `Clear fill-color of …`, and they are the ones the examples use
  from now on. This is a new standing rule for every sentence Frazaro adds:
  **where Excel has two neighbouring operations, the sentence names the one
  it performs.** (A plain "Set border-color of …" was held back under it:
  Excel's own way of doing that also draws lines you did not ask for.
  Border colour arrives instead as part of drawing the border — see above.)

- **Colour names in quotes now work wherever a colour code does.** `Set
  font-color of cell A2 to "red".` used to stop with *"'red' is not a
  color"*; it now means red, like any of the eight names. Nothing that
  worked before changes.

- **Six colours now work under Interpret.** `Make cell C4 black.` — and
  `blue`, `cyan`, `green`, `magenta` and `white` — has always translated
  and run with Run, but stopped with *"there is nothing stored at key
  'vbblack'"* under Interpret; only `red` and `yellow` worked there. All
  eight now work both ways.

- **Interpret can set five more formatting properties, each reviewed by
  name.** Interpret only touches the parts of Excel someone has read and
  listed, and refuses everything else in words. This release adds five —
  underline, strikethrough, vertical alignment, indent and text rotation —
  each of which changes how a cell looks and nothing else. No new
  *action* was added: the outline border is drawn as four single-edge
  lines rather than through a new Excel command.

  Sorting and filtering add four more, also by name: taking a sheet's
  filter buttons away (it can only take them away, never add them), and
  three that only find cells already on the sheet — a column of a range,
  the part of the sheet in use, and the visible cells of a range. Sorting
  and filtering themselves were already on the list; they now also take
  a second sort column and a second filter condition.

- **PROLOG can work with text.** A spreadsheet's whole subject is cell
  values, and until now PROLOG could not take one apart or put two
  together. Seven goals do that now:

  - `(atom-length Text N)` — how many characters. `(atom-length "hello" N)`
    gives `5`.
  - `(atom-concat A B Whole)` — join two pieces: `(atom-concat "Item " 42 X)`
    gives `Item 42`. Given the whole, it takes it apart instead:
    `(atom-concat "ID-" Rest "ID-42")` gives `Rest` = `42`,
    `(atom-concat Stem ".xlsx" "report.xlsx")` gives `report`, and with
    both halves left blank it lists every way to split.
  - `(sub-atom Text Before Length After Part)` — any piece, by position or
    by content. `(sub-atom "Frazaro" 2 3 After Part)` gives `aza`;
    `(sub-atom Email Before 1 After "@")` finds where the `@` is. Positions
    count from 0, the way Prolog counts them.
  - `(atom-number Text N)` — text to a number and back.
    `(atom-number "42" N)` gives the number `42`; text that is not a number
    simply does not match, so it doubles as the test "is this a number?".
  - `(upcase-atom Text Upper)` and `(downcase-atom Text Lower)` — case.
  - `(atomic-list-concat Parts Separator Whole)` — split and join.
    `(atomic-list-concat P "," "a,b,c")` gives the list `(list "a" "b" "c")`,
    ready for `length`, `member` and the other list goals; run the other
    way it joins a list. `(atomic-list-concat (list First Last) " "
    "Ada Lovelace")` fills in both names at once.

  The names are real Prolog's, with a hyphen where Prolog writes an
  underscore, the way `sum-list` already is. Type the underscore version
  — `atom_length`, `sub_atom` — and Frazaro tells you the hyphenated one
  instead of quietly finding nothing.

- **What a text goal hands back is text.** The result behaves exactly
  like the value of a text cell: `(atom-concat 4 2 X)` makes the text
  `"42"`, not the number 42, and to compare a result with something you
  typed, quote it — `(== X "abc")`. Compare it with an unquoted name and
  Frazaro stops and explains the difference rather than quietly answering
  no. The values you *give* a text goal are read as text whichever way
  they were written, so `(atom-concat ab c abc)` is true and
  `(downcase-atom "ENG" eng)` is true.

- **Changing case gives the same answer on every computer.** The usual
  way to change case in Excel's own language depends on the Windows
  language settings — a Turkish Windows lowercases `I` differently — so a
  cell's answer would have depended on whose machine computed it.
  `upcase-atom` and `downcase-atom` instead carry their own table: A–Z and
  the accented letters of Western and Central European languages,
  Spanish's `ñ`, `á` and `ü` among them, and Polish, Czech and Turkish
  letters too. Letters from other alphabets — Greek, Cyrillic and so on —
  are refused by name rather than passed through unchanged and looking
  finished. Characters with no case at all, like digits or `€`, pass
  through as they are. Four characters whose case is genuinely disputed
  — the micro sign `µ`, the Turkish dotless `ı` and dotted `İ`, and the
  old long `ſ` — are refused in the one direction where the authorities
  disagree.

- **Emoji and counting.** Excel counts an emoji as two characters and
  Prolog counts it as one. The goals that count or cut at a position —
  `atom-length`, `sub-atom`, and `atom-concat` when it lists every split —
  refuse such text and say why, rather than silently pick one of the two
  answers. The goals that do not count, like changing case, pass emoji
  through untouched.

- **Taking a long text apart in every possible way has a limit, and says
  so.** `(sub-atom Text B L A S)` with only the text given lists every
  piece of it: 105 pieces for a 13-character text, 120 for 14, and so on
  up to 99,681 for 445 characters, which is as much as one query may try.
  Past that it is refused by name *before* it starts, with a message
  saying to fill in more of the arguments. Anything more specific — a position, a length, the piece
  you are looking for, or a known start or end for `atom-concat` — is
  answered directly and works on a text of any length.

- **Every number PROLOG writes is now written one way.** Until now a
  number could get two different spellings depending on where it came
  from. The result of `(is ...)` followed the computer's regional
  settings, so on a Windows set to write `0,5` the next calculation
  refused it as not a number. A table cell holding `0.5`, or a total from
  `sum-list`, was written `.5` without its leading zero — and so it never
  matched a `0.5` you typed: `(sum-list (list 0.25 0.25) 0.5)` answered
  *no*. Both now read `0.5` on every machine.

- **`(length "hello" N)` now points you at `atom-length`.** It is the
  first thing most people type. It is still refused — a piece of text is
  not a list — but the message now names the goal you wanted, and
  `(append "ab" "c" X)` names `atom-concat` the same way.

- **`(whole 3)` now points at `(whole? 3)`.** It used to be an unknown
  name that quietly found nothing. The near-miss spellings Frazaro
  recognises — a hyphen for an underscore, a missing question mark — are
  now worked out automatically from the full list of names on every
  release, which is how this one was found missing.

- **PROLOG refuses Prolog's side-effect goals, for good, and says why.**
  `write`, `assert`, `random` and the rest of that family used to be
  unknown names here, and until this release an unknown name quietly found nothing — so the
  line everyone types while debugging, `(query (p X) (write X))`, answered
  with an empty column and looked like "no match". Each is now refused
  with a message giving the reason and what to do instead:

  - **Changing the program while it answers** — `assert`, `asserta`,
    `assertz`, `retract`, `retractall`, `abolish`. Excel recalculates a
    formula whenever it chooses, as often as it chooses, so an answer that
    depended on how many times it had run would be wrong in a way no one
    could see. To collect answers use `(findall X Goal Bag)`; to count
    them, `(length Bag N)`.
  - **Printing** — `write`, `writeln`, `print`, `nl`, `format`, `writeq`,
    `write_canonical`, `write_term`. A formula in a cell has nowhere to
    print; what `PROLOG` returns *is* its output. Put the variable in the
    query and it fills a column, one row per answer.
  - **Remembering a value between goals** — `b_setval`, `b_getval`,
    `nb_setval`, `nb_getval`, `gensym`. In Prolog a value kept this
    way outlives the query, and a formula must not carry anything from
    one recalculation to the next — a counter would count how often
    Excel had recalculated. Pass the value along as an argument instead.
  - **The clock and the dice** — `random`, `random_between`,
    `random_member`, `random_permutation`, `get_time`, and `random`,
    `random_float` and `cputime` inside arithmetic, as in
    `(is X (random 10))`. Their answer would change every time Excel
    recalculated. Excel's own `RAND()`, `RANDBETWEEN()` and `NOW()` do
    this in the open: put one in a table you pass to `PROLOG`.
  - **Reading from outside** — `read`, `read_term`, `consult`, `halt`.
    `PROLOG` reads only its own program and the tables you give it.

  Written bare, the way Prolog writes them — `nl`, `halt` — they are
  refused the same way, and the hyphenated spellings (`random-between`,
  `get-time`, …) are refused too. A program using `format` with one
  argument in one place and two in another is told about `format`, not
  about its arities. A rule that contains one of these goals is refused
  only when it actually runs, so a rule you never call does no harm.
  None of these is waiting to be built.

- **More reserved words.** Fifty-four names join the reserved set: the
  seven text goals, their seven underscore spellings, `whole`, the
  twenty-eight goals above and eleven hyphenated spellings of them. If a
  knowledge base of yours uses one of those as a predicate name it will
  need renaming; `write`, `read`, `print`, `format` and `whole` are
  the plausible ones. `tab` and `flag` are deliberately **not** reserved:
  a sheet tab and a flagged order are ordinary things to keep facts about.
  The same list applies to the name of a table you pass to `PROLOG`, in
  any capitalisation: a Table named `Write` or `Between` is refused as
  a reserved word rather than loaded, so rename it.

- **A predicate nothing defines now stops the formula and says so.** A
  misspelled name used to find nothing, quietly:
  `=PROLOG("(fact (parent tom bob)) (query (parnet tom X))")` answered
  with an empty column and looked like "no match". Now it says
  `'parnet' is called by this query` but nothing defines it, and where a
  definition could come from — a `(fact ...)`, a `(rule ...)` or a table
  you pass in. The silence was worst where it was wrong rather than
  empty: `(not (parnet tom bob))` answered TRUE, and a `findall` over a
  misspelled goal collected nothing and counted it as zero. Both now stop.

  PROLOG checks this before it answers anything, and only for what your
  query can reach — the predicates it calls, and the ones those rules
  call. A rule nothing calls is not examined, so one cell of rules can
  serve several formulas that pass different tables. Inside what the
  query can reach, it is checked even on a branch today's data never
  takes — the else-part of an `(if ...)` that never runs — because that
  formula would otherwise break the day the data changed. A table you
  pass that has no rows is not undefined: it is there and empty, and a
  query over it returns no rows, as before.

  If a program of yours used an undefined name on purpose, as a goal
  that always fails, write one that matches nothing instead: `(= a b)`
  fails every time.

- **PROLOG answers questions about real-sized Tables.** Until now a query
  could try only 120 things in all, and every row it looked at counted, so
  a question about a Table of more than about 120 rows stopped. So did
  "who reports to Alice, directly or not" over a Table of six. The message
  also blamed your rule for going round in circles, when the rule was
  fine. There are now separate limits for separate things, and each says
  what actually happened. Nothing you have already written changes
  meaning: the programs that answered before answer the same way, and the
  only ones whose answer changes are those that used to be refused.

  - **How deep a query goes: 120 rules deep**, the same safety limit as
    before. A rule that calls itself forever still stops here, and so does
    data that loops back on itself (A reports to B, who reports to A). The
    message now also names the third cause, a chain longer than PROLOG
    follows, and says that a DATALOG question follows a chain of any
    length.
  - **How much a query tries: 100,000 facts and rules.** This is a limit on
    SEARCHING, not on the size of your data, and the two are not the same:
    reading a Table spends one try per row, so a plain question can read
    100,000 rows, while a join of two Tables spends one try per PAIR, so two
    Tables of a few hundred rows each is also about the limit. Past it, the
    message says the query is trying too much — never that your rule is
    wrong — and suggests putting a known value where the question has an
    unknown, or asking it as a DATALOG question, which joins large Tables
    without trying every pair.

- **Longer lists, up to a thousand.** `findall` gathers up to 1,000 things
  into one list, and `length`, `member`, `==` and the rest work on the
  whole of it — where before, a list could not outlast the 120-step
  ceiling. Past a thousand it stops and says so, and says what to ask
  instead: reading a Table of any size is fine, holding all of it inside
  one value is not, and a DATALOG question counts a set of any size. (The
  limit is Excel's own: a list is a chain, and letting go of a very long
  one runs the host out of stack space. A thousand is eight times under
  where that was measured to start.) An answer too long to show in one
  cell (Excel allows 32,767 characters) is refused by name, saying which
  answer and how long, rather than left for Excel to show as something
  else. `(between 1 99999 X)` now generates all 99,999 values.

- **Cut in parentheses is refused.** Cut is written on its own, between
  the goals: `(rule (first X) (p X) !)`. Written `(!)`, it used to be a
  goal nothing could run, which quietly failed — so the rule never
  answered. It now stops and says how to write it.

- **A spelling mistake is reported as a spelling mistake.** Using a
  Prolog spelling at two sizes — `(atom X)` in one place and
  `(atom X Y)` in another, or `(integer? X)` beside `(integer? X Y)`
  — used to be blamed on the arguments: "used with 1 argument(s) in one
  place and 2 in another". It now gets the message it gets when written
  once: that `(atom ...)` is written `(atom? ...)`, or that
  `integer?` is not available and what to ask instead.

- **Two short documents for whoever decides whether Frazaro is
  allowed.** [`docs/IT_REVIEW.md`](IT_REVIEW.md) gives an IT reviewer two
  pages: what Frazaro installs and writes, including what uninstalling
  leaves behind; everything it can reach; what it sends over a network
  (nothing of its own); how it updates; and how it is removed.
  [`docs/COMPARISON.md`](COMPARISON.md) sets Frazaro beside Copilot in
  Excel, Office Scripts and Python in Excel on one page, including what
  each of those does better, and every claim about them quotes
  Microsoft's own documentation. A new check,
  `tools/check_no_network.ps1`, runs on every change and before every
  release. It fails if code that could reach the network ever appears in
  the add-in, so "no network" is kept true rather than hoped for.

### Known open security items

**Closed this release:** none — 0.5.6 is a feature release. Three changes
touch the security surface without closing an item, listed so none is a
surprise:

- *Interpret* can now set five more cell properties — underline,
  strikethrough, vertical alignment, indent and text rotation. Interpret
  only reaches the parts of Excel on a list someone has read and named
  (`SEC.1`, `docs/THREAT_MODEL.md` §1.1); these five were added to that
  list by name, each changes only how a cell looks, and anything the list
  does not name is still refused in words. No new *action* was added.
- *Interpret* can now do four more things for sorting and filtering, each
  added to that same list by name: take a sheet's filter buttons away
  (Excel lets a program only take them away, never add them), and find
  cells already on the sheet — a column of a range, the part of the sheet
  in use, and a range's visible cells. Sorting and filtering were already
  on the list; they now also take a second column and a second condition.
  Nothing here reaches outside the workbook.
- PROLOG now refuses, by name, the Prolog goals that would read from
  outside its own program and tables — `read`, `read_term`, `consult` —
  along with those that print, keep state between answers, or depend on
  the clock. None of them ever ran here; they used to be unknown names
  that quietly found nothing, and now they say why they will not.

**Written down this release, unchanged in behaviour:** writing
`docs/IT_REVIEW.md` found three things no earlier page said plainly.
`Refresh everything.` refreshes a workbook's existing data connections,
which can reach the network, and it is not held back by the `SEC.8` check
on internet-marked workbooks. Uninstalling leaves Frazaro's remembered
consents in the registry (`HKCU\Software\VB and VBA Program
Settings\Frazaro`). And uninstalling a standalone copy runs a hidden
PowerShell script, which security software may flag.

**Still open:** `SEC.3`, `SEC.7`, and two from the 2026-09-08 code
review — `SEC.10` and `SEC.15`. In plain words:
the remembered raw-VBA consent record still lives inside the workbook
(`SEC.10`);
formulas a program writes are not screened for functions that reach the
network (`SEC.15`); and effects like sending mail still run without a
permission prompt (`SEC.7`). `SEC.8` narrows that last one — it gates on
where the *workbook* came from — but does not close it: a phrasebook loaded
into a workbook of your own still reaches those verbs unprompted.

**Assessed and accepted, not fixed:** `SEC.12`, `SEC.14`, `SEC.16` and
`SEC.17`. Each needs a precondition an ordinary install does not meet —
mostly an Excel setting that ships off and that Frazaro never asks you to
turn on. The reasoning for each, and what would reopen it, is written down
rather than left implied.

The authoritative lists, kept current in one place instead of copied into
every release: [`README.md`](../README.md) in plain words, and
[`docs/BETA_REARVIEW.md`](BETA_REARVIEW.md) with the file, line, fix and
disposition for each.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from — and treat a workbook someone sent you the
same way before you press Interpret.** Frazaro makes no network call and
does not update itself; check the README's *Known open security items*
when you return for a newer build. Vulnerability reports:
`docs/SECURITY.md`. Everything else: `docs/SUPPORT.md`.

## 0.5.5

### What changed

- **PROLOG's type tests are finished.** `0.5.4` gave six — `var?`,
  `nonvar?`, `atom?`, `number?`, `atomic?` and `compound?`. Three more
  join them, and they are the three a real program actually reaches for.

  **`(ground? T)`** succeeds when `T` is completely filled in — no blanks
  anywhere inside it, however deep. `(ground? (order alice Qty))` fails
  while `Qty` is still unfilled and succeeds once it isn't. *Completely*
  is the word to read twice: one blank anywhere inside, at any depth, is
  enough to fail. This is the guard to write before anything that needs a
  finished value: printing it, storing it, comparing it.

  **`(callable? T)`** succeeds when `T` is a name or a structure —
  anything that could stand where a goal stands. A number can't, so
  `(callable? 42)` fails where `(nonvar? 42)` succeeds. That is the whole
  difference between the two, and it is the reason `callable?` exists.

  **`(is-list? L)`** succeeds when `L` really is a list — a chain that
  ends properly rather than trailing off. `(is-list? (list a b c))`
  succeeds; `(is-list? (cons a b))`, which ends in `b` instead of the
  empty list, does not; and neither does a list that is still half
  unfilled.

- **A type test answers; it never stops the query.** This is worth
  stating because the list *operations* behave differently on purpose. If
  you hand `(sum-list L N)` something that is not a list, it stops and
  says so — it cannot do its job without one. `(is-list? L)` on the same
  value simply answers *no*. That is what makes it usable as a guard:

  ```
  (fact (box (list 1 2 3)))
  (fact (box plain))
  (rule (total B N) (is-list? B) (sum-list B N))
  (query (box B) (total B N))
  ```

  gives one row — `(list 1 2 3)` and `6`. The `plain` row is skipped
  quietly instead of stopping the whole query, which is exactly what
  happens if you delete the `(is-list? B)` guard and run it again.

- **They see what a value *is*, not how it was written.** After
  `(= X 1)`, `(ground? X)` succeeds — it follows the value X was given
  rather than reading the letter `X`. The same is true all the way down:
  a list whose tail was filled in somewhere else is still a list.

- **`(integer? X)` and `(float? X)` are reserved, and Frazaro tells you
  why rather than failing quietly.** PROLOG has one kind of number
  today. `3` and `3.0` are not merely equal, they are the same value
  written down, so there is no honest answer to "is this one an integer".
  Rather than ship a test that would give a confident wrong answer — and
  one whose meaning would have to change later — the two names are
  reserved now and refused with an explanation pointing at `(number? X)`,
  the test that does exist.

  **That question is now settled, in the same release.** One kind of
  number is permanent, not a gap waiting to be filled, so `integer?` and
  `float?` stay refused for good rather than pending. What a person
  actually wanted when they reached for `integer?` — "has this got a
  fractional part?" — ships instead as **`(whole? X)`**, under a name
  that promises no type. `(whole? 3)` and `(whole? 3.0)` both succeed;
  `(whole? 3.5)` fails. The refusal now points at it by name.

- **If you write the Prolog spelling, Frazaro still tells you.**
  `(ground X)`, `(callable X)` and `(is_list X)` each stop and name the
  spelling to use instead. Note the last one: Prolog writes `is_list`
  with an underscore, Frazaro writes `is-list?` with a hyphen and a
  question mark, and typing the Prolog form gets you the Frazaro form
  rather than silence.

- **Get the spelling slightly wrong and Frazaro says so.** `sum-list` is
  written with a hyphen; Prolog writes `sum_list` with an underscore. A
  name that asks a question ends in `?`; `is-list` without one is one
  character short of `is-list?`. Each of those used to be an unknown
  predicate — which finds nothing and explains nothing — and now stops
  and names the spelling to use.

  The three it catches (`is-list`, `is_list?`, `sum_list`) are not a list
  someone thought of. They are every spelling you can reach from a real
  name by swapping a hyphen for an underscore or dropping a question
  mark, worked out mechanically over the whole reserved set, which is
  what makes the set complete rather than a guess.

- **PROLOG can say "or", and "if this then that, otherwise the other".**
  Until now the only way to say *or* was to write the same rule twice,
  which works for a whole rule and not at all for one step in the middle
  of a longer one. Two new forms fix that.

  **`(or A B ...)`** succeeds if any of its goals does, trying them left
  to right. Two or more, as many as you like:

  ```
  (rule (contact P) (or (employee P) (contractor P) (visitor P)))
  ```

  **`(if C T E)`** proves `C`; if that works it commits to the first way
  it worked and goes on to `T`, and if it doesn't it does `E` instead.
  The else-goal may be left off — `(if C T)` simply fails when `C` does.

  ```
  (query (if (member X L) (found X) (missing X)))
  ```

  Two things about `(if ...)` are worth knowing before you rely on it.
  It **commits**: once the condition succeeds one way, the other ways it
  might have succeeded are not tried. And the else-goal runs **only when
  the condition never succeeded at all** — if the condition works and the
  then-goal then fails, the whole thing fails rather than falling through
  to the else. Both are what Prolog does, and both are easy to assume
  backwards.

- **Which blanks come back as columns, when the answer had a choice in
  it.** Worth reading once, because the rule is not what you might guess
  and it is the one thing about these forms likely to look like a bug.

  A blank becomes a column of the answer **only if every way of
  succeeding fills it in**. Ask

  ```
  (query (or (p X) (q Y)))
  ```

  and you get a plain `TRUE` or `FALSE`, not two columns. That is
  deliberate: `X` is filled in only when the first branch is the one that
  worked and `Y` only when the second is, so on any given answer one of
  them is still blank — and a column of blanks reported as though it held
  values is worse than no column. Ask about a blank **both** branches
  fill in

  ```
  (query (or (p X) (q X)))
  ```

  and `X` comes back as a real column, one row per branch that succeeded.

  `(if C T E)` follows the same rule with one wrinkle worth knowing: a
  blank filled in by the **condition** counts as filled for the
  then-branch, because the condition's answers carry forward into it —
  but not for the else-branch, which is only ever reached when the
  condition failed and filled in nothing. So in

  ```
  (query (if (link W M) (tag M T) (nope T)))
  ```

  `T` is a column and `W` and `M` are not. Leave the else off, and there
  is only one way to succeed, so all three come back.

- **Why `or` and `if` rather than Prolog's `;` and `->`.** Real Prolog
  writes these `;` and `->`. **`;` is not available here and cannot be**:
  Frazaro's reader has always treated `;` as a comment that runs to the
  end of the line, long before PROLOG existed. That is not a small
  inconvenience — a `;` written mid-rule quietly deletes the rest of that
  line, and on the usual multi-line layout the rule still parses. You get
  a *different rule* rather than an error, and nothing tells you. So `;`
  had to go, and `or` and `if` are the words Frazaro already uses for
  these ideas everywhere else.

  Typing `(-> ...)` or `(\+ ...)` — Prolog's arrow and its negation —
  now gets you a short note naming the form Frazaro uses instead, rather
  than silence. The `\+` one matters most: left unrecognised it would
  have quietly *failed*, and a negation that fails looks exactly like a
  negation that worked.

- **A cut (`!`) inside one rule no longer cancels a cut in the rule that
  called it.** A real bug, present since cut shipped in `0.5.0` and found
  while building the above. If a rule used `!`, and a rule it called
  *also* used `!`, the inner one silently cancelled the outer one — so
  the outer rule kept offering answers it had been told to stop at. Rare,
  because it needs cuts at two levels at once, but wrong whenever it
  happened, and wrong quietly.

- **More reserved words.** Eighteen names join the reserved set, which is
  what lets Frazaro's advice about them always be right: `callable?`,
  `is-list?`, `ground?`, `integer?` and `float?`; the Prolog spellings
  `callable`, `is_list`, `ground`, `integer` and `float`; the three
  near-misses above; `whole?`, the whole-number test described
  earlier; and `or`, `if`, `->` and `\+` from the two control forms. If a
  knowledge base of yours uses one of those as a predicate name, it will
  need renaming — `ground`, `integer`, `float`, `or` and `if` are the
  plausible ones.

  `cons` and `nil` are still **not** reserved, unchanged from `0.5.4`: a
  predicate of your own may still be called that. Nor is `nth1`, which
  looks like a near-miss for `nth` and is not one — Prolog's `nth0` and
  `nth1` count from different places, so guessing which you meant would
  be worse than saying nothing.

- **Arithmetic grew from four operators to seventeen, in all three
  query languages at once.** Until now PROLOG, SQL and DATALOG could add,
  subtract, multiply and divide, and nothing else — no remainder, no
  integer division, no rounding. There is now `mod`, `rem`, `//`
  (integer division), `min`, `max`, `**`, `abs`, `sign`, `sqrt`,
  `truncate`, `round`, `ceiling` and `floor`.

  In PROLOG they are written the way the existing four are:
  `(is X (mod 7 3))`, `(is X (round 2.5))`. In DATALOG they go inside
  `let`: `(let S (abs A))`. SQL spells them as functions —
  `MOD(Salary, 7)`, `ABS(Salary - 95000)`, `ROUND(x)`, `POWER(2, 10)`,
  `DIV(a, b)`. The scalar two-argument minimum and maximum are
  **`LEAST`** and **`GREATEST`** in SQL, not `MIN`/`MAX`: those two are
  already aggregate functions there, and quietly changing what `MIN(x)`
  means would have altered the meaning of queries people have already
  written.

- **`mod` and `rem` both ship, because they disagree and you should not
  have to guess which one you got.** They are identical whenever the two
  numbers share a sign. When the signs differ they part company:
  `(mod -7 3)` is `2` and `(rem -7 3)` is `-1`. `mod` follows the sign of
  the number you divide *by*; `rem` follows the sign of the number you
  divide *into*. Shipping only one under the name `mod` would have been
  right half the time and silently wrong the other half.

- **`round` rounds halves away from zero.** `ROUND(2.5)` is `3`, not `2`.
  Worth stating because the rounding built into the underlying platform
  rounds halves to the nearest *even* number, which would have made
  `ROUND(2.5)` and `ROUND(3.5)` both `4`.

- **Arithmetic that has no answer now says so instead of guessing.**
  The square root of a negative, a negative number raised to a fractional
  power, a remainder or integer division by zero, and a result too large
  to represent are each refused by name, in whichever of the three
  languages you were writing. Previously several of these could surface
  as a raw platform error rather than an explanation.

- **A new release check: `tools/check_prolog_arith_operators.ps1`.** The
  code that actually performs arithmetic is shared by all three query
  languages, and each language separately decides which operators it will
  accept. That is an arrangement where an operator can be added in one
  place and forgotten in another — and the result would not be an error
  message but a *wrong number in a cell*, which looks like an answer.
  This check holds the one operator list, both refusal messages that
  advertise it, both computing functions, and SQL's own function-name
  map against each other, and fails the build if any of them disagree.

- **A new release check: `tools/check_test_assertion_safety.ps1`.** VBA
  evaluates *every* part of an `And`, even when an earlier part has
  already answered no. So a test written as one expression — check the
  answer has three rows, then read the third row — still reads that third
  row when there are only two, and stops with a raw error instead of
  reporting the failure. Harmless every day the test passes; on the one
  day it finds something, the run dies at the exact moment it was about
  to say what broke, and every later test never runs.

  This is invisible in a passing test suite by construction, so nothing
  but a static check can find it. **165 assertions across all four test
  files were carrying it**, and all 165 are fixed. The check now runs at
  every release and holds each file at zero — and fails just as loudly if
  a file comes in *under* its recorded number without the number being
  lowered too, so the ground gained cannot quietly be given back.

  It also catches the half-fixed version, which is the one that would
  have bitten: a test can be repaired where it checks the answer and left
  reading past the end of it where it *reports* the answer — and the
  report is built every single time, pass or fail.

  And it catches the opposite mistake: a test that expects an answer
  Frazaro can no longer give. When a value's printed form changes, every
  test naming the old form has to change with it — that happened once and
  two were missed by hand. The check now knows which forms are retired,
  so it lists them instead of leaving it to memory.

  Nothing here changes what Frazaro does. It changes what happens on the
  day something else goes wrong: you get told, instead of the report
  being destroyed by the test that was about to make it. The rewrite was
  checked against VBA's own rules for every shape a result can take,
  proving each replacement means exactly what it replaced, and both the
  check and that proof were mutation-tested rather than assumed to work.

### Known open security items

**Closed this release:** none — 0.5.5 is a feature release and touches no
security item.

**Still open:** `SEC.3`, `SEC.7`, and two from the 2026-09-08 code
review — `SEC.10` and `SEC.15`. In plain words:
the remembered raw-VBA consent record still lives inside the workbook
(`SEC.10`);
formulas a program writes are not screened for functions that reach the
network (`SEC.15`); and effects like sending mail still run without a
permission prompt (`SEC.7`). `SEC.8` narrows that last one — it gates on
where the *workbook* came from — but does not close it: a phrasebook loaded
into a workbook of your own still reaches those verbs unprompted.

**Assessed and accepted, not fixed:** `SEC.12`, `SEC.14`, `SEC.16` and
`SEC.17`. Each needs a precondition an ordinary install does not meet —
mostly an Excel setting that ships off and that Frazaro never asks you to
turn on. The reasoning for each, and what would reopen it, is written down
rather than left implied.

The authoritative lists, kept current in one place instead of copied into
every release: [`README.md`](../README.md) in plain words, and
[`docs/BETA_REARVIEW.md`](BETA_REARVIEW.md) with the file, line, fix and
disposition for each.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from — and treat a workbook someone sent you the
same way before you press Interpret.** Frazaro makes no network call and
does not update itself; check the README's *Known open security items*
when you return for a newer build. Vulnerability reports:
`docs/SECURITY.md`. Everything else: `docs/SUPPORT.md`.

## 0.5.4

### What changed

- **PROLOG can now ask what kind of thing a value is.** Six new goals —
  `var?`, `nonvar?`, `atom?`, `number?`, `atomic?` and `compound?` — each
  take one term and succeed or fail depending on its shape.
  `(number? Salary)` succeeds when `Salary` has been bound to a number,
  `(var? X)` succeeds only while `X` is still unfilled, `(compound? T)`
  when `T` is a structure like `(name Alice)` rather than a single value.
  They are the ordinary Prolog guards, and until now a rule had no way to
  check that what it received was the kind of thing it knew how to
  handle.

  They never change anything — a type test looks at a value and answers.
  It cannot fill in a blank, so it adds no column to your results.

  **The question mark is deliberate, and it is not decoration.** In
  Prolog you write `atom(X)`, but Frazaro's PROLOG is written in
  parentheses, where `(atom bob)` looks exactly like a piece of *data* —
  `(color red)`, `(name Alice)` and `(f a)` are all ordinary data written
  in that same shape. Prolog doesn't have that problem because its
  grammar tells a question apart from a value; ours doesn't, so the name
  has to. The question mark also matches what Frazaro already does
  elsewhere: its macro layer has long spelled its questions `null?`,
  `eq?` and `equal?` beside `car`, `cdr`, `cons` and `list`. A name that
  asks something ends in `?`; a name that produces something doesn't —
  which is why `between` below has none.

  **If you write the Prolog spelling, Frazaro tells you.** `(atom X)`
  doesn't quietly find nothing; it stops and says that this type test is
  spelled `(atom? X)`. All six work that way, and none of the six bare
  names can be used for a predicate of your own, so the advice can never
  be wrong.

- **`(between Low High X)` counts.** With `X` unfilled it produces every
  whole number from `Low` to `High` in turn, so a query can range over
  1 to 10 without ten hand-written facts. With `X` already filled it
  tests instead: `(between 1 10 5)` simply succeeds. Both bounds may be
  arithmetic, so `(between 1 (+ N 1) X)` works.

  If `Low` is greater than `High` the range is empty and the goal finds
  nothing — that is an ordinary "no rows", not an error, so a range
  computed from your data can safely come out empty.

  **A limit worth knowing before you meet it: the range can span at most
  119 values.** A whole query gets a fixed budget of resolution steps,
  and each generated value spends one. Ask for more and Frazaro refuses
  up front and tells you the range was too wide — deliberately, so that
  a wide range never gets reported as the *other* thing that exhausts
  that budget, a rule that calls itself forever. Those are different
  mistakes and now have different messages. Testing a single value is
  not affected, so `(between 1 1000000 5)` succeeds normally.

- **What the two are for, together.** A rule can now check what it was
  handed before doing arithmetic on it, and a query can feed it a range
  instead of a table:

  ```
  (rule (halved N H) (number? N) (is H (/ N 2)))
  (query (between 1 5 X) (halved X Y))
  ```

  That spills five rows — 1 through 5 beside their halves. The guard is
  what makes the rule safe to call with anything at all: without
  `(number? N)`, asking for `(halved eng H)` stops the entire query with
  an arithmetic refusal, because `eng` cannot be divided. With the guard,
  that call simply doesn't match, and the rest of the query carries on.
  That is what a type test is *for*, and until this release there was no
  way to write one.

- **Text that looks like a number is still text, and now says so.** A
  cell containing `42` typed as text is not the number 42 in PROLOG, and
  never has been — `=` and `==` have always treated the two as different
  values. The new goals follow that same rule rather than inventing a
  second one: `(atom? "42")` succeeds and `(number? "42")` does not, and
  `(between 1 10 "5")` is refused with a message that says plainly that
  a text cell reading 5 is not the number 5.

- **Comparing a text cell to an unquoted name no longer answers
  silently.** This is the one behaviour change in this release that can
  affect a program that used to run.

  A text cell reading `eng` and the bare word `eng` written in a query
  are different values in PROLOG — that has always been true, and it is
  still true. The problem was that nothing said so. Asking
  `(\= Dept eng)` — "is this department something other than eng" —
  answered **yes for every row of the table**, because the cell's value
  and the word you typed were never the same thing to begin with. The
  answer was correct about the values and useless as an answer, and
  there was no way to tell that had happened.

  Those four comparisons — `=`, `\=`, `==`, `\==` — now stop and explain
  when the only difference between the two things being compared is the
  quoting:

  ```
  the text "eng" and the name eng are different things in PROLOG - a
  text cell keeps its quotes, so if you meant the cell's own value
  write it as "eng".
  ```

  It is deliberately narrow. Two values that genuinely differ still
  compare quietly, exactly as before, so ordinary filtering is
  untouched. And it applies **only** where you explicitly asked whether
  two things are the same — never while Frazaro is matching rules
  against facts, because a query that finds a real answer must never be
  stopped by a near-miss against some *other* fact it was not asking
  about.

  The same message appears for a text `42` against a numeric `42`, where
  it says "number" rather than "name".

- **And a query that finds nothing at all now tells you why, if the
  reason is that quoting.** The same mistake shows up a second way: not
  as a comparison that answers oddly, but as a query that just comes back
  empty. `(query (emp N eng))` against a table whose department column is
  text used to return nothing and say nothing, and an empty result is the
  hardest thing to debug because it looks the same whatever caused it.

  Frazaro now checks, **only when a query found no rows whatsoever**,
  whether one of the things you asked for differs from something it
  stored by nothing but the quoting — and if so, says that instead of
  handing back a blank.

  It runs after the search, never during, so it cannot affect a query
  that works: if your query finds even one row, this check does not
  happen at all. Adding a fact that matches makes the message disappear.
  And an empty result with no such near-miss in it stays exactly as it
  was — empty, and silent, because sometimes the honest answer is that
  there is nothing there.

  Two cases it does not catch, worth knowing so the silence is not
  mistaken for a clean bill of health: a mismatch that only exists inside
  a rule's body, and one that only appears after an earlier part of the
  query has filled in a variable. Both need bookkeeping during the search
  that would slow down every query to help a few, so they are left for a
  later release.

- **A query no longer loses a column when your data is shaped like a
  goal.** If a fact stored a value such as `(not bob)` or `(atom? bob)` —
  ordinary data that happens to be written in the same shape as a
  Frazaro goal — a query reading it found the value correctly and then
  left it out of the result. One column where two were due.

  This one is worth calling out because of how it failed: nothing was
  reported, and the answer that came back was not empty or wrong-looking,
  just *narrower* than it should have been. A missing row is obvious; a
  missing column is easy to read straight past. Fixed, and both columns
  now appear.

- **A release check learned to look both ways.** PROLOG refuses to let
  you define a predicate with a reserved name, and a static check has
  been holding three places to agreeing on which names those are. It
  turned out to be checking only one direction — it would catch a name
  that was reserved but did nothing, and miss a name the solver acted on
  without ever reserving it, which would let you define a predicate that
  is then silently ignored. Nothing shipped had that fault; the check
  simply could not have caught it. It now checks both directions, and
  was verified by deliberately introducing the fault to confirm it is
  caught.

- **PROLOG has lists, and `findall` stops being a dead end.** Until now
  `findall` handed you a bag of answers and there was nothing you could
  do with it. You could put it in a cell and read it. You could not count
  it, add it up, pick the third one out of it, or ask whether something
  was in it. That was the largest missing piece in the engine, and it is
  the piece that made `findall` — the goal that gathers everything
  matching a pattern — worth much less than it looks.

  Six new goals: `(length L N)`, `(member X L)`, `(nth N L X)`,
  `(append A B C)`, `(reverse L R)` and `(sum-list L N)`. The one most
  people will reach for first:

  ```
  (fact (sale north 10)) (fact (sale north 20)) (fact (sale south 30))
  (query (findall Amount (sale north Amount) Bag) (sum-list Bag Total))
  ```

  That gathers every northern sale and totals it — 30 — in one query,
  which simply could not be written before.

- **Most of these goals run in more than one direction.** That is the
  part worth knowing, because it is what makes them worth having rather
  than being six functions with awkward spelling.

  `(member X L)` **tests** membership when `X` is already filled in, and
  **produces one row per element** when it is not. `(nth N L X)` picks
  the `N`th element when you name a position, and walks the whole list,
  position beside element, when you leave `N` blank. `(length L N)` and
  `(sum-list L N)` compute a count or a total when `N` is blank and
  **check** one when it is not.

  `(append A B C)` is the one that goes furthest. Given two lists it
  joins them. Given the *result* and either half, it hands back the
  other half — so `(append (list a) Rest Whole)` drops a known prefix,
  and `(append Front (list z) Whole)` drops a known suffix.
  Given only the result, it produces **every way the list could be split
  in two**, one row per split, both ends included. None of that is extra
  machinery; it is the same goal read in different directions, which is
  what a logic language is for.

  What it will not do is work backwards from nothing. `(append A B C)`
  with all three blank has nothing to join and nothing to split, and is
  refused by name rather than left to find nothing.

- **How a list is written, and why it is not `[a, b, c]`.** Write a list
  as `(list a b c)`. The empty list is `nil`.

  Square brackets are not available, and the reason is worth knowing
  rather than guessing at: Frazaro's reader treats `[`, `]` and `|` as
  ordinary letters, exactly like `a` or `x`. Writing `[H|T]` would not
  produce a list — it would produce one long word spelled `[H|T]`.
  Teaching the reader about brackets means changing it for all five of
  Frazaro's languages at once, which is a far larger change than this
  one.

  Underneath, a list is built from pairs — `(cons Head Tail)` — and
  `(list a b c)` is shorthand for `(cons a (cons b (cons c nil)))`. You
  can write either; they are the same list, not two kinds of list. The
  long form is what makes **taking a list apart** possible, because a
  pair is an ordinary term and matching is what this language already
  does:

  ```
  (query (findall X (color X) Bag) (= Bag (cons First Rest)))
  ```

  `First` is the first colour and `Rest` is everything after it. The same
  works in a rule's head, so a rule can be written to match only
  non-empty lists. That is the whole trade: `(list …)` to write one down,
  `(cons …)` to pull one apart.

  One thing `(list …)` is not is a goal. It builds a value, so it belongs
  in an argument. Writing `(query (list a b))` on its own stops and says
  so rather than quietly finding nothing.

- **This changes what a `findall` bag looks like in a cell.** A bag of
  three colours used to display as `(red green blue)`. It now displays as
  `(list red green blue)`. If you have a sheet that reads a bag as text,
  that text has changed.

  The old display was shorter by two characters and was also a small lie:
  you could not paste `(red green blue)` back into a program and get the
  same list — Frazaro would read it as *a thing called red with two
  parts*. The new display reads back as exactly the list it prints, which
  is the property the old one lacked.

  A list is only ever displayed this way when it really is a list. A pair
  whose tail is not a list — `(cons a b)`, which is a perfectly legal
  term and not a list at all — displays as itself, so the display can
  never make something look like a list that is not one.

- **The empty list is `nil`, and that fixed an old wrong answer.**
  A `findall` that finds nothing used to hand back `()`, and asking
  `(compound? Bag)` about it answered *yes* — because internally it was
  an empty structure. Standard Prolog says the empty list is a simple
  value, not a structure. Now that it is `nil`, `(compound? Bag)` answers
  *no* and `(atomic? Bag)` answers *yes*, which is the standard answer.
  Nothing in the type tests changed to make that happen.

- **These goals do their work in one step, on purpose.** A whole query
  gets a fixed budget of resolution steps. Written the traditional way —
  as Prolog rules that call themselves — `length` and `append` would
  spend one step per element and run out of budget on a list of about a
  hundred, which is exactly the size a real `findall` bag reaches. So
  they are built into the engine instead: counting, reversing and summing
  a list cost one step no matter how long it is. `member`, `nth` and
  `append` still cost one step per answer they hand back, because each of
  those really is a separate answer.

- **When something is not a list, Frazaro says so rather than finding
  nothing.** `(length (color red) N)` does not quietly return no rows; it
  stops and tells you it needs a list, and shows you what a list looks
  like. That includes a half-built list — `(cons a T)` where `T` has not
  been filled in yet — which some Prologs would complete for you and
  Frazaro deliberately refuses, because a query that quietly finds
  nothing is indistinguishable from one that correctly found nothing.

  Seven names are now reserved and cannot be used for your own
  predicates: `length`, `member`, `nth`, `append`, `reverse`,
  `sum-list` — and `list`, which is reserved because it is the shorthand
  marker rather than because it is a goal. `member` is the one worth
  flagging, since it is a natural name for a fact about people.

  `cons` and `nil` are **not** reserved. They are ways of writing data,
  so a predicate of your own may still be called `cons`.

### Known open security items

**Closed this release:** none — 0.5.4 is a feature release and touches no
security item.

**Still open:** `SEC.3`, `SEC.7`, and two from the 2026-09-08 code
review — `SEC.10` and `SEC.15`. In plain words:
the remembered raw-VBA consent record still lives inside the workbook
(`SEC.10`);
formulas a program writes are not screened for functions that reach the
network (`SEC.15`); and effects like sending mail still run without a
permission prompt (`SEC.7`). `SEC.8` narrows that last one — it gates on
where the *workbook* came from — but does not close it: a phrasebook loaded
into a workbook of your own still reaches those verbs unprompted.

**Assessed and accepted, not fixed:** `SEC.12`, `SEC.14`, `SEC.16` and
`SEC.17`. Each needs a precondition an ordinary install does not meet —
mostly an Excel setting that ships off and that Frazaro never asks you to
turn on. The reasoning for each, and what would reopen it, is written down
rather than left implied.

The authoritative lists, kept current in one place instead of copied into
every release: [`README.md`](../README.md) in plain words, and
[`docs/BETA_REARVIEW.md`](BETA_REARVIEW.md) with the file, line, fix and
disposition for each.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from — and treat a workbook someone sent you the
same way before you press Interpret.** Frazaro makes no network call and
does not update itself; check the README's *Known open security items*
when you return for a newer build. Vulnerability reports:
`docs/SECURITY.md`. Everything else: `docs/SUPPORT.md`.

## 0.5.3

### What changed

- **`SEC.8` — a workbook from the internet can no longer reach outside
  itself.** Office has blocked macros in files that came from the
  internet since 2022. That block reads the Mark-of-the-Web that Windows
  puts on a downloaded or emailed file, and it applies to macros — VBA
  code. A Frazaro program is not a macro: it lives in the cells. So a
  plain `.xlsx` someone mails you could carry a complete program, and
  after the ordinary "Enable Editing" click Frazaro would run it with the
  add-in's own privileges. Office's protection had never applied to it,
  because from Office's point of view there was nothing there to protect
  you from.

  Frazaro now reads that same mark itself, and refuses the things that
  reach outside the workbook: opening or saving a file, saving a copy,
  printing, exporting a PDF, composing an Outlook email, and
  password-protecting or unprotecting a sheet. Everything else runs
  normally — reading and writing cells, formatting, formulas, loops, the
  whole ordinary business of a program. A workbook from the internet
  still opens, still calculates, and can still be edited exactly as
  before. It just cannot reach off the page.

  The refusal names the thing the program tried to do, so you find out
  what a workbook wanted rather than only that something was blocked.

  **Clicking "Enable Editing" does not grant this.** Excel shows its own
  yellow banner on a file from the internet — that bar is Excel's, not
  Frazaro's, and it governs whether you can type in the sheet. Frazaro's
  check is separate and happens later, when a program actually tries to
  reach outside the workbook. Enabling editing leaves the file's origin
  mark exactly where it was, which is why the two decisions stay
  independent: you can read and edit a workbook someone sent you without
  also agreeing that its program may email, print, or save files. Only
  unblocking the file in Windows does that.

  **To allow it, you unblock the file in Windows** — close the workbook,
  right-click the file in File Explorer, choose Properties, tick
  *Unblock*, then reopen it. Frazaro deliberately offers no button of its
  own for this. That is the whole point of the design: a workbook that
  arrives from outside must not be able to carry its own permission slip,
  and anything Frazaro stored could be forged or shipped inside the
  workbook itself. Windows' own checkbox is the one channel a mailed file
  cannot reach.

  Two limits worth knowing, both deliberate. A workbook opened straight
  from a `https://` SharePoint address refuses external effect too:
  Windows cannot record an origin mark in that kind of location, so
  Frazaro cannot tell where the file came from, and it does not guess.
  (A OneDrive folder that syncs to your own disk is an ordinary local
  folder and is not affected — that is the common case.) The workaround
  is to save a local copy. And a file with no mark at all is treated as
  local, because an ordinary file you made yourself has no mark either
  and the two are genuinely indistinguishable — the protection comes from
  the mark being *present*, never from it being absent.

- **A new release check: `tools/check_sec8_provenance_gate.ps1`.** The
  self-test suite does no Office automation and cannot manufacture a file
  that claims to be from the internet, so it can check every part of the
  *decision* but never the call sites. This static check runs at every
  release and fails it if any of the nine external-effect sites loses its
  guard, or if the one line that reads the workbook's origin goes
  missing. It is mutation-tested in both directions rather than assumed
  to work, and it found a site the audit had missed on its first run.
  What it cannot do is prove the guard *functions* against a real marked
  workbook — only Excel can do that, so that rests on a live test.

- **Email problems now tell you what went wrong instead of dropping you
  into the code editor.** If Outlook was not available, or an attachment
  path did not exist, an email sentence did not report that — it stopped
  the program with Visual Basic's own `Run-time error '5'` dialog, the one
  with a *Debug* button. The message Frazaro meant to show you ("Could not
  start Outlook to create the email") existed the whole time and never got
  the chance to appear. This affected every release that has had the email
  sentence, and needing no Outlook installed is all it took to trigger it.

  The cause is a VBA quirk this project had already documented and proved
  with a standalone test: the mechanism the interpreter used to reach that
  particular helper does not carry an error back to the code that would
  have caught and displayed it. The email helper now uses a direct call
  instead, so its messages arrive the way every other refusal in Frazaro
  does.

  **Not fixed in this release, and named so it is not mistaken for
  shipped:** eight other helpers still reach you the same wrong way when
  they fail — an unrecognised colour, a missing key, and six pivot-table
  and worksheet operations with an unrecognised option. Those show
  Visual Basic's dialog rather than a Frazaro message. Only the email
  helper is fixed here, because only it was in the way of `SEC.8`; the
  rest are recorded in
  [`docs/BETA_REARVIEW.md`](BETA_REARVIEW.md) with the full list rather
  than fixed in a hurry alongside a security change.

- **A build defect found by a new release check: `VlaSlice` was never
  shipped.** Frazaro keeps two lists of modules — one for what a built
  add-in contains, one for what the development workbook reloads — and
  nothing had ever checked that they agree. `VlaSlice`, a small internal
  class the language core names directly, was in the second list and not
  the first, so a freshly built add-in would not have contained it at
  all. It has been in that state since `0.5.0`. It is now in both lists.

  This is the ninth time this project has made the same mistake, and the
  first time a machine caught it rather than a person hitting the
  resulting error. `tools/check_devrig_mods_parity.ps1` now runs at every
  release and fails it if either list is missing something the other has,
  in either direction — the more dangerous direction being a module the
  development workbook compiles happily while a built add-in would not,
  which is exactly what stayed hidden here. Modules that genuinely should
  not ship (the test suites, the builder itself) are listed in the check
  with a reason beside each. It is mutation-tested in both directions.

- **`SEC.13` — Word documents are no longer opened with macros enabled.**
  *Import Program File…* accepts Word documents, and it opened them
  through Word automation without setting Word's own
  `AutomationSecurity`. Word's default in that mode is *Low*, so a
  `.docm` carrying an `AutoOpen` or `Document_Open` macro ran that macro
  silently the instant Frazaro read the text out — no Trust Center
  prompt, because a document opened by automation does not get one. This
  was a live, one-click path in every shipped edition, on the menu and
  the ribbon both, and worse than the audit first recorded: the import
  router matches `.docm` by name, not only `.docx`/`.doc`.

  Frazaro now force-disables macros for the duration of the read, and the
  guard is deliberately set *after* the error handler is armed, so if it
  cannot be set the import refuses rather than opening the document
  unguarded. Because Frazaro reuses a copy of Word you already have open
  rather than always starting its own, it captures your Word's previous
  setting and puts it back afterwards — an application it does not own is
  handed back as it was found. If that restore should ever fail it stays
  quiet on purpose: the failure leaves Word *more* cautious than before,
  never less, and a Word restart clears it.

  Word documents still import exactly as they did — the text, the
  typography cleanup, everything. The only thing that changed is that a
  document's own macros no longer get to run on the way in.

  If you followed the previous advice in the README (*"import only Word
  files you wrote yourself, or paste the text instead"*), you no longer
  need to. That advice has been removed.

  **Not in scope, and named so it is not mistaken for shipped:** Frazaro
  still does not *ask* you before opening a document you point it at.
  Consent prompts are `SEC.7`/`SEC.8`'s subject and remain open.

- **Importing a Word document can no longer appear to hang Excel.** When
  Frazaro starts its own copy of Word to read a document (rather than
  reusing one you already have open), that copy is invisible. If Word
  decided to *ask* you something about the file rather than simply fail
  — a damaged document, a file-conversion prompt — the question appeared
  on a window you could not see or click, and Excel looked frozen with no
  way forward but Task Manager. Frazaro now tells the copy it starts not
  to raise alerts, so a document Word dislikes comes back as the ordinary
  refusal message instead of a hidden prompt.

  A copy of Word you already had open is deliberately left alone. Its
  windows are visible, so its questions are answerable — and silencing
  alerts in an application Frazaro did not start would cost you warnings
  you should see. That is the opposite direction of failure from the
  macro guard above, which is why the two are treated differently.

  Found while writing the live test for `SEC.13` rather than from a
  report: it never actually fired during testing. It is fixed as a
  hazard, not as an observed fault.

- **A new release check: `tools/check_word_automation_security.ps1`.**
  The self-test suite does no Office automation at all, so nothing in it
  could ever have caught this or its return. A static check now runs at
  every release and fails it if any code that opens a Word document does
  not force-disable macros first, in that same procedure, before the
  open. It is mutation-tested in both directions rather than assumed to
  work. What it cannot do is prove the guard *functions* — only Word can
  do that, so that rests on a live test, with a reproducible fixture
  recipe recorded in `tools/sec13_word_fixture.md`.

- **`SEC.11` — the fingerprint that remembers your consent is now a real
  one.** When a phrasebook contains `raw` — VBA written directly into a
  grammar file — Frazaro asks before loading it, and can remember your
  answer. What it remembers is tied to a fingerprint of that file's exact
  contents, so that editing the phrasebook at all asks you again.

  The fingerprint was too weak for the job. It was a simple arithmetic
  checksum, and checksums of that kind can be *aimed*: someone who wanted
  a different phrasebook to carry your fingerprint could adjust a couple
  of characters inside a comment until the numbers matched. Your "yes" to
  a file you had read could then have been silently inherited by a file
  you had never seen. It is now SHA-256, the same standard used for
  software signatures, where aiming at an existing fingerprint is not
  something anyone knows how to do.

  **You will be asked once more for phrasebooks you had already
  approved.** Old answers were filed under the old fingerprint and cannot
  be matched to the new one. Being asked again is the safe direction to
  fail, so nothing tries to convert them; the old entries are simply left
  alone and ignored.

  Two things deliberately did *not* change. The fingerprint still ignores
  spaces, tabs and line breaks, so re-indenting a phrasebook, or opening
  one that was saved on a different operating system, still counts as the
  same file rather than sending you back through the question. Changing
  what a phrasebook *says*, by even one character, still does.

  And exported *Expanded Phrasebook* files you already have keep
  working. The freshness check reads the older form, confirms the file is
  current, and tells you the next export will upgrade it. A re-exported
  file carries the new form, written out as `sha256:` followed by the
  digest so the two generations can never be mistaken for one another.
  Nothing you already have needs regenerating.

  Built without depending on anything being installed. The obvious route
  was to borrow Windows' own cryptography through .NET. Measured on the
  development machine, that turned out to fail outright — the component
  is registered against a version of .NET that Windows 11 no longer
  installs by default — so a phrasebook would have become unloadable on
  an ordinary machine. Frazaro now computes SHA-256 itself, in about 150
  lines, which works the same everywhere and can be checked completely by
  the self-test suite against the published standard test values.

- **A new release check: `tools/check_hash_twin.ps1`.** The fingerprint is
  computed in two places — in Frazaro itself, and in the release script
  that verifies an exported phrasebook is current. The two must agree
  exactly, and until now the only thing keeping them in step was a comment
  saying so. This project has been bitten nine times by that same shape.
  Both sides are now pinned to the same published test values, so either
  one drifting fails a release rather than being noticed later. The two
  were also checked against each other on the real 89,446-byte English
  phrasebook and produce the identical fingerprint.

- **`SEC.9` — a phrasebook has to be approved on this computer before it
  can replace the built-in grammar.** A phrasebook decides what your
  sentences *mean*. Frazaro used to prefer a grammar file sitting next to
  the workbook you had open over its own built-in copy, and it would
  silently reload any phrasebook path a workbook remembered — including
  paths on other machines. Between them, a workbook could quietly decide
  what every sentence in it did.

  This was not theoretical. On 2026-09-08 an old `english.vla` left over in
  a Downloads folder was picked up ahead of the add-in's own copy, simply
  because the workbook being opened was in that folder too. It failed
  noisily, but only by luck: that copy was old enough not to know a word the
  program used. A phrasebook that was merely *different* rather than *older*
  would have loaded without a word and changed what the program did.

  Frazaro now asks, once, naming the file, before using a phrasebook that
  did not come with it. Your answer is remembered **on this computer**,
  not inside the workbook — a workbook someone sends you cannot arrive with
  its own permission already granted, because the permission was never
  something a workbook could carry. Choosing a phrasebook yourself through
  *Load Phrasebook* counts as the answer, so picking a file never asks you
  about it a second time.

  Saying no is safe and is not a dead end: Frazaro simply uses its own
  built-in grammar, which is complete. The answer is remembered either way,
  so you are not asked again on every command — and if the file itself
  changes later, you are asked again, because the answer was about the file
  you were shown, not about its name.

  **Network and web locations are described, never quietly contacted.** If a
  workbook asks for a phrasebook on a network share, the question says so
  and warns that opening it would hand your Windows sign-in to that server.
  Nothing touches the location until you say yes — not even a check for
  whether the file exists, which is itself enough to leak that sign-in.

  **Changing your mind: *Forget Phrasebook Approvals*,** in the Utilities
  group. It lists every answer this computer has recorded and clears them
  all, so the next time each phrasebook is used you are asked again. Your
  answers are otherwise kept for good, which is deliberate — a phrasebook
  you declined stays declined rather than asking you the same question
  every time you open Excel, because a dialog that keeps reappearing is one
  people learn to click through without reading.

  Your recorded answers also appear in *Copy Diagnostic Report*, so if a
  sentence stops being understood you can see whether a phrasebook it
  needed was declined. That matters because the symptom on its own is
  indistinguishable from a typo: without it, Frazaro would just say it does
  not understand the sentence and never mention the phrasebook.

- **A new release check: `tools/check_sec9_phrasebook_gate.ps1`.** The
  self-test suite can check the whole decision — whether a path is remote,
  whether it belongs to Frazaro itself — but it cannot click a dialog or
  read the saved answers, so the places that *call* the check are where this
  could quietly stop working. A static check now runs at every release and
  fails it if either loader loses its gate, if the two decision functions
  are removed, or if the approval is ever moved to run after the
  file-existence probe rather than before it. That last one is the one that
  matters: reordering those two lines would restore the credential leak
  while leaving every visible behaviour, and every test, exactly as it was.
  It is mutation-tested in both directions rather than assumed to work.

- **Sixteen things a program can get wrong now say so, instead of dropping
  you into the code editor.** Type a colour Frazaro does not recognise, ask
  for a key that was never stored, or name a pivot table that is not there,
  and until now Excel's own *Run-time error '5'* box appeared, with a
  **Debug** button that opened the VBA editor at a line of Frazaro's
  internals. Frazaro had written a perfectly clear explanation for each of
  these — it just never reached the screen. You now get the ordinary Frazaro
  message saying what was wrong with what you wrote: *'notacolor' is not a
  color - use "#RRGGBB", like "#FF69B4"*, or *there is nothing stored at key
  'no-such-key'*, or *no pivot table named 'nosuchpivot' in this workbook.*

  The cause was one mechanism, not sixteen separate bugs. Frazaro reaches
  most of its built-in helpers through a general-purpose Excel facility
  that, it turns out, does not carry an error back to the code that asked
  for it. Any helper whose job includes saying no was therefore unable to
  say no. The helpers that only compute something were unaffected, which is
  why this took so long to notice: the failure was invisible until you made
  a mistake.

  Eight of the sixteen were found by hand. The other eight were found by the
  release check below, and had been missed — they refuse a misspelled pivot
  table name through a shared piece of code rather than in their own, which
  is exactly the kind of thing a person reading down a list does not see.

- **A second fix underneath it: Frazaro no longer decides whether something
  is one of its helpers by trying it and seeing what happens.** It now checks
  the name against the list of helpers first. The old approach could not tell
  "that is not a helper at all" apart from "that is a helper, and it is
  refusing" — so a deliberate refusal could have been reported as *"'vlacolor'
  is not a form..."*, naming the wrong problem with complete confidence. Being
  told the wrong thing firmly is worse than being told nothing, and this
  removes the possibility rather than making it less likely. A genuinely
  unknown name still gets exactly the same "not a form, place helper, dotted
  global, built-in, or VLA_Runtime helper" message it always did.

- **A new release check: `tools/check_runtime_raise_dispatch.ps1`.** The
  boundary this fix relies on — a helper that can refuse must be called the
  direct way — was being kept in someone's head, and had already slipped
  three times, each time found by a user hitting the crash. The check now
  works it out from the source: it reads every helper, follows the shared
  code they call, and fails the release if any helper that can refuse is
  still reached the broken way. It was written before the fix, so its first
  run listed exactly the work to do, and it is mutation-tested in both
  directions rather than assumed to work. It is what found the eight the
  hand count missed.

- **`PROLOG.7` — you can now compare two numbers in a rule.** PROLOG could
  already *calculate* — `(is Total (+ X Y))` works out a value and gives it
  a name. What it could not do was *test* one number against another. There
  was no way to write "salary over 80000", so the staffing example in the
  README had to match an exact department instead of a threshold, which is
  not what anyone actually wants to ask.

  The six comparisons now work as goals in their own right: `<`, `>`, `=<`,
  `>=`, `=:=` (equal) and `=\=` (not equal). You write them like any other
  goal in a rule body or a query:

  ```
  (rule (well-paid Name) (employee Name Salary) (> Salary 80000))
  ```

  A comparison is a question, not a calculation. It either succeeds and the
  search carries on, or it fails and that row simply isn't in the answer —
  it never invents a value or fills in a variable. Both sides can be
  arithmetic, so `(> (+ Base Bonus) 80000)` is fine.

  **A refusal now names the form you actually wrote.** Comparisons share
  the machinery that `(is ...)` uses to work out each side, and that shared
  code used to say "(is ...)" in every complaint it made. So a program
  containing `(> Salary 80000)` and no `(is ...)` at all could be told its
  problem was with `(is ...)` — pointing at a form the author never typed.
  Five separate messages did this. They now name whichever form is really
  running, and a check runs on every release to keep it that way. Nothing
  about the wording changed for `(is ...)` itself.

  The six symbols are now reserved, so they can't be used as your own
  predicate names — the same rule that already applies to `is`, `not`,
  `findall` and `!`. Unification (`=`) and structural equality (`==`) are a
  separate item, built next (below); `=:=` here is numeric equality only.

- **`PROLOG.8` — you can now match two things against each other in a
  rule.** Until now a variable only ever picked up a value one way: by
  lining a goal up against a stored fact. There was no way to say, in the
  middle of a rule, "let X be this" or "only if these two are different".
  Four goals now do that:

  ```
  (rule (corner C) (= C (point 0 0)))
  (rule (rival A B) (player A) (player B) (\= A B))
  ```

  `=` **matches, and fills in blanks while it does.** `(= X 1)` gives X the
  value 1. It works in both directions and reaches inside structures, so
  `(= (point X 7) (point 3 Y))` sets X to 3 and Y to 7 in one step.

  `\=` **succeeds when two things do NOT match** — the usual way to say
  "these must be different". It fills nothing in, whichever way it turns
  out.

  `==` **and** `\==` **ask a stricter question: are these already the same
  thing?** They never fill anything in. `(== X 1)` is false when X is still
  blank — X *could* become 1, but it isn't 1 yet — where `(= X 1)` would
  make it so. That difference is the whole reason both exist: use `=` when
  you want to set something up, `==` when you want to check it without
  changing anything.

  **These are not the number comparisons.** `(= 2.0 2)` is false: `=`
  matches things as written, and `2.0` is not written the same way as `2`.
  If you mean "the same number", that is `=:=` from `PROLOG.7`. The same
  goes for `==`.

  **Refusals name the goal you actually wrote.** Writing `(\= 1)` tells you
  about `(\= ...)` — not about `(= ...)`, `(not ...)` or `(is ...)`. `\=`
  means the same thing as "not equal to", and the shortest way to build it
  would have been to quietly rewrite it into `(not (= ...))`, but then your
  mistake would have been reported against a form you never typed. It is
  built directly instead, so the message can name what you wrote. This is
  the same principle `PROLOG.7` applied to the shared arithmetic messages.

  **Matching text that came from a table: quote it.** A word in a cell is
  stored as *text*, and a bare word in a query is a *name* — they are not
  the same thing, so they do not match. If your table has a `Dept` column
  reading `eng`, write `(== D "eng")`, with the quotes. Without them the
  test quietly succeeds on every row for `\=`, and on none for `==`, and
  both answers are technically right — you asked about a name, not about
  the text in the cell. Numbers need no quotes; this applies to text only.
  The rule is not new — it is what stops a cell reading `Alice` being
  mistaken for a blank waiting to be filled in — but these four goals are
  the first ones that make comparing against a written-out value an
  everyday thing to do, so it is worth saying here.

  All four symbols are now reserved and can't be used as your own predicate
  names, and the refusal that says so lists them. A release check now holds
  three things to agreeing about that list: what is refused as a name, what
  the solver actually runs, and what the message tells you — so a name
  can't end up forbidden but inert, or advertised but not really reserved.

### Known open security items

**Closed this release:** `SEC.8`, `SEC.9`, `SEC.11` and `SEC.13` — see above.

**Still open:** `SEC.3`, `SEC.7`, and two from the 2026-09-08 code
review — `SEC.10` and `SEC.15`. In plain words:
the remembered raw-VBA consent record still lives inside the workbook
(`SEC.10`);
formulas a program writes are not screened for functions that reach the
network (`SEC.15`); and effects like sending mail still run without a
permission prompt (`SEC.7`). `SEC.8` narrows that last one — it gates on
where the *workbook* came from — but does not close it: a phrasebook loaded
into a workbook of your own still reaches those verbs unprompted.

**Assessed and accepted, not fixed:** `SEC.12`, `SEC.14`, `SEC.16` and
`SEC.17`. Each needs a precondition an ordinary install does not meet —
mostly an Excel setting that ships off and that Frazaro never asks you to
turn on. The reasoning for each, and what would reopen it, is written down
rather than left implied.

The authoritative lists, kept current in one place instead of copied into
every release: [`README.md`](../README.md) in plain words, and
[`docs/BETA_REARVIEW.md`](BETA_REARVIEW.md) with the file, line, fix and
disposition for each.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from — and treat a workbook someone sent you the
same way before you press Interpret.** Frazaro makes no network call and
does not update itself; check the README's *Known open security items*
when you return for a newer build. Vulnerability reports:
`docs/SECURITY.md`. Everything else: `docs/SUPPORT.md`.

## 0.5.2

### What changed

- **`SEC.1` — the `CallByName` reflection fallback removed from dynamic
  dispatch, built and owner-verified live.** Previously, any `.`-member
  name the interpreter didn't recognize natively fell through to VBA's
  own `CallByName`, reaching the real Excel object model with whatever
  name and arguments a phrasebook rule supplied — a heuristic, not a
  capability gate. A full-repo census found 25 members reached only
  through it, well beyond the G-TABLES surface this item's own design
  anticipated: plain cell `.value` reads, cross-sheet `.range` lookups,
  the `listobjects`/`listrows`/`listcolumns` chain, and several
  housekeeping macros. Each got its own fixed, audited native case before
  the fallback itself came out, so anything not on that list now refuses
  in words instead of reaching arbitrary late-bound dispatch. Two
  corrections found live, not assumed clean: a corpus-only first census
  missed `.value`/`.range` entirely (caught only because the host-test
  suite reads them, not the shipped phrasebook text), and a second pass
  then found every property already native for *writing* (`name`,
  `bold`, `size`, and 17 others) had never been native for *reading* —
  fixed by mirroring the write-side list into the read side in one pass
  rather than chasing each individually. `VLA_SELF-TESTS` pure 947/947,
  host 143/143; `VerifyReports` emitter 141/141, interpreter 141/141,
  both live in Excel. G-PIVOT never used this mechanism at all (it
  dispatches through dedicated runtime helpers); keyword-argument calls
  already had no fallback to begin with. Permissioned, declared
  capabilities with real external effect (`vlasendmail` today) are a
  separate, still-open item — see `SEC.7`. Full mechanism:
  `docs/BETA_REARVIEW.md`'s own SEC.1 entry.
- **`SEC.2` — `raw` behind explicit, per-phrasebook consent, built and
  owner-verified live.** A phrasebook using `raw` (literal VBA,
  previously unconsented) now shows a modal, naming the file, before it
  loads from disk; declining refuses the whole phrasebook, not just the
  `raw`-bearing rules. Two remembered scopes, an explicit choice rather
  than a silent default: *this workbook only* (safer — forging it needs
  write access to that one file) or *every workbook on this device*
  (more convenient, a wider target, named as such in the prompt
  itself). Gated at the file-path loader specifically, not the shared
  `EnglishLoadVocabularyText` primitive `VLA_Browser.bas`'s
  already-shipped host-free translate API calls directly and documents
  as never showing a dialog — an early draft got this wrong and only
  passed the purity ratchet on a technicality, caught before it
  shipped. No test-bypass toggle anywhere in the mechanism, by design.
  Full mechanism: `docs/BETA_REARVIEW.md`'s own SEC.2 entry.
- **`GO.6` — a working Load Phrasebook button, owner live-tested.**
  Found live while hand-verifying `SEC.2`: the only mechanism that
  technically loaded an external phrasebook (`VLA_IDE.IdeVocabPath`'s
  four candidate paths) was undocumented, built for internal
  edition/dev purposes, and REPLACED the base corpus by exact filename
  match rather than adding to it — no ribbon command existed for an
  org admin or community contributor to load their own. The new "Load
  Phrasebook" button calls `EnglishLoadVocabulary` directly, so it
  inherits `SEC.2`'s raw-consent gate and `G3`'s same-shape-collision
  refusal automatically, no second loading path. ADDS rather than
  replaces — `G3`'s existing cross-file override mechanism already
  resolves collisions between sources, `GO.1`'s ratified precedence,
  no interpreter change needed — persists per workbook (one
  `VLA_LoadedPhrasebooks` custom document property, an unbounded
  vbLf-joined list; no cap, the same as a source file's own import
  statements), and shows what's currently loaded after every load
  (`EnglishLoadedSourcesReport`). A moved or deleted remembered
  phrasebook is skipped with a note rather than blocking every other
  command; a genuine content collision still refuses, unchanged. Full
  mechanism: `docs/BETA_REARVIEW.md`'s own GO.6 entry.
- **The one `AS.1` gap closed: `paint cell {r:text}`.**
  `check_rule_coverage.ps1`'s first real report (after `GEXPANDERLINT.0`
  verticalized the phrasebook artifact) found a rule with zero
  test-success proofs. It had never worked: the rule called `vlacolr`
  (no "o"), a name that resolves nowhere in the shipped modules, and
  had no cell/range slot in its pattern at all, so it could never have
  painted a cell even with the spelling fixed. Corrected to the
  `set-fill-color` idiom every sibling color rule already uses
  (`pirate.vla`'s own "paint cell" rule confirmed the intended
  semantics) and given its missing test. 150/150 phrasebook rules now
  carry at least one proof.
- **`CO.4` — versions are now something Frazaro can compare, not just
  print.** Groundwork, with no button and nothing new to say yet: it is
  what a future phrasebook will be checked against when it declares a
  minimum version (`requires: version:0.5.1`), and what a generated
  module's own version stamp will mean once it carries one. The
  decision behind it is the substance. **The grammar's compatibility
  version is the release version you already see** — the one in Copy
  Feedback and in Add/Remove Programs — under the `MAJOR.MINOR.PATCH`
  rules this project already wrote down (`SD-14`). There is no second,
  separate "grammar version" to learn, and deliberately so: `PATCH` is
  *defined* as no observable change to what your existing sentences do,
  `MINOR` means something new became sayable, and `MAJOR` means a
  sentence that once shipped changed meaning or a program could behave
  differently after upgrading. Those rules already say everything a
  compatibility check needs, so inventing a parallel number would only
  create two things to keep in step. Scoping also corrected a
  long-standing internal misreading: `VLA_CORE_VERSION` looked like a
  grammar version and is not one — it, and 23 sibling constants, name
  the work item that last touched each module, and three modules
  legitimately share one value today. It is unchanged, and no
  compatibility check reads it. Ordering is plain numeric
  `MAJOR.MINOR.PATCH`, so `0.9.0` correctly sorts *before* `0.10.0`
  rather than after it the way plain text comparison would; a version
  that cannot be read (`banana`, or a `-beta` suffix, which this
  project has never used) is refused in words rather than being quietly
  treated as either satisfied or unmet, since both hide the typo.
- **`CO.6` — a written record of when each thing you can say became
  sayable.** New file, and unlike `CO.4` above this one is meant to be
  read: [`docs/GRAMMAR_SINCE.md`](GRAMMAR_SINCE.md) lists every phrasebook
  rule and every core form with the release it first **worked** in — 278
  entries, almost all of them `0.5.0`. It answers the question `CO.4`
  leaves open. Knowing that versions compare correctly does not tell you
  *which* version you need, and the version number alone is a loose
  answer: a release can go up because a ribbon button was added, which
  tells you nothing about whether a particular sentence will work. This
  file is the precise answer, and the thing a future `requires:`
  declaration will be checked against.

  **It records when a form first worked, not when it was first spelled** —
  a distinction with a real case in it. `paint cell` appears here at
  `0.5.2`, not `0.5.0`, even though the words shipped in `0.5.0`: as the
  `AS.1` entry above describes, that rule never once painted a cell. Dating
  it `0.5.0` would tell you a phrasebook using it runs on `0.5.0`, which is
  false. Entries are never edited afterwards, so a date that went in wrong
  would stay wrong — which is why the seed was checked against the actual
  release tags rather than taken from the generated phrasebook artifact.
  That check earned its keep: the artifact looked like it gained three
  rules in `0.5.1`, and it had not — those three were already sayable in
  `0.5.0` and the artifact had simply been stale, the same staleness the
  `0.5.1` notes below record fixing.

- **`F.10` — a phrasebook can now state what it needs, and is refused
  politely when it doesn't have it.** Write a line like
  `(requires-version "0.5.2")` at the top of a phrasebook and Frazaro
  checks it *before* loading anything from that file. If the build is too
  old you get a plain sentence — a phrasebook asking for `0.9.0` on this
  build is refused with "this phrasebook needs Frazaro 0.9.0 or newer;
  this is 0.5.2" — instead of rules that load and then mysteriously do
  the wrong thing. This is what
  [`docs/GRAMMAR_SINCE.md`](GRAMMAR_SINCE.md) above exists to be checked
  against: it tells a phrasebook author which version number to write.

  Nothing loads part-way. The check happens before the first rule is
  registered, so a phrasebook is either fully in or fully refused —
  and it does not matter where in the file the line sits.

  Two other kinds of requirement are understood and both currently
  **refuse**, on purpose rather than by omission.
  `(requires-capability "...")` is the permission system that is not
  built yet (`SEC.7`, below): since nothing can grant a capability, the
  honest answer to a phrasebook asking for one is no. `(requires-form
  "...")` — needing one specific sentence rather than a whole version —
  is understood but cannot be checked yet, because the record above
  lives in the source repository and is not carried inside Frazaro
  itself. A requirement Frazaro doesn't recognise at all is also
  refused: it cannot confirm the requirement is met, so it does not
  pretend to. Note this is a phrasebook *declaring* what it needs; it is
  not yet a restriction on what the verbs themselves may do — see the
  `SEC.7` item below, which is unchanged by this release.

  Because a version requirement is only as good as the record it is
  written against, the release checks now also refuse to let a form
  reach a release with no row in
  [`docs/GRAMMAR_SINCE.md`](GRAMMAR_SINCE.md). If you are writing a
  phrasebook, that means the version number you look up there covers
  every sentence Frazaro understands, not just the ones someone
  remembered to record.

### Known open security items

- **SEC.3** — generated code does not yet carry phrasebook provenance.
- **SEC.7** — a small set of verbs with real external effect (`vlasendmail`
  today) still runs with no permission check: a phrasebook you load can
  send mail with no consent prompt. `SEC.8` above narrows this but does
  not close it — it gates on where the *workbook* came from, not on what
  a phrasebook asked permission to do, so a phrasebook loaded into a
  workbook of your own still reaches these verbs unprompted.

`SEC.1`, `SEC.2` closed this release — see above.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from.** Frazaro makes no network call and does not
update itself; check the README's *Known open security items* when you
return for a newer build. Vulnerability reports: `docs/SECURITY.md`.
Everything else: `docs/SUPPORT.md`.

## 0.5.1 — 2026-09-07

**Pre-flight for the corporate push.** The first release cut by
`tools/release.ps1` rather than by hand, and the day the static ratchets
went green again. Nothing here changes what a sentence means; every item
is the machinery that keeps the next four weeks honest.

### What changed

- **The static ratchets pass, and are now run on every push.** `0.5.0`
  shipped past a red `check_raise_ratchet.ps1` (five raw `Err.Raise` sites
  had accumulated since the ceilings were set) and a `check_id_registry.ps1`
  that misread four mid-sentence bold ids as duplicate definitions. The
  four genuinely raw refusals in `VLA_Relation`/`VLA_Sql` now carry
  catalogue ids (`relation-table-noncontiguous-areas`,
  `sql-internal-*`); the one re-raise with no id to carry is documented as
  the eighth site of `VLA`'s ceiling; the id registry counts a bold id as
  a definition only when no prose precedes it on its line.
- **Three ribbon buttons find the phrasebook again** (EDITION-VOCABPATH).
  *Translate to VLA / to VBA*, *Export Expanded Phrasebook* and *Phrasebook
  Test Coverage* had reported "vocabulary file on disk... none found" in
  the dev workbook since the phrasebook moved under `scripts/polyglotta/`.
  They now search that folder as a last candidate; Check/Compile's own
  embedded-chain loading is byte-for-byte unchanged.
- **The expanded-phrasebook staleness stamp ignores whitespace.** It was
  the source file's byte size, so a Lint VLA re-indent or a line-ending
  change made `check_rule_coverage.ps1` demand a re-export, and Beta's
  history holds four commits that changed nothing but that number. It is
  now a hash of the source's non-whitespace bytes, computed identically
  in VBA and PowerShell; only a token change counts as stale.
- **`scripts/english_expanded.vla` re-exported** against the current
  corpus, so `check_rule_coverage.ps1` reports on the phrasebook that
  actually ships. It now lives beside its source under `scripts/polyglotta/`.
- **`release.ps1` refuses the placeholder** release-notes section instead
  of publishing it.

### Known open security items

Unchanged from `0.5.0`, and still the first engineering item in the queue:

- **SEC.1** — dynamic dispatch is not yet capability-gated: a phrasebook you
  load can reach roughly what a macro in a workbook you open could reach.
- **SEC.2** — `raw` phrasebook rules (literal VBA) run without a consent
  prompt. Refused on the default Interpret path; executes only through
  Compile/Export, behind Excel's own VBA-project trust prompt.
- **SEC.3** — generated code does not yet carry phrasebook provenance.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from.** Frazaro makes no network call and does not
update itself; check the README's *Known open security items* when you
return for a newer build. Vulnerability reports: `docs/SECURITY.md`.
Everything else: `docs/SUPPORT.md`.

## 0.5.0 — 2026-09-05

**A phrasebook for your workbook.** Frazaro lets people who are not
programmers automate Excel by writing English sentences that are *checked*
before they run, refuse with an explanation when they don't parse, and mean
exactly one thing when they do. This is a known-unfinished program, released
unfinished on purpose, because the only way to learn which sentences real
people reach for is to let real people reach for them. The open items are
counted, in order, in `docs/BETA_ROADMAP.md`.

### Download

| File | Choose this if |
|---|---|
| `Frazaro_English.xlam` | The universal default. An Office document, not a program: works wherever Excel does, including managed machines. |
| `Frazaro_Espanol.xlam` | The Spanish edition (bilingual: every English sentence still works). |

Register the add-in once from inside the product (**Register for auto-load**
on the Frazaro ribbon) or through Excel's Add-ins dialog. Removal is one
button, **Uninstall Frazaro**. Full instructions: `docs/DEPLOY.md`. The
Windows installer is not part of this release; it returns in the next one.

### What works, owner-verified

The full English → VLA → interpret/compile pipeline; the worksheet IDE
(Check / Run / Undo / Known Sentences); control flow, value actions,
lookups, list and range operations; the four-layer error model; `=SQL()`,
`=DATALOG()` and `=PROLOG()` worksheet functions over your own tables. All
under a golden-file, self-test and dual-backend-parity discipline.

### Honest limits

Windows desktop Excel is the platform; Mac runs the interpreted path but
cannot export VBA. Grammar coverage is growing weekly and is not yet
complete for any one profession's vocabulary. No external users yet: the
project is looking for its first pilot.

### Known open security items

Frazaro's own threat model (`docs/THREAT_MODEL.md`) is public, and three of
its findings are open in this release:

- **SEC.1** — dynamic dispatch is not yet capability-gated: a phrasebook you
  load can reach roughly what a macro in a workbook you open could reach.
- **SEC.2** — `raw` phrasebook rules (literal VBA) run without a consent
  prompt. Refused on the default Interpret path; executes only through
  Compile/Export, behind Excel's own VBA-project trust prompt. First item
  after this release.
- **SEC.3** — generated code does not yet carry phrasebook provenance.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from.** The phrasebooks embedded in these downloads
are audited at build time. Frazaro makes no network call and does not update
itself; check the README's *Known open security items* when you return for
a newer build. Vulnerability reports: `docs/SECURITY.md`. Everything else:
`docs/SUPPORT.md`.

### Licence

Engine Apache-2.0; phrasebooks MPL-2.0 (a phrasebook you write in your own
file is yours); the one module Frazaro copies into your workbook is 0BSD;
everything Frazaro generates from your sentences is yours outright
(`OUTPUT-EXCEPTION.md`). No warranty: beta software, as-is. Take a backup
before you run a program you have not run before.
