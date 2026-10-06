# The command-line door

*Slices 1 to 4 of `docs/HORIZON.md` section 12 through their first door:
`frazaro`, one executable with no dependencies, no network and no Excel. The
add-in in `src/` is the reference implementation of the language; this door
is the second implementation's, held to the same corpus by the treaty in
`conformance/README.md`, so a sentence means the same thing here as on the
Frazaro tab, and refuses in the same words.*

This page is the reference: how to build the door, the shape every command
shares, a tour on the repository's own files with the line each step prints,
then each command with its arguments, output, exit codes and refusals. The
`USAGE` text in `cli/src/main.rs` is the short form (`frazaro help`); when a
command changes, both change.

## Building it

```powershell
cargo build --workspace
target\debug\frazaro.exe version
```

prints `frazaro 0.7.1 (core abi 1)`: the core's version, which is the
add-in's (`check_version_twin.ps1` holds them equal), and the number a page
checks before calling the core. If `cargo` is not on the PATH it is at
`%USERPROFILE%\.cargo\bin\cargo.exe`; `rust-toolchain.toml` pins the
toolchain. `cargo build --release --workspace` puts the same program, faster,
at `target\release\frazaro.exe`. Below, `frazaro` stands for whichever you
built; the commands are written from the repository root.

## The shape every command shares

- **Files in, text out.** The door reads the files it is named and prints
  the answer to stdout. The core sees text, or a workbook's bytes, and gives
  text back; the one file the door ever writes is `build`'s `--out`.
- **`--prelude <prelude.vla>`**, the standard library, `scripts\prelude.vla`.
  Every command that translates needs it; `prove` finds it beside the
  phrasebook unless told otherwise.
- **`--phrasebook <file.vla>`**, repeatable, in order: `english.vla` first,
  then an edition's dialect, then an organization's own, as the add-in loads
  them. Each loads with its proofs run, and a failing proof refuses the load,
  as Load Phrasebook refuses it on the tab.
- **`--allow-raw`**, the consent a phrasebook with a `(raw ...)` form needs
  (`SEC.2`): the same checkbox the add-in and the web page ask. Without it
  such a phrasebook is refused by name. A phrasebook that requires a
  capability is refused always, since none can be granted yet.
- **Exit codes.** 0, the answer is on stdout. 1, a refusal: its words, the
  message catalogue's, went to stderr, and nothing was written. 2, a usage
  error or a file that could not be read (the usage line, or
  `frazaro: cannot read ...`). 3, the treaty's oracle is not attempted in this
  version (today only `prove` on an engine proof file).
- **Files** are UTF-8; a byte-order mark is dropped; line endings are left as
  they are, since the core and the add-in both leave them alone.

## A ten-minute tour, on the repository's own files

Each step names what it prints, so the tour doubles as a check that the door
you built is the one this page describes.

1. **The version.** `frazaro version` prints `frazaro 0.7.1 (core abi 1)`.

2. **A program's VLA.** The build golden's fixture is seven sentences:

   ```powershell
   frazaro translate-vla scripts\build\fixture.txt --prelude scripts\prelude.vla --phrasebook scripts\polyglotta\english.vla
   ```

   prints the whole program's VLA, `(dim vla-step Long)` and `(sub main ()`
   first, each sentence under its line, `(at-line 4 (set! (range "b2") 5))`
   for `Put 5 into cell B2.`; the text `EnglishToVla` writes in Excel,
   character for character.

3. **Its VBA.** The same arguments to `frazaro translate-vba` print
   `Option Explicit`, `Dim vla_step As Long` and `Public Sub main()` onward:
   what Compile and Run writes into a module.

4. **A workbook.**

   ```powershell
   frazaro build scripts\build\fixture.txt --prelude scripts\prelude.vla --phrasebook scripts\polyglotta\english.vla --out tour.xlsx
   ```

   prints `wrote tour.xlsx: 11390 bytes, sha256 18F69AA3...EDFADBC4` and
   nothing else. Those are the bytes of `scripts\build\fixture_golden.xlsx`,
   the golden the owner opened in Excel: a `Frazaro` sheet with the
   sentences in column B and OK beside each, `Output` with the values and
   formulas, `data` with its one cell. Open it; nothing asks to be enabled.

5. **The workbook proves its build.**

   ```powershell
   frazaro rebuild tour.xlsx --prelude scripts\prelude.vla --phrasebook scripts\polyglotta\english.vla
   ```

   prints `This workbook was built from these 9 sentences by Frazaro 0.7.1:
   yes.` and exits 0. Save the file from Excel and ask again: it is refused
   as a host's save, since Excel rewrites every part, by design.

6. **The door writes one file and guards it.** Run step 4 again:
   `Refusing to write tour.xlsx: a file is already there, and it has not
   been changed. Add --replace to replace it, or name another file with
   --out.`, exit 1.

7. **Into a workbook of your own.**

   ```powershell
   frazaro build scripts\build\into.txt --prelude scripts\prelude.vla --phrasebook scripts\polyglotta\english.vla --into scripts\build\model.xlsx --out tour-into.xlsx
   ```

   prints `wrote tour-into.xlsx: 13697 bytes, sha256 4281250F...710DF798`:
   the model's own sheets `Model` and `Notes` copied byte for byte, the
   sentences' `Frazaro`, `Output` and `checks` added, `Output!B1` reading
   `=Model!B3*Rate`. `frazaro rebuild tour-into.xlsx ...` prints `This
   workbook's Frazaro sheets were built from these 6 sentences by Frazaro
   0.7.1, into a workbook whose own sheets are not checked: yes.`

8. **The workbook read back as relations.**

   ```powershell
   frazaro reflect tour.xlsx
   ```

   prints what the file holds, one form a line, in an order that never
   changes: `(sheet "Frazaro" visible)` and its two sister sheets first,
   then the `Frazaro.Build` name with its fingerprints, then the sentences
   as `(cell "Frazaro" "B1" "...")` rows with an `"OK"` cell beside each,
   then `Output`'s two values, five formulas and what each refers to,
   `(formula "Output" "C3" "=B3+B4")` followed by `(refers "Output!C3"
   "Output!B3")` and `(refers "Output!C3" "Output!B4")` among them, since a
   filled formula shows in each cell as Excel shows it, and `(cell "data"
   "A1" "west")` last: 38 lines, which are
   `scripts\reflect\build_fixture_relations.vla` exactly. No formula has a
   `cell` row, because nothing has computed one yet; save the file from
   Excel and ask again, and each gains one. `frazaro reflect tour.xlsx
   --counts` prints the sheets' counts and times alone, and `frazaro
   reflect tour.xlsx --cone Output!C2` one line of counts for C2's cone:
   `cells 3 formulas 2 inputs 1`, C2 itself, B3 that it adds and B2 that
   both read.

9. **What changed between two workbooks.**

   ```powershell
   frazaro diff scripts\build\fixture_golden.xlsx scripts\reflect\saved.xlsx
   ```

   compares the build golden with the copy Excel saved of it: six lines,
   `(changed "Frazaro.Build" ...)` for the stamp Excel rewrote as
   `_xlfn._LONGTEXT`, then `Output`'s five formulas, each `(formula
   "=B2+B3")` on the left and `(formula "=B2+B3" 15)` on the right, the
   value Excel computed and the file now holds; nothing else differs, which
   is what saving a file should do. `frazaro diff tour.xlsx
   scripts\build\fixture_golden.xlsx` prints nothing: the same bytes.

10. **Where the risks are.** `frazaro audit tour.xlsx` prints three
    lines: `(empty-reference "Output!C3" "Output!B4")`, then C4's to B4
    and to B5. The sentence filled `=B2+B3` down rows 2 to 4 of column C,
    past the two values in column B, so two of the three formulas add
    cells nothing holds. Nothing else: no constant typed over a formula,
    no inconsistent formula, no unused name, no hidden sheet, no link.
    `frazaro audit tour.xlsx --counts` prints the six counts alone.

11. **A sentence the language refuses.** A file holding `Put 5 into cell
   B2.` and `Set total to $5.` through `translate-vla` prints, on stderr,
   `I don't understand the character '$' - write the plain number (or
   'Format ... as currency.' for display) (line 2)` and exits 1: the
   add-in's words, with the line.

12. **A sentence the writer cannot hold.** A file holding `Put 5 into cell
   B2.` and `Say "hello".` through `build` prints `frazaro build writes
   only what a sheet can hold with nothing running: a value or a formula
   into a cell, a range or rows of a column, and which sheet it goes on.
   Line 2 asks for more: Say "hello". Nothing was written; run the program
   with the add-in, or take the sentence out.` and exits 1. The same
   program through `translate-vla` is fine: the sentence is good English
   that needs the add-in's Run.

13. **A phrasebook, proved and loaded.** `frazaro prove
    scripts\polyglotta\english.vla` prints `PASS 482/482`; `frazaro load
    scripts\polyglotta\english.vla --prelude scripts\prelude.vla` prints
    `loaded: 240 rules, 220 macros, 460 tests (22 expected fails) from
    scripts\polyglotta\english.vla`, the line Load Phrasebook shows.

## The commands

### `frazaro translate-vla` and `frazaro translate-vba`

```text
frazaro translate-vla <program.txt> --prelude <prelude.vla> --phrasebook <file.vla> [--phrasebook ...] [--allow-raw]
frazaro translate-vba <program.txt> --prelude <prelude.vla> --phrasebook <file.vla> [--phrasebook ...] [--allow-raw]
```

The program is a text file of sentences, one per line, a comment line
starting with `#`, a block indented by two spaces, exactly what the Frazaro
sheet's column B or a `<Frazaro>` section holds. `translate-vla` prints the
program's VLA, the text `EnglishToVla` writes: the treaty's oracle 1, which
on `scripts\instructions.txt` is `scripts\instructions_golden.vla` byte for
byte. `translate-vba` compiles that VLA with the prelude and prints the VBA
the add-in would write into a module: oracle 1a, the `.vba` golden.

A refusal is the sentence's own, with its line, on stderr, exit 1. Files it
cannot find are refused through the catalogue: `Program file not found: ...`
(`english-program-file-not-found`), the prelude (`vla-file-not-found`) and a
phrasebook (`english-vocab-file-not-found`). At least one `--phrasebook` is
required (exit 2 without). On the tab: Validate Instructions and Translate
File to VLA; Show me the VBA and Translate File to VBA.

### `frazaro build`

```text
frazaro build <program.txt> --prelude <prelude.vla> --phrasebook <file.vla> [--phrasebook ...] --out <file.xlsx> [--into <model.xlsx>] [--replace] [--allow-raw]
```

The program is translated exactly as `translate-vla` translates it, then
written as a workbook at `--out`, the one file the door writes. The
`Frazaro` sheet holds the sentences in column B, exactly as written, with OK
beside each non-blank line in column C, in the room's own layout and
colours, so the file is a live workspace the moment the add-in is loaded.
The other sheets hold what a sheet holds with nothing running: a value or a
formula into a cell, a range, or rows of a column, on the sheet the program
names (`Work on sheet Data.`, `... of sheet Output`); a filled formula is
stored shared, a newer function under its file prefix (`_xlfn.`), a formula
whose function can return an array (`IFS`, `FILTER`, `INDEX` and their kin)
as Excel stores one typed into a cell, so no `@` appears; the host computes
every formula when it opens the file. The workbook carries its stamp, the
defined name `Frazaro.Build`: the core's version and, for the sentences, the
prelude and each phrasebook, the SHA-256 of their non-whitespace bytes.

The bytes are deterministic: the same sentences and the same core give the
same file on every machine (stored entries stamped 1980-01-01, a fixed part
order, no author and no date in the properties), which is what makes
`scripts\build\fixture_golden.xlsx` a golden and the page's download the
same file. On success the door prints `wrote <out>: <n> bytes, sha256
<hex>`, the digest the web page shows beside its Download button.

With `--into <model.xlsx>`, the sentences' sheets are added to a workbook of
yours and written to `--out`; the model itself is never written. Every part
of the model is copied as the bytes it already is; only the workbook's sheet
list, names and calculation flag, its relationships, the content types and
the ends of its style lists are edited, so the model's own cells, strings,
theme and calculation chain are exactly what they were. A sentence may go to
one of the model's sheets but not write into it; new sheets' formulas read
the model's cells and names as any formula would. Building into a workbook
an earlier build added sheets to replaces those sheets.

Refusals, all exit 1 with nothing written: a file already at `--out`
without `--replace` (`build-output-exists`); a sentence that needs the
add-in's Run, quoted with its line (`build-not-representable`); a sheet a
sentence names that the program never made (`build-sheet-unknown`), or a
sheet name Excel would not take (`build-sheet-name-invalid`); a workbook too
large for the container (`build-workbook-too-large`); with `--into`, a file
that is not a workbook, an older `.xls` or an encrypted one, zip64, a broken
directory (`build-into-not-a-workbook`), a model that already has a
`Frazaro` sheet (`build-into-sheet-name-taken`), a sentence writing into one
of the model's own sheets (`build-into-model-sheet`), and a model that
cannot take a dynamic-array formula (`build-into-unsupported`). The
translation's own refusals come first, as in `translate-vla`.

On the tab there is no one button for this: Compile and Run and Interpret
and Run run the program against a live workbook, while `build` writes what
a sheet holds without running anything. The treaty's oracle 7 holds it:
`check_build_golden.ps1` and the core's tests require the two goldens byte
for byte.

### `frazaro rebuild`

```text
frazaro rebuild <file.xlsx> --prelude <prelude.vla> --phrasebook <file.vla> [--phrasebook ...] [--allow-raw]
```

Reads the sentences back out of a built workbook's `Frazaro` sheet, checks
the stamp against them and against the prelude and phrasebooks given, builds
the workbook again and compares the bytes whole, then prints one line:
`This workbook was built from these N sentences by Frazaro <version>: yes.`
(exit 0) or `... : no, <why>.` (exit 1), the why naming what differs: a
phrasebook given that is not the one the stamp names, a part whose bytes
changed. For a workbook built `--into` a model the line says `This
workbook's Frazaro sheets were built from these N sentences ..., into a
workbook whose own sheets are not checked`: the build's own sheets are
rendered again from what the stamp recorded of the model and compared part
by part; the model's own parts cannot be remade without the model.

A file that is not a build is refused (`rebuild-not-a-build`, exit 1) with
what it is: not a zip; parts no longer stored, so a host has saved it since,
which Excel does on every save; no stamp. That refusal is the design, not a
gap: a workbook anyone has saved is no longer the build. The checks run
`rebuild` on each golden and on what the door builds.

### `frazaro reflect`

```text
frazaro reflect <file.xlsx> [--counts] [--cone <Sheet!A1> ...]
```

The reader (`PORT.8`), slices 8a and 8b. The workbook's file is read, never
Excel: its sheets, names and Tables, then every cell, sheet by sheet in tab
order and cell by cell in row order, printed as relations in the proof
corpus's notation, one form a line, in an order that never changes, so that
the output is a golden and two readings of one file compare line for line:

```text
(sheet "Model" visible)
(sheet "Scratch" hidden)
(name "Rate" "0.2")
(table "Sales" "Data" "A1:B4")
(cell "Model" "B1" 1200)
(cell "Model" "B3" 400)
(formula "Model" "B3" "=B1-B2")
(refers "Model!B3" "Model!B1")
(refers "Model!B3" "Model!B2")
```

A `sheet` row's state is `visible`, `hidden` or `very-hidden`. A `name` row
holds what the name refers to as the file does, a name scoped to one sheet
spelled `Sheet!Name`; Excel's own `_xlfn.` placeholders are left out. A
`cell` row holds the value the file holds: a number as the file spells it,
a text, `true` or `false`, `(error "#DIV/0!")`, or `(date "...")` for an
ISO date cell; a date entered in Excel is its serial number. A `formula`
row holds the formula as the formula bar shows it, `=` first, the file's
`_xlfn.` prefixes dropped and a link to another workbook named by its file
(`=[Rates.xlsx]Sheet1!A1`, the spelling Excel shows while that book is
open; closed, Excel shows the path it resolved, which the file does not
hold); a formula filled down shows in each cell with
its references moved, as Excel shows it; an array formula shows in its
first cell, and the cells it spills into are values. After a formula come
its `refers` rows, one per distinct thing it refers to, sorted: a cell or a
range qualified by its sheet with its `$` marks dropped (`Model!B1`,
`Data!A:A`), a name, a structured reference, a 3D span or a reference into
another workbook as written (`Rate`, `Sales[Amount]`, `Model:Scratch!B1`,
`[Rates.xlsx]Sheet1!A1`), a spill as its cell with `#`, a `#REF!` as
`Model!#REF!`, and `(unreadable "INDIRECT")` or `(unreadable "OFFSET")`
for the two calls whose targets no reader can know from the text; the
reading inside such a call goes on, so `=OFFSET(B1,1,0)` refers to
`Model!B1` and is unreadable both. This is `AXM.7`'s reader,
`src/VLA_Refers.bas`, ported arm by arm and held to its golden. A formula
cell has a `cell` row only when the file holds its last computed value, so
a workbook `frazaro build` wrote and nothing has opened has `formula` rows
alone for its formulas, and the same workbook saved from Excel gains a
`cell` row for each. A formatted cell with nothing in it has no row. Two
files are compared by `diff` and the audit questions asked by `audit`,
below, both over this reading.

With `--counts`, nothing from inside the workbook is printed: one line a
sheet, by position, with its counts of cells, formulas, distinct formulas
(compared in R1C1 relative to their cells, so that a formula filled down
counts once), unreadable calls (`INDIRECT` and `OFFSET`) and array
formulas, the last row and column holding anything, the part's size and
the time it took to read, then one line for the workbook with its totals,
names, Tables, links to other workbooks and strings. This is for measuring
a model whose contents must not leave the machine; the lines can be pasted
anywhere.

With `--cone Sheet!A1`, once or more, a cell's cone is sized through the
file, one line a root with counts alone: the cells reached through every
reference, through names and Tables and across sheets, how many of them
hold formulas and how many are inputs, the blank cells referred to, the
sheets touched, the depth of the longest chain, where the cone is blind
(`INDIRECT` and `OFFSET`), the names and Tables resolved, the references
into other workbooks (counted, not followed), the `#REF!`s met and what the
file does not have. Excel's own Trace Precedents stops at the sheet
boundary; this does not. The root is the one address printed, because you
typed it. A root that is not a cell as `Sheet!A1` or `'Q1 Data'!A1`, or one
naming no sheet, exits 2.

Refusals, exit 1: a file that is not a workbook, with why
(`reflect-not-a-workbook`: not a zip, an older `.xls` or an encrypted one,
zip64, no workbook part, a damaged part, an OpenDocument file, which a later
slice reads); a part this version refuses to read (`reflect-xml-refused`:
one declaring a `DOCTYPE` or an entity, which the reader never expands, or
markup that never closes); a shape this version does not read, named with
its part and cell (`reflect-unsupported`: a data-table formula, a shared
formula whose first cell is missing, a string index past the table). The
reader streams, so the rows printed before a refusal stand on stdout.

On the tab there is nothing yet: `REFLECT` over a live workbook is
`AXM.8`'s, and will read through Excel what this reads from the file. The
treaty's oracle 8 holds it: `check_reflect_golden.ps1` and the core's tests
require each fixture's relations whole.

### `frazaro diff`

```text
frazaro diff <old.xlsx> <new.xlsx> [--counts]
```

The reader's third slice (`PORT.8`, 8c). Two workbooks' files are read as
`reflect` reads them and compared, and what differs is printed one form a
line, in an order that never changes, so that the output is a golden too:

```text
(sheet-removed "Secret" very-hidden)
(sheet-added "Audit" visible)
(sheet-changed "Scratch" hidden visible)
(changed "Rate" "0.2" "0.25")
(changed "Sales" "Data!A1:B4" "Data!A1:B5")
(changed "Model!B1" 1200 1300)
(changed "Model!B3" (formula "=B1-B2" 400) 500)
(changed "Model!C1" (formula "=B1*2" 2400) (formula "=B1*2" 2600))
(changed "Model!H1" blank 1)
```

Sheets are matched by name, without case. A sheet in one file alone is one
row, with its state, and its cells are not listed, so a renamed sheet is a
removal and an addition; a matched sheet whose state changed is a
`sheet-changed` row. Every other difference is a `changed` row with the
old and the new side by side: first the names and Tables, by name (a
name's sides are what it refers to, a Table's its sheet and range, `blank`
where that file has none), then sheet by sheet in the new file's tab order
and cell by cell in row order, the cell spelled as a reference spells it,
each side a value as `reflect` prints one, `(formula "...")` for a formula
whose cached value the file does not hold, `(formula "..." value)` for one
it does, or `blank`. So a formula typed over by a constant reads
`(changed "Model!B3" (formula "=B1-B2" 400) 500)`, and a formula left
alone whose result moved reads as two values under one text. A cell counts
as changed when its formula's text or its value differs; numbers are
compared as numbers, so `800` and `800.0` are one value, and everything
else by kind and text, so a text `"5"` typed over the number `5` is a
change. Two files that hold the same sheets, names, Tables and cells print
nothing. The exit code is 0 whenever the comparison ran, rows or none: the
rows are the answer.

With `--counts`, one line and nothing from inside either file: `diff:
sheets-removed 1 sheets-added 1 sheets-changed 1 names 3 tables 1 cells 17
compared 42 open 0.9 ms compared 1.8 ms`, `compared` being the addresses
either file holds on the matched sheets.

Refusals, exit 1, are the reader's (under `reflect`, above), raised for
whichever file raised them; the rows printed before one stand. Two files
are read and nothing is written. On the tab there is nothing yet: `DIFF`
over the live workbook, its Undo snapshot and a saved copy is `AXM.11`'s.
The treaty's oracle 9 holds it: `check_diff_golden.ps1` and the core's
tests require each pair's difference whole, the build golden against its
Excel-saved copy among the pairs.

### `frazaro audit`

```text
frazaro audit <file.xlsx> [--counts]
```

The reader's fourth slice (`PORT.8`, 8d): the audit list, where a
workbook's risks are, read from the file. Six questions, each a walk over
what `reflect` reads, one finding a line, in an order that never changes:

```text
(typed-over "Review!B3" 61 "=A3*2")
(inconsistent "Review!C3" "=A3+2" "=A3+1")
(unused-name "Range1" "Data!$A$2:$B$4")
(empty-reference "Review!D1" "Review!Z9")
(hidden-sheet "Scratch" hidden)
(external-link "Model!E2" "[Rates.xlsx]Sheet1!A1")
```

A **typed-over constant** is a value whose nearest neighbours above and
below in its column both hold formulas that agree in R1C1; the row shows
the value and the formula those neighbours would have put there. An
**inconsistent formula** differs in R1C1 from two such neighbours that
agree; the row shows the formula as written and the neighbours'. A cell at
the top or bottom of a column, or between neighbours that disagree, is
never reported, so a finding is one worth acting on. An **unused name** is
a defined name no formula and no other name refers to (Excel's own
`_xlnm.` names, print areas and filter databases, and Frazaro's own marks,
the build stamp and the add-in's `VLAt_` names, are never reported). An
**empty reference** is a formula's reference to one cell, on a sheet the
file has, that holds nothing: a formatted blank counts, a range is never
expanded. A **hidden sheet** is what its `sheet` row says, and an
**external link** is a formula that reaches another workbook. The order is
those six, and inside each sheet by sheet in tab order and cell by cell in
row order, names by name. A workbook with nothing to report prints
nothing, and the exit code is 0 either way: the findings are the answer.
The totals that do not foot wait for a later slice.

With `--counts`, one line and nothing from inside the file: `audit:
typed-over 1 inconsistent 1 unused-names 2 empty-references 2
hidden-sheets 2 external-links 1 open 0.9 ms indexed 1.4 ms walked 0.3
ms`, the number internal audit wants of a model it may not show.

Refusals, exit 1, are the reader's (under `reflect`, above); the whole
file is read before the first finding, so nothing is printed before one.
On the tab there is nothing yet: these questions asked of the live
workbook in English are `AXM.10`'s. The treaty's oracle 10 holds it:
`check_audit_golden.ps1` and the core's tests require each fixture's
findings whole, the build golden's three empty references among them.

### `frazaro load`

```text
frazaro load <phrasebook.vla> --prelude <prelude.vla> [--allow-raw]
```

Loads a phrasebook through the core's loader as Load Phrasebook loads it,
with the door's two gates first: a required capability refuses always, and a
`(raw ...)` form refuses unless `--allow-raw` gives the consent. Every proof
in the phrasebook runs at load, and a failing one refuses the load with its
words. On success the line `EnglishVocabStats` prints: `loaded: 240 rules,
220 macros, 460 tests (22 expected fails) from <file>`, followed by the
grammar's lint warnings when there are any. Exit 0, or 1 on a refusal.

### `frazaro prove`

```text
frazaro prove <phrasebook.vla> [--prelude <prelude.vla>] [--allow-raw]
```

The treaty's oracle 3: every proof in the phrasebook is run as loading it
in Excel runs it, every failure is printed as the add-in would have refused
it, and the last line is `PASS n/n` (exit 0) or `FAIL k/n`, k passed of n
(exit 1). The prelude is `prelude.vla` beside the file or in its parent
folder, which is how `scripts\polyglotta\english.vla` finds
`scripts\prelude.vla`; `--prelude` names another. A load that refuses before
its proofs (a malformed rule, a library file with no rule and no proof)
prints the refusal and exits 1 with no verdict line. An engine proof file
(`scripts\proofs\datalog.vla`) is not attempted in this version: `frazaro:
'prove' attempts a phrasebook's proofs; engine proofs are not attempted in
this version (PORT.9)`, exit 3, which the conformance runner reads as *not
attempted*, never as passed. Today `english.vla` is `PASS 482/482` and each
dialect passes its own.

### `frazaro compile`

```text
frazaro compile <program.vla> --prelude <prelude.vla>
```

A program already written in VLA, compiled with the prelude: its VBA to
stdout, the text Compile and Run would write into a module. The treaty's
oracle 1b: `scripts\instructions_golden.vla` less its stamp line compiles to
`scripts\instructions_golden.vba` whole. A missing program is
`vla-source-not-found`, a missing prelude `vla-file-not-found`; a compile
refusal is the compiler's own, exit 1.

### `frazaro version`, `frazaro help`

`version` (also `--version`, `-V`) prints `frazaro <version> (core abi N)`.
`help` (also no arguments, `--help`, `-h`) prints the usage text. A word that
is not a command exits 2 with `frazaro: '<word>' is not a command in this
version; 'frazaro help' lists the commands.`

## Not in this version

`check`, `run` and `ask` wait for their slices of
`docs/HORIZON.md` section 12: `ask` for the question act and the engines
(`AXM.9`, `PORT.9`), `run` for the interpreter over the sheet model
(`PORT.10`), `check` for the GitHub Action among the doors (`PORT.11`). The
usage text lists them under *Not in this version* so a reader who guesses
the command learns where it is.

## The same door on the Frazaro tab

| On the tab, or in the dev workbook | At the prompt |
|---|---|
| Validate Instructions (a refusal names its line) | `frazaro translate-vla` |
| Translate File to VLA, Translate File to VBA | `frazaro translate-vla`, `frazaro translate-vba` |
| Show me the VBA | `frazaro translate-vba` |
| Compile and Run, Interpret and Run | nothing runs here; `frazaro build` writes what a sheet holds without running |
| Load Phrasebook, and the consent it asks for a `(raw ...)` form | `frazaro load`, `--allow-raw` |
| the proofs Load Phrasebook runs | `frazaro prove` |
| Open CLI, one sentence at a time | the shell; the door takes files, so a one-line file is the sentence |
| nothing yet: `REFLECT` over the live workbook is `AXM.8`'s | `frazaro reflect`, from the file |
| `? VlaSelfTests()` in the Immediate pane | `cargo test --workspace`, then `powershell -File tools\prove.ps1 -Impl target\debug\frazaro.exe` |

## What holds the door

`tools\prove.ps1 -Impl target\debug\frazaro.exe` scores the door against the
treaty: oracle 1 and 1a on the corpus program, 1b on the compile golden,
oracle 3 on every phrasebook, oracle 7 on both build goldens with `rebuild`
run on each, oracle 8 on the six reflect fixtures, oracle 9 on the three
diff pairs, oracle 10 on the three audit rows; `25 passed, 0 failed, 2
not attempted, 1 library` today, the two not attempted being the
interpreter golden (`PORT.10`) and the engine proofs (`PORT.9`). `-Control`
first proves the runner on a fake door and a mutant. Five checks hold
floors on it: `check_compile_prefix.ps1`, `check_translate_prefix.ps1`,
`check_prove_floors.ps1`, `check_build_golden.ps1` and
`check_reflect_golden.ps1`, all taking `-Impl`;
the `core` job of `.github/workflows/checks.yml` builds the workspace,
formats, lints and tests it, builds the wasm, and runs every one of them on
every push. Before a change to the door is committed: `cargo fmt --all --
--check`, `cargo clippy --workspace --all-targets -- -D warnings`,
`cargo test --workspace`, and `powershell -File tools\run_checks.ps1`.

The web page, `web/README.md`, is the same core compiled to WebAssembly
behind a second door: its rows are the program, its *Download as .xlsx* is
`frazaro build` with the digest shown beside the button, and `frazaro
rebuild` says yes to what it downloads.
