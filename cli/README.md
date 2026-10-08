# frazaro

**The Frazaro command line: English sentences in, spreadsheets out, with
nothing installed.**

Frazaro is a language for spreadsheet work. You write the steps of a
procedure as English sentences; Frazaro checks every sentence, refuses the
ones it cannot read, in words, and carries out the rest. In Excel it is an
add-in. Here it is one executable, `frazaro`, which

- translates a text file of sentences into VLA, the language's own form, and
  into the VBA the add-in would run;
- builds an `.xlsx` workbook from the sentences, the same bytes on every
  machine, and proves later that a workbook is that build;
- reads any workbook back as relations, says what differs between two
  workbooks, and lists where a workbook's risks are, all from the file, with
  no Excel on the machine;
- loads and proves a phrasebook, the file that teaches the language its
  sentences.

It has no dependencies, opens no network connection, executes nothing, and
writes no file but the one you name. It is the same engine as the add-in,
held to the same corpus of tests, so a sentence means the same thing at the
prompt as on the Frazaro tab, and is refused in the same words.

## Installing

```text
cargo install frazaro
```

needs a Rust toolchain of 1.75 or newer and downloads nothing else, since
the crate has no dependencies. From a clone of
[the repository](https://github.com/Spreadsheet-Company/Frazaro),
`cargo build --release --workspace` puts the same program at
`target/release/frazaro`. The repository's test suite builds and runs it on
Linux on every push; it is developed on Windows. The commands on this page
are the same in PowerShell and in a Unix shell.

Nothing else is needed. The two files the language is made of, the
**prelude** (its standard library) and the **phrasebook** `english.vla` (its
English sentences), are inside the executable, byte for byte the
repository's
[scripts/prelude.vla](https://raw.githubusercontent.com/Spreadsheet-Company/Frazaro/main/scripts/prelude.vla)
and
[scripts/polyglotta/english.vla](https://raw.githubusercontent.com/Spreadsheet-Company/Frazaro/main/scripts/polyglotta/english.vla),
and a command uses them unless `--prelude` or `--phrasebook` names other
files.

## Five minutes

Save these seven sentences as `hello.txt`:

```text
Put 5 into cell B2.
Put formula "=B2*2" into cell B3.
Put formula "=B2+B3" into rows 2 to 4 of column C.
Put formula "=IFS(B2>3,""big"",TRUE,""small"")" into cell D2.

Work on sheet Data.
Put "west" into cell A1.
Put 7 into cell B1 of sheet Output.
```

Each step says what it prints, so the walk doubles as a check that the
program you installed is the one this page describes.

1. **The version.** `frazaro version` prints `frazaro 0.8.0 (core abi 1)`:
   the release's version, which is the add-in's, and the number an embedding
   checks before calling the core.

2. **What the sentences mean.**

   ```text
   frazaro translate-vla hello.txt
   ```

   prints the program in VLA: `(dim vla-step Long)` and `(sub main ()` first,
   then each sentence under its line, `(at-line 1 (set! (range "b2") 5))` for
   `Put 5 into cell B2.`, the text the add-in writes, character for
   character. `frazaro translate-vba` with the same arguments prints `Option
   Explicit`, `Dim vla_step As Long` and `Public Sub main()` onward: the VBA
   the add-in would write into a module.

3. **A workbook.**

   ```text
   frazaro build hello.txt --out hello.xlsx
   ```

   prints `wrote hello.xlsx: 10918 bytes, sha256
   03DCEC8E04577A84AA817F41EB6EEEAA141A3AEA0E29232385EE90A0A04F8918` and
   nothing else. The digest is the file's: whoever builds these seven
   sentences with this version gets these bytes, on any machine, whether the
   file was saved with Windows or Unix line endings. Open it: a `Frazaro`
   sheet holds the sentences in column B with OK beside each, `Output` holds
   the values and formulas, `data` its one cell, and nothing asks to be
   enabled.

4. **The workbook proves its build.**

   ```text
   frazaro rebuild hello.xlsx
   ```

   prints `This workbook was built from these 7 sentences by Frazaro 0.8.0:
   yes.` and exits 0. Save the file from Excel and ask again: it is refused
   as a host's save, because Excel rewrites every part on save, by design.

5. **One file is written, and guarded.** Run step 3 again: `Refusing to
   write hello.xlsx: a file is already there, and it has not been changed.
   Add --replace to replace it, or name another file with --out.`, exit 1.

6. **The workbook read back as relations.** `frazaro reflect hello.xlsx`
   prints what the file holds, one form a line, in an order that never
   changes: `(sheet "Frazaro" visible)` and its two sister sheets, the
   `Frazaro.Build` name with its fingerprints, the sentences as `(cell
   "Frazaro" "B1" "Put 5 into cell B2.")` rows with an `"OK"` cell beside
   each, then `Output`'s two values and five formulas, each followed by what
   it refers to, `(formula "Output" "C3" "=B3+B4")` then `(refers "Output!C3"
   "Output!B3")` and `(refers "Output!C3" "Output!B4")`, since a filled
   formula shows in each cell as Excel shows it, and `(cell "data" "A1"
   "west")` last: 34 lines. No formula has a `cell` row, because nothing has
   computed one yet; save the file from Excel and ask again, and each gains
   one. `frazaro reflect hello.xlsx --counts` prints counts and times alone,
   nothing from inside the file, and `frazaro reflect hello.xlsx --cone
   Output!C2` one line of counts for C2's cone: `cells 3 formulas 2 inputs
   1`, C2 itself, B3 that it adds and B2 that both read.

7. **Where the risks are.** `frazaro audit hello.xlsx` prints three lines:
   `(empty-reference "Output!C3" "Output!B4")`, then C4's to B4 and to B5.
   The third sentence filled `=B2+B3` down rows 2 to 4 of column C, past the
   two values in column B, so two of the three formulas add cells nothing
   holds. Nothing else: no constant typed over a formula, no inconsistent
   formula, no unused name, no hidden sheet, no link. `frazaro audit
   hello.xlsx --counts` prints the six counts alone.

8. **A sentence the language refuses.** Add `Set total to $5.` as a ninth
   line of `hello.txt` and translate it again: `I don't understand the
   character '$' - write the plain number (or 'Format ... as currency.' for
   display) (line 9)` on stderr, exit 1, the add-in's own words with the
   line.

9. **A sentence the writer cannot hold.** Make that line `Say "hello".`
   instead and build: `frazaro build writes only what a sheet can hold with
   nothing running: a value or a formula into a cell, a range or rows of a
   column, and which sheet it goes on. Line 9 asks for more: Say "hello".
   Nothing was written; run the program with the add-in, or take the
   sentence out.`, exit 1. The same program through `translate-vla` is fine:
   it is good English that needs the add-in's Run.

10. **A phrasebook of your own.** Save this as `own.vla`, one rule and its
    proof:

    ```text
    (english-vla
        "wobble {x:expr}"
        (debug-print {x}))
    (test-success
        "Wobble 5."
        (debug-print 5))
    ```

    `frazaro prove own.vla` prints `PASS 1/1`, and `frazaro load own.vla`
    prints `loaded: 1 rule, 0 macros, 1 test (0 expected fails) from
    own.vla`, the line the add-in's Load Phrasebook shows. Make the ninth
    line of `hello.txt` `Wobble 5.` and translate it with English and your
    phrasebook named in that order:

    ```text
    frazaro translate-vla hello.txt --phrasebook english.vla --phrasebook own.vla
    ```

    prints `(at-line 9 (debug-print 5))` under the other eight; `english.vla`
    here is the phrasebook inside the executable, since no file of that name
    is beside you. Without `--phrasebook own.vla` the ninth line is refused,
    `Don't understand: 'wobble 5.' No loaded sentence starts with 'wobble'
    ...`, and with `own.vla` alone the first line is, since `--phrasebook`
    names the whole list and English is no longer in it.

Two more steps need a clone of the repository, whose `scripts/build/` holds
these seven sentences under two comment lines as `fixture.txt`, built as
`fixture_golden.xlsx`, the golden the test suite holds the door to byte for
byte:

11. **Into a workbook of your own.**

    ```text
    frazaro build scripts/build/into.txt --into scripts/build/model.xlsx --out tour-into.xlsx
    ```

    prints `wrote tour-into.xlsx: 13697 bytes, sha256
    4281250F7AF88DDE3AA6495E3EAC47A11A3554CD7881797A2D5EDC00710DF798`: the
    model's own sheets `Model` and `Notes` copied byte for byte, the
    sentences' `Frazaro`, `Output` and `checks` added, `Output!B1` reading
    `=Model!B3*Rate`. `frazaro rebuild tour-into.xlsx` prints `This
    workbook's Frazaro sheets were built from these 6 sentences by Frazaro
    0.8.0, into a workbook whose own sheets are not checked: yes.`

12. **What changed between two workbooks.**

    ```text
    frazaro diff scripts/build/fixture_golden.xlsx scripts/reflect/saved.xlsx
    ```

    compares the build golden with the copy Excel saved of it: six lines,
    `(changed "Frazaro.Build" ...)` for the stamp Excel rewrote as
    `_xlfn._LONGTEXT`, then `Output`'s five formulas, each `(formula
    "=B2+B3")` on the left and `(formula "=B2+B3" 15)` on the right, the
    value Excel computed and the file now holds; nothing else differs, which
    is what saving a file should do. Two files that hold the same sheets,
    names, Tables and cells print nothing.

## The commands

| Command | What it prints |
|---|---|
| `frazaro translate-vla <program.txt> [--prelude <prelude.vla>] [--phrasebook <file.vla> ...]` | the program in VLA |
| `frazaro translate-vba <program.txt> [--prelude <prelude.vla>] [--phrasebook <file.vla> ...]` | the VBA the add-in would write into a module |
| `frazaro build <program.txt> [--prelude ...] [--phrasebook ...] --out <file.xlsx> [--into <model.xlsx>] [--replace]` | one line, the workbook's size and digest; the file is written |
| `frazaro rebuild <file.xlsx> [--prelude ...] [--phrasebook ...]` | one line: yes, this is that build, or no with why |
| `frazaro reflect <file.xlsx|.ods> [--counts] [--cone <Sheet!A1> ...]` | the workbook as relations, or counts alone |
| `frazaro diff <old.xlsx|.ods> <new.xlsx|.ods> [--counts]` | what differs between two workbooks |
| `frazaro audit <file.xlsx|.ods> [--counts]` | where the risks are, one finding a line |
| `frazaro view <program.txt> --sheet <name> [--window <A1:F20>] [--prelude ...] [--phrasebook ...]` | one window of the program's build as the view record, the lines a viewport draws from |
| `frazaro load <phrasebook.vla> [--prelude <prelude.vla>] [--allow-raw]` | what the phrasebook holds, its proofs run |
| `frazaro prove <phrasebook.vla> [--prelude <prelude.vla>] [--allow-raw]` | every failing proof, then `PASS n/n` or `FAIL k/n` |
| `frazaro compile <program.vla> [--prelude <prelude.vla>]` | a program already in VLA, compiled to VBA |
| `frazaro version`, `frazaro help` | the version; the usage text |

Every command that translates takes `--allow-raw`, explained below.

### What every command shares

- **Files in, text out.** The door reads the files it is named and prints
  the answer to stdout. The core sees text, or a workbook's bytes, and gives
  text back; the one file the door ever writes is `build`'s `--out`.
- **Built in.** The prelude and `english.vla` are inside the executable, so
  a command given neither flag below needs no file beside the program, and
  `frazaro build hello.txt --out hello.xlsx` is a complete command.
- **`--prelude <prelude.vla>`** names another standard library. `prove`
  looks beside the phrasebook and in its parent folder first, then uses the
  built-in one.
- **`--phrasebook <file.vla>`**, repeatable, names the whole list, in order:
  `english.vla` first, then an edition's dialect, then an organization's
  own, as the add-in loads them. `--phrasebook english.vla` with no such
  file beside you is the one inside the executable, so your own phrasebook
  over English is `--phrasebook english.vla --phrasebook own.vla` with one
  file. Each loads with its proofs run, and a failing proof refuses the
  load, as Load Phrasebook refuses it on the tab.
- **`--allow-raw`**, the consent a phrasebook with a `(raw ...)` form needs:
  the same checkbox the add-in and the web page ask. Without it such a
  phrasebook is refused by name. A phrasebook that requires a capability is
  refused always, since none can be granted yet.
- **Exit codes.** 0, the answer is on stdout. 1, a refusal: its words, the
  add-in's own, went to stderr, and nothing was written. 2, a usage error or
  a file that could not be read (the usage line, or `frazaro: cannot read
  ...`). 3, something this version declines to attempt (today only `prove`
  on an engine proof file).
- **Files** are UTF-8; a byte-order mark is dropped; line endings are left
  as they are, since the core and the add-in both leave them alone.

### `frazaro translate-vla` and `frazaro translate-vba`

```text
frazaro translate-vla <program.txt> [--prelude <prelude.vla>] [--phrasebook <file.vla> ...] [--allow-raw]
frazaro translate-vba <program.txt> [--prelude <prelude.vla>] [--phrasebook <file.vla> ...] [--allow-raw]
```

The program is a text file of sentences, one per line, a comment line
starting with `#`, a block indented by two spaces, exactly what the Frazaro
sheet's column B or a `<Frazaro>` section holds. `translate-vla` prints the
program's VLA, the text the add-in's own translator writes; in the
repository, `scripts/instructions.txt` translates to
`scripts/instructions_golden.vla` byte for byte, which the test suite holds.
`translate-vba` compiles that VLA with the prelude and prints the VBA the
add-in would write into a module, held to the `.vba` golden the same way.

A refusal is the sentence's own, with its line, on stderr, exit 1. Files it
cannot find are refused through the message catalogue: `Program file not
found: ...` (`english-program-file-not-found`), the prelude
(`vla-file-not-found`) and a phrasebook (`english-vocab-file-not-found`).

### `frazaro build`

```text
frazaro build <program.txt> [--prelude <prelude.vla>] [--phrasebook <file.vla> ...] --out <file.xlsx> [--into <model.xlsx>] [--replace] [--allow-raw]
```

The program is translated exactly as `translate-vla` translates it, then
written as a workbook at `--out`, the one file the door writes. The
`Frazaro` sheet holds the sentences in column B, exactly as written, with OK
beside each non-blank line in column C, in the add-in's own layout and
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
`scripts/build/fixture_golden.xlsx` a golden and the web page's download the
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
add-in's Run, quoted with its line (`build-not-representable`); a formula
that would reach outside the workbook on its own, `WEBSERVICE`, `HYPERLINK`,
a DDE link and their kin, named with its line and its sentence
(`build-formula-egress`, the same screen the add-in's two Run buttons
apply); a sheet a
sentence names that the program never made (`build-sheet-unknown`), or a
sheet name Excel would not take (`build-sheet-name-invalid`); a workbook too
large for the container (`build-workbook-too-large`); with `--into`, a file
that is not a workbook, an older `.xls` or an encrypted one, zip64, a broken
directory (`build-into-not-a-workbook`), a model that already has a
`Frazaro` sheet (`build-into-sheet-name-taken`), a sentence writing into one
of the model's own sheets (`build-into-model-sheet`), and a model that
cannot take a dynamic-array formula (`build-into-unsupported`). The
translation's own refusals come first, as in `translate-vla`.

### `frazaro rebuild`

```text
frazaro rebuild <file.xlsx> [--prelude <prelude.vla>] [--phrasebook <file.vla> ...] [--allow-raw]
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
gap: a workbook anyone has saved is no longer the build.

### `frazaro reflect`

```text
frazaro reflect <file.xlsx|.ods> [--counts] [--cone <Sheet!A1> ...]
```

The workbook's file is read, never Excel: its sheets, names and Tables, then
every cell, sheet by sheet in tab order and cell by cell in row order,
printed as relations in the notation of the language's proof corpus, one
form a line, in an order that never changes, so that the output is a golden
and two readings of one file compare line for line:

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
hold); a formula filled down shows in each cell with its references moved,
as Excel shows it; an array formula shows in its first cell, and the cells
it spills into are values. After a formula come its `refers` rows, one per
distinct thing it refers to, sorted: a cell or a range qualified by its
sheet with its `$` marks dropped (`Model!B1`, `Data!A:A`), a name, a
structured reference, a 3D span or a reference into another workbook as
written (`Rate`, `Sales[Amount]`, `Model:Scratch!B1`,
`[Rates.xlsx]Sheet1!A1`), a spill as its cell with `#`, a `#REF!` as
`Model!#REF!`, and `(unreadable "INDIRECT")` or `(unreadable "OFFSET")` for
the two calls whose targets no reader can know from the text; the reading
inside such a call goes on, so `=OFFSET(B1,1,0)` refers to `Model!B1` and
is unreadable both. The reading of references is the add-in's own reader,
ported arm by arm and held to its golden. A formula cell has a `cell` row
only when the file holds its last computed value, so a workbook `frazaro
build` wrote and nothing has opened has `formula` rows alone for its
formulas, and the same workbook saved from Excel gains a `cell` row for
each. A formatted cell with nothing in it has no row. Two files are compared
by `diff` and the audit questions asked by `audit`, below, both over this
reading.

An OpenDocument spreadsheet (`.ods`) reads the same: its sheets in document
order, a sheet hidden by its table style, its named expressions (a
sheet-local one as `Sheet!Name`), its database ranges as `table` rows
(Calc's named rectangle with a header row, the nearest kin of a Table) and
every cell as the file holds it, a date or a time value under `date`, a
time as its ISO duration, an error where Calc marks one or where a formula
cell's text is an Excel error literal. A formula's OpenFormula text prints
as the formula bar shows it in Excel and in Calc alike:
`of:=SUM([.B1:.B2];[Data.A1])` is `=SUM(B1:B2,Data!A1)`, a sheet quoted by
the one rule, two sheets as a 3D span, a link's file in brackets, and
`COM.MICROSOFT.` dropped from a function's name, so the `refers` rows read
the same from either file and `frazaro diff model.xlsx model.ods` says
exactly what one format holds that the other cannot. The fixture's own
twin, `scripts\reflect\opendocument.ods`, differs from
`scripts\reflect\fixture.xlsx` in six rows: a very-hidden sheet that ODF
can only hide, a `#REF!` name it cannot spell, a Table column summed as a
plain range, a date held as a value where the package holds a serial, and
a sheet-local name held workbook-level, since Excel's reading of `.ods`
keeps no sheet-local name. Excel's own `.ods` writing reads too, its
`$$Name`, its `COM.MICROSOFT.SINGLE` as `@`, its error cells and its empty
rows repeated a million times; the twin as Excel saved it,
`scripts\reflect\opendocument_saved.ods`, is a fixture of all three
oracles, and its diff against the twin is what Excel's writing changes,
nine rows.

With `--counts`, nothing from inside the workbook is printed: one line a
sheet, by position, with its counts of cells, formulas, distinct formulas
(compared in R1C1 relative to their cells, so that a formula filled down
counts once), unreadable calls (`INDIRECT` and `OFFSET`) and array formulas,
the last row and column holding anything, the part's size and the time it
took to read, then one line for the workbook with its totals, names, Tables,
links to other workbooks and strings. This is for measuring a model whose
contents must not leave the machine; the lines can be pasted anywhere.

With `--cone Sheet!A1`, once or more, a cell's cone is sized through the
file, one line a root with counts alone: the cells reached through every
reference, through names and Tables and across sheets, how many of them hold
formulas and how many are inputs, the blank cells referred to, the sheets
touched, the depth of the longest chain, where the cone is blind (`INDIRECT`
and `OFFSET`), the names and Tables resolved, the references into other
workbooks (counted, not followed), the `#REF!`s met and what the file does
not have. Excel's own Trace Precedents stops at the sheet boundary; this
does not. The root is the one address printed, because you typed it. A root
that is not a cell as `Sheet!A1` or `'Q1 Data'!A1`, or one naming no sheet,
exits 2.

Refusals, exit 1: a file that is not a workbook, with why
(`reflect-not-a-workbook`: not a zip, an older `.xls` or an encrypted one,
zip64, no workbook part, a damaged part, or an OpenDocument file that is
not a spreadsheet); a part this version refuses to read
(`reflect-xml-refused`: one declaring a `DOCTYPE` or an entity, which the
reader never expands, or markup that never closes); a shape this version
does not read, named with its part and cell (`reflect-unsupported`: a
data-table formula, a shared formula whose first cell is missing, a string
index past the table). The reader streams, so the rows printed before a
refusal stand on stdout.

### `frazaro diff`

```text
frazaro diff <old.xlsx|.ods> <new.xlsx|.ods> [--counts]
```

Two workbooks' files are read as `reflect` reads them and compared, and
what differs is printed one form a line, in an order that never changes, so
that the output is a golden too:

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
`sheet-changed` row. Every other difference is a `changed` row with the old
and the new side by side: first the names and Tables, by name (a name's
sides are what it refers to, a Table's its sheet and range, `blank` where
that file has none), then sheet by sheet in the new file's tab order and
cell by cell in row order, the cell spelled as a reference spells it, each
side a value as `reflect` prints one, `(formula "...")` for a formula whose
cached value the file does not hold, `(formula "..." value)` for one it
does, or `blank`. So a formula typed over by a constant reads `(changed
"Model!B3" (formula "=B1-B2" 400) 500)`, and a formula left alone whose
result moved reads as two values under one text. A cell counts as changed
when its formula's text or its value differs; numbers are compared as
numbers, so `800` and `800.0` are one value, and everything else by kind and
text, so a text `"5"` typed over the number `5` is a change. Two files that
hold the same sheets, names, Tables and cells print nothing. The exit code
is 0 whenever the comparison ran, rows or none: the rows are the answer.

With `--counts`, one line and nothing from inside either file: `diff:
sheets-removed 1 sheets-added 1 sheets-changed 1 names 3 tables 1 cells 17
compared 42 open 0.9 ms compared 1.8 ms`, `compared` being the addresses
either file holds on the matched sheets.

Refusals, exit 1, are the reader's (under `reflect`, above), raised for
whichever file raised them; the rows printed before one stand. Two files
are read and nothing is written.

### `frazaro audit`

```text
frazaro audit <file.xlsx|.ods> [--counts]
```

The audit list: where a workbook's risks are, read from the file. Six
questions, each a walk over what `reflect` reads, one finding a line, in an
order that never changes:

```text
(typed-over "Review!B3" 61 "=A3*2")
(inconsistent "Review!C3" "=A3+2" "=A3+1")
(unused-name "Range1" "Data!$A$2:$B$4")
(empty-reference "Review!D1" "Review!Z9")
(hidden-sheet "Scratch" hidden)
(external-link "Model!E2" "[Rates.xlsx]Sheet1!A1")
```

A **typed-over constant** is a value whose nearest neighbours above and
below in its column both hold formulas that agree in R1C1; the row shows the
value and the formula those neighbours would have put there. An
**inconsistent formula** differs in R1C1 from two such neighbours that
agree; the row shows the formula as written and the neighbours'. A cell at
the top or bottom of a column, or between neighbours that disagree, is never
reported, so a finding is one worth acting on. An **unused name** is a
defined name no formula and no other name refers to (Excel's own `_xlnm.`
names, print areas and filter databases, and Frazaro's own marks, the build
stamp and the add-in's `VLAt_` names, are never reported). An **empty
reference** is a formula's reference to one cell, on a sheet the file has,
that holds nothing: a formatted blank counts, a range is never expanded. A
**hidden sheet** is what its `sheet` row says, and an **external link** is
a formula that reaches another workbook. The order is those six, and inside
each sheet by sheet in tab order and cell by cell in row order, names by
name. A workbook with nothing to report prints nothing, and the exit code is
0 either way: the findings are the answer. A seventh question, totals that
do not foot, is not asked yet.

With `--counts`, one line and nothing from inside the file: `audit:
typed-over 1 inconsistent 1 unused-names 2 empty-references 2 hidden-sheets
2 external-links 1 open 0.9 ms indexed 1.4 ms walked 0.3 ms`, the number
internal audit wants of a model it may not show.

Refusals, exit 1, are the reader's (under `reflect`, above); the whole file
is read before the first finding, so nothing is printed before one.

### `frazaro view`

```text
frazaro view <program.txt> --sheet <name> [--window <A1:F20>] [--prelude <prelude.vla>] [--phrasebook <file.vla> ...] [--allow-raw]
```

The program is translated and built exactly as `build` builds it, into the
sheet model in memory, and nothing is written; one window of one sheet of
that model is printed as the view record, the lines a viewport draws from,
one form a line in the notation of the language's proof corpus, in an
order that never changes, so that the output is a golden. For the seven
sentences of `hello.txt` above:

```text
frazaro view hello.txt --sheet Output --window A1:F20
```

prints

```text
(sheet "Frazaro" visible)
(sheet "Output" visible)
(sheet "data" visible)
(window "Output" "A1:F20")
(extent "Output" "B1:D4")
(gridlines "Output" on)
(format 0 none general nowrap)
(cell "Output" "B1" 7)
(sentence "Output" "B1" 8)
(cell "Output" "B2" 5)
(sentence "Output" "B2" 1)
(formula "Output" "C2" "=B2+B3")
(sentence "Output" "C2" 3)
(formula "Output" "D2" "=IFS(B2>3,\"big\",TRUE,\"small\")")
(sentence "Output" "D2" 4)
(formula "Output" "B3" "=B2*2")
(sentence "Output" "B3" 2)
(formula "Output" "C3" "=B3+B4")
(sentence "Output" "C3" 3)
(formula "Output" "C4" "=B4+B5")
(sentence "Output" "C4" 3)
```

First every sheet of the model, in tab order; then the window as asked, its
sheet spelled as the model spells it; the sheet's `extent`, the rectangle
that holds its cells, or `none` for a sheet with nothing in it, so that a
viewport knows how far there is to scroll; whether the sheet shows
gridlines; a `column` row for each column of the window that has settings
(its width in characters or `none`, `shown` or `hidden`, its format or
`none`), which is how a window over the `Frazaro` sheet learns that its
column A is hidden and its columns B and C are wide and coloured; a
`format` row for format 0 and every format the window uses (the fill as
six hex digits or `none`, `general` or `text`, `wrap` or `nowrap`); then
the cells in row order, each as `reflect` prints the same cell of the
written file, a `cell` row for a value or a `formula` row for a formula's
text, followed by its `style` row when its format is not 0 and its
`sentence` row, the row of the sentence that wrote it, which is the
program's line and the row of the `Frazaro` sheet that holds the sentence.
So `B1` holds 7 because of line 8, `Put 7 into cell B1 of sheet Output.`,
and `C2`, `C3` and `C4` hold the one filled formula of line 3, each with
its references moved as Excel shows them. A `cell` row is a value the
sentences put there; a `formula` row is text the host computes when the
file opens, since the core computes nothing, and the record never shows
one as the other.

Without `--window` the sheet's whole extent is shown: `frazaro view
hello.txt --sheet Frazaro` prints the room, both columns with their widths
and formats and every sentence with its OK mark, each by its own row. A
sheet's name is matched as Excel matches one, without case, and printed as
the model spells it. The record is drawn from the model, never from a
file, so `view` runs wherever `build` would and needs no `--out`; the same
record comes out of the engine's C surface as `frazaro_view`, for a
viewport to draw from.

Refusals, exit 1: a sheet the program does not make, naming the ones it
does (`view-sheet-unknown`); a window that is not a rectangle of cells, a
whole column or row included (`view-window-not-a-range`); and before
those, the translation's and the build's refusals, as in `build`. A
missing `--sheet` is a usage error, exit 2.

### `frazaro load`

```text
frazaro load <phrasebook.vla> [--prelude <prelude.vla>] [--allow-raw]
```

Loads a phrasebook through the core's loader as Load Phrasebook loads it,
with the door's two gates first: a required capability refuses always, and a
`(raw ...)` form refuses unless `--allow-raw` gives the consent. Every proof
in the phrasebook runs at load, and a failing one refuses the load with its
words. On success the line the add-in shows: `loaded: 240 rules, 220
macros, 460 tests (22 expected fails) from <file>`, followed by the
grammar's lint warnings when there are any. Exit 0, or 1 on a refusal.

### `frazaro prove`

```text
frazaro prove <phrasebook.vla> [--prelude <prelude.vla>] [--allow-raw]
```

Every proof in the phrasebook is run as loading it in Excel runs it, every
failure is printed as the add-in would have refused it, and the last line is
`PASS n/n` (exit 0) or `FAIL k/n`, k passed of n (exit 1). The prelude is
`prelude.vla` beside the file or in its parent folder, which is how
`scripts/polyglotta/english.vla` finds `scripts/prelude.vla` in the
repository, else the one inside the executable; `--prelude` names another.
A load that refuses before its
proofs (a malformed rule, a library file with no rule and no proof) prints
the refusal and exits 1 with no verdict line. An engine proof file
(`scripts/proofs/datalog.vla` in the repository) is not attempted in this
version: `frazaro: 'prove' attempts a phrasebook's proofs; engine proofs are
not attempted in this version (PORT.9)`, exit 3, which the repository's
conformance runner reads as *not attempted*, never as passed. Today
`english.vla` is `PASS 482/482` and each dialect passes its own.

### `frazaro compile`

```text
frazaro compile <program.vla> [--prelude <prelude.vla>]
```

A program already written in VLA, compiled with the prelude: its VBA to
stdout, the text the add-in would write into a module. In the repository,
`scripts/instructions_golden.vla` less its stamp line compiles to
`scripts/instructions_golden.vba` whole. A missing program is
`vla-source-not-found`, a missing prelude `vla-file-not-found`; a compile
refusal is the compiler's own, exit 1.

### `frazaro version`, `frazaro help`

`version` (also `--version`, `-V`) prints `frazaro <version> (core abi N)`.
`help` (also no arguments, `--help`, `-h`) prints the usage text. A word that
is not a command exits 2 with `frazaro: '<word>' is not a command in this
version; 'frazaro help' lists the commands.`

## Not in this version

The usage text lists `check`, `run` and `ask` under *Not in this version*,
so that a reader who guesses a command learns where it is. Each is an item
of the repository's roadmap (`docs/BETA_ROADMAP.md`), and this page gains
its section when the item closes:

- **`frazaro ask <model.xlsx> "<question>"`**: a question in English over a
  workbook, answered from the file by the engines the add-in runs, the audit
  list's questions first (AXM.9, with PORT.9).
- **The engines in the core**: SQL, DATALOG, PROLOG and OPTIMIZE over tables,
  and `frazaro prove` attempting an engine proof file instead of exiting 3
  (PORT.9).
- **`frazaro run <program.txt> --into <model.xlsx>`**: a program carried out
  in memory against a workbook's cells, the interpreter over the sheet
  model, with a sentence that needs Excel itself refused by name (PORT.10).
- **`frazaro check`** over a repository of sentence files, the GitHub Action
  among the doors (PORT.11).

## If you use the add-in

| On the Frazaro tab | At the prompt |
|---|---|
| Validate Instructions (a refusal names its line) | `frazaro translate-vla` |
| Translate File to VLA, Translate File to VBA | `frazaro translate-vla`, `frazaro translate-vba` |
| Show me the VBA | `frazaro translate-vba` |
| Compile and Run, Interpret and Run | nothing runs here; `frazaro build` writes what a sheet holds without running |
| Load Phrasebook, and the consent it asks for a `(raw ...)` form | `frazaro load`, `--allow-raw` |
| the proofs Load Phrasebook runs | `frazaro prove` |
| Open CLI, one sentence at a time | the shell; the door takes files, so a one-line file is the sentence |
| nothing yet: the workbook as relations, the difference and the audit over the live workbook are roadmap items | `frazaro reflect`, `diff` and `audit`, from the file |
| `? VlaSelfTests()` in the Immediate pane | `cargo test --workspace`, then `powershell -File tools/prove.ps1 -Impl target/debug/frazaro` |

The web page in the repository's `web/` is the same core compiled to
WebAssembly behind a second door: its rows are the program, its *Download as
.xlsx* is `frazaro build` with the digest shown beside the button, and
`frazaro rebuild` says yes to what it downloads.

## For contributors

This page is the reference for the door, and the `USAGE` text in
`cli/src/main.rs` is its short form (`frazaro help`); when a command
changes, both change.

The add-in in the repository's `src/` is the reference implementation of
the language, and the goldens it wrote under `scripts/` are what this door
is held to, by the treaty in `conformance/README.md`:
`tools/prove.ps1 -Impl target/debug/frazaro` scores the door against it,
the corpus program translated and compiled, every phrasebook proved, both
build goldens built and `rebuild` run on each, the reflect fixtures read,
the diff pairs compared and the audit rows walked, the OpenDocument twin of
the fixture and Excel's save of it among the rows of all three; `31 passed,
0 failed, 2 not attempted, 1 library` today, the two not attempted being the
interpreter golden (PORT.10) and the engine proofs (PORT.9). `-Control`
first proves the runner on a fake door and a mutant. The checks under
`tools/` hold floors on it (`check_compile_prefix.ps1`,
`check_translate_prefix.ps1`, `check_prove_floors.ps1`,
`check_build_golden.ps1`, `check_reflect_golden.ps1`,
`check_diff_golden.ps1` and `check_audit_golden.ps1`, each taking
`-Impl`), and the `core` job of `.github/workflows/checks.yml` builds the
workspace, formats, lints and tests it, builds the wasm, and runs every one
of them on every push. Before a change to the door is committed: `cargo fmt
--all -- --check`, `cargo clippy --workspace --all-targets -- -D warnings`,
`cargo test --workspace`, and `powershell -File tools/run_checks.ps1`.

## Licence

Apache-2.0 AND MPL-2.0: the door's own code is Apache-2.0, and the
`english.vla` it carries is MPL-2.0, per file, so a phrasebook you write is
yours. The repository's `REUSE.toml` maps every file to its licence.
