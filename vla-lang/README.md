# vla-lang

**The VLA language with nothing around it: forms, their expansion, the
grid, the view record, and the machine that steps a grid.**

Frazaro is a language for spreadsheet work: you write the steps of a
procedure as English sentences, and Frazaro checks every sentence, refuses
the ones it cannot read, in words, and carries out the rest. VLA is the
Lisp beneath that English, the form every sentence translates into before
anything is emitted or built. This crate is that language alone: the reader
and the printer of its forms, the macro expander, the head table that names
the core forms, the refusal catalogue's mechanism with the language's own
refusals, the sheet model with its A1 coordinates and the references read
out of a formula's text, the view record, one window of a sheet as the
lines a viewport draws from, and recalculation's mechanism, the formulas
of a grid computed in dependency order with a budget in cells. No English,
no file format, no clock, and no dependency. Built for WebAssembly, the
module's import section is empty.

[`frazaro-core`](https://crates.io/crates/frazaro-core) stands on this crate
and adds the bridges: the English engine with a proof per sentence, the VBA
and formula emitters, the OOXML and OpenDocument writer and readers.
[`frazaro`](https://crates.io/crates/frazaro) is the command line over both.
An engine that runs a grid as a game stands on this crate and never on the
core, which is why the language is a crate of its own. The Excel add-in in
the repository is the reference implementation; this crate is held to the
same corpus of tests, so a form means the same thing here as there and is
refused in the same words.

## What it does

- **Reads and prints VLA.** `reader::read_forms` takes a text to its forms,
  each a symbol, a string, a number or a list with the line it came from;
  `printer::write_datum` and `printer::write_pretty` write a form back, as
  the add-in writes its goldens.
- **Expands macros.** `expand::expand_text` takes a program's text and a
  prelude's, defines the prelude's macros and expands every form to a
  fixpoint, refusing a malformed definition or a runaway expansion by name.
  The prelude is a library the caller passes: the repository's
  [scripts/prelude.vla](https://raw.githubusercontent.com/Spreadsheet-Company/Frazaro/main/scripts/prelude.vla)
  is the standard one, and a program may carry its own definitions.
  Twenty-seven primitives compute while a macro expands, over the data
  written in the source: the list operations (`car`, `cdr`, `cons`,
  `quote-if`, `cond` and their kin), arithmetic and comparison (`+expand`,
  `-expand`, `*expand`, `quotient-expand`, `mod-expand`, `=expand` and the
  rest), text joined or split into its characters (`&expand`,
  `characters-expand`), and a cell's address written from a row and a
  column or read back (`address-expand`, `row-expand`, `column-expand`,
  `column-letters-expand`). A name that ends in `expand` is computed while
  the macro expands; each refuses by name what it cannot compute.
- **Names the core forms.** `headtable::row` and `headtable::resolve_head_alias`
  answer for every form the language has, its aliases and its arity, from
  the table the add-in's own code exports.
- **Holds the grid.** `sheet::Workbook`, `sheet::Sheet`, `sheet::Cell` and
  `sheet::Content`: sheets of cells, each a text, a number, a truth value or
  a formula's text, with column widths, a style table and the row of the
  sentence that wrote each cell. `Sheet::set_formula` fills a range with one
  formula as Excel fills it, the references moved cell by cell.
- **Draws a number under a format's code.** `numfmt::Code::parse` takes a
  code of Excel's, such as `0.00` or `▲0.00;▼0.00;=0.00`, its sections,
  digit placeholders, thousands, thousandths, percents and literals, and
  names whatever else a code holds that this version does not draw;
  `Code::text` gives the text a value draws as under it, as Excel's
  `Range.Text` does: the number's fifteen significant digits, rounded half
  away from zero to the code's decimals; and a date's and a time's parts,
  the year, the month and the day by number or English name, the hour, the
  minute, the second and its decimals, an elapsed time and a twelve-hour
  clock, over Excel's 1900 calendar. A style's `sheet::NumFmt` is `General`,
  `Text` or a `Code`.
- **Says how a cell looks.** A style (`sheet::Style`) is a fill, a number
  format and a wrap, and what it says of its text's font, its alignment and
  its borders (`appearance`): a weight, a slant, a size in points from 1 to
  409.55 by twentieths, as Office keeps one, and a color; a horizontal and a
  vertical alignment; and each of a cell's four edges, a line of Excel's
  fourteen and a color, each said or not. The view record prints what a
  format says after its wrap, `(font bold 22 "FFFFFF")`,
  `(align center middle)` and `(border (left medium))`, and the loader reads
  it back. `Style::over` is the format a cell draws with under a look: the
  look's fill, number format and wrap, what it says of a font, an alignment
  and a border, and the cell's own format for the rest.
- **Reads references.** `refers::scan` finds every reference inside a
  formula's text, cells, ranges, columns, rows, names, Tables, external
  books and 3D spans, and names the calls it cannot read; `refers::r1c1`
  renders a formula relative to its cell, `refers::shift_a1_references` is
  the mover a fill uses, and `refers::quote_sheet` is the one sheet-quoting
  rule.
- **Draws the view record.** `view::window_of` finds a sheet and a window
  in a model, and `view::view_text` prints that window as lines: every
  sheet, the window and the sheet's extent, its columns and formats, then
  each cell's value or formula, the text a number draws under its format's
  code, its format and the row of the sentence that wrote it. It is the
  first implementation of `projection::Projection`, the
  seam a second projection enters by. `view::view_text_valued` prints the
  same record with each formula's computed value after it. The second,
  `view::ValueGrid`, is the same record less its formula rows, the values a
  host draws every frame without moving every formula's text:
  `view::view_text_values` prints it.
- **Computes the grid.** `calc::Calc::new` reads every formula off a
  `Workbook`, builds the dependency graph from the references and finds the
  order, naming every cycle; `Calc::step` evaluates the next cells of that
  order up to a budget in cells and says how far it got, so an engine can
  run a grid at a frame rate and yield; `Calc::run` evaluates everything.
  The parser reads Excel's operators and calls over the reference scanner,
  with Excel's precedence and coercions and its eight error values,
  `#CALC!` among them, as a value kind, array constants (`{1,2;3,4}`) and
  `:` between two expressions that answer references among its forms; a
  comparison holds two numbers equal to fifteen significant digits equal,
  as Excel's does (`=0.1+0.2=0.3` is TRUE), where a lookup matches a number
  exactly, and orders two texts as Excel's save ranks them, case aside, an
  accent after the plain letter, `ß` and `ss` one, a comparison by size
  that a character it did not rank decides labelled. A range where one
  value is wanted is met at the formula's cell, Excel's implicit
  intersection, in a legacy formula, and is read whole, as an array, in a
  dynamic-array formula, which `Env::dynamic` tells, in a legacy array
  formula, which `Env::array_formula` tells, and inside a parameter its
  function reads as an array in any formula (`calc::Param`). An operator,
  or a function of one value, given an array computes at each of its
  values, Excel's lift (`calc::Array`, its values shared by a count); a
  cell holds the first value of an answer of many in a legacy formula and
  a legacy array formula, and labels it a spill in a dynamic one until
  spills come. Functions enter through `calc::Library`, the seam: the
  language registers the fourteen a game's rule needs on day one (`IF`,
  `AND`, `OR`, `NOT`, `SUM`, `MIN`, `MAX`, `ABS`, `INT`, `MOD`, `ROW`,
  `COLUMN`, `CHOOSE`, `SIN`), the lookups (`INDEX`, `MATCH`, `XMATCH`,
  `VLOOKUP`, `HLOOKUP`, `LOOKUP`, `XLOOKUP`), the tests a lookup's miss is
  read by (`IFERROR`, `IFNA`, `ISERROR`, `ISNA`), the criteria (`COUNTIF`,
  `SUMIF`), `SINGLE`, Excel's `@`, the text functions (`calc::texts`:
  `LEFT`, `MID`, `LEN`, `FIND`, `SUBSTITUTE`, `TEXT` and the rest, a text
  counted by its characters as Excel counts it), the date functions
  (`calc::dates`: `DATE`, `YEAR`, `MONTH`, `WEEKDAY`, `NETWORKDAYS`,
  `YEARFRAC` and the rest, over Excel's 1900 calendar, a text a locale
  reads labelled in arithmetic), the aggregates by criteria: `SUMIFS`,
  `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS` and `MINIFS`, `SUBTOTAL`
  by each of its numbers, a hidden row (`Sheet::rows`) passed by at 101 to
  111, the database functions (`DSUM`, `DGET` and the rest) and the
  aggregates `SUBTOTAL` names (`AVERAGE`, `COUNT`, `COUNTA`, `PRODUCT`,
  `STDEV`, `VAR` and their newer names), and the array functions
  (`calc::arrays`: `SEQUENCE`, `FILTER`, `SORT`, `UNIQUE`, `SUMPRODUCT`,
  `MMULT`, the shapers `TAKE`, `HSTACK`, `TOCOL` and the rest, `#CALC!`
  for an array with nothing in it). A crate above registers more, as
  `frazaro-core` registers Excel's library by measurement. A lookup or a
  criterion asked of one range again and again searches the range's
  index (`calc::Indexes`), built at its second probe in one evaluation and
  kept until a cell of the range changes, its answers the ones a walk of the
  range gives (`calc::index`); the `SUMIFS` family's ranges are grouped by
  the keys of its equalities, every group answering the formulas that ask
  it. A formula that calls anything else is not computed, and says so naming
  the function; a cell that reads it inherits the label.
  `Calc::first_refusal` is what an engine's load asks: a cycle, a function
  outside the library, or a construct the parser does not read, a defined
  name's text among them, each refused by name from the catalogue.
- **Runs a grid as a machine.** `machine::Machine` is a grid kept between
  calls, for a game engine that steps it at a frame rate: `Machine::load`
  reads the rows a view or a reflect prints back into a grid, the record's
  inverse; `write` takes `cell`, `formula` and `derived` rows, all or none;
  `step` computes every formula in dependency order with a budget in cells,
  yields and resumes, and at the end of each frame copies every sheet's
  values into its twin, `X.last`, a hidden sheet the next frame's formulas
  read the previous frame from (`=Screen.last!B2+1`); `view` answers the
  record with each formula's value, the values (the record less its formula
  rows, what a host draws), the changes (the values' rows of the cells the
  last frame changed, so that a host repaints those alone, or `whole` where
  the machine cannot say), or the plane, a byte a cell. After each frame
  `changed` says how many formula cells it changed the value of, to the
  bit, `changed_on` how many of them on a sheet and `reached` how many it
  computed, so that a host stepping a grid until a cascade settles stops
  when nothing moves: a frame that changed nothing, with no write before
  it, is followed by frames that change nothing. Each
  formula is parsed once per shape (a filled formula, or cells whose R1C1
  text is the same) and evaluated at each cell's offset. `census` prints
  the plan as rows before a frame runs: each shape's cells, the levels of
  the order it spans, the runs it is computed in and the path each of its
  reads takes, so that a grid's cost is read ahead (`calc::Shapes::census`
  prints a plan a door made of a workbook the same way). `machine::Handles`
  keeps grids under handles from 1, never reused, sixteen at once. The C
  surface over the four calls, `vla_load`, `vla_write`, `vla_step`,
  `vla_view` and their companions, is compiled always and exported only
  under the `c-abi` feature, so that no module built on this crate exports
  a second door to its grid.
- **Refuses in words.** Every refusal is a `messages::Refusal`, a value
  carrying the catalogue's id, number, source and text, never a panic.

## Using it from Rust

```text
cargo add vla-lang
```

```rust
use vla_lang::{expand, printer, reader};

fn main() {
    let prelude = std::fs::read_to_string("prelude.vla").unwrap();
    let program = "(when (> 1 0) (debug-print \"hello\"))\n";

    // The program with its macros expanded, and how many applications it took.
    match expand::expand_text(program, &prelude, true) {
        Ok((expanded, fired)) => println!("{expanded}; {fired} applications"),
        Err(r) => eprintln!("{} ({})", r, r.id),
    }

    // The same text read as forms and printed again, one datum a line.
    for form in reader::read_forms(program).unwrap() {
        println!("{}", printer::write_datum(&form));
    }
}
```

## The contract

Bytes in, bytes out, nothing else. The crate has no dependencies, and
nothing in it reads a file, a clock, a socket or a random number at run
time. The two tables it needs, the head table and its half of the message
catalogue, are embedded at build time from `data/`, inside the crate,
exported from the add-in's own code and held to it by the repository's
checks. `vla_lang::VERSION` is the crate's version, which is the release's,
the add-in's and `frazaro-core`'s, one number for the language.

Every text the crate reads is answered or refused by name, never with a
panic, a stack overflow or a hang, so an engine can hand it a cartridge
from anyone. The limits are public constants, each refused by its own id:
- lists nested deeper than 64, in what the reader reads and in what the
  expander writes (`reader::FORM_DEPTH_LIMIT`);
- a formula nested deeper than Excel's 64 levels, or longer than its
  8,192 characters (`calc::formula::NESTING_LIMIT`, `LENGTH_LIMIT`);
- more than a million forms in what one macro application writes
  (`expand::EXPANSION_SIZE_LIMIT`), since the crate copies a macro's
  arguments into its template.

Real programs nest 10 deep at most and real formulas 6. Each stage is
tested at its limit on a 1 MiB stack, the tightest an engine is given.

```text
cargo build --release -p vla-lang --features c-abi --target wasm32-unknown-unknown
```

builds the crate as a WebAssembly module with the machine's C surface,
whose import section the repository's `tools/check_core_imports.ps1` holds
at zero entries, so that no engine built on it can reach a socket, a clock
or a file through the language, and whose exports
`tools/check_wasm_exports.ps1` holds to exactly its nine functions and its
memory.

## What comes next

Each is an item of the repository's roadmap (`docs/BETA_ROADMAP.md`), and
this page gains its section when the item closes:

- **Recalculation's next slices.** Dynamic arrays' spills, an item of their
  own after KERNEL.8, whose six slices, the lookups, implicit intersection
  and `@`, the text and the date functions, the aggregates by criteria and
  arrays within a formula, are above; the vectorized evaluation of a shared
  formula, incremental recomputation and content-addressed evaluation
  (KERNEL.20), which is what makes a frame of a 320 by 200 grid fast: today
  the machine steps one in about a tenth of a second natively. The first
  slice is here.
- **The corpus's expansion golden.** The corpus program expanded under the
  prelude, written by the reference and reproduced here, beside the golden
  of the expand-time primitives the crate is already held to.
- **An expander that shares what it walks.** A macro that walks a list
  copies the rest of the list at each step today, so its cost grows with
  the square of the list: a table generator's walk of 2,000 rows takes
  about nine seconds natively (VLA.2).

## Where it comes from

The add-in in the repository's `src/` is the reference implementation of
the language, and the goldens it wrote under `scripts/` are what this crate
is held to, by the treaty in
[conformance/README.md](https://github.com/Spreadsheet-Company/Frazaro/blob/main/conformance/README.md):
the corpus program read, expanded and compiled byte for byte through the
core, and the reference goldens the add-in writes reproduced by this
crate's own tests. The crate follows the goldens and never leads them. It
was cut out of `frazaro-core` on 2026-10-08 (PORT.12) so that an engine
can stand on the language without the bridges.

The goldens and fixtures the tests read travel with the package, under
`testdata/`, so `cargo test` runs from the published crate. The repository
holds the copies to their sources byte for byte. The expander is held to
the add-in's expansion of 135 cases, every expand-time primitive and its
refusals among them (`tests/expand_golden.rs`). Beside the goldens, the
tests hold the language's laws over generated programs (`tests/laws.rs`)
and knock at every door with mutated text (`tests/strangers.rs`, and a
longer run with `--ignored`).

## Licence

Apache-2.0. The repository's `REUSE.toml` maps every file to its licence.
