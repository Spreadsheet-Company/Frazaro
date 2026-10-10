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
- **Names the core forms.** `headtable::row` and `headtable::resolve_head_alias`
  answer for every form the language has, its aliases and its arity, from
  the table the add-in's own code exports.
- **Holds the grid.** `sheet::Workbook`, `sheet::Sheet`, `sheet::Cell` and
  `sheet::Content`: sheets of cells, each a text, a number, a truth value or
  a formula's text, with column widths, a style table and the row of the
  sentence that wrote each cell. `Sheet::set_formula` fills a range with one
  formula as Excel fills it, the references moved cell by cell.
- **Reads references.** `refers::scan` finds every reference inside a
  formula's text, cells, ranges, columns, rows, names, Tables, external
  books and 3D spans, and names the calls it cannot read; `refers::r1c1`
  renders a formula relative to its cell, `refers::shift_a1_references` is
  the mover a fill uses, and `refers::quote_sheet` is the one sheet-quoting
  rule.
- **Draws the view record.** `view::window_of` finds a sheet and a window
  in a model, and `view::view_text` prints that window as lines: every
  sheet, the window and the sheet's extent, its columns and formats, then
  each cell's value or formula, its format and the row of the sentence that
  wrote it. It is the first implementation of `projection::Projection`, the
  seam a second projection enters by. `view::view_text_valued` prints the
  same record with each formula's computed value after it.
- **Computes the grid.** `calc::Calc::new` reads every formula off a
  `Workbook`, builds the dependency graph from the references and finds the
  order, naming every cycle; `Calc::step` evaluates the next cells of that
  order up to a budget in cells and says how far it got, so an engine can
  run a grid at a frame rate and yield; `Calc::run` evaluates everything.
  The parser reads Excel's operators and calls over the reference scanner,
  with Excel's precedence and coercions and its seven error values as a
  value kind. Functions enter through `calc::Library`, the seam: the
  language registers the fourteen a game's rule needs on day one (`IF`,
  `AND`, `OR`, `NOT`, `SUM`, `MIN`, `MAX`, `ABS`, `INT`, `MOD`, `ROW`,
  `COLUMN`, `CHOOSE`, `SIN`), and a crate above registers more, as
  `frazaro-core` registers Excel's library by measurement. A formula that
  calls anything else is not computed, and says so naming the function; a
  cell that reads it inherits the label. `Calc::first_refusal` is what an
  engine's load asks: a cycle, a function outside the library, or a
  construct the parser does not read, each refused by name from the
  catalogue.
- **Runs a grid as a machine.** `machine::Machine` is a grid kept between
  calls, for a game engine that steps it at a frame rate: `Machine::load`
  reads the rows a view or a reflect prints back into a grid, the record's
  inverse; `write` takes `cell`, `formula` and `derived` rows, all or none;
  `step` computes every formula in dependency order with a budget in cells,
  yields and resumes, and at the end of each frame copies every sheet's
  values into its twin, `X.last`, a hidden sheet the next frame's formulas
  read the previous frame from (`=Screen.last!B2+1`); `view` answers the
  record with each formula's value, or the plane, a byte a cell. Each
  formula is parsed once per shape (a filled formula, or cells whose R1C1
  text is the same) and evaluated at each cell's offset. `machine::Handles`
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

- **Recalculation's next slices.** Dates, the text functions, the lookups,
  dynamic arrays and spills, implicit intersection and `@` (KERNEL.8); the
  vectorized evaluation of a shared formula, incremental recomputation and
  content-addressed evaluation (KERNEL.20), which is what makes a frame of
  a 320 by 200 grid fast: today the machine steps one in about a tenth of a
  second natively. The first slice is here.
- **An expansion golden of its own.** The corpus program expanded under
  the prelude, written by the reference and reproduced here.

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

## Licence

Apache-2.0. The repository's `REUSE.toml` maps every file to its licence.
