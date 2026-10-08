# vla-lang

**The VLA language with nothing around it: forms, their expansion, the
grid, and the view record.**

Frazaro is a language for spreadsheet work: you write the steps of a
procedure as English sentences, and Frazaro checks every sentence, refuses
the ones it cannot read, in words, and carries out the rest. VLA is the
Lisp beneath that English, the form every sentence translates into before
anything is emitted or built. This crate is that language alone: the reader
and the printer of its forms, the macro expander, the head table that names
the core forms, the refusal catalogue's mechanism with the language's own
refusals, the sheet model with its A1 coordinates and the references read
out of a formula's text, and the view record, one window of a sheet as the
lines a viewport draws from. No English, no file format, no evaluator yet,
no clock, and no dependency. Built for WebAssembly, the module's import
section is empty.

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
  seam a second projection enters by.
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
cargo build --release -p vla-lang --target wasm32-unknown-unknown
```

builds the crate as a WebAssembly module whose import section the
repository's `tools/check_core_imports.ps1` holds at zero entries, so that
no engine built on it can reach a socket, a clock or a file through the
language.

## What comes next

Each is an item of the repository's roadmap (`docs/BETA_ROADMAP.md`), and
this page gains its section when the item closes:

- **Recalculation's mechanism.** The dependency graph over a grid, the
  topological order, a cycle refused by name, Excel's error values, and a
  registry seam through which `frazaro-core` registers Excel's functions
  (KERNEL.7); the vectorized evaluation of a shared formula after it
  (KERNEL.20).
- **The engine's door.** A loaded grid kept between calls under a handle,
  with four calls, load, write, step and view, the previous frame as a
  read-only twin of every sheet, and a second projection, the plane, so a
  game engine can run a grid at a frame rate (the item after PORT.12).
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
