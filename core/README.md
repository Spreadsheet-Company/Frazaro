# frazaro-core

**The Frazaro language with no host: sentences in, spreadsheets out.**

Frazaro is a language for spreadsheet work. You write the steps of a procedure
as English sentences; Frazaro checks every sentence, refuses the ones it
cannot read, in words, and carries out the rest. This crate is the language
itself with nothing around it: the English engine, the compiler, a workbook
writer and a workbook reader, as one library of plain Rust whose one
dependency is [`vla-lang`](https://crates.io/crates/vla-lang), the language
itself as a crate of its own, and nothing from outside the repository. The
Excel add-in is the reference implementation; this crate is the second, held
to the same corpus of tests, so a sentence means the same thing here as there
and is refused in the same words. The
[`frazaro`](https://crates.io/crates/frazaro) crate is the command line over
it, and a web page runs it as WebAssembly.

## What it does

- **English to VLA to VBA.** `api::english_translate_text_to_vla` and
  `api::english_translate_text_to_vba` take a program's text, the prelude
  and one or more phrasebooks, and give back the program in VLA, the
  language's own form, or in the VBA the add-in would run, or a refusal with
  its id and the line it stands on.
- **A workbook from sentences.** `api::english_build_xlsx` writes a
  translated program as the bytes of an `.xlsx` workbook, the same bytes on
  every machine; `api::english_build_xlsx_into` adds the sheets to a
  workbook the caller supplies, whose own parts are copied as they were; and
  `api::english_rebuild_xlsx` reads a built workbook back, builds it again
  and says whether the two agree.
- **A workbook read as relations.** `api::reflect_relations` reads a
  workbook's bytes, an `.xlsx` package or an `.ods` OpenDocument
  spreadsheet, and prints its sheets, names, Tables, cells, formulas and
  references as relations in a fixed order, a formula from either format
  spelled as the formula bar shows it; `api::diff_relations` prints what
  differs between two workbooks; `api::audit_findings` prints the findings
  of six audit walks, one a line.
- **A window of the model as the view record.** `api::english_view`
  translates and builds a program into the sheet model, writes nothing, and
  prints one window of one sheet as the lines a viewport draws from: every
  sheet, the window and the sheet's extent, its columns and formats, then
  each cell's value or formula as the reader spells it, its format and the
  row of the sentence that wrote it. The first projection through the
  kernel's projections seam (`kernel::Projection`, `view::Grid`).
- **The same surface with C linkage.** `abi` exports the translation, the
  build, the reader (the relations, the audit, the difference), the view,
  the phrasebook gate and the version as `extern "C"` functions, for a wasm
  host or a native embedding. Built for `wasm32-unknown-unknown`, the
  module's import section is empty.

## Using it from Rust

```text
cargo add frazaro-core
```

Two files travel with the language rather than with the crate, and the
translation takes them as text: the prelude, the standard library, and a
phrasebook, the sentences of a language. The repository holds
[scripts/prelude.vla](https://raw.githubusercontent.com/Spreadsheet-Company/Frazaro/main/scripts/prelude.vla)
and
[scripts/polyglotta/english.vla](https://raw.githubusercontent.com/Spreadsheet-Company/Frazaro/main/scripts/polyglotta/english.vla);
an embedding ships them as it likes, since the core never reads a file.

```rust
use frazaro_core::api;

fn main() {
    let prelude = std::fs::read_to_string("prelude.vla").unwrap();
    let english = std::fs::read_to_string("english.vla").unwrap();
    let program = "Put 5 into cell B2.\nPut formula \"=B2*2\" into cell B3.\n";

    // The program in VLA, or the refusal with the line it stands on.
    match api::english_translate_text_to_vla(program, &prelude, &[&english]) {
        Ok(vla) => print!("{vla}"),
        Err(r) => eprintln!("{} (line {})", r.refusal, r.line),
    }

    // The same sentences as a workbook: the same bytes on every machine.
    let bytes = api::english_build_xlsx(program, &prelude, &[&english]).unwrap();
    std::fs::write("hello.xlsx", &bytes).unwrap();

    // Any workbook, read back as relations, one row a line.
    let relations = api::reflect_relations(&bytes, "hello.xlsx").unwrap();
    print!("{relations}");
}
```

A refusal is a value, never a panic: `Refusal` carries the message
catalogue's id, number, source and text, the same words the add-in shows in
a dialog, and the translation's `RefusalAtLine` adds the program line. A
phrasebook the embedding did not ship should pass `api::vocab_gate` first,
the two questions the add-in asks before loading a file: a phrasebook that
requires a capability is refused, and one holding a `(raw ...)` form needs
the person's consent.

The crate's rustdoc, on docs.rs once published, documents every module; `api`
is the surface an embedding needs, and the modules under it (`english`,
`emit`, `build`, `sheet`, `reflect`, `refers`, `view`) are the engine with its
parts named as the add-in names them; `kernel` writes down the boundary and
the six seams every one of them enters by; `excel` is Excel's function library,
the declared subset recalculation computes, registered through the sixth,
and `reflect::calc` the comparison `frazaro calc` prints. The language's own modules
(`form`, `reader`, `printer`, `expand`, `headtable`, `intrinsics`, `refers`,
`view`, the grid under `sheet`, and `calc`, recalculation's mechanism) are
`vla-lang`'s, re-exported here whole, so a
path that worked before the cut works after it.

## The contract

Bytes in, bytes out, nothing else. The crate's one dependency is `vla-lang`,
the repository's own crate of the language, and nothing in either reads a
file, a clock, a socket or a random number at run time. The prelude and the
phrasebooks are text the caller passes in; the tables the engine needs are
embedded at build time inside the crates, the head table and the language's
refusals in `vla-lang`'s `data/`, the core's refusals, the word tables and the
name lists in this crate's `data/`, exported from the add-in's own code and
held to it by the repository's checks. A workbook's bytes are deterministic: a
fixed part order, stored entries stamped 1980-01-01, no author and no date, so
the same sentences give the same file on every machine and a file can be a
test golden. The release profile aborts on panic, so the wasm module imports
nothing to unwind with and nothing to print with.

`frazaro_core::VERSION` is the crate's version, which is the release's and
the add-in's, one number for the language; `frazaro_core::ABI_VERSION`
changes only when an exported signature changes its meaning.

## With C linkage, and as WebAssembly

```text
cargo build --release -p frazaro-core --target wasm32-unknown-unknown
```

`abi` is the same surface for a host that is not Rust: `frazaro_alloc` and
`frazaro_free`, an allocator pair; `frazaro_translate_vla`,
`frazaro_translate_vba` and `frazaro_build_xlsx`; `frazaro_reflect`,
`frazaro_audit` and `frazaro_diff`, a workbook's bytes with a name for its
refusal in, the lines the command-line door prints out; `frazaro_view`, the
build's inputs with a sheet and a window in, the view record out; `frazaro_vocab_gate`;
and `frazaro_version_text` and `frazaro_abi_version`. Every answer is one record
in the module's memory: four little-endian `u32` (status, line, id length,
text length), then the id and the text, the one record whose text is not
UTF-8 being a build's, whose text is the workbook's bytes and whose id is
their SHA-256. Phrasebooks go in one buffer, NUL-separated. No binding
layer and no generated glue: the repository's web page calls these from a
few lines of JavaScript, and a native embedding loads the `cdylib` the same
way.

## What comes next

Each is an item of the repository's roadmap (`docs/BETA_ROADMAP.md`), and
this page gains its section when the item closes:

- **The engines in the core.** SQL, DATALOG, PROLOG and OPTIMIZE over
  tables, the add-in's worksheet functions as library calls (PORT.9).
- **The interpreter over the sheet model.** A program carried out in memory
  against a workbook's cells, with what needs Excel itself refused by name
  (PORT.10).
- **Embeddings.** The core under `wasmtime` from Python and .NET, in a
  desktop shell, and in an Office add-in (PORT.11).

## Where it comes from

The add-in in the repository's `src/` is the reference implementation of
the language, and the goldens it wrote under `scripts/` are what this crate
is held to, by the treaty in
[conformance/README.md](https://github.com/Spreadsheet-Company/Frazaro/blob/main/conformance/README.md):
the corpus program translated and compiled byte for byte, every phrasebook
proof passed, the build goldens reproduced to the byte, the reflect, diff
and audit goldens reproduced whole, and the token, refusal and reference
goldens the add-in writes reproduced by this crate's tests. The crate
follows the goldens and never leads them.

## Licence

Apache-2.0. The phrasebooks in the repository are under MPL-2.0, per file: a
phrasebook you write is yours. The repository's `REUSE.toml` maps every file
to its licence.
