# The conformance treaty

*What holds every implementation of Frazaro to one meaning. `PORT.4`, slice 0
of `docs/HORIZON.md` section 12. The runner is `tools/prove.ps1`; this page is
the contract it enforces. Append-only, as every ledger in `docs/` is: a change
to the treaty is added and dated, never written over.*

## The reference, and how it testifies

The VBA in `src/` is the reference implementation (`SD-18`: a second-host
engine follows the goldens, never leads them). It cannot run outside Excel, so
it does not run in CI and never will. It testifies through the goldens it
produced, committed in `scripts/` by the owner after each live pass. An
implementation that reproduces the goldens has matched the reference without
the reference being present, which is the only arrangement under which a VBA
program and a Rust crate can be held to one meaning on every push.

The day an implementation passes every oracle below and the reference cannot
run where the suite runs, the reference changes hands. That is a dated
amendment to `SD-18`, written by the owner, never inferred by a script.

## The oracles

One corpus, read by every implementation unchanged. The runner's inventory
mode (`powershell -File tools\prove.ps1`) lists them with their counts.

1. **The translate goldens.** `scripts/instructions.txt`, the corpus program
   in English, with `scripts/prelude.vla` and `scripts/polyglotta/english.vla`
   as inputs, translates to `scripts/instructions_golden.vla` and to
   `scripts/instructions_golden.vba`. The comparison is byte for byte after
   line endings are normalized to LF and trailing blank lines are dropped.
2. **The interpreter golden.** `scripts/interpreter_golden.txt`, the Interpret
   backend's report over the corpus. It needs a workbook model (slice 6); it
   is inventoried now and scored then.
3. **The phrasebook proofs.** Every `(test ...)` form in a source phrasebook
   under `scripts/polyglotta/` (`english_expanded.vla` is an export, not a
   source, and is not an oracle).
4. **The engine proofs.** Every `(test-<engine> ...)` form under
   `scripts/proofs/`, with clingo as the second-lineage oracle wherever
   `tools/*_lp.ps1` exports a proof; `tools/check_proofs.ps1` holds their
   floors.
5. **The message catalogue.** Every refusal id in `src/VLA_Messages.bas`. An
   implementation refuses with the same id in the same situation; the proofs
   that pin a refusal are the test of this, not a separate count.
6. **The pins** in `src/VLA_Tests*.bas`. VBA today, so the reference's alone;
   they join the suite as the METAPROOF line turns them into proof files,
   which land under `conformance/pins/` when they exist.

A level passed is a floor that never goes down, as `check_proofs.ps1` holds
the proof corpus and `run_checks.ps1` holds its own count.

## The contract an implementation implements

A command line, so that a Rust binary, a PowerShell wrapper around anything,
or a future door can all be scored by the same runner.

- `<impl> translate-vla <program.txt> --prelude <prelude.vla> --phrasebook <file.vla>`
  writes the VLA text to stdout and exits 0. A refusal writes its message to
  stderr and exits 1.
- `<impl> translate-vba <program.txt> --prelude <prelude.vla> --phrasebook <file.vla>`
  writes the VBA text to stdout and exits 0, with the same refusal rule.
- `<impl> prove <proofs.vla>` runs every proof in the file and writes, as its
  last line, `PASS <n>/<n>` or `FAIL <k>/<n>` where `n` is the number of
  proof forms the runner counts in the file; it exits 0 only when every
  proof passes.

An implementation attempts the oracles its slice reaches and says so by
exiting 3 for an oracle kind it does not attempt; the runner reports those as
*not attempted*, never as passed. The runner takes `-Impl <path>`, the path
to an executable or a `.ps1`; arguments to the implementation itself go in a
wrapper script, so that what was run is a file someone can read.

## The runner proves itself first

`tools/prove.ps1 -Control` writes a fake implementation that answers each
oracle from the golden itself, and a mutant that answers with one byte
changed and one proof failed. The control must pass and the mutant must
fail, or the runner is not trusted to score anything. `tools/run_checks.ps1
-WithExtras` runs that control.

## Amendment of 2026-10-01: oracle 1b, the compile golden (`PORT.5`)

`PORT.5` ports the reader, the macroexpander and the emitters before it
ports English, so the first implementation to be scored reads VLA, not
sentences. The translate goldens already hold the two ends of that road,
and the step between them is a golden of its own:

- **1b. The compile golden.** `scripts/instructions_golden.vla`, read with
  `scripts/prelude.vla`, compiles to `scripts/instructions_golden.vba`. The
  comparison is the one oracle 1 makes: byte for byte after line endings
  are normalized to LF and trailing blank lines are dropped.

The input is the golden *without its first line*. `VlaWriteGoldens`
(`src/VLA_Tests.bas`) transpiles the VLA text and only then writes it to
disk under a `; GENERATED` stamp line, so that Lint VLA refuses the file by
name; every `' vla:N` tag in the `.vba` therefore numbers the lines of the
text that was transpiled, which is the file from its second line on. The
runner drops that one line and hands the implementation the rest as a
file whose path the result names, so that what was compiled is a file
someone can read. The stamp is the writer's mark, not part of the program,
and the language's reader never learns of it.

The contract gains one command:

- `<impl> compile <program.vla> --prelude <prelude.vla>` writes the VBA
  text to stdout and exits 0. A refusal writes its message to stderr and
  exits 1; exit 3 says the oracle is not attempted, as above.

Until an implementation reproduces the whole golden, `tools/check_compile_prefix.ps1`
holds the length of the prefix it matches as a floor that never goes down,
in the shape every other check keeps: a hardcoded, reviewable number, and a
`-Control` on a fake and a mutant. The day the prefix reaches the end is
the first reproduction of a golden by anything but the VBA, and the commit
that lands it says so.

`-Control` covers the new kind: the fake answers `compile` from the `.vba`
golden and the mutant changes one byte of it.

## Amendment of 2026-10-02: oracle 1's stamp, oracle 3's count, the proofs floor, and the data both read (`PORT.6`)

`PORT.6` ports English to the core, slice by slice, and its step 0 read the
goldens writer and the phrasebook loader before the first slice. What it
found is recorded here before any implementation attempts the oracles it
touches.

**Oracle 1's `.vla` golden begins with the writer's stamp.** `VlaWriteGoldens`
writes `instructions_golden.vla` under the same `; GENERATED` line the
amendment of 2026-10-01 describes for 1b, for the same reason. The runner
drops that one line before comparing, as `Write-CompileInput` does, and the
control's fake answers `translate-vla` from the golden less its first line.
The `.vba` golden carries no stamp.

**Oracle 1's goldens also carry `F.9`'s section markers, which the product
path does not write.** Beside the stamp, the writer splices one
`(raw "' ---- instructions.txt:N label ----")` line before the first
statement of every paragraph of `instructions.txt` (`InsertSectionMarkers`,
`VLA_Tests.bas`): 153 lines in the `.vla` golden, 153 comment lines in the
`.vba` golden, and every `' vla:N` tag in the `.vba` numbers the lines of the
marked text. `EnglishToVla` and `EnglishToVba`, the functions this contract
names and `VLA_Browser.bas` exports, write no marker. So no implementation of
`translate-vla` can reproduce the `.vla` golden as it stands, and none of
`translate-vba` the `.vba`, with or without markers, since the tags would
differ. The owner decides between two readings; the decision is a later
dated amendment:

- (a) the goldens keep the markers, and the runner marks the implementation's
  `translate-vla` output before comparing (a port of `EnglishParagraphs` and
  `InsertSectionMarkers` into the runner); the `.vba` golden is then scored
  as the implementation's `compile` of the runner's marked text, and
  `translate-vba` the command is scored only against that; or
- (b) the goldens are regenerated without the markers, `VlaWriteGoldens`
  writing `EnglishToVla`'s text under the stamp alone, so that oracle 1 is
  exactly what the product's Compile does; `check_compile_prefix.ps1`'s
  floor is lowered once to the new whole golden (predicted 291,316
  characters), with its reason, and the `src:N` tags remain the diff
  reader's anchor.

Until the decision, no implementation attempts oracle 1, and the runner
scores it as the goldens stand.

**Oracle 3's count.** The runner counted a phrasebook's proofs as the lines
beginning `(test` at any indentation: 481 for `english.vla`, which is its 479
top-level forms plus two `(test-success` lines inside generator templates,
and not the three proofs the generators make. The loader runs 482: 460
`test-success` and 22 `test-fail`, as `english_expanded.vla` shows. From this
date n is the number of `test-success` and `test-fail` forms the loader runs
after the phrasebook's own generators have expanded: the runner reads it from
the `<name>_expanded.vla` export beside the source when there is one, after
checking the export's `source-hash` stamp against the source (a stale export
fails the oracle rather than counting), and from the source's own top-level
forms otherwise. An implementation's `prove` prints the n it ran. In
`FAIL k/n`, k is the number of proofs that passed, so that a floor can read it.

**The proofs floor.** Until an implementation passes a file whole,
`tools/check_prove_floors.ps1` holds, per source phrasebook, the number of
proofs the built door passes as a floor that never goes down, in the house
style, with its own `-Control`; an exit of 3 counts as 0 passed. The `core`
CI job runs it after the build.

**What a `test-fail` form pins.** The words: the sentence must refuse, and the
refusal's text must contain the fragment, compared without case and with
whitespace collapsed (`RunVocabFailTest`). The id is not pinned by the form;
the fifth oracle pins it through the proofs' own refusals.

**The data both implementations read grows.** `core/data/words.vla`
(`VLA_English.bas`'s nine word tables) and `core/data/names.vla`
(`VLA_SentenceEngine.bas`'s six name lists), exported by
`tools/export_words.ps1` and `tools/export_names.ps1` and held to the VBA by
`tools/check_data_exports.ps1`, in the shape of the head table and the
message catalogue. And `scripts/tokenize.txt` with `scripts/tokenize_golden.txt`:
programs, and the reference's token report of each (`VlaWriteTokenGolden`,
`VLA_Tests.bas`), a golden the core's own tests hold its tokenizer to, not a
runner kind.

**`SEC.2`'s consent is a door's question.** No shipped phrasebook holds a
`(raw ...)` form. The core's loader takes a flag: without it a phrasebook
holding one is refused with `english-vocab-raw-consent-declined`, the
reference's own id for consent declined; with it the text loads, as the
reference's text path loads it for `VLA_Browser.bas`. The runner's phrasebook
is the shipped `english.vla`, and the runner passes no flag. A door grants the
flag only by a person's explicit act, never by default: the command line's
`--allow-raw`, the page's checkbox.

## Amendment of 2026-10-02, later the same day: the section markers leave the goldens (`PORT.6`)

The owner's call, on reading step 0: reading (b). `VlaWriteGoldens` writes
`EnglishToVla`'s text under the stamp alone, so oracle 1's `.vla` golden is,
less its first line, exactly what the translate path produces, and its `.vba`
golden is exactly what the product's Compile makes of it. `F.9`'s markers,
and the writer's machinery that made them, are gone; the `' vla:N src:N` tag
every statement carries remains the reader's anchor into the corpus. The
regenerated goldens carry no writer's mark but the stamp. Oracle 1b's input is
unchanged in kind, the golden less its first line, now 229,156 characters;
`tools/check_compile_prefix.ps1`'s floor is lowered once, from 301,861 to
291,316, the whole of the regenerated `.vba`, with this amendment as its
reason. Reading (a) is not taken: a mark the product never writes does not
enter the contract, and the one stamp line the amendment of 2026-10-01
tolerated is the whole of the exception.

## Amendment of 2026-10-02, the third that day: a library is inventoried, not scored (`LX.15`)

Oracle 3 scores every `.vla` under `scripts/polyglotta/` that is not an
export. `alien.vla` is not a phrasebook: it declares no `<lingua>-vla` rule
and carries no proof form; it is a library of macros a program includes
(`Use library "alien.vla"`), and its one top-level call expands to a
`(sub ...)`, which the reference refuses as no directive and the port
reproduced. So: a file under `scripts/polyglotta/` with no rule and no proof
form is a library. The runner inventories it under its own kind (`library`),
attempts nothing on it and counts it in neither column;
`tools/check_prove_floors.ps1` keeps no floor for it; the control's attempted
count leaves it out. The rule is lexical, read from the file (a rule is
`^\([a-z]+-vla(-override)?\b`, a proof the form at column 0), so an
implementation needs no list of names, and a library that one day declares a
rule or a proof is a phrasebook from that commit, scored whole. Of the other
seven dialects beside `english.vla`, five were refused by the reference at
load for defects in the files, found when the port's loader read them;
`LX.15` mends the files and leaves the treaty as it was: oracle 3 is every
proof the reference passes, and after `LX.15` that is every proof in every
phrasebook.

## Amendment of 2026-10-02, the fourth that day: the refusal golden (`PORT.6`, slice 6f)

The fifth oracle says an implementation refuses with the same id in the same
situation, and that the proofs pin it. The proofs pin the refusals the
phrasebooks happen to provoke, twenty-two `test-fail` forms, by their words;
the English modules raise 125 ids. So the data both implementations read
grows by one more golden, in the token golden's shape: `scripts/refusals.txt`,
cases under `=== program <name>` and `=== phrasebook <name>` lines, each
named for the id it means to reach (a program translated after `english.vla`
loads, a `---` line dividing a phrasebook the case needs from its program;
or a phrasebook loaded after `english.vla` under the case's name), and
`scripts/refusals_golden.txt`, the reference's reading of each
(`VlaWriteRefusalGolden`, `VLA_Tests.bas`): one record per case,
`REFUSED<TAB><line><TAB><id><TAB><text>`, the line `EnglishLastErrorLine()`
for a program and `-` for a phrasebook, or `TRANSLATED` or `LOADED<TAB><n>
rules` where the reference does not refuse. It is a golden an
implementation's own tests hold to (`core/src/english/refusals.rs`), not a
runner kind; `tools/check_refusal_golden.ps1` holds its shape and its
coverage: the fixture's cases and the golden's agree, every case refuses,
every id is the catalogue's, and the distinct ids never go down, 113 of the
125 at this amendment. A case is named for an intention; the golden records
what the reference says, and a case the reference does not refuse is mended
or removed, never kept.

The twelve the golden cannot reach are a door's, elsewhere, or unreachable.
`english-program-file-not-found` and `english-vocab-file-not-found` are a
door's: the contract's commands refuse a missing program or phrasebook with
those ids and exit 1, a missing prelude with `vla-file-not-found` and a
missing VLA source with `vla-source-not-found`, the catalogue's text in every
case. `english-vocab-raw-consent-declined` and
`english-vocab-requires-capability-ungranted` are the door's two gates
(above), which the reference's text loader never asks. The two `-overwrite`
ids belong to file-writing commands the contract does not have. G-RENDER's
three are the reverse direction, VLA to English, not this path.
`english-extra-words-after-statement` reaches a caller only wrapped in
`english-test-failed-to-translate`, whose text carries it. And
`english-slot-value-not-one-form` and `english-unknown-slot-category-runtime`
are defensive arms that registration keeps unreachable.

## Amendment of 2026-10-03: oracle 7, the build golden (`PORT.7`)

`PORT.7` writes workbooks from the core: `frazaro build` turns a program's
sentences into an `.xlsx`, and the roadmap's oracle for it is "fixture
sentences to a golden `.xlsx`, byte for byte; the file opened in Excel,
Sheets, Calc and Numbers". This golden differs in kind from every other
golden here, and the treaty says so plainly. The VBA reference writes no
workbook file; it writes cells through Excel's object model, and Excel
saves. So there is no reference output to reproduce. The golden is the
core's own output, blessed by the owner opening it in Excel and looking at
what the slice promised, each live pass recorded in `docs/BETA_REARVIEW.md`
under `PORT.7`. What the oracle then holds is that the writer never drifts
from what was blessed without a regenerated golden committed beside the
change, and that a second implementation of the writer gives the same
bytes. `SD-18` is untouched: the language's meaning still flows from the
VBA through oracles 1 to 6; what the writer adds is a file format, whose
reference is ISO/IEC 29500 and the hosts that open it.

- **7. The build golden.** `scripts/build/fixture.txt`, a program in
  English, with `scripts/prelude.vla` and `scripts/polyglotta/english.vla`,
  builds to `scripts/build/fixture_golden.xlsx`. The comparison is byte for
  byte with no normalization: a workbook is a zip, and a zip is bytes. The
  golden is the one `.xlsx` the repository tracks (`.gitignore` names it).

The contract gains one command:

- `<impl> build <program.txt> --prelude <prelude.vla> --phrasebook <file.vla>
  [--phrasebook ...] --out <file.xlsx>` writes the workbook to the path
  `--out` names, writes nothing else anywhere, and exits 0. A refusal writes
  its message to stderr and exits 1; exit 3 says the oracle is not attempted,
  as above. The runner gives a path in its scratch directory and names it in
  the result, so that what was built is a file someone can open.

**Deterministic bytes, as a property of the golden.** Every entry of the
archive is stored, so a part's bytes are the bytes in the file; every entry
is stamped 1980-01-01 00:00:00, the zip epoch, the one value that reads as
no clock; the entries go in a fixed order with no extra field and no
comment; the two property parts name Frazaro and carry no date and no
person. The same sentences and the same core give the same bytes on every
machine. `tools/check_build_golden.ps1` holds the golden's length as a floor
that never goes down, in the house style; compares the door's build to the
golden, naming the first differing part and offset; and reads the methods
and the stamps off the golden itself, so that a clock, a compressor or a
dependency reaching the writer fails on every push even with a regenerated
golden. Its `-Control` passes a fake that copies the golden, fails a mutant
with one byte changed inside the worksheet part, and catches a copy with one
stamp moved. The `core` CI job runs it after the build. The runner's
`-Control` covers the kind too: the fake copies the golden to `--out`, and
the mutant changes one byte of it.

**What slice 7a's golden holds.** One sheet, `Frazaro`: the fixture's lines
in column B from row 1 exactly as written, blank lines included, and `OK`
in column C beside every line that is not blank, which is what the add-in's
Check marks (`DoCheck`, `VLA_IDE.bas`); the widths, the fills, the hidden
column A and the gridlines are `BuildWorkspace`'s. The fixture's sentences
are chosen from what slice 7b renders into cells (a value into a cell, a
formula into a cell and into rows of a column, `Work on sheet`, a write into
a cell of a named sheet), so that the golden grows by sheets, slice by
slice, and the fixture stays a program the add-in runs as well.

## Amendment of 2026-10-03, later: `rebuild` joins oracle 7 (`PORT.7`, slice 7c)

A built workbook carries a defined name, `Frazaro.Build`, holding one
string: the core's version and, for the sentences, the prelude and each
phrasebook in order, `sha256:<hex> over <n> non-whitespace bytes`, the
digest `EnglishSourceHash` gives a phrasebook (`VlaSha256HexSkippingWhitespace`,
`VLA_Digest.bas`), so that the stamp reads the same whatever line endings
the sources had. `core/src/sha256.rs` is that digest's third
implementation, and `tools/check_hash_twin.ps1` holds it to the same FIPS
vectors as the VBA and the PowerShell.

The contract gains one command:

- `<impl> rebuild <file.xlsx> --prelude <prelude.vla> --phrasebook <file.vla>
  [--phrasebook ...]` reads the sentences back out of the workbook's
  `Frazaro` sheet, checks the stamp against them and against the files
  given, builds again, compares the bytes whole, and prints one line:
  `This workbook was built from these N sentences by Frazaro <version>:
  yes.` or `...: no. <why>`, exiting 0 for yes and 1 for no. A workbook
  with no stamp, or one a host has saved since (its parts are no longer
  stored as the writer stores them), is refused with `rebuild-not-a-build`
  and exit 1; exit 3 says not attempted.

Oracle 7 now asks both: the build is the golden byte for byte, and
`rebuild` of what was built says yes. The runner's control fakes answer
`rebuild` by comparing the file to the golden, and `check_build_golden.ps1`
runs `rebuild` on the golden itself.

## Amendment of 2026-10-03, the third that day: a second golden, built into a workbook (`PORT.7`, slice 7d)

`frazaro build` takes `--into <model.xlsx>`: the program's sheets are added
to a workbook someone else made, whose parts are copied as the compressed
bytes they already are, byte for byte, and only the parts that must know
about the new sheets are edited at the one place each needs (the workbook's
sheet list, names and `calcPr`; its relationships; the content types; the
styles, whose lists grow at their ends). A sentence that writes into one of
the model's own sheets is refused (`build-into-model-sheet`), as is a file
that is not a workbook (`build-into-not-a-workbook`) or one whose part has
a shape this version does not edit (`build-into-unsupported`).

Oracle 7 gains a row. `scripts/build/into.txt`, built into
`scripts/build/model.xlsx` with the prelude and `english.vla`, is
`scripts/build/into_golden.xlsx` byte for byte. The model is a fixture
written by `tools/build_model_fixture.ps1`: deflated throughout, with a
string table, a calculation chain, a defined name and two sheets, the
shape Excel saves, so that the core's own inflate decodes real streams and
the merge is held against a package with everything a model has. The
contract's `build` takes `--into <model.xlsx>` for this row.

`rebuild` of a workbook a build was added to checks the stamp and the
build's own sheets, rendered again from what the stamp recorded of the
model (its fingerprint, how many cell formats it had, the metadata index),
and says so: the model's own parts cannot be remade without the model, so
the answer is `yes` for the sheets the build added, never for the workbook
whole. The fakes answer `rebuild` for this row as for the first, by
comparing the file to its golden.

The determinism pin (`check_build_golden.ps1`) reads the second golden
against the model: every entry is either one of the model's, copied with
the model's own method, checksum and size, or stored by this writer; every
stamp is the epoch either way, since the writer rewrites every header.

## Amendment of 2026-10-03, the fourth that day: the web door's build (`PORT.7`, slice 7e)

The web page builds the workbook too, through the core's C-ABI
(`frazaro_build_xlsx`): the rows typed are the program, the prelude and
phrasebooks inside the page are the files, and *Download as .xlsx* hands
the bytes over with their SHA-256 shown beside the button. The runner
scores an executable or a script, not a page, so the web door is held to
oracle 7 where its bytes are made: the core's own test puts
`scripts/build/fixture.txt` through the C-ABI record and requires the
first golden byte for byte, with the golden's digest as the record's id.
A page and the command-line door built from the same core therefore give
the same file for the same sentences, and the digest the page shows is the
line `frazaro build` prints, which is the check a person makes by eye.

## Amendment of 2026-10-03, the fifth that day: oracle 8, the reflect golden (`PORT.8`, slice 8a)

`PORT.8` reads workbooks in the core: `frazaro reflect` prints a workbook's
file as the relations of `REFLECT` (Stage 0.1 of `docs/SINGULARITY.md`;
the roadmap's `AXM.8`), with no host on the machine. As with oracle 7, the
VBA reference has no output to reproduce here: it reads a workbook through
Excel's object model and never opens the file, and `AXM.8`'s `REFLECT` over
COM is not built. So the relations' shape is fixed here in words before any
implementation prints them; the expected-relations files are the core's own
output, blessed by the owner opening each fixture in Excel and reading it by
eye, the fixture, the cell, the formula and the precedent named in
`docs/BETA_REARVIEW.md` under `PORT.8`; and the VBA's `REFLECT`, when it
comes, is held to the same files over COM through the `(workbook …)` clause
`AXM.8` names. The reference reader for a formula's references (`AXM.7`,
the `refers` relation) is the language's: it lands in the VBA first, with a
golden of its own in the token golden's shape, and slice 8b ports it; until
then a reader prints no `refers` row. `SD-18` is untouched.

**The relations, their order and their spelling.** One form a row, in the
proof corpus's notation. A reader prints `(sheet "<name>" <state>)` for
every sheet in tab order, the state `visible`, `hidden` or `very-hidden`;
then `(name "<name>" "<refers-to>")` for every defined name sorted by its
name without case, a sheet-scoped name spelled `Sheet!Name` with the sheet
quoted as a reference quotes it, the refers-to text as the file holds it,
and Excel's own `_xlfn.` placeholder names, which no one defined, left out;
then `(table "<name>" "<sheet>" "<range>")` for every Table by its sheet's
tab order then its name; then sheet by sheet in tab order, cell by cell in
document order (row-major), each cell's `(cell "<sheet>" "<addr>" <value>)`
when the file holds a value, then its `(formula "<sheet>" "<addr>"
"<text>")` when it holds a formula, then, from 8b, its `(refers
"<sheet>!<addr>" <to>)` rows sorted by their second field. A value is the
file's own text for a number, never a float round trip; a VLA string as
`WriteDatum` writes one for a text; `true` or `false`; `(error "#DIV/0!")`;
`(date "2026-10-03")` for an ISO date cell. A formula's text begins with
`=` and drops `_xlfn.`, `_xlfn._xlws.` and `_xlpm.`, which is what
`Range.Formula` returns and the formula bar shows; a shared formula's
children carry the first cell's text with its references moved by the
cell's distance, as Excel shows them; an array formula, legacy or dynamic,
is its anchor's text, and the cells it spills into are values alone; a
reference to another workbook keeps the file's `[n]` until 8b resolves it.
A formatted cell with no value has no row. A formula cell has a `cell` row
only when the file holds its cached value, so a workbook the writer built
and no host has opened has `formula` rows and no `cell` rows for them.
`changed` comes from `diff` (8c) alone and `ran` from no file: a reader
prints seven of the eight.

- **8. The reflect golden.** Four rows, each a fixture workbook to its
  expected relations: `scripts/reflect/fixture.xlsx` (written by
  `tools/build_reflect_fixture.ps1`, deflated, in Excel's shape, holding
  every case the walker has an arm for) to
  `scripts/reflect/fixture_relations.vla`; the two build goldens read back,
  `scripts/build/fixture_golden.xlsx` to
  `scripts/reflect/build_fixture_relations.vla` and
  `scripts/build/into_golden.xlsx` to
  `scripts/reflect/build_into_relations.vla`, which ties oracle 7 to oracle
  8 (what the writer wrote is what the reader reads); and the model the
  into golden was built into, `scripts/build/model.xlsx` to
  `scripts/reflect/model_relations.vla`. A fixture a host saved joins as a
  row the day the owner commits one (the Excel-saved copy of the first
  build golden, predicted in the rear-view to read as that golden's
  relations plus five `cell` rows). The comparison is the one oracle 1
  makes: byte for byte after line endings are normalized to LF and trailing
  blank lines are dropped; the goldens carry no stamp.

The contract gains one command:

- `<impl> reflect <file.xlsx>` writes the relations to stdout in the order
  above and exits 0. A refusal writes its message to stderr and exits 1,
  and the rows written before it stand, since a reader streams; exit 3 says
  the oracle is not attempted. A file that is not a workbook, a part holding
  a `DOCTYPE` or an entity (no entity is ever expanded), and a shape a
  version does not read are refused through the catalogue, under the source
  `VLA-Reflect`.

`tools/check_reflect_golden.ps1` holds each golden's line count as a floor
that never goes down, in the house style; compares the door's output to
each golden whole, naming the first differing line; and reads the fixed
order off each golden itself (the `sheet` rows first, then `name`, then
`table`, then the cell rows grouped by sheet in the order the `sheet` rows
gave), so that a golden regenerated by a printer that drifted fails on
every push. Its `-Control` passes a fake that prints the golden, fails a
mutant that changes one value, naming the line, and fails a copy of a
golden with a row out of order. The `core` CI job runs it after the build.
The runner's `-Control` covers the kind: the fake prints the golden beside
its fixture and the mutant changes one character.

`--counts` is a door's mode, not an oracle: counts and times alone, one line
a sheet by position and one for the workbook, so that a measurement over a
confidential model leaves nothing but numbers on the screen; it is how the
reader gives `AXM.1` its second number.

## Amendment of 2026-10-04: oracle 8's fifth row, a host's save (`PORT.8`, slice 8a's live pass)

The owner saved `scripts/build/fixture_golden.xlsx` from Excel 365 with
nothing changed, as `scripts/reflect/saved.xlsx`; it is kept as Excel wrote
it, as `model.xlsx` is, and its relations are
`scripts/reflect/saved_relations.vla`. They read as the amendment above
predicted: the first build golden's thirty rows, and a `cell` row before
each formula's `formula` row for the value the host computed, five in all.
One spelling the prediction did not name: a defined name whose text
constant is longer than 255 characters is rewritten by Excel on save as
`_xlfn._LONGTEXT("…","…")`, the literal split in two, and a `name` row
prints that text, since a `name` row holds what the file holds. A later
slice that reads names' formulas (8b, `refers`) meets it as a call with
two string arguments and no reference. The row joins the check's floors
(35) and the runner's kind; the fake answers it from its golden.

## Amendment of 2026-10-04, later: the refers golden (`AXM.7`, for `PORT.8` slice 8b)

The reference reader for a formula's references landed in the VBA first, as
the amendment of 2026-10-03 said it would: `src/VLA_Refers.bas`, `AXM.7`.
Its golden joins the data both implementations read, in the token golden's
shape: `scripts/refers.txt`, cases under `=== <name>` lines, each the cell
holding the formula spelled as a `refers` row spells its first field
(`Model!B3`, `'Q1 Data'!C5`) and then the formula with its `=`; and
`scripts/refers_golden.txt`, the reference's reading of each
(`VlaWriteRefersGolden`, `VLA_Tests.bas`): one record per reference in
formula order, `<kind><TAB><the token as written><TAB><the second field of
its refers row>`, or `NONE` where the formula holds none, then one record
`R1C1<TAB><the formula in R1C1 relative to the cell>`. The kinds are eleven:
`cell`, `range`, `column`, `row`, `name`, `structured`, `external`, `3d`,
`spill`, `unreadable` and `broken`, the last for `#REF!`, which the earlier
amendment did not name and which a real model holds after every deleted
row. The third field is the second field of oracle 8's `refers` row exactly:
a bare reference is qualified by the case's sheet, `$` marks are dropped and
letters upper-cased (`Model!B1`, `Data!A:A`, `'Q1 Data'!A1:B2`); a name, a
structured reference, an external reference and a 3D span stand as written;
a spill is its cell with `#` (`Model!A1#`); `(unreadable "INDIRECT")` and
`(unreadable "OFFSET")` come once per call, with the read going on inside
the call, so `OFFSET(A1,1,0)` still yields `A1`; a broken reference is
`Model!#REF!`, or `#REF!A1` as written after a deleted sheet. A sheet
qualifier is quoted when the name holds a character outside letters, digits,
the underscore and the period, starts with a digit, or is itself cell-shaped
or R1C1-shaped, an apostrophe inside doubled (the period's place and the two
shapes were settled by the owner's pass of 2026-10-04: `Q1.Data` is written
unquoted, `R1C1` and `A1` quoted); the same rule spells both ends of a row,
so the relation joins to itself whatever Excel's rule does in a corner. The R1C1 record is Excel's `FormulaR1C1` spelling: `R[-2]C[-1]`,
`R1C1`, `RC`, a whole column or row written once when its two ends render
alike (`C[-1]`, never `C[-1]:C[-1]`), every qualifier kept as written. Two
limits are in the fixture, not hidden: `LET` and `LAMBDA` binders read as
names, and an external name as the formula bar shows it (`Book.xlsx!Rate`)
reads as a sheet-qualified name, since without the workbook a book's name
and a sheet's are one shape; the file's `[1]!Rate` reads as external.

It is a golden an implementation's own tests hold to (`core/src/refers.rs`,
from slice 8b), not a runner kind. `tools/check_refers_golden.ps1` holds its
shape and its coverage, host-free, with a `-Control`: the fixture's cases and
the golden's agree, every case is a cell then a formula, every golden case
ends in its R1C1 record, every kind is one of the eleven, and the distinct
kinds never go down, all eleven at this amendment from 85 cases. The golden
was written by hand ahead of the reference's first run and is committed as
the prediction, so that the owner's empty `git diff` after
`? VlaWriteRefersGolden()` witnesses something. When 8b lands, the shared
formula's children of the earlier amendment are moved by the ported
scanner's third renderer, the writer's mover, which the build golden holds
and this golden does not.

## Amendment of 2026-10-04, the third that day: `refers` rows, a link's book named, two more counts and `--cone` (`PORT.8`, slice 8b)

Slice 8b ports `AXM.7`'s reader into the core, `core/src/refers.rs`, one arm
per arm of `src/VLA_Refers.bas`, held to the refers golden by its own test
as the amendment above says; and the reader prints what the amendment of
2026-10-03 deferred to it. Oracle 8's rows change in two ways, and the two
build goldens, the model and the saved copy change with the fixture:

- After a cell's `formula` row come its `refers` rows, `(refers
  "<sheet>!<addr>" <to>)`: the first field the cell as a reference spells
  it (`Model!B3`, `'Q1 Data'!B3`), the second the spelling the amendment
  above fixed for the refers golden's third field (`Model!B1`, `Data!A:A`,
  `Model!1:3`, `Rate`, `Sales[Amount]`, `Model:Scratch!B1`, `Model!A1#`,
  `[Rates.xlsx]Sheet1!A1`, `Model!#REF!`, `(unreadable "INDIRECT")`), one
  row per distinct second field, sorted by the second field as printed, so
  that a quoted reference comes before an unreadable form (`"` sorts before
  `(`). A formula with no reference has no `refers` row, and a cell with no
  formula has none. The relation is a set; the scanner's duplicates, kept in
  formula order, serve its renderers.
- A `formula` row's text names a linked workbook by its file: the file's
  `[1]` becomes `[Rates.xlsx]` through the external link's relationship, as
  the formula bar shows it while the linked book is open and as the first
  amendment predicted; a link with no file name keeps its `[n]`. The
  `refers` row spells the same. While the linked book is closed, Excel's
  bar shows the path it resolved for the link
  (`'C:\...\scripts\reflect\[Rates.xlsx]Sheet1'!A1`, the owner's pass of
  2026-10-04); the file holds no such path, and a row never prints one.

The five goldens are regenerated by the door, the owner reads the fixture
beside them, and the check's floors rise: 70 to 90, 30 to 38, 30 to 35, 11
to 13 and 35 to 43, the refers rows being 20, 8, 5, 2 and 8. Every row's
sheet is quoted by the reader's one rule (`quote_sheet`, the port of
`RefersQuoteSheet`), which quotes a cell-shaped or an R1C1-shaped sheet
name as the amendment above says; no golden changed by it. The writer's
mover, which moves a filled formula's references and renders a shared
formula's children, is a renderer over the ported scan from this slice;
the build golden and every shared child in the reflect goldens hold it
unchanged.

Two door's modes, not oracles, since each prints a time or holds nothing
a golden could compare:

- `--counts` gains two numbers a sheet and for the workbook: `distinct-r1c1`,
  the formulas compared in R1C1 relative to their cells so that a formula
  filled down counts once (the workbook's number is the distinct across its
  sheets; both are capped at 2^20), and `unreadable`, the `INDIRECT` and
  `OFFSET` calls, where a cone is blind. `AXM.1`'s instrument printed the
  same two numbers over COM.
- `--cone <Sheet!A1>`, given once or more: one line a root, `cone <root>:
  cells N formulas N inputs N blanks N sheets N depth N blind N names N
  tables N external N broken N unresolved N`, then `truncated` when the
  walk stopped at its cap of four million cells, then the times. The root
  is the one address printed, since the caller typed it. The cone is
  walked breadth-first from the root over the references the reader finds
  in each formula: a cell is one cell; a range, a column or a row is the
  cells inside it that the file holds (a blank inside a range feeds nothing
  and is not counted); a name is what the `name` relation says it refers
  to, scanned in turn, a sheet-scoped name found by the sheet written in
  front of it, then the formula's own sheet, then the workbook; a
  structured reference is its Table's whole range, the Table named in front
  of the brackets or the one the cell sits in; a 3D span is the sheets
  between its ends in tab order; a spill is its anchor; an external
  reference is counted and not followed; `#REF!` is counted as broken,
  `INDIRECT` and `OFFSET` as blind, and a sheet, name or Table the file
  does not have as unresolved. `cells` counts the distinct cells reached
  that the file holds, the root among them, `formulas` and `inputs`
  dividing them; `blanks` the single-cell references to cells the file does
  not hold; `depth` the longest chain, each cell counted at the shortest
  chain that reaches it. A root that is not a cell as `Sheet!A1` or
  `'Q1 Data'!A1`, or one naming no sheet, exits 2. This is `AXM.1`'s rule's
  clause 3 made exact: the cone Excel's `Precedents` could only put between
  a floor and a ceiling is sized through the file. The ten cones on a real
  model wait for `AXM.13`, as the counts do.

## Amendment of 2026-10-04, the fourth that day: oracle 9, the diff golden (`PORT.8`, slice 8c)

Slice 8c compares two workbooks' files through the reader of oracle 8:
`frazaro diff <old.xlsx> <new.xlsx>` prints the `changed` relation of
`REFLECT` (Stage 0.3 of `docs/SINGULARITY.md`; the roadmap's `AXM.11`,
whose `changed(addr, old, new)` from a saved copy or from another file this
is, with no host on the machine) and the sheets in one file alone. As with
oracle 8, the VBA reference has no output to reproduce: `AXM.11`'s `DIFF`
over COM is not built, so the rows' shape is fixed here in words first, the
goldens are the core's own output, blessed by the owner reading both files
of each pair in Excel, and the VBA's `DIFF`, when it comes, is held to the
same rows from its saved-copy and other-file sources. `SD-18` is untouched.

**The rows, their order and their spelling.** Sheets are matched by name
without case, as Excel names them. A sheet of the old file with no match is
`(sheet-removed "<name>" <state>)`, in the old file's tab order; a sheet of
the new file with none is `(sheet-added "<name>" <state>)`, in the new
file's tab order; then a matched sheet whose state differs is
`(sheet-changed "<name>" <old-state> <new-state>)`, in the new file's tab
order, the name as the new file spells it; the states are the `sheet` row's.
A sheet in one file alone is never matched and its cells are not listed, so
a renamed sheet is a removal and an addition, as `AXM.11` says. Every other
difference is one row of one relation, `(changed "<key>" <old> <new>)`:
first the names and Tables, which share one namespace in a workbook,
matched by name without case and sorted by it, the key a name as its `name`
row spells it (`Rate`, `Model!Local`) or a Table's name, a name's side its
refers-to text quoted and a Table's its sheet and range as `Data!A1:B5`,
`blank` where that file has none; then sheet by sheet in the new file's tab
order, cell by cell in document order, the key the cell as a reference
spells it (`Model!B3`, `'Q1 Data'!B2`), which is the spelling `refers` uses
for its second field, so that `AXM.11`'s rules join `changed` to `refers`
with no parser and a repointed name is a changed precedent of every formula
that refers to it. A cell's side is `blank` where the file holds nothing at
the address, its value as a `cell` row prints it, `(formula "<text>")` for
a formula whose cached value the file does not hold, or `(formula "<text>"
<value>)` for one it does. A cell is changed when its formula's text or its
value differs; two numbers are one value when both read as a number and are
equal (`800` and `800.0`), since a writer other than Excel spells a float
with its point and a model touched by a script would otherwise show every
float cell changed; every other value compares by kind and text, so a text
`"5"` over a number `5` is a change. A formula whose text is unchanged and
whose cached value moved is a change, which is what `cause` needs.

- **9. The diff golden.** Three rows, each a pair of fixture workbooks to
  their difference: the first build golden against its Excel-saved copy,
  `scripts/build/fixture_golden.xlsx` and `scripts/reflect/saved.xlsx` to
  `scripts/reflect/build_fixture_saved_diff.vla` (the stamp name's
  `_xlfn._LONGTEXT` rewrite and the five cached values, six rows); the
  model against the into golden built into it, `scripts/build/model.xlsx`
  and `scripts/build/into_golden.xlsx` to
  `scripts/reflect/model_into_diff.vla` (what `--into` added: three sheets
  and the stamp name, four rows); and the reader's fixture against a
  changed copy of it, `scripts/reflect/fixture.xlsx` and
  `scripts/reflect/changed.xlsx` to
  `scripts/reflect/fixture_changed_diff.vla` (one edit for each arm, listed
  in `tools/build_reflect_fixture.ps1`, which writes the copy with
  `-Changed`; twenty-four rows, and no row for a `800` that became
  `800.0`). The changed copy joins oracle 8 as its sixth row,
  `scripts/reflect/changed_relations.vla`. The comparison is the one oracle
  1 makes: byte for byte after line endings are normalized to LF and
  trailing blank lines are dropped; the goldens carry no stamp.

The contract gains one command:

- `<impl> diff <old.xlsx> <new.xlsx>` writes the rows to stdout in the order
  above and exits 0, rows or none: the rows are the answer, as `reflect`'s
  are, and two files that hold the same sheets, names, Tables and cells
  print nothing. A refusal, which is the reader's (oracle 8's three, raised
  for whichever file raised it), writes its message to stderr and exits 1,
  the rows written before it standing; exit 3 says the oracle is not
  attempted.

`tools/check_diff_golden.ps1` holds each golden's line count as a floor
that never goes down; compares the door's output to each golden whole,
naming the first differing line; and reads the fixed order off each golden
itself (the `sheet-removed`, `sheet-added` and `sheet-changed` rows in that
order and first, then the `changed` rows whose key is not shaped like a
cell, then the cell rows grouped by sheet and ascending within one), so that
a golden a drifted printer regenerated fails on every push. Its `-Control`
passes a fake that prints the golden, fails a mutant that changes one value,
naming the line, and fails a copy of a golden with a row out of order. The
`core` CI job runs it after the build, and `tools/prove.ps1` scores the
three pairs as the kind `diff`, its `-Control` covering the kind as it does
for oracle 8.

`--counts` is a door's mode, not an oracle: `sheets-removed N sheets-added N
sheets-changed N names N tables N cells N compared N` and the times,
`compared` being the addresses either file holds on the matched sheets,
nothing from inside either file.

## Amendment of 2026-10-05: oracle 10, the audit golden (`PORT.8`, slice 8d)

Slice 8d asks Stage 2.3's audit list of a workbook's file: `frazaro audit
<file.xlsx>` prints where the workbook's risks are, six of the seven
questions of the roadmap's `AXM.10` (the totals that do not foot wait for a
meaning of "total"), each a named walk over the reader's relations with no
engine and no host. As with oracles 8 and 9, the VBA reference has no
output to reproduce: `AXM.10`'s queries over the live workbook are not
built, so each walk is defined here in words first, the goldens are the
core's own output blessed by the owner reading the fixture in Excel, and the
VBA's audit, when it comes as `DATALOG` over `REFLECT`, is held to the same
findings. `SD-18` is untouched.

**The walks, in words.** A *typed-over constant* is a cell holding a value
and no formula whose nearest cells above and below in its column that the
file holds both hold formulas with one text in R1C1 relative to their own
cells. An *inconsistent formula* is a formula cell whose nearest such
neighbours agree in R1C1 while it differs. A cell at either end of a
column, or whose two neighbours do not agree, is never judged; a blank
between is passed over, since "nearest the file holds" is the rule; an
array formula's spilled cells are values, so one between two like formulas
reads as typed over, a limit the fixture avoids and this sentence records.
An *unused name* is a defined name that no formula's `refers` names and no
other name's refers-to names, a name matched by its bare text without case
and without scope; Excel's own `_xlfn.` placeholders and its `_xlnm.` names
(print areas, filter databases) are never reported, nor are Frazaro's own
marks, the build stamp `Frazaro.Build` that `rebuild` reads and the add-in's
`VLAt_` names, which no formula refers to by design (a finding every built
workbook carried would teach the reader to skip the line). An *empty reference* is
a single-cell reference in a formula, not a range, a column, a row, a name,
a spill or a link, to a sheet the file has, whose cell has no row: a
formatted blank is empty, and a sheet the file lacks is not judged here
(the cone's `unresolved` counts it). A *hidden sheet* is a `sheet` row whose
state is not `visible`. An *external link* is a formula's reference into
another workbook, as `refers` spells it.

**The rows, their order and their spelling.** One finding is one row:
`(typed-over "<cell>" <value> "<expected>")`, the cell as a reference
spells it, its value as a `cell` row prints it, and the neighbours' formula
rendered at the cell, which is the formula a fill would have put there;
`(inconsistent "<cell>" "<formula>" "<expected>")`, the formula as written
and the neighbours' rendered at the cell; `(unused-name "<name>"
"<refers-to>")`, the name as its `name` row spells it; `(empty-reference
"<cell>" "<target>")`, the formula's cell and the target as its `refers`
row spells it; `(hidden-sheet "<name>" <state>)`; `(external-link "<cell>"
"<target>")`. The order: the six relations in that order; inside one, sheet
by sheet in tab order and cell by cell in document order, a formula's
targets sorted as its `refers` rows are and listed once each, the names by
name without case, the hidden sheets in tab order.

- **10. The audit golden.** Three rows, each a fixture to its findings:
  `scripts/reflect/fixture.xlsx` to `scripts/reflect/fixture_audit.vla`
  (the fixture gains a sheet, `Review`, holding one case of each column
  walk and the two empty references; nine findings with the unused names,
  the hidden sheets and the link it had); the changed copy,
  `scripts/reflect/changed.xlsx` to `scripts/reflect/changed_audit.vla`
  (six: no hidden sheet, `Deep` unused, one empty reference, since the
  copy fills `Model!H1`); and the first build golden,
  `scripts/build/fixture_golden.xlsx` to
  `scripts/reflect/build_fixture_audit.vla` (three: `Output!C3` and `C4`
  refer to `B4` and `B5`, which nothing holds, a finding in a file the
  writer made from a sentence that filled a formula past its data). The
  reflect goldens of the fixture and the changed copy grow by the new
  sheet's forty rows, and the diff golden of the pair by one row,
  `Review!D2`, recomputed in the copy. The comparison is the one oracle 1
  makes.

The contract gains one command:

- `<impl> audit <file.xlsx>` writes the findings to stdout in the order
  above and exits 0, findings or none: a workbook with nothing to report
  prints nothing. A refusal is the reader's, exits 1, and since the index is
  built before any walk, no finding is printed before one; exit 3 says the
  oracle is not attempted.

`tools/check_audit_golden.ps1` holds each golden's line count as a floor
that never goes down; compares the door's output to each golden whole,
naming the first differing line; and reads the fixed order off each golden
itself (the six relations in order, the cell rows of one grouped by sheet
and never descending, the names never descending), so that a golden a
drifted printer regenerated fails on every push; its `-Control` passes a
fake, fails a mutant naming the line, and fails a copy with a row out of
order. The `core` CI job runs it after the build, and `tools/prove.ps1`
scores the three rows as the kind `audit`.

`--counts` is a door's mode, not an oracle: `typed-over N inconsistent N
unused-names N empty-references N hidden-sheets N external-links N` and the
times, nothing from inside the file: the number internal audit wants of a
model it may not show.

## Amendment of 2026-10-05, later: OpenDocument, the reader's second format (`PORT.8`, slice 8e)

Slice 8e reads an OpenDocument spreadsheet, `.ods`, into the same relations:
`frazaro reflect`, `diff` and `audit` take either file, and oracles 8, 9 and
10 hold both formats to one spelling. The VBA reference opens an `.ods`
through Excel, which converts it on the way in, so as with every row of
these oracles the goldens are the core's own output, blessed by the owner
opening the fixture in Excel, which opens `.ods`, and reading it by eye.

**What is read.** The `mimetype` entry must say spreadsheet; another
OpenDocument kind is refused by name. Everything else is read from
`content.xml`: the automatic styles, for a table style whose `display` is
`false` (a hidden sheet; the format has no very-hidden); each `table:table`
as a sheet in document order; the named expressions, a sheet-local one
behind its sheet as `Model!Local` (the fixture carries none sheet-local,
since Excel's reader of `.ods` removed one as unreadable content on
2026-10-05 and a fixture Excel is to open keeps its names workbook-level;
the reader's own test holds the sheet-local case), a named range's address
and a named expression's text both spelled as a `name` row's refers-to; the database
ranges as `table` rows (Calc's named rectangle with a header row, the
nearest kin of a Table; the anonymous ones autofilters leave are not); and
the other files the formulas reach. Styles, settings, the manifest and the
metadata are never decoded, and a `DOCTYPE` is refused as in a package.

**A cell's value** is what the file holds: `office:value` as written for a
float, a percentage or a currency, never a float round trip; a boolean; a
date, or a time as its ISO duration, under `date`; a string from its
paragraphs, each `text:p` a line, a `text:s` its spaces, with annotations
and drawings passed over; an error where Calc's `calcext:value-type` says
`error` or, lacking that marker, where a formula cell's string is one of
Excel's seven error literals, the gap the format leaves, or where Excel's
own writing says `office:value-type="error"`, a value type the format does
not define. A repeated row or cell is read once per repetition, which is
what the file says, and an empty one, however often repeated, costs
nothing (Excel writes the empty rows below the data as one row repeated a
million times holding one cell repeated sixteen thousand); a matrix
formula is its anchor's, an array anchor, and the cells it spills into are
values; a formatted cell with nothing in it has no row, as in a package.

**A formula's text** is OpenFormula (`of:=SUM([.B1:.B2];[Data.A1])`) and is
printed as the formula bar shows it in Excel and in Calc alike
(`=SUM(B1:B2,Data!A1)`): the namespace prefix dropped; each bracketed
reference spelled in A1, the sheet quoted by the one rule and a sheet
marked absolute with `$` unmarked, two sheets as a 3D span
(`Model:Scratch!B1`, `'Q1 Data:Q4 Data'!A1`), a source as the file's name
in brackets (`['file:///…/Rates.xlsx'#$Sheet1.A1]` as
`[Rates.xlsx]Sheet1!A1`); `;` between arguments as `,`; the reference
operators `~` and `!` as `,` and a space; `COM.MICROSOFT.` dropped from a
function's name, so that `COM.MICROSOFT.SEQUENCE(2)` reads `SEQUENCE(2)`.
Excel's own writing is read too: a name written `$$Name` reads `Name`, and
`COM.MICROSOFT.SINGLE(x)`, Excel's spelling of implicit intersection, reads
`@x`, which is what Excel's bar shows for it. A link whose address is
relative, which the format allows, Excel's reader drops as unreadable
content (the owner's openings of 2026-10-05), so the fixture's is absolute
(`'file:///C:/Rates.xlsx'#$Sheet1.A1`), and a reader prints the file's
name alone either way. So the one reference reader reads both formats and
every `refers` row keeps its spelling.

- **The rows oracles 8, 9 and 10 gain.** The fixture written again in
  ODF's terms, `scripts/reflect/opendocument.ods`
  (`tools/build_reflect_ods_fixture.ps1`, its stored `mimetype` first as
  the format requires), to `scripts/reflect/opendocument_relations.vla`
  (oracle 8's seventh row: the package fixture's 130 lines less `Broken`'s,
  with `Secret` hidden, E1 a plain range and A6 a date value); the package
  fixture against it, `scripts/reflect/fixture.xlsx` and
  `scripts/reflect/opendocument.ods` to
  `scripts/reflect/fixture_opendocument_diff.vla` (oracle 9's fourth pair:
  six rows, four of them things one format cannot say and two the scope of
  the name `Local`, sheet-local in the package and workbook-level in the
  twin for Excel's sake, and nothing else); and
  its findings, `scripts/reflect/opendocument_audit.vla` (oracle 10's
  fourth row: eight). And the twin as Excel 365 saved it on 2026-10-05,
  `scripts/reflect/opendocument_saved.ods`, kept as Excel wrote it: oracle
  8's eighth row, `scripts/reflect/opendocument_saved_relations.vla` (127
  lines: Excel's own dialect, `Q1 Data` renamed `Q1_Data`, the link's
  formula gone); oracle 9's fifth pair, the twin against it,
  `scripts/reflect/opendocument_opendocument_saved_diff.vla` (nine rows,
  what Excel's ODF writing changes and nothing else); oracle 10's fifth
  row, `scripts/reflect/opendocument_saved_audit.vla` (seven, the link's
  finding gone with its formula). The comparison is the one oracle 1
  makes.

`tools/check_reflect_golden.ps1`, `check_diff_golden.ps1` and
`check_audit_golden.ps1` each gain the rows with their floors, and
`tools/prove.ps1` the six rows.

## Amendment of 2026-10-05, later still: the web door's reader (`PORT.8`, slice 8f)

The web page reads a workbook too, through the core's C-ABI
(`frazaro_reflect`, `frazaro_audit`, `frazaro_diff`): a file the person
picks is read by the browser into the module's memory, and the Reflect
pane prints the record's text, the lines `frazaro reflect`, `audit` or
`diff` print for the same file. The runner scores an executable or a
script, not a page, so the web door is held to oracles 8, 9 and 10 where
its text is made: the core's own test puts `scripts/reflect/fixture.xlsx`,
its changed copy and the OpenDocument twin through the C-ABI records and
requires the goldens whole, `fixture_relations.vla`, `fixture_audit.vla`,
`fixture_changed_diff.vla` and `opendocument_relations.vla`. A page and
the command-line door built from the same core therefore print the same
lines for the same file, and what a person copies from the pane is what
the door prints. A refusal names the file as the page gave its name, where
the door prints its path, and is otherwise the catalogue's text; a name
the page does not give reads `the file`.

## Amendment of 2026-10-05, the egress golden (`SEC.15`, the formula sink)

Every formula a program writes now passes one sink per implementation,
and the sink refuses by name a formula that reaches outside the workbook on
its own: in the reference, `VlaSetFormula` (`src/VLA_Runtime.bas`, above
the inject boundary, so the compiled program carries it), which Compile's
emitted code calls where it wrote `.Formula2 =` and Interpret's `Formula`
member calls too; in the core, `build`'s `write_formula`, so that
`frazaro build` and the web page's download refuse the same text
(`build-formula-egress`, the line, the function and the sentence; the
reference's two backends say it with the runtime's own `rt-formula-egress`,
the function and the cell, through the program's step report). The scan
is the reference's `VlaFormulaEgress`, ported as `core/src/egress.rs`: any
text written through the formula member, whatever it begins with; a string
literal, a quoted name and a bracket group stepped over; a run of
identifier characters followed by spaces, tabs or line breaks and then `(`
a call, the file prefixes `_xlfn.` and `_xlws.` stripped from its front,
the rest compared without case to the list whole; a `|` outside the three
groups a DDE link; the first hit in text order named in upper case, a DDE
link as `DDE`. The list is 23 names, each with its reason in SEC.15's
entry: `WEBSERVICE`, `FILTERXML`, `HYPERLINK`, `RTD`, `IMAGE`,
`STOCKHISTORY`, `CALL`, `REGISTER`, `REGISTER.ID`, `EXEC`, `INITIATE`,
`EXECUTE`, `POKE`, `REQUEST`, `SEND.MAIL`, `DDE`, `IMPORTXML`,
`IMPORTDATA`, `IMPORTHTML`, `IMPORTRANGE`, `IMPORTFEED`, `GOOGLEFINANCE`
and `GOOGLETRANSLATE`.

The golden joins the data both implementations read, in the token golden's
shape: `scripts/egress.txt`, cases under `=== <name>` lines, each one text
as the sink receives it (joined with a line break where there are more
than one); and `scripts/egress_golden.txt`, the reference's reading of each
(`VlaWriteEgressGolden`, `VLA_Tests.bas`): one record, `REFUSED<TAB><the
name>` or `WRITTEN`. It is a golden an implementation's own tests hold to
(`core/src/egress.rs`, `the_egress_golden_is_reproduced`), not a runner
kind. `tools/check_egress_golden.ps1` holds its shape and its coverage,
host-free, with a `-Control`: the fixture's cases and the golden's agree in
order, every record is `WRITTEN` or `REFUSED` with a listed name or `DDE`,
the reference's list (the `listed = "…"` line of `VlaFormulaEgress`) and
the core's (`EGRESS_NAMES`) hold the same names, every listed name is
reached by a `REFUSED` record and `DDE` by the link shape, and the list
never shrinks below its floor, 23 at this amendment from 55 cases, 37
refused. The golden was written by hand ahead of the reference's first run
and is committed as the prediction, so that the owner's empty `git diff`
after `? VlaWriteEgressGolden()` witnesses something. Oracle 1b's golden
changes with the emitter, its eighteen formula writes now calls to
`VlaSetFormula`; the comparison is unchanged in kind.

## Amendment of 2026-10-07: oracle 11, the view golden (`KERNEL.4`, the view record)

`KERNEL.4` draws one window of the sheet model as the view record, the
first implementation of the kernel's projections seam (`CONTRIBUTING.md`,
"The kernel and its seams"; `web/CALLOSUM.md` §7, decisions 1 and 2):
`frazaro view` translates a program and builds it as `frazaro build`
builds it, into the model in memory, writes nothing, and prints one window
of one sheet as the lines a viewport draws from. As with oracles 7 to 10,
the VBA reference has no output to reproduce here: it draws nothing, since
Excel is its view. So the record's shape is fixed here in words before a
viewport reads it; the goldens are the core's own output, blessed by the
owner reading each against the fixture workbook open in Excel; and
`SD-18` is untouched, the record being a projection of the model the
writer builds and the reader reads, never a new meaning.

**The record, its order and its spelling.** One form a row, in the proof
corpus's notation. First `(sheet "<name>" visible)` for every sheet of
the model in tab order (a model nothing has hidden); then one `(window
"<sheet>" "<range>")`, the window as asked, its sheet spelled as the
model spells it and its range as a reference spells one (`B2` for one
cell, `A1:F20` for more); then one `(extent "<sheet>" "<range>")`, the
sheet's used rectangle, or `(extent "<sheet>" none)` for a sheet with no
cell; then one `(gridlines "<sheet>" on|off)`; then `(column "<sheet>"
"<letters>" <width|none> shown|hidden <format|none>)` for every column of
the window's span that has settings, in column order, the width in
characters as `ColumnWidth` counts them; then `(format <index> <fill|none>
general|text wrap|nowrap)` for format 0 and every format the window's
columns and cells use, ascending, the fill six hex digits; then the cells
of the window in row-major order, each as the reader spells it, `(cell
"<sheet>" "<addr>" <value>)` for a value or `(formula "<sheet>" "<addr>"
"<text>")` for a formula, followed by `(style "<sheet>" "<addr>" <index>)`
when the cell's format is not 0 and `(sentence "<sheet>" "<addr>" <row>)`
when a sentence wrote it, the row being the program line, which is the row
of the `Frazaro` sheet holding that sentence. A value is a literal or a
folded formula; a formula is text the host would compute, and the record
never shows one as the other, since the core computes nothing until
recalculation lands (`KERNEL.7`). A shared formula's children show the
master's text with its references moved, as the reader shows them; a
dynamic-array formula shows its text in each cell it was written to.

**Without a window** the door shows the sheet's whole extent, or `A1`
alone for a sheet with no cell. A sheet the program does not make is
refused naming the ones it does (`view-sheet-unknown`); a text that is not
a rectangle of cells, a whole column or row included, is refused as
written (`view-window-not-a-range`); both under the source `VLA-View`, and
both the core's alone, as the writer's and the reader's refusals are.

- **11. The view golden.** Six rows, each a fixture program, a sheet and a
  window to its expected record under `scripts/view/`:
  `scripts/build/fixture.txt`, sheet `Frazaro`, no window, to
  `fixture_frazaro.vla`; sheet `Output`, window `A1:F20`, to
  `fixture_output.vla`; sheet `Output`, window `B2:C3`, to
  `fixture_output_b2_c3.vla`; sheet `Data`, no window, to
  `fixture_data.vla`; `scripts/build/into.txt`, sheet `Output`, no window,
  to `into_output.vla`, and sheet `Checks`, no window, to
  `into_checks.vla` (the second fixture viewed without its model: its
  formulas read the model as text, which a view can show). The comparison
  is the one oracle 8 makes: byte for byte after line endings are
  normalized to LF and trailing blank lines are dropped; the goldens carry
  no stamp.

**The free oracle.** The window is drawn from the model; the file the
writer writes is written from the model; the reader reads the file. So the
`sheet`, `cell` and `formula` rows of a built model's sheet viewed whole
are `reflect`'s rows for that sheet of the written file, row for row, and
the core's own test (`core/src/view.rs`,
`the_view_of_a_built_model_is_the_reflect_of_its_file`) holds it for both
build fixtures. Oracles 7, 8 and 11 are one model seen three ways.

The contract gains one command:

- `<impl> view <program.txt> --sheet <name> [--window <A1:F20>] --prelude
  <prelude.vla> --phrasebook <file.vla>` writes the record to stdout in
  the order above and exits 0. A refusal, the translation's, the build's
  or the view's, writes its message to stderr and exits 1; exit 3 says the
  oracle is not attempted.

`tools/check_view_golden.ps1` holds each golden's line count as a floor
that never goes down, in the house style; compares the door's output to
each golden whole, the corpus prelude and `english.vla` named, naming the
first differing line; and reads the fixed order off each golden itself
(the `sheet` rows, one `window`, one `extent`, one `gridlines`, the
`column` rows, the `format` rows with 0 first and ascending, then the
cells in row-major order with each cell's `style` and `sentence` rows
right after it, every body row on the window's sheet), so that a golden
regenerated by a printer that drifted fails on every push. Its `-Control`
passes a fake that prints the golden, fails a mutant that changes one
value, naming the line, and fails a copy of a golden with a format row
moved below a cell row. The `core` CI job runs it after the build. The
runner gains the kind: `tools/prove.ps1` inventories the six rows and
scores them as it scores the reflect kind, the corpus prelude and
`english.vla` named as the build kind names them; its `-Control` fake
answers each from the golden of the window, found by the program's name,
the sheet and the window, and its mutant changes one character.

## Amendment of 2026-10-07, later: `prove` over a distro folder (`KERNEL.2`)

A distro is a folder, the unit the build tools take (`web/CALLOSUM.md`
§14.1; `distros/english/README.md` has the shape): its `distro.vla` names
the prelude, the base phrasebook, an overlay an edition loads after it, the
dialects a language picker offers, the libraries and the examples, by
reference, each path relative to the folder. Three readers take the one
file, the page builder, the add-in builder and the core, and
`tools/check_distro.ps1` holds every manifest to the shape all three rely
on. The treaty gains no new golden: a distro is proved by the proofs its
phrasebooks already carry, oracle 3 applied to the folder.

The contract gains the folder form of one command:

- `<impl> prove <folder>`, the folder holding `distro.vla`, proves the
  distro whole and writes one line a book: the base phrasebook alone, as
  oracle 3 proves a file (`english: PASS 482/482`), then the overlay and
  each dialect loaded over the base, as the page and the add-in load them
  (`pirate over english: PASS 20/20`), a library listed and not scored, and
  as its last line the total, `PASS n/n` or `FAIL k/n`, where n is the sum
  over the books the manifest names; exit 0 only when every book passes.
  Each failing proof is printed before its book's line, as the file form
  prints it. The manifest's prelude is the prelude unless `--prelude` names
  another; a folder with no `distro.vla` is refused
  (`distro-manifest-missing`), a manifest the implementation cannot read
  is refused naming what is wrong (`distro-manifest-invalid`), and a file
  the manifest names that is not there is refused as a named file is.

The runner gains the kind: `tools/prove.ps1` inventories `distros/english`
and `distros/espanol` and scores each as it scores a phrasebook, n summed
over the manifest's books with each book counted as the file form counts
it (a book the manifest names twice, as the espanol distro names its
overlay among the dialects, counts twice, since the door proves it twice);
its `-Control` fake reads the manifest and answers `PASS` per book and in
total, and its mutant answers `FAIL`. At this amendment: english, 632
proofs over eight books; espanol, 662 over nine.

## Amendment of 2026-10-08: oracle 12, the calc golden, and oracle 11's `value` row (`KERNEL.7`, recalculation)

`KERNEL.7` lands recalculation's first slice: a parser for operators and
calls over the reference scanner `refers.rs` already has, the dependency
graph from `refers`, evaluation in topological order, a cycle refused by
name, Excel's error values as a value kind, and a registry seam for
functions, in the language crate (`vla-lang/src/calc/`), with Excel's
library registered through the seam from `frazaro-core` (`core/src/excel.rs`)
as the declared subset. The VBA reference has no recalculation of its own:
Excel is its evaluator, so the reference for this oracle is Excel's own
saved values, as `web/CALLOSUM.md` §7 decision 3 recorded when it amended
`HORIZON.md` §12.2 and as this treaty's decision 7 there foresaw. `SD-18`
is untouched: the language's forms, the grid and the references are still
held to the goldens the add-in writes, and what this amendment adds is a
second reference for one thing the add-in never did.

**The subset, by measurement.** A function enters the subset with its
fixture and its measured frequency, never by taste (`KERNEL.6`'s rule).
At this amendment the subset is the fourteen functions the engine's first
cartridges need (`Alonzo/SPEC.md` §13: `IF`, `AND`, `OR`, `NOT`, `SUM`,
`MIN`, `MAX`, `ABS`, `INT`, `MOD`, `ROW`, `COLUMN`, `CHOOSE`, `SIN`, the
comparisons and the arithmetic), which are the language's own; six more by
the two corpora of Jansen and Hermans (2015, Table IV: `AVERAGE`, `ROUND`,
`SUMIF`, `COUNTIF`, `ISNUMBER`, `ISBLANK`); and `IFS`, on its fixture
`scripts/reflect/saved.xlsx`. Every function's evidence is a row of
`excel::SUBSET`, and the one line that admits a function is a row there with
its fixture. A formula outside the subset is not an error: its cell shows
`not computed here` with the function named, and a cell that reads it
inherits the label.

**The tolerance, stated before the oracle.** The language's own functions
carry none: their tests hold exact values. The agreement of a computed
value with a host's cached one carries one stated tolerance: two numbers
agree when they round to the same fifteen significant digits, the
precision Excel documents for its arithmetic; two texts when equal byte for
byte; two truth values or two errors when the same. A number the file
spells that no double holds, a date cell, and an error outside the seven a
cell can hold (`#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`,
`#N/A`) are not compared. The negative zero a double can hold is spelled
`0`, since no cell holds one.

**The record, its order and its spelling.** `frazaro calc <file>` reads
the workbook through the reader of oracle 8 into the grid, computes every
formula, and prints one form a row, in the proof corpus's notation. First
`(cycle "<Sheet!A1>" "<Sheet!B2>" ...)` for every cycle, its cells sorted,
the cycles by their first cell; then one `(calc "<sheet>" "<addr>"
"<formula>" <computed> <cached> <verdict>)` per formula cell, sheets in tab
order and cells in row-major order within a sheet, the formula as the bar
shows it with its `=`, the computed value spelled as a `cell` row's value
is (`15`, `"big"`, `true`, `(error "#DIV/0!")`) or `(not-computed
"<name>")` naming the function or the construct this version does not
read, or `(not-computed cycle)` for a cell in a cycle or reading one; the
cached value the file holds, spelled as the reader spells it, or `none`;
and the verdict, `agree`, `differ` or `unchecked`, the last whenever there
is nothing to compare. The constructs not read in this slice, each named
as written: an array constant, implicit intersection by `@` or by a space,
a union in parentheses, a spill range, a structured reference, a link into
another workbook, and a range where one value is wanted, which a legacy
formula intersects and a dynamic one spills, and which the text alone does
not tell apart (`KERNEL.8` takes them with the `@` operator). A cycle, an
unreadable reference and a function outside the subset refuse nothing
here, since a workbook is read as it is; the engine's `load` refuses the
first two by name (`calc-cycle`, `calc-function-not-computed`, under the
source `VLA-Calc`, the language's own), as `Alonzo/SPEC.md` §2 asks.

- **12. The calc golden.** Seven rows, each a fixture workbook to its
  recalculation under `scripts/recalc/`: `scripts/reflect/saved.xlsx`, the
  first build golden as Excel 365 saved it, to `saved_calc.vla`, every row
  `agree`; `scripts/reflect/fixture.xlsx`, the reader's fixture with every
  reference kind, to `fixture_calc.vla`, no row `differ`; its OpenDocument
  twin `scripts/reflect/opendocument.ods` to `opendocument_calc.vla`, no
  row `differ`; `scripts/build/fixture_golden.xlsx`, the core's own build
  with no cached value, to `build_fixture_calc.vla`, every row `unchecked`;
  `scripts/recalc/subset.xlsx`, the subset's fixture, built by the door
  from `scripts/recalc/subset.txt` with one formula per function and per
  coercion this slice claims, to `subset_calc.vla`, every row `unchecked`;
  `scripts/recalc/subset_saved.xlsx`, the owner's save of that fixture from
  Excel 365, to `subset_saved_calc.vla`, no row `differ` and 102 `agree`,
  the four this slice does not compute named; and `scripts/recalc/edges.xlsx`,
  built from `edges.txt`, to `edges_calc.vla`, every row `unchecked` until an
  Excel save of it beside, `edges_saved.xlsx`, joins this list at no row
  `differ`. That is how every rule of coercion enters the oracle with the
  fixture that shows it: Excel's save of the subset corrected two of this
  slice's first rules the evening it was made, a number's General spelling
  (fixed while the decimal exponent is between -14 and 14, scientific
  outside, the boundaries LibreOffice's automatic number format draws and
  the edges fixture holds) and a text read through a reference in a logical
  test (`IF` coerces the text `TRUE` found in a cell as it coerces a typed
  one). The comparison is the one oracle 8 makes: byte for byte after line
  endings are normalized to LF and trailing blank lines are dropped; the
  goldens carry no stamp.

**Oracle 11 gains a row.** After each `(formula ...)` row of the view record
comes `(value "<sheet>" "<addr>" <v>)`, the formula's computed value under
the same library, spelled as a `cell` row's value is, or `(not-computed
...)` as above: the row `Alonzo/SPEC.md` §4.5 and decision 14 named for
this item, so that the record stays the oracle of the engine's plane. A
view of a program built without its model shows the values its own sheets
give, so a reference to a sheet or a name the model lacks is `#REF!` or
`#NAME?`, as Excel would show them in a workbook lacking them
(`into_checks.vla`). The six view goldens were regenerated with the row,
`check_view_golden.ps1`'s floors rose with them (26, 18, 12 and 10; the two
windows with no formula unchanged), and its order rule places the `value`
row right after its formula row, before the `style` and `sentence` rows. A
grid before its first step shows no `value` row, which is `view_text`'s
record and the engine's at load.

**The free oracle.** The study's workbooks (`scripts/study/`, `KERNEL.3`)
carry cached values a PowerShell evaluator computed, so they are no
fixture of this oracle; but `check_recalc_golden.ps1` holds their
agreement, 80, 80, 8 and 8 formulas with none differing, two lineages of
evaluation agreeing on every value, with no golden stored.

The contract gains one command and one mode:

- `<impl> calc <file.xlsx|.ods>` writes the rows to stdout in the order
  above and exits 0 whenever the calculation ran, `agree` or `differ`;
  with `--counts`, one line of counts and times alone (`formulas`,
  `computed`, `not-computed`, `cycles`, `agree`, `differ`, `unchecked`).
  A refusal, the reader's, writes its message to stderr and exits 1; exit 3
  says the oracle is not attempted.
- `<impl> reflect <file> --functions` is a door's mode, not an oracle:
  one `(function "<NAME>" <n>)` row per function the formulas call, `n`
  the formula cells calling it, most called first, counts alone, so that
  `KERNEL.6`'s histogram can be taken over anyone's models.

`tools/check_recalc_golden.ps1` holds each golden's line count and its
`agree` count as floors that never go down, in the house style; the fixed
order and the row shapes read off the golden itself; each fixture's pin
(every row `agree`, no row `differ`, or every row `none unchecked`); the
door's rows against each golden whole, naming the first differing line;
and the study's agreement. Its `-Control` passes a fake that prints the
golden and the agreement, fails a mutant that changes one value and reports
one `differ`, fails a copy of the saved golden with one `agree` turned to
`differ`, and fails a copy with a cycle row after the calc rows. The `core`
CI job runs it after the view check. The runner gains the kind:
`tools/prove.ps1` inventories the seven rows and scores them as it scores
the reflect kind; its `-Control` fake answers each from the golden by the
fixture's name, and its mutant changes one character.

## Amendment of 2026-10-09: oracle 11's `row` and `look` rows, a sheet's own state, and the loader its inverse (`KERNEL.22`, the handle and the four calls)

An engine keeps a grid in the language's memory and views it through the
same record (`vla_lang::machine`; `Alonzo/SPEC.md` §4), which asks three
things of oracle 11 and moves no golden, since no model a program builds
holds any of them:

- **A `(sheet ...)` row prints the sheet's own state**, `visible`, `hidden`
  or `very-hidden`, the reader's three words, where every row said
  `visible` before. An engine's twin of a sheet `X`, `X.last`, the previous
  frame, is `hidden` and listed after the sheets it twins, in their order.
- **`(row "<sheet>" <n> <height|none> shown|hidden)`**, one row per row of
  the window that has settings, ascending, after the `column` rows: the
  `column` row's twin (`Alonzo/SPEC.md` decision 19).
- **`(look "<sheet>" <value> <format>)`**, one row per entry of the sheet's
  look map, in the order declared, after the `format` rows, each look's
  format among the formats printed: every cell of the sheet holding the
  value draws with that format in place of its own (the same decision).

**The record's inverse.** The language's loader (`vla_lang::machine::load`)
takes every row the record and the reader print, `sheet`, `cell` and
`formula` over a cell or a range, `value` after its formula, `name`,
`gridlines`, `column`, `row`, `format`, `look`, `style` and `sentence`, and
skips by name `window`, `extent`, `refers` and `table`; a record of a whole
sheet loaded and viewed again is the same record, row for row. That is
held by the language's test over the five view goldens whose window holds
the sheet's extent (the sixth, `fixture_output_b2_c3`, clips its window
inside the sheet, so its `extent` row is not the loaded cells') and over a
model holding every row kind, stepped and saved, as an engine's save is
made. No new oracle number: the four calls are one implementation's
surface, with no second to hold to them; 13 stays the expansion golden's.

## Amendment of 2026-10-10: a fresh build holds the values it computes (`KERNEL.20`, its fifth slice)

A workbook the core builds is computed once under the declared subset, as
the view record's `value` rows are, and each formula cell the subset
computes holds its value, as a file a host saved holds the values the host
computed: a number in the writer's spelling, zero as `0`; a text with
`t="str"`; a truth value with `t="b"`; an error with `t="e"`; a number no
cell can hold as `#NUM!`, as a value cell spells it. A formula the subset
does not compute (a function outside it, a construct not read, a cycle, a
cell reading one) holds none. The workbook's `calcPr` still asks the host
to compute every formula when it opens the file, so a host shows its own
values and a reader that does not calculate reads the build's. A build
added to a model (`--into`) holds no value, since its formulas may read the
model's own cells and names, which the build does not read.

What it moves:

- **Oracle 7.** The first build golden is regenerated, each of its five
  formulas holding its value (11,443 bytes, the floor raised); the into
  golden is the same bytes. `rebuild` builds again and compares whole, as
  before, and says yes to both.
- **Oracle 8.** The slice-8a sentence that a workbook the writer built and
  no host has opened has no `cell` rows for its formulas now holds only for
  a build added to a model: `build_fixture_relations.vla` gains a `cell` row
  before each of its five formula rows, the five `saved_relations.vla`
  shows Excel computing, and is 43 rows.
- **Oracle 9.** `build_fixture_saved_diff.vla`, the build golden against
  Excel's save of it, loses its five `changed` rows: the values the build
  holds are the ones Excel computed and saved, spelled as Excel spelled
  them. The stamp's rewrite is the one row left, and the floor is lowered
  from 6 to 1 for that reason.
- **Oracle 12.** The three fixtures the door builds, `fixture_golden.xlsx`,
  `subset.xlsx` and `edges.xlsx`, are built again and hold the build's own
  values, so their rows read `agree` where the subset computes (5, 102 and
  16) and `none unchecked` where it does not (the subset's four). Their
  `agree` is agreement with the build's own evaluation, not with a host;
  the rows that hold the evaluator to a host stay `saved_calc.vla` and
  `subset_saved_calc.vla`. `check_recalc_golden.ps1` pins the three
  `built`: no row differs, and no value the evaluator computes is missing
  from the file.

The measurement the amendment rests on: the subset's fixture built again
and compared with the owner's Excel 365 save of it (`subset_saved.xlsx`)
differs in six rows, where the build without values differed in 107: the
stamp, a formula Excel respelled (`1E307` as `1E+307`, both `#NUM!`), and
the four formulas this version does not compute. Every one of the 102
values the build holds is the value Excel saved, spelled as Excel saved it.

## Amendment of 2026-10-10: a fill Excel cannot share, written one formula a cell (`PORT.13`)

The writer held every formula filled into a range as one shared formula,
and Excel refused to open a build whose fill read another sheet: [MS-XLS]
`SharedParsedFormula` forbids a shared formula a reference to another sheet
or workbook, an array constant, a union, an intersection, a range operator
and a `#REF!`, and Excel writes such a fill one formula a cell. Now, so
does the writer: each cell of a fill holding any of these, or a table's
column, a spill, a newer function, or a text the evaluator's parser does
not read, holds its own formula, the first cell's with its references
moved. A fill that holds none of them is one shared formula as before, the
shared formulas left in a sheet numbered from 0 in their order.

- **Oracle 7 gains a row.** `scripts/build/sheets.txt`, built with the
  prelude and `english.vla`, is `scripts/build/sheets_golden.xlsx` byte for
  byte: two fills that read the sheet `data`, each cell its own formula,
  beside a fill that reads only its own sheet and stays shared, blessed as
  the build goldens are, by the owner opening it in Excel. The two goldens
  before it are the same bytes, neither holding a fill Excel cannot share;
  `rebuild` says yes to all three, and the runner's fakes answer the new row
  as the first, by its golden.
- **Oracles 8 to 12 are unchanged.** A reader prints each cell of a fill as
  its formula, whether the file shares it or not, so no relation, change,
  finding, view or recalculation row moves.

## Amendment of 2026-10-12: oracle 11's values, the record less its formula rows (`KERNEL.37`)

An engine draws a running grid from the view record every frame, and the
record spells every formula's text at every view, each shared formula's
cell its master's text moved to it: an 80 by 50 window of one short fill is
viewed in about 10 ms where the frame computing it takes 0.2. A host draws
values; a formula's text is the bar's, one cell at a time. So the language
prints a second projection of the same window, the values: the view record
with its `formula` rows left out.

- **The values record.** Every row of the view record of the same window
  over the same values, in the record's order and spelling, but the
  `(formula ...)` rows: the `sheet`, `window`, `extent`, `gridlines`,
  `column`, `row`, `format` and `look` rows, then each cell's `cell` row, or
  its formula's `value` row once it has one, then the cell's `style` and
  `sentence` rows. A formula with no value prints no row of its own, its
  `style` and `sentence` rows kept; a shared formula's cell whose master the
  sheet no longer holds has no text, and prints nothing in either record. So
  for any window the values record is the view record with its lines that
  open `(formula ` removed, byte for byte, which is the oracle and the
  record's whole definition.
- **No golden moves and none is added.** No door prints the values yet: an
  engine asks for them by name (`values`, beside `grid` and `plane`), and a
  door's choice among projections waits for the projections seam
  (`KERNEL.12`). The six view goldens, each less its `formula` lines, are
  the values of their windows, held by the core's test through the door's
  own recalculation and by the language's over the drafts loaded from five
  of them, and a law holds the oracle over generated grids before a frame,
  after each and after each write between frames.
- **The extent where the values keep it.** Both records read a sheet's
  `extent` row, and the window of a sheet asked for without a range, from
  the extent their values keep, where they keep one: an engine's machine
  keeps every sheet's for its evaluator, and a recalculation those of the
  model it was made over. The rows are the ones a walk of the sheet gives,
  which the same law holds.

## Amendment of 2026-10-15: oracle 11's changes, the rows a frame changed (`KERNEL.42`)

A host drawing a running grid's values viewed its whole window at every
frame to learn whether anything had changed, and most frames of a game's
sheets change nothing. An engine's machine knows which cells a frame
changed, so it answers a window's changes: what to replace in the values a
host holds, and nothing when nothing changed.

- **The changes.** A `(changes "<sheet>" "<range>" <since> <at> <n>)` row,
  the sheet and the range as the values' own `window` row spells them,
  `at` the last complete frame and `since` the one before it, `n` the cells
  changed; then each changed cell's rows as the values of the same window
  print them, in their order: its `cell` row or its formula's `value` row,
  then its `style` and `sentence` rows; a cell that prints no row now,
  `(empty "<sheet>" "<addr>")`. A cell is changed when its rows are, as the
  values spell them: a formula's 0 turned to -0, its `value` row 0 both
  times, is not, where a twin's `cell` row, which spells -0, is.
- **Whole.** Where the machine cannot say which cells changed, the row ends
  `whole` in place of `n` and no row follows: before the first frame; at a
  frame that changed more cells than it keeps room for, or gave a twin a
  cell or took one; after a write that changed a formula or added a cell,
  until the next frame's end and at it; and while a value written since the
  last frame lies in the window. The host then views the window's values.
- **The oracle, free.** A window's values at frame `since`, each cell the
  changes name given their rows in place of its own, are its values now,
  byte for byte, and the changes name a cell exactly when its rows
  changed, so none exactly when the two are equal, or the changes answer
  whole: held by the language's eleventh law
  over generated grids, every sheet's and twin's window, right after a
  frame, after the writes before the next and while it yields, and by the
  count of the cells a view visits, a still window's changes visiting none.
  No door prints them, a door having no frames, and no golden moves.

## Amendment of 2026-10-17: oracle 11's number formats, a code and a number's text (`KERNEL.47`)

The record's formats were `general` or `text`, and it spells a number at
its shortest round trip, which a viewport never respells, so a cell
holding 1 drew `1` where a workbook's author wrote `0.00`. A format may now
hold a code of Excel's, and the record prints the text the code draws.

- **The code.** A `format` row's third field is `general`, `text` or a code
  in quotes, written as `Range.NumberFormat` takes it, spelled the same on
  every machine: `(format 2 none "0.00" nowrap)`. The language draws one to
  four sections, `0` and `#`, the point, the thousands, thousandths and
  percents, and literals; a code holding anything else is refused by name
  (`grid-format-code`), and so are `"General"` and `"@"`, which are spelled
  `general` and `text`.
- **The text.** A cell whose number draws under a code, or whose text draws
  under a code's fourth section, prints `(text "<sheet>" "<addr>" "<text>")`
  after its `cell` or `value` row and before its `style` row: the text
  Excel's `Range.Text` gives it in a column wide enough, under the code of
  the format it draws with, its look's where its value has one. A number
  is taken at its nearest fifteen significant digits, an exact tie toward
  zero, then rounded half away from zero to the code's decimals, and one
  section's minus is dropped where the number draws as zero, as Excel
  draws them. The grid, the values and a window's changes print the row
  alike; a load skips it, so that a save loads as before; and a record
  whose formats hold no code prints none, so that no golden of oracle 11
  moves.
- **The oracle.** Excel's own text: `scripts/formats.txt` holds codes and
  the values each draws, and `scripts/formats_golden.txt` is what
  `VlaWriteFormatsGolden` (`VLA_Tests.bas`) writes of them in Excel, each
  code as Excel reads it back, or Excel's refusal, and each value's
  `Range.Text`. The language's test (`vla-lang/tests/formats_golden.rs`)
  holds every case to that file whole, and
  `tools/check_formats_golden.ps1` holds the pair's shape and the
  constructs the cases reach. The language's third, ninth and eleventh laws
  hold over grids whose formats hold codes.
- **The writer.** A code that is a built-in Excel draws alike in every
  locale is written by its id, as Excel writes it, and any other is defined
  in the styles part's `<numFmts>` from 164; into a workbook of one's own,
  a code the workbook defines takes its id and the rest take ids past its
  highest. The readers print no format yet, as they print no look at all.

## Amendment of 2026-10-18: oracle 12's lookups, each answered as its scan (`KERNEL.8`, its first slice)

The language computes the lookups and the error tests a lookup's miss is
read by: `INDEX`, `MATCH`, `XMATCH`, `VLOOKUP`, `HLOOKUP`, `LOOKUP` and
`XLOOKUP`; `IFERROR`, `IFNA`, `ISERROR` and `ISNA`; and `SUMIF` and
`COUNTIF`, moved down from `frazaro-core` with their meanings and their
tests, since a vector's index keys by a criterion's equality. Each is a row
of `excel::SUBSET` with its evidence, the slice's fixture, so the subset is
thirty-two functions. Each rule is Excel's as its documentation states it;
where the documentation is silent, the fixture's row and the owner's save of
it decide.

**The index.** A lookup or a criterion asked of one range again and again
answers from the range's index (`vla_lang::calc::Indexes`), built at its
second probe in one evaluation and kept until a cell of its range changes;
an approximate match takes a stated bisection over the vector as it stands,
and needs none. The scan stays the meaning: it answers every first probe,
every pattern and every criterion by size, and a test switch answers every
probe both ways and compares, over the generated grids with writes between
frames, so that an index a change should have reached fails by name. The
index moves no answer, only the time a probe takes.

**Labelled, not read.** Excel reads `SUMIF`'s sum range from its top-left
cell in the range's shape, and `LOOKUP`'s results however far past their
end: cells the formula's references do not name, which a plan orders after
nothing. A sum range smaller than its range, and results read past their
end, are labelled as written (`(not-computed "B1")`), where reading them
could read a formula not computed yet. A larger sum range is read in the
range's shape, as Excel reads it. `INDEX`'s pick of many cells where one
value is wanted is labelled as a range is, until implicit intersection
(the item's second slice).

**The fixture.** `scripts/recalc/lookups.txt`, built by the door into
`lookups.xlsx`, holds 165 formulas over three sheets: a Table of vectors
ascending, descending, of texts in two cases, of every kind with an empty
cell, an error and duplicates, and unsorted, on which the bisection's
candidate rules answer differently; a Lookups sheet of one formula per
function, mode, search and rule, `INDEX`'s every form, every error a
function answers, and the idioms (`LOOKUP(9.99E+307,...)`,
`IFERROR(VLOOKUP(...),0)`, `IF(ISNA(MATCH(...)),...)`); and a Criteria
sheet of `SUMIF` and `COUNTIF` over the index's equality and the scan's
sizes and patterns, an error in a sum range at a row met and at a row not
met, an empty cell as a criterion, and the two questions the equality turns
on: a letter's case beyond ASCII (`=F1=F2` over `É` and `é`) and two
numbers equal to fifteen digits and not to the bit (`=0.1+0.2=0.3`, and
`MATCH` of the one in a cell holding the other). `lookups_calc.vla` is the
door's record of it: 161 rows `agree` with the values the build folded and
four labelled holding none, pinned `built`. The owner's Excel 365 save of
it, `lookups_saved.xlsx`, joins the table at `no-differ`, its agreements a
floor; where Excel's save answers otherwise, the rule changes, in the
lookups alone or in every comparison at once.

**The subset's fixture.** `VLOOKUP(10,A1:B2,2)` and the cell reading it
compute, so `subset.xlsx`, built again, agrees in 104 rows, and the owner's
save of it, `subset_saved.xlsx`, agrees in the same 104, the build's
prediction; `CONCATENATE` and the range where one value is wanted stay
named.

## Amendment of 2026-10-19: oracle 12 held to Excel's save of the lookups (`KERNEL.8`, its first slice)

The owner's Excel 365 save of the lookups' fixture,
`scripts/recalc/lookups_saved.xlsx`, joins the table at `no-differ`:
`lookups_saved_calc.vla` holds its 165 rows, 160 `agree` and five
labelled, its census beside the fixture's and its calc row `prove.ps1`'s
ninth. It answered otherwise than the build in twelve rows, and seven rules
changed to its answers:

- `MATCH`'s type -1 walks the vector from its first cell, keeping each cell
  of the value's kind not under it and stopping at the first under it,
  where the build bisected (`Lookups!A26`: 1 over an unsorted vector, where
  every bisection keeps a later cell). Type 1 and every other approximate
  match keep the bisection, which the save agreed with over the same
  unsorted vectors.
- `XMATCH`'s and `XLOOKUP`'s match mode 3, Excel's regular expressions, is
  taken and labelled `(not-computed "match mode 3")`, where the build
  answered `#VALUE!` (`A46`).
- A criterion's `<>` of a number holds of a text spelling the number
  (`"<>20"` counts `"20"`, `Criteria!H4` and `H20`), while `=` counts it
  equal (`H1` to `H3`): a text spelling a number meets `=` alone.
- An empty cell given as a criterion is the number 0, counting the zeros
  and not the blanks (`H16`).
- An error in `SUMIF`'s sum range answers only at a row the criterion meets
  (`H23`; `H24`, at a row met, answers it).
- Two texts compare without case beyond A to Z too, by Unicode's simple
  lowercase mapping: `="É"="é"` is TRUE (`H29`), and so in `COUNTIF`,
  `SUMIF` and the lookups (`H34` to `H36`); an accent is no case.
- Two numbers equal to fifteen significant digits are equal to the
  comparison operators and to the criteria, `=0.1+0.2=0.3` TRUE and
  `COUNTIF` of the one against the other 1 (`H30`, `H32`), while a lookup
  matches a number to the bit, as the save's `MATCH` and `VLOOKUP` of the
  one in the other answer `#N/A` (`H31`, `H33`).

The build was made again, and its rows agree with the values it folds: 160,
since `XMATCH`'s mode 3 is labelled now, holding none, so the agree floor of
`lookups_calc.vla` moved from 161 to 160 with that reason written. Five
rows are labelled in both files, `LOOKUP`'s results read past their end,
`INDEX`'s pick of many cells twice and `SUMIF`'s smaller sum range; Excel's
values for them (`"r8"`, 20, `#VALUE!`, 11) wait for the slices that compute
them. The identity of a name is untouched: the language's `fold`, which
decides that two names are one, stays ASCII's alone.

## Amendment of 2026-10-19: oracle 12's implicit intersection and a formula's kind (`KERNEL.8`, its second slice)

A legacy formula reads a range where one value is wanted at its own cell,
Excel's implicit intersection: an area one column wide at the formula's row,
one a row high at its column, any other at both, on any sheet; `#VALUE!`
where the area holds no such cell, and `#REF!` for a reference spanning
sheets, as Excel's save holds them. A dynamic-array formula's range where
one value is wanted is an array, labelled as written until the item's
arrays. `SINGLE`, Excel's `@` as the file spells it (`_xlfn.SINGLE`, without
which prefix Excel reads a function of no add-in, `#NAME?`), meets the range
at the cell in a formula of either kind and answers a reference to the cell
met, a span `#REF!`.

The kind comes from the file. The xlsx reader tells an array formula's
anchor (`t="array"`) dynamic where its `cm` names the dynamic-array
metadata (`XLDAPR`) and legacy where it does not, and the OpenDocument
reader tells a matrix formula, legacy, its format telling no dynamic one
from it; each tells it to the sink beside the formula row, and the printer
ignores it, so that oracle 8's records are unchanged. `frazaro calc` loads
an array formula of either kind as a dynamic one: a legacy array formula
over single values computes as Excel's array evaluation computes it.

**The fixture.** `scripts/recalc/intersect.txt`, built by the door into
`intersect.xlsx`, holds 31 formulas over a Data sheet's column, row and
block: 21 written legacy, each met inside its area and past it, through an
operator, a condition, a sign, `&`, a whole column, a whole row, a span of
sheets alone and through `SINGLE`, and `SINGLE`; and 10 written dynamic by
the functions they call, their arrays labelled and `SINGLE` meeting them all
the same. `intersect_calc.vla` is the door's record of it: 28 rows `agree`
with the values the build folded and three labelled holding none, pinned
`built`. Excel 365's first save of it, of a build older than the writer's
prefix, read `SINGLE` as a function of no add-in (`_xludf.SINGLE`, `#NAME?`)
in its eight cells and the span as `#REF!` where the door had answered
`#VALUE!`, its other eighteen rows agreeing. Built again, its second save,
`intersect_saved.xlsx`, joins the table at `no-differ`: 31 rows, 28 agreeing
and the three arrays unchecked beside the values Excel spilled. Excel writes
`@` bare where a legacy formula's top holds it (`C1`, `C10`, `B6`), the
value the same, and keeps `_xlfn.SINGLE` under `ROW` and in the dynamic
formulas.

**The subset's fixture.** `IF(A1>5,A1:A2,0)`, met at row 27 outside
`A1:A2`, is `#VALUE!`, as Excel's save holds it, so `subset.xlsx`, built
again, and `subset_saved.xlsx` agree in 105 rows each.

**An engine's load and its twins.** A defined name whose text the parse
does not read, read by a formula or by a name it reads, is refused at an
engine's load with the constructs (`calc-construct-not-read`); and a
`.last` twin holding a reason beside its cells answers it to a range over
the twin as its sheet answers a range over itself, the twin offering no
array to a walk while the reason stands.

## Amendment of 2026-10-19: oracle 12's text functions and oracle 11's dates (`KERNEL.8`, its third slice)

The language computes twenty-five text functions, `CONCATENATE`, `CONCAT`,
`TEXTJOIN`, `LEFT`, `RIGHT`, `MID`, `LEN`, `FIND`, `SEARCH`, `SUBSTITUTE`,
`REPLACE`, `UPPER`, `LOWER`, `PROPER`, `TRIM`, `REPT`, `EXACT`, `VALUE`,
`T`, `N`, `UNICODE`, `UNICHAR`, `CHAR`, `CODE` and `TEXT`, each as Excel's
documentation states it and as Excel 365's save of the fixture answers. A
text is counted by its characters, as Excel counts it, an emoji one; a
count, a start or an instance is read down to the whole number at or below
it, so that `LEFT(text,-0.5)` is `#VALUE!`. Excel's limit on a text, 32,767,
is in UTF-16 units, an emoji two: a text `REPT`, `SUBSTITUTE`, `REPLACE`,
`CONCAT`, `TEXTJOIN` or `TEXT` would make past it is `#VALUE!`, and one `&`
or `CONCATENATE` makes is cut to it, labelled where the cut would halve a
surrogate pair, a text of the language holding no half. `UPPER` and
`LOWER` map a letter by Unicode's simple mapping, one character for one,
`LOWER("İ")` an `i` and `UPPER("ß")` a `ß`. `CHAR` and `CODE` compute over
ASCII and label a code page's 128 to 255; `VALUE` labels a text holding a
digit that no operator reads, which a host reads by its locale; `TEXT`
draws by a number format's code and labels a code this version does not
draw. `FIXED` and `DOLLAR`, a locale's separators and currency, are not
computed.

**The order of texts.** Two texts compare as Excel's save ranks them: by
their characters' places (the space, the punctuation and the currency, the
other signs, the digits, then the letters, each with its accented forms,
then the save's five of other scripts), then by their accents (none, acute,
grave, circumflex, diaeresis, ring), case aside, and last by their
apostrophes, hyphens and em dashes, passed by until then, a text without one
first and of two with one at different places the later first; `ß` is `ss`
and `æ` is `ae`. The operators, a lookup's search, an index's keys and a
criterion all take it. The table of places holds the characters the save
ranks and what its levels give: ASCII's printable characters, nine signs
beyond it, the em dash, the letters with an acute, a grave, a circumflex or
a diaeresis, `å`, `ç`, `ñ`, `æ`, `ß`, `α`, `ω`, `ж`, `😀` and `日`, each
letter in both its cases. A comparison by size that a character beyond the
table decides, its place unknown, is labelled `(not-computed "a character's
order")`; whether two texts are one is answered always, by the fold beyond
the table.

**Oracle 11's dates.** A number format's code may hold a date's and a
time's parts: the year, the month by number or English name, the day or
the weekday's English name, the hour, the minute, the second and its
decimals, an elapsed time and a twelve-hour clock, drawn over Excel's 1900
calendar, its 29 February 1900 among its days; a negative number and one
past 31 December 9999 draw no date. `scripts/formats.txt` holds nine cases
of them beside its first date's, which was refused and draws now, and
`scripts/formats_golden.txt` the text Excel draws for each, written ahead
as the prediction and written again by the reference's
`VlaWriteFormatsGolden` unchanged.

**The fixture.** `scripts/recalc/texts.txt`, built by the door into
`texts.xlsx`, holds 357 formulas: a Functions sheet of 192, one a function
and a rule, `TEXT`'s dates among them, and the comparisons the order of
texts answers; and an Order sheet of 163 characters and texts, each ranked
among them by `SUMPRODUCT`, which this version does not compute, so that
Excel's save holds its own order of texts for the language to be held to.
`texts_calc.vla` is the door's record of it: 190 rows `agree` with the
values the build folded, an empty text among them, and 167 labelled, pinned
`built`. The owner's Excel 365 save of it, `texts_saved.xlsx`, joins the
table at `no-differ`, its 190 computed rows agreeing with Excel's and its
Order sheet's 163 ranks the order the language's own tests hold rank for
rank. The save differed in 17 rows until the language counted characters,
cut `&` and `CONCATENATE` at the limit, read a count down, lowered `İ` and
ordered texts as Excel does.

**A formula's empty text.** The xlsx reader reads `t="str"` and an empty
`<v>`, as a host saves a formula's `""`, as the empty text, where it read
nothing; any other kind's empty `<v>` still holds nothing.

**The subset's fixture.** `CONCATENATE("a","b")` computes `"ab"`, as
Excel's save holds it, so that `subset.xlsx`, built again, and
`subset_saved.xlsx` agree in 106 rows each.

## Amendment of 2026-10-20: oracle 12's date functions (`KERNEL.8`, its fourth slice)

The language computes seventeen date functions, `DATE`, `YEAR`, `MONTH`,
`DAY`, `WEEKDAY`, `EOMONTH`, `EDATE`, `DAYS`, `NETWORKDAYS`, `WORKDAY`,
`TIME`, `HOUR`, `MINUTE`, `SECOND`, `DATEDIF`, `YEARFRAC` and `WEEKNUM`,
each as Excel's documentation states it and, where it is silent, as Excel
365's own save answers. A date is a serial of Excel's 1900 calendar: 1 is 1
January 1900, 60 the 29 February 1900 that never was, which Excel keeps, 0
the day before the first, and 2,958,465 31 December 9999; a serial below 0
or past the last day is `#NUM!`, as is a date `DATE`, `EOMONTH` or `EDATE`
would make outside them, or a month they reach before 1900 or after 9999.
February 1900 has 28 days to `EOMONTH`, `EDATE` and `YEARFRAC`, and 1900 is
no leap year to `YEARFRAC`. A date's parts, its weekday among them, and a
time's are read from the serial rounded to the second, as `numfmt` draws it,
so that a time a hair before midnight is the next day's. A whole number an
argument asks for is read as Excel's save reads it: `DATE`'s and `TIME`'s
parts, `WORKDAY`'s days, a type and a basis by their floor, `EOMONTH`'s and
`EDATE`'s months truncated toward zero. `NETWORKDAYS` and `WORKDAY` pass by
Saturdays, Sundays and their holidays, each holiday counted once.
`DATEDIF`'s `"YD"` walks the months past the years from the start's own
month in its own year and adds the days past the months, as Excel's saves
reckon it. Two edges are labelled, where two of Excel's saves of one formula
answered differently: a year outside 0 to 9999 that `DATE`'s months carry
back into it, `(not-computed "a year outside the calendar")`, and
`WEEKNUM`'s types past 21, which its documentation does not name,
`(not-computed "WEEKNUM's undocumented type")`.

**A locale's number.** An argument of a date function, an operand of `+`,
`-`, `*`, `/`, `^`, `%` and the unary minus, and `VALUE`'s, read as a
number: a text spelling one by the language's rule its number, a text
holding a digit that rule does not read, `"1,000"` or `"1/1/2026"`, labelled
`(not-computed "a locale's number")`, since a host reads it by its locale,
and any other text `#VALUE!`. The operators answered `#VALUE!` for such a
text before. `DATEVALUE`, `TIMEVALUE`, `TODAY` and `NOW` are not computed,
named.

**The 1904 system.** The xlsx reader reads `workbookPr`'s `date1904`, a
Mac's workbook counting its dates from 1904, and a door computing it labels
every function whose answer turns on the calendar, `(not-computed "the 1904
date system")`, and `TEXT` of a date's code, while `DAYS`, `TIME`, `HOUR`,
`MINUTE` and `SECOND` compute; a workbook without the attribute counts from
1900.

**The fixture.** `scripts/recalc/dates.txt`, built by the door into
`dates.xlsx`, holds 1,377 formulas: a Functions sheet of 364, one a function
and a rule, the edges Excel's documentation leaves unsaid among them (a
fraction where a whole number is asked, a time a hair before midnight, the
1900 calendar's first days, `DATEDIF`'s `"MD"`, `YEARFRAC`'s bases), and
arithmetic over the texts a locale reads; and a Calendar sheet of every year
from 1900 to 2101, its first day's serial, February's length, the first
day's weekday and ISO week, the day before's date and the last day's ISO
week, from which with the months' fixed lengths every day's date and weekday
follow, as a test of the core holds every day of them to Excel's save.
Excel 365's first save of the first 297 differed in 25 rows, each folded to
its answer; the 51 after them are the edges the fold chose, for its second.
The second differed in 13: two of the 297 answered otherwise than in the
first, whose edges are labelled, and the rest folded; the 11 after the 348
are `DATEDIF`'s `"YD"` and the two edges again, for its third. The third
differed in one, `"YD"` from 22 June 1988 to 11 May 2012, 323, whose rule
it corrected; the 5 after the 359 are where that rule parts from the
anniversary's, and its fourth agreed with each. `dates_calc.vla` is the
door's record of it: 1,329 rows `agree` with the values the build folded
and 48 labelled, pinned `built`; the owner's fourth save of it,
`dates_saved.xlsx`, joins the table at `no-differ` as
`dates_saved_calc.vla`, its 1,329 computed agreeing with Excel's.

## Amendment of 2026-10-20: oracle 12's aggregates by criteria (`KERNEL.8`, its fifth slice)

The language computes thirty-one functions more, each as Excel's
documentation states it and, where it is silent, as Excel 365's own save
answers: `SUMIFS`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS` and
`MINIFS`; `SUBTOTAL`; the database functions `DSUM`, `DAVERAGE`, `DCOUNT`,
`DCOUNTA`, `DGET`, `DMAX`, `DMIN`, `DPRODUCT`, `DSTDEV`, `DSTDEVP`, `DVAR`
and `DVARP`; and the aggregates `SUBTOTAL` names, `AVERAGE`, `COUNT`,
`COUNTA`, `PRODUCT`, `STDEV`, `STDEVP`, `VAR` and `VARP`, with `STDEV.S`,
`STDEV.P`, `VAR.S` and `VAR.P`. A criterion reads as `COUNTIF`'s (the
amendment of 2026-10-18), and the ranges of `SUMIFS` and its family are of
one shape, `#VALUE!` otherwise. A criterion's text with no `*` or `?` is
compared as written, its tildes too: Excel's save met no `~tilde` with
`"~~tilde"`, where its documentation reads `~~` as one `~`. A lookup's
reads every `~` as that documentation does: its second save found
`~tilde` with `MATCH` of `"~~tilde"` and nothing with `"~tilde"`.
A database's criteria are rows of tests under its columns' names, found
without case: a record meets a row whose every test it meets, and the
criteria where it meets any row, a blank row meeting every record; a text
with no operator is a record's text beginning with it, as Excel's filters
read one; a criteria range of one row, its names alone, is `#VALUE!`. A
database's rows end with its sheet's last cell, so that a whole column's
records are the rows its sheet holds anything in, and `DGET` of an empty
cell is `#VALUE!`. A sum and a mean add in the order a range is walked,
from -0.0, as `SUM`'s do, and a variance is taken in two passes, the mean
first.

**A hidden row.** The xlsx reader reads a row's `hidden`, and the
OpenDocument reader a row's `table:visibility` of `collapse` or `filter`,
as one flag. `SUBTOTAL`'s 101 to 111 pass every hidden row by. Its 1 to 11
count a row hidden by hand and pass by a row a filter hides, which an xlsx
file tells apart only by its filter's range, which the reader does not
read: where a hidden row holds a cell of a reference they are labelled
`(not-computed "a row a filter may hide")`, and where none does they
compute. Every `SUBTOTAL` passes by a cell whose formula calls `SUBTOTAL`.

**Labelled.** A database's test under a name no column of it holds: a
formula there, which Excel computes record by record, `(not-computed "a
computed criterion")`, and a value there `(not-computed "a criterion of no
field")`; and an `AVERAGEIF` whose average range is smaller than its range,
the range named, as `SUMIF`'s is.

**The tolerance, amended.** Two numbers agree when their fifteen significant
digits, as each spells them, are at most one unit of the fifteenth apart,
the unit the larger spelling's; until this amendment they had to round to
the same fifteen. Two computations of one quantity a few units in the last
place apart then agree whichever side of a rounding boundary each falls on:
Excel's save gave `VARP` of the amounts 9174.2080555555549 where the
language's two passes give 9174.208055555557, 0.21 of a unit apart and
spelled ...555 and ...556, and `VAR(0.1,0.2,0.3)` 0.009999999999999988 where
they give 0.009999999999999998, spelled 0.00999999999999999 and
0.0100000000000000, a tenth of the larger's unit apart. Past the fifteen
digits Excel documents, its variance family follows no one formulation its
save shows. Every earlier row of the table agrees under both readings.

**The fixture.** `scripts/recalc/criteria.txt`, built by the door into
`criteria.xlsx`, holds 228 formulas: a Functions sheet of 219, one a
function and a rule, the edges Excel's documentation leaves unsaid among
them, over a Data sheet of twelve records, its criteria ranges and a
`SUBTOTAL` among the cells it sums, and a Far sheet whose cells reach below
its database's. `criteria_calc.vla` is the door's record of it: 225 rows
`agree` with the values the build wrote and 3 labelled, pinned `built`.
The owner's save hides the Data sheet's row 6 before Excel computes it, so
that `SUBTOTAL`'s 101 to 111 pass a row by and its 1 to 11 are labelled
over it. Excel's first save of it differed in seven rows, four of them the
rules above that it settled, three the variance's last digits; its second,
of the rules' reach, in two, `MATCH`'s tilde, which keeps the documented
reading. The owner's second save, `criteria_saved.xlsx`, joins the table
at `no-differ` as `criteria_saved_calc.vla`, its 213 computed agreeing
with Excel's, 205 to the bit and the variance family's 8 within one unit
of the fifteenth digit, and the 15 labelled.

## Amendment of 2026-10-20: oracle 12's arrays within a formula (`KERNEL.8`, its sixth slice)

**An array** is rows by columns of values that a formula computes and no
cell holds: an array constant, `{1,2;3,4}`, `,` between the values of a row
and `;` between rows, each a number, a text, a truth value or an error; a
range read whole, every place of it, an empty cell an empty value; and a
function's answer. An operator, and a function of one value given an array,
compute at each of its values, Excel's lift: the answer has the most rows
and the most columns of its operands, an operand of one row standing at
every row and one of one column at every column, and `#N/A` at a place one
of them does not reach. `SUM`, the aggregates, the lookups' vectors and
every function reading a range read an array as a range; `COUNTIF`'s and
the database functions' ranges are references, an array there `#VALUE!`.

**A formula's kind and its parameters' classes.** A dynamic-array formula
reads a range where one value is wanted whole, as an array; a legacy one
meets it at its cell, implicit intersection (the amendment of 2026-10-19),
but inside a parameter Excel reads as an array, `SUMPRODUCT`'s, `MMULT`'s,
`TRANSPOSE`'s, `ROWS`' and `COLUMNS`', `INDEX`'s reference, a lookup's
vector and table, `XLOOKUP`'s arrays and every array function's array, the
argument's expression is read as arrays in a formula of either kind. A
parameter of one value lifts its function over an array given, and over a
range where the formula reads arrays: `LEN(B1:B6)` is six lengths there, and
`COUNTIF(A1:A9,{1,3})` two counts in any formula. A parameter of one value
forced, `IFERROR`'s and `IFNA`'s, and every one of `VLOOKUP`'s and
`HLOOKUP`'s but the table, is one value in a legacy formula: its expression
is computed as outside every array parameter, a range met at the cell, and
an array given is its first value, as Excel's own saves compute
`IFERROR({1,#N/A,3},0)` as 1, `VLOOKUP(3,A1:C6,{1,3},FALSE)` as its first
column and `VLOOKUP({3,4},A1:C6,3,FALSE)` as the row of 3, where
`MATCH({4,9},A1:A6,0)` lifts; a dynamic formula lifts the function over it.
`IF`'s condition and `CHOOSE`'s index keep their own class in a legacy
formula: an array lifts the function, and a range is met at the cell even
inside an array parameter, as `SUMPRODUCT(IF(A1:A9>2,1,0))` asks for
Ctrl+Shift+Enter in Excel. `ROW` and `COLUMN` of a range read as arrays are
its rows' and columns' numbers. A legacy array formula, entered with
Ctrl+Shift+Enter, reads as arrays as a dynamic one does.

**A cell's answer.** An array of one value is that value. An answer of many
is, in a legacy formula and a legacy array formula, its first value, as
Excel's `@` takes the top-left value of an array; in a dynamic formula a
spill, labelled `(not-computed "a spill")` at its cell until spills come, as
a dynamic formula's range answered is labelled by its text (the amendment of
2026-10-19). An array with nothing in it is `#CALC!`, the eighth error
value, which no formula writes as a literal. The file holds it only as a
rich value, `#VALUE!` beside a `vm` naming an `_error` of `errorType` 13
([MS-XLSX] 2.3.6.1.3), `#SPILL!` its 8, and the reader reads the error the
rich value names; the build writes a `#CALC!` answer with no cached value,
since Excel repairs a bare one. An array of more than 1,048,576 values, the
grid's own limit, is labelled `(not-computed "an array past the grid's
limit")`; Excel's own save answers `SEQUENCE(2000000)` `#VALUE!`, which its
documentation does not state, and the label stands at the owner's decision
of 2026-10-21, since that answer is of a sequence past the sheet's rows, and
an array past the limit within them may compute in Excel. An array a
function wanting one value meets where the formula reads arrays is
`(not-computed "an array where one value is wanted")`.

**A correction to the amendment of 2026-10-19.** `TEXTJOIN` and `CONCAT`
past the text limit are `#CALC!`, not the `#VALUE!` their documentation and
that amendment state: Excel's own save of the text fixture holds its `A142`,
`LEN(TEXTJOIN(",",FALSE,data!B:B))`, as a rich `#CALC!` whose fallback
`#VALUE!` the reader took before it read rich values, and its second save of
the arrays' fixture holds `LEN(CONCAT(REPT("x",20000),REPT("y",20000)))`
(`A214`) the same. `REPT`, `SUBSTITUTE`, `REPLACE` and `TEXT` past it stay
`#VALUE!`, and `&` and `CONCATENATE` cut their text to it.

**The functions.** `SEQUENCE`, `FILTER`, `SORT`, `SORTBY`, `UNIQUE`,
`TRANSPOSE`, `SUMPRODUCT`, `MMULT`, `ROWS` and `COLUMNS`, and `CHOOSECOLS`,
`CHOOSEROWS`, `TAKE`, `DROP`, `HSTACK`, `VSTACK`, `TOCOL`, `TOROW`,
`WRAPROWS`, `WRAPCOLS` and `EXPAND`, each as Excel documents it: `FILTER`
keeping no row is its third argument or `#CALC!`; `SORT` and `SORTBY` order
numbers, then texts in Excel's order of texts, then FALSE and TRUE, then the
errors, an empty value last either way, rows of one order keeping theirs;
`UNIQUE` tells two texts one without case; `SUMPRODUCT` counts a value that
is no number as 0 and adds place by place in rows from 0, `MMULT` along each
row from its first column (`Alonzo/docs/ALGEBRA.md` section 7); the shapers
pad with `#N/A`. `IF`, `CHOOSE`, `IFERROR` and `IFNA` lifted take at each
place the value of the operand they choose there, `#N/A` where that one
falls short, an operand not chosen free to: Excel's own save computes
`IF({FALSE,FALSE,FALSE},{1,2},5)` as three 5s. `COUNTA` counts an array's
empty value, where it passes a range's empty cell by, as Excel's save
counts `COUNTA(TRANSPOSE(B1:B6))` over an empty `B6`.

**`:` between expressions.** `A1:INDEX(A:A,5)`, `INDEX(A1:A9,2):A9` and
`IF(c,A1,B1):C5` are the rectangle spanning the two references their ends
answer, on their one sheet; an end that answers no reference is `#VALUE!`.
The plan's edges for such a range are its reach, the rectangle spanning
every reference either end may answer (a reference's own, `INDEX`'s its
reference's, `IF`'s and `CHOOSE`'s their branches'); a range whose end is a
defined name, which no plan bounds, is labelled `(not-computed "a range
between computed ends")`.

**The goldens it moves.** The reflect fixture's legacy array formula,
`=B1:B2*2` (`fixture.xlsx`'s `E3`), and the OpenDocument fixture's matrix
formula, the same `E3`, compute their first value, 2400, and agree; the
reflect fixture's dynamic `=SEQUENCE(2)` is labelled a spill, and the
OpenDocument fixture's, a plain formula there, computes its first value
and agrees. The intersection fixture's dynamic `=CHOOSE(1,Data!B1:F1)+0` is
labelled a spill where its range was. The text fixture's 163 `SUMPRODUCT`
ranks of its Order sheet, labelled `SUMPRODUCT` before, compute and compare
texts the save did not rank, labelled `a character's order`.

**The fixture.** `scripts/recalc/arrays.txt`, built by the door into
`arrays.xlsx`, holds 210 formulas: a Functions sheet of 205 from row 10, one
a rule, array constants and the lift in legacy formulas, `SUMPRODUCT`,
`ROWS` and `COLUMNS`, the array functions, a range read whole in a dynamic
formula, the lookups over arrays and `:` between expressions, the edges
Excel's documentation leaves unsaid among them, and two spills with room
beside them, over a Data sheet; the build writes a formula calling an array
function as a dynamic one, as it writes every such formula.
`arrays_calc.vla` is the door's record of it: 198 rows `agree` with the
values the build wrote, 9 `#CALC!` answers it holds no value for, and 3
labelled, the two spills and the grid's limit. Excel's first save of it, on
2026-10-21, agreed at every cell the door computes once the classes, the
lifts and the reading above were folded to it; its second, the same day,
with three legacy array formulas typed in and 15 rows of the folds' reach
(`A200` to `A214`), agreed too once the value `VLOOKUP` seeks was forced and
`CONCAT` past the limit was `#CALC!`. That save, `arrays_saved.xlsx`, joins
the table at `no-differ` (`arrays_saved_calc.vla`): 213 formulas, the 210
the door computes agreeing with Excel's, the legacy array formulas among
them, and 3 labelled, the two spills and the grid's limit.
