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

**The data both implementations read grows.** `scripts/words.vla`
(`VLA_English.bas`'s nine word tables) and `scripts/names.vla`
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
