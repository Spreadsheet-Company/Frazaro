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
