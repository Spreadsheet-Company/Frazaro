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
