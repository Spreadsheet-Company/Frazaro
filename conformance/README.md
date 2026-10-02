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
