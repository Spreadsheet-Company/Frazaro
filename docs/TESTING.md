# Testing checklist

The standard verification path for changes to VLA/Frazaro's grammar,
phrasebook (`english.vla`), or runtime (`VLA_Runtime.bas` and friends).
Six passes, in order - each depends on the previous one passing clean,
and not every pass applies to every change (see each pass's own "when
this applies" note). Where a step says "confirm your new ...", fill in
the specific rule/pin/check the current pass actually added or touched.

## The lab bench, from nothing (once per clone)

`VLA.xlsm`, the development workbook every pass below runs in, is not
tracked (since the `0.5.0` public import; see `.gitignore` for why). It is
nothing but a blank workbook with the modules imported, so a fresh clone
rebuilds it in three steps:

1. In Excel, create a new blank workbook and save it as `VLA.xlsm`
   (macro-enabled) in the repository root, beside `src/`. The reload rig
   finds `src` by that adjacency.
2. Once, in File → Options → Trust Center → Trust Center Settings → Macro
   Settings, check **"Trust access to the VBA project object model"** —
   the reload rig imports modules through the VBProject, which needs it
   (the shipped add-in's Interpret path does not; this is a developer
   setting only).
3. Alt+F11, File → Import File…, import `src\VLA_DevRig.bas` by hand (a
   module cannot import itself), then in the Immediate window run
   `VlaDevReload`. It imports every other module from `src/` — standard
   modules, the classes, and `frmCLI` — then **Debug → Compile VBAProject**.

That is the whole lab bench. Everything the passes below need beyond code
(`Output`, `Demo`, `GStruct`, `GFormat`, the `VLA_Log` sheet) is created by the runs
themselves. Building the distributable add-ins from it is `DEPLOY.md`'s
"Building the add-in" section.

## Pass 0 — Reload changed source

Immediate window: `VlaDevReload` (no arguments needed - it defaults to
the `src` folder beside the workbook). One call re-imports every module
in its own list (`VLA_DevRig.bas`'s `mods` array) from disk, whether or
not each one actually changed - safe to run every time rather than
tracking which files you touched. It cannot reload `VLA_DevRig.bas`
itself (a module can't replace itself while running) - update that one
by hand if it's what changed.

`VlaDevReload` re-imports the code but does not recompile it - after it
prints its "N module(s) refreshed" summary, do **Debug > Compile
VBAProject** in the VBE before running anything else, same as the
tool's own closing note says. Skipping this step means you're about to
test against a mix of old compiled code and new uncompiled source.

`.vla` vocabulary files and `.txt` scripts (`english.vla`,
`instructions.txt`) are read from disk at runtime, not imported - no
action needed there beyond them being in place at the expected path.

## Pass 1 — VlaSelfTests (static, safe on any workbook state)

*Always applies.*

1. Immediate window: `? VlaSelfTests()` (the leading `?` is not
   required - a bare `VlaSelfTests` runs it identically, same
   `Debug.Print` output either way - it's there because the function
   returns a `Boolean`, and `?` prints that final True/False verdict as
   its own line after the summary, so you don't have to read and parse
   the "pure PASS, host PASS" text yourself to know pass/fail.)
2. Expect: `pure PASS (N/N), host PASS (N/N)` on the summary line, then
   `True`. AS7.2: if either half failed, its full failure list (name +
   detail, not just the count) now reprints again right below this
   line too - no need to scroll back up to the point in the run where
   that half's own tail already printed it once.
3. Confirm any regression pin(s) added for this change appear and read
   `PASS` (a new pin proves the fix; its absence from the run means it
   isn't wired into `VlaSelfTest`/`VlaSelfTestHost`'s dispatch list yet).
4. Confirm pins for anything this change's edit could plausibly have
   disturbed still read `PASS` - especially a shared dispatch block
   (a `Select Case`, a shared helper) that gained a new branch, since
   that's exactly the shape of edit that can silently break a sibling
   case.

If anything fails, stop here - don't proceed to later passes until this
is clean.

## Pass 2 — retired

Was `EnglishFormPathSelfCheck`: loaded a vocabulary twice, once with
F.2's form path on and once off, and confirmed the `test:`/`fail:`
corpus agreed both ways. Retired along with the toggle itself and the
old `Replace()`-text fallback it was proving safe - there is only one
path now (F.2's own closure bar is met; see `BETA_ROADMAP.md`'s F.2
entry), so there is nothing left to compare against. Ordinary
vocabulary loading already re-verifies every `test:`/`fail:` proof
through that one path on every load, which is what this pass's own
comparison reduces to now - no separate step needed. (Steps 5-7
retired with it; Pass 3 below still starts at step 8, not renumbered,
so cross-references elsewhere in this file stay valid.)

## Pass 3 — VlaGoldens (static, regenerates + diffs, no live workbook needed)

*Applies whenever `english.vla`, `instructions.txt`, or anything in the
compile/emit pipeline (`VLA.bas`, `VLA_English.bas`, a `defmacro`, a new
`VLA_Runtime.bas` helper) changed - i.e. almost every pass.* Distinct
from Pass 1's `TestGoldens`, which only proves translation/transpilation
are deterministic - it never touches the golden files on disk. This
pass is what actually regenerates and inspects them.

8. Immediate window: `? VlaGoldens()` - runs `VlaWriteGoldens` (rewrites
   `instructions_golden.vla`/`instructions_golden.vba` beside
   `instructions.txt`) then `VlaWriteInterpreterGoldens` (rewrites
   `scripts/interpreter_golden.txt`), one call. Needs no live workbook -
   pure transpile/emit over files already on disk.
9. Expect: `True`, no `FAILED` lines.
10. `git diff` the three golden files. This diff is the total-pipeline
    witness, in the tool's own words: **empty means behavior preserved;
    any change is a contract change and must be read, not skimmed.**
    Confirm every changed line is something this change actually
    intended (the new rule's emitted VBA, a changed step number from an
    earlier insertion, etc.) - an unexpected line anywhere else in the
    diff means something this change wasn't supposed to touch, changed.
11. Commit the regenerated golden files alongside the source change -
    an uncommitted golden diff left behind means the next person's
    `git diff` shows noise that isn't theirs.

## Pass 4 — Live Excel run (both backends, in the dev workbook)

*Applies whenever the change touches the object model - a new or
changed `defmacro`, a new Tier-2 `VLA_Runtime.bas` helper, or anything
`instructions.txt`/`VerifyReportChecks` was extended to exercise.*

**Precondition:** if this change added anything to `instructions.txt`
that creates a NAMED object (a table, pivot table, sheet, defined name,
etc.) that didn't exist in the file before, use a fresh workbook, or
delete/clear the relevant sheets from a prior run first. Rerunning over
leftover state will collide (Excel refuses to recreate a same-named
table/pivot/sheet). The interpreter side is self-cleaning
(`VerifyReportInterpreter` deletes-and-recreates the `Output` sheet
itself); the compile side is not.

12. Trigger the compile-backend run (ribbon "Run", or Immediate window
    `EnglishIdeRun`).
13. Expect: no error dialog - `instructions.txt` compiles and runs clean.
14. Immediate window: `? VerifyReport()`
15. Expect: `0 failed`, then `True` (same Boolean-return convention as
    Pass 1), including every check added or changed for this pass
    reading `PASS`.
16. Immediate window: `? VerifyReportInterpreter()`
17. Expect: same - `0 failed`, then `True`, same checks passing, run
    independently through the interpreter backend (it compiles and
    executes `instructions.txt` itself - no dependency on step 12's run).

Shortcut for 14-17: `? VerifyReports()` runs `VerifyReport` then
`VerifyReportInterpreter` in that order - but step 12 still has to
happen first for `VerifyReport`'s own half. AS7.2: the combined summary
line now shows counts for both (`emitter PASS (N/N), interpreter PASS
(N/N)`), and either side's failure list reprints again right below it
if that side failed.

## Pass 5 — Optional visual spot-check

*Worth doing whenever a change exercises a genuinely new kind of COM
interaction - something no existing helper or macro has exercised
before - since an existence/property check can miss a result that's
structurally wrong in a way the check didn't think to ask about.*

18. Open the relevant sheet(s) and eyeball the objects this change
    created or modified - does it actually look like what the English
    sentence said it would do, not just "did a check pass."

Not required for correctness when Pass 4 already asserts the resulting
state programmatically - this is a sanity net for the specific class of
bug that a same-shape existing check wouldn't catch on a first-of-its-
kind interaction.

## Pass 6 — VlaBuildAddin clickthrough (the shipped artifact, a fresh workbook, no dev-only modules)

*Applies whenever the shipped module list changed (`VLA_Build.bas`'s own
`mods` array), a shipped module now references something new by name, or
before treating a change as release-ready.* Every earlier pass runs
inside the dev workbook, which carries every dev-only module too - it
cannot catch a shipped module calling something that isn't actually
shipped. That exact failure has hit this project multiple times
(`VLA_Build.bas`'s own F5.0/IN.6/IN.7 history notes): a fresh-built
add-in failing **Debug > Compile** with "Variable not defined," only
visible once something outside the dev workbook actually loads it.

19. In the dev workbook's Immediate window: `VlaBuildAddin`.
20. Expect the "Add-in built" message box, naming `Frazaro.xlam` next to
    the dev workbook, with no phrasebook-audit or ribbon-injection
    failure noted in it.
21. Open a genuinely new, blank Excel workbook - not the dev workbook,
    no prior `VlaDevReload` state.
22. Load `Frazaro.xlam` as an add-in for that session (open the file
    directly, or install it via Add-ins) so the **Frazaro** ribbon tab
    appears.
23. Click **Set Up Workspace** (Program group). Expect a workspace sheet
    to appear and activate, no error.
24. Type a trivial one-line program into it, e.g. `Put 1 into cell A1.`
25. Click **Compile Instructions** (Compiler group). Expect no VBA
    runtime-error dialog and cell A1 to actually hold `1` - proves the
    compile backend works standalone, with only the shipped module list.
26. Click **Interpret Instructions** (Interpreter group) on the same
    program. Expect the same - no crash, same resulting value - proving
    the interpreter backend works standalone too.

This pass is intentionally not about full corpus correctness (Pass 4
already covers that in depth) - it is specifically checking "does the
thing a real user would install even survive being opened," which nothing
else in this checklist can catch.
