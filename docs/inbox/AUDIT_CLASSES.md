# Standing audit classes

*The defect classes a cloud session re-scans the whole tree for, so that a
class found once is either closed by a ratchet or watched until it is. Each
entry gives the recipe, the count at the last pass, the candidates it
produced, and whether it could become a `tools/check_*.ps1`. The goal named
by the owner, 2026-09-27: catch the classes that produce Terrarium items
before they produce them.*

*Last full pass: 2026-09-27, against `f0fb1db`, candidates in
[`2026-09-27-terrarium-audit.md`](2026-09-27-terrarium-audit.md). A class
with a ratchet moves to the bottom table and stops being audited by hand.*

---

## How a pass runs

1. Re-run every recipe below over the whole tree and compare the count with
   the one recorded here. A count that grew is the first thing reported.
2. Read every new hit, not a sample. Classify it *real* or *benign* and say
   why. A benign hit that a future edit could turn real is still recorded.
3. Write new candidates to a fresh dated inbox file. Update the counts and
   the "last pass" line here, and nothing else in this file.
4. When a class has a stable benign set, propose its ratchet: the recipe
   plus an allowlist of the benign sites, in the house shape of
   `check_ptrsafe_declares.ps1` (hardcoded, reviewable, mutation-tested).

The classes are grouped by what goes wrong, not by module. Most were found
in more than one slice, which is what makes them classes.

---

## 1. Correctness: a value's spelling standing in for the value

**A1 · Locale-following intrinsics on source or cell text.** INTRINSICS #1
and #2 say identifiers fold through `VLA_Identity.Fold` and numbers go
through `Val`/`Str$`. REBUILD's R6 promises a lint for it, but none exists.
- Recipe: `grep -nE '\b(LCase|UCase)\$?\(|\bIsNumeric\(|\bCDbl\(|\bCStr\(CDbl|\bCLng\(' src/*.bas`,
  excluding `VLA_Tests*`. For each hit, ask whether the argument can be
  source text or a cell's String. A real `Double` argument is benign.
- Last count: in the translator, 24 `LCase$/UCase$` and 1 `IsNumeric`; in
  VLA.bas, 3 `IsNumeric`; in the interpreter, 2 `IsNumeric` and the `CDbl`
  literal path; in DATALOG, 1 unguarded `CDbl` of 7; in OPTIMIZE, 9
  `CDbl`/`CLng` on validated digits (2 guarded).
- Candidates: C5, C9, C20, C34, C67 and C73; the unbounded `CLng` half is under E1.
- Ratchet: yes, and it is the highest-value one. Ban the five intrinsics in
  shipped modules outside an allowlist of justified sites, each with a
  one-line reason.

**A2 · Spelling used as identity.** Tuple keys, group keys, join keys and
literal matching built from `CStr(value)`, or from a literal as typed,
where the value has a canonical form (`InvariantNumberText`).
- Recipe: `grep -nE 'CStr\((t|arr|tup|l|r|v|e)\b' src/VLA_Relation.bas src/VLA_Datalog.bas src/VLA_Sql.bas src/VLA_Prolog.bas`.
  Also list every place a leaf is compared with `=`/`StrComp` and check
  that both sides passed through the same normaliser.
- Last count: about 14 sites across SQL, DATALOG and Relation, and 3
  families in PROLOG.
- Candidates: C8, C10, C12, C13 and C41.
- Ratchet: partly. One `ValueKey(v)` function, and a scan that forbids a
  `CStr(` inside any function whose name ends in `Key`.

**A3 · The two backends disagree on semantics the goldens do not
exercise.** Interpret and Compile are held to parity only where the corpus
reaches.
- Recipe: for every VLA head with declaration, handler or loop meaning
  (`dim`, `on-error`, `resume`, `label`, `goto`, `for`, TCO, `+`/`-`, and
  every `TryEvalBuiltin` case), write down VBA's rule and the `ExecX` rule
  side by side, and compare builtin arity against what the phrasebooks
  emit (`(instr 1 x y vbtextcompare)` against a two-argument case).
- Last count: 8 real (dim in a loop, handler disarm, TCO locals, the error
  model, operator coercion, the for-counter, declared types, `instr`
  arity).
- Candidates: C1, C3, C4, C19, C27, C35, C36, C37 and C70.
- Ratchet: the arity half, yes. For each `Case "<name>"` in
  `TryEvalBuiltin`, fail if a shipped phrasebook emits `(<name> …)` with
  more arguments than the case reads. The semantic half needs a
  differential corpus, which is a cloud task: a program per head, run
  both ways by the owner's `VerifyReports`.

## 2. State that outlives its moment

**B1 · Host state set without a paired restore on every exit path.**
`DisplayAlerts`, `EnableEvents`, `Calculation`, `StatusBar`, sheet
protection, and `(on-error goto 0)` inside macros.
- Recipe: `grep -nE 'Application\.(StatusBar|Calculation|EnableEvents|DisplayAlerts|Cursor|ScreenUpdating)\s*='`
  in `src/`. For each procedure, confirm that its error path restores the
  saved value (not a constant). In `*.vla`, look for
  `(set! application.\w+ …)` and `(on-error goto 0)` inside a `defmacro`.
- Last count: 3 real in VBA, and 2 real macros (`delete-sheet`,
  `show-all-rows`).
- Candidates: C27, C33, C62 and C81.
- Ratchet: yes for the macro half (a `defmacro` that sets host state must
  wrap its body in a `with-*` bracket).

**B2 · Parse-context and module flags set and cleared, not saved and
restored.** Recursion or an error leaves them wrong for the next caller.
- Recipe: find `m<Flag> = True … m<Flag> = False` pairs around a recursive
  `Parse*` call. List the side entry points (RunVocabTest, EnglishExplain,
  AuditCrossRuleShadow, EnglishTryRule …) that do not reset what
  `EnglishToVla` resets.
- Last count: 3 flags and about 8 per-run collections.
- Candidates: C19 and C60.
- Ratchet: no. A `SaveParseContext`/`RestoreParseContext` design fixes the
  class.

**B3 · Caches whose value is not a pure function of their key.**
- Recipe: for every memo or cache `Put`, list the fields of the stored
  value that can depend on the clock, the host, or module state.
- Last count: 1 in OPTIMIZE (the guard outcome) and 1 in the translator
  (leaking action names).
- Candidates: C49 and C60.
- Ratchet: no. Audit every new cache.

**B4 · Re-entrancy through Excel events.** A program's own writes, or the
IDE's, fire user handlers mid-run.
- Recipe: every interpreter entry (`VlaInterpret*`, `VlaDispatch*`) must
  sit inside an `EnableEvents` save/set/restore. Every `mApp_*` handler
  must use its `Sh`/`Target`.
- Last count: 2 unguarded entries and 1 handler that discards its target.
- Candidates: C16, C17 and C52.
- Ratchet: yes for the entry bracket.

**B5 · Data persisted into the user's workbook.** Very-hidden sheets and
document properties travel with a shared file.
- Recipe: `grep -nE 'xlSheetVeryHidden|CustomDocumentProperties|\.Names\.Add' src/VLA_IDE.bas`.
- Last count: 3 kinds.
- Candidates: C18 and C63.
- Ratchet: no. It needs a standing decision on what may persist, and then
  a list check.

## 3. Destruction by replacement

**C1 · Replace instead of restore.** Recovery that swaps an object (a sheet
by copy-and-delete, a range by writing `.Value` back) rather than
restoring its contents, which silently breaks everything that pointed at
the object or at its formulas.
- Recipe: `.Delete` near `.Copy` in `VLA_IDE`; `.Value =` of a whole range
  read earlier with `.Value` (slab write-back).
- Last count: 2 real.
- Candidates: C15 and C2.
- Ratchet: partly. Forbid slab write-back over a range that `HasFormula`
  could make non-empty.

## 4. Text crossing into another language

**D1 · User text spliced into code or formula text.** VBA source, Excel
formulas, and the `=DATALOG`/`=SQL` strings a phrasebook builds. SEC.4's
guard covers one sink shape; SEC.15 is the screening half.
- Recipe: in `src/`, `ToVbaString`, `EmitFormula`, `FormulaQuote`,
  `QuoteDatum` and `\.Value2?\s*=`. In `*.vla`, `(interpolate "=` and
  `(& ` near `set-formula`.
- Last count: 7 value sinks (4 real), 3 formula-building templates (2
  real), and 4 source-emitting sites (2 real).
- Candidates: C6, C58, C30 and C28.
- Ratchet: yes. Every sink calls one guarded writer, and a scan fails on
  any direct `.Value =`/`.Formula =` outside it.

**D2 · Encoding at the edges.** ANSI fallbacks that never trigger or
trigger silently, BOMs, and UTF-8 decoded as ANSI.
- Recipe: `grep -nE 'ADODB\.Stream|Charset|Open .* For Input|Line Input' src/*.bas`.
- Last count: 3 readers with 3 different policies.
- Candidates: C26, C31 and C32.
- Ratchet: yes. There should be one reader (`VlaUtf8Decode` exists), and
  the scan fails on any other.

**D3 · Hand-rolled lexers beside the real reader.** Each handles a
different subset of strings, comments and escapes.
- Recipe: `inLit`, `inText`, `inName` and `c = """"` state machines
  outside `VLA.bas`'s `Tokenize`.
- Last count: 5 (EnglishResolveCheck, VocabTextHasRawForm, EnTokenize's
  raw capture, Lint's splitter, OptimizeCallArgs).
- Candidates: C25, C71 and C51.
- Ratchet: no. Route through the reader, or pin each against a shared
  corpus of lexical edge cases.

## 5. Refusals that are not refusals

**E1 · Raw VBA errors reaching the user.** `CStr` of an error Variant or an
array, `CLng`/`CDbl` overflow on an unbounded literal, error cells in
Tables, and short directives.
- Recipe: `CStr(` on a Variant with no preceding `IsError`/`IsArray`/
  `VarType` guard; `CLng(` on anything derived from user text with no
  length or bound check.
- Last count: about 12 across slices.
- Candidates: C24, C38, C39, C42, C44, C50 and C78.
- Ratchet: partly (the `CLng`-on-text half).

**E2 · Registration accepts what can never work.** Rules whose literals the
tokenizer cannot produce, template slots not in the pattern, unchecked
directive arity, and macro-name overrides.
- Recipe: for each assumption `TryPhrase`/`FormSubstitute` makes, check
  that `AddPhraseRule`/`ValidateRuleItems` refuses its violation.
- Last count: 5 unvalidated properties.
- Candidates: C21, C22, C24 and C61.
- Ratchet: no. Each becomes a load-time refusal with a `fail:` proof.

## 6. Performance cliffs

**F1 · Data-sized quadratic loops.** Positional `Collection.Item(i)` inside
a loop over a data-sized Collection, and `s = s & …` building unbounded
output.
- Recipe: `For (\w+) = .* To (\w+)\.Count` with `\2.Item(\1)` in the body;
  `(\w+) = \1 & ` inside a `For Each`.
- Last count: 4 data-sized in PROLOG/Unify; about 5 renderers; the vocab
  load path.
- Candidates: C64, C66 and C68.
- Ratchet: yes for the named hot modules, as
  `check_optimize_search_discipline.ps1` already does for one module.

## 7. Drift

**G1 · Two copies of one rule.** OPEN-slot matching (VBA and `.iss`),
ribbon against legacy menu, four OPTIMIZE entry bodies, three pipeline
prefixes, and a hyphen-mangling copy.
- Recipe: read for it. Any function body duplicated at 80% or more is a
  candidate (`jscpd`-style token diff).
- Last count: 6.
- Candidates: C53, C65 and C80.
- Ratchet: per pair, in the shape of `check_devrig_mods_parity.ps1`.

**G2 · Comments asserting an invariant that later code broke.**
- Recipe: grep comments for "never", "cannot", "not part of this subset",
  "one-line alias" and "cannot drift", then check each claim against the
  current code.
- Last count: 5 (VLA_Relation "never VLA_Messages"; EvalBool's division
  note; OPTIMISE "alias"; menu "cannot drift"; english.vla:4275 "the macro
  doubles them").
- Candidates: C45, C58, C65, C74, C79 and C80.
- Ratchet: no. It is a reading audit, and cheap for a cloud session.

## 8. The gates themselves

**H1 · A check that can pass vacuously.** A discovery glob with no
non-empty assertion, a report-only script counted as green, an unchecked
native exit code, or a stale floor.
- Recipe: for each `tools/check_*.ps1`, rename its inputs in a copy of the
  tree and confirm that it goes red. `grep -cE 'exit 1'` = 0 means it
  cannot fail. After `git`/`&` calls in `tools/*.ps1`, look for
  `$LASTEXITCODE`.
- Last count: 3 vacuous-pass checks, 3 report-only, 4 unchecked exit
  codes, and 1 stale floor.
- Candidates: C54, C55, C56, C57, C82 and C84.
- Ratchet: yes. A meta-check that runs every check against an empty tree
  and requires non-zero from each.

---

## Classes already under a ratchet

*Listed so a pass does not audit them by hand. A blind spot in one of these
ratchets is still a finding (see H1).*

| Class | Ratchet |
|---|---|
| Module headers, `Option Explicit` | `check_module_heads.ps1` |
| 64-bit `Declare` discipline | `check_ptrsafe_declares.ps1` |
| Network reach | `check_no_network.ps1` |
| Host touches on the translate path | `check_translate_purity.ps1` |
| Message slot arity, catalogue ids | `check_message_slots.ps1` |
| Raw `Err.Raise` growth | `check_raise_ratchet.ps1` |
| Runtime raise dispatch | `check_runtime_raise_dispatch.ps1` (blind to Excel-error helpers: C39) |
| SEC.8 / SEC.9 gates | `check_sec8_provenance_gate.ps1`, `check_sec9_phrasebook_gate.ps1` |
| Word macro opening | `check_word_automation_security.ps1` (Word only: C59) |
| PROLOG budgets, reserved names, operators | `check_prolog_*.ps1` |
| Search data discipline | `check_optimize_search_discipline.ps1` |
| Test assertions that cannot fail | `check_test_assertion_safety.ps1` |
| Licence map | `check_spdx.ps1` |
| ID reuse | `check_id_registry.ps1` (passes when it scans nothing: C55) |
