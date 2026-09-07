# CUTS — where the code splits, and where the licence splits

*Two cuts, decided the night before the beta uploads: the module cuts
`REBUILD.md` describes, priced against the code as it actually is today, and
the repo/licence cuts `SIG.0` and `GO.4` have been waiting on. They are one
question wearing two coats: a licence boundary that does not fall on a file
boundary is unenforceable, and a file boundary that does not know which
licence it sits under gets drawn in the wrong place. This file exists so the
two agree before either is made.*

*Written 2026-09-04 against commit `d82b32c`, measured rather than recalled:
every count below came from a scan of `src/` that evening, and Appendix B
records the scans so the numbers can be re-run. The working tree at the time
also carried the owner's uncommitted `L-FORMAT` edits (585 insertions across
`VLA.bas`, `VLA_Interpreter.bas`, `VLA_Messages.bas`, the tests and
`english.vla`); every number here is at `HEAD`, before them. Decisions the owner has to
make are marked ⚖. Nothing here is legal advice; §4.8 names the two bespoke
clauses a lawyer should read before they ship.*

---

## 0. THE VERDICT IN ONE SCREEN

**Do not split source files tonight.** The user of the beta never sees a
`.bas` file; they see one `.xlam` built from all of them. A split changes the
build manifest, the reload rig, every module's `Private`/`Public` surface, and
line endings, with the owner as the only test lab — the one night the build
must not wobble is the wrong night to move fourteen thousand lines. The split
is a post-upload programme, and §3 prices it: **five free cuts** (zero shared
state, an afternoon each), **three state-bearing cuts** (a state record first,
then the move — two sessions each), and **one genuinely hard cut**
(`VLA_SentenceEngine`, whose 56 cross-section module variables are the entire
difficulty of this project's file layout, exactly as `LX5.2`'s header already
confessed).

**Do decide the licence tonight.** `SIG.0` is a document, it is the cheapest
open item on the roadmap, and every fact that decides it is *more* favourable
today than it will ever be again: **zero outside contributors** (git has one
author across 226 commits), so there is no relicensing debt; no `LICENSE` file
at all, so there is nothing to un-publish; and the product's own architecture
already puts an organization's private rules in a **separate file** from the
public corpus, which is the exact line a file-scoped licence draws. §4 gives
four buckets and recommends one: **Apache-2.0 engine, MPL-2.0 corpus, 0BSD on
the one module that gets copied into customers' workbooks, an explicit output
exception, the name held as a trademark, DCO instead of a CLA.** It keeps the
most doors open, needs the least paperwork, and is the only bucket where a
compliance buyer's lawyer, an org writing a private phrasebook, and a future
acquirer all get the answer they need without a phone call.

**The one place the two cuts must agree, stated so it is not rediscovered:**
`VLA_Runtime.bas` above its `EN_RUNTIME INJECT BOUNDARY` is copied *verbatim*
into every exported customer workbook as `Frazaro_EN_Runtime`. Whatever licence
that text carries, the customer carries. It is the GCC-runtime-exception
problem in one file, and it is also why that module must stay whole (`R7`)
and why it gets the most permissive licence in the repository.

---

## 1. WHAT `REBUILD.md` PREDICTED, AND WHAT EXISTS

`REBUILD.md` was written against a 16,662-line `.bas` body. The body is now
**54,005 lines** across 24 standard modules, three classes and one form, and
the shape it proposed has been partly built without anyone filing it as
"the rebuild" — which is the incremental path §7 of that document said had the
better odds, working. The ledger, so the plan below starts from what is rather
than from what was planned:

| `REBUILD.md` target | Status today | Evidence |
|---|---|---|
| Layer 0 — `VLA_Identity` | ✅ built, exactly two verbs | `VLA_Identity.bas`, 62 lines |
| Layer 0 — `VLA_Messages` | ✅ built; 311/343 raise sites migrated (`LX.2`) | `VLA_Messages.bas`, 674 lines |
| Layer 0 — `VLA_Context` / `VlaFrame` | 🟡 `VlaFrame.cls` exists for the interpreter; compile state still lives at module level in `VLA.bas` (33 variables) | §2 |
| Layer 1 — `VLA_HeadTable` | 🟡 built as a *catalogue* (IN.1, IN5.0); by its own header "still NOT the emitter's dispatch itself, for six reasons" | `VLA_HeadTable.bas`, 344 lines |
| Layer 1 — `VLA_Forms`, `VLA_Reader`, `VLA_Expander`, `VLA_VbaStatements`/`Expressions`, `VLA_FormulaEmitter` | ⬜ all still inside `VLA.bas` (5,378 lines), which has since absorbed listops, quasiquote, `cond`, slabs and `for-each-row` | §3.2 B |
| Layer 1 — `VLA_Interpreter` | ✅ built, and it is the default runtime (IN.9) | 3,524 lines |
| Layer 2 — `VLAX_RuntimeHelpers` | ✅ in substance (`VLA_Runtime.bas`, whole, R7 honoured, boundary marked); ⬜ the `VLAX_` prefix and R3's derived manifest | 1,810 lines |
| Layer 3 — `VLA_EnglishWords` | ✅ in substance: `LX5.2` left `VLA_English.bas` holding exactly the English vocabulary tables and the keyword-alias seam | 303 lines |
| Layer 3 — `VLA_Sentences`, `VLA_Matcher`, `VLA_Phrasebooks`, `VLA_EnglishRenderer` | ⬜ all inside `VLA_SentenceEngine.bas` (8,162 lines) | §3.2 D |
| Layer 4 — `VLA_Commands`, `VLA_Panel`, `VLA_Workspace`, `VLA_ModuleInjection` | ⬜ all inside `VLA_IDE.bas` (3,211 lines) — but with only **2** module-level variables, the cheapest big split in the project | §3.2 A |
| Layer 4 — `VLA_Capabilities` | ⬜ not started | — |
| Layer 4 — `VLA_AddinBuild` | 🟡 `VLA_Build.bas` builds every edition; manifest still hand-typed | `VLA_Build.bas:371` vs `VLA_DevRig.bas:156` |
| Layer 5 — the `VLAT_` harness | 🟡 split by concern (`F.8`: Tests / Grammar / Host / Query) but under the `VLA_Tests_` prefix; four files total 16,693 lines | §3.2 F |
| Lints (`VLAT_Coverage`) | 🟡 seven `tools/check_*.ps1` scripts exist, host-independent, house style; none checks R2 layering, R4 line budget, or R3 prefixes | `tools/` |
| Not in `REBUILD.md` at all | `VLA_Sql` (3,831), `VLA_Datalog` (2,162), `VLA_Prolog` (2,204), `VLA_Relation` (1,311), `VLA_Unify`, `VLA_Events`, `VLA_Lint`, `VLA_Browser` — an entire query tier the topology document never anticipated | §3.2 E |

**Fourteen of twenty-four modules are over R4's 1,200-line budget.** That is
the headline, and it is less alarming than it sounds, because the budget
was written for modules that hide *coupling*, and §2 shows that most of the
overage is in modules that hide none.

---

## 2. THE MEASUREMENT THAT PRICES A SPLIT

In VBA a split has exactly one cost that matters, and it is not line count.
A `Private` module-level variable cannot be seen from another file. So every
variable that two would-be modules both touch forces one of two moves:

- **Make it `Public`** — one word changed, zero behaviour change, one more
  name in the single project-wide namespace (constraint 1 of `REBUILD.md` §1).
  Cheap, honest, and ugly in a way that is visible in one file.
- **Put it in an object** — `R8`'s answer, the `VlaFrame` pattern. Correct,
  and a real refactor of every procedure that reads the variable.

Everything else about a split is clerical: the `Attribute VB_Name` line, the
two hand-maintained manifests (`VLA_Build.bas:371`, `VLA_DevRig.bas:156`),
`DEPLOY.md`'s module list, and the line-ending convention (every file in
`src/` is LF-only at `HEAD` — confirmed per file, not assumed; a new module
written CRLF would be the first). Private *procedures* that cross the cut go
`Public` the same way, with one trap: a `Public` name that already exists in
another module is a compile-time "Ambiguous name" the moment it is called
unqualified. Those collisions were counted.

| Module | Lines | Procs | Module-level vars | Vars shared across its own sections | Private names already defined elsewhere | Split grade |
|---|---|---|---|---|---|---|
| `VLA_SentenceEngine` | 8,162 | 205 | **76** | **56** (`mPatItems` alone from 9 of 22 sections) | 8 | **Hard** — the rule store is one blob |
| `VLA.bas` | 5,378 | 132 | 33 | 23 — but two sections (datum helpers, `deflambda`) touch **zero** | 13 | Mixed: two free cuts, then a state record |
| `VLA_Interpreter` | 3,524 | 81 | 13 | 12 — all non-local control signals (`mCaughtErr*`, `mLoopBreak`, `mProcReturn`, `mGotoLabel`) | 8 | One free cut (member dispatch), rest is one machine |
| `VLA_IDE` | 3,211 | 95 | **2** | 1 | **0** | **Free** |
| `VLA_Sql` | 3,831 | 121 | **0** | 0 | 2 | Free, low value |
| `VLA_Datalog` | 2,162 | 46 | **0** | 0 | — | Free, low value |
| `VLA_Prolog` | 2,204 | 36 | **0** | 0 | — | Free, low value |
| `VLA_Runtime` | 1,810 | 71 | 5 | — | — | **Do not split** (R7, constraint 4) |
| `VLA_DevRig` | 1,766 | 41 | — | — | — | Dev-only; split when convenient |
| `VLA_Tests*` (4 files) | 16,693 | 205 | — | — | — | Dev-only; split by banner when a file is open anyway |

*Three of the largest modules in the repository carry no module-level state
at all* (the query tier was written after `R8` was known and it shows). Their
overage is length, not coupling, and the proportion test in `REBUILD.md` §4
says length alone does not buy a file.

The collision lists are worth reading once, because they are the same names
every time: `AssignVar`, `ClickHandlerSlug`, `CollGet`, `SbAdd`/`SbText`,
`IsList`/`Nth`/`HeadSym`/`SymText`/`StrLitContent`, `Tokenize`, `ProfMs`.
These are `R7` duplicates and rule-12 duplicates — small helpers deliberately
copied per module so the injectable and dev-only modules compile alone. A
split that makes them `Public` in one module collides with the copy in
another. **The rule for the split, then: a helper that exists in two modules
as `Private` stays `Private` in both and is copied into the new module too;
only single-definition names cross a cut as `Public`.** Thirteen names in
`VLA.bas`, eight each in the sentence engine and interpreter — a checklist,
not a blocker.

---

## 3. THE SPLIT PLAN

### 3.1 One split, one commit — the procedure

`REBUILD.md` §6's bootstrap, made concrete for this codebase. Every split runs
this list; a split that skips a line is not done.

1. **Predict the diff.** Write down, before touching anything, that
   `VlaWriteGoldens` will produce an empty diff and that `VlaSelfTest` /
   `VlaSelfTestHost` will print the roadmap's last recorded baseline
   (903 / 119 at `EDITION-MANIFEST`; re-read the current count first).
2. **Create the new file** with `Attribute VB_Name = "<name>"`,
   `Option Explicit`, a `Public Const <NAME>_VERSION`, and the `REBUILD.md`
   Appendix C header — `LAYER`, `MAY CALL`, `SHIPS`, `PAYS INTO`, `REASON`.
   LF line endings, matching every other file in `src/` at `HEAD`.
3. **Move procedures by section banner**, whole sections only. Do not
   reorganize inside a section in the same commit.
4. **Promote crossing names.** Every `Private` the old module still calls
   becomes `Public` in the new one — after `grep -l "Sub <name>\b\|Function
   <name>\b" src/*` shows exactly one definition. Names on §2's collision
   list are copied, not promoted.
5. **Promote crossing state** — or, for the state-bearing cuts, land the
   state record *first* (its own commit, §3.2 B/D) so this step is a rename.
6. **Both manifests, and `DEPLOY.md` step 3's list.** Until `R3` derives
   them, they are edited by hand, in the same commit, or the reload rig and
   the build disagree — the drift `EDITION-PARITY`'s own text already
   records happening once.
7. **Debug → Compile**, then the three plates: self-test counts, golden
   diff, `VerifyReport` on a real Run. The owner runs these; the handoff
   names each one.
8. **Commit** with the module name in the subject and the predicted-vs-actual
   diff in the body. Delete-the-old-shell is its own later commit.

Ten lines of clerical work per split, once, plus whatever step 5 costs — and
step 5 is the whole table in §2.

### 3.2 The cuts, in the order they pay

Priority is value over owner-hours, because owner-hours at a keyboard with
Excel open are the constraint (`F.11`, `CONTINUITY.md`). Names follow
`REBUILD.md` Appendix D where a name exists there and `R11` where one does not.

**A. `VLA_IDE.bas` → four modules.** *Free. First.* Two module-level
variables, zero name collisions, sections already labelled. The payoff is
`REBUILD.md`'s "`VLA_Commands` never talks to a human" — the pure/host split
at the product layer that `U.14`, `AC.1` and `F.11` all cite, and it makes the
`SEC.*` attribution question ("which module opened a dialog?") answerable by
filename.

| New module | From `VLA_IDE.bas` | Lines (approx.) |
|---|---|---|
| `VLA_Commands` | ambient-state capture (`:264`), Check/Run/Undo orchestration, the recovery-path truth table (`:2082`) | ~900 |
| `VLA_CommandLine` | the `frmCLI` bridge (`:785`–`:1490`), 22 procedures | ~700 |
| `VLA_Workspace` | program identity (`:1864`), snapshots and undo (`:2122`), the file bridge (`:2917`), feedback export (`:3117`) | ~1,000 |
| `VLA_AddinRegistration` | self-registration (`:2531`), Uninstall (`:2644`), the Add-ins menu (`:2826`) | ~500 |
| *(stays)* `VLA_Panel` — the ribbon callbacks and every `MsgBox` | what is left | ~100, and the only module allowed to grow a dialog |

⚖ `REBUILD.md` files registration/uninstall nowhere; `VLA_ModuleInjection`
is VBProject manipulation, which today lives in `VLA_Runtime`'s injector and
`VLA_Loader`, not here. `VLA_AddinRegistration` is a new name, offered under
`R11`: it decides whether Excel knows the add-in exists.

**B. `VLA.bas` → two free cuts, then three state-bearing ones.**

*Free, immediately:* the **datum helpers** (`:2108`–`:2239`, 12 procedures,
zero state) become **`VLA_Forms`** — the form ADT `F.2` needs and the only
Layer 1 module the interpreter, the emitters and the sentence layer all
share. The **`deflambda` / formula path** (`:4890`–end, 15 procedures, zero
state) becomes **`VLA_FormulaEmitter`**. Two afternoons; two of the four
backends become visible in the file listing.

*Then the state record.* Twenty-three variables cross the remaining
sections. Land **`VLA_CompileState`** — one module, every one of those
variables as `Public`, no procedures, header stating it is `R8`'s interim
record and will become fields on `VlaFrame` — as its own commit with an
empty golden diff. It is not the architecture; it is the architecture's
coupling made visible in one file, which is the precondition for fixing it.

*Then the moves:* **`VLA_Reader`** (tokenizer + parser + `include`,
`:1702`–`:2108`, 16 procedures), **`VLA_Expander`** (the macro system,
`:2239`–`:3791`, 37 procedures — the largest single section in the
repository), and **`VLA_VbaEmitter`** (top-level + statements + expressions,
`:3791`–`:4890`, 25 procedures, ~1,100 lines). `REBUILD.md` split the emitter
in two; measured, the expression half is 2 procedures and 123 lines, so one
module fits the budget and the two-way split is deferred until it does not.
The Public API and `macroexpand` tooling stay in `VLA.bas`, which ends near
1,000 lines and keeps `VLA_RELEASE_VERSION`.

**C. `VLA_Interpreter.bas` → one cut.** Section 6 (`:2281`–`:3456`, 30
procedures, 1,175 lines) is the `CallByName` member-dispatch machinery — the
heuristic `THREAT_MODEL.md` §1.2 names and the exact site `SEC.1`'s
capability gate will live at. Make it **`VLA_MemberDispatch`**: the security
architecture then has a filename, and a reviewer reading `SIG.1` can be
pointed at one module. It touches 9 of the 13 module variables, all
control-signal state (`mCaughtErr*`, `mErrMode`, `mLoopBreak`, `mProcReturn`,
`mGotoLabel`) — promote them `Public`, or better, land them as fields on
`VlaFrame`, which already exists and is where `R8` says they go. The
`Exec`/`Eval` halves stay together: they share the same nine signals and are
one machine; a 2,300-line interpreter is a *documented* R4 exception, not a
missed one.

**D. `VLA_SentenceEngine.bas` → a state record, then five modules.** *The
hard one, and the reason this document exists.* `LX5.2`'s header already
found that the rule store is "NOT separable along the grammar sections" —
§2 confirms it: 56 of 76 variables cross sections, `mPatItems` is touched from
nine. There is no cut through this file that does not cross the store, so the
store leaves first:

*First:* **`VLA_GrammarStore`** — all 56 cross-section variables, `Public`,
no logic, one header. Same shape as `VLA_CompileState`, same empty-diff
commit. (⚖ `R11`: it holds the registered patterns, their sources, their
tests, the dispatch index and the diagnostics scratch — "grammar store" is
what a reader would guess; `VLA_RuleStore` is the alternative.)

*Then, by section, in this order:*

| New module | Sections (by banner line) | Procs | Why this grouping |
|---|---|---|---|
| **`VLA_EnglishRenderer`** | G-RENDER `:1752`–`:2894` | 33 | `REBUILD.md`'s own Layer 3 correction; 15 store variables, all read-only |
| **`VLA_Phrasebooks`** | grammar registration `:1305`, vocabulary loading `:5526`, `VOCABDIFF` `:6075`, rule-usage profiling `:6570`, the audit `:8088` | 49 | everything that turns a `.vla` file into store rows or reports on them; the `F.1` ABI check and the dot count land here |
| **`VLA_Sentences`** | tokenizer `:2894`, statement/block parsing `:3440`, expressions `:4742`, conditions `:5135`, sub-assembly `:5259` | 66 | the structural grammar — and the module the "no English literals" lint (`REBUILD.md` Layer 3) runs against |
| **`VLA_Matcher`** | phrase matching `:4354`, best-partial-match diagnostics `:6948`, reference-shape validators `:6999` | 25 | `REBUILD.md` files near-miss and conflict analysis with the matcher because they are the same math |
| **`VLA_PhrasebookAuthoring`** | call checking `:7396`, the front door `:7475`, translate `:7507`, function-word registry `:7686`, `EnglishExplain` `:7795`, the scratchpad `:7808` | 29 | the author's tools — `Explain`, `Audition`, the scratchpad — which `DO.4` will document as a surface and which no user of the product ever calls |

`VLA_Sentences` lands near 2,000 lines, over budget; the tokenizer (22
procedures, 7 store variables) is the second cut when the file is open again.
The Public API section (`:751`, 9 procedures, **41** store variables —
`EnglishToVla` drives the whole parse loop) stays in `VLA_SentenceEngine.bas`
as the front door, exactly as `LX5.2` said it must.

**E. The query tier — leave it.** `VLA_Sql`, `VLA_Datalog`, `VLA_Prolog` carry
no module-level state, have clean internal banners, and are each one
concept. R4 says split; the proportion test says a 3,800-line SQL engine
split into parser and evaluator buys navigation and nothing else. Split by
banner the next time one of them is open for a feature, never as its own
pass.

**F. The harness, the prefix, and the manifest — one pass.** Rename
`VLA_Tests*` and `VLA_DevRig` to `VLAT_*`, `VLA_Runtime` to
`VLAX_RuntimeHelpers`, and in the *same* commit make `VLA_Build` and
`VLA_DevRig` derive their module lists from the prefix (`R3`). The rename
only pays for itself if the manifests stop being hand-typed the day the
prefixes become meaningful; done separately, it is nineteen renames for
nothing. The `VLAT_Coverage` lints (`R2` layering from headers, `R4` line
budget, `R3` prefix) are three more `tools/check_*.ps1` scripts in the house
style — PowerShell, host-independent, hardcoded baseline, never wired into
`VlaSelfTest`.

### 3.3 The price, honestly

| Cut | Owner sessions | Behaviour risk |
|---|---|---|
| A. `VLA_IDE` → 4 | 1–2 | nil — no shared state |
| B1. `VLA_Forms`, `VLA_FormulaEmitter` | 1 | nil |
| B2. `VLA_CompileState` record | 1 | nil (Private → Public) |
| B3. `VLA_Reader`, `VLA_Expander`, `VLA_VbaEmitter` | 2–3 | low — reentrancy is unchanged, just visible |
| C. `VLA_MemberDispatch` | 1–2 | low; the effect-log golden is the witness |
| D1. `VLA_GrammarStore` record | 1 | nil |
| D2. five sentence modules | 4–6 | **medium** — 160 private procedures, 8 collisions, and `Explain`'s attribution state (`A4` goldens) must survive |
| E. query tier | 0 | — |
| F. prefixes + derived manifest + three lints | 2 | low, but touches the build |
| **Total** | **13–19 sessions** | |

Call it six to eight weeks of evenings, sequenced so every commit ships. That
is the number to put against the alternative, which is not "never split" but
"split the day a second language or a security review forces it, in a hurry."

### 3.4 What tonight's upload does not need

None of §3. The `.xlam` is built from whatever `src/` contains; `REBUILD.md`
§1 constraint 3 is explicit that more modules buy no compile safety. What the
upload *does* need from this document is §4 — because the beta goes out under
some licence or under none, and "none" is all-rights-reserved, which the
roadmap's own `SIG.0` text calls a 100% procurement block.

---

## 4. THE LICENCE CUTS

### 4.1 The facts that decide it — measured, not assumed

1. **One author.** `git shortlog` shows 226 commits, one name, one email.
   `README.md` already declines contributions until a licence exists. There
   is no contributor whose permission a relicence would need. This is the
   single largest reason to decide now: it is the last day the decision is
   free.
2. **Code leaves the building.** On the export path, `VLA_Runtime.bas`'s
   injectable half is copied verbatim into the customer's workbook as
   `Frazaro_EN_Runtime`, beside a generated `Frazaro_EN_<program>` module
   whose text is assembled from `prelude.vla` macro bodies and `english.vla`
   templates. A customer who emails that workbook is *redistributing* every
   licence those three sources carry. Whatever the engine's licence, those
   three need the lightest terms in the repository or an explicit exception.
   `VENTURE.md` §12 already calls the output exception non-negotiable; this
   is the mechanism it was talking about.
3. **Org phrasebooks are already separate files.** `IdeVocabPath` loads an
   organization's own `.vla` from beside its workbook, layered over the
   embedded base; the `override:` directive replaces a base rule *from the
   overlay file*, never by editing the base. That is precisely the line a
   file-scoped copyleft draws, which is why one of the buckets below can be
   share-alike on the corpus without touching a private phrasebook.
4. **But overlays depend on the base.** `espanol.vla`'s own header: *"Carga
   english.vla PRIMERO"* — it calls macros the base defines. Under a
   strong-copyleft base, that dependency is the Emacs-and-elisp argument
   that an overlay is a derivative work. Any bucket that puts GPL on
   `english.vla` needs a written exception for overlays, or the "private
   org phrasebook" business does not exist.
5. **Phrasebooks are code.** They carry `test:` proofs, `defmacro` bodies,
   and templates that emit VBA — seven `raw` sites in `english.vla` splice
   VBA text directly. Creative Commons licences say in their own FAQ not to
   use them for software; a phrasebook gets a software licence.
6. **SD-10 binds.** Semantics are never a pricing boundary: every core form,
   backend and grammar section is free. So the engine and the base corpus
   are open in every bucket, and the only closable things are what `DI.5`
   already lists — seats, support, registry access, hosted services.
7. **SD-13 removes the usual reason for a source-available engine.** The
   product makes no network call. The threat that BSL/FSL/SSPL exist to
   answer — a cloud vendor hosting your engine as a service — cannot happen
   to a desktop VBA add-in. Fair-source terms belong on the *services*, if
   and when they exist, not on the engine.
8. **Authorship is human-directed and documented.** The method (`VENTURE.md`
   §10: assistant writes blind, human is compiler, runtime and QA) is the
   kind of human specification, selection and verification current US
   Copyright Office guidance treats as authorship, and the repo's
   adjudication trail is unusually good evidence of it. But copyleft's
   teeth are copyright's teeth; a permissive licence does not care how the
   question resolves. ⚖ A counsel item, not a blocker — and a mild thumb on
   the scale toward permissive terms for the engine.

### 4.2 The territories

What actually gets a licence line, so the buckets can be read as a table:

| Territory | What it is, concretely | Where it lives today |
|---|---|---|
| **Engine** | every `VLA_*.bas`, the classes, `frmCLI` | `src/` |
| **Injectable runtime** | `VLA_Runtime.bas` above the boundary — copied into customer workbooks | `src/VLA_Runtime.bas` |
| **Generated output** | `Frazaro_EN_<program>`, the user's own `.vla`/English text, anything `Compile`/`Export` writes | the customer's workbook |
| **Prelude** | `prelude.vla` — Layer 1 by `REBUILD.md`'s reading; its macro bodies land in output | `scripts/` |
| **Base corpus** | `english.vla` | `scripts/polyglotta/` |
| **Editions and dialects** | `espanol.vla` and the seven siblings | `scripts/polyglotta/` |
| **Org phrasebooks** | an organization's private overlay | *never in this repo* — theirs |
| **Fixtures and goldens** | `instructions*.txt/.vla/.vba`, `interpreter_golden.txt` | `scripts/` |
| **Docs** | everything in `docs/`, `README.md` | `docs/` |
| **Tools, installer** | `tools/*.ps1`, `installer/` | as named |
| **Services** *(future)* | registry server, compliance-pack attestation tooling (`VENTURE.md` §6.3–6.4) | do not exist |
| **The mark** | "Frazaro" — also appears in emitted module names | everywhere |

### 4.3 The buckets

Four coherent knots. Each is internally consistent; mixing rows across
buckets is how projects end up with a licence nobody can explain.

#### Bucket 1 — *The Gift, insured.* Permissive everywhere, the name held.

| Territory | Licence |
|---|---|
| Engine, prelude, tools, installer, fixtures | Apache-2.0 |
| Injectable runtime | 0BSD (zero-attribution) |
| Generated output | not covered — stated in `OUTPUT-EXCEPTION.md` |
| Base corpus, editions | Apache-2.0 |
| Docs | CC-BY-4.0 |
| Services | FSL-1.1-Apache-2.0, separate repo, when they exist |
| Mark | held; `TRADEMARK.md` |
| Inbound | DCO |

*Buys:* maximum adoption; zero legal friction for any buyer; no CLA ever
(Apache inbound can be used in anything, so relicensing flexibility is free);
the AI-authorship question is moot. *Costs:* a competitor may fork the
corpus closed, and phrasebook improvements need not flow back. `VENTURE.md`
§8 already argues the corpus's moat is the *discipline* that produced it, not
the licence — this bucket takes that argument at its word. *Destiny served:*
The Gift, The Dowry (cleanest possible acquisition), The Lifestyle Product.

#### Bucket 2 — *Open engine, share-alike corpus.* ★ Recommended.

| Territory | Licence |
|---|---|
| Engine, prelude, tools, installer, fixtures | Apache-2.0 |
| Injectable runtime | 0BSD |
| Generated output | not covered — `OUTPUT-EXCEPTION.md` |
| **Base corpus, editions** | **MPL-2.0** |
| Docs | CC-BY-4.0 |
| Services | FSL-1.1-Apache-2.0, separate repo, when they exist |
| Mark | held; `TRADEMARK.md` |
| Inbound | DCO |

MPL-2.0 is *file-scoped* copyleft: the obligation attaches to a file that
contains covered text, and it triggers only on **distribution**. Read against
fact 3 and fact 4 above:

- an org's overlay in its own file, calling `english.vla`'s macros, is an MPL
  "Larger Work" — not covered, theirs, closed if they like;
- an org that edits `english.vla` itself and *ships* the result must ship
  that file's source — which is the flow-back `VENTURE.md` §12 wanted;
- an org that edits `english.vla` and uses it *internally* owes nothing at
  all, because nothing was distributed;
- `espanol.vla` and future editions inherit MPL and stay open, which is the
  right default for the edition line.

*Buys:* everything Bucket 1 buys for the engine (no CLA, no friction, full
owner flexibility), plus a legal edge on the compounding asset — and the
licence line and the product's `override:` design say the same thing, so
the architecture is the FAQ. *Costs:* MPL on a DSL file is unusual enough
that an org's lawyer will ask one question ("is our overlay a Modification?");
§4.7's `PHRASEBOOK-TERMS.md` answers it in four sentences. Fragments of
corpus *templates* appear in generated output, which is why
`OUTPUT-EXCEPTION.md` is written as an additional permission from the
copyright holder rather than left to inference. *Destiny served:* The
Register (a governed commons with private editions is exactly this shape),
The Lifestyle Product, The Gift; The Dowry loses nothing, since the engine is
permissive and the corpus was never going to be sold closed.

#### Bucket 3 — *Capture-resistant.* Copyleft engine, dual-licensed.

| Territory | Licence |
|---|---|
| Engine, prelude | GPL-3.0-or-later **with a Runtime and Output Exception** (GCC RLE pattern) — and a second exception for phrasebooks loaded by the interpreter |
| Injectable runtime | covered by the exception |
| Generated output | covered by the exception |
| Base corpus, editions | MPL-2.0 |
| Docs | CC-BY-4.0 |
| Services | proprietary |
| Mark | held |
| Inbound | **CLA, mandatory** (dual licensing is impossible without it) |
| Commercial | a paid non-GPL engine licence for buyers whose policy bans copyleft |

*Buys:* forks must stay open; a second revenue lever (the commercial
licence). *Costs:* two bespoke exceptions, each a review item for every
buyer's counsel ("GPL text in our workbook?" is a question the exception
answers, but it is still asked); a CLA from the first outside contributor;
AGPL — the only variant that would cover a hosted port — is a hard no on
many procurement checklists, and SD-13 means the hosting threat it answers
does not apply anyway; and it is the bucket most exposed to fact 8, since
copyleft is exactly as strong as the copyright behind it. `VENTURE.md` §12
said "without that clause every compliance buyer walks"; this is the bucket
that has to write the clause and then defend it. *Destiny served:* The Gift
in its FSF reading; The Lifestyle Product via dual licensing. Weakens The
Dowry (an acquirer inherits the exceptions and the CLA archive).

#### Bucket 4 — *Fair-source core.* Considered for the engine; rejected; kept for the services.

| Territory | Licence |
|---|---|
| Engine | FSL-1.1-Apache-2.0 or BSL-1.1 (converts to Apache after 2 / ≤4 years) |
| everything else | as Bucket 2 |

*Buys:* protection against a hosted competitor, which fact 7 says cannot
exist for this product. *Costs:* the project stops being able to say "open
source" without a footnote; `VENTURE.md` §11's "our continuity plan is
`git clone`" and `PREMORTEM.md`'s escrow argument both weaken; the tension
with the mission paragraph is real. **FSL is the right licence for the
registry server and the compliance-pack tooling** — source-available,
non-compete for two years, then Apache — because those are the services SD-10
lets be priced and the only places a hosted competitor is even conceivable.
That is where every other bucket already puts it.

### 4.4 The scorecard

Against the project's own constraints, not a generic one:

| Constraint | B1 Gift | **B2 Share-alike corpus** | B3 Copyleft engine | B4 Fair-source engine |
|---|---|---|---|---|
| SD-10 (semantics free) | ✅ | ✅ | ✅ | ⚠ open-eventually |
| Output exception airtight | ✅ trivially | ✅ trivially (+ one stated permission) | ⚠ bespoke clause, must be defended | ✅ |
| Private org phrasebook legal with no contract | ✅ | ✅ (Larger Work; internal use never triggers) | ⚠ needs second exception | ✅ |
| Compliance buyer's counsel: zero questions | ✅ | ✅ one question, answered in `PHRASEBOOK-TERMS.md` | ❌ | ❌ |
| CLA needed | no | no | **yes** | for relicensing, yes |
| Corpus improvements flow back | no | **yes, on distribution** | yes | yes |
| Owner's future flexibility (tighten later) | full | full on engine | only via CLA | full |
| Exposure to the authorship question (fact 8) | none | corpus only | **engine** | engine |
| Lawyer hours to ship | ~0 | ~1 (read two short files) | many | some |

Reversibility deserves its own sentence, because "cut twice, measure once" is
the joke and irreversibility is the risk. The owner is sole copyright holder;
under Buckets 1 and 2, any *future* version can be released under stricter
terms if the destiny changes (the HashiCorp move), while the released
versions stay open — the path from permissive to strict is always open to
the holder. The path from copyleft-with-contributors back to permissive is
closed the day the first outside patch lands without a CLA. **Bucket 2 keeps
the most doors open with the least paper.**

### 4.5 ⚖ The recommendation — **decided: Bucket 2, owner, 2026-09-04**

*The files in §4.7 now exist at the repository root, `tools/check_spdx.ps1`
verifies the map, and the three disclosure steps from the SEC.\* assessment
(README *Known open security items*, `DEPLOY.md`'s download copy,
`installer/POST_INSTALL.txt`) landed in the same pass. The paragraph below
is kept as the record of why.*

**Bucket 2.** Apache-2.0 on the engine and prelude; MPL-2.0 on `english.vla`
and the editions; 0BSD on `VLA_Runtime.bas`; an output exception stated as an
additional permission; CC-BY-4.0 on `docs/`; the name held under a short
trademark policy; DCO inbound; FSL reserved for services that do not yet
exist, in a repository that does not yet exist.

If the owner's read of the moat matches `VENTURE.md` §8 exactly (the
discipline is the moat, the licence is decoration), Bucket 1 is the same
decision with one fewer licence text. The difference between them is
whether "improvements to the public phrasebook flow back" is a value the
project wants to *write down* — and `VIABILITY.md` N.2 says the copyleft-
versus-permissive question on the *corpus* is the hinge between a commons
and a market. Bucket 2 chooses the commons for the corpus and the market for
the engine, which is the only pairing that matches what the product's own
loader already does.

### 4.6 The repo cuts

**Do not move a single file tonight.** The last time `english.vla` moved
directories it broke three things the same session (`EDITION-MANIFEST`'s
record: the build's hardcoded path, the dev fallback, the installer's
`[Files]`). Licence territories do not need directories; they need
**per-file declarations plus a path map**, and there is a standard for
exactly this: the REUSE specification (`reuse.software`). Its shape is three
things and it is lint-able in the house style:

1. `LICENSES/` holding the full text of every licence used —
   `Apache-2.0.txt`, `MPL-2.0.txt`, `0BSD.txt`, `CC-BY-4.0.txt`.
2. `REUSE.toml` at the root mapping path globs to SPDX identifiers and the
   copyright line — the *path map* that lets `scripts/polyglotta/*.vla` be
   MPL while `scripts/prelude.vla` beside it is Apache, without a move.
3. An `SPDX-License-Identifier:` comment in every source file, which is
   what a reader and a scanner both actually check. A `tools/check_spdx.ps1`
   in the existing `check_*.ps1` shape — host-independent, hardcoded
   baseline, never wired into `VlaSelfTest` — verifies every `.bas`, `.cls`,
   `.frm`, `.vla`, `.ps1` carries one and that it agrees with `REUSE.toml`.

Where the header goes, per file type, because the first line is spoken for:

```vb
Attribute VB_Name = "VLA_Reader"
Option Explicit
' SPDX-License-Identifier: Apache-2.0
' SPDX-FileCopyrightText: 2025-2026 <copyright holder>
```

```lisp
; SPDX-License-Identifier: MPL-2.0
; SPDX-FileCopyrightText: 2025-2026 <copyright holder>
; english.vla - ...
```

**`VLA_Runtime.bas` is the one file where the header must sit *inside* the
injectable region**, above the boundary, so it travels into the customer's
workbook with the code it licenses. The fence check (`TestRuntimeModule`,
`VLA_Tests_Host.bas`) scans that region for forbidden procedure names; a
licence comment names none. A two-line 0BSD notice there is the entire
runtime-exception mechanism, and it costs nothing.

**When the directory move does happen** — with cut B in §3, when `prelude.vla`
becomes the expander's file — the territories become directories and
`REUSE.toml` shrinks: `src/` Apache, `phrasebooks/` MPL, `docs/` CC-BY. Not
before.

*Built 2026-09-04:* the repository skeletons now exist as `../Delta/`
beside this tree, one folder per future GitHub repository (`frazaro`,
`frazaro-phrasebooks`, `frazaro-org-phrasebook-template`,
`Frazaro-Services`), each with its own licence, governance files, and a
`HYDRATE.md` naming which `Beta/` files fill it; `Delta/README.md` is the
map and the hydration order.

**A second repository the day services exist, not before.** FSL code in a
public Apache repository is a mistake waiting for a copy-paste; and the
registry server is the one artifact whose *absence* from the public repo is
the point. Org phrasebooks live in the org's own repository, under the org's
own terms, and `PHRASEBOOK-TERMS.md` says so in writing. A community
phrasebook registry (`GO.3`) is a third repository with its own trust model,
and it inherits MPL from the corpus it extends.

### 4.7 Closing `SIG.0` tonight — the files, drafted

Once ⚖ 4.5 is decided, `SIG.0` is nine small files and one README edit. The
canonical licence texts are pasted from `spdx.org/licenses`, never retyped.
The bespoke text is drafted here so tonight is a paste, not a composition.
*Legal-name and year placeholders are deliberate; fill from the record, not
from memory.*

**`LICENSE`** — the full Apache-2.0 text, verbatim. (Root file; what GitHub
and every scanner read first.)

**`LICENSES/`** — `Apache-2.0.txt`, `MPL-2.0.txt`, `0BSD.txt`,
`CC-BY-4.0.txt`, verbatim.

**`REUSE.toml`**

```toml
version = 1
SPDX-PackageName = "Frazaro"
SPDX-PackageSupplier = "<copyright holder> <english@spreadsheet.company>"

[[annotations]]
path = ["src/**", "scripts/prelude.vla", "scripts/*.txt", "scripts/*.vla", "scripts/*.vba", "tools/**", "installer/**"]
SPDX-FileCopyrightText = "2025-2026 <copyright holder>"
SPDX-License-Identifier = "Apache-2.0"

[[annotations]]
path = ["src/VLA_Runtime.bas"]
precedence = "override"
SPDX-FileCopyrightText = "2025-2026 <copyright holder>"
SPDX-License-Identifier = "0BSD"

[[annotations]]
path = ["scripts/polyglotta/**"]
SPDX-FileCopyrightText = "2025-2026 <copyright holder>"
SPDX-License-Identifier = "MPL-2.0"

[[annotations]]
path = ["docs/**", "README.md"]
SPDX-FileCopyrightText = "2025-2026 <copyright holder>"
SPDX-License-Identifier = "CC-BY-4.0"
```

**`NOTICE`**

```text
Frazaro
Copyright 2025-2026 <copyright holder>

This product includes software developed by <copyright holder>.
Phrasebooks under scripts/polyglotta/ are licensed separately under the
Mozilla Public License 2.0; VLA_Runtime.bas under the BSD Zero Clause
License. See LICENSES/ and REUSE.toml.
```

**`OUTPUT-EXCEPTION.md`** — the additional permission fact 2 requires:

> **Frazaro Output Exception**
>
> As an additional permission granted by the copyright holder under every
> licence in this repository:
>
> 1. **Your programs are yours.** English sentences, `.vla` programs, and
>    worksheets you write are your own work. Frazaro's licences do not apply
>    to them.
> 2. **Generated code is yours.** VBA modules, formulas, or other output that
>    Frazaro generates from your programs — including any text copied into
>    that output from `prelude.vla`, from a phrasebook template, or from the
>    injectable runtime helpers (`Frazaro_EN_Runtime`) — may be used,
>    modified, and distributed under terms of your choosing, without
>    attribution and without any obligation under Apache-2.0, MPL-2.0, or any
>    other licence in this repository.
> 3. **Workbooks are not derivative works.** Distributing a workbook that
>    contains generated code or the injected runtime helpers does not make the
>    workbook, or any other code in it, subject to Frazaro's licences.
>
> This exception does not permit removing the licence notice from a copy of
> Frazaro *itself*, and does not grant any right in the Frazaro name.

**`PHRASEBOOK-TERMS.md`** — the four sentences an org's counsel will ask for:

> `english.vla` and the edition phrasebooks in `scripts/polyglotta/` are
> licensed under the Mozilla Public License 2.0, which applies **per file**.
> A phrasebook you write in your own file — including one that uses
> `override:` to replace a rule, or that calls macros the base phrasebook
> defines — is a *Larger Work* under MPL §3.3: it is yours, under any terms
> you choose, including none. If you edit `english.vla` itself and
> **distribute** the edited file, MPL requires you to make that file's source
> available under MPL; using an edited copy inside your own organization
> creates no obligation at all. Generated code is covered by
> `OUTPUT-EXCEPTION.md`, not by this file.

**`TRADEMARK.md`**

> "Frazaro" is a trademark of <copyright holder>. The source code licences
> in this repository do not grant rights to the name or logo. You may use the
> name to accurately refer to the software, including in the module names
> (`Frazaro_EN_*`) the software itself writes into your workbook, and to say
> that your product or phrasebook works with Frazaro. Please do not use it as
> the name of a modified version, a fork, or a service in a way that suggests
> it is official. Ask first for anything else: `english@spreadsheet.company`.

⚖ Registration is a separate, later, paid step (a class-9 filing); a policy
file costs nothing and establishes use in commerce from tonight.

**`CONTRIBUTING.md`** — replaces the README's "not accepting contributions"
line, and is *only* needed if that line is being retired tonight; otherwise
it waits for `GO.2`:

> Contributions are accepted under the Developer Certificate of Origin
> (`developercertificate.org`): sign each commit with `git commit -s`. By
> contributing you agree your work is licensed under the licence of the file
> it touches (see `REUSE.toml`) and the Output Exception. No CLA.

**`README.md` — "License and status", replacement text:**

> Frazaro is open source. The engine is Apache-2.0; the phrasebooks are
> MPL-2.0 (file-scoped: your own phrasebook file is yours); the one module
> Frazaro copies into your workbook is 0BSD; and everything Frazaro generates
> from your sentences is yours outright — see `OUTPUT-EXCEPTION.md`. The
> name is a trademark; see `TRADEMARK.md`. Full texts in `LICENSES/`, per-
> file map in `REUSE.toml`.
>
> **No warranty.** This is beta software, released unfinished on purpose
> (see the top of this file). It is provided as-is, without warranty of any
> kind, as the licences say in longer words. Take a backup before you run a
> program you have not run before; Frazaro takes a snapshot before every Run
> and offers Undo, and that is a convenience, not a guarantee.
>
> **What a program can do with your data.** Frazaro itself makes no network
> connection, ever, and phones nothing home. A program *you* run can do what
> the sentences say: read and write cells, sheets and files, and, if a
> sentence asks for it, open a draft email in your own Outlook for you to
> send (Frazaro never sends mail itself). A phrasebook
> can define new sentences, and a phrasebook rule marked `raw` can contain
> arbitrary VBA; load phrasebooks from people you trust, the same way you
> would open a macro-enabled workbook from them. `docs/THREAT_MODEL.md` says
> this at length and without flattery.

That last paragraph is `SIG.0`'s data-handling statement, written to what is
*true today* (`THREAT_MODEL.md` §1.4, §2: no trust layers, `raw` ungoverned)
rather than to what `SEC.1`/`SEC.2` will make true. When they ship, the
paragraph gets shorter, not longer.

### 4.8 What a lawyer should read, and what they need not

Two files carry bespoke language: `OUTPUT-EXCEPTION.md` and
`PHRASEBOOK-TERMS.md`. Everything else is verbatim standard text or a
policy statement. An hour of counsel on those two, once, before the first
paid pilot — not before tonight's upload, because an unlicensed beta is a
worse state than a beta with a well-formed exception awaiting review.

---

## 5. THE CASE AGAINST THIS DOCUMENT

**The split plan schedules C-activity.** `REBUILD.md` §7's own test: is the
new structure being leaned on by someone shipping grammar rules, this month?
If the answer after the beta is "no one is shipping anything," the correct
response is `PI.1`, and §3's thirteen-to-nineteen sessions are the most
comfortable possible way to avoid finding that out. The order in §3.2 is
chosen so the first three cuts (A, B1, C) each pay a security or product
item directly; if those land and nothing downstream needs cut D, stop there.

**MPL on a phrasebook is a novelty, and novelties cost explanation.** Every
org will ask the question once. The counter is that the answer is four
sentences and the product's own `override:` design makes it obvious; the
honest residual is that a very conservative buyer may still prefer Bucket 1's
"Apache on everything," and if the first pilot's counsel says so, Bucket 1
is a one-line change to `REUSE.toml` and a re-header of nine `.vla` files —
*in that direction*, permissive-ward, which the sole copyright holder can
always make.

**The output exception is drafted by an engineer.** It says what the project
means; whether it says it in words a court would read the same way is §4.8's
hour. The GCC and Bison exceptions exist because the FSF's lawyers found the
inference insufficient; this project should assume the same.

**Trademark without registration is a policy, not a right.** It is still
worth writing tonight, because the alternative is a fork named "Frazaro Pro"
appearing before the paperwork does, and `VENTURE.md` §8 ranks the name as
the actual commercial boundary in open-core.

**And the whole of §4 rests on there being someone to license it to.**
`VENTURE.md` §9: users, zero. A licence is a door; this document has spent
its evening choosing the hinges. The beta uploads tonight so that someone
can walk through it.

---

## APPENDIX A — CURRENT → TARGET, REVISED

`REBUILD.md` Appendix A, re-measured on 2026-09-04.

| Today | Lines | `REBUILD.md` said | Measured, this document says |
|---|---|---|---|
| `VLA.bas` | 5,378 | 8 modules + `prelude.vla` | `VLA_Forms`, `VLA_FormulaEmitter` (free); `VLA_CompileState` (record); `VLA_Reader`, `VLA_Expander`, `VLA_VbaEmitter`; `VLA.bas` keeps the API |
| `VLA_SentenceEngine.bas` | 8,162 | 4 modules | `VLA_GrammarStore` (record); `VLA_EnglishRenderer`, `VLA_Phrasebooks`, `VLA_Sentences`, `VLA_Matcher`, `VLA_PhrasebookAuthoring`; the front door stays |
| `VLA_English.bas` | 303 | `VLA_EnglishWords` | already is; rename only under cut F |
| `VLA_IDE.bas` | 3,211 | 4 modules | `VLA_Commands`, `VLA_CommandLine`, `VLA_Workspace`, `VLA_AddinRegistration`; `VLA_Panel` is what remains |
| `VLA_Interpreter.bas` | 3,524 | *(new)* | `VLA_MemberDispatch` out; the rest stays as a documented R4 exception |
| `VLA_Runtime.bas` | 1,810 | whole | whole; 0BSD; `VLAX_` prefix under cut F |
| `VLA_Sql` / `_Datalog` / `_Prolog` / `_Relation` | 9,508 | *(absent)* | unchanged; split by banner when next open |
| `VLA_Tests*` | 16,693 | 8 `VLAT_` modules | `VLAT_` prefix + derived manifest, one pass; further splits by banner |
| `VLA_Loader.bas` | 145 | merged away | still present; merge with cut B (`include` is the reader's) |
| `VLA_Build.bas` | 855 | `VLA_AddinBuild` | rename + derived manifest under cut F |

## APPENDIX B — THE NUMBERS, AND HOW TO RE-RUN THEM

All counts from a Git Bash scan of `src/` at `d82b32c`, 2026-09-04.

| Measure | Count |
|---|---|
| `.bas` lines total | 54,005 (was 16,662 when `REBUILD.md` was written) |
| Modules over R4's 1,200 lines | 14 of 24 |
| Module-level variables, `VLA_SentenceEngine.bas` | 76; **56** touched from two or more of its 22 banner sections; `mPatItems` from 9 |
| Module-level variables, `VLA.bas` | 33; 23 cross-section; datum helpers and `deflambda` touch 0 |
| Module-level variables, `VLA_Interpreter.bas` | 13; 12 cross-section; all control signals |
| Module-level variables, `VLA_IDE.bas` | 2 |
| Module-level variables, `VLA_Sql` / `_Datalog` / `_Prolog` | 0 / 0 / 0 |
| Private procedure names already defined in another module | `VLA.bas` 13 · `VLA_SentenceEngine` 8 · `VLA_Interpreter` 8 · `VLA_IDE` 0 · `VLA_Sql` 2 |
| Hand-maintained module manifests | still 2 (`VLA_Build.bas:371`, `VLA_DevRig.bas:156`), plus `DEPLOY.md`'s prose list |
| Line endings, every file in `src/` at `HEAD` | LF |
| Git authors | 1, across 226 commits |
| `LICENSE` files | 0 |
| `raw` sites in `english.vla` | 7 |
| Qualified cross-module references, `VLA_SentenceEngine` → | `VLA_Messages` 98 · `VLA` 71 · `VLA_Identity` 62 · `VLA_IDE` 10 |
| Qualified cross-module references, `VLA_Interpreter` → | `VLA_Messages` 62 · `VLA` 42 · `VLA_Runtime` 40 · `VLA_Identity` 29 |

The section-coupling scan: for each module, list module-level declarations
outside any procedure; for each procedure, record which of those names its
body mentions; assign procedures to sections by the banner lines listed in
§3.2; report variables whose sections number two or more. Twenty lines of
Perl, re-runnable in a minute, and the right first step before any cut in §3
is attempted, since the banners move.
