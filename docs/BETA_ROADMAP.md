# FRAZARO — THE ROADMAP

*Open work only. Every item here is ⬜ open, 🟡 in progress, 🔒 gated,
⛔ parked or 🛡️ accepted, cut to one short summary paragraph and filed
under the department accountable for it. The day an item becomes ✅
(owner-tested AND committed) it is moved, verbatim, to
[`BETA_REARVIEW.md`](BETA_REARVIEW.md), the rear-view mirror: that file
carries the mission argument, the ordering doctrine, the full reasoning
attached to every standing decision, every closed item's build record, and
the expanded scoping behind each item below. When an item here is too thin
to act on, read its ID there. (History: forked from the master roadmap on
2026-08-28 as a second file; trimmed 2026-09-08; renamed to this file and
emptied of closed items on 2026-09-27, when the master roadmap became
`BETA_REARVIEW.md`.)*

*Item markers: ⬜ open · 🟡 in progress (built, awaiting owner test + commit) ·
🔒 gated · ⛔ parked/vetoed · 🛡️ accepted risk (assessed, deliberately not fixed,
with the mitigating control named and a stated condition that reopens it —
kept here because that condition is still being watched). ✅ appears here
only for as long as the commit that moves it to the rear-view.*

> **CONSTRAINT, this version:** a human at a Windows box (throughput) and the
> absence of any user but the owner (learning). Work that is at neither is, by
> Goldratt's definition, waste this version. Re-name the constraint at every
> version open; if it has not changed in three versions, that is itself a
> finding.

> **BETA, defined, because four gates in this file point at it:** *one named
> person outside the project runs their own SOP, on their own machine, on a
> Monday, with the owner unreachable.* Owner: the owner. Until that sentence is
> true, "gates on first external user" is a deferral with good manners.

---

## THE DEPARTMENTS

*Sections below are grouped by accountable department — a filing cabinet, not a queue. Re-bet from the whole document at every version open (cost-of-delay ÷ appetite, subject to 🔒 gates and SD-12's grammar floor).*

### 🔧 THE MACHINE — *"It works, it's correct, it's fast, everywhere."*
Fortifications · Environment · Assurance · Optimization · VLA · Performance
Reviewer profile: compiler engineer. Done means: pinned, measured, reproducible.

### 🗣 THE LANGUAGE — *"What can be said, and does it keep meaning what it meant?"*
Linguistics · Compatibility · Grammar
Reviewer profile: linguist / domain expert who has lived in spreadsheets.
Done means: auditioned, proven, and promised. **This department is the moat.**
Anyone can clone a compiler; nobody can clone a curated dialect plus a corpus
of tested phrasings plus the judgment that shaped them.

### 🪟 THE PRODUCT — *"Can a human find it, use it, install it?"*
Interface · Learnability · Accessibility · Distribution
Reviewer profile: designer. Done means: a stranger succeeded unaided.

### 🌍 THE COMMONS — *"Can others join, extend, and coexist?"*
Documentation · Governance · Interoperability
Reviewer profile: maintainer. Done means: someone else did it without asking.

### 🧑 THE PATIENT — *"Is anyone actually being helped, this week?"*
Pilot · SOP Import · Acquisition · Gap reporting
Reviewer profile: the user, who has not read this document and never will.
Done means: their SOP ran on their machine, on a Monday, without you in the room.

### 🛡 THE ADVERSARY — *"What can a hostile author, a hostile program, or a hostile input do?"*
Security
Reviewer profile: someone trying to break it — a profile no other department seats.
Done means: a threat model exists, the reachable object-model surface is default-deny,
and a phrasebook cannot act beyond it without visible, revocable consent.

### 🖋 THE SIGNATORY — *"What has to be true before someone with budget authority can say yes?"*
Procurement · Licensing · Positioning
Reviewer profile: the person who signs the form — the other reader no department here
ever seated. Done means: a licence, an IT-facing security summary, support terms, a
VPAT, and a one-page competitive positioning all exist as documents a stranger can
download without asking.

---

## THE STANDING-DECISION REGISTER

*Append-only; titles only here. Reasons, exceptions and drift notes for each are in `BETA_REARVIEW.md`.*

- **SD-1 — VBA is a backend, not the language.**
- **SD-2 — no refusal ships as a raw string.**
- **SD-3 — the audition checklist is append-only.**
- **SD-4 — a shipped spelling keeps its meaning.**
- **SD-5 — every backend supports or explicitly refuses every core form.**
- **SD-6 — the corpus family is a fixture.**
- **SD-7 — no grammar section is scheduled without a sentence that needs it.**
- **SD-8 — identifiers are folded invariantly, and identity never depends on a locale.**
- **SD-9 — IDs are never reused across documents, and a retired ID is never re-minted.**
- **SD-10 — semantics are never a pricing boundary.**
- **SD-11 — demolish an assumption when its benefit re-expresses as an optional feature; keep it when the benefit is structural.**
- **SD-12 — every version ships at least one grammar slice, whatever else is being paid down.**
- **SD-13 — no outbound network call, ever, without a dedicated re-litigation of this decision.**
- **SD-14 — `VLA_RELEASE_VERSION` is `MAJOR.MINOR.PATCH`, standard names, project-specific triggers.**
- **SD-15 — dynamic dispatch beyond the native, reviewed tier requires a declared capability; the default is deny.**
- **SD-16 — the phrasebook's pattern language is the only sentence grammar; no DCG, no backtracking parser, and no Prolog-shaped matcher ever parses a sentence.** *Precision amendment, 2026-09-11 (the owner's call, with `G-PROLOG`; not a re-litigation):* a slot consumes one token, one quoted string, one G6 list, or one phrase of a BUILT-IN sub-grammar — `expr`, `cond`, and `G-PROLOG`'s `conditions`, which is held to be regular — and, since `G-PROLOG` slice 2 (2026-09-14), its `clause` and `question`, both regular too. A phrasebook-defined nonterminal, and anaphora of any kind, stay forbidden. *(more: BETA_REARVIEW.md)*
- **SD-17 — blind spots are hunted on a cadence, not collected in a file: every roadmap fork is preceded by one outside-persona review, the persona must be one not yet used, and a review that mints or kills no roadmap item is decoration.**
- **SD-18 — VBA remains the reference; a second-host engine (web, desktop shell, or otherwise) follows the goldens, never leads them.**
- **SD-19 — where Excel has neighbouring operations, each sentence names the one it performs.**

---

# 🛡 ADVERSARY · SECURITY

*specimens: 1 — an audit finding, not a live exploit: this session's own read of `VLA_Interpreter.bas`/`VLA.bas` found `Application` plausibly reachable through the `CallByName` fallback with no capability gate, and `raw` splicing arbitrary VBA with no consent. Flagged for a live test, not yet confirmed as exploited — the same honesty this file already holds `AS.8`'s heuristics and `IN.11`'s parity claims to. This tranche exists because no department above ever seated a reader whose job was to look for this.*

- ⬜ **SEC.3 — phrasebook provenance for capability-requiring effects.** Which *layer* — base corpus, org phrasebook, community phrasebook — introduced a given `raw` splice or mail-send, attached to the emitted code itself, not just to a load-time log line. Extends `LISTOPS-PROVENANCE`'s existing `gen-row`/`at-row` tagging by one field rather than building new machinery: that mechanism already answers "which table row produced this rule"; the same answers "which trust layer authorized this effect." *Why now:* GO.3 already names the phrasebook registry "a supply chain" with no mechanism; this is the mechanism, and it turns "who do I hold accountable for this effect" from an investigation into a lookup. *Depends on:* SEC.1's tiers existing to have something worth attributing. `~days`

- ⬜ **SEC.6 — a security review before wide release.** Gated on SEC.0–SEC.2 landing and live-tested. *Expiry condition, G9-shaped, now met:* the interpreter's default-runtime dispatch surface is default-deny (SEC.1) and `raw` is gated (SEC.2), both committed — unlocked, not itself done. Not gated on SEC.7. `~days`–`~weeks`

- ⬜ **SEC.7 — Tier 1: permissioned, declared capabilities.** The small set of verbs with real external effect — `vlasendmail` today; file I/O once G-FILES ships — each meant to carry a `requires: capability:<name>` tag, checked at load, refused in words when absent. Scoped, not built: build a minimal general `F.10` first (`requires:` generalized from grammar-version dependency to permission, one shared parser rather than two); reuse SEC.2's own two-scope consent UX, not DI.1's blanket per-publisher trust; build the requires:/consent machinery generically, not single-purpose to `vlasendmail`; check at phrasebook load time, matching SEC.2's own chokepoint. *Depends on:* SEC.1, F.10. *(more: BETA_REARVIEW.md)* `~weeks`

*SEC.8–SEC.17 — the 2026-09-08 audit tranche, found by reading dispatch/build code against `THREAT_MODEL.md` (SEC.0) while scoping F.10. Audit findings, not confirmed live exploits (SEC.0's own framing); each names the file:line and the fix in BETA_REARVIEW.md. Ranked most-severe first.*

- ⬜ **SEC.10 — workbook-scoped `raw` consent is attacker-fillable.** SEC.2's `SEC2RawConsentV2 <hash>` (renamed by `SEC.11`) record lives in the host workbook's CustomDocumentProperties — which the attacker authored. Chained with SEC.9 a raw-bearing phrasebook loads with no dialog; the threat model assumed the workbook is the user's. *Fix:* keep workbook-scope consent device-side, keyed by (hash, workbook path), or drop the scope. *(more: BETA_REARVIEW.md)* `~hours`–`~days`
- 🛡️ **SEC.12 — the compile path emits a call to any unknown head; `english-function` accepts any target.** `EmitExpr`'s `Case Else` emits `SymName(h)(args)`, `EnglishAddFunctionWord` validates nothing, and Compile runs the module immediately — `(english-function "launch of" shell)` is `raw` without SEC.2's dialog. Compile-path only (needs VBOM trust); the interpreter refuses it (SEC.1). *Fix:* an emitted-callable allowlist — head-table forms plus the program's own `defun`s; unknown refuses in words. **🛡️ Accepted risk 2026-09-08:** Compile path only, and `RunProgram` refuses without `VlaHasVbProjectTrust()` - a setting Office ships OFF and managed IT routinely disables (that code's own words). Interpret touches no VBProject; SEC.1 already refuses both shapes at dispatch. Exposed population: someone who deliberately enabled the setting, uses Compile not Interpret, on a phrasebook they did not write. *Reopens if* Compile stops needing VBProject trust, or a shipped edition starts telling users to enable it. *(more: BETA_REARVIEW.md)* `~days`
- 🛡️ **SEC.14 — no step budget or cancel in the interpreter; unbounded reader recursion.** `while`/`until`/`repeat` have no cap, nothing touches `EnableCancelKey`, and `ParseForm` recurses per paren with no depth guard — while macro-expand (200), include (16), Prolog (120) and Datalog (10,000) all are capped. A hostile program hangs Excel. *Fix:* a step/wall-clock budget in the loop primitives, refused by name; a nesting cap in the reader. Adjacent to IN.14. **🛡️ Accepted risk 2026-09-08, with one half honestly uncovered:** nothing in src/ sets `Application.EnableCancelKey` (grepped), so it stays at VBA's default `xlInterrupt` and **Ctrl+Break interrupts a runaway loop**. That control does NOT cover `ParseForm`'s recursion, which is a VBA stack overflow - a hard crash Ctrl+Break cannot catch, and unsaved work elsewhere can be lost. Accepted because the outcome is availability and data loss, not compromise: no code runs, nothing leaves, nothing persists, and a workbook that crashes Excel on open is self-defeating as an attack. A robustness item wearing a security label. *Reopens with* IN.14, or if a crash is shown to leave exploitable state. *(more: BETA_REARVIEW.md)* `~days`
- ⬜ **SEC.15 — formula writes are an ungoverned egress channel.** SEC.4 guards only the `Value` sink; `set-formula` → `.Formula` accepts `WEBSERVICE`/`FILTERXML` (network via Excel), `HYPERLINK`, DDE, XLM `CALL`. A rule can plant one under an innocent sentence. *Fix:* refuse those function names in written formula text by name, and/or make formula writes a SEC.7 capability. *(more: BETA_REARVIEW.md)* `~days`
- 🛡️ **SEC.16 — trusted code loads from user-writable locations without integrity.** `prelude.vla`/`english.vla` (and the `.xlam`) beside the add-in in `%AppData%` override the embedded copies with no integrity check. *Fix:* make overrides opt-in with a visible indicator, or check them against an embedded hash list. **🛡️ Accepted risk 2026-09-08:** every route presupposes an attacker already running code as the user - a mailed workbook cannot write next to the add-in or into %AppData%. This is a persistence mechanism, not an entry point. The written fix also depends on SEC.11's digest, and hashing a file to defend against someone who can edit the hasher is close to circular; the honest defence is a non-user-writable install location, an installer decision. **Not** the same as SEC.9, where a mailed workbook's OWN directory wins the search order with no privilege at all - which is why SEC.9 stays open and severe. *Reopens if* a route is found that does not presuppose local code execution. *(more: BETA_REARVIEW.md)* `~days`
- 🛡️ **SEC.17 — emitted `Names.Add` with program-controlled `RefersTo`.** A defined name called `Auto_Open` with an XLM `RefersTo` is classic workbook persistence; compile-path only. *Fix:* refuse reserved names and non-formula `RefersTo`. **🛡️ Accepted risk 2026-09-08:** same control as SEC.12's, which is why they are accepted together - Compile path only, behind VBProject trust; the interpreter emits no `Names.Add` at all. Narrowest of the tranche and its `~hours` estimate is honest; accepted rather than done only because it would land a fix on a path no default-configured user can reach while SEC.9, reachable by anyone who opens a workbook, stays open. *Reopens with* SEC.12. *(more: BETA_REARVIEW.md)* `~hours`
- ⬜ **SEC.18 — the effect ledger: a declared effect class for every reachable form.** *Found while scoping the SEC.8–SEC.17 tranche; filed as substrate rather than folded into a consumer.* **Confirmed by absence:** nothing in `VLA_HeadTable.bas` or `VLA_Interpreter.bas` records whether a form can reach outside the sheet — SEC.1's tiers exist as prose here and as the shape of a hand-written `Select Case`, nowhere as data. Three open items are each about to derive the same table by hand: SEC.7 (which verbs need a capability), SEC.8 (what to refuse on an internet-zone workbook), SEC.15 (a formula write is network-class). **Build:** a ninth `effect` column on `AddRow` from a closed set (`none`/`host-ui`/`formula-write`/`file`/`mail`/`process`), a census over the three dispatching surfaces — 65 head-table rows, 128 core arms (via `check_emitter_coverage.ps1 -ListArms`, reused not re-parsed), 55 `VLA_Runtime` public procedures — and `tools/check_effect_ledger.ps1` in the house ratchet shape, baseline 0 unclassified, so a new form cannot reach a release unclassified. CO.6's pattern one axis over: that ledger dates a form, this one says what it can reach. **The irreversible half:** the effect vocabulary is an interface, since SEC.7's `requires-capability` names will draw from it and will then live in phrasebooks in the wild — choose it deliberately and record why. *Pays into:* SEC.7, SEC.8, SEC.15, and SEC.6 (the first artifact an external reviewer asks for). *Depends on:* nothing. *(more: BETA_REARVIEW.md)* `~days`

---

# 🖋 SIGNATORY · PROCUREMENT

*specimens: 0 — no outside procurement conversation has happened yet, which is exactly the finding: none of these six items require one to start. Every one is a document, not a feature, and every one is blocking not because it is hard but because it has never been written.*

- ⬜ **SIG.4 — the VPAT.** A public-sector or large-enterprise buyer will ask for one. 🔒 *Expiry:* AC.1 shipping — filing this before AC.1's own "two hours today" fix lands would just restate the gap it exists to close; this item's real dependency is that two-hour fix, open across multiple versions already, not new work of its own. `~days`, once AC.1 is done.

- ⬜ **SIG.5 — `0.5.0`'s own acceptance criteria, written down.** SD-14 defines what a version *number* means; nothing states what the specific `0.5.0` beta cutoff must satisfy to ship on its own targeted date — which tranches, which known-open bug classes are acceptable to ship with, what the release notes promise a downloader. *Why now:* a date with no definition of done is the one thing this file otherwise never permits — see SD-12's own floor, applied here to the release itself rather than to a single version's grammar work. `~hours`

- ⬜ **SIG.8 — the signing certificate as a published artifact, for a trusted-publisher deployment.** *Minted 2026-09-12 while closing SIG.1, which found the gap: the VBA project is signed (`CN=Frazaro VBA Signing`) and its public certificate is published nowhere, so one person can click "Trust all from publisher" but an IT department has nothing to deploy.* Microsoft's recommended macro baseline is "disable all except digitally signed macros" plus a required trusted publisher, whose certificate "needs to be installed as a Trusted Publisher on users' devices" — under that baseline Frazaro is undeployable today, whatever its engineering quality. Publish the `.cer` as a release asset, print its SHA-256 thumbprint in the notes and in `IT_REVIEW.md`, write the Group Policy/Intune steps, and state a rotation policy. Four questions it must answer rather than assume: certificate expiry and whether VBA signatures are timestamped (an untimestamped signature stops validating when the certificate expires, breaking installed copies on an unwatched date); the Windows-wide blast radius of a self-signed trusted publisher, which is where DI.1's deferred CA purchase gets re-argued; that this does **not** lift the Mark-of-the-Web block on a downloaded `.xlam`; and that every edition must carry the same publisher identity. *Depends on:* DI.1, SIG.1. *Expiry:* the first pilot whose IT policy requires it (`PI.7`). *(more: BETA_REARVIEW.md)* `~days`

---

# 🧑 PATIENT · THE PILOT

*specimens: 0 — which is the entire point of the tranche.*

- ⬜ **PI.0 — dogfood, honestly.** Has a real SOP *of the owner's own* been run through Frazaro, start to finish, on a workbook that matters? If yes, the specimen count is not zero and this tranche can be read at half strength. If no, it is one afternoon and it comes first. `~hours`

- ⬜ **PI.1 — name it.** One human, one SOP, one workbook, one date. Written down. `~hours`

- ⬜ **PI.2 — the transcript.** Their procedure in their words, verbatim, before any grammar work. This file — not `pareto.txt` — is the corpus's true north. `~days`

- ⬜ **PI.3 — the concierge run.** Every sentence the corpus refuses gets its VLA hand-written by the owner. The user gets a working automation on day one; the project gets the **gap log**. *Why this one matters most:* it is the only item in this document that acquires a specimen this month, and it converts "which sentences matter" from an argument into a measurement. `~days`

- ⬜ **PI.4 — the gap log ranks Grammar.** 🔒 SD-7 in force. `~hours`

- ⬜ **PI.5 — kill criteria, written BEFORE the run.** What observation would mean this product should not exist? Alpha 1's pair still stands: at least one program still clicked after a month of Mondays; at least one user modified a program alone. Both fail → the pre-chosen pivot. `~hours`

- ⬜ **PI.6 — the Monday test.** Their program runs unattended, on their machine, with the owner unreachable. **This is the beta definition, made concrete.** `~days`

- ⬜ **PI.7 — the trust reading.** During PI.1–PI.3, record what their IT actually permits: macro-enabled add-ins, VBProject trust, signed publishers, `.xlsm` in email. *Why here rather than in Distribution:* it is a five-minute question that ranks the whole interpreter effort, and it can only be asked of a real organization. `~hours`

---

# 🧑 PATIENT · SOP IMPORT

*The raw-material problem, not the sentence-grammar problem — SD-16 already settled how one sentence gets parsed, and this tranche's own owner-simplified design (this session) goes further than the pattern this file's other grammar-adjacent tooling usually takes: no triage heuristic of any kind runs before a line ever reaches the phrasebook — not even a sentence-blind one. A real SOP lives today in a Word document or a PDF, with inline screenshots and side-notes mixed into the numbered steps, and today there is no tooling between "raw document" and `PI.2`'s own hand-typed transcript at all. Named explicitly so it does not stay an implicit assumption inside `PI.2`/`PI.3`: this tranche is what makes THOSE items tractable for a real Word/PDF document instead of a pre-cleaned transcript; it does not replace them. specimens: 0 until PI.2 (LE.6/LE.7's own precedent — no specimen exists until a real transcript does).*

- ⬜ **SOP.2 — one line in, one row out, no classification at all.** Every Word paragraph or pasted line becomes exactly one row, in order; no guessing whether it is an instruction, a comment, or a header. **Rewritten 2026-09-18: blank lines are kept, with no exception.** A blank line is syntax: it ends every open block, and it separates `Otherwise:` and `If that fails:` from the block before them. Dropping blanks would silently re-nest a procedure. This is already today's import behavior; what remains is a pin test. Hazard kept on purpose: an empty Word paragraph used for spacing inside a block ends the block. The proposed Validate warning for that is an owner decision and outside this item. *(more: BETA_REARVIEW.md)* `~hours`

- ⬜ **SOP.3 — an embedded image becomes a `#`-prefixed comment row, in place.** During the same walk, an inline picture is emitted as its own row (`# [image1]`, `# [image2]`…), auto-numbered in encounter order. *(more: BETA_REARVIEW.md)*

- ⬜ **SOP.4 — Verify, then a hashtag, not a workbench.** Every imported row runs through the interpreter's own validate-and-refuse step exactly as authored; a failing row gets one guided fix — prepend `#` to make it a comment. No drafting, no AI bridge, no candidate generation. *(more: BETA_REARVIEW.md)*

- ⬜ **SOP.5 — the acceptance sentence, written before the run.** `PI.5`'s own kill-criteria discipline, pointed at this tranche specifically, decided now so success cannot be quietly redefined once something ships. Proposed: *a real SOP — a genuine Word document or PDF, containing at least one inline image and at least one side-comment with no `#` already on it — gets imported, run through Verify, and hand-marked or corrected line by line into a running Frazaro program, by the same named pilot user `PI.1` already committed to, without the owner hand-authoring any VBA outside the phrasebook.* If that sentence cannot be written truthfully after SOP.1–4 ship, the tranche is not done regardless of what got built. `~hours`

- ⛔ **SOP.7 — per-section languages.** Parked 2026-09-26 as a niche nice-to-have (the owner's call). Documents with more than one language want one language per run — an edition or a loaded phrasebook, already served, with SOP.6's language word refusing cleanly in the wrong edition — or "run only the section in my language", which is selection and now SOP.8's; as filed, it would run a bilingual SOP twice. Reopens if a pilot needs two languages in one run. `~weeks` *(more: BETA_REARVIEW.md)*

- ⬜ **SOP.8 — named sections, and running just the ones you pick.** Minted 2026-09-26 when SOP.7 was parked; scoped, not built. `<Frazaro "Accruals">` names a section, in the quoted slot SOP.6 reserved. Load still pours every section, each under a header row, and the choice is made at Run time by selection: **Run Selected Sections** runs the named sections the selection touches, in document order, as one Run with one Undo. Not a dialog after Load, which would be hidden state, and whose checkboxes would need a UserForm. Measure first: whether a section run alone silently reads a name only another section sets — if so, an engine item comes first. `~days` *(more: BETA_REARVIEW.md)*

---

# 🔧 MACHINE · FORTIFICATIONS

*Structure over discipline. Every boundary is currently held by habit. specimens: 3 (three blind-fix incidents, one first-user Undo report).*

- ⬜ **F.7 — the formula dialect as a declared subset.** One operator table, plus a test that fails when the VBA emitter gains a case the formula dialect neither supports nor refuses. *Now generalized by SD-5 to cover every backend.* Two latent bugs are its first cases, found 2026-09-27 while Contemplation 9's collapse 1 was scoped, and unfixed: `(- x)` emits `(x)`, because `EmitFormula`'s operator arm (`VLA.bas`) writes the operator only between operands, so a lone operand loses its sign. And a string's inner `"` is never doubled for Excel — `EmitFormula`'s and `FormulaQuote`'s string arms wrap the text as it stands — so a `deflambda` whose string holds a quote mark writes a formula Excel cannot read. `~days`

- ⬜ **F.16 — pre-flight runtime checks: literal arguments validated at Check.** Found live in G-FORMAT slice 2: `Format cell A1 as percent with 2.5 decimals.` passes Check, then refuses at run time — after the program's earlier sentences have already run. Pure runtime helpers (no workbook, no host object) are marked as such; Check calls them wherever their arguments are all literals and reports a refusal against the sentence's line, like `EnglishResolveCheck`. Values from variables or cells stay run-time refusals. First cases: decimals, colour codes, the freeze count, and slice 1's indent/rotation bounds (which today surface Excel's own message). Only marked-pure helpers may run at Check; `check_translate_purity.ps1` accepts exactly those. *(more: BETA_REARVIEW.md)* `~days`

- ⬜ **F.17 — a long program compiles: the emitter splits `main` before VBA's procedure-size limit.** Minted 2026-09-29, the owner's call, from `G-FORMULA` slice 1's live pass: Compile and Trace refused `instructions.txt` with VBA's own "Procedure too large" once its main body reached 755 steps (4,185 lines of generated VBA), where 731 (4,054) had compiled. Every top-level sentence goes into one `Sub main`, about five and a half lines each with step tracking, and VBA caps one procedure's compiled size, so a long SOP cannot be compiled at all, and the refusal is a VBE dialog, not Frazaro's words. Interpret and Run, the default, has no such limit. The corpus was split by hand (two sections became `To …:` steps) - discipline, where this item is structure: the emitter cuts a long `main` into parts called in order, each with its own step handler, cutting only between top-level sentences, moving `main`'s names to module level (as the interpreter already holds them), and stopping the run after a part that failed (`TER-11`'s case). *(more: BETA_REARVIEW.md)* `~days`

---

# 🗣🔧 LANGUAGE + MACHINE · THE TWO NEUTRALITIES

*The mission's section, in two co-equal parts and one appendix. Part A removes the assumption that the language is English. Part B removes the assumption that the machine is VBA. They are the same move — an assumption baked into a core that is still small enough to unbake it — and they are ranked together because delaying either one converts a chokepoint into a hundred sites. specimens: 0 for Part A, 0 for Part B until PI.7.*

## Part A — language-neutrality

**A2 — the catalogues (defer past beta)**

- ⬜ **LX.9 — surface-audition checklist as a lint.** *SD-3 banks the discipline; this is the tooling.* `~weeks`

- ⬜ **LX.12 — the product's own chrome is not covered by LX.2.** LX.2's catalogue covers refusal *text*; the ribbon captions, dialog text, sheet names (`Output`, `Trace`, `Known Sentences`), and column headers are hardcoded English with no equivalent mechanism. A Spanish-phrasebook user meeting an English ribbon is only half the mission delivered. *Why now:* named, not scoped — a real gap the falsification test (`LX.10`) never touched, because `LX.10` tested vocabulary, not chrome. `~days` to scope, `~weeks` to build — a second, UI-shaped catalogue, the same `VLA_Messages` shape `LX.2` already proved out.

## Part B — target-neutrality: the interpreter as the runtime

*Restored from `ALPHA1_ROADMAP.md` Phase F, where it was specified in full and then lost to a homograph (SD-9). Renumbered `IN.*` so it can never collide with F.1 again. Alpha 1's evaluator design is preserved verbatim; its **ordering** is not — Alpha 1 filed the interpreter as "a toggle beside the transpiler," and IN.9 below revises that to default-and-export, with reasons.*

**B2 — the evaluator (weeks; the runtime, so no longer contingent)**

- ⬜ **IN.4 — "Show me the VBA," and the export.** Two features under one ID; the first is shipped, the second is scoped and waiting on the owner's own design calls. **"Show me the VBA": done.** `EnglishIdeShowVba` (`VLA_IDE.bas`) reuses `VlaTranspile` - already pure text, already Mac-safe, no new compile machinery - and writes the result to a "Generated VBA" sheet (gridlines off, monospace, `NumberFormat "@"`), the same rendering shape `RenderTraceReport` already established. Deliberately shows the code WITH step-tracking instrumentation intact (unlike `InterpretProgram`'s own `EnglishStepTracking False` bracket) - an auditor should see the real thing Compile would inject, not a cleaned-up stand-in. *(more: BETA_REARVIEW.md)*

- ⬜ **IN.14 — cancel and progress in the interpreter's execution loop.** `IN.8`'s own numbers: a loop-shaped program can run 100×–2,700×+ slower than compiled VBA, worsening with N, and the interpreter has no `DoEvents` inside its statement loop — `IN.8`'s own benchmark run already froze Excel for roughly 100 seconds with zero warning before this was worked around by shrinking the *benchmark's* own sizes, not by fixing the interpreter itself. A real user's long-running program hangs Excel with no progress indication and no escape; the only recovery today is a task-kill, which a prior session's own memory records as capable of corrupting session state. *Why now:* this is a data-loss risk, not a performance nit — the interpreter is the default runtime (IN.9), so every user meets this loop eventually, not just an edge case. *(more: BETA_REARVIEW.md)*

- ⬜ **IN.16 — the interpreter cannot write a cell on another sheet by name.** Minted 2026-09-14 from `U.19`'s live pass. Under Interpret Instructions, `Put 1 into cell B2 of sheet Omega19.` was refused with SEC.1's "'range' is not in this interpreter's native dynamic-dispatch allowlist", though the phrasebook documents it as a one-sentence cross-sheet write. The phrase expands to `(set! (. (worksheets s) range r) e)`. The interpreter's write path (`WalkMemberSet`, then `DynamicSet`) carries only the value, never the `r` that `Range(r)` needs; reading the same cell works. It is the only phrasebook write of that shape, and no test pins it on either backend. The refusal's own text also exposes internals, against `LX.8`. Workaround: `Go to sheet Data.` then `Put 5 into cell B2.` Fix it within SEC.1's allowlist (no late-bound fallback) and keep SEC.4's formula guard on the write. Then pin the sentence under both backends and reword the refusal. *(more: BETA_REARVIEW.md)* `~hours`

## Part C — the dialect (the style guide)

- ⬜ **LX.7 — the Business English dialect spec.** The constrained sentence forms Frazaro accepts. *SD-3 accretes it; this is the write-up.* *Pays into:* Grammar (authors have a target), F.4, Learnability, Governance, **and the AI bridge — a model can be constrained to a published finite grammar and cannot be constrained to a mood.** `~weeks`

- ⬜ **LX.8 — the refusal-message style guide.** Name the problem, teach the fix, never blame, never expose internals. Refusals are the most-read prose in the product. `~days`

- ⬜ **LX.11 — the comment-syntax adjudication.** `#` (status quo) vs `[bracketed asides]` (Inform 7 precedent, multi-line free, unclosed brackets error loudly) vs `Note:` paragraphs. *Carried from Alpha 1's E1, where it was correctly filed as needing pilot parse-failure evidence.* `!` and `;` were explored and vetoed with reasons — recorded, not reopened. 🔒 *Expiry:* first pilot transcript; it is a **surface** decision, so it must land before contact or CO.1 inherits it. `~hours`

---

# 🔧 MACHINE · ENVIRONMENT

*The machine the user actually has, not the one you develop on. specimens: 1.*

- ⬜ **EN.1 — capability probe at load.** Excel version, LAMBDA/dynamic-array support, bitness, locale, VBProject trust — detected once, reported in words. **Revised by IN.9:** the probe no longer picks a runtime — there is one. It reports what the machine permits, which decides whether the *export* is available and what to say if it is not. `~days`

- ⬜ **EN.2 — formula locale policy.** `.Formula` vs `.FormulaLocal`. Choose one, document it, pin it. `~days`

- ⬜ **EN.4 — decimal and thousands separators in the reader.** A silently wrong number is the worst class of bug a spreadsheet tool can ship. `~days`

- ⬜ **EN.5 — date literal and date-format policy.** `~days`

- ⬜ **EN.6 — honest platform refusal.** Mac Excel has no VBProject object model. **Revised by IN.9:** Mac is no longer a refusal at all — it runs the default runtime and cannot export. EN.6 stops being an apology and becomes one honest sentence about a missing button, which is the first concrete dividend of SD-1. `~hours`

- ⬜ **EN.7 — the environment matrix.** Versions × locales × bitness × **backend**. `~weeks`

- ⬜ **EN.9 — `Workbook.Path` as a cloud URL breaks every `Dir$`-based file check.** A OneDrive Known-Folder-Move workbook reports `https://d.docs.live.net/...` from `.Path`; `Dir$` either crashed on it (fixed, `SafeFileExists`) or now silently reports "not found" for every candidate built from it. Found live during EDITION-MANIFEST's override check. The fix is scoped in the full entry. *(more: BETA_REARVIEW.md)*

---

# 🔧 MACHINE · ASSURANCE

*Coverage of the coverage. specimens: 0.*

- ⬜ **AS.3 — mutation testing.** Break the emitter in N known ways; assert the suite notices each. The only way to learn whether pins are load-bearing. `~weeks`

- ⬜ **AS.4 — property tests.** Expansion idempotence and "no program crashes the reader." The round-trip half shipped as G-RENDER's `EnglishRenderSelfCheck` (116/116 rendered, 115 round-tripped, the one failure a pre-existing corpus bug), so it is no longer a seed here. *(more: BETA_REARVIEW.md)*

- ⬜ **AS.5 — reader fuzzing.** Refuse in words, never crash, never hang. `~weeks`

- ⬜ **AS.9 — grammar coverage against real external text.** `AS.1` measures test density per rule; nothing measures what fraction of a *real* procedure Frazaro's corpus can express. A fixed sample of external SOP text (public procedure manuals, help-forum questions, a recorded-macro corpus) run through Check, with the refusal rate published alongside the existing test counts. *Why now:* the moat is described as "a curated dialect plus a corpus of tested phrasings"; at 122 rules that claim is currently unmeasured, and this is the instrument that would measure it, the same family as `AS.1`/`AS.2`/`AS.8`. `~days`

---

# 🗣 LANGUAGE · COMPATIBILITY

*Keeping promises to files you cannot see. specimens: 1 (the first-user Undo report — the standing override's founding incident).*

- ⬜ **CO.2 — the frozen compatibility corpus.** *SD-4 is the promise and is in force now; the file is bookkeeping.* `~days`

- ⬜ **CO.3 — a version stamp and a backend stamp inside generated modules**, so a support question is answerable from the workbook alone. Real scope, corrected while scoping SEC.7 (this item's own text used to restate "`requires:` in phrasebooks" as a second thing to build — it's F.10's own mechanism, cited here rather than duplicated): a version marker and a backend marker (interpreter vs. compiled emitter — no new scheme needed). **The version half is now answered rather than still a choice:** CO.4 (✅) settled it — the stamp's value is `VLA_RELEASE_VERSION` under SD-14's triggers, made orderable by `VlaVersionCompare`/`VlaVersionAtLeast`. This item stamps the value and names the backend; it neither picks a scheme nor writes a comparison. *Depends on:* F.10 (the `requires:` half), CO.4 (the ordering, already built). *(more: BETA_REARVIEW.md)* `~days`

- ⬜ **CO.5 — the migration tool.** Worthless until there is history to migrate, but CO.1 must exist first or there is nothing to migrate *to*. `~weeks`
- ⬜ **CO.7 — the shipped corpus, audited under SD-19.** Every shipped rule read against SD-19's test (could a speaker expect a sibling operation?); each failure gains an explicit sibling and becomes a documented legacy spelling (SD-4 keeps it working). Known candidates: `Add border to range` and `Clear color of cell` (both given siblings in G-FORMAT slice 1), `Format cell … as currency` and `… as date` (siblings `as dollars` and `as a short date` in slice 2), `Sort range … by column B1` (hidden header), `Keep only rows … where column 3 is …` (reads as deleting; position not letter) and `Show all rows` (reads as unhiding) — all three given siblings in G-SORTFILTER — and `Clear cell|range` (not yet). *Added by LE.6 (2026-09-18), a gap in specificity:* `Make cell|range … green|white|…` sets the **fill**, but reads just as naturally as the text colour. `Make cell B5 white.` erased a green result cell, and the same sentence turned four navy sample headers white. No colour-word sibling names what it colours (candidates: `Fill cell … green` / `Make the text of cell … green`). Separately, the colour words bind VBA's saturated constants (`vbGreen` = `#00FF00`). **Note, engine finding:** the refusals SD-19 creates teach the wrong rule — the near-miss ranking favours whichever rule an open slot (`{f:text-list}`, `{v:var}`) carried furthest, so `Remove the border from range …` suggests the pivot rule and `Set border-color of …` suggests `set {v:var} to {e:expr}`; live in G-SORTFILTER, `Sort range A1:C50 by column B1 with a header row.` expects `descending` and steers back to the legacy sort. Correct refusals, wrong lesson; the fix is LX.8's ranking, not this audit's. *Added by G-TEXT (2026-09-25):* `Set … to trimmed …` is VBA's `Trim`, the two ends only, where `Remove extra spaces from …` is Excel's TRIM; a reader who knows `=TRIM()` will expect `trimmed` to close up the inside too. *(more: BETA_REARVIEW.md)* `~days`

---

# 🗣 LANGUAGE · GRAMMAR

*The corpus. The moat, and the reason everything else exists. It has no position in this file, because under SD-12 it never stops. specimens: 0 until PI.2.*

- ⬜ **G-CONDFORMAT** — conditional formatting, `pareto.txt` §7 (12 entries). Split from G-FORMAT 2026-09-10: every entry but one needs a Tier-2 helper and pareto files it P1, so it is different work from G-FORMAT's one-property macros and would have held that item open. *(more: BETA_REARVIEW.md)* `~weeks`

- ⬜ **G-TABS** — sheets, `pareto.txt` §14: 18 entries, P0 (rename, copy, move before/after, add before/after, hide/unhide, exists, for each sheet, count, name, clear, tab colour). Several ship already (add, delete, go to, work on, protect/unprotect, tab colour). Waits on L-SHEET-HELPERS for the copy/move/add-after helpers. Part of the workhorse middle. `~weeks`

- 🟡 **G-FORMULA** — formulas and calculation, `pareto.txt` §11: 16 entries, P0 (a formula into a range or filled down, calculation mode, recalculate, sum/count/average-if, lookup, counts, median, standard deviation, subtotals). Sum, average, sum-where, count-matching and lookup ship already. **✅ Slice 1, built, owner-verified live and committed 2026-09-29:** a formula into a range and into rows counted as the program runs (adjusted row by row, measured; nothing written when there are no data rows), largest and smallest of a range, average, largest and smallest of a range into a cell, and a range's empty or filled cells counted - five rules, no runtime helper, no interpreter change. Average-where was cut from it: it needs AVERAGEIF in the interpreter, and SUMIF ÷ COUNTIF gives a different number when a matching row is blank. **✅ Slice 2, built, owner-verified live and committed 2026-09-30, for `0.8.0`:** a range's median and its standard deviation, each set and put into a cell, the reading named at the end (`as a sample` is STDEV.S, `as the population` STDEV.P; a sentence naming neither is refused), and average-where over AVERAGEIF - five new rules beside slice 1's (the ledger dates a rule by its pattern, so no alternation of slice 1's was widened) and four interpreter arms (`median`, `stdev_s`, `stdev_p`, `averageif`). A range with too few numbers stops the run at the sentence, as a lookup does; saying so in Frazaro's words is `U.27`. **Still open:** calculation mode and recalculate (slice 3, with the audit's `C62` restore as its precursor), `median of` a remembered range as a function word (a later slice, the owner's call), a put form for the where family, R1C1 and subtotals. Part of the workhorse middle. *(more: BETA_REARVIEW.md)* `~weeks`

- 🟡 **G-FILES — workbooks and files.** Scoped in `scripts/pareto.txt` section 15 (16 surfaces) - this file previously (wrongly) claimed zero templates existed for this section; a cross-check found six already shipped and already running through F.2's mechanism: `open-workbook`, `save-workbook` (as `save-current-workbook`), `save-as` (as `save-workbook-as`), `save-copy` (as `save-copy-as`), `close-workbook` (simplified - closes THIS workbook, not pareto's named-workbook form), `export-sheet-pdf` (as `export-sheet-as-pdf`). None needed the `{p:path}` slot §15's own preamble calls for - each takes `{p:expr}` bound to a pre-set variable rather than an inline quoted literal. *(more: BETA_REARVIEW.md)*

- 🔒 **G9** — conjoined predicates. STAYS GATED on pilot evidence. *The best-formed gate in this document: it names the observation that opens it.*

- ⬜ **G-TAIL** — charts, printing, validation, dates, email. Reordered by real sentence-gap reports, not by this file.

- **The query family — English sentence templates targeting each DSL engine's own text, not a fifth S-expression dialect.** Phrasebook rows are recognizers, not generators (settled in `SQL.1`); SQL's parse target stays open between canonical SQL text and the engine's AST until a real English→SQL row needs one. *(more: BETA_REARVIEW.md)*

  - ⬜ **G-SQL** — templates for `SQL`'s frozen subset. Engine complete (SQL.1–SQL.7); arguably the cheapest win in the family, since SQL's clause vocabulary already reads close to a business question. *(more: BETA_REARVIEW.md)*

  - ⬜ **G-RELATIONS — declared relation templates: a relation said in words, with its arguments anywhere.** Minted 2026-09-13 by `G-PROLOG` slice 2, the owner's call over an inferred multi-word reading, which is binary-only and whose guesses SD-4 would freeze. A relation is declared once (`The relation is: an invoice is submitted by a person.`) and clauses, conditions and questions are read against the finite declared set — Logical English's templates, SBVR's fact types, Prolog's `op/3`. To scope: where a declaration lives, the ambiguity audit over declarations, an undeclared relation refused at Check, the render reverse. Slice 5 reserves `starts with`, `ends with` and `contains`, which the audit must hold apart too. *(more: BETA_REARVIEW.md)* `~weeks`
  - ⬜ **G-DECISIONS — decision tables: a policy held as rows in a Table, with its hit policy named.** Minted 2026-09-14 by `G-PROLOG` slice 5's scoping, the owner's call over keeping the question deferred. DMN keeps a classification's rules as rows — input columns of conditions, an output column, and a hit policy (Unique, First, Priority, Any, Collect) for when several match — and in Excel a decision table is naturally a Table a business user edits. It sits beside slice 5's "otherwise" rather than replacing it; DMN practice prefers Unique, since First hides overlaps. To scope: a corpus entry that needs it (SD-7), how a header says which comparison its cells make, which hit policies and how Unique refuses an overlap, and what the rules compile to. *(more: BETA_REARVIEW.md)* `~weeks`

  - ⬜ **G-OPTIMIZE** — templates targeting `OPTIMIZE(rules, tables...)`; named for the function (renamed from `G-SOLVE` with the engine, 2026-09-18; `G-SOLVE` retired under SD-9). The grammar is the guardrail: its sentences may only write the tractable shapes `OPTIMIZE.0` names — cardinality words become native counters, "in a row" a precomputed pair relation. Nothing to recognize against until `OPTIMIZE.3` exists. *(more: BETA_REARVIEW.md)*

---

# 🪟 PRODUCT · PERFORMANCE

*Emitted-code speed AND authoring-time latency - both what the user experiences, distinct from OPTIMIZATION's compiler-internals speed. Owner-caught gap (2026-09-03): PF.1-PF.6 below only ever covered the first half - how fast a COMPILED program runs against a big workbook - never the second: how long Check/Compile/vocab-load itself takes for the person AUTHORING the SOP, felt on every edit-test cycle regardless of whether that SOP ever touches a 200k-row workbook. "Loading an SOP cannot take 5-10 seconds - that feels broken to any user" is exactly the product concern this section's own tagline already claimed to cover and didn't. PF.7-PF.9 are that missing half, found during P-TOK's own "what's next" pass (same session P-TOK closed a real, measured instance of this exact shape - compiler latency this time, not authoring latency, but the identical accidentally-quadratic-loop mechanism) - not a new tranche. specimens: 0.*

- ⬜ **PF.1 — the benchmark corpus.** A representative 200k-row workbook and a timing harness. Everything below cites it. `~days`

- ⬜ **PF.2 — automatic environment bracketing.** Screen-updating and calculation bracketed automatically. The single largest constant-factor win in VBA. `~days`

- ⬜ **PF.3 — never emit `.Select`.** Recorded macros are full of it; generated code has no excuse. *Pays into:* correctness too. `~hours`

- ⬜ **PF.5 — loop-invariant hoisting.** `~weeks`

- ⬜ **PF.6 — published numbers, pinned.** The performance claim becomes a test that fails when a change makes generated code slower. `~days`

- ⬜ **PF.8 — `RegisterVocabMacro`'s own O(macros²) string-build of `mVocabMacros`.** **Not** the finding **P-PROBE** already closed (✅, MACHINE · OPTIMIZATION) - verified against current code before filing, not assumed: P-PROBE's fix removed the re-TRANSPILE of the accumulated macro corpus on every registration (`VLA.VlaProbeMacroForm` now validates only the new macro's own text). This is a separate, still-live cost in the same function: `mVocabMacros = mVocabMacros & IIf(Len(mVocabMacros) > 0, vbCrLf, "") & macText` (`VLA_SentenceEngine.bas`) appends via `&` once per macro registered - the identical string-concatenation antipattern PF.7 names for `VlaEmbeddedText`, just accumulating vocabulary source text instead of sheet-cell text. Unmeasured - may already be dwarfed by other costs at today's corpus size (P-PROBE's own real-corpus number, 586 ms post-fix, doesn't obviously show it dominating), real risk as a vocabulary corpus keeps growing, same "flat-looking now, a cliff later" shape P-TOK's own opening quote already named once for `TokAt`. `~hours`

- ⬜ **PF.9 — `IdeLoadVocab`'s own multi-file embedded-sheet loop.** Lowest-confidence of the three - flagged for completeness, not a known live cost. `IdeLoadVocab` (`VLA_IDE.bas`) loops `n = 1 To 8` reading `VLAe_Source`, `VLAe_Source2`, ... via `VlaEmbeddedText` each iteration; the ceiling is small and fixed (its own comment: "no edition needs more than a couple of overlay files today"), so this loop itself is probably fine - its real cost, if any, is likely just PF.7's own per-call cost multiplied by however many chain files an edition actually carries, not a new quadratic shape of its own. Worth a look only once PF.7 lands and this gets re-measured; may close itself. `~hours`

---

# 🔧 MACHINE · THE MIDDLE LAYER

*The middle layer. Grows only as Grammar demands — never speculatively. The model tranche: demand-driven by construction. specimens: n/a.*

- ⬜ **L.11** — runtime `@doc:` annotations surfacing through apropos. `~days`

- ⬜ **L-SHEET-HELPERS** — tab create/copy/move/rename. Smallest surface, unblocks G-TABS. `~days`

- ⬜ **L-FILE-HELPERS** — CSV/text import. **Adjudicate the mechanism on the ledger first** (Workbooks.Open vs QueryTables vs line I/O). `~days`

- ⬜ **L-PIVOT-HELPERS** — pivot plumbing behind honest one-call verbs. Wants G6 first or the sentences are unwritable. `~weeks`

- ⬜ **L-TIER2** — classes and `defprop`. *Carried from Alpha 1's Phase C owner question, with its recommendation intact: hold the ordering.* Tier 2 changes `VlaTranspile`'s contract from text→text to source→*component set*, rippling through the compile path, the IDE, and the build, and it multiplies the injection surface D1 exists to harden. **Revised by the interpreter:** its payoff (events) is now reachable via IN.7 without any of that, which is an argument for IN.7 and against L-TIER2, not merely a reordering. 🔒 *Expiry:* a pilot sentence that needs a class and cannot be served by IN.7. `~weeks`

- ⬜ **L-TIER3** — self-hosting: port the phrasebook loader to VLA. *Reframed honestly on the Alpha 1 ledger:* VLA declines procedural macros, gensym, and eval, so this is dogfooding and middle-layer inspectability, **not** homoiconic flattening. `~weeks`

- ⛔ **DR2** — the reload consolidation, parked after DR1's crash.

---

# 🗣🔧 LANGUAGE + MACHINE · THE METAMETAMACRO LINE

*Vocabulary as computable, inspectable data - the F.13 claim ("english.vla is real VLA now") pushed as far as one session's worth of scoped passes could push it. Ordered by dependency, house style: each shipped item is what made the next one cheap, not just next on a list. The line's own constitution, adjudicated this session and binding on everything below it: determinism is the wall between this project and "the Wild West of LLMs" (the owner's own words) - `gensym` breaches it by inventing values from nothing and is vetoed outright; `car`/`cdr`/`cons`-style recursion over `quote` data does NOT breach it, provided it never reads anything but the literal, finite source text being compiled - the same wall, guarding a different door. *(more: BETA_REARVIEW.md)*

- ⛔ **GENSYM** — hygienic-macro identifier invention. **Vetoed by the owner, this session, in these words:** *"gensym is off the table, because determinism is the Chinese wall between this project and the Wild West of LLMs."* Not a style objection - a gensym'd name appears directly in emitted VBA with no trace to anything a person wrote, and its determinism across recompiles/reorderings cannot be guaranteed the way every other identifier in the pipeline already is. Confirms L-TIER3's own earlier "declines... gensym" rather than reopening it. `eval` was never proposed and stays off the table by the identical reasoning - runtime-arbitrary-code-execution is the same wall's other, more obvious breach. Kept here, permanently, precisely so a future session doesn't re-propose it without this context.

- ⬜ **DIALECT-REGEN** — regenerate `pirate.vla`/`latin.vla`/etc. mechanically from `english.vla`'s own loaded `mPatForms`/`mPatTexts` (real data already) plus a glossary table, rather than hand-translated as this session's seven files were. The seven hand-built files become the acceptance corpus for this generator, not artifacts to maintain by hand forever.

- ⬜ **METAPROOF** — proofs as forms, all the way down. `test-success` and `test-fail` are the `car` and `cdr` of testing: two primitives the loader runs, and every other proof shape a `defmacro` over them, written in the phrasebook. Minted 2026-09-24 from a brainstorming session, the same night as SPITBALLS' Lisp list; this is the one entry from it that earns an ID, because it retires a ratchet and pays into `AS.1` and `EDITION-PARITY` at once. First line, `~hours`: a third primitive, `(test-refuses-with-id sentence msg-id)`, so a proof can pin a refusal's *identity* instead of its English text — the thing `EDITION-MESSAGES` needs before a Spanish catalogue can exist without breaking `english.vla`'s own proofs. Then the macros, each one `~hours`: `test-same` (two spellings, one form: a metamorphic proof of paraphrase invariance, replacing every duplicated `in`/`into` pair), `test-every-surface` (enumerate a pattern's alternations and optionals — finite, because `SD-16` keeps them regular — and prove every branch, which closes `AS.1`'s "only one proof" gap by construction), `test-round-trip` (sentence → form → sentence through `G-RENDER`, a proof of the arrow and its inverse), and `test-parity` beside each query program, which retires `DatalogParityPrograms`, its scraper and `check_optimize_parity.ps1` the same day. Two catches, stated at minting: `requires-*` refuses when it arrives from a macro expansion, so proof macros must be dispatched on the other side of that scan or exempted by name; and generated proof names must be hand-derivable from the rule's own text under the `vla-` prefix — no gensym, per the veto above. Later, unscheduled: a proof-carrying certificate `(audited …)` emitted into `english_expanded.vla`, and per-rule goldens. Done: METAPROOF.1–3, in `BETA_REARVIEW.md`'s closed ledger. Open: METAPROOF.4–13, below, filed 2026-09-28. *(more: BETA_REARVIEW.md)* `~days`

  - ⬜ **METAPROOF.4 — `(tables ...)`: a proof hands its program table arguments.** The notation's next clause, decided now that tests needing it are next in line. About thirty DATALOG pins wait on it: keyed atoms, keyed queries, and number cells read from a Table-shaped array, in `TestDatalogKeyedAtoms` and in the unknown-predicate, ground, negated and text-test Subs. `check_proofs.ps1` learns the clause, and the clingo translation declines it until it can state it faithfully. `~days`
  - ⬜ **METAPROOF.5 — `(refuses id (slot value))`: a refusal pinned by what it names.** `VLA_Messages`' recorder keeps the last refusal's slots too, so a proof can say `(refuses datalog-unknown-predicate (predicate zz1))`: which predicate, variable or hint the message names, without pinning English, which is what `EDITION-MESSAGES` wants. About fifty DATALOG refusal pins that check their words wait on it. `check_proofs.ps1` checks that each slot is one the message's own text uses. `~days`
  - ⬜ **METAPROOF.6 — retire the parity table, its scraper and `check_optimize_parity.ps1`.** The parent's step 5, reached by another road: both parity loops read the proof corpus, so every test that moves shrinks `DatalogParityPrograms` (192 → 99 so far). After METAPROOF.4 and 5 it holds about thirty programs whose tests must stay in VBA — live workbooks, engine internals, the text entry point, a program VBA builds — and retiring it takes a decision about those. `~hours`, after 4 and 5
  - ⬜ **METAPROOF.7 — clingo checks more of the corpus.** Widen `tools/proofs_lib.ps1`'s translation, each construct with its argument and a witness: queries answered TRUE or FALSE (a fact, and one under `not`), comparisons over whole numbers, `let` arithmetic, `sum`, then quoted text. Quoted text first needs `count`'s case question settled, and is where clingo would have caught it. Forty-seven answer proofs are declined today. Also list only the counted variables in a `#count` tuple, as clingo noted at METAPROOF.3, and have the owner re-run the files that change. `~days`
  - ⬜ **METAPROOF.8 — PROLOG proofs on the same judge.** `test-prolog` beside `test-datalog`; `ProofAnswerVerdict` already knows nothing of DATALOG. About 704 call sites in `VLA_Tests_Query.bas`, the bulk of the whole item, taken a group at a time. The proof runner moves to a module of its own here, the day a second suite runs proofs. `~weeks`
  - ⬜ **METAPROOF.9 — SQL and OPTIMIZE proofs.** `test-sql` and `test-optimize` on the same judge: 56 and 46 call sites. `~days`
  - ⬜ **METAPROOF.10 — `(test-refuses-with-id sentence msg-id)` in phrasebooks.** The parent's step 1. `test-fail` pins an English substring, so `english.vla`'s own proofs would break the day a Spanish catalogue loads. Nothing blocks it, since METAPROOF.1's recorder is its substrate, and `EDITION-MESSAGES` needs it first. `~hours`
  - ⬜ **METAPROOF.11 — `test-same`: two spellings, one form.** The parent's step 2: paraphrase invariance stated once, replacing every duplicated `in`/`into` pair that today copies its expected form twice. `~hours`
  - ⬜ **METAPROOF.12 — `test-every-surface`: every branch of a pattern proven.** The parent's step 3: one `test-success` for each alternation and optional, with the expected form built from the rule's own template. It closes `AS.1`'s "exactly one test" gap by construction, and `check_rule_coverage.ps1` becomes a macro. `~hours`
  - ⬜ **METAPROOF.13 — `test-round-trip`: sentence → form → sentence.** The parent's step 4, through `G-RENDER`, equal up to optional words, so `G-RENDER`'s missing reverses become visible gaps. Needs `G-RENDER`. `~hours`

- ⬜ **VOCAB-MIGRATE** — version a generation spec, diff the rules it produces against what's currently loaded (VOCABDIFF), and report the delta - schema migrations for a spoken grammar. No known precedent to crib from; scoping this means designing close to first principles.

- ⬜ **PARETO-SPEC** — *correcting an overclaim made out loud this session, not proposing something new*: `pareto.txt` is prose with its own marker convention, not VLA - no generator can read it directly, ever, without a from-scratch parser for a format that was never meant to be parsed. The only real path is a human translating its rows into a VLA-syntax spec by hand, at which point it is a new artifact, not "pareto.txt, compiled." Filed to keep the corrected version on the record, not the first one.

- ⬜ **WORKBOOK-SPEC** — a generator reading its own row list from a live Excel range at build time. Thematically inevitable for a project this spreadsheet-native; collides directly with "the assistant writes VBA blind, no Excel available," and - stated plainly, not discovered later

- ⬜ **LISPIMPORT** — catalogue, not a chokepoint: once QUASIQUOTE and LISTOPS both exist and have real use behind them, survey further Lisp primitives worth porting against friction the shipped items above actually hit, not against "Lisps have this." `apply` (expand-time variadic calls) is the one candidate on the table; `gensym` is not - see the veto above - and `eval` never was.

---

# 🗣🔧🪟 LANGUAGE + MACHINE + PRODUCT · THE EDITION LINE

*One download per language — the owner's own call, this session, and it resolves more than it costs. Excel's ribbon (`customUI` XML) is embedded per file and is not runtime-swappable in practice, so chrome was always going to need a build-time answer regardless of what the rest of this line did; making every per-language surface build-time-selected rather than runtime-dispatched also retires a real VBA wall that would otherwise have blocked it outright — two `Public` procedures of the same name cannot coexist in one compiled project (`VLA_DevRig.bas`'s own `VlaAuditDuplicates` exists because this already happened once, by accident, on this very project), so an edition-per-download design is what makes a literal `VLA_Spanish.bas` *possible*, not merely simpler. Ordered by dependency, house style: each item below is the instrument the next one is built with, not just next on a list. The owner's own founding principle, on the record, this session: *"the language in which a person thinks should be the language in which a person works."* Measured here as the distance still standing between LX.5's proven grammar seam and a Spanish speaker's whole session — vocabulary, refusals, and the ribbon they read — staying in Spanish. specimens: 0 — no live non-English user has asked yet, the same standing this line's own `LX.10` proof fixture had before it was built; this is instrument-building on principle, the way that item was, not a response to a demand. *(more: this session's own transcript — CONTEMPLATIONS.md's Contemplation 3 and REBUILD.md's Layer 3, both written independently of this line and both landing on the same shape: vocabulary as tables an axis owns, not code duplicated per language.)*

- 🟡 **EDITION-VOCABPATH** — Built 2026-09-07, awaiting owner test: `IdeVocabPath` gained `includePolyglotta`, adding `scripts\polyglotta\<edition file>` as a fifth candidate for the three ribbon callers that need a real file (Translate to VLA/VBA, Export Expanded Phrasebook, Coverage), while `IdeLoadVocab`'s pre-check keeps the default so EDITIONMANIFEST.5's regression cannot reopen. *Why it jumped the queue:* it was the only thing that could regenerate the stale `english_expanded.vla` the release ratchet reads. *(more: BETA_REARVIEW.md)* `~hours`

- ⬜ **EDITION-MODULETAG** — `OUT_MODULE`/`RT_MODULE` hardcode `EN` into every compiled program's module name regardless of edition, so a workbook compiled by both editions collides — exactly the bug IO6.0's tag exists to prevent. Make the tag per-edition. Found during EDITION-MANIFEST's scoping, not guessed. *(more: BETA_REARVIEW.md)*

- ⬜ **EDITION-PARITY** — the verification discipline, built before any edition has real content to verify: a built edition's emitted VBA/VLA for an equivalent program must match the English reference byte-for-byte, `SD-18`'s own "follows the goldens, never leads them" applied to a second *language* edition instead of a second host, plus each edition's own language-specific `test-success`/message proofs (`TestLx10NonEnglishFixture`'s own shape, generalized). *Why before content:* pouring vocabulary/message/chrome work into an edition with no parity check is exactly how `VLA_Build.bas`/`VLA_DevRig.bas` drifted — a canary from day one is cheaper than a recount later. *Depends on:* EDITION-MANIFEST. `~days`

- ⬜ **EDITION-VOCAB** — extend LX5.1's `keyword-alias` directive shape to the vocabulary helpers still hardcoded as VBA `Select Case` in `VLA_English.bas` (`NumberWord`, `OrdinalWord`, `IsColorWord`, `IsNoiseWord`, `SkipArticles`, `ExprOpWord`, …) so each is table-driven from a loaded file — one compiled engine however many editions exist. Two categories first, to prove the mechanism generic (G6.0's precedent). *(more: BETA_REARVIEW.md)*

- 🟡 **EDITION-MACROS** — the `.vla`-side `defmacro` layer an author writes rule bodies in: `espanol.vla` had Spanish patterns over English macro-call bodies, so a Spanish speaker could read a sentence but not extend the corpus. Owner directive: a thin Spanish-named wrapper macro over every english.vla action macro. Verification correction on record: the dev-rig `expand` REPL seeds only `prelude.vla`, so phrasebook macros are proven through a real compile, not the REPL. *(more: BETA_REARVIEW.md)*

- ⬜ **EDITION-MESSAGES** — activate `VLA_Messages.bas`'s dormant seam: swap `AddEntries`' bodies for a Spanish catalogue with zero changes at the ~250 `RaiseMsg` call sites. Error text is the most-read prose in the product (LX.8), which is why this comes before full-coverage translation. *(more: BETA_REARVIEW.md)*

- ⬜ **EDITION-CHROME** — `LX.12`, executed: give ribbon captions, dialog text, sheet names (`Output`, `Trace`, `Known Sentences`), and column headers the same id-plus-table treatment EDITION-MESSAGES just gave refusals. The one item in this line with zero existing seam today (checked directly, not assumed: `VLA_IDE.bas`'s `"Trace"`/`"Known Sentences"` are bare literals with no indirection at all) — real new infrastructure, not an extension, which is why it sits after two items that already established the pattern it gets to copy rather than invent. `~weeks`

- ⬜ **EDITION-ESPANOL** — the dissection the rack above was built for: author real Spanish content across vocabulary, messages, and chrome using every instrument in this line, build `Frazaro_Espanol.xlam` for real through EDITION-MANIFEST, verify it through EDITION-PARITY, and take `DEPLOY.md`/`SIG.1`/`SIG.5` through their own first per-edition pass — Distribution's and Procurement's own departments, not just Linguistics's. `espanol.vla` already exists, already has real (if partial) coverage, and is the reason Spanish rather than any other language is first. The proof that the whole rack was worth building, the same role LX.10's fixture played for the grammar seam. *Depends on:* every item above. `~weeks`

---

# 🔧 MACHINE · OPTIMIZATION

*Compiler speed. Under the hood; the user never sees it directly. specimens: 0.*

- ⬜ **P-CONS.1** — exploratory: does a real cons-cell chain even win on VBA's own substrate? Filed from the owner's question during P-NTH ("why not go all the way"), answered honestly rather than deferred: it means the whole list representation and hundreds of construction sites, weeks not days — so measure first, and build only on a number. *(more: BETA_REARVIEW.md)*

- ⬜ **P-PRELUDE** — *largely subsumed by F.3.* `~hours`

- ~~**P-BULK** — array-slab runtime helpers.~~ Retired (2026-09-03), absorbed into **PF.4** (PRODUCT · PERFORMANCE, above) - confirmed, not just suspected: this tranche is compiler-internals speed the user never sees, but "array-slab runtime helpers" is code a *compiled program* runs, squarely PF.4's own "emitted-code speed" domain. Wrong tranche, not a second item.

---

# 🔧 MACHINE · QUERY AND LOGIC

*Aspirational, not demand-driven — the opposite contract from THE MIDDLE LAYER above. specimens: 1 (owner's own long-standing want, stated live 2026-08-28 — "you have no idea how long I've wanted both a SQL function and a PROLOG function inside Excel" — not corpus demand, so this tranche stays thin and appetite-boxed rather than hydrated). Buy the decision now; defer the artifact.*

- **`SQL(table, query)`** — shipped; its entry is in BETA_REARVIEW.md. Open items:
  - ⬜ **SQL.8 — `LEFT`/`RIGHT`/`FULL OUTER JOIN`, `CROSS JOIN`, and table aliasing (`AS`) for self-joins.** Both gaps already refuse by name (`sql-outer-join-not-supported`, `sql-duplicate-join-table`). Not unknown territory: `DATALOG`'s `not` proves the substrate can express an anti-join, which is the outer-join half. *(more: BETA_REARVIEW.md)*
  - ⬜ **SQL.9 — `NULL`, `IS NULL`/`IS NOT NULL`, and three-valued `WHERE`/`HAVING` logic.** An undocumented gap, not merely an unbuilt feature: a blank cell flows through `RangeToRows` as a raw `Empty` with whatever VBA's coercion happens to do, not a chosen behavior. Real SQL's `NULL` is foundational — comparisons yield `UNKNOWN`, aggregates skip it. *(more: BETA_REARVIEW.md)*
  - ⬜ **SQL.10 — `LIKE`, `IN (...)`, `BETWEEN`, and `CASE WHEN`.** Four ordinary conveniences with no presence in the grammar; today each falls into the generic `sql-unsupported-keyword` bucket. New cases at `ParseComparison`'s own precedence level; `LIKE` as a small anchored-match routine, never a regex engine. *(more: BETA_REARVIEW.md)*
  - ⬜ **SQL.11 — subqueries: scalar, `IN (SELECT ...)`, and `EXISTS`/`NOT EXISTS`.** The last structural gap against ordinary relational SQL, not a ceiling item — a subquery is a single non-recursive `SELECT`, so the termination guarantee is untouched. Scalar reuses the pipeline wholesale; `IN`/`EXISTS` are the semi-join and anti-join `RelJoin`'s hash-join law already extends to. *(more: BETA_REARVIEW.md)*

- ⬜ **`PROLOG(knowledgebase, query)`** — `=PROLOG(clauses, tables…)`: real unification and backtracking over structured facts and rules; free-form English-to-query is a phrasebook problem by the owner's own call, never this engine's. **Shipped, PROLOG.1–PROLOG.6:** `VLA_Unify.bas` one-way match (PROLOG.1) and two-way unification with the occurs check refused by name (PROLOG.2); ground facts and conjunctive queries (PROLOG.3); `(rule …)` forms with SLD resolution and backtracking (PROLOG.4); `is`/arithmetic, negation-as-failure, `findall`, cut (PROLOG.5.1–5.4); table-sourced and named-column facts (PROLOG.6). A 120-step ceiling refuses an infinite rule by name. **Open, below.** *(more: BETA_REARVIEW.md)*
  - ⬜ **PROLOG.16 — STANDARD ORDER OF TERMS: `@<`/`@=<`/`@>`/`@>=`, `compare/3`, `msort`/`sort`, `setof`/`bagof`.** `PROLOG.7` gave numeric comparison, `PROLOG.8` structural equality; there is still no way to ORDER two terms that are not both numbers, so a program cannot sort what it derived. All of it depends on one decision — a total order over var/number/atom/compound which, for an auditable engine, must be **deterministic and documented** rather than inherited from whatever `StrComp` does under the machine's locale (`VLA_Identity` already draws that distinction). `setof` is `findall` plus sort-and-dedup, nearly free once the order exists; `bagof`'s free-variable grouping is the hard half and may be worth refusing by name rather than approximating. Blocked by `PROLOG.13` for anything returning a list. `~week`. *(more: BETA_REARVIEW.md)*
  - ⬜ **PROLOG.25 — the Name/Arity half of ISO's existence_error: a table-backed goal with the wrong number of arguments.** A table's arity is never recorded and `UnifyArgsOnly` walks the goal's positions, so — read from the code, not yet seen live — fewer arguments than columns match every row on a prefix, silently, and more come back as VBA's raw "Subscript out of range". Measure live first, then likely seed the arity check from each table's column count. `~days` *(more: BETA_REARVIEW.md)*
  - ⬜ **PROLOG.26 — refusing reserved names on a branch the data never reaches stay silent.** `PROLOG.22` refuses an undefined name statically, but `(atom X)`, `(write X)` and the other always-refused names still refuse only when reached, so `(if (p X) (q X) (atom X))` carries a misspelling until the data changes. Decide whether `PROLOG.22`'s walk refuses them too, and what that makes of their solve-time arms and rule C. `~days` *(more: BETA_REARVIEW.md)*
  - ⬜ **PROLOG.27 — `true`, `fail`, `false`: ISO's control constants.** None exist; `(fail)` worked as a silent fail only by accident until `PROLOG.22` made it refuse as undefined. Reserve and solve them, reserve and refuse pointing at `(= a b)`, or leave them free. `~hours` *(more: BETA_REARVIEW.md)*
  - ⬜ **PROLOG.29 — first-argument indexing: stop spending a unit of work per row to ask one keyed question.** Minted 2026-09-12 out of the question `PROLOG.28`'s live pass raised - an Excel power user's Tables run past 100,000 rows, so a budget of 100,000 reads as arbitrary. It is a budget on SEARCHING, not on size (a scan spends one try per row, a join one per pair), and the refusal now says so; it cannot honestly become a TIME budget, since the same formula would then answer on one machine and refuse on a slower one. The fix is a smaller cost, not a bigger budget: index a predicate's clauses by their first argument when the goal binds it, and a keyed question over a 10,000-row Table costs about one unit instead of 10,000. Left the stated-ceiling list to become this item. `~days` *(more: BETA_REARVIEW.md)*

- **`DATALOG(tables, rules)`** — shipped; its entry is in BETA_REARVIEW.md. Open items:
  - ⬜ **DATALOG.6 — `min`/`max`/`avg` aggregation, alongside `count`/`sum`.** DATALOG.2 built exactly `BI_COUNT`/`BI_SUM`; `SQL` next door already has all five through `VLA_Relation`'s shared accumulator. Real, scoped work rather than a switch flip: Datalog's rule-body aggregate forms use their own small walk over derived tuples, not SQL's row-major `GROUP BY`. *(more: BETA_REARVIEW.md)*
  - ⬜ **DATALOG.7 — a refusal names the column a writer left out, never an internal variable.** Minted 2026-09-13 by `G-PROLOG` slice 1's review of generated names. Under `(not ...)`, a keyed atom's unmentioned column is refused as the unsafe variable `VlaAnonB<item>C<column>` — a name no writer typed, which renames itself when a Table's columns are reordered. It reaches no code or cell, only this message, so it is an `LX.8` defect rather than a breach of the `GENSYM` veto. Say which table's column was never mentioned, and how to fix it. *(more: BETA_REARVIEW.md)* `~hours`
  - ⬜ **Avoiding a full re-parse/re-fixpoint on every recalc — profiled first, not yet built.** Named precisely: `DATALOG` declares no `Application.Volatile`, so Excel's dependency graph already skips it unless an argument cell changed; the real, narrower cost is that a legitimate rerun redoes the *entire* fixpoint when one row changed. A cache keyed on the rules text plus each table's address needs a cheap "has this table changed" signal before it is worth building. *(more: BETA_REARVIEW.md)*

- ⬜ **`OPTIMIZE(rules, tables...)`** — the fourth engine, and the only one that produces a decision: `DATALOG` plus a choice, a constraint and an objective, in s-expressions like the other three (the owner, 2026-09-18). **Named `OPTIMIZE` 2026-09-18, the owner's call** — it names what the user wants rather than what the engine does, every manager already says it, its cognates carry it into the phrasebook languages, no worksheet function begins with `OPT`, and it positions the engine against Solver, where its value is; `OPTIMISE` is a one-line alias. `SOLVE` was refused as naming no task, `CHOOSE` is Excel's own, `ASSIGN`/`ARRANGE`/`ALLOCATE` name the mechanism, `ALLOT` is unclaimed and unused by any manager. `SOLVE.1`–`SOLVE.9` and `G-SOLVE` are retired under SD-9, and every other mention of the old name in the repository was renamed with the engine (the owner, 2026-09-18: semantic drift should not calcify in a months-old project). **Four standing decisions:** the answer is the decision, spilled with headers, searched once and read many times by Excel's own functions or a `DATALOG` question over the spill (a spilled range's first row now counts as headers — substrate, approved), with a content-keyed memo making "once" literal; **determinism is the hill** — same inputs, same answer, on every machine, budgets counted in work not seconds, every answer the best under a stated order (the objective, then fewest changes from the kept schedule, then Table order), 80% of answer-set programming from 20% of its machinery; negation stays stratified, choice the only source of alternatives; `VLA_Optimize.bas` new, substrate in `VLA_Relation.bas`, `optimize-*` ids in `VLA_Messages.bas`. **The nine dissections at commit `8368111` are replaced, on arithmetic:** they enumerated C(5,2)^7 = 10^7 worlds for a five-person toy (2.8 hours at a millisecond each), 10^128 for twenty staff over a fortnight — naive search never finishes at real size, `PROLOG.28`'s cliff a hundred orders of magnitude sooner. The engine now grounds once and searches with propagation from its first search item, every narrowing is scheduled and gated by a measurement, and conflict learning is an item with a trigger, not a ceiling. No feature here ships proven only at toy size. *(more: BETA_REARVIEW.md)*
  - ⬜ **OPTIMIZE.4 — `count` and `sum` as native constraints.** At least 2 a shift, at most 5 a week, no more than 40 hours — counters and weighted sums over chosen atoms, propagated as bounds, never expanded: the family's largest trap, closed. *(more: BETA_REARVIEW.md)* `~days`–`~weeks`
  - ⬜ **OPTIMIZE.5 — rules over chosen atoms.** `(assign S P)` chosen, `(overtime P)` following by rule, read by a constraint or the objective — grounded where every body atom is possible and given to the search as definitional clauses, never a fixpoint per candidate. The 80/20 cut: no recursion through a choice in this version, refused by name — recursion over the DATA is unaffected, being settled before any search. **The surface word stays `rule` and recursion arrives beneath it later (the owner, 2026-09-19), with the stable-model semantics fixed now so support is a loosening that changes no existing answer; it wants a §17 sentence first (`SD-7`): choosing links so every office is connected.** *(more: BETA_REARVIEW.md)* `~weeks`
  - ⬜ **OPTIMIZE.6 — the objective: branch-and-bound, and "best found" against "proven best".** `(minimize term)` over a weighted sum; partial worlds no better than the best are abandoned; the first solution is the first bound; values tried in the objective's order. Two results never confused, the name's own honesty rule. Where `OPTIMIZE.9`'s trigger is expected to fire. *(more: BETA_REARVIEW.md)* `~weeks`
  - ⬜ **OPTIMIZE.7 — the kept schedule: fewest changes, stability across edits, a warm start that cannot change the answer.** The item managers should fall in love with: `(keep Table)` names the published roster as an explicit input, adds one-per-difference behind the objective, tries kept values first. The warm-start trap named: only explicit inputs steer the search, so answers hold across sessions and versions. The command form's natural loop, and the ground for a large-neighbourhood mode past the ladder if wanted. *(more: BETA_REARVIEW.md)* `~days`
  - ⬜ **OPTIMIZE.8 — the structural narrowers, each measured:** symmetry breaking, component decomposition, root probing, dynamic fail-first — ~~none changes an answer~~ **three change no answer; dynamic fail-first would, since `OPTIMIZE.3`'s fork 6 made the Tables' own order decide which schedule a program with no objective gets, and it is out as written** — each switchable, each kept on its own ladder number. *(more: BETA_REARVIEW.md)* `~weeks`
  - ⬜ **OPTIMIZE.9 — conflict learning:** learned clauses, backjumping, watched literals, deterministic restarts — the old ceiling, now built when the ladder shows propagation stalling short of the reference roster, all integers and iterative, every answer identical to `OPTIMIZE.3`'s. *(more: BETA_REARVIEW.md)* `~weeks`
  - ⬜ **OPTIMIZE.10 — how many, and show me a few,** within a budget, last, since with determinism the first answer is the answer. *(more: BETA_REARVIEW.md)* `~days`

---

# 🪟 PRODUCT · INTERFACE

*The panel. The thing a user actually touches. specimens: 4 (owner catches, S3).*

- ⬜ **U.15 — lazy vocabulary self-heal** after state wipes. Small, removes a daily papercut. `~hours`

- ⬜ **V.1 — `verify:` rows.** The users' own harness: a program that checks itself. *Why high:* it extends the project's core philosophy — proof accompanies capability — from the compiler *to the user's programs*, which is the most Frazaro-shaped feature on this roadmap, and it is the difference between an automation that breaks silently and one that refuses in words on the morning someone upstream changes a column. `~weeks`

- ⬜ **U1 · U3 · U2 · U5** — the panel wave, in the owner's recorded order.

- ⬜ **U.12 — apropos in the panel** (three tiers plus worksheet functions).

- ⬜ **U.14 — `VlaTryTranspile`.** Retires the modal class from expected-error smokes. `~days`

- ⬜ **U.16 — the backend switch, in words.** Where the user sees which backend is running, why, and how to change it. *Depends on:* IN.4, IN.6, EN.1. `~days`

- ⬜ **U.13 · U.11 · U.6 · U.7 · U.8** — exploration rig and the remaining items.

- ⬜ **U.17 — snapshot-before-run on every path.** `TakeRunSnapshot` already exists and already runs for the workspace-sheet flow; the CLI (`IN.13`'s own documented gap: "no Undo snapshot, `TakeRunSnapshot` is keyed to a workspace sheet's own tag and a CLI command has none") and any future non-sheet entry point have no equivalent. A program that fails partway through — `IN.12`'s own `add-sheet-called` incident is the proof this is not hypothetical — should never leave a workbook in an unrecoverable, half-mutated state on *any* path. *Why now:* the cheap half of the transactional story; `V.1` (above) is the expensive half. `~days`

- ⬜ **U.18 — a per-run log, user-facing.** `IN.3.5`'s effect-log golden already exists but is a dev-side test artifact regenerated by `VlaWriteGoldens`, not a persisted, per-run record a user or auditor ever sees. The auditability pitch — "the procedure she wrote is the procedure the auditor reads," `docs/AUDIT.md` Part III — currently has no artifact behind it: no timestamp, no user, no program hash, no effect list survives a run. *Why now:* the single feature that turns "auditable automation" from a slogan into something a compliance reviewer can actually inspect, and the machinery it needs — the effect log — already exists; this is wiring, not invention. *(more: BETA_REARVIEW.md)*

- ⬜ **U.24 — a ribbon button runs the add-in's own code, whatever workbook is active.** Minted 2026-09-19 from the `0.6.1` Uninstall fix. Every Frazaro ribbon button calls `onAction="VlaRibbonAction"` with no workbook name (`VLA_Build.RibbonBtn`). Live, with `VLA.xlsm` active beside `Frazaro_Beta.xlam`, the add-in's Uninstall button ran the dev workbook's copy, so `ThisWorkbook` was `VLA.xlsm`. In `0.6.0` that deleted it. `0.6.1` guards the one destructive command (`VlaUninstallRefusal`); every other button still runs the dev copy while `VLA.xlsm` is active, and would run any open workbook's procedure of that name. Measure the rule first, in a standalone two-workbook repro. Then choose between a workbook-qualified `onAction` (which breaks renamed copies like `Frazaro_Beta.xlam` unless they are re-injected) and a guard in `VlaRibbonAction` that refuses in words outside an add-in. *(more: BETA_REARVIEW.md)* `~hours`–`~days`

- ⬜ **U.27 — when Excel's function has no answer, the stop says so in Frazaro's words.** Minted 2026-09-30, the owner's call, from `G-FORMULA` slice 2's scoping. Eleven WorksheetFunction members stand behind the phrasebook's figures: SUM, AVERAGE, MAX, MIN, COUNTIF, SUMIF and VLOOKUP, and from `0.8.0` MEDIAN, STDEV.S, STDEV.P and AVERAGEIF. Where a formula would show an error - a lookup that finds nothing, an average or median of no numbers, a sample of fewer than two, no matching row, an error value already in the range - the run stops at the sentence and puts the sheets back (`U.25`), but the words it shows are Excel's object model's: `Unable to get the VLookup property of the WorksheetFunction class`. One pure mapping, called where a stop's words are built (`VlaIdeStopMessage`, which both backends reach), gives all eleven Frazaro's words, naming the function and its likely reason. *(more: BETA_REARVIEW.md)* `~hours`

---

# 🪟 PRODUCT · LEARNABILITY

*Not documentation. Documentation answers "how do I do X." Learnability answers **"what can I say?"** — the actual first-contact problem. specimens: 0.*

- ⬜ **LE.3 — in-sheet autocomplete.** A constrained language is an autocompletable one. `~weeks`

- ⬜ **LE.4 — the first-run tutorial workbook.** Runnable, not readable. `~days`

- ⬜ **LE.5 — progressive disclosure.** A new user meets 40 verbs, not 340. `~days`

- ⬜ **LE.7 — the AI drafting bridge.** *Carried from Alpha 1's F2, and independently rediscovered by the audit's Part III.* "Describe what you want" → the phrase catalogue as the constraint → candidate sentences into column A → Check validates deterministically → red rows drive retry. *Why it is strategically large:* an LLM writing VBA produces code the person who asked cannot read, verify, or sign. An LLM writing Frazaro sentences produces output in a published finite grammar, refused mechanically when malformed, readable by the requester, testable by the corpus protocol, and failing as refusals with directions. **Frazaro becomes the safe target language for AI-generated spreadsheet automation** — the layer that makes a model's output auditable by the person who has to sign it. *Depends on:* LX.7, LE.1, F.4, LX.8. `~weeks`

- ⬜ **LE.8 — the phrasebook's own readability pass.** Five metasyntaxes is four too many: `=>` → `means`, `test:` → `example:`, `fail:` → `never:`, `{r:cell}` → `{a cell}`, `into|in` → `into (or in)`; rule rationale moves from `#` comments to a machine-visible `note:` line so LE.1 can show the best prose in the file, which is currently invisible to the product. `~days`

- ⬜ **LE.9 — split `kernel.vocab` out.** Macros, function mappings, and any rule that must touch a dot-form move to `kernel.vocab`; `english.vla` keeps the rest and contains **no dots at all**. The readability policy made physical: *if you are editing a file with parentheses in it, you are in the wrong file, and the filename says so.* *Depends on:* F.1. `~days`

- ⬜ **LE.10 — author the phrasebook in a workbook.** `english.vla` becomes a build artifact generated from a sheet — still the file of record, still diffable, still what CI checks, no longer what a human edits. The `example:` column is filled by the same act that writes the rule; the "Proven" column is live because the loader already refuses vocabularies whose tests fail; and **LE.1 is the read-only view of the identical table.** On-brand to the point of being slightly embarrassing that it is not already true. `~weeks`

---

# 🌍 COMMONS · INTEROPERABILITY

*Coexisting with everything already in the user's workbook. specimens: 0. **Revised by IN.9:** this tranche now defends the export path only. On the default runtime there is no module to collide and no project to lock, so what was infrastructure-wide risk is now scoped to a deliberate act — which is most of the reason IN.9 went the way it did.*

- ⬜ **IO.2 — locked/signed VBProject handling.** Refuse in words, name the reason — **or switch backends** (IN.6). `~days`

- ⬜ **IO.3 — existing macros, Power Query, connections** left demonstrably untouched. `~days`

- ⬜ **IO.4 — shared/co-authored workbook reality check.** `~days` **Three concrete findings this item inherits, from `docs/CONSULTANT.md`'s own audit, not yet acted on:** (1) the table macros (`table-add-row`/`table-delete-row`/`table-column`, and G-TABLES generally) resolve their target via `(activesheet.listobjects n)` — a program's meaning depends on which sheet happened to be active when it ran, a race condition by design, the exact class `PF.3` ("never emit `.Select`") already polices for emitted code but not for the grammar's own macros. (2) `IN.9`'s own flagship export pitch — email a self-contained `.xlsm` to a colleague — fights Windows' Mark-of-the-Web default (macro-blocked-by-default for internet-sourced files) on unmanaged home machines too, not only managed ones; the pitch needs a caveat or a companion "how to unblock this file" note, not a rewrite. *(more: BETA_REARVIEW.md)*

- ⬜ **IO.5 — other add-ins** in the same session. `~days`

---

# 🪟 PRODUCT · DISTRIBUTION

*Getting it onto a machine, and keeping it current. specimens: 0.*

- ⬜ **DI.4 — offline/air-gapped install.** Many finance environments are. `~days`

- ⬜ **DI.5 — licensing enforcement points**, if and where a commercial layer exists. 🔒 **SD-10 governs:** enforcement may gate seats, support, registry access, and hosted services — never a core form, a backend, or a grammar section. *The tempting inversion, named so it is not rediscovered:* module injection looks like an enterprise feature and is very nearly the opposite — enterprises block the trust it needs and quarantine the `.xlsm` it produces, while small unmanaged shops want exactly that file in exactly that email. `~weeks`

---

# 🪟 PRODUCT · ACCESSIBILITY

*Thin by design, with one item that is not optional. specimens: 0.*

- ⬜ **AC.1 — status must not be colour-only.** Roughly 1 in 12 men has a colour-vision deficiency; this is both an accessibility failure and, in some jurisdictions, a procurement blocker. **Two hours today, a UI refactor after the panel wave.** `~hours`

- ⬜ **AC.2 — keyboard-only operation.** `~days`

- ⬜ **AC.3 — screen-reader labelling.** `~days`

- ⬜ **AC.4 — font scaling and high contrast.** `~days`

---

# 🌍 COMMONS · DOCUMENTATION

*Deliberately deferred — with one carve-out. specimens: 0.*

- ⬜ **DO.1 — the dialect spec** *(lives in Linguistics as LX.7)*.

- ⬜ **DO.2 — the message catalogue** *(lives in Linguistics as LX.2; it is the error reference by construction)*.

- ⬜ **DO.3 — the user guide.** Post-alpha, as planned. `~weeks`

- ⬜ **DO.4 — the phrasebook authoring guide.** 🔒 *Expiry:* first outside contributor. `~weeks`

- ⬜ **DO.5 — whitepaper/onboarding refresh** at beta cutoff. `~days`

- ⬜ **DO.6 — the SOP cookbook.** *Acquisition instrument; belongs in the pilot bundle with LE.6.* `~weeks`

---

# 🌍 COMMONS · GOVERNANCE

*Thin on purpose. Extract only the decisions that are load-bearing on time. specimens: 0.*

- ⬜ **GO.2 — contribution ladder and review standards.** `~weeks`

- ⬜ **GO.3 — a community phrasebook registry and its trust model.** Vocabulary rules emit code; the registry is a supply chain. `~weeks`

- 🟡 **GO.4 — licensing split, CLA, trademark.** The split and the trademark policy landed with SIG.0 (2026-09-04; `docs/CUTS.md` §4.5–4.7): engine Apache-2.0, phrasebooks MPL-2.0, `TRADEMARK.md`. The CLA was deliberately *not* built — permissive inbound under the DCO (`CONTRIBUTING.md`) gives the owner every relicensing option a CLA would, with none of the paperwork, and `CUTS.md` §4.4 records why. Still open here: registering the mark, and the second repository for fair-source services (`FSL-1.1-ALv2`) once any exist. `~weeks`

- ⬜ **GO.5 — the escalation path** for a disputed surface. `~hours`

---

# 🪴 TERRARIUM · DEFERRED, NON-CRITICAL BUGS

*A parking lot, not a department - owner-created after L0.2's own live regression turned up two real bugs neither one was looking for. The point of writing them down here instead of chasing them on the spot: finding a bug while doing something else is not itself a mandate to fix it right now, and without a place to put it, "just this one small fix" is exactly how a regression run for one item turns into three. Deliberately outside F.12's ID namespace - these are found, not planned, and triaging them properly (real repro, root cause, a fix) is itself real work, owed its own turn rather than squeezed in as a rider on whatever else was already in flight. Picked up after the roadmap above is closed out, not before; nothing here blocks anything above it.*

- ⬜ **TER-1 — `Check Instructions` appears to re-run a prior `Compile and Trace`.** **Measured 2026-09-18 and deferred by the owner the same day**; left ⬜ because nothing is built. *As filed:* owner-found live 2026-08-27 during L0.2's regression - `Compile and Trace` run, then `Check Instructions` clicked right after, and it visibly re-ran; only the first click after a run does it. **The measurement** (dev workbook on the built add-in's ribbon, a five-row program whose Run adds 1 to a cell and stamps the time, two `Compile and Trace` runs each followed by two Validates, once with nothing clicked in between and once with the program's tab clicked first): the program's effects never happened twice (counter 2 after two Runs, one timestamp each), six commands produced exactly six translations with flat timings (vocabulary 430-469 ms), and `VLAt_TraceOn` read `=0` - so the filed armed-trace lead is dead, and `DoCheck` neither runs a step nor redraws the trace. It did not reproduce at five rows. The 2026-08-27 sighting was a few hundred rows, where the two surviving explanations become visible: Check marks column C row by row with screen updating ON (the same sweep a Run paints, since a Run Checks first), and a Run paints with updating OFF, so Excel can defer its redraw to the next command - one-shot by construction. Re-measure with `instructions.txt` before theorising. *(more: BETA_REARVIEW.md)*

- ⬜ **TER-3 — audit `instructions.txt`/`VerifyReportChecks` coverage in both directions.** Owner-created 2026-08-27, prompted by this session's G-STRUCT pass: building real `VerifyReportChecks` assertions for G-STRUCT's own new `GStruct` sheet caught a genuine argument-order bug in `make {a} look like {b}` (copied the wrong direction) that a `test-success` proof could never have caught, and the SAME pass found the `Demo` sheet's entire "prelude vocabulary" block (styling, rows/columns, sort/filter/dedupe on a small table, sheet protect/unprotect) has run on every regression pass since it was written but has never once had a single `Report`/`CheckV` assertion - a wrong constant or reversed argument there currently shows up as nothing worse than "the script didn't crash." *(more: BETA_REARVIEW.md)*

- ⬜ **TER-9 — the per-row and whole-program checks disagree about a sentence's final period.** Found in LE.6's live pass, 2026-09-18. `Show "… approval."` (the period inside the quotes, so the sentence has none) got a green OK on its row. The whole-program pass then refused it on the blank row *below*, with a message exposing the tokenizer's internal `|` marker. Nothing runs, so this is a teaching failure, not a safety one. Fix: one terminator rule for both passes, the refusal on the sentence's own row, in words. *(more: BETA_REARVIEW.md)* `~hours`

- ⬜ **TER-11 — in a compiled program, a failure inside one of its own actions does not stop the run.** Found 2026-09-26 building `U.25`, by reading the generated code, not yet seen live. Every procedure a program defines (`To stamp, …:`) gets its own step handler, which reports the failure and returns, so the caller goes on to its next sentence. The interpreter stops there (SD-5). The Run button still puts the sheets back, and its message claims nothing about what ran after, but a program run on its own shows the dialog and carries on. *(more: BETA_REARVIEW.md)* `~hours`

## THE SHORT ANSWER, IF YOU READ NOTHING ELSE

*Re-bet at every version open; the previous bet and its corrections are recorded in `BETA_REARVIEW.md`. SD-12's floor applies regardless: whatever else is on the list, one grammar slice ships.*
