# FRAZARO — THE MASTER ROADMAP (2)

*Forked from [`BETA_ROADMAP1.md`](BETA_ROADMAP1.md) on 2026-08-28. This file is the **quickly
scrollable list**: every item, open and closed, cut to one short summary
paragraph. Everything else — the mission argument, the ordering doctrine, the
full reasoning attached to every standing decision, every ✅ item's build
record, and the expanded scoping behind each item below — lives in
`BETA_ROADMAP1.md`, the massive reference file. When an item here is too thin
to act on, read its ID there. (Trimmed 2026-09-08: the query-and-logic, edition,
optimization, SOP, and PORT blocks, plus every item over ~1,500 characters, were
moved into BETA_ROADMAP1.md verbatim and summarized here — nothing was deleted,
only relocated.)*

*Item markers: ⬜ open · 🟡 in progress (built, awaiting owner test + commit) ·
🔒 gated · ⛔ parked/vetoed. An item becomes ✅ only after it is owner-tested AND
committed; ✅ items are deleted from this file at the next fork, not kept.*

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

*Append-only; titles only here. Reasons, exceptions and drift notes for each are in `BETA_ROADMAP1.md`.*

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
- **SD-16 — the phrasebook's pattern language is the only sentence grammar; no DCG, no backtracking parser, and no Prolog-shaped matcher ever parses a sentence.**
- **SD-17 — blind spots are hunted on a cadence, not collected in a file: every roadmap fork is preceded by one outside-persona review, the persona must be one not yet used, and a review that mints or kills no roadmap item is decoration.**
- **SD-18 — VBA remains the reference; a second-host engine (web, desktop shell, or otherwise) follows the goldens, never leads them.**

---

# 🛡 ADVERSARY · SECURITY

*specimens: 1 — an audit finding, not a live exploit: this session's own read of `VLA_Interpreter.bas`/`VLA.bas` found `Application` plausibly reachable through the `CallByName` fallback with no capability gate, and `raw` splicing arbitrary VBA with no consent. Flagged for a live test, not yet confirmed as exploited — the same honesty this file already holds `AS.8`'s heuristics and `IN.11`'s parity claims to. This tranche exists because no department above ever seated a reader whose job was to look for this.*

- ✅ **SEC.0 — the threat model, written down.** `docs/THREAT_MODEL.md`, owner-reviewed: what a program may reach, what a phrasebook may reach, who is trusted at each layer — and the finding that "zero-trust runtime" (IN.9's phrase) means "no Trust Center prompt," not "every capability verified." Written against the real dispatch code, not the roadmap's description of it. *(more: BETA_ROADMAP1.md)*

- ✅ **SEC.1 — capability gating of dynamic dispatch (Tier 0 + Tier 2 only).** Built and owner-verified live (`VLA_SELF-TESTS` pure 947/947, host 143/143; `VerifyReports` emitter 141/141, interpreter 141/141) and committed. The `CallByName` fallback in `DynamicGet`/`DynamicCall`/`DynamicSet` is removed from shipped builds; a 25-member census (not just G-TABLES — plain cell-`.value` reads and cross-sheet `.range` lookups too) was promoted to native `Select Case` arms first, closed by subtraction second. Tier 1 (permissioned/declared capabilities like `vlasendmail`) is scoped separately as `SEC.7`, not part of this item's own completion. *(more: BETA_ROADMAP1.md)*

- ✅ **SEC.2 — `raw` behind explicit, per-phrasebook consent.** Gated at `EnglishLoadVocabulary` (the file loader), not the shared host-free `EnglishLoadVocabularyText` — the correction a purity-ratchet check forced. Two chained dialogs; two remembered scopes (this workbook via a document property, this device via the registry), keyed by `EnglishSourceHash` so any edit re-prompts; no test-bypass toggle anywhere. Owner-verified live. *(more: BETA_ROADMAP1.md)*

- ⬜ **SEC.3 — phrasebook provenance for capability-requiring effects.** Which *layer* — base corpus, org phrasebook, community phrasebook — introduced a given `raw` splice or mail-send, attached to the emitted code itself, not just to a load-time log line. Extends `LISTOPS-PROVENANCE`'s existing `gen-row`/`at-row` tagging by one field rather than building new machinery: that mechanism already answers "which table row produced this rule"; the same answers "which trust layer authorized this effect." *Why now:* GO.3 already names the phrasebook registry "a supply chain" with no mechanism; this is the mechanism, and it turns "who do I hold accountable for this effect" from an investigation into a lookup. *Depends on:* SEC.1's tiers existing to have something worth attributing. `~days`

- ✅ **SEC.4 — formula/CSV-injection guard.** A program-written cell *value* beginning with `=`/`+`/`-`/`@` is neutralized with a leading apostrophe (`NeutralizeFormulaInjection`, the `Value` sink only). Owner-tested live, including the formula-bar check. Deliberately not applied to formula writes, which are a feature. *(more: BETA_ROADMAP1.md)*

- ✅ **SEC.5 — `SECURITY.md` and a disclosure contact.** Owner-reviewed and committed (`docs/SECURITY.md`) — contact address (`english@spreadsheet.company`) and response commitment (3-business-day acknowledgment, no fixed remediation deadline, severity-prioritized, reporter kept informed) both filled in by the owner rather than guessed. A place a finder is told to report to, and a stated response commitment — the minimum credible artifact before any external pilot (`PI.*`) or public download. *Honest tension, named rather than solved here:* `SD-13`'s no-network stance means a disclosed vulnerability has no push-update path to the people already running an affected build; this item does not resolve that, it only makes sure a report has somewhere to land. `~hours`

- ⬜ **SEC.6 — a security review before wide release.** Gated on SEC.0–SEC.2 landing and live-tested. *Expiry condition, G9-shaped, now met:* the interpreter's default-runtime dispatch surface is default-deny (SEC.1) and `raw` is gated (SEC.2), both committed — unlocked, not itself done. Not gated on SEC.7. `~days`–`~weeks`

- ⬜ **SEC.7 — Tier 1: permissioned, declared capabilities.** The small set of verbs with real external effect — `vlasendmail` today; file I/O once G-FILES ships — each meant to carry a `requires: capability:<name>` tag, checked at load, refused in words when absent. Scoped, not built: build a minimal general `F.10` first (`requires:` generalized from grammar-version dependency to permission, one shared parser rather than two); reuse SEC.2's own two-scope consent UX, not DI.1's blanket per-publisher trust; build the requires:/consent machinery generically, not single-purpose to `vlasendmail`; check at phrasebook load time, matching SEC.2's own chokepoint. *Depends on:* SEC.1, F.10. *(more: BETA_ROADMAP1.md)* `~weeks`

*SEC.8–SEC.17 — the 2026-09-08 audit tranche, found by reading dispatch/build code against `THREAT_MODEL.md` (SEC.0) while scoping F.10. Audit findings, not confirmed live exploits (SEC.0's own framing); each names the file:line and the fix in BETA_ROADMAP1.md. Ranked most-severe first.*

- ⬜ **SEC.8 — programs in cells bypass Office's macro policy (Mark-of-the-Web unread).** The architectural one. A plain `.xlsx` carries its program in cells; after "Enable Editing" the interpreter runs it with the add-in's privilege, reaching `Workbooks.Open`/`SaveAs`/`SaveCopyAs`, PDF export, and Outlook natively — no `Zone.Identifier`/Protected-View check anywhere in `src/`. Microsoft's 2022 default blocks internet macros; Frazaro reintroduces the capability. *Fix:* read the host workbook's zone, refuse external-effect forms on internet-zone workbooks until trusted. Partly absorbed by SEC.7; zone-awareness is not in its scope. *(more: BETA_ROADMAP1.md)* `~days`–`~weeks`
- ⬜ **SEC.9 — phrasebook search-order hijack + silent replay of workbook-stored paths.** `IdeVocabPath` prefers `<workbook dir>\scripts\english.vla` over the add-in's own copy (DLL-hijack with a `.vla`), and GO.6's `VLA_LoadedPhrasebooks` document property replays absolute/UNC paths with no prompt (UNC also leaks NTLM on Check). Ship a workbook with a grammar file beside it and every sentence means what the attacker says. *Fix:* never auto-load grammar from the workbook's own directory or its carried paths without a one-time consent naming the file; refuse UNC/URL. *(more: BETA_ROADMAP1.md)* `~days`
- ⬜ **SEC.10 — workbook-scoped `raw` consent is attacker-fillable.** SEC.2's `SEC2RawConsent <hash>` record lives in the host workbook's CustomDocumentProperties — which the attacker authored. Chained with SEC.9 a raw-bearing phrasebook loads with no dialog; the threat model assumed the workbook is the user's. *Fix:* keep workbook-scope consent device-side, keyed by (hash, workbook path), or drop the scope. *(more: BETA_ROADMAP1.md)* `~hours`–`~days`
- ⬜ **SEC.11 — the consent hash is a 32-bit polynomial.** `EnglishSourceHash` is `h*31+b mod 2³²` over non-whitespace bytes; second preimages are trivial (edit two comment bytes), so even device-scope consent transfers to a crafted phrasebook. *Fix:* SHA-256 via `System.Security.Cryptography` (COM-visible from VBA, no network); `check_rule_coverage.ps1`'s twin changes in step. *(more: BETA_ROADMAP1.md)* `~hours`
- ⬜ **SEC.12 — the compile path emits a call to any unknown head; `english-function` accepts any target.** `EmitExpr`'s `Case Else` emits `SymName(h)(args)`, `EnglishAddFunctionWord` validates nothing, and Compile runs the module immediately — `(english-function "launch of" shell)` is `raw` without SEC.2's dialog. Compile-path only (needs VBOM trust); the interpreter refuses it (SEC.1). *Fix:* an emitted-callable allowlist — head-table forms plus the program's own `defun`s; unknown refuses in words. *(more: BETA_ROADMAP1.md)* `~days`
- ⬜ **SEC.13 — Word automation opens untrusted documents with macros enabled.** **Reachable today**, not gated on SOP.1: *Import Program File…* (menu and ribbon, `EnglishIdeImport`) already accepts `.docx`/`.doc` and routes them through `ReadWordFile`, which never sets `AutomationSecurity` — so Word's automation default (*Low*) runs a `.docm`'s AutoOpen silently; `GetObject` also attaches to the user's live Word. SOP.1's own text asked for this check; the code shipped without it. *Fix (one line):* `wordApp.AutomationSecurity = 3` before `Documents.Open`, restored afterward. *(more: BETA_ROADMAP1.md)* `~hours`
- ⬜ **SEC.14 — no step budget or cancel in the interpreter; unbounded reader recursion.** `while`/`until`/`repeat` have no cap, nothing touches `EnableCancelKey`, and `ParseForm` recurses per paren with no depth guard — while macro-expand (200), include (16), Prolog (120) and Datalog (10,000) all are capped. A hostile program hangs Excel. *Fix:* a step/wall-clock budget in the loop primitives, refused by name; a nesting cap in the reader. Adjacent to IN.14. *(more: BETA_ROADMAP1.md)* `~days`
- ⬜ **SEC.15 — formula writes are an ungoverned egress channel.** SEC.4 guards only the `Value` sink; `set-formula` → `.Formula` accepts `WEBSERVICE`/`FILTERXML` (network via Excel), `HYPERLINK`, DDE, XLM `CALL`. A rule can plant one under an innocent sentence. *Fix:* refuse those function names in written formula text by name, and/or make formula writes a SEC.7 capability. *(more: BETA_ROADMAP1.md)* `~days`
- ⬜ **SEC.16 — trusted code loads from user-writable locations without integrity.** `prelude.vla`/`english.vla` (and the `.xlam`) beside the add-in in `%AppData%` override the embedded copies with no integrity check. *Fix:* make overrides opt-in with a visible indicator, or check them against an embedded hash list. *(more: BETA_ROADMAP1.md)* `~days`
- ⬜ **SEC.17 — emitted `Names.Add` with program-controlled `RefersTo`.** A defined name called `Auto_Open` with an XLM `RefersTo` is classic workbook persistence; compile-path only. *Fix:* refuse reserved names and non-formula `RefersTo`. *(more: BETA_ROADMAP1.md)* `~hours`
- ⬜ **SEC.18 — the effect ledger: a declared effect class for every reachable form.** *Found while scoping the SEC.8–SEC.17 tranche; filed as substrate rather than folded into a consumer.* **Confirmed by absence:** nothing in `VLA_HeadTable.bas` or `VLA_Interpreter.bas` records whether a form can reach outside the sheet — SEC.1's tiers exist as prose here and as the shape of a hand-written `Select Case`, nowhere as data. Three open items are each about to derive the same table by hand: SEC.7 (which verbs need a capability), SEC.8 (what to refuse on an internet-zone workbook), SEC.15 (a formula write is network-class). **Build:** a ninth `effect` column on `AddRow` from a closed set (`none`/`host-ui`/`formula-write`/`file`/`mail`/`process`), a census over the three dispatching surfaces — 65 head-table rows, 128 core arms (via `check_emitter_coverage.ps1 -ListArms`, reused not re-parsed), 55 `VLA_Runtime` public procedures — and `tools/check_effect_ledger.ps1` in the house ratchet shape, baseline 0 unclassified, so a new form cannot reach a release unclassified. CO.6's pattern one axis over: that ledger dates a form, this one says what it can reach. **The irreversible half:** the effect vocabulary is an interface, since SEC.7's `requires-capability` names will draw from it and will then live in phrasebooks in the wild — choose it deliberately and record why. *Pays into:* SEC.7, SEC.8, SEC.15, and SEC.6 (the first artifact an external reviewer asks for). *Depends on:* nothing. *(more: BETA_ROADMAP1.md)* `~days`

---

# 🖋 SIGNATORY · PROCUREMENT

*specimens: 0 — no outside procurement conversation has happened yet, which is exactly the finding: none of these six items require one to start. Every one is a document, not a feature, and every one is blocking not because it is hard but because it has never been written.*

- ✅ **SIG.0 — the licence, the warranty disclaimer, and a data-handling statement.** Owner-decided Bucket 2 (`docs/CUTS.md` §4): Apache-2.0 engine, MPL-2.0 phrasebooks (file-scoped, so an organization's own phrasebook is its own), 0BSD injectable runtime, CC-BY-4.0 docs; DCO, no CLA; trademark; holder Spreadsheet Company. `tools/check_spdx.ps1` enforces the map. Built and owner-verified 2026-09-04. *(more: BETA_ROADMAP1.md)*

- ⬜ **SIG.1 — the IT-facing security and architecture summary.** A two-page document a stranger can download: what Frazaro installs, what it can reach (SEC.0/SEC.1's own answer), what it sends over a network (SD-13's no-network stance, stated as a selling point, not an apology), how it updates, how it is removed (DI.2's uninstall button — built, live-tested, and worth citing here as evidence, not just in `DEPLOY.md`). *Why now:* this is the actual artifact that gets software approved inside a company, and nothing resembling it exists — `DEPLOY.md` is written for the developer, not the reviewer who has to sign the exception. *Depends on:* SEC.0. `~hours`, once SEC.0 exists.

- ⬜ **SIG.2 — the competitive positioning one-pager.** Frazaro against Copilot in Excel, Office Scripts, and Python in Excel — leading with the three properties none of those has: determinism (`GENSYM`/`eval` vetoed outright, the METAMETAMACRO LINE's own wall), offline/no-network (SD-13), and auditability (the procedure a human reads is the procedure that ran). *Why now:* the honest differentiators already exist and are simply not written down anywhere a procurement conversation could point to. *Pays into:* LE.7 — the AI bridge is the engineering answer to the same competitive question; this is the one page that exists before LE.7 ships. `~hours`

- ✅ **SIG.3 — support terms and an incident-response process.** `docs/SUPPORT.md` — where a user reports a bug, what response to expect, a named security contact: the ordinary-support half of SEC.5's `SECURITY.md`. Built and owner-confirmed. *(more: BETA_ROADMAP1.md)*

- ⬜ **SIG.4 — the VPAT.** A public-sector or large-enterprise buyer will ask for one. 🔒 *Expiry:* AC.1 shipping — filing this before AC.1's own "two hours today" fix lands would just restate the gap it exists to close; this item's real dependency is that two-hour fix, open across multiple versions already, not new work of its own. `~days`, once AC.1 is done.

- ⬜ **SIG.5 — `0.5.0`'s own acceptance criteria, written down.** SD-14 defines what a version *number* means; nothing states what the specific `0.5.0` beta cutoff must satisfy to ship on its own targeted date — which tranches, which known-open bug classes are acceptable to ship with, what the release notes promise a downloader. *Why now:* a date with no definition of done is the one thing this file otherwise never permits — see SD-12's own floor, applied here to the release itself rather than to a single version's grammar work. `~hours`

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

- ⬜ **SOP.1 — Word and PDF intake, via Word itself.** Both formats route through Word's object model over COM, never a hand-built parser; PDF through Word's own conversion. Two limits named up front: Word must be installed and automatable (refused by name otherwise), and a scanned PDF has no recoverable text. Plain-text paste stays the fallback. Two open questions for its scoper: whether automating Word is SD-15 territory, and — ADVERSARY's name-check — an opened document must never have its macros triggered. *(more: BETA_ROADMAP1.md)* `~days`–`~weeks`

- ⬜ **SOP.2 — one line in, one row out, no classification at all.** Every Word paragraph or pasted line becomes exactly one row, unmodified, in order; no guessing whether it is an instruction, a comment, or a header. *(more: BETA_ROADMAP1.md)*

- ⬜ **SOP.3 — an embedded image becomes a `#`-prefixed comment row, in place.** During the same walk, an inline picture is emitted as its own row (`# [image1]`, `# [image2]`…), auto-numbered in encounter order. *(more: BETA_ROADMAP1.md)*

- ⬜ **SOP.4 — Verify, then a hashtag, not a workbench.** Every imported row runs through the interpreter's own validate-and-refuse step exactly as authored; a failing row gets one guided fix — prepend `#` to make it a comment. No drafting, no AI bridge, no candidate generation. *(more: BETA_ROADMAP1.md)*

- ⬜ **SOP.5 — the acceptance sentence, written before the run.** `PI.5`'s own kill-criteria discipline, pointed at this tranche specifically, decided now so success cannot be quietly redefined once something ships. Proposed: *a real SOP — a genuine Word document or PDF, containing at least one inline image and at least one side-comment with no `#` already on it — gets imported, run through Verify, and hand-marked or corrected line by line into a running Frazaro program, by the same named pilot user `PI.1` already committed to, without the owner hand-authoring any VBA outside the phrasebook.* If that sentence cannot be written truthfully after SOP.1–4 ship, the tranche is not done regardless of what got built. `~hours`

---

# 🔧 MACHINE · FORTIFICATIONS

*Structure over discipline. Every boundary is currently held by habit. specimens: 3 (three blind-fix incidents, one first-user Undo report).*

- ⬜ **F.7 — the formula dialect as a declared subset.** One operator table, plus a test that fails when the VBA emitter gains a case the formula dialect neither supports nor refuses. *Now generalized by SD-5 to cover every backend.* `~days`

- ✅ **F.10 — `requires:` in phrasebooks.** `(requires-version "0.5.2")`, `(requires-capability …)`, `(requires-form …)`, checked by a lexical pre-pass before a single rule registers. The specified `requires: ns:value` spelling could not exist after F.13's all-forms migration; the namespace rides in the head, and a bare atom would have named a function. Shared parse, two enforcement sites: `version`/`form`/unknown refuse host-free in `EnglishLoadVocabularyText`; `capability` in `EnglishLoadVocabulary` (SEC.2's trap). Unknown namespace refuses, with version-first ordering. `capability` can only refuse until SEC.7; `form` refuses as not-yet-enforceable. Carried CO.6's ratchet, `check_grammar_since.ps1` (150/150, 128/128). Owner-verified live 2026-09-07 (pure 985/985). *(more: BETA_ROADMAP1.md)* `~days`

- ⬜ **F.15 — the two hardcoded module manifests, cross-checked.** `VLA_Build.bas`'s `mods` array and `VLA_DevRig.bas`'s reload list must agree and are not checked against each other — `REBUILD.md`'s own R3 names this, and `LX.2`'s own session was bitten by it live the same day (`VLA_Messages` missing from `VlaDevReload`'s own list, surfacing only as a "Variable not defined" compile error on a fresh workbook). *Why now:* cheap and mechanical, the exact shape of bug this project's own tooling already catches for other pairs (`F.12`'s ID checker, `AS.8`'s parity checker) — a `tools/*.ps1` script comparing the two lists is an afternoon, not a redesign, and the failure mode it prevents has already fired once live. `~hours`

---

# 🗣🔧 LANGUAGE + MACHINE · THE TWO NEUTRALITIES

*The mission's section, in two co-equal parts and one appendix. Part A removes the assumption that the language is English. Part B removes the assumption that the machine is VBA. They are the same move — an assumption baked into a core that is still small enough to unbake it — and they are ranked together because delaying either one converts a chokepoint into a hundred sites. specimens: 0 for Part A, 0 for Part B until PI.7.*

## Part A — language-neutrality

**A2 — the catalogues (defer past beta)**

- ✅ **LX.5 — split `VLA_English.bas`** into language-neutral sentence machinery and English-specific rules. *Why it mattered:* it decided whether `VLA_Spanish.bas` is a sibling or a fork — answer: sibling, proven live, not just argued. **Session finding, sharpened by a direct audit:** the structural remainder isn't just the four headline forms — `ParseStmt`/`ParseCond`/`ParseExpr`/`ParsePrim`/the Oxford-comma list logic (5 call sites) compare tokens against dozens of literal English words, and the existing head-table alias system (`VlaHeadTableAliasMap`/`ResolveHeadAlias`) turns out to be no help at all: it only resolves *emitted* VLA s-expression heads inside `VLA.bas`, never English sentence tokens.
  **Phase 1 (committed) — the canonicalization seam:** a new `CanonicalizeStructuralWords` pass (runs right after every `EnTokenize` call that feeds `ParseStmt`) rewrites any bare-word token a phrasebook has aliased via a new `(keyword-alias "si" "if")` directive to its canonical English spelling — same alias-map shape `VlaHeadTableAliasMap` already proved, staged one step earlier, so every existing comparison site is untouched and English behavior is provably unchanged (empty alias table = no-op). Oxford-comma-list grammar's own positional rule stays universal, fixed core syntax, per LX.1's own "keep, structural" ruling — only its trigger words go through the table.
  **Phase 2 (committed) — the physical split:** a direct variable-usage audit of all 216 procedures against the file's ~65 shared `Private` Collections found the split was NOT the "~2/3 neutral" first estimate — `mPatItems` alone was touched directly from 13 different regions spanning the whole file, so grammar registration, the DCG matcher, G-RENDER, statement/condition/expression parsing, vocabulary loading, and the front door itself (`EnglishToVla` drives the parse loop directly) are one connected, inseparable unit. **204 of 216 procedures moved** to a new `VLA_SentenceEngine.bas`; `VLA_English.bas` keeps only the twelve pure, stateless functions LX.1's audit already called "data" (`NumberWord`/`OrdinalWord`/`ExprOpWord`/`IsExprOpWord`/`IsNoiseWord`/`IsDroppedWord`/`SkipArticles`/`IsColorWord`/`SlotDesc`/`StrayCharHint`) plus Phase 1's own keyword-alias mechanism. Every moved procedure kept its exact name and visibility — VBA resolves an unqualified `Public` call project-wide regardless of which module hosts it, so no caller anywhere in the project needed to change. `VLA_Build.bas`/`VLA_DevRig.bas` manifests, `VLA_Tests_Host.bas`'s stray-helper scan, and the `check_raise_ratchet.ps1`/`check_translate_purity.ps1` module-name ratchets all updated and re-verified clean against the new file.
  **Phase 3 (committed) — the second-language proof:** `scripts/polyglotta/espanol.vla` gets five `(keyword-alias ...)` entries (si/es/repetir/veces/mientras) and three `test-success` proofs exercising real Spanish `if`/`repeat`/`while` sentences — the first non-English sentences in this project to reach `ParseStmt`'s own structural dispatch, not just the DCG phrase matcher. `TestLx10NonEnglishFixture` (`VLA_Tests_Grammar.bas`) validated these for free: `EnglishLoadVocabulary` raises on any `test-success` mismatch at load time, so the existing fixture running clean *was* the proof, not a weaker proxy for it — no new test needed. **Confirmed live:** `VlaSelfTests()` 903/903 pure + 119/119 host, `VlaGoldens()` diff empty, `TestLx10NonEnglishFixture` itself PASS. **`for each` deliberately excluded, found not guessed — the one real remainder:** `en`/`para` are already ordinary vocabulary elsewhere in `espanol.vla` ("... en la celda", "envia un correo PARA ..."), and `CanonicalizeStructuralWords` rewrites every occurrence of an aliased word globally, not just in structural position — aliasing either would break existing rules in the same file. Real fix (named, not built): canonicalize only when the *preceding* token already resolved to a structural word expecting a second one, not on every token unconditionally — a future item, now with a real, found reason rather than a hypothetical one. *(more: BETA_ROADMAP1.md)*

- ⬜ **LX.9 — surface-audition checklist as a lint.** *SD-3 banks the discipline; this is the tooling.* `~weeks`

- ⬜ **LX.12 — the product's own chrome is not covered by LX.2.** LX.2's catalogue covers refusal *text*; the ribbon captions, dialog text, sheet names (`Output`, `Trace`, `Known Sentences`), and column headers are hardcoded English with no equivalent mechanism. A Spanish-phrasebook user meeting an English ribbon is only half the mission delivered. *Why now:* named, not scoped — a real gap the falsification test (`LX.10`) never touched, because `LX.10` tested vocabulary, not chrome. `~days` to scope, `~weeks` to build — a second, UI-shaped catalogue, the same `VLA_Messages` shape `LX.2` already proved out.

## Part B — target-neutrality: the interpreter as the runtime

*Restored from `ALPHA1_ROADMAP.md` Phase F, where it was specified in full and then lost to a homograph (SD-9). Renumbered `IN.*` so it can never collide with F.1 again. Alpha 1's evaluator design is preserved verbatim; its **ordering** is not — Alpha 1 filed the interpreter as "a toggle beside the transpiler," and IN.9 below revises that to default-and-export, with reasons.*

**B2 — the evaluator (weeks; the runtime, so no longer contingent)**

- ⬜ **IN.4 — "Show me the VBA," and the export.** Two features under one ID; the first is shipped, the second is scoped and waiting on the owner's own design calls. **"Show me the VBA": done.** `EnglishIdeShowVba` (`VLA_IDE.bas`) reuses `VlaTranspile` - already pure text, already Mac-safe, no new compile machinery - and writes the result to a "Generated VBA" sheet (gridlines off, monospace, `NumberFormat "@"`), the same rendering shape `RenderTraceReport` already established. Deliberately shows the code WITH step-tracking instrumentation intact (unlike `InterpretProgram`'s own `EnglishStepTracking False` bracket) - an auditor should see the real thing Compile would inject, not a cleaned-up stand-in. *(more: BETA_ROADMAP1.md)*

- ⬜ **IN.14 — cancel and progress in the interpreter's execution loop.** `IN.8`'s own numbers: a loop-shaped program can run 100×–2,700×+ slower than compiled VBA, worsening with N, and the interpreter has no `DoEvents` inside its statement loop — `IN.8`'s own benchmark run already froze Excel for roughly 100 seconds with zero warning before this was worked around by shrinking the *benchmark's* own sizes, not by fixing the interpreter itself. A real user's long-running program hangs Excel with no progress indication and no escape; the only recovery today is a task-kill, which a prior session's own memory records as capable of corrupting session state. *Why now:* this is a data-loss risk, not a performance nit — the interpreter is the default runtime (IN.9), so every user meets this loop eventually, not just an edge case. *(more: BETA_ROADMAP1.md)*

**B3 — a second host, prepared (SD-18's own infrastructure; target-neutrality applied one substrate further than IN.\* ever needed to)**

- ✅ **PORT.1 — pure, host-free translate entry points.** `VLA_Browser.bas` (named for the client, not the network — nothing about it implies an SD-13-violating connection): `EnglishTranslateTextToVla`/`ToVba` take phrasebook text in memory, never touch a file, never show a dialog. *(more: BETA_ROADMAP1.md)*
- ✅ **PORT.2 — the translate-path purity ratchet.** `tools/check_translate_purity.ps1`: scans a named set of (module, function) pairs for forbidden host touches — body-only by design, so a callee is verified by reading, not by the exit code. *(more: BETA_ROADMAP1.md)*
- ✅ **PORT.3 — the intrinsics spec.** `docs/INTRINSICS.md`: the VBA-specific string/array/comparison behaviors a non-VBA port must match, each earned by a real citation. *(more: BETA_ROADMAP1.md)*
- *Together:* SD-18's infrastructure — a port with no purity guarantee and no intrinsics reference would re-litigate every behavior. *(more: BETA_ROADMAP1.md)*

## Part C — the dialect (the style guide)

- ⬜ **LX.7 — the Business English dialect spec.** The constrained sentence forms Frazaro accepts. *SD-3 accretes it; this is the write-up.* *Pays into:* Grammar (authors have a target), F.4, Learnability, Governance, **and the AI bridge — a model can be constrained to a published finite grammar and cannot be constrained to a mood.** `~weeks`

- ⬜ **LX.8 — the refusal-message style guide.** Name the problem, teach the fix, never blame, never expose internals. Refusals are the most-read prose in the product. `~days`

- ⬜ **LX.11 — the comment-syntax adjudication.** `#` (status quo) vs `[bracketed asides]` (Inform 7 precedent, multi-line free, unclosed brackets error loudly) vs `Note:` paragraphs. *Carried from Alpha 1's E1, where it was correctly filed as needing pilot parse-failure evidence.* `!` and `;` were explored and vetoed with reasons — recorded, not reopened. 🔒 *Expiry:* first pilot transcript; it is a **surface** decision, so it must land before contact or CO.1 inherits it. `~hours`

---

# 🔧 MACHINE · ENVIRONMENT

*The machine the user actually has, not the one you develop on. specimens: 1.*

- ⬜ **EN.1 — capability probe at load.** Excel version, LAMBDA/dynamic-array support, bitness, locale, VBProject trust — detected once, reported in words. **Revised by IN.9:** the probe no longer picks a runtime — there is one. It reports what the machine permits, which decides whether the *export* is available and what to say if it is not. `~days`

- ⬜ **EN.2 — formula locale policy.** `.Formula` vs `.FormulaLocal`. Choose one, document it, pin it. `~days`

- ⬜ **EN.3 — number-format locale policy.** Named formats resolve to locale-correct patterns. 🔒 **Blocks: G-FORMAT's number-format rules — about twelve of them.** Nothing in G-STRUCT, G-ROWLOOP, G-TEXT, or the rest of G-FORMAT needs a locale-correct currency pattern. *The previous revision read "before Grammar §6" and was taken as blocking seventy rules to protect twelve.* *Expiry:* first non-US-locale user. `~days`

- ⬜ **EN.4 — decimal and thousands separators in the reader.** A silently wrong number is the worst class of bug a spreadsheet tool can ship. `~days`

- ⬜ **EN.5 — date literal and date-format policy.** `~days`

- ⬜ **EN.6 — honest platform refusal.** Mac Excel has no VBProject object model. **Revised by IN.9:** Mac is no longer a refusal at all — it runs the default runtime and cannot export. EN.6 stops being an apology and becomes one honest sentence about a missing button, which is the first concrete dividend of SD-1. `~hours`

- ⬜ **EN.7 — the environment matrix.** Versions × locales × bitness × **backend**. `~weeks`

- ⬜ **EN.8 — 64-bit declaration discipline** (`PtrSafe`) as a lint rule. `~hours`

- ⬜ **EN.9 — `Workbook.Path` as a cloud URL breaks every `Dir$`-based file check.** A OneDrive Known-Folder-Move workbook reports `https://d.docs.live.net/...` from `.Path`; `Dir$` either crashed on it (fixed, `SafeFileExists`) or now silently reports "not found" for every candidate built from it. Found live during EDITION-MANIFEST's override check. The fix is scoped in the full entry. *(more: BETA_ROADMAP1.md)*

---

# 🔧 MACHINE · ASSURANCE

*Coverage of the coverage. specimens: 0.*

- ⬜ **AS.3 — mutation testing.** Break the emitter in N known ways; assert the suite notices each. The only way to learn whether pins are load-bearing. `~weeks`

- ⬜ **AS.4 — property tests.** Expansion idempotence and "no program crashes the reader." The round-trip half shipped as G-RENDER's `EnglishRenderSelfCheck` (116/116 rendered, 115 round-tripped, the one failure a pre-existing corpus bug), so it is no longer a seed here. *(more: BETA_ROADMAP1.md)*

- ⬜ **AS.5 — reader fuzzing.** Refuse in words, never crash, never hang. `~weeks`

- ⬜ **AS.9 — grammar coverage against real external text.** `AS.1` measures test density per rule; nothing measures what fraction of a *real* procedure Frazaro's corpus can express. A fixed sample of external SOP text (public procedure manuals, help-forum questions, a recorded-macro corpus) run through Check, with the refusal rate published alongside the existing test counts. *Why now:* the moat is described as "a curated dialect plus a corpus of tested phrasings"; at 122 rules that claim is currently unmeasured, and this is the instrument that would measure it, the same family as `AS.1`/`AS.2`/`AS.8`. `~days`

---

# 🗣 LANGUAGE · COMPATIBILITY

*Keeping promises to files you cannot see. specimens: 1 (the first-user Undo report — the standing override's founding incident).*

- ⬜ **CO.2 — the frozen compatibility corpus.** *SD-4 is the promise and is in force now; the file is bookkeeping.* `~days`

- ⬜ **CO.3 — a version stamp and a backend stamp inside generated modules**, so a support question is answerable from the workbook alone. Real scope, corrected while scoping SEC.7 (this item's own text used to restate "`requires:` in phrasebooks" as a second thing to build — it's F.10's own mechanism, cited here rather than duplicated): a version marker and a backend marker (interpreter vs. compiled emitter — no new scheme needed). **The version half is now answered rather than still a choice:** CO.4 (✅) settled it — the stamp's value is `VLA_RELEASE_VERSION` under SD-14's triggers, made orderable by `VlaVersionCompare`/`VlaVersionAtLeast`. This item stamps the value and names the backend; it neither picks a scheme nor writes a comparison. *Depends on:* F.10 (the `requires:` half), CO.4 (the ordering, already built). *(more: BETA_ROADMAP1.md)* `~days`

- ✅ **CO.4 — grammar semantic versioning.** **Decided (owner, 2026-09-07): the grammar's compatibility version IS `VLA_RELEASE_VERSION` under SD-14's triggers — no second axis** (a census of 25 `VLA_*_VERSION` constants showed they are per-module item tags, not versions). Built `VlaVersionParse`/`VlaVersionCompare`/`VlaVersionAtLeast` in `VLA.bas`, plain numeric MAJOR.MINOR.PATCH, 26 pins led by "0.9.0 < 0.10.0". Owner-verified live. Left open on purpose: a version alone is a loose gate — CO.6 is the per-form record. *(more: BETA_ROADMAP1.md)* `~hours`

- ⬜ **CO.5 — the migration tool.** Worthless until there is history to migrate, but CO.1 must exist first or there is nothing to migrate *to*. `~weeks`
- ✅ **CO.6 — the grammar since-ledger.** `docs/GRAMMAR_SINCE.md`: 278 rows (150 rules + 128 dispatch arms) dating each form to the release it first *worked* in, append-only. Two seeding traps found: the generated artifact was stale at v0.5.0 (inventory from the artifact, dates from source), and source alone undercounts (three generator-emitted rules). Owner-accepted seed; its ratchet, `check_grammar_since.ps1`, landed with F.10. *(more: BETA_ROADMAP1.md)* `~hours`

---

# 🗣 LANGUAGE · GRAMMAR

*The corpus. The moat, and the reason everything else exists. It has no position in this file, because under SD-12 it never stops. specimens: 0 until PI.2.*

- ⬜ **G-FORMAT** — formatting and number-format sections (~70 rules, mostly thin). Pure Tier-1, no new plumbing, and where a beta looks thin or finished. `~weeks`

- ⬜ **G-SORTFILTER**, ⬜ **G-TABS**, ⬜ **G-FORMULA**, ⬜ **G-TEXT** — the workhorse middle. `~weeks` each

- 🟡 **G-FILES — workbooks and files.** Scoped in `scripts/pareto.txt` section 15 (16 surfaces) - this file previously (wrongly) claimed zero templates existed for this section; a cross-check found six already shipped and already running through F.2's mechanism: `open-workbook`, `save-workbook` (as `save-current-workbook`), `save-as` (as `save-workbook-as`), `save-copy` (as `save-copy-as`), `close-workbook` (simplified - closes THIS workbook, not pareto's named-workbook form), `export-sheet-pdf` (as `export-sheet-as-pdf`). None needed the `{p:path}` slot §15's own preamble calls for - each takes `{p:expr}` bound to a pre-set variable rather than an inline quoted literal. *(more: BETA_ROADMAP1.md)*

- 🔒 **G9** — conjoined predicates. STAYS GATED on pilot evidence. *The best-formed gate in this document: it names the observation that opens it.*

- ⬜ **G-TAIL** — charts, printing, validation, dates, email. Reordered by real sentence-gap reports, not by this file.

- **The query family — English sentence templates targeting each DSL engine's own text, not a fifth S-expression dialect.** Phrasebook rows are recognizers, not generators (settled in `SQL.1`); SQL's parse target stays open between canonical SQL text and the engine's AST until a real English→SQL row needs one. *(more: BETA_ROADMAP1.md)*

  - ✅ **G-DATALOG** — sentence templates against `DATALOG(tables, rules)`. Engine complete; the first real English query anyone wrote against it — *"Show every Staffing name whose salary is over 80000 as rich in cell E2"* — is DATALOG.5's own specimen, and named-column atoms exist because a positional grammar could not say it. *(more: BETA_ROADMAP1.md)*

  - ⬜ **G-SQL** — templates for `SQL`'s frozen subset. Engine complete (SQL.1–SQL.7); arguably the cheapest win in the family, since SQL's clause vocabulary already reads close to a business question. *(more: BETA_ROADMAP1.md)*

  - ⬜ **G-PROLOG** — fact/rule authoring rows producing `(fact …)`/`(rule …)` forms, plus query-goal recognition ("is Bob an ancestor of Liz"). Its engine dependency (PROLOG.1–.6) has since shipped; not pilot-gated. *(more: BETA_ROADMAP1.md)*

  - ⬜ **G-SOLVE** — templates targeting `SOLVE(facts, program)`; named for the function, not the discipline. Nothing to recognize against until a SOLVE MVP exists. *(more: BETA_ROADMAP1.md)*

- ✅ **L-INTERPOLATE — a template-string primitive** (`interpolate`; filed as L-FORMAT, renamed before shipping). Found during G-DATALOG's grammar work. Forced to be an engine primitive rather than a `defmacro`: a macro is parameterized by whole forms and cannot see inside a string. Owner-tested live and committed. *(more: BETA_ROADMAP1.md)*

---

# 🪟 PRODUCT · PERFORMANCE

*Emitted-code speed AND authoring-time latency - both what the user experiences, distinct from OPTIMIZATION's compiler-internals speed. Owner-caught gap (2026-09-03): PF.1-PF.6 below only ever covered the first half - how fast a COMPILED program runs against a big workbook - never the second: how long Check/Compile/vocab-load itself takes for the person AUTHORING the SOP, felt on every edit-test cycle regardless of whether that SOP ever touches a 200k-row workbook. "Loading an SOP cannot take 5-10 seconds - that feels broken to any user" is exactly the product concern this section's own tagline already claimed to cover and didn't. PF.7-PF.9 are that missing half, found during P-TOK's own "what's next" pass (same session P-TOK closed a real, measured instance of this exact shape - compiler latency this time, not authoring latency, but the identical accidentally-quadratic-loop mechanism) - not a new tranche. specimens: 0.*

- ⬜ **PF.1 — the benchmark corpus.** A representative 200k-row workbook and a timing harness. Everything below cites it. `~days`

- ⬜ **PF.2 — automatic environment bracketing.** Screen-updating and calculation bracketed automatically. The single largest constant-factor win in VBA. `~days`

- ⬜ **PF.3 — never emit `.Select`.** Recorded macros are full of it; generated code has no excuse. *Pays into:* correctness too. `~hours`

- ✅ **PF.4 — array slabs for range reads/writes.** Two orders of magnitude on row loops, owner-verified live (2026-09-04). Deliberately *not* the invisible, pattern-detected rewrite this bullet first promised: the owner flagged the determinism/auditability risk and `docs/CONSULTANT.md`'s own doctrine ("every emitted line is derivable by hand from visible source") confirmed it, so the fast path is visible in the source rather than silently inferred. *(more: BETA_ROADMAP1.md)*
**PF.4a — array-element `set!`/read parity, ✅ owner-verified live (2026-09-03).** A real gap found during scoping, both directions: `(set! (arr i) v)` and `(arr i)` already worked compiled via plain text substitution (`EmitStmt`/`EmitExpr`'s own generic `Case Else`), but the interpreter had neither - `ExecSet` raised `interp-set-place-not-object`, `EvalDynamicHead` raised `interp-head-unresolved` - nothing in the existing corpus ever needed either direction before. Fixed as two purely-additive fallback tiers. **A real bug caught by the owner's own first live run, not reasoned through in advance:** both tiers' first draft called `VLA_Runtime.VlaDictGet` directly to test "is this bound to an array" - but `VlaDictGet` is deliberately loud on a miss (`rt-dict-key-missing`), not a safe try-get, so a plain `(range "a1")` place (also `Count = 2`) raised "there is nothing stored at key 'range'" for every ordinary bare-place Range assignment - 10 host failures, all `TestStmtParity`/one IN.2 case, never the pure suite. Fixed by guarding both tiers with `VlaDictHas` first, the existing non-raising check already built for exactly this shape. `VlaSelfTestsAll`: 910/910 pure, 125/125 host, clean on the second run.
**PF.4b — bulk read/write-back runtime helpers, owner-verified live (2026-09-04).** `VLA_Runtime.VlaSlabRead`/`VlaSlabWrite` - **a real design correction caught before landing, not after:** the first plan was delegating to `VLA_Relation.SourceToArray`'s already-correct bulk-read pattern, until re-reading `VLA_Runtime.bas`'s own EN_RUNTIME INJECT BOUNDARY (the exact text copied verbatim into a standalone user workbook, with NO other add-in module available there - the same V5.3 problem this module was split out to solve originally) showed that would raise "Sub or Function not defined" the moment the add-in isn't present - exactly IN.9's own "export" scenario PF.4's performance claim is actually about, not just interactive add-in use. Built self-contained instead (R7's own precedent - this module's `Fold` already duplicates rather than depends, for the identical reason). Neither function is ListObject-aware or strips a header row (a SQL/DATALOG-specific convention, not a general one); `.Value` throughout, not `.Value2`, to behave exactly like the `Cells(i, j).Value` loop it replaces. Tests: `TestArraySlabHelpers` (multi-row/col with a real mutate-then-verify, 1x1, 1-row/N-col, N-row/1-col).
**PF.4c — the `for-each-row` form itself, owner-verified live (2026-09-04).** `(for-each-row (row range) body...)`, wired to `PF.4a`/`PF.4b`. Bulk-read the range once into `arr`; `row` is not a real emitted variable at all - `(row i)`/`(set! (row i) v)` compile DIRECTLY to `arr(iVar, i)` (a dynamic-extent substitution in `EmitExpr`, checked ahead of its normal dispatch, save/restored around the body the same way `mAtLine`/`mGenRow` already are). One bulk write-back after the loop. All-or-nothing commit falls out of that shape with nothing special-cased: `VlaSlabWrite` is the one commit point, called once, after the loop, and because nothing is ever deferred into a separate buffer, a bare `exit-for` needs no cooperation of its own - whatever's mutated up to and including the break's own row is already sitting in `arr` the moment it fires. `return` does not commit (`EmitReturn`'s bare `Exit Sub`/`Exit Function` always leaves the whole procedure) - the interpreter's own `ExecForEachRow` checks `mProcReturn`/`mGotoLabel` first to match, a real cross-backend asymmetry caught during scoping. "a raised error discards everything" - argued from the shape, then owner-confirmed live (2026-09-03): a genuine out-of-bounds error mid-loop (rows 1-2 already mutated in memory, row 3 raising before ever reaching it) left the range fully unchanged after both Run and Interpret, on a real workbook. Body is in-place read/modify/write only for now - no raw `Cells`/`Range` calls inside it. Tests: `TestStmtParity` gains a full round-trip case and a break-commits-partial case. **Three real bugs caught live across two owner-run passes, not reasoned through in advance:** (1) a hidden bookkeeping name, `cVar`, collided with `CVar` - a real VBA intrinsic (Convert-to-Variant) - turning `Dim cVar As String`, alone on its own line, into a bare "Syntax error." (2) the string-building was reshaped from one long continuation into separate `r = r & ...` appends, matching `EmitFor`/`EmitForEach`/`EmitIf`/`EmitSelect`'s own style. (3) the real design flaw, found by `VlaSelfTestsAll`'s own break-commits test: the original fix for (1)/(2) still sliced each row into its own small array, ran the body, then copied it back into `arr` in a SEPARATE loop placed AFTER the body - so `exit-for` skipped that trailing copy-back, losing the mutation on the exact row where the break fired (`20 + 20 + 30 = 70` measured live, not the expected `90`). Fixed by removing the separate row array entirely for the direct-substitution design above, deleting the copy-back step - and the whole bug class - rather than special-casing `exit-for` to route through it correctly. `~weeks`

- ⬜ **PF.5 — loop-invariant hoisting.** `~weeks`

- ⬜ **PF.6 — published numbers, pinned.** The performance claim becomes a test that fails when a change makes generated code slower. `~days`

- ✅ **PF.7 — `VlaEmbeddedText`'s own O(n²) string-build.** Every embedded-sheet load (Check/Compile's vocabulary and prelude, `IdeLoadVocab`'s chain) concatenated cell by cell through COM; replaced with one bulk `.Value` read and a `Join`. Found during P-TOK's what's-next pass; owner-verified live. *(more: BETA_ROADMAP1.md)*

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

*Vocabulary as computable, inspectable data - the F.13 claim ("english.vla is real VLA now") pushed as far as one session's worth of scoped passes could push it. Ordered by dependency, house style: each shipped item is what made the next one cheap, not just next on a list. The line's own constitution, adjudicated this session and binding on everything below it: determinism is the wall between this project and "the Wild West of LLMs" (the owner's own words) - `gensym` breaches it by inventing values from nothing and is vetoed outright; `car`/`cdr`/`cons`-style recursion over `quote` data does NOT breach it, provided it never reads anything but the literal, finite source text being compiled - the same wall, guarding a different door. *(more: BETA_ROADMAP1.md)*

- ⛔ **GENSYM** — hygienic-macro identifier invention. **Vetoed by the owner, this session, in these words:** *"gensym is off the table, because determinism is the Chinese wall between this project and the Wild West of LLMs."* Not a style objection - a gensym'd name appears directly in emitted VBA with no trace to anything a person wrote, and its determinism across recompiles/reorderings cannot be guaranteed the way every other identifier in the pipeline already is. Confirms L-TIER3's own earlier "declines... gensym" rather than reopening it. `eval` was never proposed and stays off the table by the identical reasoning - runtime-arbitrary-code-execution is the same wall's other, more obvious breach. Kept here, permanently, precisely so a future session doesn't re-propose it without this context.

- ⬜ **DIALECT-REGEN** — regenerate `pirate.vla`/`latin.vla`/etc. mechanically from `english.vla`'s own loaded `mPatForms`/`mPatTexts` (real data already) plus a glossary table, rather than hand-translated as this session's seven files were. The seven hand-built files become the acceptance corpus for this generator, not artifacts to maintain by hand forever.

- ⬜ **VOCAB-MIGRATE** — version a generation spec, diff the rules it produces against what's currently loaded (VOCABDIFF), and report the delta - schema migrations for a spoken grammar. No known precedent to crib from; scoping this means designing close to first principles.

- ⬜ **PARETO-SPEC** — *correcting an overclaim made out loud this session, not proposing something new*: `pareto.txt` is prose with its own marker convention, not VLA - no generator can read it directly, ever, without a from-scratch parser for a format that was never meant to be parsed. The only real path is a human translating its rows into a VLA-syntax spec by hand, at which point it is a new artifact, not "pareto.txt, compiled." Filed to keep the corrected version on the record, not the first one.

- ⬜ **WORKBOOK-SPEC** — a generator reading its own row list from a live Excel range at build time. Thematically inevitable for a project this spreadsheet-native; collides directly with "the assistant writes VBA blind, no Excel available," and - stated plainly, not discovered later

- ⬜ **LISPIMPORT** — catalogue, not a chokepoint: once QUASIQUOTE and LISTOPS both exist and have real use behind them, survey further Lisp primitives worth porting against friction the shipped items above actually hit, not against "Lisps have this." `apply` (expand-time variadic calls) is the one candidate on the table; `gensym` is not - see the veto above - and `eval` never was.

---

# 🗣🔧🪟 LANGUAGE + MACHINE + PRODUCT · THE EDITION LINE

*One download per language — the owner's own call, this session, and it resolves more than it costs. Excel's ribbon (`customUI` XML) is embedded per file and is not runtime-swappable in practice, so chrome was always going to need a build-time answer regardless of what the rest of this line did; making every per-language surface build-time-selected rather than runtime-dispatched also retires a real VBA wall that would otherwise have blocked it outright — two `Public` procedures of the same name cannot coexist in one compiled project (`VLA_DevRig.bas`'s own `VlaAuditDuplicates` exists because this already happened once, by accident, on this very project), so an edition-per-download design is what makes a literal `VLA_Spanish.bas` *possible*, not merely simpler. Ordered by dependency, house style: each item below is the instrument the next one is built with, not just next on a list. The owner's own founding principle, on the record, this session: *"the language in which a person thinks should be the language in which a person works."* Measured here as the distance still standing between LX.5's proven grammar seam and a Spanish speaker's whole session — vocabulary, refusals, and the ribbon they read — staying in Spanish. specimens: 0 — no live non-English user has asked yet, the same standing this line's own `LX.10` proof fixture had before it was built; this is instrument-building on principle, the way that item was, not a response to a demand. *(more: this session's own transcript — CONTEMPLATIONS.md's Contemplation 3 and REBUILD.md's Layer 3, both written independently of this line and both landing on the same shape: vocabulary as tables an axis owns, not code duplicated per language.)*

- ✅ **EDITION-MANIFEST** — `VlaBuildAddin` builds every known edition in one pass (`VlaEditionNames` — English and Espanol today); one edition's phrasebook defect is reported rather than aborting the rest, and `VlaBuildAddin "Espanol"` builds one. Each edition audits its own phrasebook chain, embeds it, and gets its own named `.xlam`. Owner-verified live. The long build record — including the external-override regression (EDITIONMANIFEST.5/.6) that put the polyglotta fallback in `IdeLoadVocab` rather than `IdeVocabPath` — is in BETA_ROADMAP1.md. *(more: BETA_ROADMAP1.md)*

- 🟡 **EDITION-VOCABPATH** — Built 2026-09-07, awaiting owner test: `IdeVocabPath` gained `includePolyglotta`, adding `scripts\polyglotta\<edition file>` as a fifth candidate for the three ribbon callers that need a real file (Translate to VLA/VBA, Export Expanded Phrasebook, Coverage), while `IdeLoadVocab`'s pre-check keeps the default so EDITIONMANIFEST.5's regression cannot reopen. *Why it jumped the queue:* it was the only thing that could regenerate the stale `english_expanded.vla` the release ratchet reads. *(more: BETA_ROADMAP1.md)* `~hours`

- ⬜ **EDITION-MODULETAG** — `OUT_MODULE`/`RT_MODULE` hardcode `EN` into every compiled program's module name regardless of edition, so a workbook compiled by both editions collides — exactly the bug IO6.0's tag exists to prevent. Make the tag per-edition. Found during EDITION-MANIFEST's scoping, not guessed. *(more: BETA_ROADMAP1.md)*

- ⬜ **EDITION-PARITY** — the verification discipline, built before any edition has real content to verify: a built edition's emitted VBA/VLA for an equivalent program must match the English reference byte-for-byte, `SD-18`'s own "follows the goldens, never leads them" applied to a second *language* edition instead of a second host, plus each edition's own language-specific `test-success`/message proofs (`TestLx10NonEnglishFixture`'s own shape, generalized). *Why before content:* pouring vocabulary/message/chrome work into an edition with no parity check is exactly how `VLA_Build.bas`/`VLA_DevRig.bas` drifted — a canary from day one is cheaper than a recount later. *Depends on:* EDITION-MANIFEST. `~days`

- ⬜ **EDITION-VOCAB** — extend LX5.1's `keyword-alias` directive shape to the vocabulary helpers still hardcoded as VBA `Select Case` in `VLA_English.bas` (`NumberWord`, `OrdinalWord`, `IsColorWord`, `IsNoiseWord`, `SkipArticles`, `ExprOpWord`, …) so each is table-driven from a loaded file — one compiled engine however many editions exist. Two categories first, to prove the mechanism generic (G6.0's precedent). *(more: BETA_ROADMAP1.md)*

- 🟡 **EDITION-MACROS** — the `.vla`-side `defmacro` layer an author writes rule bodies in: `espanol.vla` had Spanish patterns over English macro-call bodies, so a Spanish speaker could read a sentence but not extend the corpus. Owner directive: a thin Spanish-named wrapper macro over every english.vla action macro. Verification correction on record: the dev-rig `expand` REPL seeds only `prelude.vla`, so phrasebook macros are proven through a real compile, not the REPL. *(more: BETA_ROADMAP1.md)*

- ⬜ **EDITION-MESSAGES** — activate `VLA_Messages.bas`'s dormant seam: swap `AddEntries`' bodies for a Spanish catalogue with zero changes at the ~250 `RaiseMsg` call sites. Error text is the most-read prose in the product (LX.8), which is why this comes before full-coverage translation. *(more: BETA_ROADMAP1.md)*

- ⬜ **EDITION-CHROME** — `LX.12`, executed: give ribbon captions, dialog text, sheet names (`Output`, `Trace`, `Known Sentences`), and column headers the same id-plus-table treatment EDITION-MESSAGES just gave refusals. The one item in this line with zero existing seam today (checked directly, not assumed: `VLA_IDE.bas`'s `"Trace"`/`"Known Sentences"` are bare literals with no indirection at all) — real new infrastructure, not an extension, which is why it sits after two items that already established the pattern it gets to copy rather than invent. `~weeks`

- ⬜ **EDITION-ESPANOL** — the dissection the rack above was built for: author real Spanish content across vocabulary, messages, and chrome using every instrument in this line, build `Frazaro_Espanol.xlam` for real through EDITION-MANIFEST, verify it through EDITION-PARITY, and take `DEPLOY.md`/`SIG.1`/`SIG.5` through their own first per-edition pass — Distribution's and Procurement's own departments, not just Linguistics's. `espanol.vla` already exists, already has real (if partial) coverage, and is the reason Spanish rather than any other language is first. The proof that the whole rack was worth building, the same role LX.10's fixture played for the grammar seam. *Depends on:* every item above. `~weeks`

---

# 🔧 MACHINE · OPTIMIZATION

*Compiler speed. Under the hood; the user never sees it directly. specimens: 0.*

- ✅ **P-PROF** — per-phase timing behind one switch (`VlaProfileAll`). The before-number over 228 transpiles — tokenize 0.9%, parse 27.6%, **expand 71.3%**, emit 0.2% — pointed P-DICT/P-NTH at the right place and surfaced P-TOK's compile-side twin. Built and owner-verified (2026-09-03). *(more: BETA_ROADMAP1.md)* `~days`

- ✅ **P-DICT** — registry lookups off `Collection` error-traps: `mMacros` became a `Scripting.Dictionary` (`.Exists`, direct overwrite). Read honestly: no clean win on this corpus — `expand` flat, `parse` up ~10% — its value was ruling this out so P-NTH could find the real cost. Built and owner-verified (2026-09-03). *(more: BETA_ROADMAP1.md)*

- ✅ **P-NTH** — `Nth`'s Collection walk on hot paths. `cdr`/`cddr`'s O(n²) full-copy `ListTail` replaced with O(1) structural sharing via `VlaSlice.cls`, a read-only view over a Collection's tail. The decisive number: `expand` 42,918 → 1,012 ms (−97.6%). Built and owner-verified live (2026-09-03). *(more: BETA_ROADMAP1.md)*

- ⬜ **P-CONS.1** — exploratory: does a real cons-cell chain even win on VBA's own substrate? Filed from the owner's question during P-NTH ("why not go all the way"), answered honestly rather than deferred: it means the whole list representation and hundreds of construction sites, weeks not days — so measure first, and build only on a number. *(more: BETA_ROADMAP1.md)*
- ✅ **P-TOK** — single-scan tokenizer buffer. Three positionally-indexed `Collection`s (`TokAt` translate-side, `TokRawLine`/`ParseForm` compile-side, and `TokLine`, found during scoping) are now materialized once into 1-based arrays for O(1) access. Built and owner-verified live (2026-09-03). *(more: BETA_ROADMAP1.md)*

- ⬜ **P-PRELUDE** — *largely subsumed by F.3.* `~hours`

- ✅ **P-PROBE** — batched/cached load-time probes. `RegisterVocabMacro` re-transpiled every previously accumulated macro on each new registration — O(macros²), a registration-count cost independent of LISTOPS-BUDGET's depth finding. Built and owner-verified: probes are context-pushed (`VLA.VlaProbeMacroForm`) and one full Pass-1 compile per load replaces one per macro. *(more: BETA_ROADMAP1.md)*

- ~~**P-BULK** — array-slab runtime helpers.~~ Retired (2026-09-03), absorbed into **PF.4** (PRODUCT · PERFORMANCE, above) - confirmed, not just suspected: this tranche is compiler-internals speed the user never sees, but "array-slab runtime helpers" is code a *compiled program* runs, squarely PF.4's own "emitted-code speed" domain. Wrong tranche, not a second item.

---

# 🔧 MACHINE · QUERY AND LOGIC

*Aspirational, not demand-driven — the opposite contract from THE MIDDLE LAYER above. specimens: 1 (owner's own long-standing want, stated live 2026-08-28 — "you have no idea how long I've wanted both a SQL function and a PROLOG function inside Excel" — not corpus demand, so this tranche stays thin and appetite-boxed rather than hydrated). Buy the decision now; defer the artifact.*

- ✅ **`SQL(table, query)`** — a relational query engine as a real worksheet function, `=SQL(query, tables…)`. **Shipped, SQL.1–SQL.7, each owner-tested live:** single-table `SELECT`/`WHERE` (SQL.1); computed expressions, `AS`, `DISTINCT` (SQL.2); `INNER JOIN … ON` on `VLA_Relation`'s hash join (SQL.3); `GROUP BY`/aggregates/`HAVING` (SQL.4); `ORDER BY`/`LIMIT` (SQL.5); `UNION`/`INTERSECT`/`EXCEPT` (SQL.6); `WITH` and recursive `WITH`, the convergence proof (SQL.7). Its own tokenizer and recursive-descent parser (`VLA_Sql.bas`), never `VlaReadForms` — real SQL syntax already lives in users' heads. **Open, below — each already refused by name today.** *(more: BETA_ROADMAP1.md)*
  - ⬜ **SQL.8 — `LEFT`/`RIGHT`/`FULL OUTER JOIN`, `CROSS JOIN`, and table aliasing (`AS`) for self-joins.** Both gaps already refuse by name (`sql-outer-join-not-supported`, `sql-duplicate-join-table`). Not unknown territory: `DATALOG`'s `not` proves the substrate can express an anti-join, which is the outer-join half. *(more: BETA_ROADMAP1.md)*
  - ⬜ **SQL.9 — `NULL`, `IS NULL`/`IS NOT NULL`, and three-valued `WHERE`/`HAVING` logic.** An undocumented gap, not merely an unbuilt feature: a blank cell flows through `RangeToRows` as a raw `Empty` with whatever VBA's coercion happens to do, not a chosen behavior. Real SQL's `NULL` is foundational — comparisons yield `UNKNOWN`, aggregates skip it. *(more: BETA_ROADMAP1.md)*
  - ⬜ **SQL.10 — `LIKE`, `IN (...)`, `BETWEEN`, and `CASE WHEN`.** Four ordinary conveniences with no presence in the grammar; today each falls into the generic `sql-unsupported-keyword` bucket. New cases at `ParseComparison`'s own precedence level; `LIKE` as a small anchored-match routine, never a regex engine. *(more: BETA_ROADMAP1.md)*
  - ⬜ **SQL.11 — subqueries: scalar, `IN (SELECT ...)`, and `EXISTS`/`NOT EXISTS`.** The last structural gap against ordinary relational SQL, not a ceiling item — a subquery is a single non-recursive `SELECT`, so the termination guarantee is untouched. Scalar reuses the pipeline wholesale; `IN`/`EXISTS` are the semi-join and anti-join `RelJoin`'s hash-join law already extends to. *(more: BETA_ROADMAP1.md)*

- ⬜ **`PROLOG(knowledgebase, query)`** — `=PROLOG(clauses, tables…)`: real unification and backtracking over structured facts and rules; free-form English-to-query is a phrasebook problem by the owner's own call, never this engine's. **Shipped, PROLOG.1–PROLOG.6:** `VLA_Unify.bas` one-way match (PROLOG.1) and two-way unification with the occurs check refused by name (PROLOG.2); ground facts and conjunctive queries (PROLOG.3); `(rule …)` forms with SLD resolution and backtracking (PROLOG.4); `is`/arithmetic, negation-as-failure, `findall`, cut (PROLOG.5.1–5.4); table-sourced and named-column facts (PROLOG.6). A 120-step ceiling refuses an infinite rule by name. **Open, below.** *(more: BETA_ROADMAP1.md)*
  - ⬜ **PROLOG.7 — comparison operators (`<`, `>`, `=<`, `>=`, `=:=`, `=\=`) as goals, not folded into `(is ...)`.** Owner-flagged: PROLOG.5.1 shipped arithmetic strictly as a *binding* form, so there is no `(> Salary 80000)` to write — the README's own staffing example had to route around it. New dispatch in `SolveGoalList`, sibling to `not`/`findall`/`!`: a comparison goal succeeds or fails, it never binds. *(more: BETA_ROADMAP1.md)*
  - ⬜ **PROLOG.8 — `=`/`\=` (unification/dis-unification) and `==`/`\==` (structural equality) as ordinary goals.** No way today to unify two terms as a goal mid-body, only implicitly at head-matching time. `\=` desugars to `(not (= X Y))` over `SolveNegation`'s existing bounded sub-call; `==`/`\==` are the stricter, no-binding pair. *(more: BETA_ROADMAP1.md)*
  - ⬜ **PROLOG.9 — type-checking goals (`var`/`nonvar`/`atom`/`number`/`atomic`/`compound`) and `between/3`.** None exist today (confirmed against the dispatch and reserved-name list). Each type test is a cheap shape check over `IsVarAtom`/`TermHasVariable`'s existing classification; `between` is the bounded generate-and-test complement to `findall`. Lower priority than PROLOG.7/.8 — nothing downstream needs a guard clause yet. *(more: BETA_ROADMAP1.md)*

- ✅ **`DATALOG(tables, rules)`** — `=DATALOG(rules, tables…)`, the guaranteed-terminating engine, built first as the cheapest proof of the shared-substrate bet (`VLA_Relation.bas`). **Shipped:** function-free Horn clauses only — compound terms refused at parse time, and that restriction *is* the termination proof; facts from a range or a `(fact …)` block; semi-naive fixpoint; spilled-array results; then stratified negation (DATALOG.1), grouped `count`/`sum` (DATALOG.2), column-scoped table arguments (DATALOG.3), comparison and arithmetic built-ins with safety checks (DATALOG.4), named-column atoms (DATALOG.5). **Open, below.** *(more: BETA_ROADMAP1.md)*
  - ⬜ **DATALOG.6 — `min`/`max`/`avg` aggregation, alongside `count`/`sum`.** DATALOG.2 built exactly `BI_COUNT`/`BI_SUM`; `SQL` next door already has all five through `VLA_Relation`'s shared accumulator. Real, scoped work rather than a switch flip: Datalog's rule-body aggregate forms use their own small walk over derived tuples, not SQL's row-major `GROUP BY`. *(more: BETA_ROADMAP1.md)*
  - ⬜ **Avoiding a full re-parse/re-fixpoint on every recalc — profiled first, not yet built.** Named precisely: `DATALOG` declares no `Application.Volatile`, so Excel's dependency graph already skips it unless an argument cell changed; the real, narrower cost is that a legitimate rerun redoes the *entire* fixpoint when one row changed. A cache keyed on the rules text plus each table's address needs a cheap "has this table changed" signal before it is worth building. *(more: BETA_ROADMAP1.md)*

- ⬜ **`SOLVE(facts, program)`** — Answer Set Programming, the fourth engine, for "what are all the self-consistent ways this could be, and which is best": Datalog plus choice rules and integrity constraints. Not `ASP()` — a real finance-spreadsheet collision. The killer case is staff scheduling, the thing people fight Excel's Solver over, from named, inspectable facts. **Scoped in nine dissections, below, none built; CDCL-style search stays a stated ceiling.** *(more: BETA_ROADMAP1.md)*
  - ⬜ **SOLVE.1 — the zero-choice base case: `DATALOG`'s own engine, unchanged, wearing the `SOLVE` name.** No choice rule and no constraint is exactly a `DATALOG` program: hand the same `(fact ...)`/`(rule ...)` forms to `RunStratifiedFixpoint` and report its one model as "the one answer set." Proves the framing — and the `#SOLVE!` malformed-input refusal — before either new ingredient exists. *(more: BETA_ROADMAP1.md)* `~days`
  - ⬜ **SOLVE.2 — integrity constraints (`:- body.`) against that single candidate world.** Still zero search: a constraint is a post-hoc check over the one derived model, reusing `DATALOG`'s own body-atom evaluator. The first point where "no answer set exists" is a possible outcome — a first-class, auditable result, distinct from the malformed-input refusal. *(more: BETA_ROADMAP1.md)*
  - ⬜ **SOLVE.3 — one choice rule (`{ p(X) : q(X) } = N`), naive enumerate-and-backtrack.** The first real search: `q(X)`'s facts name the candidate pool; enumerate every N-sized subset, check each against SOLVE.2's constraints, return the first survivor. The try/recurse/backtrack shape is `SolveGoalList`'s own loop with a different generator. Chosen atoms may not yet feed other rules — that composition is SOLVE.4, so a bug here has one place to be. *(more: BETA_ROADMAP1.md)*
  - ⬜ **SOLVE.4 — a choice rule's chosen atoms feeding ordinary rules before constraints are checked.** The real scheduling shape: `assign(S,P)` chosen, `overtime(P)` derived by the stratified fixpoint, constraints checked against the *fully derived* model. Composes SOLVE.1 and SOLVE.3 — `DATALOG`'s per-call cost, paid once per candidate. *(more: BETA_ROADMAP1.md)* `~days`–`~weeks`
  - ⬜ **SOLVE.5 — more than one choice rule in the same program.** One choice per shift, not one total: nested generators, still fully naive (the Cartesian product of every choice's subsets through SOLVE.4's pipeline). A small increment, and the first point where combinatorial cost is real — which is why SOLVE.9 waits for it. *(more: BETA_ROADMAP1.md)* `~days`
  - ⬜ **SOLVE.6 — aggregates over a choice's chosen set (`#count`, `#sum`) inside a constraint or cardinality bound.** "Every shift has at least 2 people" counts over the *current candidate's* chosen atoms, not stored facts — the same accumulation as DATALOG.2/.6, reading from a per-candidate source. *(more: BETA_ROADMAP1.md)* `~weeks`
  - ⬜ **SOLVE.7 — first-answer-set vs. all-answer-sets mode, spilled and `LIMIT`-capped.** A caller-facing mode switch over an already-complete search, reusing `DATALOG`'s spilled-array return. Real asks, not speculative: "how many valid schedules are there," "show me a few options." *(more: BETA_ROADMAP1.md)* `~days`
  - ⬜ **SOLVE.8 — `#minimize`/`#maximize` (weak constraints): the best answer set, not just a valid one.** The most business-valuable mode — the thing people fight Excel's Solver over. Runs SOLVE.7's search to exhaustion (or an incremental bound-tightening variant, an owner-decidable choice left open), scoring each survivor with `DATALOG`'s arithmetic evaluator. The one item whose cost can genuinely explode on a workbook-sized problem. *(more: BETA_ROADMAP1.md)* `~weeks`
  - ⬜ **SOLVE.9 — search pruning (unit-propagation-style short-circuiting) over SOLVE.3–SOLVE.8's naive backtracking, profiled first.** Performance, not semantics: every answer stays correct without it. Real profiling against a scheduling-sized workbook before building anything, per the recalc-caching precedent; stops short of CDCL, which stays at the ceiling. *(more: BETA_ROADMAP1.md)* `~weeks`

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

- ⬜ **U.18 — a per-run log, user-facing.** `IN.3.5`'s effect-log golden already exists but is a dev-side test artifact regenerated by `VlaWriteGoldens`, not a persisted, per-run record a user or auditor ever sees. The auditability pitch — "the procedure she wrote is the procedure the auditor reads," `docs/AUDIT.md` Part III — currently has no artifact behind it: no timestamp, no user, no program hash, no effect list survives a run. *Why now:* the single feature that turns "auditable automation" from a slogan into something a compliance reviewer can actually inspect, and the machinery it needs — the effect log — already exists; this is wiring, not invention. *(more: BETA_ROADMAP1.md)*

---

# 🪟 PRODUCT · LEARNABILITY

*Not documentation. Documentation answers "how do I do X." Learnability answers **"what can I say?"** — the actual first-contact problem. specimens: 0.*

- ✅ **LE.1 — the sentence palette.** A browsable, searchable catalogue of every rule with a runnable example, generated from the corpus so it cannot go stale. *Reference by construction (Diátaxis), which is why it can be built early while the user guide correctly waits.* **Owner-verified:** the "What can I say?" ribbon button now writes a real Excel Table named "Phrasebook" (`EnglishPhraseRows`, VLA_English.bas; rendered by `EnglishIdeShowPhrases`, VLA_IDE.bas) - Template/Example columns, browsable, generated fresh from the live grammar every time, worked examples are real proof-verified `test:` sentences, never synthesized. A real `ListObject` rather than plain cells, so every column carries a native AutoFilter search dropdown - searchable for free, no custom search UI needed. Sheet tab and table both named "Phrasebook" for parity. `~weeks`

- ⬜ **LE.3 — in-sheet autocomplete.** A constrained language is an autocompletable one. `~weeks`

- ⬜ **LE.4 — the first-run tutorial workbook.** Runnable, not readable. `~days`

- ⬜ **LE.5 — progressive disclosure.** A new user meets 40 verbs, not 340. `~days`

- ⬜ **LE.6 — worked SOP templates.** *Acquisition instrument, not documentation — belongs in the pilot bundle.* `~days`

- ⬜ **LE.7 — the AI drafting bridge.** *Carried from Alpha 1's F2, and independently rediscovered by the audit's Part III.* "Describe what you want" → the phrase catalogue as the constraint → candidate sentences into column A → Check validates deterministically → red rows drive retry. *Why it is strategically large:* an LLM writing VBA produces code the person who asked cannot read, verify, or sign. An LLM writing Frazaro sentences produces output in a published finite grammar, refused mechanically when malformed, readable by the requester, testable by the corpus protocol, and failing as refusals with directions. **Frazaro becomes the safe target language for AI-generated spreadsheet automation** — the layer that makes a model's output auditable by the person who has to sign it. *Depends on:* LX.7, LE.1, F.4, LX.8. `~weeks`

- ⬜ **LE.8 — the phrasebook's own readability pass.** Five metasyntaxes is four too many: `=>` → `means`, `test:` → `example:`, `fail:` → `never:`, `{r:cell}` → `{a cell}`, `into|in` → `into (or in)`; rule rationale moves from `#` comments to a machine-visible `note:` line so LE.1 can show the best prose in the file, which is currently invisible to the product. `~days`

- ⬜ **LE.9 — split `kernel.vocab` out.** Macros, function mappings, and any rule that must touch a dot-form move to `kernel.vocab`; `english.vla` keeps the rest and contains **no dots at all**. The readability policy made physical: *if you are editing a file with parentheses in it, you are in the wrong file, and the filename says so.* *Depends on:* F.1. `~days`

- ⬜ **LE.10 — author the phrasebook in a workbook.** `english.vla` becomes a build artifact generated from a sheet — still the file of record, still diffable, still what CI checks, no longer what a human edits. The `example:` column is filled by the same act that writes the rule; the "Proven" column is live because the loader already refuses vocabularies whose tests fail; and **LE.1 is the read-only view of the identical table.** On-brand to the point of being slightly embarrassing that it is not already true. `~weeks`

---

# 🌍 COMMONS · INTEROPERABILITY

*Coexisting with everything already in the user's workbook. specimens: 0. **Revised by IN.9:** this tranche now defends the export path only. On the default runtime there is no module to collide and no project to lock, so what was infrastructure-wide risk is now scoped to a deliberate act — which is most of the reason IN.9 went the way it did.*

- ⬜ **IO.2 — locked/signed VBProject handling.** Refuse in words, name the reason — **or switch backends** (IN.6). `~days`

- ⬜ **IO.3 — existing macros, Power Query, connections** left demonstrably untouched. `~days`

- ⬜ **IO.4 — shared/co-authored workbook reality check.** `~days` **Three concrete findings this item inherits, from `docs/CONSULTANT.md`'s own audit, not yet acted on:** (1) the table macros (`table-add-row`/`table-delete-row`/`table-column`, and G-TABLES generally) resolve their target via `(activesheet.listobjects n)` — a program's meaning depends on which sheet happened to be active when it ran, a race condition by design, the exact class `PF.3` ("never emit `.Select`") already polices for emitted code but not for the grammar's own macros. (2) `IN.9`'s own flagship export pitch — email a self-contained `.xlsm` to a colleague — fights Windows' Mark-of-the-Web default (macro-blocked-by-default for internet-sourced files) on unmanaged home machines too, not only managed ones; the pitch needs a caveat or a companion "how to unblock this file" note, not a rewrite. *(more: BETA_ROADMAP1.md)*

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
- ✅ **GO.6 — a Load Phrasebook button.** Found live while verifying SEC.2: no ribbon command existed for an author to load their own phrasebook, and the one working mechanism (`IdeVocabPath`'s undocumented override paths) *replaced* the base corpus rather than adding to it. Built and owner-verified: a file dialog into `EnglishLoadVocabulary` (so SEC.2's raw gate applies), loaded on top of the base corpus, the path remembered per workbook (`VLA_LoadedPhrasebooks` document property) and replayed on Check. *(more: BETA_ROADMAP1.md)*

---

# 🪴 TERRARIUM · DEFERRED, NON-CRITICAL BUGS

*A parking lot, not a department - owner-created after L0.2's own live regression turned up two real bugs neither one was looking for. The point of writing them down here instead of chasing them on the spot: finding a bug while doing something else is not itself a mandate to fix it right now, and without a place to put it, "just this one small fix" is exactly how a regression run for one item turns into three. Deliberately outside F.12's ID namespace - these are found, not planned, and triaging them properly (real repro, root cause, a fix) is itself real work, owed its own turn rather than squeezed in as a rider on whatever else was already in flight. Picked up after the roadmap above is closed out, not before; nothing here blocks anything above it.*

- ⬜ **TER-1 — `Check Instructions` appears to re-run a prior `Compile and Trace`.** Owner-found live, 2026-08-27, during L0.2's regression: `Compile and Trace` run successfully, then `Check Instructions` clicked right after - it visibly re-ran, as if it were another Compile/Trace, not a translate-only Check. A second `Check Instructions` click immediately after does NOT repeat this - only the first click after a run does. Root cause not yet investigated; a plausible lead, not a diagnosis: leftover armed trace/step-tracking state (`VLA_IDE.bas`'s `ArmTrace`) surviving from the Compile/Trace run into `DoCheck`'s own transpile probe - the same class of leftover-armed-state fragility the S5.3/S5.5/S5.6 history already documents once, for a different pair of callers.

- ⬜ **TER-2 — stale row-error marks survive a Check after the row's content is deleted.** Owner-found live, 2026-08-27, during L0.2's regression, root cause checked against the code, not just observed: `ClearMarks` (`VLA_IDE.bas`) only clears `B1:B<lastRow>`, and `lastRow` is `IdeLastRow(ws)` recomputed fresh on every call - if the row carrying the failing mark was the sheet's last non-empty row, deleting that row's content shrinks `lastRow` itself, so the clear range no longer reaches the very cell holding the stale mark it needs to erase. Fix direction, not yet built: `ClearMarks` needs to clear against the sheet's prior extent (or a safely-large fixed bound), not the freshly-recomputed one.

- ⬜ **TER-3 — audit `instructions.txt`/`VerifyReportChecks` coverage in both directions.** Owner-created 2026-08-27, prompted by this session's G-STRUCT pass: building real `VerifyReportChecks` assertions for G-STRUCT's own new `GStruct` sheet caught a genuine argument-order bug in `make {a} look like {b}` (copied the wrong direction) that a `test-success` proof could never have caught, and the SAME pass found the `Demo` sheet's entire "prelude vocabulary" block (styling, rows/columns, sort/filter/dedupe on a small table, sheet protect/unprotect) has run on every regression pass since it was written but has never once had a single `Report`/`CheckV` assertion - a wrong constant or reversed argument there currently shows up as nothing worse than "the script didn't crash." *(more: BETA_ROADMAP1.md)*

## THE SHORT ANSWER, IF YOU READ NOTHING ELSE

*Re-bet at every version open; the previous bet and its corrections are recorded in `BETA_ROADMAP1.md`. SD-12's floor applies regardless: whatever else is on the list, one grammar slice ships.*
