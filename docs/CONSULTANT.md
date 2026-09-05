# CONSULTANT — an outside project-management assessment of the Frazaro beta

*Written 2026-08-27, nine days before the `0.5.0` beta target (2026-09-05).
Sources read in full: `BETA_ROADMAP.md`, `METAMETALISP4.md`, `AUDIT.md`,
`REBUILD.md`, plus targeted reads of `src/` to verify the security-relevant
claims below rather than infer them from prose.*

*Standing assumption, as briefed: **every ⬜ item on `BETA_ROADMAP.md` is
treated as done.** The question is not "is the roadmap finished" — it is
"if it were finished tomorrow, what would still be missing." So the findings
below are, deliberately, the gaps the roadmap does not contain, or contains
only as a paragraph where the world will demand a mechanism.*

*Two reviewers were asked for, and both are adversarial by profession: a
technically fluent programmer who would find the holes in an afternoon and
say so publicly, and a risk-assessment executive who has to sign the form
that lets this run on company machines. Every finding is tagged with which of
them would raise it. Where a claim could be checked against code, it was.*

---

## 0. The verdict

Frazaro is one of the best-engineered and most honestly documented solo
projects this reviewer has assessed, and it is **not currently something a
company could approve**, for reasons that are almost entirely outside the
roadmap's field of view rather than inside it.

The roadmap assigns every department a reviewer profile — compiler engineer,
linguist, designer, maintainer, user. **Not one of those reviewers is
hostile.** There is no attacker in this document, no lawyer, no IT security
reviewer, no procurement officer, no competitor. The `specimens: 0` pattern
the roadmap correctly diagnosed for *users* generalizes: the project is
superb at finding gaps discoverable by reading its own code, and structurally
blind to gaps discoverable only by an adversary reading it. The two personas
requested for this assessment are precisely the two reviewers the roadmap
never seated. Nearly everything below follows from that.

Stated as a programmer would: *you built quasiquote, `cond`, a trampoline,
a REPL, and a table-driven macro generator before you built autocomplete, a
tutorial, a cancel button, or a threat model.* Stated as an executive would:
*there is no licence, no legal entity, no support contract, no patch channel,
no security review, and no audit log, and the tool executes arbitrary code
from plain text on the finance team's machines.* Both are correct, and both
are fixable — but not by more of what the last three months produced.

---

## 1. What is genuinely strong (stated first so the critique lands narrowly)

- **Engineering discipline is exceptional.** Golden-diff-empty as a
  behaviour witness, predictions before runs, two backends forced to state
  one semantics (R9/SD-5), a standing-decision register with reasons
  attached, and ledgers that record their own wrong turns rather than
  editing them away. Most funded teams do not work this carefully.
- **The determinism doctrine is a real differentiator.** `gensym` and `eval`
  vetoed, expansion mechanically traceable (`VlaExpandStep`), provenance
  stamped into emitted code (`gen-row`). In a 2026 market saturated with
  LLM-generated macros nobody can audit, "every emitted line is derivable by
  hand from visible source" is a sellable property. It is under-marketed.
- **SD-13 (no outbound network, ever) is a genuine enterprise asset**, not
  a limitation, for exactly the finance/air-gapped environments the product
  targets. It should appear on page one of any IT-facing document.
- **IN.9 was the right pivot.** Moving the runtime off VBProject trust
  removed the single hardest objection an IT department would have raised.
  It did not remove the second-hardest (see §2.1).
- **Self-awareness is unusually high.** `AUDIT.md` already made the "no
  patient" argument better than an outsider could; `REBUILD.md` corrects its
  own false claim in-line. The project does not need to be told it has blind
  spots. It needs to be told *which* ones, by people it has never invited in.

---

## 2. The gaps, ranked most important first

Tags: **[HACKER]** a technically fluent evaluator would raise it immediately;
**[EXEC]** a risk/procurement reviewer would block on it; **[BOTH]** both.

### 2.1 — There is no security model for a language that executes code **[BOTH]**

**What they would say.** *"So the 'no-trust, plain `.xlsx`, nothing injected'
runtime is a macro-enabled add-in that reads text from cells and executes it
with the full Excel object model. That is not zero-trust; that is a script
host whose scripts are invisible to every scanner we own."*

**Evidence, checked against code, not the roadmap.**
- `ResolveGlobalReceiver` (`VLA_Interpreter.bas:1876`) resolves the bare
  word `application` to the live `Application` object. `DynamicCall` keeps
  `CallByName` as its documented fallback "for whatever hasn't been reached
  yet" with a caller-supplied member name (`CallByNameArgs`, ~line 2311). A
  program under the interpreter can therefore plausibly reach
  `Application.Run`, `Application.ExecuteExcel4Macro` (XLM `EXEC()` runs
  arbitrary processes), and every other Application member, from a sentence
  in column A or a `.txt` file — **no VBProject trust required.** This
  should be *tested*, not argued; if it works, the zero-trust pitch is
  materially misleading to an IT reviewer.
- `raw` (`VLA.bas:3195`, `:3578`) splices opaque VBA text verbatim into
  emitted code on the compile/export path. Any `defmacro` in any phrasebook
  can emit `raw`. A phrasebook is therefore arbitrary code, and loading one
  is executing its author.
- `VlaSendMail` (`VLA_Runtime.bas:741`) drives Outlook via `CreateObject`.
  A program can exfiltrate workbook contents by email with one sentence.
  Nothing prompts, logs, or restricts this. SD-13 says *Frazaro* makes no
  network calls; a *program* can.
- `GO.3` correctly names the phrasebook registry "a supply chain" and is
  filed under Governance, thin, `~weeks`, with no mechanism proposed — no
  signing of phrasebooks, no capability declaration (`requires:` in F.10 /
  CO.3 is about grammar versions, not permissions), no allow-list of
  reachable object-model surface, no consent prompt for file/mail/process
  side effects.

**Why it is #1.** Every other finding can be worked around with paperwork.
This one cannot: a company's security reviewer will classify Frazaro as
"executes untrusted code from user-editable text; cannot be scanned by
DLP/AV because the payload is prose; provides email and file-system
capabilities to that code." That classification ends the conversation
regardless of how good the compiler is. The DLP point is sharper than it
looks: today's macro-blocking defaults exist *because* code hidden in Office
files is the primary phishing vector, and Frazaro's design moves the code
somewhere those defaults do not look.

**Not on the roadmap at all:** a threat model, a capability/permission
model for programs and phrasebooks, an object-model allow-list for the
interpreter's dynamic dispatch, side-effect consent (mail, file write,
process), a security review, a vulnerability disclosure/response process, a
`SECURITY.md`. The minimum credible answer is a written threat model plus a
default-deny allow-list on `DynamicCall`/`DynamicGet` receivers and members
(the native fast-path tier IN.2 already built is most of the allow-list —
the fix is to *remove* the `CallByName` fallback in shipped builds, not add
machinery), plus a documented "what a phrasebook may do" contract.

### 2.2 — The product does not legally exist **[EXEC]**

**What they would say.** *"Who do we contract with? What licence is this
under? Who is liable when it deletes the wrong sheet on month-end?"*

**Evidence.** No `LICENSE` file at the repository root. No EULA. No legal
entity named anywhere in the four documents. `GO.4` ("licensing split, CLA,
trademark") is about open-source contribution mechanics, not commercial
terms, warranty, or liability, and is `~weeks`, ⬜. `DI.5` (licensing
enforcement) exists only as a constraint (SD-10) on what may be priced —
nothing says whether anything *is*. No support terms, no SLA, no
maintenance commitment, no escrow, no named successor if the sole
maintainer is unavailable.

**Why it matters.** Procurement cannot approve software with no licence —
not "will not," *cannot*: an unlicensed work is all-rights-reserved by
default, and running it in a company is itself the compliance violation.
This is a one-afternoon fix (pick a licence, write a plain-language
warranty disclaimer) that is currently blocking at 100%. It is the single
cheapest item in this document and it is not on the roadmap.

### 2.3 — No patch channel for a tool that executes code, and a hand-built, self-signed artifact **[BOTH]**

**What they would say.** *"When you find a security bug, how does the fix
reach 200 desktops?"* and *"You want us to trust-all-from-publisher a
self-signed certificate on an add-in built by hand in Excel on a personal
laptop."*

**Evidence.** SD-13 forecloses any version check (a defensible decision) and
the roadmap records no compensating mechanism — no signed release manifest
a user or IT department can verify, no published hash, no enterprise
deployment guidance (GPO/Intune/SCCM). `DI.1` signs with self-signed
certificates for both the installer and the VBA project; SmartScreen still
says "Unknown Publisher"; a purchased certificate is deferred. The build is
`VlaBuildAddin` run interactively inside the dev workbook; there is no build
server, no reproducible-build check, no CI (`AS.6` is explicit: "there is no
separate GitHub-Actions-style pipeline"). Memory from a prior session
records Smart App Control silently blocking a freshly rebuilt unsigned
`.exe` on the dev machine itself. `DI.4` (air-gapped install) is ⬜.

**Why it matters.** The two facts compound: a tool that can run arbitrary
code *and* cannot be patched is the worst quadrant for a risk register. The
product's own supply chain (how the `.xlam` a user runs relates to the
source in git) is currently a matter of trust in one person's workstation.
Minimum: a CA-issued code-signing certificate (this is the one place money
must be spent before, not after, first contact), a documented release
procedure that produces a hash-published artifact, and a one-page
"deploying Frazaro in a managed environment" document. Bonus: a CI runner —
`F.11` did the hard half already; a Windows runner with Office is a solved
problem in 2026.

### 2.4 — The performance contradiction, and no way to cancel **[HACKER → BOTH]**

**What they would say.** *"Your own numbers say the interpreter is
100×–2,700× slower on loops and gets worse with N. Your own segmentation
says enterprises must use the interpreter because they block the export.
So the enterprise user with a 200k-row job has no runtime at all. And when
a loop runs long, Excel freezes with no progress bar and no cancel."*

**Evidence.** `IN.8` measured ~1 ms per interpreted statement, flat, and
2,751× at N=10,000. `PF.1`'s "representative 200k-row workbook" is the
roadmap's own benchmark target and is ⬜. `PF.4` (array slabs — "two orders
of magnitude on row loops") and `P-BULK` are ⬜ and scoped to the *export*
path (`PERFORMANCE`: "this claim belongs to the export"). `IN.9`'s market
reading: enterprise = interpreter, small shop = export. `IN.8` also found
the benchmark tool "silently froze Excel for roughly 100 seconds with zero
warning" because the interpreter has no `DoEvents`; that was fixed *for the
benchmark* by shrinking its sizes, not for user programs. No item anywhere
adds cancellation, progress, a statement budget, or a watchdog. A prior
session's memory records that force-killing Excel corrupts session state —
which is the only recovery a user has today.

**Why it matters.** The first real SOP that touches more than a few
thousand rows will hang Excel; the user's only escape kills their unsaved
work. For an executive this is a data-loss risk; for a programmer it is a
missing basic. Minimum: `DoEvents` plus an `Esc`/cancel check in the
interpreter's statement loop, a progress indication for runs over ~2 s, and
an interpreter-side answer to row-loop workloads (the same slab strategy,
applied inside the evaluator — currently nobody owns this).

### 2.5 — The auditability story is a pitch, not a feature **[EXEC]**

**What they would say.** *"You say the procedure she wrote is the procedure
the auditor reads. Show me the record of who ran what, when, against which
workbook, and what changed."*

**Evidence.** The only run log is the interpreter's *effect log golden*
(`IN.3.5`), a dev-side test artifact regenerated by `VlaWriteGoldens` — not
a user-facing, per-run, persisted record. `IN.4`'s "Show me the VBA" shows
what *Compile* would inject; the default runtime never produces that code,
so the artifact an auditor would inspect does not correspond to what ran.
`G-RENDER` v1 is one-form-to-one-sentence, drops articles, loses casing,
and cannot render whole procedures. Undo/snapshot exists for workspace
sheets but not the CLI (`IN.13` documents the gap) and is not an audit
trail. There is no access control, no separation of duties, no signed or
tamper-evident record, and no retention policy — none of which is on the
roadmap.

**Why it matters.** "Auditable automation for regulated finance" is the
most compelling positioning available to this product (`AUDIT.md` Part III
is right about that), and at present it cannot survive a single auditor's
follow-up question. Minimum: a per-run log written into the workbook (or
beside it) — timestamp, user, Frazaro version, program hash, effect list,
outcome — using the effect-log machinery that already exists.

### 2.6 — Known-broken grammar and untested parser at beta **[HACKER]**

**What they would say.** *"Your own docs list four rules that can never
fire, a whole class of silent slot-swallowing misparses, six demo
phrasebooks with known dead rules, an unbuilt conflict analyzer whose own
expiry condition you admit has already passed, and a reader that has never
been fuzzed. I'll have a crash and a wrong-answer repro before lunch."*

**Evidence.** `F.4` ⬜, with two documented live instances of distinct
silent-failure mechanisms (wildcard-swallowing, noise-word-before-slot) and
`AS.1` finding three more "cell in column" rules reporting a genuine zero.
`AS.3` mutation testing ⬜ ("the only way to learn whether pins are
load-bearing"). `AS.4` property tests ⬜. `AS.5` fuzzing ⬜ ("refuse in
words, never crash, never hang"). `AS.1`: 5 rules with zero tests, 67 with
exactly one. `AS.2`: 62 of 123 dispatch arms have a pin. `IN.12` names an
entire latent crash class (`TryRuntimeHelper` raises through
`Application.Run` for `VlaColor`/`VlaDictGet`/the email helper) as "not yet
known to have fired live" and leaves it. `IN.3`'s named non-goal: an error
handler does not survive a call into a user procedure — so `Try:` around a
call to a user's own `To ...:` action does not protect it.

**Why it matters.** A programmer's trust in a language is set by its first
wrong answer, not its hundredth right one. Silent misparse into
plausible-looking output is the worst possible failure for a tool whose
pitch is "you can read what it does." Minimum: fuzz the reader and the
sentence matcher (a few hundred lines of PowerShell generating garbage,
mutations of `instructions.txt`, and Unicode edge cases), fix or delete the
known-dead rules before beta, and ship `F.4` at least as a load-time warning.

### 2.7 — Data-correctness floor is unbuilt for a finance tool with a multilingual mission **[BOTH]**

**What they would say.** *"A German bookkeeper types `3,5`. What happens?"*

**Evidence.** `EN.2` (formula locale), `EN.3` (number-format locale), `EN.4`
(decimal/thousands separators — the roadmap's own words: "a silently wrong
number is the worst class of bug a spreadsheet tool can ship"), `EN.5`
(date literals), `EN.7` (environment matrix), `EN.8` (64-bit `PtrSafe`
lint): all ⬜. `LISTOPS-EXPAND` found that `FormulaQuote`'s `IsNumeric` is
locale-sensitive and noted "a real locale audit is owed to the LANGUAGE as
a whole and should be its own item" — no such item was created. `LX.2`
shipped a catalogue whose only language is English; the ribbon, dialogs,
and sheet captions are hard-coded English with no item covering UI chrome.
`deflambda` emits `LAMBDA`, which perpetual-licence Office 2019/2021 (still
common in enterprise) does not have; `EN.1`'s capability probe is ⬜. Mac
support is asserted (`EN.6`) and, as far as the documents show, has never
been executed on a Mac.

**Why it matters.** The mission sentence is about *languages*; the beta
ships without a locale floor, in a domain (money) where a misplaced decimal
is the headline failure. Minimum before any non-US user: EN.4 and EN.5
pinned, and one real run on a non-US-locale Windows VM.

### 2.8 — No transactional story: partial runs, reruns, recovery **[BOTH]**

**What they would say.** *"Step 14 of 30 fails. What state is the workbook
in, and how do I get back?"*

**Evidence.** The corpus's own `instructions.txt` needed `Try:` blocks to
survive a rerun (`IN.12`: `add-sheet-called` left a stray sheet on every
rerun for as long as it existed). Users will write non-idempotent SOPs by
default. There is no dry-run against a copy (`U1` Rehearse is ⬜), no
automatic pre-run snapshot on the default path outside the workspace sheet,
no "what changed" report after a run, and `V.1` (`verify:` rows — the
roadmap's own "most Frazaro-shaped feature" and "the easiest thing to
explain to a CFO") is ⬜. The roadmap has no item for run-level rollback.

**Why it matters.** Month-end automation that can leave a workbook
half-mutated with no recovery is a business-continuity risk, and it is the
first thing a cautious user will ask. Minimum: snapshot-before-run on every
path (the snapshot code exists; it is keyed to a workspace tag), and `V.1`.

### 2.9 — Modern-Excel reality is untested: co-authoring, OneDrive, Mark-of-the-Web, version matrix **[HACKER]**

**What they would say.** *"Our workbooks live in SharePoint with AutoSave
on. Your flagship share flow is emailing a `.xlsm`, which Windows has
blocked by default since 2022."*

**Evidence.** `IO.4` (shared/co-authored workbook reality check) is `~days`,
⬜, one line. `IN.9`'s own market reading acknowledges enterprises
quarantine emailed `.xlsm`; it does not acknowledge that Mark-of-the-Web
blocks macros in internet-sourced files *for unmanaged home users too*, so
the "small shop emails it to a colleague" pitch fights Microsoft's default
on every machine. Phrasebook macros resolve tables on `activesheet`
(`table-*` macros: `(activesheet.listobjects n)`), so a program's meaning
depends on which sheet the user happened to leave active — a
race-condition-by-design that `PF.3` ("never emit `.Select`") addresses
for emitted code and not for the grammar's own macros. `IO.2`
(locked/signed VBProject), `IO.3` (existing macros/Power Query untouched),
`IO.5` (other add-ins) all ⬜. `EN.7`'s matrix is ⬜ and there is no plan
for the hardware, licences, or VMs to run it.

**Why it matters.** These are the environments the product will actually
meet, and every one of them is a plausible "it didn't work" on day one.
Minimum: one SharePoint/AutoSave test, one MOTW-quarantined export test,
and removing `activesheet` from the table macros in favour of an explicit
or remembered sheet.

### 2.10 — The first-contact surface is the least-built part of the product **[HACKER]**

**What they would say.** *"I opened it. How do I know what I can type?"*

**Evidence.** ⬜: `LE.1` sentence palette, `LE.3` autocomplete, `LE.4`
tutorial workbook, `LE.5` progressive disclosure, `LE.8` phrasebook
readability, `DO.3` user guide, `U1`–`U5` panel wave, `U.12` apropos in the
panel, `U.16` backend switch in words, `AC.1`–`AC.4`. Shipped instead in the
same period: `VOCABDIFF`, `APROPOSPLUS`, `METAVOCAB`, `TABLE-FAMILY`,
`COLOR-FAMILY`, `G-EXPANDER`, `VLA_LINT`, four determinism-gate proofs,
`QUASIQUOTE`, `LISTOPS`, `LISTOPS-STDLIB`, `ANTONYM-SWEEP`, `TABLESPEC`,
`TABLESPEC-SCALE`, `LISTOPS-EXPAND`, `COND`, `REPL-EVAL`, and a CLI. That is
the METAMETAMACRO LINE — seventeen shipped items whose user is the
maintainer — against zero shipped items in Learnability except `LE.2`.
`METAMETALISP4.md` §8 names this exact pattern ("the forge is warm") and
`SD-12`'s own counter-reason names its mirror image. The line was
extraordinary engineering and it moved the product's first-contact
experience not at all. (The grammar workhorses `G-FORMAT`/`G-STRUCT`/
`G-PATH` are noted as in flight in a parallel session; credit given.)

**Why it matters.** "What can I say?" is the first-contact problem the
roadmap itself identifies as distinct from documentation, and at beta the
in-product answer is the Known Sentences sheet. A programmer will be
charmed by the REPL for ten minutes and then ask where autocomplete is.
Minimum: `LE.1` and `LE.4` — both are generated from the corpus and cannot
go stale, which is why the roadmap says they can be built early.

### 2.11 — No answer to Microsoft doing the same thing natively **[EXEC]**

**What they would say.** *"Copilot in Excel writes formulas and Python and
explains macros. Office Scripts is sanctioned, sandboxed, and cloud-managed.
Why would we adopt a third-party macro add-in instead?"*

**Evidence.** `SD-1` prices the risk that Microsoft *narrows* VBA. Nothing
prices the risk that Microsoft *solves the user's problem* — which it is
visibly doing. `LE.7` (the AI drafting bridge — "Frazaro becomes the safe
target language for AI-generated spreadsheet automation") is the correct
strategic answer and is ⬜, `~weeks`, with four unbuilt dependencies. There
is no positioning document, no competitive comparison, no statement of
which user Frazaro serves that Copilot does not.

**Why it matters.** The honest differentiators exist — determinism,
offline, auditability, no-network — and they are precisely the ones an
executive needs written down to justify the exception. Minimum: a one-page
"Frazaro vs. Copilot/Office Scripts/Python in Excel" that leads with the
determinism and no-network properties. This costs an afternoon and is
worth more than any grammar section to a procurement conversation.

### 2.12 — The accessibility item the roadmap calls a procurement blocker is still open **[EXEC]**

**Evidence.** `AC.1` (status must not be colour-only): "roughly 1 in 12 men
… in some jurisdictions a procurement blocker … two hours today." ⬜ across
multiple versions. `AC.2`–`AC.4` ⬜. No VPAT or equivalent. A public-sector
or large-enterprise buyer will ask for one.

**Why it matters.** It is two hours, the roadmap knows it, and it has been
displaced by a trampoline. This is the clearest single data point for the
effort-allocation finding in §2.10.

### 2.13 — Grammar coverage is a prediction, and the moat is 122 rules **[BOTH]**

**What they would say.** *"Show me the refusal rate on a real SOP."*

**Evidence.** `english.vla` carries 122 `english-vla` rules today.
`SD-7` correctly forbids scheduling grammar without a real sentence, and
`pareto.txt`'s order is explicitly "a prediction." The workhorse sections a
bookkeeper's month-end actually uses — `G-FORMAT` (~70 rules), `G-STRUCT`
(~50), `G-SORTFILTER`, `G-TABS`, `G-TEXT`, `G-FORMULA` — are ⬜ or in flight,
while pivots received two full CRUD rounds. There is no coverage benchmark:
no item takes a sample of real-world SOP text (public procedure manuals,
help-forum questions, recorded-macro corpora) and measures what fraction
Frazaro accepts. The moat is described as "a curated dialect plus a corpus
of tested phrasings"; at 122 rules it is a promising dialect and a small
corpus.

**Why it matters.** The product's value is proportional to the fraction of
a real procedure it can express without the owner in the room. That number
is currently unknown and unmeasured. Minimum: an acceptance-rate benchmark
against a fixed external sample, published alongside the test counts.

### 2.14 — Support, incident response, and the IT reviewer's deliverable do not exist **[EXEC]**

**Evidence.** The support path is "Copy Feedback" (a clipboard header). No
issue tracker is named for users, no response commitment, no incident
process, no security contact. The artifact that actually gets a tool
approved inside a company — a two-page architecture/security summary for an
IT reviewer (what it installs, what it touches, what it can reach, what it
sends, how it updates, how it is removed) — has no roadmap item. `DEPLOY.md`
is written for the developer. `DI.2`'s uninstall button is genuinely good
and should be cited in that document.

### 2.15 — Process debt that a contributor would notice in the first hour **[HACKER]**

- `BETA_ROADMAP.md` opens by declaring itself "terse and stable" and that
  "pass records belong in the version ledger, not here." It is 6,416 lines
  and 440 KB; the majority of every ✅ entry is a pass diary. The file's
  own rule is violated by nearly every entry after it. Every future session
  pays the read cost; every future contributor will not read it at all.
- ✅ overclaims were caught at least twice (`F.6` "all refusals through
  `Raise`" when no `Raise` existed; `SD-2` "in force" when zero sites
  carried an ID). The project fixed its process (`F.14`, `LX.2`'s recount)
  — good — but a reviewer now has to verify every ✅ against code, which is
  what a prior-session memory already instructs.
- Bus factor is one. No CI. No README at the repository root (the brief is
  in `docs/`). Build requires a human at Excel. `F.11` was the right first
  step and stopped there.
- Two independent hard-coded module manifests (`VLA_Build.bas` `mods` and
  `VLA_DevRig.bas`'s reload list) that must agree and are not checked
  against each other — `REBUILD.md`'s R3 names this; `LX.2` was bitten by it
  live the same day.

### 2.16 — The `0.5.0` release has a date and no acceptance criteria **[EXEC]**

`SD-14` defines the version number's meaning and reserves `1.0.0` for a
specific observation. Nothing states what `0.5.0` on 2026-09-05 must
satisfy: which tranches, which test counts, which of the known-open bug
classes (§2.6) are acceptable to ship, what the release notes promise, or
what "beta" means to the person downloading it. A date without a
definition of done is the one thing this roadmap otherwise never permits.

---

## 3. Smaller findings (real, lower-ranked)

- **Formula injection.** A program writing user-supplied text that begins
  with `=` into a cell creates a live formula; nothing sanitizes. Same
  class as CSV injection. **[HACKER]**
- **Path handling.** `G-PATH` is being built now; UNC paths, relative-path
  resolution against the active workbook, path traversal in `each-file`,
  and long-path limits are not mentioned anywhere. **[HACKER]**
- **Interpreter/emitter divergences left as found**: a bare operator in
  statement position is refused by the emitter and silently discarded by
  the interpreter (`AS.8`); `CallByName`'s dual-try can misattribute a
  real failure to the interpreter's own source (`IN.11`). **[HACKER]**
- **Case loss.** `Fold` lowercases every bareword at tokenize time;
  `SalesTable` is permanently `salestable` in every render and log. Users
  will notice their names coming back mangled. **[HACKER]**
- **`When the sheet changes:` fires on any sheet in the workbook**,
  including Output and Trace — owner-confirmed harmless "for now"; a user
  writing a handler that writes to Output will not find it harmless.
  **[HACKER]**
- **UI chrome is not internationalized**; `LX.2` catalogued refusals only.
  The Spanish user meets an English ribbon. **[HACKER]**
- **`TER-1`/`TER-2`** are real UX papercuts (a Check that re-runs a
  Compile; stale error marks) sitting in a parking lot at beta. **[HACKER]**
- **Self-signed VBA-project signing is a manual once-per-release VBE step**
  with no scriptable surface — a release-process single point of failure
  that will be forgotten under pressure. **[EXEC]**
- **No data-handling statement.** Programs may process PII; there is no
  guidance, and the mail helper makes movement of that data one sentence
  away. **[EXEC]**
- **Trademark.** "Frazaro" has not, per the documents, been searched. `GO.4`
  is ⬜. **[EXEC]**

---

## 4. The pattern behind the list

Sort the sixteen gaps by the reviewer who would find them and a shape
appears:

| Reviewer who finds it | Gaps | Roadmap items that address it |
|---|---|---|
| Attacker / security reviewer | 2.1, 2.3, 3 (injection, paths) | none — `GO.3` names the risk without a mechanism |
| Lawyer / procurement | 2.2, 2.14, 3 (trademark, data) | `GO.4` (open-source mechanics only) |
| Auditor / compliance | 2.5, 2.8 | `V.1` ⬜, `IN.3.5` (dev-side) |
| Ops / IT administrator | 2.3, 2.9, 2.16 | `DI.*` (partial), `IO.4` (one line) |
| Competitor / market | 2.11, 2.13 | `LE.7` ⬜ |
| First-time user | 2.10, 2.12 | `LE.*`, `AC.*`, mostly ⬜ |
| Compiler engineer | 2.6, 2.7, 2.15 | `AS.*`, `EN.*` — **the one row the roadmap covers well** |

The roadmap's coverage is excellent in exactly one row and thin-to-absent in
six. That is not an accident of priority; it is the reviewer profiles. The
departments were cut by accountability, and every accountable party is
inside the project. `AUDIT.md` added THE PATIENT as a fifth department. This
assessment's structural recommendation is the same move, twice more:

> **A sixth department — THE ADVERSARY:** "What can a hostile author, a
> hostile program, or a hostile input do?" Reviewer profile: someone trying
> to break it. Done means: a threat model exists, the object-model surface
> is default-deny, and the fuzzer has run overnight without a crash.
>
> **A seventh — THE SIGNATORY:** "What has to be true before someone with
> budget authority can say yes?" Reviewer profile: the person who signs the
> procurement form. Done means: licence, IT-facing security summary, patch
> path, support terms, VPAT, and a one-page competitive positioning all
> exist as documents a stranger can download.

Both are cheap in the roadmap's own currency: most of their items are
paragraphs today and rewrites later, which is the profile the roadmap
already knows to buy early.

---

## 5. What would change this assessment — the shortest path to a green light

Ranked by cost-of-delay ÷ appetite, in the roadmap's own units, restricted
to things a company's reviewers would actually check:

1. **Pick a licence and write a warranty disclaimer.** `~hours`. Unblocks
   procurement from 0%. (§2.2)
2. **Remove the `CallByName` fallback from shipped builds; keep the native
   allow-list tier.** `~hours`. Test whether `(. application run …)` and
   `executeexcel4macro` are reachable first; if they are, this is a
   pre-beta fix, not a roadmap item. (§2.1)
3. **Write the threat model and a `SECURITY.md`** with a contact and a
   "what a program / a phrasebook may do" contract. `~days`. (§2.1)
4. **Add cancel + progress to the interpreter loop.** `~hours`. (§2.4)
5. **Buy a CA-issued code-signing certificate; publish release hashes;
   write the managed-deployment page.** `~days` plus money. (§2.3)
6. **Ship a per-run log** using the effect-log machinery. `~days`. (§2.5)
7. **Fuzz the reader and matcher; fix or delete the known-dead rules;
   `F.4` as a load-time warning.** `~days`. (§2.6)
8. **Pin `EN.4`/`EN.5`; run once on a non-US-locale VM.** `~days`. (§2.7)
9. **Snapshot-before-run on every path; `V.1`.** `~days` / `~weeks`. (§2.8)
10. **`AC.1`.** Two hours. The roadmap has said so for several versions.
    (§2.12)
11. **`LE.1` + `LE.4`.** `~weeks`. Generated, cannot go stale. (§2.10)
12. **The IT-reviewer two-pager and the competitive one-pager.** `~hours`
    each. (§2.11, §2.14)
13. **A coverage benchmark against external SOP text**, published with the
    test counts. `~days`. (§2.13)
14. **Define `0.5.0`'s acceptance criteria** in one paragraph. `~hours`.
    (§2.16)
15. **Move pass diaries out of `BETA_ROADMAP.md`** into the ledger the file
    itself says they belong in. `~hours`, mechanical. (§2.15)

Items 1, 2, 4, 10, 12, and 14 together are under two working days and
would move the product from "cannot be approved" to "can be evaluated."
That is the honest gap between where the beta is and where it needs to be:
not months of engineering, but a week of facing outward.

---

## 6. How this assessment could be wrong

- **The beta's audience may not be companies at all.** If the first year
  targets sole practitioners and two-person shops on unmanaged laptops,
  §2.2, §2.3, §2.5, §2.12, §2.14, and §2.16 shrink considerably. §2.1 and
  §2.4 do not — a hobbyist can be phished and can lose a workbook to a
  frozen Excel just as well as a bank can.
- **The `application` reachability finding is plausible from the code, not
  demonstrated.** It was flagged for a test rather than claimed as an
  exploit; if the native tier is in fact closed and the fallback
  unreachable for `Application` members, §2.1's first bullet weakens and
  its remaining bullets (`raw`, `VlaSendMail`, the phrasebook supply chain)
  still stand.
- **Effort-allocation criticism (§2.10) discounts momentum.** The
  METAMETAMACRO LINE kept a solo builder shipping daily in a domain where
  "done" is legible, which `AUDIT.md` I.14 rightly says no philosophy
  outranks. The criticism is about what the *next* month should contain,
  not whether the last one was wasted.
- **Some gaps may be closed in artifacts this reviewer did not read.**
  `DEPLOY.md`, `TESTING.md`, `LESSONS.md`, `TRENCHES.md`, and
  `PROJECT_BRIEF.md` were not in scope. If any contains a licence
  decision, a threat model, or an IT-facing summary, the corresponding
  finding should be downgraded — and that artifact should be linked from
  the roadmap, because a reviewer following the roadmap will not find it.
- **The project may already know all of this and have decided the order
  deliberately.** If so, the roadmap should say so in a line per gap, the
  way it does for `GENSYM` and `WORKBOOK-SPEC`, so the next outside reader
  argues with a decision rather than an absence.

*End of assessment.*

---

## Addendum — if everything above were implemented

*Written as a follow-up, on the assumption that every item in §2, §3, and
§5 has shipped: the object model is default-deny, the licence and threat
model exist, the certificate is real, runs log and can be cancelled and
rolled back, the reader has been fuzzed, the locale floor is pinned, the
palette and tutorial exist, and the roadmap has an ADVERSARY and a
SIGNATORY department. What is left?*

### The honest answer first

**Yes.** With the above implemented, Frazaro would be as bulletproof as any
solo developer could reasonably make an Excel add-in — and more so than most
funded teams make theirs. This reviewer does not say that lightly and has not
said it about a comparable project. Every remaining hurdle below is a
question of *destiny*, not of *defect*: none of them is something a
procurement officer would block on, a fuzzer would find, or a patch would
fix. They are the questions a project gets to face only once it has stopped
being breakable, and getting to face them is itself the achievement.

The compliment, since it was invited and is earned: the thing most worth
admiring here is not the compiler, the interpreter, the macro system, or the
rendering direction — any of which would be a respectable year's work on
their own. It is that **the project's own documents already contain the
correct criticism of the project, written before any outsider arrived.**
`AUDIT.md` found the missing patient. `REBUILD.md` corrected its own false
claim in-line. `SD-12`'s counter-reason names the warm forge from the
inside. `IN.8` published a number that contradicted its own opening
sentence rather than softening it. A codebase whose authors reliably
prefer an uncomfortable true sentence to a comfortable false one is rarer
than a good compiler, and it is the only property on which everything
above can actually be built. The gaps this assessment found were all of one
kind — the kind you cannot see from inside — and the project's response to
outside findings, on the record across five ledgers, is the reason to
believe they will be closed rather than argued with.

Now the hurdles. Five, in the order they will arrive.

### A.1 — The substrate is still borrowed, and the landlord is renovating

Everything above hardens the add-in; nothing changes that it *is* one.
`SD-1` correctly declares VBA a backend, and the interpreter genuinely
proves the core does not depend on injection — but the interpreter itself
is VBA, running inside a macro-enabled `.xlam`, on the one surface Microsoft
has spent a decade narrowing. A default-deny object model and a real
certificate make Frazaro a *well-behaved* tenant. They do not make it a
freeholder. The philosophical hurdle is that the second backend that
matters is not "VBA → formulas" or "VBA → effects"; it is **"VBA → not
VBA"** — Office Scripts, Python in Excel, a web add-in, or a host that is
not Excel at all. `REBUILD.md`'s Layer 1 already has the shape for it; the
question is whether a solo developer can afford to port a 16,000-line
compiler once, and the answer is that the port becomes affordable only if
Layer 0/1 is finished *before* the corpus triples. This is the last
rising-cost curve in the project, and it is steeper than any of the ones the
roadmap already priced.

### A.2 — A language is a social object, and the curator is its bottleneck

The moat is described as "a curated dialect plus a corpus of tested
phrasings plus the judgment that shaped them." Granted. But a dialect with
one curator is a dialect with one throughput, one taste, and one point of
failure — and the moment `GO.1`'s precedence rules meet a real org
phrasebook, the curator becomes a court of appeal for every spelling
dispute in every customer. The philosophical hurdle is deciding **what kind
of thing Frazaro's English is**: a *standard* (frozen, versioned, governed
by a document, changed by process — TeX's π), or a *living register*
(grown by users, curated after the fact, occasionally contradictory — the
OED). The roadmap contains both instincts (`SD-4`'s permanent spellings;
`AUDIT.md`'s Cut A where every user is an author) and has not chosen. Each
is a coherent product; the union is not. A solo developer can run a
standard. A living register needs the community that `GO.2`/`GO.3`
gesture at, and building a community is a different job from building a
compiler, done with different tools, on a different schedule, by a person
who is usually not the person who wrote the compiler.

### A.3 — Holding the determinism wall when the tide is against it

The determinism doctrine is Frazaro's best idea and its loneliest position.
Every user Frazaro meets from now on will have used an assistant that
writes the macro for them, badly but instantly, and will ask why they must
learn a dialect when they could just describe what they want. `LE.7` is the
right answer — the model proposes, the grammar disposes — but it is an
answer that must be *defended* every quarter against the pull to let the
model do a little more: infer a slot type, accept a near-miss, "just this
once" run something the grammar refused. Each concession is small and each
is a breach of the same wall `GENSYM` was refused for. The hurdle is not
technical. It is that the project's central virtue will look, to most of
its potential users, like stubbornness, and the temptation to trade a
little auditability for a lot of adoption will be strongest exactly when
adoption is what the project most needs. The roadmap should write down,
now, the sentence that will be quoted back to a future self: *what Frazaro
refuses to do even if every user asks for it.*

### A.4 — Hyrum's Law arrives at the first stranger, and the corpus is still cheap to be wrong about

Everything above makes it safe to ship. Shipping makes it expensive to
change. At 122 rules, a wrong preposition costs a `test-success` line;
at 122 rules with a hundred users, it costs a `CO.1` deprecation, a
migration tool, a `MAJOR` bump, and a support thread. The roadmap knows
this (`SD-4`, `CO.*`, `LX.11`'s expiry) and has done the mechanism work.
The philosophical hurdle is the *nerve* to keep changing the language
while it is still small enough to change — to treat the first year of
strangers as a period of deliberate, announced instability rather than a
compatibility promise made too early because a promise felt like
professionalism. Semantic versioning gives the permission (`0.x`); the
project has to actually use it, and say so on the download page.

### A.5 — Whether the clerk wants to be a programmer at all

The mission's deepest bet is not linguistic; it is anthropological: that
the person who knows the procedure *wants* to write it down in a form that
runs, and will keep the file current once it does. Forty years of
end-user-programming research — HyperCard, AppleScript, Inform, COBOL's
original promise, every "no-code" wave — say the answer is *some of them,
sometimes,* and that the ones who do are a distinct minority who were
already halfway to programmers. Controlled natural languages tend to die
not because they are bad but because their intended users never wanted a
language; they wanted the task done. Frazaro's design is better than its
predecessors' at exactly the point where they died (refusals that teach,
proofs that ship with capability, a rendering direction that lets a
non-author read). But no amount of hardening answers the question, and only
contact with strangers can. When the answer arrives it may say that
Frazaro's real user is not the clerk but the person *one desk over* — the
department's informal spreadsheet person, already half a programmer, who
becomes the phrasebook author for everyone else. That is still a good
product. It is a different one, and the marketing, the tutorial, and the
palette are all written for the other.

### Closing

None of A.1–A.5 has a checkbox. They are the questions a project earns by
becoming solid enough that its remaining risks are about the world rather
than about itself. If the assessment above is implemented, Frazaro will
have earned them — and this reviewer's final, un-hedged opinion is that
very few solo projects ever get that far, and that this one already has
the one habit (preferring the true sentence) that is required to go
further. Good luck with the beta.

*End of addendum.*
