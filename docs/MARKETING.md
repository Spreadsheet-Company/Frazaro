# MARKETING.md — finding the people who will break the beta

*The operating manual for one bullet of `VENTURE.md` §7 — "open source as
top-of-funnel" — and the acquisition-side companion to `LE.6`/`DO.6`, which
the roadmap already labels "acquisition instrument, not documentation." It
is not a roadmap item; it feeds one: `PI.1` (name the pilot — one human, one
SOP, one workbook, one date). Everything below exists to put real sentences
from real people into `PI.2`'s transcript and real refusals into `PI.4`'s gap
log, faster than a solo developer can invent them.*

*Status disclosures, in the house style: **zero external users** as of
2026-09-08; `0.5.2` is the current release; the product makes no network
call by standing decision (`SD-13`), so nothing here is measured by
telemetry — see §7 for what is measured instead. Written the same day
`SEC.8`–`SEC.17` were filed, which is why §1 leads with honesty items rather
than a launch date.*

---

## 0. What each audience is actually being sold

One sentence per audience, the artifact that proves it, and the objection
that will arrive in the first comment. The rest of this document expands
each row.

| Audience | The sentence | Proof artifact | First objection |
|---|---|---|---|
| VBA / Excel developers | A checked English that *compiles to readable VBA* — the macro you would have written, generated from sentences your users can read. | Show me the VBA on the README program; `check_backend_parity` | "Recorded macros already exist." (They're unreadable and unmaintainable — that's the point.) |
| Lisp / PL people | A Lisp whose target is VBA, with the surface language as *data* — eight phrasebooks over one macro layer, including one written entirely in glyphs. | `alien.vla`, `prelude.vla`, `scripts/polyglotta/` | "Another DSL." (The deterministic first-match grammar and the refusal doctrine are the argument, not the syntax.) |
| Prolog people | `=PROLOG(clauses, tables…)`: unification, backtracking, `not`, `findall`, cut — over Excel Tables as facts, with a step ceiling that refuses by name. | README's staffing example | "Where's tabling / clause indexing?" (Stated ceilings in the roadmap; say so.) |
| Datalog people | `=DATALOG(rules, tables…)`: function-free Horn clauses, stratified negation, aggregates, named-column atoms — provably terminating, refused at parse time otherwise. | README's ancestor example; `datalog-*` refusals | "Why not just SQL recursive CTEs?" (Both exist; Datalog is the one that *cannot* not terminate.) |
| ASP people | `SOLVE()` — choice rules and integrity constraints over the same substrate, scoped in nine dissections, not built. | The SOLVE entries in `BETA_ROADMAP2.md` | "Vaporware." (Correct today; the ask is design review, not adoption.) |
| SQL / data people | `=SQL("SELECT …", Table)` as a spilled array: joins, GROUP BY, recursive `WITH`, a frozen SQLite-leaning subset that refuses everything else *by name*. | `sql-outer-join-not-supported` and friends | "No outer joins, no NULL." (Both are open items with IDs — link them.) |
| NLP / controlled-language people | A controlled natural language for spreadsheets: deterministic, first-match-wins, no statistics, refusals that teach. | The Datalog refusal quoted in the README | "That's not NLP." (Agreed — it's CNL, and the ACE lineage is the right shelf to put it on.) |
| Accountants | Your month-end checklist as sentences that either run exactly as written or refuse before touching a number. | A reconciliation-prep workbook (to be built — §1) | "I'm not a programmer." (That is the entire premise; the first tutorial is for you.) |
| Finance / FP&A | Procedures that survive handover — the sheet of sentences *is* the documentation, and the generated VBA is the audit trail. | Undo-snapshot, run provenance, exported module | "We use Python in Excel / Office Scripts now." (Those are for programmers; VENTURE §8.) |
| Audit / risk / EUC controls | A manual control that leaves evidence: what a reviewer reads is what ran; a checked-before-run procedure is a control, not an attestation. | `THREAT_MODEL.md`, run log, `SD-13` | "What can a loaded phrasebook reach?" (README's known-open list — never bluff this audience.) |
| Operations / SOP / quality | Paste the SOP; every line becomes one row; a line that can't run gets a `#` and stays a comment. | `SOP.1`–`SOP.4` (scoped); *Import Program File…* already takes `.docx`, safe to show only once `SEC.13` is fixed | "Our SOPs are in Word." (That is `SOP.1`'s exact premise.) |
| Security researchers | A public threat model, ten open findings from our own review, a disclosure address, and a product with no network stack to hide behind. | `THREAT_MODEL.md`, `SECURITY.md`, `SEC.8`–`SEC.17` | "You shipped with these open?" (Yes, listed, in a beta — the invitation is to find the eleventh.) |

---

## 1. Week 0 — what has to exist before the first post

Nothing in §2 onward is worth doing until every item here is true. Each is
concrete and most are a day or less.

**Accuracy first — the things a first reader will catch.**

1. **`README.md` *Known open security items* is stale.** It lists `SEC.1` as
   open; `0.5.2` closed it. It does not list `SEC.7` or `SEC.8`–`SEC.17`.
   `DEPLOY.md`'s "Before either download" paragraph names `SEC.1`/`SEC.2` as
   open for the same reason. A Hacker News thread will find this in the
   first ten minutes; the fix is an hour, and it converts a liability into
   the single most persuasive thing on the page ("we filed ten findings
   against ourselves this week; here they are").
2. **`README.md` *Install* still says the installer "returns with the next
   release that rebuilds and signs it."** Two releases have shipped without
   it. Either ship `FrazaroSetup.exe` with `0.5.3` or reword to "standalone
   add-in only, for now."
3. **Sign the `.xlam` VBA projects** (`0.5.3` — signing joined the release
   sequence 2026-09-08). Unsigned, a downloader gets a bare *Enable Macros*
   modal with no trust option; signed, they get *Trust all from publisher*
   once and silence forever. This is the difference between "scary" and
   "fine" for the non-developer audiences.
4. **Fix `SEC.13`** (Word automation opens documents with macros enabled)
   before anyone is invited to import an SOP — *Import Program File…*
   already accepts `.docx` in every shipped edition. It is one line.

**The landing — what a stranger sees in the first sixty seconds.**

5. **A GIF at the top of the README.** There is no image, GIF, or video
   anywhere in the repository today. Twenty seconds: type
   `Make cell A1 bold.` → Check → a deliberate typo → the refusal in words
   → fix → Interpret → Undo. Record it once at 1280×720, no narration, no
   cursor tour. It goes at the top of the README, the top of
   `spreadsheet.company/frazaro`, and in every post.
6. **An `examples/` folder** — none exists. One workbook per audience,
   listed in Appendix C, each with its Instructions sheet pre-filled and
   its data on a second sheet, so "open, click Check, click Interpret" is
   the whole tutorial.
7. **The Unblock instruction, in the README, above the download link.** A
   downloaded `.xlam` carries Mark-of-the-Web and Excel refuses it with
   *"This file type is not supported in Protected View"* — a hard stop, not
   a banner (`DEPLOY.md`, "Protected View"). Right-click → Properties →
   Unblock, or `Unblock-File`. Every first-time downloader hits this; if it
   is not the first thing they read, the first fifty testers bounce.
8. **Three issue templates and two Discussion categories** (Appendix B):
   *a sentence I expected to work*, *a refusal that was wrong*, *a
   phrasebook I wrote*; and *Sentences* / *Show and tell*. Today
   `CONTRIBUTING.md` exists (DCO, no CLA) and nothing else does. The
   templates are the corpus intake — `PI.2` insists on the user's own words,
   verbatim, and a template that asks for the exact sentence is how those
   words arrive un-paraphrased.
9. **Three five-minute tutorials at `spreadsheet.company/frazaro`**, one per
   door: *Your first sentence* (Excel user, no code), *Your first
   phrasebook* (VBA developer), *Prolog in a cell* (logic / Lisp). Each ends
   with a download of the matching `examples/` workbook. The site is
   allowed to measure page views; the add-in never measures anything.

**The one thing to decide, not build:** the pilot sign-up. `PI.1` needs one
named human, one SOP, one workbook, one date. A single form — "What is the
procedure? How long does it take you each time? Windows Excel? When could
you spend an hour with us?" — linked from the README's *Today* section,
turns every reader who nods into a `PI.1` candidate. The concierge motion
(`VENTURE.md` §7, `PI.3`) is the product's real first sale; marketing's job
is to fill that one form.

---

## 2. The sequence

Six weeks, in dependency order. The dates are relative to Week 0 finishing,
not to today.

| Week | Motion | Why this order |
|---|---|---|
| 0 | §1, all of it. | Nothing below survives a stale README or an unsigned add-in. |
| 1 | **Soft launch** — r/vba, r/lisp, MrExcel and ExcelForum's add-in boards, the SWI-Prolog Discourse. Small, kind, technical. | Twenty people hit the first-run friction where fixing it is cheap and nobody important is watching. |
| 2 | **Show HN**, logic-programming angle (Appendix A.1). Cross-post to lobste.rs, r/ProgrammingLanguages, r/prolog, r/datalog if active, the Potassco/ASP list for the `SOLVE` design review. | HN's first comment decides the thread; the logic-programming story is the one that survives HN's reflexive skepticism of "program Excel in English." |
| 3 | **The Excel circuit** — r/excel with its showcase flair, pitches to the Excel YouTube channels (Appendix A.7), the Spanish-language Excel ecosystem for the Espanol edition, EuSpRIG abstract submitted. | The audiences that will actually *use* it, reached after the technical audiences have shaken the obvious bugs out. |
| 4 | **Second HN post**, phrasebooks angle: "A Lisp that compiles to VBA, with English phrasebooks in eight languages (including pirate, and one for an alien)." Plus dev.to / Hashnode long-form. | Don't spend that title before the phrasebook tutorial exists to catch the people it brings. |
| 5 | **Accounting / FP&A / audit** — LinkedIn carousel, r/Accounting, r/FinancialModeling, the EUC-controls angle for ISACA/IIA-adjacent groups. | These audiences need the `examples/` workbooks and the signed add-in; by week 5 both exist and have been exercised. |
| 6 | **Retrospective post**: what refused, what we fixed, who found what — credited by handle. Then re-bet the channels (§7). | The daily cadence has produced five weeks of material; the digest is the proof the beta *listens*, which is the only marketing a beta really has. |

**One fixed date that does not wait for the sequence:** 19 September is
International Talk Like a Pirate Day, eleven days from the day this file
was written. `pirate.vla` is real, has nineteen rules, and
`"stow {e:expr} in the cell {r:cell}, arr"` is a working sentence. That is
a free news hook for X and Instagram (§4.5) regardless of where the rest
of the sequence stands.

---

## 3. Community by community

Each entry: who they are, where they are, the hook, the ask (what feedback
we specifically want back), what they will object to and the honest answer,
and one orthogonal move that is not "post a link."

### 3.1 VBA and Excel developers

**Who.** People who already automate Excel and are tired of being the only
person who can read the result. The most likely phrasebook authors, and
the audience that will read the generated VBA line by line.

**Where.** r/vba, r/excel (read the current rules — showcase posts go
through the flair the mods designate, and naked self-promotion is
removed), MrExcel and ExcelForum's add-in and VBA boards, Stack Overflow
(answer, never post), the Excel MVP circle, the Microsoft Tech Community
Excel hub.

**The hook.** *Show me the VBA.* Write the README program, click the export,
show the module. The claim "readable, 1:1 VBA wearing parentheses, no
semantic smoothing" is one they can verify in thirty seconds, and it is
the opposite of what a recorded macro gives them.

**The ask.** Three things, in this order: (1) a macro they already have,
pasted, so we can see which sentences the corpus lacks (§4.2, *Port my
macro*); (2) a phrasebook file for their domain — it is theirs under MPL's
per-file scope (`PHRASEBOOK-TERMS.md`), which is the sentence to lead with;
(3) the generated VBA they would have written differently, with the reason.

**Objections, honestly.** "Recorded macros exist" — unreadable, not
maintainable, and not checked before they run. "I can write this in VBA in
ten minutes" — yes, and the next person cannot read it; the sheet of
sentences is the documentation. "Why a Lisp in the middle?" — because the
grammar is data and the emitter is one function; `README.md` *How it
works* is the two-paragraph answer, and `LESSONS.md` is the long one.

**Orthogonal.** A *VBA-to-Frazaro diff* series: one real, ugly recorded
macro per week, beside the six sentences that replace it, beside the VBA
Frazaro generates from those sentences. Three columns, one screenshot. It
is the single most legible artifact for this audience and it doubles as
corpus mining, because every macro that *cannot* be expressed yet is a
`PI.4` gap with a specimen attached.

### 3.2 Lisp and programming-languages people

**Who.** The people who will read `prelude.vla` before the README.

**Where.** Hacker News, lobste.rs, r/lisp, r/ProgrammingLanguages,
r/Racket and r/scheme (a smaller, kinder audience), the Lisp Discord
servers, the European Lisp Symposium and its mailing list, PLATEAU
(programming-language usability), and the "spreadsheets are code"
research lineage (Hermans; Panko on spreadsheet error rates) — the people
who already believe the spreadsheet is a programming environment.

**The hook.** Not "English." Three things this audience has not seen
together: a Lisp whose *target* is VBA; a surface language that is
literally data, eight of them over one macro layer; and `alien.vla` — the
VLA layer with every name replaced by a glyph, docstrings as the Rosetta
stone (`"=|> : inscribe - put the payload <*> in the cell ~o~"`). Lead with
the alien. Then the doctrine: first-match-wins, no backtracking, the loader
*proves* two rules don't overlap rather than hoping; refusals as a language
feature (`LX.8`); a golden-file parity suite holding two backends to the
same semantics.

**The ask.** Design review, not adoption. Specifically: the `SOLVE`
dissection (nine scoped items, none built — is the sequencing right?); the
decision to keep VLA strictly 1:1 with VBA and never add semantic smoothing
(README, *How it works*, layer 2); and the unknown-namespace refusal
doctrine in `F.10`. These are the arguments this audience enjoys, and the
roadmap already contains them written out, so the invitation is to
disagree with a specific paragraph.

**Objections, honestly.** "Another DSL" — the grammar is the DSL; the
phrasebook is a table; the argument is determinism, not syntax. "VBA is
dead" — Excel is not, and the interpreter needs no VBA trust at all.
"Why not Racket/Guile/…" — because the runtime must be the spreadsheet,
and `SD-13` forbids a network hop to anything else.

**Orthogonal.** Submit `alien.vla` to the lisp-adjacent aggregators as
what it is — an easter egg with a serious point — and write the
*Frazaro is Esperanto for phrasebook* post for r/Esperanto: a small
community that will be delighted, and `esperanto.vla`
(`"metu {e:expr} en la chelon {r:cell}"`) is real. Also: an
awesome-list PR (awesome-lisp, awesome-excel, awesome-prolog,
awesome-datalog) is cheap, permanent, and reaches people who search
rather than scroll.

### 3.3 Prolog and logic programming

**Who.** Practitioners and teachers who have wanted Prolog somewhere
non-programmers already are.

**Where.** r/prolog, the SWI-Prolog Discourse, the Logtalk community, the
ICLP/LPNMR crowd (academic; a short paper or a demo-track submission is
the right form), Prolog Day when it runs, the Association for Logic
Programming newsletter.

**The hook.** The README's staffing example: three ordinary Excel Tables
become facts, one rule states the policy in a supervisor's words, the
result spills as a roster. Then the honest engineering: `VLA_Unify.bas` is
substrate, not engine-private; the occurs check is refused *by name*
(`prolog-occurs-check`), never skipped for speed; an infinite rule hits a
120-step ceiling and says so.

**The ask.** Break the engine, kindly. `PROLOG.7`–`PROLOG.9` are the known
gaps (comparison goals, `=`/`==`, type tests, `between/3`) — we want the
gap they hit *first* in a real knowledge base, with the query. And the
harder question: is a step ceiling the right termination story, or should
tabling (a stated ceiling item) come before comparison goals?

**Objections, honestly.** "No tabling, no indexing" — correct; both are
named ceilings in `BETA_ROADMAP1.md`'s `PROLOG` entry, neither is promised.
"120 steps is tiny" — it is a beta constant chosen to refuse loudly; the
ask above is exactly what number it should be. "Not ISO" — never claimed;
it is Prolog-shaped over S-expressions because VLA's reader was free.

**Orthogonal.** A *logic puzzle in a cell* series — the classic zebra
puzzle, a small scheduling problem, a family tree — each as one
`=PROLOG(...)` formula in a downloadable workbook, posted where puzzle
people gather (r/puzzles is not the venue; the Prolog and Excel
communities both are). Sudoku is explicitly *not* the pitch (the `SOLVE`
entry says so), but it is an honest hello-world for a thread.

### 3.4 Datalog

**Who.** A smaller, denser community: database theorists, the Soufflé and
Logica users, the Datomic/Datascript people, static-analysis folks.

**Where.** The Datalog workshop circuit (the "Datalog 2.0" line of
workshops), r/datalog if it is active, the Soufflé and Logica issue
trackers' discussion tabs (sparingly, as a peer, never as an advertisement),
HN comment threads whenever Datalog comes up — which is often.

**The hook.** Guaranteed termination as a *user-facing* property: compound
terms refused at parse time, stratification checked before the fixpoint
runs, arithmetic built-ins range-restricted (`datalog-builtin-unsafe-variable`),
a 10,000-round ceiling as the last resort. And the thing this community
has not seen: named-column atoms — `(staffing (name X) (salary S))` reads
the columns your table actually has, because the first real English
sentence anyone wrote against the engine could not be said positionally
(`DATALOG.5`'s own history).

**The ask.** Real rule sets. Org charts, bills of materials, dependency
closures — the reachability-shaped questions the README names — from
real tables, with the answer they expected. And a review of the safety
checks: is there a rule that should be refused and isn't?

**Objections, honestly.** "SQL recursive CTEs do this" — `SQL.7` ships them;
Datalog is the one that *cannot* fail to terminate, and the refusal is at
parse time. "Re-running the fixpoint on every recalc" — named, profiled,
and deliberately not yet built (`BETA_ROADMAP2.md`, the recalc-caching
item under `DATALOG`).

**Orthogonal.** A one-page *Datalog for spreadsheet people* explainer on
the site, with the org-chart example — the piece that does not exist
anywhere and that this community would link to. Datalog people are
teachers by temperament; give them the handout.

### 3.5 Answer Set Programming

**Who.** The Potassco/clingo community, ASP researchers, the LPNMR crowd,
and — orthogonally — everyone who has ever fought Excel's Solver add-in
over a staff rota.

**Where.** The Potassco mailing list and GitHub discussions, LPNMR/ICLP,
r/prolog (ASP threads land there), the constraint-programming Discords.

**The hook.** There is nothing to download. `SOLVE()` is scoped in nine
dissections and none is built. The hook is the *design*: choice rules and
integrity constraints over the same substrate as `DATALOG`, `SOLVE.1` being
literally Datalog's engine wearing the name, `#minimize`/`#maximize` as the
business-valuable mode, CDCL-style search deliberately at the ceiling.
And the reason it is not called `ASP()` — a real collision with "average
selling price" in finance spreadsheets.

**The ask.** Design review before a line is written: is enumerate-and-
backtrack (`SOLVE.3`) the right first search for problems that are
workbook-sized? What is the smallest scheduling case that breaks it? Would
a clingo user find the syntax tolerable or offensive? Post the nine
dissections and ask.

**Objections, honestly.** "Vaporware" — yes, and the roadmap says so with
estimates. "Use clingo" — a spreadsheet user cannot; the whole bet is that
the solver lives where the facts already are, offline, with no runtime to
install.

**Orthogonal.** Publish the killer case as a *challenge before the
feature exists*: a small, real rota problem in a workbook (shifts,
certifications, leave, "every shift needs two people, nobody works two in
a row") and ask the community how *they* would write it. The replies are
the syntax review, and the workbook becomes `SOLVE`'s first golden test the
day it ships.

### 3.6 SQL and data people

**Who.** Analysts who know SQL and resent VLOOKUP; the DuckDB/SQLite crowd
who will recognize the "frozen SQLite-leaning subset" phrase immediately.

**Where.** r/SQL, r/dataengineering, the DuckDB and SQLite community
forums (as a peer — "we chose a SQLite-leaning subset and refuse by name
outside it; here is what we froze"), the data-tooling newsletters, LinkedIn
data-analyst circles.

**The hook.** `=SQL("SELECT Name, Salary FROM staff WHERE Salary > 80000", Staff)`
spilling with headers, over a real Table, offline. Then the discipline:
every unsupported keyword has a *named* refusal; the subset is documented
as frozen; recursive `WITH` converges by proof (`SQL.7`).

**The ask.** The query they reach for that refuses. `SQL.8`–`SQL.11` are
the known gaps — outer joins, `NULL`, `LIKE`/`IN`/`BETWEEN`/`CASE`,
subqueries — and we want them ranked by how often a real analyst hits each
one, not by how interesting each is to build. `SQL.9` in particular: what
*should* a blank cell mean, and does the current silent `Empty` behavior
ever produce a wrong number for them?

**Objections, honestly.** "Power Query exists" — yes, and it is a GUI with
a refresh button; this is a formula. "No outer joins" — filed, with a
message that says so by name. "Why not Python in Excel?" — for
programmers, with a cloud round-trip; `SD-13` forbids the round-trip.

**Orthogonal.** A *SQL refusal gallery* — every `sql-*` message from
`VLA_Messages.bas`, rendered as a page, with the sentence that triggers each.
The refusal doctrine is itself the marketing: the error messages *are* the
documentation, and no other spreadsheet SQL tool has a page like it.

### 3.7 NLP and controlled natural language

**Who.** Computational linguists, the controlled-natural-language community
(Attempto Controlled English and its descendants), the people who write
"but what does the user *mean*" papers.

**Where.** The CNL workshop series and its mailing list, r/LanguageTechnology
(carefully — it is dominated by ML), the ACL-adjacent communities only if a
short paper exists, PLATEAU and VL/HCC (visual languages and human-centric
computing), where "an English that refuses" is squarely on topic.

**The hook.** Do not call it NLP; it isn't, and saying so is the hook.
It is a CNL with a spreadsheet as its semantic domain: deterministic,
first-match-wins, no statistical model anywhere, a published grammar, and
refusals that state what was understood, where expectation diverged, and
what to write instead (`LX.8`). The Datalog refusal quoted in the README is
the specimen. The polyglot phrasebooks are the second hook: the same macro
layer under eight surface grammars is a claim about where language
variation actually lives.

**The ask.** Ambiguity hunting. The shadow audit refuses ambiguous rule
sets outright — find the pair of sentences it *should* have refused and
didn't. And the slot-category design (`{r:cell}`, `{f:text-list}`, the
alternation and optional forms): is it expressive enough for the second
language a real user brings, or will the fifth phrasebook bend it?

**Objections, honestly.** "This is 1980s CNL" — the lineage is
acknowledged; the novelty is the substrate and the refusal-as-feature
discipline, not the grammar formalism. "LLMs made this obsolete" — the
README's bet is the opposite, stated in one sentence: *a small, checked
English is better than an unlimited, guessed one.* Argue that sentence,
not the tooling.

**Orthogonal.** A short *position note* — two pages — for the CNL / VL/HCC
venues: "refusal as the primary UI of a controlled language," with the
message catalogue as data. It costs a weekend and puts the project on the
one shelf where it is genuinely novel.

### 3.8 Accountants

**Who.** The analyst in `VENTURE.md` §2 — hours. Month-end formatting,
reconciliation prep, report assembly, the 90-minute weekly procedure that is
75 hours a year.

**Where.** r/Accounting (read the rules; contribute first), the AICPA and
national-body community forums, the accounting-tech newsletters and
podcasts, LinkedIn (the accounting community there is large and
un-ironic), and — orthogonally — the bookkeeping communities, where the
procedures are smaller and more repetitive.

**The hook.** Not the engine. A recognizable procedure, done. The
`examples/` workbook for this audience is a month-end close checklist:
*Work on sheet Trial Balance. Clear formatting of range A1:F200. Make row 1
bold. Set total to sum of range F2:F199. If total is not equal to 0, put
"OUT OF BALANCE" into cell H1.* — every sentence real, every sentence
checked, Undo behind it. The sentence to say out loud: **it either runs
exactly as written or refuses before touching a number.**

**The ask.** Their checklist, verbatim, in their words — `PI.2`'s
transcript. Not "what features do you want": *paste the procedure you did
last Friday*. Every sentence the corpus refuses is a gap with a specimen,
which is the only kind of gap `SD-7` lets us schedule.

**Objections, honestly.** "I'm not a programmer" — the first tutorial
assumes it. "IT won't let me install things" — the standalone add-in is an
Office document, not a program; `PI.7`'s trust reading is the honest answer
about what IT actually permits, and we want *their* IT's answer as data.
"What if it deletes something?" — snapshot first, Undo button; and it
cannot do anything a sentence does not provably say.

**Orthogonal.** *Bring your checklist* — a public, scheduled, recorded
hour where one volunteer's real month-end procedure is turned into
sentences live, refusals and all. It is `PI.3`'s concierge run performed
in public; the recording is the tutorial nobody could write in advance,
and the volunteer is a `PI.1` candidate by construction.

### 3.9 Finance, FP&A, and financial modeling

**Who.** The manager in `VENTURE.md` §2 — fragility. Procedures that live
in one person's head and break on handover. Also the modeling-standards
crowd (FAST, ICAEW's twenty principles), who already believe spreadsheets
need discipline and have written it down.

**Where.** r/FinancialModeling, r/FPandA, the Wall Street Prep / CFI
learner communities, LinkedIn FP&A groups, the Financial Modeling World
Cup and Excel esports audience (orthogonal — §4.6), the modeling-standards
bodies' forums.

**The hook.** *Procedures that survive people.* The sheet of sentences is
the handover document; the generated VBA is the artifact a successor can
read; the run snapshot is the rollback. Against Python in Excel and Office
Scripts, the line from `VENTURE.md` §8: those are for programmers, the
constituency that was never the problem.

**The ask.** The handover story that went wrong — the procedure that broke
when someone left — as a specimen. And the modeling-standards question:
which of ICAEW's principles or FAST's rules does a checked-English
procedure satisfy or violate? That review, from someone who knows the
standards, is a whitepaper (`DO.5`) written by the audience.

**Objections, honestly.** "We have Power Automate / macros / a BI team" —
and the 90-minute procedure is still manual, because none of those is
written in the analyst's language. "Windows only" — correct for export;
Mac interprets. "Is it audited?" — `THREAT_MODEL.md` is public and
`SEC.8`–`SEC.17` are open; say so before they ask.

**Orthogonal.** A *Frazaro solves the case* write-up for a past Financial
Modeling World Cup or Excel championship case: the procedural part of the
case as sentences, the modeling part left alone. The esports audience is
young, online, competitive, and shares things; nobody has pitched them a
language.

### 3.10 Audit, risk, and EUC controls

**Who.** The executive and the auditor in `VENTURE.md` §2 — attestation.
SOX 404 key-spreadsheet owners, internal audit, the End-User Computing
risk function that today buys monitoring tools at five figures to *watch*
spreadsheets without making them safer.

**Where.** EuSpRIG (the European Spreadsheet Risks Interest Group — an
annual conference with a call for papers, and the one room where every
attendee already agrees with the premise), ISACA and IIA chapter
communities, the SOX/ICFR practitioner groups on LinkedIn, the EUC-risk
vendors' own webinars (as a question-asker, not a competitor).

**The hook.** A manual control that leaves evidence. Today the evidence for
a manual procedure is a human's attestation that they followed a document;
a Frazaro procedure *is* the document, is checked before it runs, refuses
otherwise, snapshots first, and can hand the reviewer generated source. The
`SD-13` posture (no network, no telemetry, no update channel) is a
*feature* to this audience, not a limitation — say it in their units:
nothing leaves the machine, ever, by standing decision.

**The ask.** Two things. The control question: what would this need to
show an auditor for a checked-English procedure to count as a control
rather than an attestation? (That answer is `SIG.6`'s hypothesis test,
written by the people who would sign.) And `PI.7`'s trust reading, as a
survey: *what does your IT permit — macro-enabled add-ins, VBProject
trust, downloaded `.exe`?* Ten honest answers reshape the distribution
roadmap (`DI.4`, offline/air-gapped, exists because "many finance
environments are").

**Objections, honestly.** "What can a loaded phrasebook reach?" — the
README's known-open list, complete and current (§1 item 1), never a bluff;
this audience reads threat models for a living. "Unsigned add-in" — signed
from `0.5.3`. "Self-signed certificate" — yes, and `DEPLOY.md` says exactly
what that does and does not buy.

**Orthogonal.** Submit to EuSpRIG. A ten-page paper — *checked English as a
spreadsheet control: procedures that refuse to run wrong* — with the
threat model as the honesty section. Abstract in Appendix A.6. It is
peer-reviewed by the exact people who buy EUC tooling, and the proceedings
are indexed, which means the paper markets for years.

### 3.11 Operations, SOP, and quality people

**Who.** The people who *write* SOPs: quality managers, ISO 9001 process
owners, Lean/Six Sigma practitioners, the person who owns the runbook.

**Where.** r/QualityAssurance and the quality-management forums, Lean and
Six Sigma communities, the process-documentation tool communities
(where people already complain that the document and the doing drift
apart).

**The hook.** `SOP.1`–`SOP.4`'s design, said plainly: paste the SOP, every
line becomes one row, no classification, no AI; a line that cannot run gets
a `#` and stays a comment beside the lines that can. The document and the
procedure stop being two things. (*Import Program File…* already accepts
`.docx` today, and until `SEC.13` is fixed it opens the document in Word
with macros enabled — so do not *promote* the Word path; the plain-text
paste path is the honest offer until then.)

**The ask.** A real SOP, as text, with permission to quote it. Not a
template — the one they actually follow, with its numbered steps and its
"if the total doesn't match, call Priya" lines. Every non-runnable line is
data about what a procedure language must tolerate.

**Objections, honestly.** "Our SOPs are in Word with screenshots" — `SOP.1`
and `SOP.3` exist for exactly that; today, paste the text. "Steps aren't
all Excel" — correct; the `#` convention is the design answer, and the
question of what fraction of a real SOP is Excel-shaped is one we want
measured, not assumed.

**Orthogonal.** An *SOP swap*: people post a short procedure, someone else
writes it as sentences, the original author says whether it is faithful.
Translation-as-review is how the corpus learns what "faithful" means to a
process owner, and it is a format that rewards participation rather than
downloads.

### 3.12 Educators and students

**Who.** Teachers of introductory programming, business-school Excel
courses, and — orthogonally — Latin and classics teachers.

**Where.** r/CSEducation, the SIGCSE community, business-school teaching
forums, r/latin and the classics-teaching lists (`latin.vla` is real:
`"pone {e:expr} in cellula {r:cell}"`).

**The hook.** A checked language whose refusals teach is a *teaching*
language by accident. A student who writes a wrong sentence gets told what
was understood, where it diverged, and what to write — the feedback a
grader gives, instantly. And the polyglot phrasebooks make one lesson
possible nowhere else: the same program in English, Spanish, Latin,
Esperanto, and pirate, with the same generated VBA underneath, as a
concrete demonstration that syntax is surface.

**The ask.** A lesson plan, tried once. Which refusals confused a student,
and what wording would not have? Refusal text is the most-read prose in
the product (`LX.8`), and students are its harshest, most useful readers.

**Objections, honestly.** "We teach Python" — this is not a replacement;
it is a first language for people who will never take a second. "Windows
Excel only" — for export; interpretation runs on Mac.

**Orthogonal.** *Automate Excel in Latin* — a single post, aimed at
classics teachers, with a five-line workbook. It will travel farther than
anything aimed at programmers, because nobody has ever offered that
audience anything.

### 3.13 No-code, citizen developers, and automation people

**Who.** The r/nocode and Zapier/Make/Power Automate crowd; people who
already believe non-programmers should automate and are used to tools that
guess.

**Where.** r/nocode, r/automation, Product Hunt (later — after tutorials
and a signed build), the no-code newsletters, Makerpad-style communities.

**The hook.** The opposite of the prompt box. Against every tool in this
category that accepts anything and is confidently wrong sometimes,
Frazaro accepts a small grammar and is never wrong about what a sentence
means — or refuses. For automations that touch numbers, the pitch is the
refusal.

**The ask.** The automation they built that silently did the wrong thing
once. Every such story is the README's *idea* section in someone else's
words, and the best of them belong in the tutorial as motivation.

**Objections, honestly.** "It can't do X" — probably true; the gap log is
public and every gap needs a specimen. "Copilot does this" — Copilot
guesses; this refuses; pick the one you want near a general ledger.

**Orthogonal.** A Product Hunt launch is the conventional move for this
audience and it should wait for week 6 or later: PH rewards polish and
punishes an unsigned download with a Protected View wall. When it
happens, the tagline is the README's: *a phrasebook for your workbook.*

### 3.14 The Spanish-speaking Excel world

**Who.** The largest non-English Excel community by a wide margin, with
its own YouTube channels, forums, and courses (Excel Total, ExcelyVBA, and
their peers), and no tool that speaks to it in Spanish at the sentence
level.

**Where.** Those channels and forums; r/excel's Spanish-language
counterparts; LinkedIn Latin America; the Spanish-language accounting and
FP&A communities.

**The hook.** `Frazaro_Espanol.xlam` is a real second edition — a
separate download whose phrasebook has Spanish patterns, with
`EDITION-MACROS` giving Spanish-named macro bodies so a Spanish speaker
can *extend* the corpus, not just use it. Messages and chrome are still
English (`EDITION-MESSAGES`, `EDITION-CHROME` are open) — say so; a
half-translated product announced as whole is the fastest way to lose this
audience.

**The ask.** Sentences — the Spanish a real analyst would write — and the
messages they hit in English that they would want in Spanish first.
`EDITION-MESSAGES` is prioritized by exactly that list.

**Orthogonal.** Pitch the demo to one Spanish-language Excel channel
*before* the English ones. The English channels have been pitched a
thousand add-ins; the Spanish ones have not been offered a product with a
Spanish edition on day one.

### 3.15 Security researchers

**Who.** People who read `THREAT_MODEL.md` for sport.

**Where.** The security subreddits are the wrong venue for a beta;
the right ones are the Office-security and macro-malware research
community, the people who write about VBA and XLM abuse, and — once
`SEC.8` (Mark-of-the-Web) is understood — the enterprise-security folks
who will ask about it first.

**The hook.** Total candor. A public threat model, a disclosure address
with a 3-business-day acknowledgment commitment (`SECURITY.md`), ten
findings filed against ourselves in one afternoon (`SEC.8`–`SEC.17`, each
with file and line), a product with no network stack, and an explicit
statement that "zero-trust runtime" here means "no Trust Center prompt,"
not "every capability verified" (`THREAT_MODEL.md` §0).

**The ask.** The eleventh finding. Point them at the shipped list and ask
what we missed — with `SECURITY.md`'s reporting shape (version, minimal
`.vla` or sentence, which backend). Credit by handle in `RELEASES.md` and
`THREAT_MODEL.md`'s next revision.

**Objections, honestly.** "You shipped a beta with these open" — yes,
listed, ranked, with fixes scoped; the alternative was a beta with the same
holes and no list. "Self-signed cert" — `DEPLOY.md` says what it buys.

**Orthogonal.** Publish `THREAT_MODEL.md`'s next revision *as a blog post*,
including the audit-and-cleared list (what was checked and found sound).
Nobody publishes the negative results; the security crowd notices when
someone does.

---

## 4. Orthogonal approaches

Twelve moves that are not "post a link," each with the community it
serves and the feedback it yields.

1. **The refusal leaderboard.** A public list of wrong refusals — a
   sentence that *should* have worked and didn't, credited by handle, with
   the release that fixed it. The refusal doctrine makes "make it refuse
   wrongly" a game, and every entry is a `PI.4` gap with a specimen. (All
   audiences; feeds Grammar.)

2. **Port my macro.** A standing thread: paste a VBA macro, the community
   writes the sentences, the corpus author says what refused. Three
   columns — the macro, the sentences, the VBA Frazaro emits — one
   screenshot per week. (VBA developers; feeds Grammar and `LE.6`.)

3. **The phrasebook jam.** A weekend: write a phrasebook for your
   profession. It is *yours* under MPL's per-file scope
   (`PHRASEBOOK-TERMS.md`) — that sentence is the recruiting line — and
   `GO.3`'s registry gets its first real supply-chain question the moment
   two exist. Winner criterion: the one that makes the most sentences
   *refuse correctly* in `test-fail` proofs, not the one with the most
   rules. (VBA, domain experts; feeds `GO.2`/`GO.3`.)

4. **Adopt a sentence.** Every daily release note becomes one tweet — the
   sentence of the day, a GIF of it working — with the ask: *reply with the
   sentence you would have written.* Replies are `PI.2` corpus lines in the
   user's own words, credited in `RELEASES.md`. (X, Instagram; feeds `PI.2`.)

5. **Talk Like a Pirate Day, 19 September.** `pirate.vla` has nineteen
   working rules. One thread: the README program in pirate, the same
   generated VBA underneath, the alien phrasebook as the closer. It is the
   only day of the year a spreadsheet-automation post is *funny* on its
   own merits; the phrasebook already exists; the cost is an afternoon.
   (Everyone; feeds nothing but attention — which is the point.)

6. **Frazaro solves the case.** Take a past Financial Modeling World Cup
   or Excel-championship case, write its procedural part as sentences,
   post the workbook where that audience gathers. Nobody has pitched
   Excel esports a language. (FP&A, the young Excel crowd; feeds
   `LE.6`.)

7. **Bring your checklist — live.** A scheduled, recorded hour: one
   volunteer's real month-end procedure turned into sentences in public,
   refusals included. `PI.3`'s concierge run as a broadcast; the volunteer
   is a `PI.1` candidate by construction; the recording is the tutorial.
   (Accounting, FP&A; feeds `PI.1`–`PI.3`.)

8. **The SOP swap.** Post a procedure; someone else writes it as sentences;
   the author grades fidelity. Translation-as-review teaches the corpus
   what "faithful" means to a process owner. (Operations; feeds `SOP.*`.)

9. **The SOLVE challenge, before the feature.** Publish the rota problem
   (shifts, certifications, leave, two-per-shift, no back-to-backs) as a
   workbook and ask ASP people how they would write it. The replies are the
   syntax review; the workbook becomes `SOLVE`'s first golden test. (ASP;
   feeds `SOLVE.1`–`SOLVE.9`.)

10. **Conference papers as durable marketing.** EuSpRIG (the audience that
    buys), the CNL / VL/HCC venues (the shelf where it is novel), an
    ICLP/LPNMR demo (the engines). Proceedings are indexed and cited; a
    Reddit post is gone in a day. (Audit, NLP, logic programming.)

11. **Awesome-lists and Wikipedia-adjacent surfaces.** PRs to awesome-excel,
    awesome-lisp, awesome-prolog, awesome-datalog; GitHub topics on the
    repository; a Console.dev submission (a newsletter that spotlights
    open-source tools weekly). Cheap, permanent, reaches searchers rather
    than scrollers. (Lisp, Prolog, Datalog, VBA.)

12. **The negative-results post.** Publish what was audited and found
    sound alongside what was found broken, the kill criteria (`PI.5`)
    before the pilot runs, and the channels this document retired for
    yielding nothing (§7). Nobody publishes the negatives; the audiences
    that matter here — auditors, security people, Lisp people — notice when
    someone does. (Everyone who reads carefully.)

---

## 5. Feedback machinery — turning attention into corpus

The product's own doctrine constrains the intake, and the constraints are
the design:

- **`SD-7`: nothing is scheduled without a specimen.** Every feature
  request must arrive with the sentence, the query, or the SOP line that
  motivated it, verbatim. The issue templates (Appendix B) make the
  specimen the first field, not an afterthought, so a request without one
  cannot be filed by accident.
- **`PI.2`: the user's own words.** Never paraphrase a submitted sentence
  into house style before it is recorded. The transcript is the corpus's
  true north precisely because it is unedited.
- **`LX.8`: refusals teach.** A refusal that did not teach — the user could
  not tell what to write instead — is a bug in the message catalogue, filed
  under *a refusal that was wrong* even when the refusal was correct.
- **`SUPPORT.md`'s response commitment** applies to feedback as to bugs:
  acknowledged within 3 business days, no promised fix date, kept
  informed. Say it on every intake surface so the promise is uniform.

The surfaces, in order of how much they ask of the person:

| Surface | What it captures | Effort asked |
|---|---|---|
| **Reply to a sentence-of-the-day post** | One sentence in their words | Ten seconds |
| **Issue: *a sentence I expected to work*** | Sentence + what happened + Frazaro version | Two minutes |
| **Issue: *a refusal that was wrong*** | Sentence + refusal text + what they meant | Two minutes |
| **Issue: *a phrasebook I wrote*** | The `.vla` file, MPL-scoped as theirs | An hour, once |
| **Discussion: *Show and tell*** | Workbooks, screenshots, the SOP swap | Variable |
| **The in-product *Copy Feedback* button** (`SUPPORT.md`) | Diagnostics the product already gathers, pasted into an issue | One click |
| **The pilot form** (§1) | One human, one SOP, one workbook, one date — `PI.1` | Five minutes, and then an hour with us |
| **`PI.7`'s trust survey** | What their IT actually permits | Three questions |

Closing the loop is the part most betas skip: every fixed specimen is
credited by handle in `RELEASES.md`'s section for the release that fixed
it, and the leaderboard (§4.1) is regenerated from those credits. A
contributor who sees their sentence named in a release note comes back;
one who filed into silence does not.

---

## 6. The content engine

The release cadence (`DEPLOY.md`: a `0.5.N` patch at the end of each
working day, a `0.N.0` minor at the end of each week) is already a
publishing schedule. Each tier has one format:

- **Daily patch → the sentence of the day.** One sentence from that day's
  work, a fifteen-second GIF of it working, the ask (§4.4). Pulled from
  `RELEASES.md`'s section, so the content is written before the release,
  never after — the same rule the release script enforces.
- **Weekly minor → the before/after.** One real procedure — a recorded
  macro, a manual checklist, a Solver fight — beside the sentences that
  replace it. Long enough for LinkedIn and dev.to.
- **Monthly → the digest.** What refused, what we fixed, who found what,
  what the channels yielded, and the negative results (§4.12).
- **Refusal of the week.** A real refusal, quoted, and the sentence that
  fixes it. The message catalogue is the most-read prose in the product;
  publishing it one line at a time is documentation and marketing in the
  same artifact.

**The memes stay.** The Instagram and X accounts (`@spreadsheet.company`,
`@spreadsheetery`) are already running; the one rule to add is that every
meme links to a sentence that works *today* — a meme that lands and a
download that refuses is the worst outcome available.

---

## 7. Measurement without telemetry

`SD-13` forbids the add-in from phoning home, permanently. Adoption is
therefore measured by proxy, and saying so publicly is a differentiator to
exactly the audiences that matter (§3.10, §3.15). The proxies:

- **Release download counts**, per asset, from the GitHub releases API —
  the closest thing to installs, and per-edition, so the Espanol edition's
  uptake is visible.
- **Issues by template** — the specimen count, which is the only number
  `SD-7` cares about.
- **Discussions, phrasebook PRs, stars and forks** — the community's
  pulse, in that order of meaning.
- **Site analytics** at `spreadsheet.company/frazaro` — the site may
  measure page views and tutorial completions; the add-in never measures
  anything. Draw the line publicly and keep it.
- **Pilot-form submissions** — the number that matters most, because one
  of them becomes `PI.1`.
- **`PI.7` trust-survey answers** — a small, priceless dataset about what
  enterprise IT permits.

**Channel kill criteria, `PI.5`-shaped.** Written before the sequence
starts, per channel: *if this channel has produced zero specimens (issues
with a sentence in them) in thirty days, it is retired and the retirement
is published in the digest.* Attention that does not produce a sentence is
not attention this project can use.

---

## 8. Guardrails — what not to say, what not to do

- **Never say "zero-trust."** `THREAT_MODEL.md` §0 is explicit: the phrase
  means "no Trust Center prompt," and a security audience will hear the
  other meaning. Say "no trust toggles by default" and point at the
  known-open list.
- **Never hide an open `SEC.*` item.** The README's list is complete and
  current or the launch waits (§1). Ten open findings, listed, is a
  strength; one discovered by a commenter is a thread.
- **Never call it AI.** The README's bet is the opposite of the prompt box;
  the word invites the comparison the product exists to refuse.
- **Never claim Mac parity.** Mac interprets; it cannot export VBA
  (README, *Today*).
- **Never astroturf.** One account per platform, real name, "I built this."
  Reddit's self-promotion rules are enforced by people who can smell a
  sockpuppet; HN's are enforced by everyone. Do not ask for upvotes, ever.
- **Respect the trademark in both directions.** `TRADEMARK.md`: "for
  Frazaro" and "Frazaro-compatible" are free for anyone; a fork needs its
  own name. Say this up front in the phrasebook jam so nobody is surprised.
- **DCO, no CLA, and the licence table** (`CONTRIBUTING.md`) go in the first
  reply to anyone who offers code, before they write it.
- **Do not promote the Word intake** until `SEC.13` is fixed. It already
  exists — *Import Program File…* accepts `.docx` — which is exactly why it
  must not be advertised: today it opens the document in Word with macros
  enabled. The plain-text paste path is the honest offer.
- **Do not announce `SOLVE()` as existing.** It is scoped; the ask is
  review.
- **The beta is the beta.** Its definition is one sentence
  (`BETA_ROADMAP2.md`'s header): *one named person outside the project runs
  their own SOP, on their own machine, on a Monday, with the owner
  unreachable.* Until that sentence is true, every claim of traction is a
  claim about attention, and the two should never be confused in public.

---

## 9. The calendar

Relative to Week 0's completion. One fixed date: 19 September.

| When | What | Owner artifact |
|---|---|---|
| Week 0 | §1 complete: README accuracy, signed `0.5.3`, `SEC.13` fixed, GIF, `examples/`, Unblock instruction, issue templates, three tutorials, pilot form. | Commit; release `0.5.3` |
| 19 Sep (fixed) | Pirate Day thread on X and Instagram (§4.5). | Thread + GIF |
| Week 1 | Soft launch: r/vba, r/lisp, MrExcel, ExcelForum, SWI-Prolog Discourse. Fix what the first twenty hit. | Patch releases as needed |
| Week 2, Tue–Thu, 8–10 a.m. ET | Show HN (Appendix A.1). Cross-posts same day. Author present in the thread for six hours. | The thread; issues filed from it |
| Week 2 | Awesome-list PRs, GitHub topics, Console.dev submission (§4.11). | PRs |
| Week 3 | r/excel showcase post; YouTube pitches, Spanish channel first (Appendix A.7); EuSpRIG abstract submitted (A.6); the SOLVE challenge posted (§4.9). | Posts, emails, abstract |
| Week 4 | Second Show HN: the phrasebooks angle. dev.to / Hashnode long-form. First *Bring your checklist* live hour (§4.7). | Post; recording |
| Week 5 | Accounting / FP&A / audit: LinkedIn carousel, r/Accounting, r/FinancialModeling, `PI.7` trust survey circulated. *Frazaro solves the case* (§4.6). | Posts; survey; workbook |
| Week 6 | The digest (§6): what refused, what we fixed, who found what, channels retired. Re-bet the channels. Product Hunt decision. | The digest post |

---

## Appendix A — post drafts

### A.1 Show HN (Week 2)

**Title:** Show HN: Frazaro – Prolog, Datalog and SQL as Excel worksheet
functions (offline, no code injected)

**Body:**

> Frazaro is an Excel add-in. Three of its worksheet functions are query
> and logic engines over your actual Tables, returned as spilled arrays:
>
> `=SQL("SELECT Name, Salary FROM staff WHERE Salary > 80000", Staff)` — a
> frozen, SQLite-leaning subset (joins, GROUP BY, recursive WITH); everything
> outside it refuses by name rather than guessing.
>
> `=DATALOG(...)` — function-free Horn clauses, stratified negation,
> aggregates, atoms that address table columns by header name. Provably
> terminating; anything that would break that is refused at parse time.
>
> `=PROLOG(...)` — real unification and backtracking, `not`, `findall`,
> cut; an infinite rule hits a step ceiling and says so.
>
> All three share one substrate (hash join, grouping, spilled output) and
> one reader — VLA, a small Lisp whose target is VBA. The English layer on
> top (sentences that either parse or refuse, in words) is the actual
> product, but it's a longer argument and I'd rather you tried the
> functions first.
>
> Honest limits: Windows Excel for export (Mac interprets); it's a beta
> with a public threat model and ten open findings we filed against
> ourselves last week, listed in the README; no telemetry or network calls
> by standing decision, so I can't see whether you used it — please tell
> me. Apache-2.0 engine, MPL-2.0 phrasebooks, DCO, no CLA.
>
> What I'd most like: a query against your own table that refused, pasted
> as an issue.

**First comment (yours, posted immediately):** the shared-substrate story
in four sentences; why `SOLVE()` isn't built yet and what the nine
dissections are; the `alien.vla` link for the Lisp people; the
`THREAT_MODEL.md` link for the security people. Then stay six hours.

### A.2 r/excel showcase

**Title:** Recursive queries (org charts, BOM explosions) as a single
formula — a free, offline add-in I built [OC]

Lead with the `=DATALOG(...)` ancestor example over a two-column table,
one screenshot of the spilled result. Install steps *in the post*: download,
right-click → Unblock, open, trust the publisher once. Say it's a beta,
say what's open, link the README. Ask for the query they'd write.

### A.3 r/lisp

**Title:** A Lisp that compiles to VBA, with the surface language as data —
and a phrasebook written entirely in glyphs

Lead with `alien.vla`'s header comment and one docstring. Then the
doctrine paragraph (first-match-wins, no backtracking, the loader proves
non-overlap). Ask for review of the "strictly 1:1 with VBA, no semantic
smoothing" decision — it's the one this audience will want to argue.

### A.4 SWI-Prolog Discourse / r/prolog

**Title:** Prolog as an Excel worksheet function: unification and
backtracking over Tables as facts (beta; wants breaking)

Lead with the README's staffing example. State the ceilings honestly:
120-step limit, no tabling, no indexing, no comparison goals yet
(`PROLOG.7`). Ask which of `PROLOG.7`–`PROLOG.9` they'd hit first in a real
KB, and whether the step ceiling is the right termination story.

### A.5 LinkedIn carousel (accounting / FP&A)

Six slides. (1) "Your month-end checklist, as sentences." (2) The five-line
close procedure. (3) A refusal, quoted — "it won't run if it doesn't
understand." (4) Undo. (5) "It never leaves your machine. No network, no
telemetry, by standing decision." (6) "Beta. Windows Excel. Free. Tell me
what refused." Link to the accounting tutorial, not the repo.

### A.6 EuSpRIG abstract (~150 words)

> **Checked English as a spreadsheet control: procedures that refuse to
> run wrong.** Most spreadsheet risk tooling watches risky spreadsheets
> without making them less risky; the manual procedure — the thing that
> actually drifts — stays manual, and its evidence is a human's attestation.
> We present Frazaro, an open-source Excel add-in in which a procedure is
> written as sentences in a small, deterministic, published English
> grammar; every sentence is checked before execution and refused, with an
> explanation, otherwise. The sheet of sentences is the procedure document;
> execution snapshots first; an export path yields readable VBA for review.
> We describe the grammar's determinism guarantee, the refusal-as-feedback
> design, the no-network posture, and the product's own published threat
> model, including open findings. We report the beta's first external
> results and invite the community to define what evidence such a
> procedure would need to count as a control rather than an attestation.

### A.7 Email to an Excel YouTube channel (Spanish channel first)

Five lines. Who you are; one sentence on what it is; the five-minute demo
file attached and a 30-second GIF linked; "you keep full editorial control
— including saying it's not ready"; the Espanol edition, if it's a
Spanish-language channel. No feature list. No follow-up before ten days.

### A.8 Talk Like a Pirate Day thread (19 September)

Post 1: the README program's first three lines, in pirate, from
`pirate.vla` — `"stow {e:expr} in the cell {r:cell}, arr"` is the shape —
with the GIF of it running. Post 2: the same VBA underneath, "the compiler
doesn't care what language you swear in." Post 3: the alien phrasebook, one
docstring, "first contact, in parentheses." Post 4: the download and the
Unblock instruction. Do not explain the joke.

---

## Appendix B — issue templates and Discussion categories

Three templates for `.github/ISSUE_TEMPLATE/`. The specimen is the first
field in each, on purpose (`SD-7`).

**`sentence-expected-to-work.md`**

```markdown
---
name: A sentence I expected to work
about: Frazaro refused or misread a sentence you think it should understand
labels: sentence, grammar
---

**The sentence, exactly as you typed it** (one line; don't tidy it):

**What Frazaro said** (paste the refusal text, or describe what it did):

**What you meant it to do:**

**Frazaro version** (click **Copy Feedback** in the ribbon and paste — the report's second line is `Frazaro <version>`; or read the first line of the Known Sentences sheet):

**Edition:** English / Espanol    **Backend:** Interpret / Compile / both
```

**`refusal-was-wrong.md`**

```markdown
---
name: A refusal that was wrong
about: The sentence was refused, and the refusal did not tell you what to write instead
labels: refusal, messages
---

**The sentence, exactly as you typed it:**

**The refusal text, exactly as shown:**

**What would have told you what to do?** (a rewording, or "I still don't know"):

**Frazaro version:**
```

**`phrasebook-i-wrote.md`**

```markdown
---
name: A phrasebook I wrote
about: Share a .vla phrasebook, or report a problem authoring one
labels: phrasebook
---

Your phrasebook is yours (MPL-2.0 is per-file; see PHRASEBOOK-TERMS.md).
Attach or link the .vla, and tell us:

**What domain / profession is it for?**

**One sentence it handles that english.vla does not:**

**Anything that refused while loading it** (paste the message):

**May we link it from the README's community list?** yes / no
```

Discussion categories: **Sentences** (the standing *Adopt a sentence*
thread and *Port my macro*), **Show and tell** (workbooks, screenshots,
the SOP swap, the refusal leaderboard).

---

## Appendix C — the `examples/` workbooks

One per door, each self-contained: Instructions sheet pre-filled, data on
a second sheet, "open → Check → Interpret → Undo" as the whole tutorial.

| File | Audience | Contents |
|---|---|---|
| `01-first-sentences.xlsx` | Everyone | The README program, verbatim, with a deliberate typo on line 4 so the first thing a user sees is a refusal that teaches. |
| `02-month-end-close.xlsx` | Accounting | A trial-balance sheet and a five-sentence close check: clear formatting, bold headers, sum, out-of-balance flag, timestamp. |
| `03-handover.xlsx` | FP&A | A weekly report-assembly procedure, twelve sentences, with a second tab "how to read this" that is the handover document — i.e., the same sentences. |
| `04-org-chart-datalog.xlsx` | Datalog / data | An `Employees` table with `manager_id`; the two-rule transitive-closure `=DATALOG(...)` from the roadmap's killer case. |
| `05-staffing-prolog.xlsx` | Prolog | The README's shifts/staff/leave tables and the `can-cover` rule, plus the `findall` variant. |
| `06-sales-sql.xlsx` | SQL | One Staff table; six `=SQL(...)` cells escalating from `SELECT` to recursive `WITH`, and one cell that hits `sql-outer-join-not-supported` on purpose. |
| `07-rota-solve-challenge.xlsx` | ASP | The rota problem as data only — no formula yet — with the challenge text (§4.9). |
| `08-polyglot.xlsx` | Lisp / educators | The same five-line program in English, Spanish, Latin, Esperanto, and pirate, one sheet each, with the identical generated VBA shown beside. |
| `09-port-my-macro.xlsm` | VBA developers | One ugly recorded macro; the sentences that replace it; the VBA Frazaro emits. The template for §4.2. |

---

## Appendix D — staleness found while writing this

Filed here so it is not lost; each belongs in a Week 0 commit.

- ~~`README.md` *Known open security items*: lists `SEC.1` as open (closed
  in `0.5.2`); omits `SEC.7` and `SEC.8`–`SEC.17`.~~ **Fixed 2026-09-08**,
  the same day this file was written.
- ~~`DEPLOY.md` "Before either download": names `SEC.1`/`SEC.2` as open.~~
  **Fixed 2026-09-08.**
- `README.md` *Install*: the installer "returns with the next release that
  rebuilds and signs it" — two releases on, it has not.
- `VENTURE.md` §9: "`=SQL()` through INNER JOIN" — now through recursive
  `WITH`; `=PROLOG()` unmentioned. Dated 2026-08-31, so honest, but the
  investor-facing status line should move when the product does.
- No `examples/`, no images, no issue templates, no Discussions
  configuration exist in the repository today (§1).
