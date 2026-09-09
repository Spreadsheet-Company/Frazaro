# Frazaro

## What is a beta?

*With apologies to Kierkegaard, whose* Either/Or *opens with a poet
in the same predicament.*

> What is a beta? An unfinished product which conceals profound gaps
> in its features, but whose interface is so styled that when it is
> clicked it purrs like a functional program. Its fate is like that
> of the unfortunates whom the tyrant Phalaris shut inside his brazen
> bull and roasted over a slow fire: their cries could not reach him
> as cries, but arrived at his ear as sweet music. So too the roadmap
> groans, and by the time the groan has passed through a ribbon
> button it sounds like a feature. And people crowd about the beta
> and say, "Do release again soon" — which is to say, "May new gaps
> open in your grammar, but may your buttons stay fashioned as
> before; for the roadmap would only distress us, but the purring,
> the purring is delightful." And the reviewers come forward and say,
> "Quite right; that is how it should be, according to the principles
> of software engineering." Now a reviewer resembles a developer to a
> hair, except that they have no gaps in their heart and no purr on
> their lips. I tell you, I would rather ship a beta and be
> understood by the roadmap than ship a 1.0 and be misunderstood by
> men.

Less lyrically: this is a known-unfinished program, released
unfinished on purpose, because the only way to learn which sentences
real people reach for is to let real people reach for them.
Bullet-proofing software for industry adoption is a hydra — every
gap closed uncovers two that the closing revealed — and the heads
are counted, in order, in
[docs/BETA_ROADMAP2.md](docs/BETA_ROADMAP2.md) (current) and
[docs/BETA_ROADMAP1.md](docs/BETA_ROADMAP1.md) (its predecessor,
kept for the trail). Before reporting a missing feature, please check
whether it is already a head on that list. If it is, the complaint is
heard and queued. If it is not, that is a genuinely useful report,
and exactly what the beta is for.

---

## What is Frazaro?

**A phrasebook for your workbook.** Frazaro lets people who are not
programmers automate Excel by writing English sentences — sentences
that are *checked* before they run, refuse with an explanation when
they don't parse, and mean exactly one thing when they do.

```text
Work on sheet Output.
Put "Test Report" into cell A1.
Make cell A1 bold.
Create a number called total.
Set total to 0.

Repeat 5 times:
  Increase total by counter.
  If counter is divisible by 2, log "even step " joined with counter.

Put total into cell B2.
Set grand to sum of range B2:B3.
If grand is greater than 40:
  Put "PASS" into cell C4.
```

That is not pseudocode. It is a working excerpt from this repository's
regression corpus: you type sentences like these into a worksheet,
press **Check Instructions** to have every row validated (errors land on
their exact row, in words), then **Interpret Instructions** to execute —
with a snapshot taken first and an **Undo Last Run** button behind it.
(**Compile Instructions** does the same through generated VBA, when you
want the artifact.)

---

## The idea

> The language in which a person thinks (source thought) should be
> the same language in which a person works (source code).

Everyone who lives in spreadsheets has procedures — month-end
checklists, report formatting, data cleanup — that they can *describe
in one breath* and cannot automate without learning VBA. The last few
years added a new option, the AI prompt box, which accepts anything
and is confidently wrong just often enough to matter. Frazaro takes
the opposite bet:

**A small, checked English is better than an unlimited, guessed one.**

- Every sentence either parses against a published grammar or is
  refused *before anything runs* — there is no "it did something,
  hopefully what you meant."
- One sentence has one meaning. The grammar is deterministic by
  construction (first match wins, no backtracking), so the loader can
  *prove* two rules don't overlap rather than hope.
- Refusals teach. A failed sentence gets told what was understood,
  where expectation diverged, and what to write instead — in the same
  register a colleague would use. A real one, from the Datalog engine:

  > fact 'parent' uses 'X', which looks like a variable (it starts
  > with a capital letter) — facts must be fully specific; did you
  > mean to write a rule instead?

## How it works

Three layers, each one honest about what it is:

1. **English grammar** — sentence shapes plus phrase rules loaded from
   *phrasebooks* (data files, not code). A rule is a pattern and a
   template:

   ```text
   "make cell {r:cell} {d:bold|italic}"   →   (make-{d} (range {r}))
   ```

   So `Make cell A1 bold.` translates mechanically — no statistics,
   no generative black box.

2. **VLA** — the middle layer the templates emit: s-expression VBA
   with strictly 1:1 semantics ("VBA wearing parentheses"). No
   semantic smoothing, ever — the moment the middle layer gets nicer
   than VBA, you are maintaining a real compiler alone, forever. This
   project declines that trapdoor on principle and keeps declining it.

3. **Execution, two ways.** By default your program is **interpreted**
   in place — no code is injected into your workbook, no scary trust
   settings are required. When you want the artifact, **Show me the
   VBA** / export produces readable VBA modules: the procedure you
   wrote in English, as inspectable source. Both backends are held to
   parity by a golden-file test corpus; a behavioral difference
   between them is a failing test, not a surprise.

## What makes it different

- **Deterministic, not probabilistic.** Against Copilot and its
  cousins: Frazaro cannot hallucinate. It can only do what a sentence
  provably says, or refuse in words. For procedures that touch real
  numbers, that is the entire point.
- **Auditable by construction.** The sheet of sentences *is* the
  program. What a reviewer reads is what ran — and the export path
  hands them the generated VBA if they want to go deeper.
- **Offline, permanently.** Frazaro makes no network calls — no
  update checks, no telemetry, nothing — by standing decision, not by
  current accident. What you download is the whole product, and it
  keeps working whether or not this project does.
- **No trust toggles by default.** The interpreted runtime needs none
  of Excel's frightening VBA-project trust settings. Only the explicit
  "give me VBA modules" path touches the VBA project, and it says so.
  (This is not "zero trust" in the security-architecture sense; what a
  loaded phrasebook can reach is stated plainly under *Known open
  security items* below.)
- **Undo means undo.** Runs snapshot first. The undo button restores.

## Query and logic (the fun recent turn)

The same engine now ships **worksheet functions** for querying and
logic programming over your actual tables — spilled dynamic arrays,
no add-ins beyond Frazaro itself, no code injected anywhere:

```text
=SQL("SELECT Name, Salary FROM staff WHERE Salary > 80000", Staff)
```

Real SQL text (a frozen, SQLite-leaning subset — everything outside
it refuses by name rather than guessing), read straight from your
Table, returned as a spilled array with headers.

```text
=DATALOG("(fact (parent tom bob)) (fact (parent bob liz))
          (rule (ancestor X Y) (parent X Y))
          (rule (ancestor X Z) (parent X Y) (ancestor Y Z))
          (query ancestor)")
```

Recursive queries — org charts, bills of materials, anything
reachability-shaped — that plain formulas cannot express. Atoms can
address table columns *by header name*, so `(staffing (name X)
(salary S))` reads the columns your table actually has.

```text
=PROLOG("(rule (can-cover Shift Who)
           (shifts (shift Shift) (needs Cert) (minlevel Min))
           (staff (name Who) (cert Cert) (level Level))
           (>= Level Min)
           (not (leave (name Who) (shift Shift))))
         (query (can-cover Shift Who))", Shifts, Staff, Leave)
```

Full unification with backtracking, so the question can be *shaped
like the policy*. Three ordinary Excel Tables — the shifts, with the
certification and the seniority each one needs; who holds which
certification, and at what level; who is on leave when — become facts,
and one rule states the staffing policy in the same words a supervisor
would: a person can cover a shift if they hold the certification it
needs, are cleared to at least the level it asks for, and are not on
leave that day.
The result spills as a roster of every legal pairing, one row per
solution, in derivation order. Swap the last line for
`(query (shifts (shift Shift) (needs C)) (findall Who (can-cover
Shift Who) Bag))` and you get one row per shift with the candidates
gathered into a list. Negation-as-failure, arithmetic via `(is ...)`,
numeric comparison (`<`, `>`, `=<`, `>=`, `=:=`, `=\=`) as goals in
their own right, cut, and `findall` are all in; an infinite rule is
stopped by a step ceiling and refused by name, never left spinning.
`SOLVE()` (answer set programming) is scoped and coming.

## Phrasebooks all the way down

English is not the product — English is the **first phrasebook**. The
grammar is data, so the same machinery loads other surfaces: the
repository carries demonstration phrasebooks for Esperanto, French,
German, Spanish, Danish, Latin — and pirate, and one for an alien —
as proof of the seam, not finished translations. (*Frazaro* is
Esperanto for "phrasebook." The name is the roadmap.)

Phrasebook authors get real tooling: patterns with alternations and
optionals, typed reference slots that shape-check at Check time,
`override:` with provenance, `fail:` proofs that pin refusal wording
at load, and a shadow audit that refuses ambiguous rule sets outright.

## Today

*This section is dated — 2026-08-31, approaching the `0.5.0` beta —
and expects to be rewritten as the project moves. The sections above
it should barely change; this one should.*

- **Works now, owner-verified:** the full
  English → VLA → interpret/compile pipeline; the worksheet IDE
  (Check / Interpret / Compile / Undo / Known Sentences); control flow, value
  actions, lookups, list and range operations; the four-layer error
  model; `=SQL()`, `=DATALOG()`, and `=PROLOG()`; the build that
  produces the distributable add-in and refuses defective
  phrasebooks.
- **Honest limits:** Windows desktop Excel is the platform; Mac runs
  the interpreted path but cannot export VBA (the object model
  doesn't exist there). Grammar coverage is growing weekly and is
  not yet complete for any one profession's vocabulary — that is
  what the beta is for.
- **Who builds this:** one developer, pair-programming with an AI
  assistant, under an unusually heavy testing discipline (self-test
  suite, golden files, dual-backend parity, build-time phrasebook
  audits) precisely *because* of that arrangement.
- **No external users yet.** The project is at the
  looking-for-its-first-pilot stage: one person, one real SOP, one
  workbook. If that could be you, read on.

## Install

Two paths, both documented in [docs/DEPLOY.md](docs/DEPLOY.md):

- **Standalone add-in** — the universal default; it is an Office
  document, not a program, so it works wherever Excel does, including
  managed machines. Download the latest build directly:
  [Frazaro_English.xlam](https://github.com/Spreadsheet-Company/Frazaro/releases/latest/download/Frazaro_English.xlam)
  or
  [Frazaro_Espanol.xlam](https://github.com/Spreadsheet-Company/Frazaro/releases/latest/download/Frazaro_Espanol.xlam)
  (every release is listed under
  [Releases](https://github.com/Spreadsheet-Company/Frazaro/releases)).
- **Installer** (`FrazaroSetup.exe`) — one-click setup with a normal
  Windows uninstall entry, for machines you control. English edition
  only. `0.5.0`–`0.5.2` shipped the standalone add-ins alone; from
  `0.5.3` the release script refuses to publish without a freshly built,
  signed installer, so it is a release asset beside them:
  [FrazaroSetup.exe](https://github.com/Spreadsheet-Company/Frazaro/releases/latest/download/FrazaroSetup.exe).

Either way, removal is one honest in-product button.

## The shelf

This repository's [docs/](docs/) folder is unusually complete —
strategy, standing decisions, campaign histories of every hard bug,
and a set of adversarial self-reviews (a consultant's audit, a
premortem, a devil's-advocate pass, a succession audit, a platform
history) commissioned against the project's own blind spots. If you
want to evaluate the engineering culture before the code, start with
[docs/LESSONS.md](docs/LESSONS.md); if you want the current plan,
[docs/BETA_ROADMAP2.md](docs/BETA_ROADMAP2.md).

## Known open security items

Frazaro's own threat model ([docs/THREAT_MODEL.md](docs/THREAT_MODEL.md))
is public, and so is the list of what it has not closed yet. Items are
listed here by roadmap ID so a downloader hears it from this page rather
than from the repository. Frazaro does not update itself and makes no
network call, so a copy you download today stays as it is until you come
back; check this section or the roadmap to see when each closes.

**Closed.** SEC.1 (`0.5.2`): a member reference the interpreter does not
recognize now refuses in words instead of falling through to VBA's own
late-bound dispatch. SEC.2: a phrasebook rule marked `raw` (literal VBA)
shows an explicit consent dialog, naming the phrasebook, before it loads
from disk — declining refuses the whole phrasebook, not just the
`raw`-bearing rules. SEC.13 (`0.5.3`): *Import Program File…* now opens
Word documents with macros force-disabled, so a document that carries an
`AutoOpen` macro no longer runs it when Frazaro reads the text out.
SEC.8 (`0.5.3`): a workbook that Windows has marked as having come from
the internet can still be read and edited as usual, but a program in its
cells is now refused when it tries to reach OUTSIDE the workbook —
opening or saving files, printing, exporting a PDF, composing an email,
or password-protecting a sheet. To allow it, unblock the file
in Windows first (right-click the file, Properties, tick Unblock);
Frazaro deliberately has no button of its own for this, so that a
workbook arriving from outside cannot carry its own permission slip.

**Open from the original threat model:**

- **SEC.3 — generated code does not yet carry phrasebook provenance.**
  Emitted VBA says what it does, not which phrasebook layer introduced
  each line.
- **SEC.7 — verbs with real external effect are not yet permissioned.**
  A phrasebook you load can open or save workbooks to a path it names,
  export a sheet to PDF, and compose an Outlook email (it is displayed
  for you, never sent silently) — with no consent prompt.

**Open from the project's own code review of 2026-09-08.** The review found
ten items, SEC.8 through SEC.17. **Four are fixed**, all in `0.5.3` —
above: SEC.8, SEC.13, **SEC.11** (the fingerprint your consent is keyed to
was weak enough to forge, and is now SHA-256) and **SEC.9** (a grammar file
beside a workbook, or a path a workbook remembers, now needs this computer's
approval before it can override the built-in grammar). Four were assessed
and **accepted** rather than fixed,
with the mitigating control written down and a stated condition that reopens
each — see *Assessed and accepted* below. Two remain open and are listed
here, most-severe first. Both are audit findings read from the code rather
than exploits anyone has run. Each one's file, line, and fix is in
[docs/BETA_ROADMAP1.md](docs/BETA_ROADMAP1.md). In plain words:

- **SEC.10** — the "remember my consent for this workbook" record is
  stored inside the workbook, so a workbook someone sends you can arrive
  with consent already granted.
- **SEC.15** — formulas a program writes are not screened for functions
  that reach the network or the shell (`WEBSERVICE`, DDE).

**Assessed and accepted — deliberately not fixed, and why.** Each of these
needs a precondition an ordinary install does not meet. The full reasoning,
and the condition that would reopen each one, is in
[docs/BETA_ROADMAP1.md](docs/BETA_ROADMAP1.md); in short:

- **SEC.12** and **SEC.17** — both are on the *Compile* path, which refuses
  to run at all unless you have turned on Excel's *Trust access to the VBA
  project object model*. That setting is off in every Office install by
  default, and managed IT departments routinely disable it outright.
  Frazaro never asks you to turn it on, and *Interpret* — the ordinary way
  to run a program — does not touch it.
- **SEC.14** — a runaway program can hang Excel, and Ctrl+Break stops it.
  A deeply nested program can instead overflow VBA's stack and crash Excel,
  which Ctrl+Break cannot stop and which can lose unsaved work. Accepted
  because the worst outcome is a lost session rather than a compromise:
  nothing runs, nothing leaves the machine, nothing persists.
- **SEC.16** — replacing the grammar files beside the add-in requires
  already being able to run programs on your machine as you. It makes an
  existing compromise durable; it does not create one.

Until these close: **load phrasebooks only from people you would accept
a macro-enabled workbook from — and treat a workbook someone sent you
the same way before you press Interpret.** The phrasebooks embedded in
the downloads are audited at build time; a `.vla` file someone sends you
is not. Found something? [docs/SECURITY.md](docs/SECURITY.md) says where
to report it and what response to expect.

## License and status

Frazaro is open source. The engine is **Apache-2.0**; the phrasebooks
are **MPL-2.0** (file-scoped: a phrasebook you write in your own file is
yours, see [PHRASEBOOK-TERMS.md](PHRASEBOOK-TERMS.md)); the one module
Frazaro copies into your workbook is **0BSD**; and everything Frazaro
generates from your sentences is yours outright, see
[OUTPUT-EXCEPTION.md](OUTPUT-EXCEPTION.md). Full texts are in
[LICENSES/](LICENSES/), the per-file map is [REUSE.toml](REUSE.toml),
and `tools/check_spdx.ps1` keeps the two honest. The name is a
trademark; see [TRADEMARK.md](TRADEMARK.md). Contributions are welcome
under the DCO, no CLA; see [CONTRIBUTING.md](CONTRIBUTING.md).

**No warranty.** This is beta software, released unfinished on purpose
(see the top of this file). It is provided as-is, without warranty of
any kind, as the licences say in longer words. Take a backup before you
run a program you have not run before; Frazaro snapshots before every
Run and offers Undo, and that is a convenience, not a guarantee.

**What a program can do with your data.** Frazaro itself makes no
network connection, ever, and phones nothing home. A program *you* run
can do what the sentences say: read and write cells, sheets and files,
and, if a sentence asks for it, open a draft email in your own Outlook
for you to send (Frazaro never sends mail itself). A phrasebook can
define new sentences, and a phrasebook rule marked `raw` can contain
arbitrary VBA; load phrasebooks from people you trust, the same way you
would open a macro-enabled workbook from them.
[docs/THREAT_MODEL.md](docs/THREAT_MODEL.md) says this at length and
without flattery.

The project is at the looking-for-its-first-pilot stage, and it *is*
looking for that first pilot user, and for skeptical readers of the docs
above.
