# Releases

*Newest first. `tools/release.ps1 -Version X.Y.Z` publishes the section headed `## X.Y.Z` as that release's notes and refuses to run without one, so the notes are written before the release, never after. Cadence: a `0.5.N` patch at the end of each working day, a `0.N.0` minor at the end of each week; security and safety fixes ride the patches, larger features the minors. Keep the *Known open security items* block in every section until the items close.*

## 0.5.1

*(next: write this before running the release)*

## 0.5.0 — 2026-09-05

**A phrasebook for your workbook.** Frazaro lets people who are not
programmers automate Excel by writing English sentences that are *checked*
before they run, refuse with an explanation when they don't parse, and mean
exactly one thing when they do. This is a known-unfinished program, released
unfinished on purpose, because the only way to learn which sentences real
people reach for is to let real people reach for them. The open items are
counted, in order, in `docs/BETA_ROADMAP2.md`.

### Download

| File | Choose this if |
|---|---|
| `Frazaro_English.xlam` | The universal default. An Office document, not a program: works wherever Excel does, including managed machines. |
| `Frazaro_Espanol.xlam` | The Spanish edition (bilingual: every English sentence still works). |

Register the add-in once from inside the product (**Register for auto-load**
on the Frazaro ribbon) or through Excel's Add-ins dialog. Removal is one
button, **Uninstall Frazaro**. Full instructions: `docs/DEPLOY.md`. The
Windows installer is not part of this release; it returns in the next one.

### What works, owner-verified

The full English → VLA → interpret/compile pipeline; the worksheet IDE
(Check / Run / Undo / Known Sentences); control flow, value actions,
lookups, list and range operations; the four-layer error model; `=SQL()`,
`=DATALOG()` and `=PROLOG()` worksheet functions over your own tables. All
under a golden-file, self-test and dual-backend-parity discipline.

### Honest limits

Windows desktop Excel is the platform; Mac runs the interpreted path but
cannot export VBA. Grammar coverage is growing weekly and is not yet
complete for any one profession's vocabulary. No external users yet: the
project is looking for its first pilot.

### Known open security items

Frazaro's own threat model (`docs/THREAT_MODEL.md`) is public, and three of
its findings are open in this release:

- **SEC.1** — dynamic dispatch is not yet capability-gated: a phrasebook you
  load can reach roughly what a macro in a workbook you open could reach.
- **SEC.2** — `raw` phrasebook rules (literal VBA) run without a consent
  prompt. Refused on the default Interpret path; executes only through
  Compile/Export, behind Excel's own VBA-project trust prompt. First item
  after this release.
- **SEC.3** — generated code does not yet carry phrasebook provenance.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from.** The phrasebooks embedded in these downloads
are audited at build time. Frazaro makes no network call and does not update
itself; check the README's *Known open security items* when you return for
a newer build. Vulnerability reports: `docs/SECURITY.md`. Everything else:
`docs/SUPPORT.md`.

### Licence

Engine Apache-2.0; phrasebooks MPL-2.0 (a phrasebook you write in your own
file is yours); the one module Frazaro copies into your workbook is 0BSD;
everything Frazaro generates from your sentences is yours outright
(`OUTPUT-EXCEPTION.md`). No warranty: beta software, as-is. Take a backup
before you run a program you have not run before.
