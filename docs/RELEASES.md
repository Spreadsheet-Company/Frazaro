# Releases

*Newest first. `tools/release.ps1 -Version X.Y.Z` publishes the section headed `## X.Y.Z` as that release's notes and refuses to run without one, so the notes are written before the release, never after. Cadence: a `0.5.N` patch at the end of each working day, a `0.N.0` minor at the end of each week; security and safety fixes ride the patches, larger features the minors. Keep the *Known open security items* block in every section until the items close.*

## 0.5.2

### What changed

- **`SEC.1` — the `CallByName` reflection fallback removed from dynamic
  dispatch, built and owner-verified live.** Previously, any `.`-member
  name the interpreter didn't recognize natively fell through to VBA's
  own `CallByName`, reaching the real Excel object model with whatever
  name and arguments a phrasebook rule supplied — a heuristic, not a
  capability gate. A full-repo census found 25 members reached only
  through it, well beyond the G-TABLES surface this item's own design
  anticipated: plain cell `.value` reads, cross-sheet `.range` lookups,
  the `listobjects`/`listrows`/`listcolumns` chain, and several
  housekeeping macros. Each got its own fixed, audited native case before
  the fallback itself came out, so anything not on that list now refuses
  in words instead of reaching arbitrary late-bound dispatch. Two
  corrections found live, not assumed clean: a corpus-only first census
  missed `.value`/`.range` entirely (caught only because the host-test
  suite reads them, not the shipped phrasebook text), and a second pass
  then found every property already native for *writing* (`name`,
  `bold`, `size`, and 17 others) had never been native for *reading* —
  fixed by mirroring the write-side list into the read side in one pass
  rather than chasing each individually. `VLA_SELF-TESTS` pure 947/947,
  host 143/143; `VerifyReports` emitter 141/141, interpreter 141/141,
  both live in Excel. G-PIVOT never used this mechanism at all (it
  dispatches through dedicated runtime helpers); keyword-argument calls
  already had no fallback to begin with. Permissioned, declared
  capabilities with real external effect (`vlasendmail` today) are a
  separate, still-open item — see `SEC.7`. Full mechanism:
  `docs/BETA_ROADMAP1.md`'s own SEC.1 entry.
- **`SEC.2` — `raw` behind explicit, per-phrasebook consent, built and
  owner-verified live.** A phrasebook using `raw` (literal VBA,
  previously unconsented) now shows a modal, naming the file, before it
  loads from disk; declining refuses the whole phrasebook, not just the
  `raw`-bearing rules. Two remembered scopes, an explicit choice rather
  than a silent default: *this workbook only* (safer — forging it needs
  write access to that one file) or *every workbook on this device*
  (more convenient, a wider target, named as such in the prompt
  itself). Gated at the file-path loader specifically, not the shared
  `EnglishLoadVocabularyText` primitive `VLA_Browser.bas`'s
  already-shipped host-free translate API calls directly and documents
  as never showing a dialog — an early draft got this wrong and only
  passed the purity ratchet on a technicality, caught before it
  shipped. No test-bypass toggle anywhere in the mechanism, by design.
  Full mechanism: `docs/BETA_ROADMAP1.md`'s own SEC.2 entry.
- **`GO.6` — a working Load Phrasebook button, owner live-tested.**
  Found live while hand-verifying `SEC.2`: the only mechanism that
  technically loaded an external phrasebook (`VLA_IDE.IdeVocabPath`'s
  four candidate paths) was undocumented, built for internal
  edition/dev purposes, and REPLACED the base corpus by exact filename
  match rather than adding to it — no ribbon command existed for an
  org admin or community contributor to load their own. The new "Load
  Phrasebook" button calls `EnglishLoadVocabulary` directly, so it
  inherits `SEC.2`'s raw-consent gate and `G3`'s same-shape-collision
  refusal automatically, no second loading path. ADDS rather than
  replaces — `G3`'s existing cross-file override mechanism already
  resolves collisions between sources, `GO.1`'s ratified precedence,
  no interpreter change needed — persists per workbook (one
  `VLA_LoadedPhrasebooks` custom document property, an unbounded
  vbLf-joined list; no cap, the same as a source file's own import
  statements), and shows what's currently loaded after every load
  (`EnglishLoadedSourcesReport`). A moved or deleted remembered
  phrasebook is skipped with a note rather than blocking every other
  command; a genuine content collision still refuses, unchanged. Full
  mechanism: `docs/BETA_ROADMAP1.md`'s own GO.6 entry.
- **The one `AS.1` gap closed: `paint cell {r:text}`.**
  `check_rule_coverage.ps1`'s first real report (after `GEXPANDERLINT.0`
  verticalized the phrasebook artifact) found a rule with zero
  test-success proofs. It had never worked: the rule called `vlacolr`
  (no "o"), a name that resolves nowhere in the shipped modules, and
  had no cell/range slot in its pattern at all, so it could never have
  painted a cell even with the spelling fixed. Corrected to the
  `set-fill-color` idiom every sibling color rule already uses
  (`pirate.vla`'s own "paint cell" rule confirmed the intended
  semantics) and given its missing test. 150/150 phrasebook rules now
  carry at least one proof.

### Known open security items

- **SEC.3** — generated code does not yet carry phrasebook provenance.
- **SEC.7** — a small set of verbs with real external effect (`vlasendmail`
  today) still runs with no permission check: a phrasebook you load can
  send mail with no consent prompt, the same as before this release.

`SEC.1`, `SEC.2` closed this release — see above.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from.** Frazaro makes no network call and does not
update itself; check the README's *Known open security items* when you
return for a newer build. Vulnerability reports: `docs/SECURITY.md`.
Everything else: `docs/SUPPORT.md`.

## 0.5.1 — 2026-09-07

**Pre-flight for the corporate push.** The first release cut by
`tools/release.ps1` rather than by hand, and the day the static ratchets
went green again. Nothing here changes what a sentence means; every item
is the machinery that keeps the next four weeks honest.

### What changed

- **The static ratchets pass, and are now run on every push.** `0.5.0`
  shipped past a red `check_raise_ratchet.ps1` (five raw `Err.Raise` sites
  had accumulated since the ceilings were set) and a `check_id_registry.ps1`
  that misread four mid-sentence bold ids as duplicate definitions. The
  four genuinely raw refusals in `VLA_Relation`/`VLA_Sql` now carry
  catalogue ids (`relation-table-noncontiguous-areas`,
  `sql-internal-*`); the one re-raise with no id to carry is documented as
  the eighth site of `VLA`'s ceiling; the id registry counts a bold id as
  a definition only when no prose precedes it on its line.
- **Three ribbon buttons find the phrasebook again** (EDITION-VOCABPATH).
  *Translate to VLA / to VBA*, *Export Expanded Phrasebook* and *Phrasebook
  Test Coverage* had reported "vocabulary file on disk... none found" in
  the dev workbook since the phrasebook moved under `scripts/polyglotta/`.
  They now search that folder as a last candidate; Check/Compile's own
  embedded-chain loading is byte-for-byte unchanged.
- **The expanded-phrasebook staleness stamp ignores whitespace.** It was
  the source file's byte size, so a Lint VLA re-indent or a line-ending
  change made `check_rule_coverage.ps1` demand a re-export, and Beta's
  history holds four commits that changed nothing but that number. It is
  now a hash of the source's non-whitespace bytes, computed identically
  in VBA and PowerShell; only a token change counts as stale.
- **`scripts/english_expanded.vla` re-exported** against the current
  corpus, so `check_rule_coverage.ps1` reports on the phrasebook that
  actually ships. It now lives beside its source under `scripts/polyglotta/`.
- **`release.ps1` refuses the placeholder** release-notes section instead
  of publishing it.

### Known open security items

Unchanged from `0.5.0`, and still the first engineering item in the queue:

- **SEC.1** — dynamic dispatch is not yet capability-gated: a phrasebook you
  load can reach roughly what a macro in a workbook you open could reach.
- **SEC.2** — `raw` phrasebook rules (literal VBA) run without a consent
  prompt. Refused on the default Interpret path; executes only through
  Compile/Export, behind Excel's own VBA-project trust prompt.
- **SEC.3** — generated code does not yet carry phrasebook provenance.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from.** Frazaro makes no network call and does not
update itself; check the README's *Known open security items* when you
return for a newer build. Vulnerability reports: `docs/SECURITY.md`.
Everything else: `docs/SUPPORT.md`.

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
