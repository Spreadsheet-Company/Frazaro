# Releases

*Newest first. `tools/release.ps1 -Version X.Y.Z` publishes the section headed `## X.Y.Z` as that release's notes and refuses to run without one, so the notes are written before the release, never after. Cadence: a `0.5.N` patch at the end of each working day, a `0.N.0` minor at the end of each week; security and safety fixes ride the patches, larger features the minors. Each section carries a short *Known open security items* block: the standing advice, what closed in that release, and a pointer to the authoritative list. It does NOT re-enumerate every open item — that list lives in `docs/BETA_ROADMAP1.md` (full, with dispositions) and `README.md` (plain words), which are edited once rather than copied into every release forever. Sections written before `0.5.3` keep their longer blocks as published; they are history, not a template.*

## 0.5.3

### What changed

- **`SEC.8` — a workbook from the internet can no longer reach outside
  itself.** Office has blocked macros in files that came from the
  internet since 2022. That block reads the Mark-of-the-Web that Windows
  puts on a downloaded or emailed file, and it applies to macros — VBA
  code. A Frazaro program is not a macro: it lives in the cells. So a
  plain `.xlsx` someone mails you could carry a complete program, and
  after the ordinary "Enable Editing" click Frazaro would run it with the
  add-in's own privileges. Office's protection had never applied to it,
  because from Office's point of view there was nothing there to protect
  you from.

  Frazaro now reads that same mark itself, and refuses the things that
  reach outside the workbook: opening or saving a file, saving a copy,
  printing, exporting a PDF, composing an Outlook email, and
  password-protecting or unprotecting a sheet. Everything else runs
  normally — reading and writing cells, formatting, formulas, loops, the
  whole ordinary business of a program. A workbook from the internet
  still opens, still calculates, and can still be edited exactly as
  before. It just cannot reach off the page.

  The refusal names the thing the program tried to do, so you find out
  what a workbook wanted rather than only that something was blocked.

  **Clicking "Enable Editing" does not grant this.** Excel shows its own
  yellow banner on a file from the internet — that bar is Excel's, not
  Frazaro's, and it governs whether you can type in the sheet. Frazaro's
  check is separate and happens later, when a program actually tries to
  reach outside the workbook. Enabling editing leaves the file's origin
  mark exactly where it was, which is why the two decisions stay
  independent: you can read and edit a workbook someone sent you without
  also agreeing that its program may email, print, or save files. Only
  unblocking the file in Windows does that.

  **To allow it, you unblock the file in Windows** — close the workbook,
  right-click the file in File Explorer, choose Properties, tick
  *Unblock*, then reopen it. Frazaro deliberately offers no button of its
  own for this. That is the whole point of the design: a workbook that
  arrives from outside must not be able to carry its own permission slip,
  and anything Frazaro stored could be forged or shipped inside the
  workbook itself. Windows' own checkbox is the one channel a mailed file
  cannot reach.

  Two limits worth knowing, both deliberate. A workbook opened straight
  from a `https://` SharePoint address refuses external effect too:
  Windows cannot record an origin mark in that kind of location, so
  Frazaro cannot tell where the file came from, and it does not guess.
  (A OneDrive folder that syncs to your own disk is an ordinary local
  folder and is not affected — that is the common case.) The workaround
  is to save a local copy. And a file with no mark at all is treated as
  local, because an ordinary file you made yourself has no mark either
  and the two are genuinely indistinguishable — the protection comes from
  the mark being *present*, never from it being absent.

- **A new release check: `tools/check_sec8_provenance_gate.ps1`.** The
  self-test suite does no Office automation and cannot manufacture a file
  that claims to be from the internet, so it can check every part of the
  *decision* but never the call sites. This static check runs at every
  release and fails it if any of the nine external-effect sites loses its
  guard, or if the one line that reads the workbook's origin goes
  missing. It is mutation-tested in both directions rather than assumed
  to work, and it found a site the audit had missed on its first run.
  What it cannot do is prove the guard *functions* against a real marked
  workbook — only Excel can do that, so that rests on a live test.

- **Email problems now tell you what went wrong instead of dropping you
  into the code editor.** If Outlook was not available, or an attachment
  path did not exist, an email sentence did not report that — it stopped
  the program with Visual Basic's own `Run-time error '5'` dialog, the one
  with a *Debug* button. The message Frazaro meant to show you ("Could not
  start Outlook to create the email") existed the whole time and never got
  the chance to appear. This affected every release that has had the email
  sentence, and needing no Outlook installed is all it took to trigger it.

  The cause is a VBA quirk this project had already documented and proved
  with a standalone test: the mechanism the interpreter used to reach that
  particular helper does not carry an error back to the code that would
  have caught and displayed it. The email helper now uses a direct call
  instead, so its messages arrive the way every other refusal in Frazaro
  does.

  **Not fixed in this release, and named so it is not mistaken for
  shipped:** eight other helpers still reach you the same wrong way when
  they fail — an unrecognised colour, a missing key, and six pivot-table
  and worksheet operations with an unrecognised option. Those show
  Visual Basic's dialog rather than a Frazaro message. Only the email
  helper is fixed here, because only it was in the way of `SEC.8`; the
  rest are recorded in
  [`docs/BETA_ROADMAP1.md`](BETA_ROADMAP1.md) with the full list rather
  than fixed in a hurry alongside a security change.

- **A build defect found by a new release check: `VlaSlice` was never
  shipped.** Frazaro keeps two lists of modules — one for what a built
  add-in contains, one for what the development workbook reloads — and
  nothing had ever checked that they agree. `VlaSlice`, a small internal
  class the language core names directly, was in the second list and not
  the first, so a freshly built add-in would not have contained it at
  all. It has been in that state since `0.5.0`. It is now in both lists.

  This is the ninth time this project has made the same mistake, and the
  first time a machine caught it rather than a person hitting the
  resulting error. `tools/check_devrig_mods_parity.ps1` now runs at every
  release and fails it if either list is missing something the other has,
  in either direction — the more dangerous direction being a module the
  development workbook compiles happily while a built add-in would not,
  which is exactly what stayed hidden here. Modules that genuinely should
  not ship (the test suites, the builder itself) are listed in the check
  with a reason beside each. It is mutation-tested in both directions.

- **`SEC.13` — Word documents are no longer opened with macros enabled.**
  *Import Program File…* accepts Word documents, and it opened them
  through Word automation without setting Word's own
  `AutomationSecurity`. Word's default in that mode is *Low*, so a
  `.docm` carrying an `AutoOpen` or `Document_Open` macro ran that macro
  silently the instant Frazaro read the text out — no Trust Center
  prompt, because a document opened by automation does not get one. This
  was a live, one-click path in every shipped edition, on the menu and
  the ribbon both, and worse than the audit first recorded: the import
  router matches `.docm` by name, not only `.docx`/`.doc`.

  Frazaro now force-disables macros for the duration of the read, and the
  guard is deliberately set *after* the error handler is armed, so if it
  cannot be set the import refuses rather than opening the document
  unguarded. Because Frazaro reuses a copy of Word you already have open
  rather than always starting its own, it captures your Word's previous
  setting and puts it back afterwards — an application it does not own is
  handed back as it was found. If that restore should ever fail it stays
  quiet on purpose: the failure leaves Word *more* cautious than before,
  never less, and a Word restart clears it.

  Word documents still import exactly as they did — the text, the
  typography cleanup, everything. The only thing that changed is that a
  document's own macros no longer get to run on the way in.

  If you followed the previous advice in the README (*"import only Word
  files you wrote yourself, or paste the text instead"*), you no longer
  need to. That advice has been removed.

  **Not in scope, and named so it is not mistaken for shipped:** Frazaro
  still does not *ask* you before opening a document you point it at.
  Consent prompts are `SEC.7`/`SEC.8`'s subject and remain open.

- **Importing a Word document can no longer appear to hang Excel.** When
  Frazaro starts its own copy of Word to read a document (rather than
  reusing one you already have open), that copy is invisible. If Word
  decided to *ask* you something about the file rather than simply fail
  — a damaged document, a file-conversion prompt — the question appeared
  on a window you could not see or click, and Excel looked frozen with no
  way forward but Task Manager. Frazaro now tells the copy it starts not
  to raise alerts, so a document Word dislikes comes back as the ordinary
  refusal message instead of a hidden prompt.

  A copy of Word you already had open is deliberately left alone. Its
  windows are visible, so its questions are answerable — and silencing
  alerts in an application Frazaro did not start would cost you warnings
  you should see. That is the opposite direction of failure from the
  macro guard above, which is why the two are treated differently.

  Found while writing the live test for `SEC.13` rather than from a
  report: it never actually fired during testing. It is fixed as a
  hazard, not as an observed fault.

- **A new release check: `tools/check_word_automation_security.ps1`.**
  The self-test suite does no Office automation at all, so nothing in it
  could ever have caught this or its return. A static check now runs at
  every release and fails it if any code that opens a Word document does
  not force-disable macros first, in that same procedure, before the
  open. It is mutation-tested in both directions rather than assumed to
  work. What it cannot do is prove the guard *functions* — only Word can
  do that, so that rests on a live test, with a reproducible fixture
  recipe recorded in `tools/sec13_word_fixture.md`.

- **`SEC.11` — the fingerprint that remembers your consent is now a real
  one.** When a phrasebook contains `raw` — VBA written directly into a
  grammar file — Frazaro asks before loading it, and can remember your
  answer. What it remembers is tied to a fingerprint of that file's exact
  contents, so that editing the phrasebook at all asks you again.

  The fingerprint was too weak for the job. It was a simple arithmetic
  checksum, and checksums of that kind can be *aimed*: someone who wanted
  a different phrasebook to carry your fingerprint could adjust a couple
  of characters inside a comment until the numbers matched. Your "yes" to
  a file you had read could then have been silently inherited by a file
  you had never seen. It is now SHA-256, the same standard used for
  software signatures, where aiming at an existing fingerprint is not
  something anyone knows how to do.

  **You will be asked once more for phrasebooks you had already
  approved.** Old answers were filed under the old fingerprint and cannot
  be matched to the new one. Being asked again is the safe direction to
  fail, so nothing tries to convert them; the old entries are simply left
  alone and ignored.

  Two things deliberately did *not* change. The fingerprint still ignores
  spaces, tabs and line breaks, so re-indenting a phrasebook, or opening
  one that was saved on a different operating system, still counts as the
  same file rather than sending you back through the question. Changing
  what a phrasebook *says*, by even one character, still does.

  And exported *Expanded Phrasebook* files you already have keep
  working. The freshness check reads the older form, confirms the file is
  current, and tells you the next export will upgrade it. A re-exported
  file carries the new form, written out as `sha256:` followed by the
  digest so the two generations can never be mistaken for one another.
  Nothing you already have needs regenerating.

  Built without depending on anything being installed. The obvious route
  was to borrow Windows' own cryptography through .NET. Measured on the
  development machine, that turned out to fail outright — the component
  is registered against a version of .NET that Windows 11 no longer
  installs by default — so a phrasebook would have become unloadable on
  an ordinary machine. Frazaro now computes SHA-256 itself, in about 150
  lines, which works the same everywhere and can be checked completely by
  the self-test suite against the published standard test values.

- **A new release check: `tools/check_hash_twin.ps1`.** The fingerprint is
  computed in two places — in Frazaro itself, and in the release script
  that verifies an exported phrasebook is current. The two must agree
  exactly, and until now the only thing keeping them in step was a comment
  saying so. This project has been bitten nine times by that same shape.
  Both sides are now pinned to the same published test values, so either
  one drifting fails a release rather than being noticed later. The two
  were also checked against each other on the real 89,446-byte English
  phrasebook and produce the identical fingerprint.

- **`SEC.9` — a phrasebook has to be approved on this computer before it
  can replace the built-in grammar.** A phrasebook decides what your
  sentences *mean*. Frazaro used to prefer a grammar file sitting next to
  the workbook you had open over its own built-in copy, and it would
  silently reload any phrasebook path a workbook remembered — including
  paths on other machines. Between them, a workbook could quietly decide
  what every sentence in it did.

  This was not theoretical. On 2026-09-08 an old `english.vla` left over in
  a Downloads folder was picked up ahead of the add-in's own copy, simply
  because the workbook being opened was in that folder too. It failed
  noisily, but only by luck: that copy was old enough not to know a word the
  program used. A phrasebook that was merely *different* rather than *older*
  would have loaded without a word and changed what the program did.

  Frazaro now asks, once, naming the file, before using a phrasebook that
  did not come with it. Your answer is remembered **on this computer**,
  not inside the workbook — a workbook someone sends you cannot arrive with
  its own permission already granted, because the permission was never
  something a workbook could carry. Choosing a phrasebook yourself through
  *Load Phrasebook* counts as the answer, so picking a file never asks you
  about it a second time.

  Saying no is safe and is not a dead end: Frazaro simply uses its own
  built-in grammar, which is complete. The answer is remembered either way,
  so you are not asked again on every command — and if the file itself
  changes later, you are asked again, because the answer was about the file
  you were shown, not about its name.

  **Network and web locations are described, never quietly contacted.** If a
  workbook asks for a phrasebook on a network share, the question says so
  and warns that opening it would hand your Windows sign-in to that server.
  Nothing touches the location until you say yes — not even a check for
  whether the file exists, which is itself enough to leak that sign-in.

  **Changing your mind: *Forget Phrasebook Approvals*,** in the Utilities
  group. It lists every answer this computer has recorded and clears them
  all, so the next time each phrasebook is used you are asked again. Your
  answers are otherwise kept for good, which is deliberate — a phrasebook
  you declined stays declined rather than asking you the same question
  every time you open Excel, because a dialog that keeps reappearing is one
  people learn to click through without reading.

  Your recorded answers also appear in *Copy Diagnostic Report*, so if a
  sentence stops being understood you can see whether a phrasebook it
  needed was declined. That matters because the symptom on its own is
  indistinguishable from a typo: without it, Frazaro would just say it does
  not understand the sentence and never mention the phrasebook.

- **A new release check: `tools/check_sec9_phrasebook_gate.ps1`.** The
  self-test suite can check the whole decision — whether a path is remote,
  whether it belongs to Frazaro itself — but it cannot click a dialog or
  read the saved answers, so the places that *call* the check are where this
  could quietly stop working. A static check now runs at every release and
  fails it if either loader loses its gate, if the two decision functions
  are removed, or if the approval is ever moved to run after the
  file-existence probe rather than before it. That last one is the one that
  matters: reordering those two lines would restore the credential leak
  while leaving every visible behaviour, and every test, exactly as it was.
  It is mutation-tested in both directions rather than assumed to work.

- **Sixteen things a program can get wrong now say so, instead of dropping
  you into the code editor.** Type a colour Frazaro does not recognise, ask
  for a key that was never stored, or name a pivot table that is not there,
  and until now Excel's own *Run-time error '5'* box appeared, with a
  **Debug** button that opened the VBA editor at a line of Frazaro's
  internals. Frazaro had written a perfectly clear explanation for each of
  these — it just never reached the screen. You now get the ordinary Frazaro
  message saying what was wrong with what you wrote: *'notacolor' is not a
  color - use "#RRGGBB", like "#FF69B4"*, or *there is nothing stored at key
  'no-such-key'*, or *no pivot table named 'nosuchpivot' in this workbook.*

  The cause was one mechanism, not sixteen separate bugs. Frazaro reaches
  most of its built-in helpers through a general-purpose Excel facility
  that, it turns out, does not carry an error back to the code that asked
  for it. Any helper whose job includes saying no was therefore unable to
  say no. The helpers that only compute something were unaffected, which is
  why this took so long to notice: the failure was invisible until you made
  a mistake.

  Eight of the sixteen were found by hand. The other eight were found by the
  release check below, and had been missed — they refuse a misspelled pivot
  table name through a shared piece of code rather than in their own, which
  is exactly the kind of thing a person reading down a list does not see.

- **A second fix underneath it: Frazaro no longer decides whether something
  is one of its helpers by trying it and seeing what happens.** It now checks
  the name against the list of helpers first. The old approach could not tell
  "that is not a helper at all" apart from "that is a helper, and it is
  refusing" — so a deliberate refusal could have been reported as *"'vlacolor'
  is not a form..."*, naming the wrong problem with complete confidence. Being
  told the wrong thing firmly is worse than being told nothing, and this
  removes the possibility rather than making it less likely. A genuinely
  unknown name still gets exactly the same "not a form, place helper, dotted
  global, built-in, or VLA_Runtime helper" message it always did.

- **A new release check: `tools/check_runtime_raise_dispatch.ps1`.** The
  boundary this fix relies on — a helper that can refuse must be called the
  direct way — was being kept in someone's head, and had already slipped
  three times, each time found by a user hitting the crash. The check now
  works it out from the source: it reads every helper, follows the shared
  code they call, and fails the release if any helper that can refuse is
  still reached the broken way. It was written before the fix, so its first
  run listed exactly the work to do, and it is mutation-tested in both
  directions rather than assumed to work. It is what found the eight the
  hand count missed.

### Known open security items

**Closed this release:** `SEC.8`, `SEC.9`, `SEC.11` and `SEC.13` — see above.

**Still open:** `SEC.3`, `SEC.7`, and two from the 2026-09-08 code
review — `SEC.10` and `SEC.15`. In plain words:
the remembered raw-VBA consent record still lives inside the workbook
(`SEC.10`);
formulas a program writes are not screened for functions that reach the
network (`SEC.15`); and effects like sending mail still run without a
permission prompt (`SEC.7`). `SEC.8` narrows that last one — it gates on
where the *workbook* came from — but does not close it: a phrasebook loaded
into a workbook of your own still reaches those verbs unprompted.

**Assessed and accepted, not fixed:** `SEC.12`, `SEC.14`, `SEC.16` and
`SEC.17`. Each needs a precondition an ordinary install does not meet —
mostly an Excel setting that ships off and that Frazaro never asks you to
turn on. The reasoning for each, and what would reopen it, is written down
rather than left implied.

The authoritative lists, kept current in one place instead of copied into
every release: [`README.md`](../README.md) in plain words, and
[`docs/BETA_ROADMAP1.md`](BETA_ROADMAP1.md) with the file, line, fix and
disposition for each.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from — and treat a workbook someone sent you the
same way before you press Interpret.** Frazaro makes no network call and
does not update itself; check the README's *Known open security items*
when you return for a newer build. Vulnerability reports:
`docs/SECURITY.md`. Everything else: `docs/SUPPORT.md`.

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
- **`CO.4` — versions are now something Frazaro can compare, not just
  print.** Groundwork, with no button and nothing new to say yet: it is
  what a future phrasebook will be checked against when it declares a
  minimum version (`requires: version:0.5.1`), and what a generated
  module's own version stamp will mean once it carries one. The
  decision behind it is the substance. **The grammar's compatibility
  version is the release version you already see** — the one in Copy
  Feedback and in Add/Remove Programs — under the `MAJOR.MINOR.PATCH`
  rules this project already wrote down (`SD-14`). There is no second,
  separate "grammar version" to learn, and deliberately so: `PATCH` is
  *defined* as no observable change to what your existing sentences do,
  `MINOR` means something new became sayable, and `MAJOR` means a
  sentence that once shipped changed meaning or a program could behave
  differently after upgrading. Those rules already say everything a
  compatibility check needs, so inventing a parallel number would only
  create two things to keep in step. Scoping also corrected a
  long-standing internal misreading: `VLA_CORE_VERSION` looked like a
  grammar version and is not one — it, and 23 sibling constants, name
  the work item that last touched each module, and three modules
  legitimately share one value today. It is unchanged, and no
  compatibility check reads it. Ordering is plain numeric
  `MAJOR.MINOR.PATCH`, so `0.9.0` correctly sorts *before* `0.10.0`
  rather than after it the way plain text comparison would; a version
  that cannot be read (`banana`, or a `-beta` suffix, which this
  project has never used) is refused in words rather than being quietly
  treated as either satisfied or unmet, since both hide the typo.
- **`CO.6` — a written record of when each thing you can say became
  sayable.** New file, and unlike `CO.4` above this one is meant to be
  read: [`docs/GRAMMAR_SINCE.md`](GRAMMAR_SINCE.md) lists every phrasebook
  rule and every core form with the release it first **worked** in — 278
  entries, almost all of them `0.5.0`. It answers the question `CO.4`
  leaves open. Knowing that versions compare correctly does not tell you
  *which* version you need, and the version number alone is a loose
  answer: a release can go up because a ribbon button was added, which
  tells you nothing about whether a particular sentence will work. This
  file is the precise answer, and the thing a future `requires:`
  declaration will be checked against.

  **It records when a form first worked, not when it was first spelled** —
  a distinction with a real case in it. `paint cell` appears here at
  `0.5.2`, not `0.5.0`, even though the words shipped in `0.5.0`: as the
  `AS.1` entry above describes, that rule never once painted a cell. Dating
  it `0.5.0` would tell you a phrasebook using it runs on `0.5.0`, which is
  false. Entries are never edited afterwards, so a date that went in wrong
  would stay wrong — which is why the seed was checked against the actual
  release tags rather than taken from the generated phrasebook artifact.
  That check earned its keep: the artifact looked like it gained three
  rules in `0.5.1`, and it had not — those three were already sayable in
  `0.5.0` and the artifact had simply been stale, the same staleness the
  `0.5.1` notes below record fixing.

- **`F.10` — a phrasebook can now state what it needs, and is refused
  politely when it doesn't have it.** Write a line like
  `(requires-version "0.5.2")` at the top of a phrasebook and Frazaro
  checks it *before* loading anything from that file. If the build is too
  old you get a plain sentence — a phrasebook asking for `0.9.0` on this
  build is refused with "this phrasebook needs Frazaro 0.9.0 or newer;
  this is 0.5.2" — instead of rules that load and then mysteriously do
  the wrong thing. This is what
  [`docs/GRAMMAR_SINCE.md`](GRAMMAR_SINCE.md) above exists to be checked
  against: it tells a phrasebook author which version number to write.

  Nothing loads part-way. The check happens before the first rule is
  registered, so a phrasebook is either fully in or fully refused —
  and it does not matter where in the file the line sits.

  Two other kinds of requirement are understood and both currently
  **refuse**, on purpose rather than by omission.
  `(requires-capability "...")` is the permission system that is not
  built yet (`SEC.7`, below): since nothing can grant a capability, the
  honest answer to a phrasebook asking for one is no. `(requires-form
  "...")` — needing one specific sentence rather than a whole version —
  is understood but cannot be checked yet, because the record above
  lives in the source repository and is not carried inside Frazaro
  itself. A requirement Frazaro doesn't recognise at all is also
  refused: it cannot confirm the requirement is met, so it does not
  pretend to. Note this is a phrasebook *declaring* what it needs; it is
  not yet a restriction on what the verbs themselves may do — see the
  `SEC.7` item below, which is unchanged by this release.

  Because a version requirement is only as good as the record it is
  written against, the release checks now also refuse to let a form
  reach a release with no row in
  [`docs/GRAMMAR_SINCE.md`](GRAMMAR_SINCE.md). If you are writing a
  phrasebook, that means the version number you look up there covers
  every sentence Frazaro understands, not just the ones someone
  remembered to record.

### Known open security items

- **SEC.3** — generated code does not yet carry phrasebook provenance.
- **SEC.7** — a small set of verbs with real external effect (`vlasendmail`
  today) still runs with no permission check: a phrasebook you load can
  send mail with no consent prompt. `SEC.8` above narrows this but does
  not close it — it gates on where the *workbook* came from, not on what
  a phrasebook asked permission to do, so a phrasebook loaded into a
  workbook of your own still reaches these verbs unprompted.

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
