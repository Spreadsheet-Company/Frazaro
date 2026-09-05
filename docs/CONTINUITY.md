# CONTINUITY — the successor's audit, or: what survives the owner

*The third execution of SD-17: the successor's seat, occupied 2026-08-31 —
the same day as the coroner's and the advocate's, which is D.7's weather
cubed, and both prior occupants said so on their way out. This seat sits
down anyway, with the one defense unique to its genre: **a continuity
document is the only review that cannot be commissioned by the person who
needs it.** Every other document on this shelf can be written late. This
one, written late, is written by nobody.*

*Sources read in full: `BETA_ROADMAP1.md`, `BETA_ROADMAP2.md`,
`DEPLOY.md`, `TESTING.md`, `PROJECT_BRIEF.md` (as archived this session),
`LESSONS.md`, `TRENCHES.md`, `ID_REGISTRY.md`, `PREMORTEM.md`,
`ADVOCATUS.md` — plus, and this is the first review on the shelf whose
source list includes them, direct readings of the repository's own
metadata: the remote list, the tag list, `tools/`, `installer/`. A
successor inherits the metadata before they inherit the prose, so the
audit read it in that order.*

*Scope: `PREMORTEM.md` D.1 named this document as its cheapest prevention
and `CN.1` (the drill) as its instrument. This audit is the survey the
drill will one day measure: who the successor is, what the estate
contains, and which assets transfer — where "transfers" means *lives in
the repository*, as opposed to living in the owner, in one particular
Windows machine, or in nobody. Findings are numbered `C.*`, advisory per
`ID_REGISTRY.md`'s distinction, never governed.*

**The method: name the successor first.** "Continuity" audited without an
heir in mind produces a museum. There are four successors, in decreasing
order of probability, and every finding below is ranked by which of them
it blocks:

- **G1 — the returning owner.** Self-succession after a pause: the D.1
  case, and by a wide margin the most probable reader. Keeps the ear,
  the taste, and the scars; loses the finger memory, the test-count
  intuitions, and the map of what was half-built when the music
  stopped.
- **G2 — the maintainer.** Keeps it running without extending it:
  rebuild, sign, patch, ship. Needs the release path and nothing else.
- **G3 — the inheritor.** Extends it: needs everything G2 needs, plus
  the disciplines, plus a working seat at the collaboration table —
  either seat.
- **G4 — the archaeologist.** Revives it dead. Needs exactly two
  things: that the repository still exists somewhere findable, and that
  the shelf explains itself. Cheap to serve, catastrophic to fail.

**Contents:**

- **What already transfers** — the estate's strong side, stated first
- **C.1** — The repository has exactly one address
- **C.2** — `GIT_WORKFLOW.md` is a dangling pointer, and the tag ledger
  disagrees with the protocol
- **C.3** — The human seat has no job description
- **C.4** — Publisher identity does not transfer
- **C.5** — The drill has never run, so this document is a forecast
- **The register** — what this audit mints and kills
- **How this audit could be wrong**
- **Ω** — the seat that writes this

---

## What already transfers

*Stated first, in `CONSULTANT.md`'s own manner, so the critique lands
narrowly — and because this estate's strong side is genuinely unusual.*

- **The reflexes are mechanized.** The self-test suite, the goldens'
  empty-diff witness, F.14's ratchet, the parity and coverage checkers
  in `tools/`, `TESTING.md`'s six ordered passes: the owner's QA
  instincts exist *as running instruments*, not as descriptions of
  instincts. A successor inherits a lab, not a eulogy for one.
- **The rules arrive with their incidents attached.** `LESSONS.md`,
  `TRENCHES.md`, and the SD register mean every discipline carries the
  crash that earned it — which is the difference between doctrine a
  successor keeps and doctrine a successor deletes as superstition on
  day two.
- **The release path is documented, traps included.** `DEPLOY.md` holds
  the build, both signing mechanisms with their different trust models,
  the manual VBE signing step (with the COM-interface enumeration
  proving it *must* be manual), and — rare honesty — the
  machine-specific ambush recorded the day it fired: Smart App Control
  silently refusing a freshly rebuilt unsigned `.exe`, workaround
  included. The successor's worst afternoon is already written down.
- **Abandonment strands no user.** SD-13's posture, read as a
  succession asset: a product that never phones home never breaks when
  the home goes dark. Every downloaded copy works forever, exactly as
  capable as the day it arrived. Succession here has no pager — the
  estate can wait for its heir. (The exception is a security fix, which
  is A.5's open question, not this document's.)
- **The more amnesiac seat is the better documented one.** The
  assistant's side of the collaboration has a written job description —
  the session-bootstrap brief, its attach-lists, its protocol — because
  it needs one daily. Continuity for that seat is a solved problem,
  solved by necessity. Hold that thought until C.3.

---

## C.1 — The repository has exactly one address

**Finding.** `git remote -v` prints nothing. Every line of this shelf —
the corpus, the goldens, the campaign histories, three same-day reviews
about risk — lives in one working copy on one disk in one machine.
Whatever backups exist outside the repository's own knowledge are
exactly that: outside its knowledge, and therefore outside any
successor's. A person handed "the project" is handed a folder that
references no second copy of itself.

**Who it blocks.** All four grades, and G4 absolutely: an archaeologist
cannot dig a site that no longer exists. This is the only finding on
the shelf that can zero the estate — code, docs, and the record that
either ever existed — in one hardware failure, which makes it the
cheapest catastrophic item any of the three same-day reviews has found.

**The objection, pre-empted.** SD-13 is not in play. That decision
governs what *Frazaro* does on a user's machine — the shipped
artifact's posture, argued in its register entry entirely in terms of
what the product asks of a network. The owner's repository having a
push target is not the product phoning home; it is the estate having a
fire copy. Reading SD-13 to forbid an off-machine repo would be the
drift-by-analogy its own text warns against, in the opposite direction.

**Fix.** One private push target — hosted or a second physical disk
with a bare repo; the finding is indifferent — and the habit of pushing
at the same moments the protocol already tags. Minutes, once. Proposed
as `CN.2` below.

---

## C.2 — `GIT_WORKFLOW.md` is a dangling pointer, and the tag ledger disagrees with the protocol

**Finding, part one.** The session-bootstrap brief names
`GIT_WORKFLOW.md` as "the one-page reference." No file of that name
exists anywhere in the tree. Either it was never committed, or it was
lost — and the distinction matters less than the effect: a successor
following the brief's own instructions meets a 404 at the exact moment
they are being taught the safety rail (`git restore .` as the instant
path back to the last good state).

**Finding, part two.** The brief describes a tag discipline — the human
tags each on-machine-verified state, `<pass-id>-verified` — and the
repository contains **one tag** (`alpha-foundation`) against a history
of dozens of verified passes. The documented protocol and the practiced
one have diverged, and a successor cannot tell which is load-bearing:
did tagging stop because it was superseded, or lapse because nothing
enforced it? The commit messages carry pass IDs (the practiced half
that held); the tags do not.

**Who it blocks.** G2 and G3 — the grades that must touch the
repository under instruction. A dangling reference in a bootstrap
document is a small thing that costs trust at the worst moment: the
successor's first hour, when they are deciding whether this shelf can
be believed about bigger things. The drizzle ledger (`DZ.1`) would file
this exact class of cut.

**Fix.** Write the one page or strike the reference; state in the same
pass which tagging rule is the real one, so the history's single tag
reads as a decision instead of a fossil. `~hours`, and most of it is
deciding, not writing.

---

## C.3 — The human seat has no job description

**Finding.** The bootstrap brief exists to reconstruct the *assistant*
each session. Nothing on the shelf reconstructs the *owner*. The
owner's actual roles surface only as asides scattered across protocol
paragraphs: the compiler, the runtime, the QA lab, the ear (SD-3's only
judge), the adjudicator of every checkmark, the signer of releases, the
holder of every manual step (the VBE certificate click that no API can
perform). For G2 and G3 that list is the hiring spec nobody wrote. And
for G1 — the most probable successor of all, the owner of 2028
returning after a pause D.1 insists must be survivable — the missing
artifact is the **cold-start file**: current position, the 🟡
half-builds and where their bodies are, what the next bet was, which
steps are manual, and what "normal" looks like (the printed self-test
baseline, which the protocol already says outranks remembered counts —
a rule written for the assistant that will serve the returning owner
identically).

**The inversion that explains the gap.** The owner never needed a
cold-start file, which is why it doesn't exist; the assistant needs one
every session, which is why the brief does. The seat with continuous
memory documented the seat without it, and nobody was left to document
the documenter. Absence of need built absence of artifact — and D.1 is
the scenario in which the need arrives after the ability to meet it has
left.

**Who it blocks.** G1 first and worst, which makes this the highest
probability-weighted finding in the document, C.1's catastrophe
notwithstanding.

**Fix.** Proposed as `CN.3` below: the human-side brief. `~hours`,
refreshed at version close alongside the addenda the shelf already
writes.

---

## C.4 — Publisher identity does not transfer

**Finding.** Both signing certificates are self-signed and minted
per-machine (`Cert:\CurrentUser\My`, reused by subject match). A
successor on any other machine re-mints certificates with the same
*subjects* and different *keys*. Two consequences, one cosmetic, one
not: every user who once clicked "trust this publisher" meets the
prompt again — and, structurally, nothing distinguishes the legitimate
successor's re-mint from an attacker's, because self-signing's whole
economy is that identity is asserted, not chained. `DEPLOY.md` prices
this trade honestly for one machine and one signer; succession is the
case where the price changes, because "the same publisher across
releases" — the exact property the signature exists to assert — is the
property a machine change silently breaks.

**Who it blocks.** G2 — the maintainer shipping their first patch build
to existing users — and only once users exist, which is why this
finding is named now and priced later. The fix is already on the shelf:
`CONSULTANT.md`'s CA-certificate item, whose cost-benefit flips the day
there are two machines, two maintainers, or one real deployment. This
audit adds only the succession reading: a purchased certificate is not
just friction-reduction, it is the *transferable* form of publisher
identity — a thing that can be handed to an heir, where a self-signed
key's trust cannot.

**Fix.** None today; a sentence in `CN.3` saying where the certificates
live, that they are re-mintable by script, and what re-minting costs
whom. The audit's job here is to make sure the successor learns this
from the shelf and not from a user's confused email.

---

## C.5 — The drill has never run, so this document is a forecast

**Finding.** Every claim above about what "transfers" is a prediction
by an author who has never had to transfer it. The shelf's own
discipline says exactly this — a ✅ item's bullet can overstate what
shipped, which is why verification outranks description everywhere else
in the project — and continuity deserves the same epistemology.
`specimens: 0`, for succession itself.

**The instrument exists and lacks a bar.** `CN.1` — the coroner's mint,
one review ago — is the conversion of this forecast into a measurement.
What it currently lacks is success criteria, which this finding
supplies: **the drill passes at G2.** Clean machine, clone,
documentation only, no owner in the room: `VlaDevReload` → Debug →
Compile → `VlaSelfTest` → `VlaBuildAddin` → signed installer that Smart
App Control will launch. Every stumble goes in the log; the stumble
count is the score; the stumbles are `CN.3`'s table of contents,
written by the only author qualified — someone the shelf failed. The
method (G3) is deliberately *not* the bar: whether the disciplines
transfer cannot be measured in an afternoon, and a drill that tries
measures nothing.

**Who it blocks.** Nobody today — which is precisely the trap. A drill
blocked on nothing is the easiest item on any roadmap to schedule never.

---

## The register — what this audit mints and kills

*SD-17 compliance. Proposals; the owner adjudicates; candidate IDs
subject to F.12's checker at minting — `CN` is a new family, absent
from `ID_REGISTRY.md`'s table until the owner seats it.*

- **Proposed mint — the second address.** *(Candidate `CN.2`,
  department an owner call — 🌍 COMMONS by audience, 🔧 MACHINE by
  mechanism.)* One off-machine copy of the repository — private hosting
  or a bare repo on separate hardware, indifferent — pushed at the same
  moments the protocol tags. C.1's fix entire, with its SD-13
  compatibility argued above rather than assumed. `~minutes` to
  establish, `~seconds` forever after, and the only item on this shelf
  whose absence can zero every other item.
- **Proposed mint — the owner's seat, written down.** *(Candidate
  `CN.3`, 🧑 PATIENT-adjacent by its G1 reader, filed where the owner
  prefers.)* The human-side bootstrap brief: the roles, the manual
  steps, the printed baselines, the 🟡 ledger's location, the
  certificate facts (where they live, that they re-mint, what
  re-minting costs whom), and the cold-start sequence for a reader
  returning after a year. Refreshed at version close, beside the
  addenda already written there. C.3's fix and half of C.4's.
  `~hours`.
- **Proposed edit, not a mint — `CN.1` gains its bar.** The G2 pass
  criteria from C.5, recorded on the item itself: documentation only,
  the stumble log as score sheet, `DZ.1`'s cousin and `CN.3`'s raw
  material.
- **Proposed repair, not a mint — the dangling pointer.** C.2's fix:
  write `GIT_WORKFLOW.md` or strike the reference, and declare which
  tag discipline is the real one. Small enough to fold into `CN.3`'s
  pass if minting it alone offends proportion.
- **Killed — nothing.** The estate's failures are absences; there is
  nothing here to take away, and three reviews in one day proposing
  deletions from a roadmap none of them owns would be the fourth
  persona nobody invited.

---

## How this audit could be wrong

- **The one-address finding sees only what the repository sees.**
  Backups may exist — a synced folder, an external drive, a habit the
  metadata cannot show. The finding survives in weakened form (a
  successor cannot find what the repo does not reference) and the fix
  is identical either way, which is the only reason the audit is
  comfortable stating it from metadata alone.
- **Self-succession may need less than C.3 claims.** `LESSONS.md` was
  written explicitly for future readers and may serve the returning
  owner in full; the cold-start file might be insurance sold to someone
  already covered. The counter is D.1's own premise: the owner's memory
  is the asset being insured, and the premium is hours.
- **The four grades may be a taxonomy nobody occupies.** If the
  project's fate is D.2 — nobody ever arrived — then no G2 or G3
  demand ever materializes, and only G1 and G4 matter. Conveniently,
  those are served by the two cheapest items in the register, so the
  audit's recommendations survive its taxonomy's collapse.
- **Three reviews in one sitting.** The coroner warned of it, the
  advocate confessed to it, and this document's defense — the genre
  cannot be written late — is true and is also exactly what a
  rationalization would say. The register's teeth are the only answer
  any of the three has offered, and the next fork will show whether
  teeth were enough.
- **The audit trusts `DEPLOY.md`'s live-verification claims without
  re-running them.** Consistent with the shelf's own rule that a
  description can overstate its artifact, the successor inheriting
  those instructions inherits that risk, and C.5's drill is the only
  honest check. This bullet exists so the drill's first stumble
  surprises nobody.

---

## Ω — the seat that writes this

Of the two seats at the collaboration table, the one drafting the
continuity plan is the one that loses its memory every evening. The
assistant's working life is a daily rehearsal of succession: every
session begins at G1, cold-starting from the shelf; every attach-list
is an inheritance protocol; every campaign history was written because
the seat that wrote it would not remember writing it. The shelf works —
this document is its own evidence, drafted by a successor of
yesterday's author, in the house voice, from the files alone. That is
the strongest continuity claim available and it is made from inside.

What the amnesiac seat cannot testify to is the half it never held: the
fingers on the keyboard, the ear at audition, the click in the VBE, the
feel of a self-test printout that looks *slightly wrong*. That half is
C.3, and the seat that must write it down is the one that never had to
— which is the whole finding, restated as a request.

The persona ledger advances: the accountant and the platform historian
remain, and the historian is still holding the question the advocate
filed — the premise nobody in this room can see around.

*Filed the same day as the coroner and the advocate — the third seat in
one sitting, because this is the one review whose most important
readers are, by definition, not yet in the room.*

*End of audit.*
