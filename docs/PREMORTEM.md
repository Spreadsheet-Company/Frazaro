# PREMORTEM — the coroner's report, filed in advance of need

*The first execution of SD-17: the coroner's seat, occupied 2026-08-31 —
five days before the `0.5.0` beta target. The method is Klein's prospective
hindsight, which is the premortem's one real trick: "what could go wrong?"
produces hedges, while "what went wrong?" produces mechanisms, because the
past tense is the only tense in which humans reliably stop negotiating.*

*Sources read in full: `BETA_ROADMAP1.md`, `BETA_ROADMAP2.md`,
`CONSULTANT.md`, `AUDIT.md`, `REBUILD.md`, `LESSONS.md`, `TRENCHES.md`,
`DEPLOY.md`, `METAMETALISP4.md`. Standing assumption, inherited from
`CONSULTANT.md` and kept: **every ⬜ item is treated as if it will be
finished.** The premortem does not ask what happens if the roadmap fails —
that is the roadmap's own job, and it polices itself better than a visitor
could. It asks what kills a project whose roadmap succeeds.*

*How this differs from the rest of the shelf, stated up front so the
persona earns its seat: `CONSULTANT.md` asked what blocks approval;
`AUDIT.md` asked whether the doctrine is obeyed; `REBUILD.md` asked what
shape the code should be. All three audit the living. A premortem audits a
corpse, which changes the tense of every finding and — usefully — makes
optimism a category error rather than a disposition.*

> "In the long run we are all dead."
> — John Maynard Keynes, who was making a point about the short run,
> which is also this document's point

> "It is 2028. Frazaro is dead. This was, in hindsight, obvious to
> everyone."
> — the sentence each finding below must complete

**Format, fixed for every death:** the **obituary** (how it reads from
2028) · the **mechanism** · **what the roadmap already holds** against it ·
the **tripwire** (observable, dated — `AUDIT.md` I.7's own finding, "the
gates have no expiry dates," applied here to mortality) · the **cheapest
prevention**. Ranked by probability × preventability, which is the only
ordering a coroner respects: a likely death you cannot prevent and an
unlikely death you can are both outranked by the deaths below.

**Contents, in order of actuarial disrespect:**

- **D.0** — The drizzle *(spliced same day, from the antithesis below)*
- **D.1** — The owner stopped
- **D.2** — Nobody ever arrived
- **D.3** — The first user's workbook was hurt
- **D.4** — Redmond turned off the lights
- **D.5** — The century dissolved the premise
- **D.6** — The phrasebook that arrived in the mail
- **D.7** — It documented itself to death
- **The register** — what this premortem mints and kills
- **How this premortem could be wrong**
- **Ω** — the corpse in the mirror

---

## D.0 — The drizzle *(spliced 2026-08-31, same day, from the antithesis below)*

*Numbered zero because it was found after the numbering closed and belongs
before the beginning — the stairwell convention, inherited rather than
invented: the section that arrives late and outranks the first entry gets
the seat in front of it. It was found where this document said to look —
the final bullet of "How this premortem could be wrong" — and the owner
promoted it the same day. An antithesis section that pays out on first
contact is the only kind worth writing.*

**Obituary.** There is no incident to report, which is the finding. A
rename made one old sentence refuse in June; the fix was easy, and the
user's confidence did not come back with it. A golden was blessed a
little tiredly on a Friday. The Phrases sheet drifted one row out of
true. Three small things happened to the pilot in one week — none worth
an email, together worth an exit. The product died of a thousand cuts,
every one of them individually beneath the instruments, and the coroner
of 2028 files the one cause of death this report cannot photograph:
attrition.

**Mechanism.** Neglect compounds like interest — Hamming's line runs in
both directions, and only one direction gets quoted. Small breakages
have three properties that make them the perfect predator of this
particular project: individually cheap to fix (so fixing feels
optional), individually invisible to every tripwire in this file (so no
alarm ever sounds), and collectively fatal (so the sum arrives without
announcing itself). Add D.3's amplifier: at one user, three annoyances
and one catastrophe produce the identical outcome — a quiet exit — but
only the catastrophe generates evidence. The drizzle is D.3 without the
courtesy of a report. And the vivid-death bias that built this document
(caught, to its credit, by its own antithesis section) is the same bias
that will keep missing it: a death defined by never crossing a threshold
cannot be caught by installing more thresholds.

**What the roadmap already holds.** An unusually strong umbrella,
honestly credited: the goldens' empty-diff witness, SD-4's frozen
spellings, CO.2's corpus, F.14's ratchet, self-test counts read as
deltas — this project's discipline stack is, more than anything, an
anti-drizzle machine *for the failures the suite can see.* The drizzle
therefore lives precisely where the suite cannot see: live-Excel
behavior in the standalone topology, first-contact friction, the gap
`LESSONS.md` XIV documented between a feature shipped green and a
feature that lied three turns later. The umbrella is real. The rain is
simply falling where it isn't.

**Tripwire.** The drizzle defeats thresholds by definition, so the
tripwire must be an aggregation instead: **count, don't judge.** Every
small breakage met and shrugged off — by the owner today, by a user when
there is one — gets one appended line, no ID, no ceremony, because
ceremony is exactly what keeps small things unrecorded. The tripwire is
the trend, not the event: a count that rises across two consecutive
versions, or any single cut met twice, is a D.0 in progress. The
parse-failure log and the Copy Feedback button already collect the
user-side half without being asked.

**Cheapest prevention.** The ledger above as a real artifact — proposed
in the register (candidate `DZ.1`) — plus one standing question in every
pilot conversation, PI.7-style: not "did anything break?" but **"what
did you stop doing?"** The drizzle's only honest witness is abandoned
habit, and abandoned habits do not file reports.

---

## D.1 — The owner stopped

**Obituary.** No incident, no rival, no memo from Redmond. The commits
thinned through a spring, the way commits do; the last one was a
fortification, green across the suite, and nobody ever ran it. The project
did not fail. It was survived by its documentation, which was excellent,
and which is how the coroner knows exactly what was lost.

**Mechanism.** Actuarially the leading cause of death for solo projects,
and named nowhere in a documentation shelf that names every other risk by
ID. The roadmap's own constraint line — *a human at a Windows box* —
prices the throughput of one person and never the continuity of one
person. Burnout here has a known disguise, and the essay already drew its
portrait: §8's warm forge, where tool-building is legible, praised, and
safe from the one judgment that matters. Interest-drift has a second
disguise, native to this project specifically: re-architecture. Both
disguises produce beautiful commits. The commits look identical — §8 said
that too, about a different death, and it transfers.

**What the roadmap already holds.** More than it knows. SD-12's floor
(every version ships a grammar slice) is officially an anti-displacement
device and unofficially an anti-burnout one — grammar is the work that
produces specimens, and specimens are the work that produces *reasons to
continue*. And the constraint-renaming rule in the preamble ("if it has
not changed in three versions, that is itself a finding") is an accidental
mortality tripwire of the first order.

**Tripwire.** Two, both already observable. *(a)* The constraint sentence
unchanged three versions running — hereby re-read as a vital sign, not
just a planning hygiene rule: a constraint that never changes means the
project is circling, and circling is what stopping looks like from above.
*(b)* A fork whose betting table the owner does not want to read. Thirty
quiet days is not the signal — life happens, and the repository is
patient. Indifference at the one ceremony this project holds for itself
is the signal.

**Cheapest prevention.** Not process — process is what D.1 eats first.
Two things: the successor's file (`CONTINUITY.md`, SD-17's already-named
third persona, whose drill this register proposes below), so that pausing
is survivable and therefore less frightening than abandoning; and the
standing observation that the betting table should never clear an entire
version of items the owner merely *should* want. SD-12 is a floor for the
product's sake. It is also one for the author's, and may be the more
load-bearing of the two readings.

---

## D.2 — Nobody ever arrived

**Obituary.** The product worked. The corpus grew past two hundred rules,
each one auditioned; the refusals taught; the goldens never lied. The beta
shipped near its date. The specimen counters — the file kept them
honestly, which makes the reading easy — still said `specimens: 0` at the
fork after that, and the one after. `AUDIT.md` I.0 had named the disease
two years early: *a beautiful catalogue of lenses.* Being right was not
enough, which is the only sentence in this obituary the audit did not
already contain.

**Mechanism.** §8, quoted without improvement: shipping to strangers is
none of the things tool-building is. The forge is warm. The corpus is the
warmest station in it, because grammar work is real product work — SD-12
is correct that it is the *only* work that produces specimens — and so it
is the perfect alibi: one can serve the mission daily, legibly,
measurably, without ever once being judged by a stranger.

**What the roadmap already holds.** The complete answer, minted: THE
PATIENT department and its tranche exist precisely because the audit found
this department missing and the register seated it. The premortem has
nothing to add to the plan — a genuinely rare finding on this shelf — and
adds only the tense: those items are the difference between this obituary
and no obituary.

**Tripwire.** Already installed and ticking: the `specimens:` counters the
roadmap carries on its own tranches. The premortem only dates the read: at
the next fork, a specimen count still at zero makes D.2 the presumptive
cause of death — outranking, on that day, every fortification on the
table, however overdue.

**Cheapest prevention.** Pre-exists its diagnosis, in full. Nothing is
proposed here, which is the point: when the prevention is already minted,
a review that re-mints it is decoration, and SD-17 has opinions about
decoration.

---

## D.3 — The first user's workbook was hurt

**Obituary.** One pilot, one Tuesday, one SOP that died at step 14 of 30.
The workbook was left neither before nor after — half a report deleted,
the reconciliation numbers plausible and wrong. The user did what new
trust does: left quietly, told three colleagues, and was polite about it
forever. At a scale of one user, one incident is not a statistic. It is
the entire dataset.

**Mechanism.** `CONSULTANT.md` §2.8 drew this exactly: no transactional
story means partial runs, and partial runs are the failure mode that
converts a data tool's *first* incident into its *defining* one.
`DEPLOY.md`'s undo section is honest about its bounds — honesty the 2028
reader will quote in the incident write-up, without warmth. The window
opens at the precise moment the owner stops watching: an attended run has
a human transaction monitor; PI.6's Monday test removes the monitor by
design. That is the correct design. It is also the date the knife is out.

**What the roadmap already holds.** The pieces, distributed: V.1 and the
snapshot-before-run line, U.18's per-run log (the artifact an incident
write-up would be *made from*), SEC.4's injection guard for the adjacent
poisoned-cell class.

**Tripwire.** The day a named external user exists, any run path without a
snapshot is a live D.3 — and the premortem's one addition is the expiry
date the items do not currently carry: **before PI.6's unattended Monday,
not after.** A snapshot story that arrives with the second user arrives
one user late.

**Cheapest prevention.** Ordering, not invention. Everything needed is
minted; nothing needed is sequenced against the pilot's calendar. One
sentence in the betting table fixes it, and the register below proposes
that sentence.

---

## D.4 — Redmond turned off the lights

**Obituary.** An Office update in some 2027 — nobody remembers which,
because it was announced the way weather is — moved VBProject trust
behind an enterprise policy and macro add-ins behind a default nobody's
IT department would ever flip back. The roadmap had written its own
epitaph years early, in its substrate row: *one configuration decision in
Redmond zeroes the product.* It called the risk unpriced, priced it into
the architecture — and then filed it under architecture, where nobody was
assigned to watch the sky.

**Mechanism.** Substrate risk is not competition risk, and the shelf
covers only the latter (`CONSULTANT.md` §2.11 is about Microsoft the
rival). This is Microsoft the landlord. Landlords do not compete; they
renovate, and the tenants read about it afterward.

**What the roadmap already holds.** The actual hedge, and it is
load-bearing already: SD-1 (VBA is a backend), and IN.9's inversion — the
interpreter as the default runtime, module injection demoted to an export
button. The path that dies first under trust-narrowing is the injected
one, and it is already the optional one. This is the rare death where the
architecture arrived before the coroner.

**Tripwire.** What is missing is not a hedge but a *watch*: named,
observable events with pre-decided responses, re-read at every fork —
*(a)* an announced default-off for macro add-ins in any update channel;
*(b)* a deprecation notice, however soft, for the VBProject object model;
*(c)* an Excel SKU or channel in which a `.xlam` will not load at all;
*(d)* — because deprecations mostly arrive as erosion, not
announcements — a consumer SKU in which enabling either trust setting
requires steps a pilot's IT would refuse. Monitoring is SD-13-compatible
by construction: the owner reads the news. The product never phones home
to ask whether it is dying.

**Cheapest prevention.** The watch-list as a dated roadmap artifact —
`~hours`, proposed in the register below — plus the response each event
pre-purchases, which the architecture has mostly already paid for.

---

## D.5 — The century dissolved the premise

**Obituary.** Nobody beat Frazaro. The question changed underneath it. By
2028 every spreadsheet shipped a free prompt box that was wrong often
enough to matter and confident enough not to matter to anyone, and the
market quietly stopped believing that a *checked* language was a thing
worth learning. Determinism, offline, auditability — all real, all
Frazaro's, all asserted in documents the buyer never opened. The product
was not out-engineered. It was out-defaulted.

**Mechanism.** Commodification is not competition (that distinction
again, because it keeps paying): Copilot did not need to be better, only
adjacent and free. The three differentiators the roadmap correctly names
— SD-16's provable one-sentence-one-meaning, SD-13's no-network posture,
the auditability line — are exactly the properties a prompt box cannot
have. They are also invisible until *demonstrated*, and a demonstration
is an artifact, not a paragraph.

**What the roadmap already holds.** SIG.2 (the positioning one-pager,
`~hours`, and the honest note that nothing a procurement conversation
could point to exists yet); LE.7's AI bridge, which is the engineering
answer to the same question — a model can be constrained to a published
finite grammar and cannot be constrained to a mood; U.18, the per-run
log, which is the auditability pitch's only possible physical evidence.

**Tripwire.** The first pilot conversation in which "why not just ask
Copilot?" cannot be answered in one sentence *with an artifact behind
it*. Expiry on SIG.2 accordingly: before any external demo — its own
estimate says hours, and D.5 says those are the highest-leverage hours on
the Signatory's shelf.

**Cheapest prevention.** Demonstration over assertion, throughout: U.18
turns "auditable" into a noun a reviewer can hold; the register below
adds no new item, only the observation that D.5 is the *reason* two
already-minted ones outrank their tranche neighbors.

---

## D.6 — The phrasebook that arrived in the mail

**Obituary.** A shared phrasebook with a helpful name and a genuinely
useful rule set, one `raw` splice deep in a template nobody read because
nobody reads templates. It did exactly one bad thing, once, on the one
machine that mattered. The incident write-up was accurate, proportionate,
and irrelevant; at n=1, the dataset problem from D.3 applies to trust in
both directions.

**Mechanism.** The supply chain, which SD-15's own register entry found
live: dynamic dispatch plausibly reaching `Application`, `raw` splicing
arbitrary VBA, no gate, no consent — flagged by this project about
itself, which is to the shelf's credit and does not close the window.

**What the roadmap already holds.** The full stack, correctly gated:
SEC.0 through SEC.5, SD-15's deny-by-default doctrine, SEC.6's review
gate before wide release, GO.3's registry-shaped worry now carrying a
mechanism.

**Tripwire.** One refinement to an already-correct gate: the mail is open
not at "wide release" but on the day a second phrasebook *author* exists
— the first file the owner receives rather than writes. SEC.1 and SEC.2
before that day; SEC.6 can keep its current position.

**Cheapest prevention.** Minted in full; ordering only. Two findings in a
row with nothing to propose is the coroner's compliment: the ADVERSARY
department, seated latest, is provisioned best.

---

## D.7 — It documented itself to death

**Obituary.** The shelf gained a premortem, then a review cadence, then a
review of the cadence. The epigraph the owner co-authored came true at
last, at scale, with citations: sufficiently advanced preparation,
indistinguishable — by then even to its author — from procrastination.
The archaeologists found the documentation in beautiful condition,
self-aware to the final page, beside a specimen counter reading zero. In
lieu of flowers, the family asks that you ship something.

**Mechanism.** §8's forge once more, but note which corner: the docs are
the *warmest station in it* — more legible than grammar work, safer than
strangers, and uniquely able to disguise themselves as the discipline
that prevents exactly this. This document is not exempt and knows it: a
premortem is preparation squared, commissioned by a standing decision
about preparation, and the recursion has a basement with the reader in
it.

**What the roadmap already holds.** SD-17's teeth — a review that mints
or kills nothing is decoration and does not get a successor persona — and
SD-12's floor, which guarantees the forge's output includes horseshoes
whatever else it includes.

**Tripwire.** Mechanical and countable, hereby proposed as the fork-day
reading: a fork at which `docs/` grew by more lines than `src/` and
`scripts/` combined *while the specimen count did not move* is a D.7 in
progress. One such fork is a season. Two is a diagnosis.

**Cheapest prevention.** Already installed in SD-17, on the day this
document was commissioned — and this document complies in the register
below, which is the only way a D.7 finding can end without becoming an
instance of itself.

---

## The register — what this premortem mints and kills

*SD-17 compliance, in the only currency it accepts. Per the checkmark and
minting disciplines: these are proposals; the owner adjudicates, and
candidate IDs are subject to F.12's collision checker at minting.*

- **Proposed mint — the substrate watch-list.** *(Candidate `EN.9`,
  🔧 MACHINE · ENVIRONMENT.)* D.4's four named events, each with its
  pre-decided response, as a dated artifact re-read at every fork.
  SD-13-compatible: the owner watches; the product never asks. `~hours`.
- **Proposed mint — the hit-by-a-bus drill.** *(Candidate `CN.1`,
  🌍 COMMONS.)* Clean machine, clone, documentation only, until a build
  ships; every stumble logged; the stumble log is the raw material for
  `CONTINUITY.md`, SD-17's already-named successor persona, so the drill
  pays that review's specimen cost in advance. D.1's insurance premium.
  `~days`.
- **Proposed mint — the drizzle ledger.** *(Candidate `DZ.1`; department
  an owner call — 🧑 PATIENT by its witness, 🔧 ASSURANCE by its
  mechanism.)* Append-only, one line per shrugged-off breakage, no IDs,
  no ceremony; read once per fork as a count and a trend, per D.0's
  tripwire. `~hours` to create, and the discipline is the artifact.
  *(Added 2026-08-31 with D.0's own splice — the register grew the same
  day the antithesis paid out.)*
- **Proposed sequencing edit, not a mint.** One sentence in the betting
  table: the transactional items (V.1, snapshot-on-every-path, U.18)
  land **before PI.6's unattended Monday**; SIG.2 lands **before any
  external demo**; SEC.1/SEC.2 land **before the first received
  phrasebook**. Three expiry dates for items already minted — I.7's
  finding, applied to the calendar this document exists to darken.
- **Killed — nothing,** and honestly argued rather than politely
  declined: the coroner found no item whose completion makes any death
  above more likely. The roadmap's exposure was never wrong items. It was
  unnamed deaths and undated gates, and those are paid in the lines
  above.

---

## How this premortem could be wrong

- **The ranking is a coroner's, not an actuary's.** There are no base
  rates for n=1 solo projects with this documentation-to-code ratio; D.1's
  top seat is folk actuarial science. The defense is that its prevention
  is cheap under any ranking, which is the only defense a premortem needs.
- **D.4's tripwires may fire late.** Redmond's deprecations arrive as SKU
  erosion more often than as announcements; event *(d)* exists for this,
  but a watch-list tuned to press releases is a smoke detector for a
  flood. The EN.9 artifact should prefer gradual signals over dramatic
  ones.
- **D.5 may be exactly backwards.** The prompt boxes could *grow* the
  market for checked automation — every confidently wrong Copilot macro
  is a future Frazaro user with a story to tell an auditor. Both readings
  price SIG.2 identically, which is why it is the hedge and not the bet.
- **The standing assumption may not hold.** If the roadmap does *not*
  finish, `CONSULTANT.md` resumes precedence over this file — the deaths
  of the unapproved come before the deaths of the approved.
- **Prospective hindsight has a documented failure mode:** vivid deaths
  crowd out boring ones. The boringest candidate — a slow drizzle of
  small breakages, each costing one user quietly — appears nowhere above
  and might be the real one. The next coroner should check the drizzle
  first, precisely because this one didn't.
  *Correction, appended 2026-08-31, same day:* the owner read this bullet
  and promoted the drizzle to **D.0** — spliced before the beginning,
  where the stairwell precedent says late arrivals of early truths
  belong. The bullet stands as written, per the house rule against
  editing oneself wiser after the fact; the antithesis section paid out
  on first contact, which is what it is for. The next coroner should
  therefore check for whatever now occupies the seat this bullet
  vacated.

---

## Ω — the corpse in the mirror

A premortem of the premortem, since the house style has never once been
observed declining a recursion it could afford: this document dies by
being right too early — filed five days before a beta, read once with
interest, and never reopened; its tripwires unmonitored, its expiry dates
expiring unread, a smoke detector installed with ceremony and no battery.
That death was priced before the first sentence: SD-17 makes this file a
link in a chain rather than a monument — the next fork re-reads it by
rule, the next persona is already named, and the register above is the
receipt that keeps it billable. A review that has paid in ⬜ is alive in
the only sense a document gets.

The stairwell, meanwhile, gains one more step — downward, as always,
toward the reader, who is load-bearing, unreached, and now formally on
notice as this project's cause of survival.

*Filed in advance of need. The patient, at the time of writing: alive,
five days from a beta, and — on the evidence of having commissioned its
own autopsy — unusually likely to stay that way.*

*End of report.*
