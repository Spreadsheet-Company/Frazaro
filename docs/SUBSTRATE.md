# SUBSTRATE — the platform historian's survey, or: the landlord's record, read

*The fifth execution of SD-17, and the last seat of the original ledger:
the platform historian, occupied 2026-08-31 — same sitting as the
coroner, the advocate, the successor, and the accountant, who between
them have already confessed the cadence four different ways; the
historian merely notes that the sitting is now itself the shelf's most
flagrant D.7 exhibit, and that history will judge it by one criterion
only: whether the fork collects.*

*The commission, inherited twice: `PREMORTEM.md` D.4 named the substrate
death and proposed the watch-list (`EN.9`) without contents;
`ADVOCATUS.md` filed the premise no seat before this one was positioned
to doubt — "Excel is where the users are" — and `VIABILITY.md` co-signed
it as the largest unbookable line item on its books. Both land here.*

*Sources read in full: `BETA_ROADMAP1.md`, `BETA_ROADMAP2.md`,
`DEPLOY.md`, `PREMORTEM.md`, `ADVOCATUS.md`, `CONTINUITY.md`,
`VIABILITY.md`, `CONTEMPLATIONS.md` — plus one census of the codebase
itself (the `Declare` sites, counted below), and the public record of
the platform, which is the historian's actual archive: every exhibit
cited is a dated, checkable event in Office's own history. Findings are
numbered `H.*`, advisory per `ID_REGISTRY.md`'s distinction, never
governed.*

*Method, stated because it differs from every seat before it: the other
four reasoned from the project outward. The historian reasons from the
record inward — substrates die in characteristic ways, this landlord's
record is thirty-three years long, and base rates beat vivid scenarios
everywhere base rates exist. The premortem imagined the lights going
out; this survey reads the landlord's actual history of light
switches.*

> "History doesn't repeat itself, but it often rhymes."
> — Mark Twain, who never said it, which for a document about
> misattributed obsolescence is a credential

> Exhibit zero: in 2008 the landlord shipped Mac Office **without VBA**
> — removal, actually tried, once. In 2011 it was back. The tenants
> won in one release cycle.

**Contents:**

- **The record** — the dated exhibits, 1993–2024
- **H.1** — The rungs, and which ones Frazaro stands on
- **H.2** — The premise, cross-examined: "Excel is where the users
  are"
- **H.3** — The hedges already built, named as hedges
- **H.4** — The census, and the kill hunt the accountant ordered
- **The watch-list** — `EN.9`, hydrated
- **The register** — what this survey mints, rescopes, and kills
- **How this survey could be wrong**
- **Ω** — the method's own landlord

---

## The record

*The exhibits, in order. Each is public and checkable; the readings are
the historian's.*

- **1993** — VBA arrives (Excel 5). **~1998–2000** — the VBE reaches
  essentially its final form and never materially changes again: a
  quarter-century of maintenance without investment. *Reading:*
  embalmment. A landlord renovates the floors he plans to rent.
- **~2002–2008** — VSTO and the managed-code succession attempt. Fails
  to displace VBA. *Reading:* succession attempt #1; the deposit was
  already too large.
- **2008 → 2011** — the Mac removal and reversal (exhibit zero,
  above). *Reading:* the only attempted eviction on record, and the
  landlord lost — to the same enterprise macro deposit that secures
  the Windows tenancy many times over.
- **2013 →** — Office JS add-ins: web-first succession attempt #2.
  Coexists; does not displace. *Reading:* the heirs multiply; none is
  anointed.
- **2018** — AMSI integration: Defender gets runtime eyes inside VBA
  macros. *Reading:* the signature of **quarantine**, not removal —
  you don't build surveillance into a building you're demolishing.
- **2021** — LAMBDA ships and Microsoft's "world's most popular
  programming language" marketing attaches to *formulas*. *Reading:*
  evangelism has a new address. The formula layer is where the
  landlord now invests; VBA is where he polices.
- **2021–22** — **XLM macros disabled by default**, twenty-nine years
  after their successor shipped. The one completed macro-substrate
  death on the record, and its shape is the playbook: never removed,
  killed by a default flip, admin re-enable retained, announced in a
  blog post. *Reading:* this is what the end looks like, and how long
  the fuse burns.
- **2022** — VBA blocked by default for Mark-of-the-Web files.
  *Reading:* the XLM playbook's first act applied to VBA — but
  provenance-scoped, at the edges, not blanket. Quarantine tightening
  on schedule.
- **2019–2024** — Office Scripts (web-first, desktop late and
  partial); **Python in Excel** (2023 — executing *in Microsoft's
  cloud*, an architecture that tells you the successor languages are
  not expected to run on your machine); Copilot. *Reading:*
  succession attempts #3–5, still none anointed.

**The base-rate verdict.** In thirty-three years: one attempted
removal (lost), one completed default-death (XLM, on a 29-year fuse,
with its successor installed the whole time), and otherwise a single
consistent practice — **quarantine**: trust gates, provenance checks,
default-narrowing at the edges, tooling embalmment, evangelism
relocation. Substrates here die by default-setting, not deletion, and
the fuses are measured in decades. `PREMORTEM.md` D.4's obituary
("announced the way weather is") was written before this record was
read, and the record confirms its mechanism while stretching its
timeline: the weather is real, and it is glacial.

---

## H.1 — The rungs, and which ones Frazaro stands on

**The ladder the record implies**, from sturdy to condemned: running
installed macro code (the deposit itself) · loading new macro
artifacts (provenance-gated, 2022) · **programmatic access to the VBA
project** (the VBOM trust setting — off by default as long as anyone
can remember, AV-heuristic-watched, the single most quarantined corner
of the platform) · VBA on new platforms (web Excel: was never granted
a tenancy at all).

**Where Frazaro stands, rung by rung.** The *product*, since IN.9,
stands on the sturdiest rung: the default runtime interprets, touches
no VBProject, and needs no trust setting beyond running the add-in at
all. The *export button* stands one rung down, gated and honest about
it. But the survey's one new finding lives here: **the factory stands
on the condemned rung.** `VlaBuildAddin` constructs the shipped
`.xlam` through the VBProject object model; `VlaDevReload` and the
whole dev rig live there too. The shelf has carefully moved the
product off VBOM and never once audited that the *means of production*
still stands entirely on it — quarantine tightening could break the
factory while every shipped copy keeps running perfectly, which is a
failure mode none of the four previous seats priced: not death, but
**sterility**. Proposed as `EN.10` below.

---

## H.2 — The premise, cross-examined: "Excel is where the users are"

**As filed:** the advocate could not argue it because both sides of
that document stand on it; the accountant could not book it. The
historian can at least *decompose* it, which is what the record is
for.

**The premise Frazaro actually needs is narrower by three qualifiers:**
*Windows desktop* Excel *with macros permitted* is where the users
are. Rate each leak separately:

- **"Excel"** — the sturdiest install base in commercial software,
  sitting atop a forty-five-year spreadsheet lineage, with its
  formula layer *growing* (LAMBDA, dynamic arrays — the layer F.7's
  dialect targets). This clause is not the leak.
- **"Windows desktop"** — leaking, slowly: web Excel's share rises,
  and the generational pipeline (students on Sheets and web apps)
  prices the far horizon. But note *which* users: Frazaro's
  demographic — people whose procedures live in workbooks, the SOP
  and back-office cohort — is the desktop-stickiest population Excel
  has. The premise decays slowest exactly where the product aims.
  (This is also the survey's least-evidenced sentence; see the
  how-wrong section.)
- **"With macros permitted"** — the active leak, per the entire
  record above: this clause is what quarantine erodes, edge by edge,
  default by default. It is the clause the watch-list watches.

**Verdict.** The premise holds for the pilot's decade, weakens
clause-by-clause rather than collapsing, and the correct instrument is
not a hedge-in-advance but the watch: each fork re-reads the three
clauses against the exhibits. For `VIABILITY.md`'s books: the Register
destiny's long horizon is the one priced against the pipeline clause;
the near destinies are priced against the macro clause only.

---

## H.3 — The hedges already built, named as hedges

*Credited precisely, because a hedge that isn't named as one gets
refactored away by someone who thinks it's an accident:*

- **SD-1** ("VBA is a backend, not the language") — the constitutional
  hedge; everything below is its case law.
- **The head-table fork (F.1/F.2)** — the option premium already
  paid: a second backend is a peer, not a port, *while* the dispatch
  is one table. This is the mechanism that makes every response in
  the watch-list affordable.
- **IN.9** — the biggest shipped hedge on the shelf: the product off
  the condemned rung, module injection demoted to a button, Mac
  converted from apology to missing button (EN.6). Built for product
  reasons; worth double as substrate insurance.
- **F.7, the formula dialect** — the quiet one: a backend aimed at
  the one layer the landlord is *investing* in. If every macro gate
  slams, the formula dialect is the backend that still ships — even
  on web Excel, where VBA never had a tenancy at all.
- **SD-16's wall** — the hedge against the other weather system
  (D.5): whatever the prompt boxes commodify, a provable
  one-sentence-one-meaning language is the artifact they cannot be.

**And the hedge deliberately not proposed:** a JS/Office-Scripts
backend, on spec, today, would be an ark built from a forecast —
SD-7's spirit applies to backends exactly as to grammar sections (no
backend without a user who needs it), and the head table holds the
door open at no carrying cost. The correct spend is the watch, not
the ark. The historian, of all people, declines to cosplay Noah.

---

## H.4 — The census, and the kill hunt the accountant ordered

**The standing instruction, honored:** `VIABILITY.md`'s register
warned that if a fifth review also kills nothing, the personas may
have gotten comfortable. The historian therefore went hunting with a
measurement, not an opinion — and reports the hunt whichever way the
evidence points.

**The census.** `Declare` statements in the codebase: **one site.**
`frmCLI.frm`, four Win32 API functions, already correctly
dual-branched under `#If VBA7` with `PtrSafe`/`LongPtr` on the modern
side and the legacy signatures preserved. That is the *entire*
bitness-sensitive surface of the project.

**What the measurement rescopes — two items, neither a clean kill:**

- **EN.8** ("64-bit declaration discipline as a lint rule,
  `~hours`") — the hazard set it guards is one already-compliant
  site. The honest item is smaller than the minted one: not a
  discipline-lint but a **one-assertion ratchet** in F.14's exact
  shape — fail if any `Declare` appears outside the blessed site, or
  any new one arrives without the dual branch. Rescope, don't kill:
  the ratchet is cheaper than the lint and stricter.
- **EN.7's bitness axis** — the environment matrix multiplies
  versions × locales × bitness × backend, and the census just
  collapsed one axis: the project's bitness sensitivity is "does
  `frmCLI` load," a single cell, not a dimension. Kill the axis;
  keep the matrix.

**The hunt's verdict, filed for the accountant:** asked, per your
instruction. The evidence pointed at reduction, not deletion — a
measured shrink of two items rather than a ⛔. If that is comfort, it
is at least comfort with a grep attached.

---

## The watch-list — `EN.9`, hydrated

*`PREMORTEM.md` D.4 minted the candidate; this section is its
contents, so that minting it costs filing, not writing. Gradual
signals ranked above dramatic ones — the record shows deprecations
arrive as erosion, and `PREMORTEM.md`'s own how-wrong section demanded
a smoke detector that works on floods. Cadence per the original
proposal: re-read at every fork; the owner reads the news;
SD-13-compatible by construction — the product never phones home to
ask if it is dying.*

| # | Signal (observable, dated when seen) | Pre-decided response |
|---|---|---|
| 1 | VBOM trust setting removed or policy-locked in consumer SKUs | Factory risk, not product risk (H.1): execute `EN.10`'s recovery page; keep one blessed builder environment; product and users unaffected |
| 2 | Provenance gates (MOTW-style) extended to local or intranet files | Already doctrine: the standalone-add-in path plus signing (`DEPLOY.md`); verify the trust-publisher flow still clears it |
| 3 | **The XLM playbook announced for VBA** — "disabled by default, admin re-enable" | The big one; the fuse is years by precedent. Accelerate target-neutrality Part B; the interpreter keeps working wherever re-enabled; F.7 carries the read-only half everywhere else |
| 4 | An Office channel or SKU where `.xlam` add-ins do not load | EN.1's capability probe refuses in words (already scoped); begin measuring that SKU's share of the pilot demographic |
| 5 | Defender/AMSI heuristics flag the built artifact or `VlaBuildAddin`'s behavior (code writing code is the exact heuristic class) | Signing posture plus SIG.1's transparency page; add a stock-Defender machine to the release clickthrough |
| 6 | Evangelism signals: VBA reference archived, absent from new-feature announcements two years running, VBE finally touched (in either direction) | No action; recalibrate every fuse estimate above and re-read H.2's clauses |

---

## The register — what this survey mints, rescopes, and kills

*SD-17 compliance. Proposals; the owner adjudicates; candidate IDs
subject to F.12's checker at minting.*

- **Seconded, and hydrated — `EN.9`, the substrate watch-list.** The
  coroner's candidate, still unminted; its contents now exist above,
  so the mint is a filing act. The registry's own table already names
  `EN.9` the family's next free slot.
- **Proposed mint — the factory's trust dependencies, enumerated.**
  *(Candidate `EN.10`, 🔧 MACHINE · ENVIRONMENT.)* H.1's new finding:
  which build and dev steps require VBOM or macro trust
  (`VlaBuildAddin`, `VlaDevReload`, the signing clickthrough), and
  the one-page recovery if a Windows update flips any of them — the
  successor audit's C.3 cousin, for the machine instead of the human.
  The product is off the condemned rung; the factory is not, and
  nobody had written that down. `~hours`.
- **Proposed rescope — `EN.8`** from discipline-lint to one-assertion
  ratchet (F.14's shape), per the census. Cheaper and stricter than
  what it replaces.
- **Proposed kill — `EN.7`'s bitness axis.** The matrix keeps
  versions × locales × backend; bitness collapses to one cell
  (`frmCLI` loads) by measurement. The first kill any of the five
  reviews has drawn blood on, and it is a dimension, not an item —
  which the historian reports as exactly what a well-policed roadmap
  should surrender first.

---

## How this survey could be wrong

- **Base rates assume the landlord rhymes with himself.** The record
  has no AI chapter: Copilot-era pressure to close macro attack
  surface (or to clear floor space for the anointed successor it
  finally has) could compress every fuse the XLM precedent suggests.
  The watch-list's signal #3 is the check; the fuse estimate is the
  guess.
- **Exhibit zero is aging.** The tenants who won the Mac reversal in
  2011 were enterprises with macro deposits; those deposits are aging
  with their authors. The counter-reading: workbooks outlive authors
  (the COBOL lesson, which is a lesson about *both* longevity and
  what longevity costs a platform's reputation) — but the historian
  notes the strongest exhibit is also the oldest.
- **The desktop-stickiest-cohort claim is folk demography.** H.2's
  most load-bearing sentence has no citation, and the survey knows
  it. PI.7's trust reading is its only near-term check, and the
  clause should be re-read there, not here.
- **Half-successors may shorten fuses, not lengthen them.** The
  survey read "no anointed heir" as "no eviction date," per XLM
  (which burned its fuse only after VBA stood ready). A security
  incident could invert that overnight — policy deaths, unlike
  platform deaths, need no successor at all. Signal #5 exists for
  this.
- **Five reviews, one sitting.** The record will show this day as
  either the fork's best-collected receivable or the shelf's
  monument to D.7, and the historian — uniquely among the five —
  will not get to write that verdict. The fork will.

---

## Ω — the method's own landlord

The historian's final exhibit is the room he is standing in. This
survey audits one landlord's power over the product while being
written under another's: the collaboration protocol runs on a
specific AI vendor's models — the method has a substrate too, with
its own trust gates, its own pricing, its own deprecations and
retirements, and the seat drafting this sentence is itself the
tenancy. Every argument above about quarantine, embalmment, and
anointed successors applies, clause for clause, to the other half of
the dyad — and no document on the shelf has ever priced it, because
the seat that would write that survey is the one with the conflict.

The mitigation, it turns out, was built years ago for other reasons,
and the successor's audit already proved it works: everything is
written down, in model-neutral prose, on the shelf — so the method
transfers across assistants the way the product transfers across
backends. SD-1 has a twin nobody minted, and the historian leaves it
on the table in the register's own voice: **the assistant is a
backend, not the method.**

With that, the original persona ledger closes: coroner, advocate,
successor, accountant, historian — all seated, all filed, all in one
sitting that the next fork will judge. SD-17's non-repetition rule
now demands what it was always going to demand eventually:
**invention.** The cheap seats are taken. The register's own text
said the list was not closed, and the next reviewer must be a persona
nobody at this table has thought of — which is, the historian notes
in closing, exactly how every seat at this table got here.

*Filed fifth and last of the sitting. The reading room's lights are
off; the fork knows where the switch is.*

*End of survey.*
