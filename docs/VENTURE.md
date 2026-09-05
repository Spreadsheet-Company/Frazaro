# VENTURE — the commercial case for Frazaro, written for outside capital

> **⚠ EDIT BEFORE DISTRIBUTING — three manual passes this document
> requires before it leaves the building:**
>
> 1. **Check the numbers.** Every market figure in §5 is directional
>    and flagged as such; replace with independently verified figures
>    (Excel user base, RPA market size, EUC tooling price points) or
>    delete the claim. A VC team will check; the document must survive
>    it.
> 2. **Decide the ask.** §13 deliberately frames scenarios instead of
>    announcing a raise. Before sending: choose raise / no-raise /
>    partnership, replace the conditional framing with the decision,
>    and name what the capital (if any) buys, in order.
> 3. **Re-price against the pilot.** §5 and §9 pin their hypotheses to
>    PI.2 (the transcript) and PI.7 (the trust reading). If the pilot
>    has run by the time this ships, fold its evidence in — traction
>    with one real user re-prices every section. If it has not,
>    consider whether sending this before the pilot is the right
>    order at all.

*Prepared 2026-08-31, five days before the `0.5.0` beta. Status
disclosures up front, because this project's documentation culture is
part of its pitch: **pre-revenue, pre-pilot, zero external users**, one
developer pair-programming with an AI assistant, working product,
unusually heavy internal verification discipline. Every claim below is
either sourced from this repository, labeled as an estimate, or labeled
as a hypothesis awaiting the pilot. Directional market figures are
marked as such and should be independently checked — this document
would rather survive diligence than dazzle it.*

*Relationship to the rest of `docs/`: this is the investor-facing
distillation of `VIABILITY.md` (the internal economics audit) and the
revenue assessment that followed it. The five adversarial self-reviews
cited throughout (`PREMORTEM.md`, `ADVOCATUS.md`, `CONTINUITY.md`,
`VIABILITY.md`, `SUBSTRATE.md`) are not marketing — they are the
diligence a careful investor would commission, already written, against
the project's own interests. Read them second.*

---

## 1. Executive summary

**Frazaro lets non-programmers automate Excel by writing checked
English sentences.** Sentences are validated against a published,
deterministic grammar *before anything runs*; a sentence that parses
has exactly one meaning; a sentence that doesn't is refused with a
teaching explanation. Programs run interpreted by default (no code
injected, no trust settings), or export to readable VBA on demand. The
same engine ships `=SQL()` and `=DATALOG()` worksheet functions for
querying and logic programming over live tables.

**The commercial thesis in three sentences:**

1. The AI era made *plausible* automation free and *provable*
   automation scarce — and the people who run month-end close,
   regulated reports, and operational checklists need provable.
2. The product's architecture (deterministic grammar, effect logging,
   no network calls, open core) is accidentally-on-purpose the exact
   shape that compliance-grade spreadsheet automation requires — a
   category (End-User Computing risk) that today spends real money on
   tools that merely *watch* the problem.
3. Revenue comes from services around evidence, assurance, and
   governance — never from the language itself, which stays open. That
   split is already project doctrine (standing decision SD-10),
   decided before it was a strategy.

**The honest stage:** working product approaching beta; the pilot
program (one named user, one real SOP) is the next milestone and the
falsification test for everything below.

---

## 2. The problem, in the buyers' own units

Spreadsheet automation has three constituencies, and they experience
the same gap three different ways:

**The analyst — hours.** Recurring manual procedures: month-end
formatting, data cleanup, report assembly, reconciliation prep. A
90-minute weekly procedure is ~75 hours a year; most analysts run
several. The existing options are: learn VBA (they won't — that's a
career change, not a feature), record a macro (brittle, unreadable,
unmaintainable), or ask an AI prompt box (instant, plausible, and
confidently wrong just often enough to be disqualifying for numbers
that matter).

**The manager — fragility.** The real cost isn't the hours; it's that
the procedure lives in one person's head. It breaks on handover,
drifts with every re-telling, and turns staff turnover into
operational risk. Managers don't want automation so much as they want
*procedures that survive people*.

**The executive and the auditor — attestation.** Under SOX 404 and
equivalent regimes, key spreadsheets in financial reporting are
controls, and today the evidence for a manual procedure is *a human's
attestation that they followed a document*. "End-User Computing (EUC)
risk" is a named compliance category with dedicated budgets; an
industry of monitoring tools (inventory, change-detection, access
logs) exists at five-figure annual price points to *watch* risky
spreadsheets without making them less risky. The procedure itself —
the thing that actually drifts — stays manual.

The gap, stated once: **there is no way for a non-programmer to turn a
written procedure into something that runs exactly as written, refuses
to run otherwise, and leaves evidence.** That is the product.

---

## 3. The product, briefly

*(The repository's README carries the full tour; this is the
sixty-second version, with what each property is commercially for.)*

- **Checked English.** `Make cell A1 bold.` `Repeat 5 times:` `If
  grand is greater than 40: Put "PASS" into cell C4.` A phrasebook
  grammar (patterns + templates, loaded as data) translates
  mechanically to a 1:1 s-expression layer over VBA — no statistics,
  no model, no temperature. *Commercially:* determinism is the entire
  compliance pitch; it is also the property no LLM product can offer.
- **Check before run.** Every sentence validates first; errors land on
  their exact row, in words, in a teaching register. *Commercially:*
  refusals are the onboarding — the product teaches its own grammar at
  the moment of failure.
- **Two backends, one contract.** Interpreted by default (zero trust
  settings, nothing injected — the version IT can say yes to);
  exported readable VBA on demand (the version an auditor can read).
  Held to parity by a golden-file corpus. *Commercially:* the
  IT-approval path and the audit path are both first-class.
- **Query and logic.** `=SQL("SELECT Name, Salary FROM staff WHERE
  Salary > 80000", Staff)` — real SQL text, frozen subset, spilled
  results; `=DATALOG(...)` with recursive rules for
  reachability-shaped questions plain formulas can't express.
  *Commercially:* the demo that makes technical evaluators sit up, and
  a second wedge into the same workbooks.
- **Offline, permanently.** No network calls — no telemetry, no update
  checks — by standing decision (SD-13), not by roadmap gap.
  *Commercially:* the single sentence that moves a managed-environment
  security review from "no" to "keep talking."
- **Undo means undo.** Runs snapshot first. *Commercially:* the demo
  moment that converts the risk-averse.

---

## 4. Why now

Three trends converge, none of which this project controls and all of
which it benefits from:

1. **LLMs commodified plausible automation** — and thereby created,
   for the first time, a *market distinction* between "probably does
   what you meant" and "provably does what you said." Before Copilot,
   determinism was an implementation detail; now it is a category
   axis. Frazaro sits alone on the far end of it in the Excel context
   (`ADVOCATUS.md` A.1 argues the bear case; its falsification test is
   the pilot).
2. **Microsoft is quarantining the macro layer** (Mark-of-the-Web
   blocking 2022, AMSI inspection, trust-gate narrowing — the full
   thirty-year record is read in `SUBSTRATE.md`). Every tightening
   makes traditional VBA automation scarier for IT — and Frazaro's
   default runtime needs none of the frightening settings. The
   platform's own risk posture is a tailwind for the zero-trust
   design.
3. **EUC/spreadsheet-risk scrutiny keeps rising** post-scandal after
   post-scandal, while the tooling industry that grew around it only
   monitors. A product that *removes* the manual-execution risk, in
   the user's own language, with evidence, enters a budget line that
   already exists.

---

## 5. Market, sized honestly

*Directional figures; the bottom-up wedge is the number that matters
at this stage.*

- **Top of the funnel (context, not TAM):** Excel's user base is
  commonly cited at three-quarters of a billion to over a billion;
  Microsoft markets the formula layer as the world's most popular
  programming language. Nobody captures markets that size; the figure
  matters only because it makes every niche below enormous in
  absolute terms.
- **The adjacent spend (evidence the budget exists):** RPA — the
  enterprise answer to "humans repeat procedures" — is a
  multi-billion-dollar category, and its persistent mid-market
  failure mode (too heavy, too IT-dependent, too far from the
  spreadsheet where the work actually lives) is exactly the gap a
  worksheet-native tool enters. EUC-risk tooling sustains multiple
  vendors at five-figure-per-year price points for monitoring alone.
- **The wedge (the honest target):** mid-market finance, accounting,
  and operations teams — organizations large enough to have SOPs,
  month-end close, and audit exposure, small enough to have no RPA
  program and no development staff. In the US alone this is tens of
  thousands of firms; each has a handful of analysts and dozens of
  recurring procedures. The beachhead unit is not a company but a
  *procedure*: value per procedure (Section 6) times procedures per
  team makes even a few hundred customers a durable business.
- **What the pilot must establish:** that one real team's SOPs land
  within teaching distance of the grammar (`PI.2`'s transcript is
  the designed falsification test) and that the trust posture clears
  one real IT review (`PI.7`). Every market claim above is downstream
  of those two facts.

---

## 6. Business model — four streams, all open-core-compatible

The constitution came first: **SD-10** (in force since long before
this document) — *price seats, support, the registry, governance,
hosted services; never a core form, a backend, or a grammar section.*
The language is free forever. Certainty, evidence, and governance are
the products.

### 6.1 Concierge SOP translation — *paid discovery* (active first)

"Send us your SOP; get back a running Frazaro program, plus the
grammar we grew to hold it." Priced per procedure in the low thousands
— defensible against the consultancy rates the same buyer already
pays for far more brittle automation. **Margin honesty:** this is
founder time, the scarcest currency; volume must stay capped, and the
stream's real yield is not the fee — it is *specimens*. Every paid
engagement feeds the corpus, the gap log, and the refusal vocabulary:
the customer funds exactly the asset that compounds (the phrasebook is
the moat — Section 8). This is a discovery program that pays for
itself, not a services business to scale.

### 6.2 Assurance subscription — *the Red Hat move* (activates at beta)

The software is free; **certainty costs**: guaranteed response times,
a security-advisory channel (signed release hashes and notices the
*customer* subscribes to — which resolves the project's
no-phone-home doctrine and its patch-delivery tension in one move,
`ADVOCATUS.md` A.5), and version-pinned long-term support. Shape:
hundreds of dollars per organization per month, sold to **dozens of
organizations, not thousands of seats** — deliberately, because
support capacity is the binding constraint (`VIABILITY.md` N.3) and
the pricing must make dozens sufficient. Fifty organizations at a
mid-three-figures monthly price is meaningful annual recurring
revenue at software margins, inside one person's support capacity.

### 6.3 The compliance pack — *the executive's line item* (activates with per-run evidence)

Signed, hash-chained run logs; retention policy; the IT-reviewer
documentation bundle; a control-attestation report generator. The
pitch, in the buyer's units: *"Your key spreadsheet control stops
being a document a human attests to and becomes the executable
itself. The procedure that ran is the procedure the auditor reads;
deviation is not detected but impossible; the run log is the
evidence."* Priced per legal entity per year, five figures at the top
— calibrated against what the same buyer already pays EUC tools that
only watch. **Dependency, stated plainly:** this stream waits on the
per-run evidence feature (roadmap item `U.18` — the machinery exists;
the user-facing artifact is unbuilt) and on one practicing auditor's
verbatim reaction (`SIG.6`), which the project's own devil's-advocate
review demanded before believing the pitch. This is the
highest-willingness-to-pay stream and the least proven; the document
refuses to pretend otherwise.

### 6.4 The org phrasebook registry — *governance as a service* (activates with a commons)

A company's own vocabulary — its account names, its report verbs, its
approved procedures — as a private, signed, provenance-tracked
phrasebook layer, hosted and access-controlled. The public commons
stays free; private registries are seat-priced. This is the true
platform stream (network effects live here) and the furthest out; it
waits for a contributor commons to exist at all.

**The sequence is the strategy:** concierge funds and feeds the corpus
now → assurance monetizes trust at beta → compliance monetizes
evidence when `U.18` ships → registry monetizes governance when there
is a commons to govern. Each stream's activation gate is an existing
roadmap item, not an invention of this document.

---

## 7. Go-to-market

- **Concierge-led land.** The first motion is manual and
  high-touch by design: pilot-shaped engagements that produce a
  running procedure in days. The analyst gets hours back and becomes
  the champion; the transcript becomes corpus.
- **Manager expand.** The champion's procedure survives a handover —
  the moment the manager notices, the conversation moves from "a tool
  Dana uses" to "how our team's checks run." Assurance subscription
  enters here.
- **Compliance up-sell.** At audit season, the run log is shown to an
  auditor; the compliance pack conversation starts with evidence
  already in hand. (This is a hypothesis with a designed test —
  `SIG.6` — not an assumption.)
- **Open source as top-of-funnel.** The repo, the README, and the
  unusually complete documentation shelf are the technical
  evaluator's diligence path; the no-network posture and readable
  generated code are the IT reviewer's. Both personas can say yes
  without a sales call — which is the only acquisition channel a
  solo company can afford to operate.

---

## 8. Competition and moat

| Alternative | What it offers | Where Frazaro wins |
|---|---|---|
| **Copilot / LLM prompt boxes** | Instant, free, zero syntax | Determinism: provably does what the sentence says or refuses; no hallucinated numbers; evidence trail. The categories don't overlap where money is regulated. |
| **Office Scripts / Python in Excel** | Real languages, Microsoft-native | They are *for programmers* — the constituency that was never the problem. Python in Excel executes in Microsoft's cloud (a non-starter for the offline/compliance buyer). |
| **Recorded macros / VBA consultants** | Familiar, cheap to start | Unreadable, unauditable, brittle artifacts; consultant output has no check-before-run and no grammar to teach the next maintainer. |
| **RPA (UiPath et al.)** | Enterprise-grade orchestration | Heavy, IT-owned, and outside the workbook; the mid-market team with five SOPs will never deploy it. |
| **EUC monitoring tools** | Compliance budget incumbents | They watch the risk; Frazaro removes it. Potentially channel partners rather than competitors. |

**The moat, honestly ranked:**

1. **The corpus.** The phrasebook — hundreds of auditioned grammar
   rules with pinned refusal wording — is the asset that compounds,
   and the concierge stream feeds it with paid specimens. A
   competitor can fork the engine (it's open); the years of
   audition decisions are the part that doesn't fork meaningfully
   forward without the discipline that produced them.
2. **Provable determinism as a claim.** The grammar's restriction
   (SD-16: no backtracking, first-match-wins, load-time ambiguity
   refusal) makes "one sentence, one meaning" a *provable* property,
   not a marketing sentence. Competitors built on LLMs cannot make
   the claim at all; competitors built on parsers must rebuild the
   restriction and the teaching-refusal machinery it enables.
3. **The trademark, not the license.** In open-core, the code forks
   but the name doesn't. "Frazaro" is the enforcement point
   (roadmap `GO.4`).
4. **Switching costs on the buyer's side.** Once a team's SOPs *are*
   Frazaro programs, the programs are readable English — trivially
   portable in theory, and in practice the operational habit, the
   org phrasebook, and the audit history are the lock-in. The moat
   is sticky in the polite way: the customer could leave and has no
   reason to.

---

## 9. Traction and current state

*Stated with the same bluntness the internal docs use, because a VC
will find these facts in an afternoon and should hear them from the
document first:*

- **Working:** the full English → VLA → interpret/export pipeline;
  the worksheet IDE; `=SQL()` (through INNER JOIN, as of the day of
  writing) and `=DATALOG()` (recursive rules, named-column atoms) —
  all owner-verified live under a golden-file + self-test + dual-
  backend-parity discipline that most funded teams do not match.
- **Version `0.5.0` beta targeted 2026-09-05.**
- **Users: zero.** The pilot program — one named user, one real SOP,
  one workbook, kill criteria written before the run — is scoped as
  the immediate next milestone and is deliberately designed as the
  falsification test for the mission premise, not as a friendly
  demo.
- **Revenue: zero.** Everything in Section 6 is design, not report.

---

## 10. Team and method

One developer, pair-programming with an AI assistant, under a
documented collaboration protocol: the assistant writes VBA blind,
the human is compiler, runtime, and QA lab, and the gap is closed by
machinery — golden files, self-test suites, build-time phrasebook
audits, dual-backend parity checks. The honest reading of that
arrangement cuts both ways and the project wrote both readings down
itself (`ADVOCATUS.md` A.6, `CONTINUITY.md` Ω): the redundancy is
mechanical rather than human, *and* the defect-escape record is
documented in a way most teams' are not. The method — every rule
carrying the incident that earned it, every review adversarial, every
claim owner-verified before a checkmark — is itself a due-diligence
artifact: an investor evaluating process maturity can read it
directly instead of inferring it from interviews.

---

## 11. Risks — pre-disclosed, with the diligence already written

This project commissioned five adversarial reviews of itself in one
sitting and published them in-repo. A VC's risk section is usually
extracted under questioning; this one is shelved and indexed. The top
five, distilled, with where the full argument lives:

1. **Key-person risk** — one founder, evenings-funded
   (`PREMORTEM.md` D.1; mitigations minted in `CONTINUITY.md`:
   off-machine repo, succession brief, bus-factor drill).
2. **Demand risk** — checked English may lose to free-form AI in
   perceived value even where it wins on merit (`ADVOCATUS.md` A.1,
   the strongest bear case in the repo; falsification test designed:
   the pilot transcript).
3. **Platform risk** — Microsoft quarantines the macro layer further
   or faster than the thirty-year base rates suggest
   (`SUBSTRATE.md`: full historical read, watch-list with
   pre-decided responses; the default runtime already needs none of
   the endangered trust settings).
4. **Compliance-sales risk** — long cycles, vendor-stability
   objections against a solo company (mitigation unique to the
   model: the open core *is* the escrow; "our continuity plan is
   `git clone`" is an answer funded competitors cannot give).
5. **Death by documentation** — the failure mode where process
   excellence substitutes for users (`PREMORTEM.md` D.7 — named,
   tripwired, and watched by the project's own review cadence rule,
   SD-17).

---

## 12. Open-source structure and licensing plan

The segmentation the revenue model implies (final license choices are
roadmap item `SIG.0`/`GO.4`, deliberately made *after* the model is
chosen, in this order):

- **Open, forever:** the engine, both backends, the English
  phrasebook, the worksheet IDE, `=SQL()`/`=DATALOG()`. License
  either permissive (Apache-2, adoption-maximizing) or copyleft
  (GPLv3, capture-resistant) — **with, non-negotiably, an explicit
  output exception**: generated VBA embedded in customers' workbooks
  must be unambiguously theirs (the GCC-runtime-exception pattern).
  Without that clause every compliance buyer walks.
- **Open, share-alike (candidate):** the corpus itself — extensions
  to the public phrasebook flow back, keeping the compounding asset
  common even if the engine is permissive.
- **Closed or fair-source:** the compliance pack's report/attestation
  tooling and the registry server — services around the language,
  never the language.
- **Held:** the trademark, which is the actual commercial boundary in
  open-core.
- **Docs:** permissive (they are, as this document demonstrates, part
  of the sales motion).

---

## 13. The ask, framed honestly

This document precedes a raise decision rather than announcing one.
The base case is explicitly viable without capital: the project is
evenings-funded, the sequence in Section 6 is solvent at every step
(concierge pays as it discovers; assurance prices to fit inside solo
support capacity), and `VIABILITY.md`'s books say a
dozens-of-customers business at software margins is a durable outcome
— a real company, and honestly perhaps not a venture-scale one.

**The venture case is the platform outcome,** and it is a different
bet: if "the SOP is the executable and the run is the evidence"
becomes how mid-market finance teams operate, the registry layer
(Section 6.4) carries network effects, the phrasebook model extends
beyond English (the architecture is language-neutral by standing
decision) and eventually beyond Excel (the backend seam is already a
fork point, held open at no carrying cost). Capital accelerates
exactly three things: the compliance pack's build-out (`U.18` plus
certification-grade evidence handling), a real go-to-market motion
into the wedge market, and a second pair of hands that converts the
key-person risk from mitigated to retired.

What this document will not do is pretend the base case is the
venture case. An investor who wants the distinction blurred is not
the right investor for a project whose entire product thesis is that
provable beats plausible.

---

## Appendix — the diligence shelf

| Document | What it holds |
|---|---|
| [README.md](../README.md) | The product, for a technical evaluator |
| [BETA_ROADMAP2.md](BETA_ROADMAP2.md) | Current plan, open items, standing decisions |
| [VIABILITY.md](VIABILITY.md) | Internal economics audit (the accountant) |
| [PREMORTEM.md](PREMORTEM.md) | Eight ways this dies, with tripwires (the coroner) |
| [ADVOCATUS.md](ADVOCATUS.md) | The bear case, argued to win (the advocate) |
| [CONTINUITY.md](CONTINUITY.md) | Key-person risk, audited (the successor) |
| [SUBSTRATE.md](SUBSTRATE.md) | Platform risk, thirty years of base rates (the historian) |
| [CONSULTANT.md](CONSULTANT.md) | The original outside assessment |
| [LESSONS.md](LESSONS.md) | Engineering culture, shown not told |
| [DEPLOY.md](DEPLOY.md) | Distribution and signing reality |

*End of memorandum.*
