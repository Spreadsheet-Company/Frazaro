# AUDIT — `BETA_ROADMAP.md` read against `METAMETALISP4.md`
 
*Three parts, written in one pass. **Part I** audits the master roadmap under
the essay's own doctrine — with §7's invitation taken literally, so the
antithesis arrives wearing several other people's uniforms. **Part II** answers
the phrasebook-readability question. **Part III** is the horizon, which is the
only part written to be enjoyed rather than acted on.*
 
*Standing assumption, given by the owner and applied throughout: **launching a
beta outranks engineering a perfect beta.** New Jersey over Cambridge. Where
this audit and the roadmap disagree, that assumption is usually why.*
 
---
 
# PART I — THE ROADMAP AUDIT
 
## I.0 — The verdict, in one paragraph
 
`BETA_ROADMAP.md` is an unusually good strategy document that is optimizing for the
wrong bottleneck. It correctly identifies rising-cost work, correctly refuses to
hydrate distant tranches, correctly cuts departments by accountability, and
correctly names the microscope discipline it inherited. What it does not contain
— anywhere in fifteen tranches, roughly a hundred items, and a closing "read
nothing else" summary — is **a single item whose completion requires another
human being.** The essay anticipated this exact failure and named it twice: §5's
field mark (*"the fake microscope keeps acquiring lenses; the real one keeps
acquiring specimens"*) and §8's warning that the forge is warm. The roadmap is
currently a beautiful catalogue of lenses. Every recommendation below is
downstream of that single observation.
 
---
 
## I.1 — What the roadmap already gets right
 
Stated first because it is true, and because the critique lands differently once
it is clear the critique is narrow.
 
1. **The flat/rising distinction is correctly generalized.** §3's cost-of-delay
   mechanism is not just cited, it is *applied to decisions the essay didn't
   cover* — exactly the transfer the essay hoped for. The observation that a
   surface-compatibility policy gets more expensive per rule shipped is a real
   expansion-count argument, arrived at independently.
2. **Departments cut by accountability, not subject.** "A different question it
   must answer, a different reviewer profile, a different definition of done" is
   a better cut than Computer/Human/Community and would survive contact with an
   actual team.
3. **Thin later tranches, declared as discipline rather than apologized for.**
   Governance-in-detail-today-would-be-fiction is correct and rarely said aloud.
4. **F.1 is genuinely the highest-leverage item in the document**, and the
   roadmap knows it. One sentence of doctrine, expansion count in the hundreds of
   rules. It is also — see Part II — the single largest readability lever on
   `english.vla`, which the roadmap does not currently notice.
5. **The closing extraction paragraph is the philosophy working.** "Each is a
   paragraph today and a rewrite later" is the whole essay compressed, and the
   list (GO.1, CO.1, EN.2/3, LX.6, AC.1) is well chosen.
6. **AC.1 is honest.** A two-hour accessibility fix, ranked by when it stops
   being two hours, in a tranche the author openly calls thin. That is what
   integrity looks like in a backlog.
 
---
 
## I.2 — Finding 1: the roadmap's one unbound symbol
 
`first external user` is referenced as a gate in **Tranche 4** ("the whole
tranche gates on first external user"), as a precondition in **DO.4** ("required
before the first outside contributor"), as a payoff target in **LE.6** and
**DO.6** ("pays into beta acquisition"), and as a temporal landmark in **DO.5**
("at beta cutoff"). It is defined and scheduled **nowhere.**
 
A compiler would refuse this file. Four call sites, no binding.
 
This is not a documentation nit. It means the roadmap's most consequential
sequencing decisions — what is cheap now versus ruinous later — are all
expressed relative to an event that no item in the document causes. The gate is
load-bearing and unowned, which is the structural signature of a date that
slips forever without anyone deciding to slip it.
 
**Fix:** define beta in one line, with a scope and a date, in the strategy file
itself, and give it an owner. Then every "gates on first external user" becomes
a real constraint rather than a deferral with good manners.
 
---
 
## I.3 — Finding 2: specimen count zero
 
§5's field mark, applied bluntly:
 
| | Count |
|---|---|
| Tranches | 15 |
| Items | ~100 |
| Items whose "done" is a stranger doing something | **0** |
| Real user sentences cited as motivation for any grammar section | **0** |
| Instances of the word "pilot" in `BETA_ROADMAP.md` | **0** |
 
The Alpha 5 ledger *does* contain a pilot (E3, with a kill criterion and a
near-miss log, and a note that "the pilot's log re-ranks everything"). That was
the right instinct and it did not survive the promotion into strategy. The
master roadmap inherited the engineering and dropped the patient.
 
Notice also what this does to the Grammar tranche's ordering. It is sequenced by
`pareto.txt` — a *predicted* pareto, derived from reasoning about what SOPs
probably contain. G-TAIL alone is honest about it: *"reordered by real
sentence-gap reports, not by this file."* That sentence should govern the whole
tranche. Right now the corpus is being planned the way the essay's §8 warns
about: *a team that does not yet know what it is building will construct the
wrong microscope, beautifully.*
 
Every hour spent guessing the grammar's priority order is an hour that one
afternoon with one bookkeeper's month-end file would settle by measurement.
 
---
 
## I.4 — Finding 3: the roadmap names its own bottleneck and ranks it eleventh
 
**F.11 — split the suite into pure / host-required halves.** *Why now: a human
at a Windows box is the velocity ceiling and the bus factor. Pays into: every
tranche, forever.*
 
That is a Theory-of-Constraints identification, written in the roadmap's own
hand, and then filed last in Tranche 0. Goldratt's rule (§7) is unambiguous:
improvement anywhere except the constraint is waste by definition. If a human at
a Windows box is the ceiling, F.11 outranks F.2, F.4, F.5 and probably F.1 on
throughput grounds, and the entire Optimization tranche is — under TOC — noise.
 
Two candidate constraints are visible in this project. The roadmap has already
identified one (the Windows box) and buried it. It has not identified the other
(zero users), which §7 says is the constraint that no amount of tooling can
move. **A strategy file should open by naming its constraint and should re-name
it every version.** Right now the constraint has to be inferred from an aside in
item eleven.
 
---
 
## I.5 — Finding 4: Tranche 1 is anticipatory where the essay demands measured
 
The mission — *whatever language a person prefers to think in* — is genuinely
beautiful and probably correct. The elevation of Linguistics to near-first
position rests on three claims, and they do not all hold equally.
 
| Claim | Assessment |
|---|---|
| Rising cost | **True, and quantified** — 102 fold sites, 63 head words, four raise sites. This is exactly the evidence §5.1 demands. |
| Gates Grammar | **True for LX.7** (a dialect spec authors can target). Weak for LX.2/LX.5, which gate a *second language*, not a first grammar. |
| Can invalidate the mission if deferred | **Unproven.** It can invalidate the mission *if there is a non-English user*. There is no user at all. |
 
§7 is direct about this shape of argument: *"TOC usually wins, because TOC is
measured and Metametalisp is anticipatory — one reads the gauge, the other reads
the future, and gauges have the better track record."*
 
The tranche is not wrong; it is **undifferentiated**. It bundles items with
wildly different curves under one elevation:
 
- **Cheap, irreversible, hours: do now.** LX.3 (locale-safe fold — a live
  correctness bug), LX.4 (alias table at one chokepoint before it is 63), LX.6
  (identifier policy — a paragraph), LX.1 (*if written as a decision list rather
  than an audit project*).
- **Expensive, reversible, weeks: defer past beta.** LX.2 (the catalogue as an
  artifact), LX.5 (splitting a 3,000-line module), LX.9 (audition lint).
- **Cheap, high-information, timeboxed: do now, as a test.** LX.10 — twenty
  sentences in one non-English language. The roadmap already calls this "the
  only honest proof." One day, not one tranche. It is the falsification
  experiment for the entire mission and it costs almost nothing; it should
  arguably run *before* the expensive half of Part A rather than after.
 
**Recommendation:** split Tranche 1 into **1a — chokepoints (hours, do now)**
and **1b — catalogues (weeks, post-beta)**. The mission survives intact; the
schedule stops paying for a market it has not met.
 
---
 
## I.6 — Finding 5: the unpriced irreversibility
 
The roadmap prices every internal irreversibility carefully and does not price
the external one at all.
 
Frazaro compiles to VBA, injects through the VBProject object model, and ships
as a macro-enabled add-in. That entire substrate is a single vendor's deprecated
surface. Microsoft has spent years narrowing it — macro blocking by default for
internet-sourced files, Office Scripts and Python-in-Excel as the sanctioned
successors, VBProject trust as a managed-tenant policy switch. **One
configuration decision in Redmond can zero this product**, and the roadmap's
only acknowledgements are EN.6 (Mac has no VBProject) and DI.1 (signing).
 
The mitigation is nearly free *today* and structural later, which is precisely
§3's definition of a decision to buy while it is cheap:
 
> **Standing decision:** VLA is the target-independent middle layer. VBA is a
> backend, not the language. No Tier-1 form, no prelude macro, and no
> phrasebook template may assume the VBA backend; anything that must is a
> declared, listed exception in the runtime helpers.
 
That sentence costs nothing to adopt now, because the architecture is already
90% compliant — English emits VLA, VLA emits VBA. It becomes a rewrite the day a
hundred templates have quietly grown `xl*` constants (Part II counts 11 already)
and `application.` reaches (61 rules already). Note that **F.1 is also this
mitigation** — the ABI that keeps templates off raw dot-forms is exactly what
keeps a second backend possible. F.1's expansion count is even larger than the
roadmap claims.
 
Second, related, unpriced: **Hyrum's Law arrives at the beta, not after it.**
Tranche 4 says surface changes are free before contact and expensive after,
which is right, and then places the whole tranche after contact. CO.1 (retired
spellings refuse with directions) is *one refusal path*. It must exist **at**
first contact, not be scheduled from it. Split CO into policy-before and
tooling-after.
 
---
 
## I.7 — Finding 6: the gates have no expiry dates
 
§7's reconciliation with real-options theory is the sharpest sentence in the
essay: *decide early only where the option is about to expire — irreversibility
is the expiry date.*
 
The roadmap has five 🔒 gates. None of them says when it expires or what
observation would open it.
 
| Gate | Currently | Proposed expiry condition |
|---|---|---|
| F.1 🔒 before Grammar | Absolute | Keep absolute. It is one sentence and it is free. |
| F.2 🔒 before Grammar | Absolute | **Expires at rule #175, or first outside contributor.** String-concatenated templates are painful in proportion to rule count; that curve is linear, not vertical. |
| F.4 🔒 before Grammar | Absolute | **Expires at first user-authored or org phrasebook.** At 102 rules with a load-time duplicate-shape audit already running, a human can still audition. At the moment strangers write rules, silence becomes fatal. |
| EN.3 🔒 before Grammar §6 | Absolute | Keep — but note it is really "before the first non-US-locale user," which may be the same day as the beta. |
| G9 🔒 on pilot evidence | Correct | **This is the best-formed gate in the document** — it names the observation that opens it. Copy its form to the other four. |
 
An absolute gate with no expiry is not sequencing. It is a deferral that has
been given a lock icon so nobody argues with it.
 
---
 
## I.8 — Finding 7: no time units anywhere
 
`BETA_ROADMAP.md` contains no estimates, no appetites, no durations, and no dates.
That is defensible for a strategy file — but it has a specific consequence: a
document with ~100 items and no denominator **cannot be prioritized by
cost-of-delay at all**, because cost of delay is a rate, and a rate needs a
duration to rank against.
 
Reinertsen's CD3 (cost of delay ÷ duration) is the missing arithmetic. Applied
casually to items the roadmap already describes:
 
| Item | Cost of delay | Duration | CD3 |
|---|---|---|---|
| LX.3 locale fold | Silent correctness bug, 102 sites | hours | **very high** |
| F.6 one raise site | 4 sites → 100 | hours | **very high** |
| AC.1 colour-only status | 2h now, refactor later | hours | **high** |
| F.1 ABI doctrine | Every rule shipped | one sentence | **effectively infinite** |
| LX.5 module split | Fork vs sibling, later | weeks | moderate |
| AS.3 mutation testing | Unknown pin quality | weeks | low **pre-users** |
| Tranche 5 (all) | Compiler latency at 1,000 rules | weeks | low today |
 
The roadmap's own "short answer" ordering is close to CD3-correct at the front
and drifts after. Adding a crude appetite tag — `~hours` / `~days` / `~weeks` —
to every item would take an afternoon and would visibly reorder the middle.
 
---
 
## I.9 — The cheapest move in the document: convert artifacts into standing decisions
 
This is the single recommendation with the best ratio in this audit, and it is
the essay's own central claim applied to the roadmap that cites it.
 
§1 is explicit: **a metametamacro is a standing decision, not a tool.** *"The
linter is a tool. 'CI blocks the merge when the linter objects' is a
metametamacro."* Several roadmap items are currently specified as **artifacts**
(weeks of work) when the rising-cost half of their value is available today as a
**decision** (one sentence). Buy the decision now; defer the artifact past beta.
 
| Item, as specified | The decision that captures most of its value today | What you defer |
|---|---|---|
| **LX.2 — the message catalogue** | *"No refusal ships as a raw string. Every refusal goes through `Raise` with a stable ID and named parameters, even while the English text lives inline."* | Building the catalogue, the lookup, the translations. The IDs accrete for free; the catalogue is later assembled from something that already exists. |
| **LX.7 — the dialect spec** | *"No rule ships without an audition, and any audition decision not already covered by the checklist is appended to the checklist that day."* | Writing a spec up front for a grammar you have not finished discovering. The spec accretes as a byproduct of the work, with reasons attached, which is what §3 actually asks for. |
| **CO.2 — the frozen compatibility corpus** | *"A sentence that has ever appeared in a shipped `instructions.txt` keeps its spelling, or is retired through CO.1's refusal."* | The separate file, the tooling, the split. The promise is the expensive part; the file is bookkeeping. |
| **F.7 — formula dialect as declared subset** | *"The formula backend must either support or explicitly refuse every emitter case; a new case with neither is a failing test."* | The operator table, until a formula rule needs it. |
| **AS.6 — corpus-family contract** | Already a decision. Good. Cite it as the model. | — |
| **Backend plurality (I.6)** | *"VBA is a backend, not the language."* | Everything. Costs nothing today. |
 
Six sentences. Most of the rising-cost protection of roughly two months of
tranche work, bought before the beta rather than instead of it. This is
arbitrage in exactly §3's sense.
 
**Corollary, structural:** the roadmap should carry a **standing-decision
register** — numbered, dated, with reasons attached — separate from the backlog.
Today that register exists as one five-clause paragraph ("the fifth thing, which
is not a department: the Protocol"). It is the most valuable content in the file
and it is the least elaborated. §3's civic duty (*declare them with their
reasons attached*, so the engineer of 2031 can tell doctrine from fossil) is not
satisfied by a list of five noun phrases. Expand it into an ADR-style register;
it is append-only, it never needs pruning, and it is the artifact most likely to
outlive the codebase.
 
---
 
## I.10 — The three tests, applied per tranche
 
§5, run mechanically. *Dogfood* is scored strictly, as the essay demands: **you,
this week** — not the team, not Q3, not a persona.
 
| Tranche | Rising cost named with a number? | Proportion (tool < work) | Dogfood (you, now) | Verdict |
|---|---|---|---|---|
| 0 Fortifications | ✅ 63 heads, 4 raise sites, 102 folds | ✅ mostly one-sentence | ✅ | **Keep first.** Reorder F.11 up. |
| 1 Linguistics | ✅ for LX.3/LX.4; ❌ for LX.2/LX.5 | ⚠️ catalogue > current need | ❌ except LX.3/LX.10 | **Split 1a/1b.** |
| 2 Environment | ✅ per-rule inheritance (EN.2/3) | ✅ | ⚠️ | Keep policies, defer matrix. |
| 3 Assurance | ⚠️ "557 without a denominator" is a good argument | ⚠️ AS.3/AS.5 are large | ⚠️ | AS.1 now; rest post-beta. |
| 4 Compatibility | ✅ Hyrum | ✅ CO.1 is one path | ❌ (no strangers yet) | **CO.1 before contact**, rest after. |
| 5 Optimization | ❌ no user-visible number | ⚠️ | ❌ | **Demote.** Correctly gated on P-PROF; wrong position at 5. |
| 6 VLA | ✅ demand-driven by design | ✅ | ✅ | Model tranche. Leave alone. |
| 7 Grammar | ✅ | ✅ | ✅ | **This is the product.** Re-sequence from pilot gaps, not `pareto.txt`. |
| 8 Performance | ✅ "expert idioms automatically" is the pitch | ✅ PF.2/PF.3 are cheap | ✅ | **Promote above 5.** |
| 9 Interface | ⚠️ | ✅ | ✅ U.15 | Keep; V.1 is undervalued (see Part III). |
| 10 Learnability | ⚠️ | ✅ | ✅ LE.6 | **Promote LE.1/LE.6 into the pilot bundle.** |
| 11 Interoperability | ✅ IO.1 breaks first on a real machine | ✅ policy | ❌ | IO.1 is a pilot prerequisite. |
| 12 Distribution | ✅ DI.1 gates enterprise | ✅ | ❌ | DI.1/DI.2 are pilot prerequisites. |
| 13 Accessibility | ✅ AC.1 quantified | ✅ | ⚠️ | AC.1 now, as written. |
| 14 Documentation | ✅ carve-out reasoning is sound | ✅ | ⚠️ | DO.6 is really acquisition; move to 10. |
| 15 Governance | ✅ GO.1 rising and structural | ✅ paragraph | ❌ | GO.1 paragraph now; rest fiction, correctly. |
 
Two patterns fall out. **First:** the tranches that fail the dogfood test are
exactly the ones written for a population that does not exist yet — and the
roadmap's own defence of thin tranches ("specifying Governance in detail today
would be fiction") applies to them with equal force but was not applied.
**Second:** the roadmap ranks compiler-internal speed (5) above emitted-code
speed (8), interface (9), learnability (10), interoperability (11) and
distribution (12) — that is, above every tranche a user can perceive. Under a
launch-first assumption that ordering inverts.
 
---
 
## I.11 — Imports, per §7's invitation
 
§7 asks for the antithesis honestly. Here it is in five other people's words,
each with one concrete change to make.
 
**Theory of Constraints (Goldratt).** *Name the constraint at the top of the
file; re-name it every version.* Today: the Windows-box human (F.11) and the
absent user. Everything not at one of those two is, by definition, waste this
version.
 
**Shape Up (Basecamp).** *Appetite, not estimate.* Give each item a fixed time
budget and let scope vary. A six-week appetite on "the pilot slice of grammar"
produces a shipped slice; an unbounded G-FORMAT produces seventy immaculate
rules and no user. Also adopt the **betting table**: at each version open,
re-bet from the whole roadmap rather than draining tranches in order. The
roadmap half-says this already ("tranches are a gradient, not a queue") and then
publishes a queue.
 
**Lean / hypothesis-driven development.** *Each tranche states the belief it
tests and the observation that would falsify it.* Tranche 1's belief is "a
non-English market exists and is reachable." LX.10 is its experiment. Run the
experiment before funding the tranche. This generalizes: the roadmap has no way
to be **wrong**, and a plan that cannot be wrong is §7's "mood with citations."
 
**Cynefin (Snowden).** *Complicated ≠ complex.* Compiler internals, locale
folds, and emitter coverage are **complicated**: expertise suffices, analysis
works, front-loading is correct. Which sentences people actually need is
**complex**: unknowable by analysis, discoverable only by probing. The roadmap
applies one method (analyze, then build) to both domains, and serializes them.
The correct move is to run them **in parallel** — probe the complex domain with
a pilot while the complicated work proceeds — which also happens to be the
cheapest way to keep a solo builder from stalling.
 
**Wardley mapping.** Place components on evolution: the **dialect + corpus** is
genesis (the moat, correctly identified); the **compiler** is heading to
product; **performance** is utility. Wardley's advice is to invest in genesis and
be ruthlessly pragmatic elsewhere — which agrees with demoting Tranche 5 and
promoting Grammar. The map also makes the gap visible: *user need sits at the
top of the map and is currently attached to nothing.*
 
**Concierge MVP / design partner (YC, and every services-to-product company).**
*Do it manually for one customer before automating.* For Frazaro this is
unusually powerful, because the manual fallback is already built: for any
sentence the corpus refuses, the owner writes the VLA form by hand. The user
gets a working automation on day one; the project gets a **gap log**, which is
the empirically-ordered grammar backlog that `pareto.txt` is currently guessing.
 
**Diátaxis (Procida).** For Tranche 14, free structure: tutorial / how-to /
reference / explanation. Note that LE.1 (palette) and DO.2 (message catalogue)
are *reference* by construction and are generated — which is why they can be
built early and never go stale, while DO.3 (the user guide, *explanation*) is
correctly deferred. The distinction the roadmap draws by instinct is a known
taxonomy; borrowing it makes the carve-out argument in one line instead of a
paragraph.
 
**And one anti-import, for honesty.** Everything above pushes the same
direction, which should itself be suspicious. The essay's §9 lists the
conditions under which its own philosophy is correct — long horizon, high rework
cost, immovable foundations, builder-is-user, compounding domain. **Frazaro
scores five out of five.** A language with a corpus is the most build-upon-able
artifact in the trade. So the roadmap's front-loading instinct is not a mistake;
the audit's claim is narrower and should not be over-read: *front-load the
irreversible, and stop front-loading everything else because it lives in the
same document as the irreversible.*
 
---
 
## I.12 — The New Jersey reordering
 
### The missing department
 
The roadmap has four departments plus the Protocol. It needs a fifth, and its
absence is the whole finding:
 
```
### 🧑 THE PATIENT — "Is anyone actually being helped, this week?"
Pilot · Acquisition · Gap reporting
Reviewer profile: the user, who has not read this document and never will.
Done means: their SOP ran on their machine, on a Monday, without you in the room.
```
 
The roadmap's organizing metaphor is *microscope before dissection*. Carried
faithfully, the metaphor has a patient in it — and §5's field mark says the real
instrument is the one that keeps acquiring **specimens**. There is currently no
department accountable for the existence of a patient.
 
### The missing tranche
 
```
# 🧑 TRANCHE −1 — THE PILOT
*Numbered below zero deliberately: §−1 of the essay is the section proving that
the ground of the whole system was never in the document — it was the person
holding it. The same is true here.*
 
- ⬜ PI.1 — name it. One human, one SOP, one workbook, one date. Written down.
- ⬜ PI.2 — the transcript. Their procedure in their words, verbatim, before any
  grammar work. This file — not `pareto.txt` — is the corpus's true north.
- ⬜ PI.3 — the concierge run. Every sentence the corpus refuses gets its VLA
  hand-written by the owner. The user sees a working automation; the project
  gets the gap log.
- ⬜ PI.4 — the gap log ranks Grammar. 🔒 Standing decision: no grammar section
  is scheduled without a sentence in PI.2 that needs it. (G-TAIL's rule,
  promoted to govern the tranche.)
- ⬜ PI.5 — kill criteria, written BEFORE the run, not after. What observation
  would mean this product should not exist?
- ⬜ PI.6 — the Monday test. Their program runs unattended, on their machine,
  with the owner unreachable. This is the beta definition (I.2) made concrete.
```
 
PI.3 is the important one. It is the only item in this audit that acquires a
specimen *this month*, and it converts the project's largest open question —
"which sentences matter" — from an argument into a measurement.
 
### The revised short answer
 
Replacing the closing section of `BETA_ROADMAP.md`:
 
> **Do next, in this order:**
> **F.1** (one sentence, largest expansion count, and the readability lever of
> Part II) → **LX.3** (hours, live correctness bug) → **F.6** (hours; also buys
> LX.2's *discipline* without LX.2's *artifact*) → **PI.1–PI.3** (name a human,
> transcribe their SOP, run it concierge) → **the six standing decisions of
> I.9** (an afternoon) → **the grammar slice the gap log names** (almost
> certainly G-FORMAT / G-STRUCT / G-ROWLOOP, appetite-boxed) → **CO.1 + IO.1 +
> DI.1/DI.2** (the three things that break on a real machine at a real company)
> → **PI.6, the Monday test.**
>
> **Then, and only then:** F.2 and F.4 at their expiry conditions, LX.1b, the
> rest of Grammar as the gap log re-ranks it, Assurance, Optimization.
>
> **Deferred with a clean conscience** (§7's real-options clause): every
> catalogue, every matrix, every second-language artifact, and the whole of
> Tranche 5. Their options are not expiring.
 
---
 
## I.13 — Concrete edits to `BETA_ROADMAP.md`
 
Small, mechanical, and each one removable if you disagree.
 
1. **Add a `CONSTRAINT:` line under the title**, re-stated each version. One
   sentence. Currently: *the Windows-box human and the absence of a user.*
2. **Define beta** in one line beside it, with a date. Bind the symbol.
3. **Add `appetite:` to every item** — `~hours` / `~days` / `~weeks`. An
   afternoon of work; it makes CD3 possible and will visibly reorder the middle.
4. **Add `specimens: N`** to each tranche header — the number of real user
   sentences or incidents motivating it. Most will read `0` today. That is the
   point of the field, and it will feel bad in exactly the productive way.
5. **Give every 🔒 an expiry condition**, in G9's form. G9 is the model: it
   names the observation that opens it.
6. **Promote F.11** to the top of Tranche 0, on TOC grounds, with its own
   reasoning quoted back at it.
7. **Split Tranche 1** into 1a (chokepoints, hours) and 1b (catalogues, weeks).
8. **Split Tranche 4** into CO-before (CO.1 policy + one refusal path) and
   CO-after (CO.2–CO.5).
9. **Swap Tranches 5 and 8.** Emitted-code speed is the pitch; compiler speed is
   invisible to everyone but the builder.
10. **Move DO.6 (SOP cookbook) and LE.6 (worked SOP templates) into the pilot
    bundle.** They are acquisition instruments, not documentation.
11. **Expand the Protocol paragraph into a numbered standing-decision register**
    with reasons and dates attached, per §3, and append the six decisions of
    I.9 to it.
12. **Add the backend-plurality decision** (I.6) as register entry #1.
 
---
 
## I.14 — How this audit could be wrong
 
Written because §7 exists and because an audit that cannot lose is also a mood.
 
- **The corpus may genuinely be the product, and corpora do not benefit from
  early users the way applications do.** A dialect shipped at 102 rules to a
  bookkeeper may simply *fail* — not teach — and a failed first contact with a
  small market is expensive in a way the essay's framework does not price. The
  counter-counter: PI.3's concierge model is designed precisely to absorb that,
  because the human fills the gaps the grammar cannot.
- **The owner is the user.** §9's fourth row — *the tool-builder is the
  tool-user* — may already be satisfied, in which case the dogfood test passes
  everywhere and this audit's central complaint weakens considerably. The
  question worth asking honestly: **has a real SOP of your own been run through
  Frazaro, start to finish, on a workbook that matters?** If yes, the specimen
  count is not zero and Part I should be re-read at half strength. If no, that
  is PI.0 and it costs one afternoon.
- **Hyrum's Law cuts the other way too.** Every argument for shipping early is
  an argument for freezing a surface early. If the grammar is still moving in
  its fundamentals, a beta buys evidence at the price of the very
  irreversibility the roadmap is trying to protect. CO.1 before contact is the
  hedge, and it is the reason it appears in the revised short answer rather than
  after it.
- **Solo psychology runs both ways.** §8 warns that the forge is warm. It is
  equally true that shipping-anxiety can dress itself up as "one more
  fortification," *and* that audits written by outsiders systematically
  under-weight how much a solo builder's momentum depends on working in a domain
  where they can tell what "done" means. If the pilot would cost you the ability
  to work at all for two months, the pilot is wrong, and no philosophy outranks
  that.
 
---
 
# PART II — READING `english.vla` WITHOUT BEING A PROGRAMMER
 
## II.0 — The seam is horizontal, not vertical
 
The question assumes the cut runs *between* `instructions.txt` and
`english.vla`. Measured, it does not. It runs **along** the vocab file,
down the middle of every rule, at the `=>`.
 
```
put {e:expr} into|in cell {r:cell}   =>   (set! (range {r}) {e})
└──────────── almost English ────────┘      └──── not English at all ────┘
```
 
The left side of `=>` is a sentence with holes in it. A bookkeeper can read
`put {e:expr} into|in cell {r:cell}` on their second look, and could probably
*write* one on their fifth. The right side is a Lisp form addressing a COM
object model. Nothing about the left side is the problem. **The file is not
Prolog-to-Lisp so much as English-with-holes to Lisp**, and only the second half
of that is unreadable.
 
Counted across the 102 rules in the file today:
 
| | Count | Share |
|---|---|---|
| Rules whose template reaches the VBA object model directly (`(. …)`, `application.`, `activesheet`, `worksheets`, …) | **61** | 60% |
| Rules whose template names an `xl*` or `vb*` constant | **11** | 11% |
| Rules whose template is already just a named verb — `(set-font-color …)`, `(clear-contents …)` | **40** | 39% |
 
That last row is the good news and it is F.1's whole thesis, already
demonstrated at 39% coverage without anyone having declared it a policy. The
worst offenders are exactly the ones you would predict:
 
```
paste values of range {a:range} into range {b:range} =>
    (begin (. (range {a}) copy)
           (. (range {b}) pastespecial :paste xlpastevalues)
           (set! application.cutcopymode false))
```
 
Three COM incantations, a magic constant, and a clipboard-state cleanup — which
is to say, three pieces of Excel folklore that the *sentence* on the left has
already named perfectly. Under F.1 this becomes:
 
```
paste values of range {a:range} into range {b:range} =>
    (paste-values (range {a}) (range {b}))
```
 
**Therefore: the single highest-leverage readability change to
`english.vla` is already item F.1 on the roadmap, and its readability
payoff is not currently listed among its benefits.** Add it. A useful,
countable, falsifiable target:
 
> **The dot count.** `english.vla` contains 61 rules that reach the object
> model. F.1 is finished when that number is 0 and the dots live in
> `prelude.vla` (F.3) instead. Publish the number; watch it fall; make a test
> that fails when it rises.
 
A phrasebook a non-programmer can read is, to a surprisingly exact
approximation, **a phrasebook with no dots in it.**
 
## II.1 — But the question was about tongs, so: where does the cut actually go?
 
F.1 buys legibility, not authorship. `(paste-values (range {a}) (range {b}))` is
readable; it is still not something a bookkeeper *writes*. The Dedekind-cut
question — what lives strictly between `instructions.txt` and `english.vla`? —
is the right question, and the essay's §−1.5 gives the shape of the answer
before it gives the answer: **the thing standing in the cut is a person**, and
the interval is dense, so the correct move is not to name one intermediate
notation but to open the interval and let it fill.
 
Two candidate cuts. They are complementary, and the first one is nearly built
already.
 
### Cut A — the definitional cut (the one you already have)
 
`instructions.txt` already contains this:
 
```
To stamp, with row-number of 1 and value of "ok":
  Put value into column F row row-number.
```
 
Read it again with the phrasebook in mind. That is **a vocabulary rule written
entirely in English, with named parameters, defaults, and a body in the language
itself.** It defines a new thing to say. It ships in the corpus. It works today.
 
The distance from that to a phrasebook entry is smaller than it looks, and it is
exactly three properties:
 
| | `To stamp …` today | A `.vocab` rule |
|---|---|---|
| What it defines | a **word** with named arguments (`Stamp with row-number of 2`) | a **sentence shape** — word order, prepositions, alternations |
| Where it lives | inside one program | in a phrasebook: persistent, shared, versioned |
| When it happens | runtime (a VBA `Sub`) | compile time (a macro expansion) |
 
Close two of those three and the tongs exist. The proposal:
 
> **Let `To` take a phrase pattern instead of a name, and let a phrasebook be a
> program that contains nothing but `To` definitions.**
 
```
To put {amount} into cell {where}:
  Set the value of cell {where} to {amount}.
 
To highlight the row of cell {c}:
  Make the row of cell {c} yellow.
  Make the row of cell {c} bold.
 
To close out {month}:
  Go to sheet Summary.
  Copy values of range A1:D50 into tab {month}.
  Export this sheet as pdf ...
```
 
Every line of every body is a sentence the user has **already learned to write**
by using the product. Nothing new is taught. No braces-with-colons, no `=>`, no
parentheses, no dots, no `xl` constants, no second notation of any kind. The
author's entire mental model is: *"I am teaching it a new way of saying things,
using the ways it already knows."*
 
And this is precisely the essay's §−1 in product form. Hart's memo observed that
a finished interpreter is an invitation, and that **every user of a system is a
potential author of its next extension**. `english.vla` becomes the
kernel — the small, programmer-maintained set of rules that bottom out in VLA —
and everything above it is written by people who never leave English. The
document grows downward toward its reader, one definition at a time.
 
What this costs, honestly:
 
- **Slot typing.** How does the machine know `{where}` is a cell? Two options,
  and the good one is free: *infer it from the body*. `{where}` is passed where
  a `:cell` is expected, so it is a cell, and the Check-time shape refusals (G2)
  keep working unchanged. Fall back to an English gloss when inference is
  ambiguous — `{where: a cell}` — which reads fine and needs no punctuation
  training.
- **Conflict detection becomes mandatory.** User-authored sentence shapes *will*
  collide. This is **F.4's real justification** — not "340 rules" but "strangers
  writing rules" — and it is the expiry condition proposed in I.7. Same for
  **GO.1** (precedence: base corpus vs org phrasebook vs this user's
  definitions) and for the refusal that must say *which layer* refused.
- **Expansion depth must be bounded.** Definitional layers stack, and a
  definition can name a definition. The essay hands you the discipline for free
  in §−1.5's appendix: the expander is `macroexpand-1`, single-stepped, because
  the fixpoint does not halt. Bound the depth, report it in words when it is
  exceeded, and refuse actual infinity while keeping the potential kind. *(In
  practice: a definition may use definitions, and a cycle is a worded refusal
  naming both sentences.)*
- **F.1 is a prerequisite, not an optimization.** If templates still bottom out
  in dot-forms, an English body compiles into dot-forms and the middle layer
  leaks the moment anything goes wrong. The ABI is what makes the English bodies
  honest.
- **F.2 is a prerequisite too.** English-emitting-forms is what lets a body be
  *compiled* into a template rather than *string-substituted* into one.
 
Which yields a genuinely useful reordering fact for Part I: **F.1, F.2 and F.4
are not merely Grammar's gates. They are the gates on the phrasebook ever being
authorable by a non-programmer** — which is a far larger prize than "rules get
shorter," and it belongs in their *why now* lines.
 
### Cut B — the extensional cut (show, don't specify)
 
The second tongs, cheaper and independently useful. Today an author writes the
rule **and** its proofs, in two notations, saying the same thing twice:
 
```
put formula {f:text} into|in cell {r:cell} =>
    (set-formula (range {r}) {f})
test: Put formula "=B2*2" into cell B3. =>
    (set-formula (range "b3") "=B2*2")
```
 
Invert it. Let the author supply only what they already know how to write —
**sentences** — and let the machine propose the rule:
 
```
Author types:   Put formula "=B2*2" into cell B3.
                Put formula "=A1+1" in cell D7.
 
Frazaro replies: I think you mean:
                   Put formula <some text> into (or in) cell <a cell>.
                 I'll write that as: set-formula.
                 Say another example, or say "yes".
```
 
The pattern, the alternation, the slot types, and the two `test:` lines are then
**machine output**, not human input. The metasyntax that makes the file look
like Prolog stops being something anyone types. It survives in the file as a
compiled artifact — which is the correct fate for a notation whose audience is a
parser.
 
This wants three things the roadmap already contains, which is a good sign:
**F.4** (the near-miss/conflict math), **LE.2** ("did you mean" — the same
matcher pointed the other way), and **U1's Rehearse**. It is the same machinery
aimed at authorship instead of diagnosis.
 
### The mirror both cuts need: show a rule's extension, not its intension
 
The deepest readability problem with any rule notation is that a rule is an
**intension** — a description of a set — while human beings check their
understanding against the **extension**, the members. Nobody reads
`sort range {r:range} by column {k:cell} [ascending]` and *knows* whether
`Sort range A1:C50 by column B1.` is accepted. They guess.
 
So render it:
 
```
sort range {r} by column {k} [ascending]
    accepts, among others:
      Sort range A1:C50 by column B1.
      Sort range A1:C50 by column B1 ascending.
      Sort range "Q1 Data" by column B1.
    refuses:
      Sort range banana by column B1.   ("banana" is not a range)
```
 
The sampler is small — expand the alternations and optionals, fill slots from a
fixture table — and it pays into four roadmap items at once: **LE.1** (the
palette *is* this, aggregated), **AS.4** (round-trip property tests), **G4** (the
`fail:` proofs get generated rather than hand-written), and **U1's Rehearse**.
It also makes the second direction obvious and valuable: **render a program back
into English from its forms**, so a rule can be checked by round-trip and a
workbook can be read by someone who did not write it. That capability turns out
to matter enormously in Part III, so it is worth building for this reason and
collecting the other one for free.
 
## II.2 — Cheap wins for the file exactly as it stands
 
If neither cut ships this year, the file itself can be improved in an afternoon.
The principle: **five metasyntaxes is four too many**, and every one that
survives should read as English punctuation rather than as regex.
 
| Today | Proposed | Why |
|---|---|---|
| `=>` | `means` | A word, not an arrow. Reads aloud. |
| `test:` | `example:` | "Test" says *this file is for engineers*. "Example" says *this file is for readers*. Identical semantics. |
| `fail:` | `never:` | Same. And `never:` states the intent, not the mechanism. |
| `{r:cell}` | `{a cell}` / `{some text}` / `{a column}` | Type names as English nouns. Machine-readable, human-obvious, no colon dialect to learn. |
| `into|in` | `into (or in)` | The pipe is the single most alien character in the file. |
| `[ascending]` | keep | Square brackets already mean *optional* in every dictionary and grammar book in English. This one was right. |
| `center/ed` | keep, but document as "the slash means both spellings" in one line, not four | |
| `# comments` carrying rule rationale | a machine-visible `note:` line | Today the reasoning is in comments, so **LE.1's palette cannot show it**. That is a real loss: the best prose in the file is invisible to the product. |
| `# ---- 56-68: the expansion pack ----` | `# ---- Finding and fixing text ----` | The numeric ranges are archaeology. They date the rules; they do not organize them. |
 
Two structural moves in the same spirit:
 
- **Split the kernel out.** `kernel.vocab` holds the `macro:` definitions, the
  `function:` mappings, and any rule that must touch a dot-form.
  `english.vla` keeps the rest and contains no dots at all. The split is
  the readability policy made physical: *if you are editing a file with
  parentheses in it, you are in the wrong file, and that is now obvious from the
  filename.*
- **Publish the verb catalogue at the top of the phrasebook.** After F.1, the
  right-hand side of every rule draws from a closed set of named verbs. Print
  that set — with the one-line docstrings that already exist on the `defmacro`s
  — as the phrasebook's own glossary. An author then has a vocabulary to
  compose from rather than an object model to explore.
 
## II.3 — The tongs nobody has proposed yet: the phrasebook is a workbook
 
One more, offered because it is on-brand to the point of being slightly
embarrassing that it is not already true.
 
Frazaro's users live in spreadsheets. Its authoring surface is a text file with
five metasyntaxes, edited in Notepad. **Author the phrasebook in a workbook.**
 
| What you say | For example | What it does | Proven |
|---|---|---|---|
| Put {some text} into cell {a cell} | Put "Q1" into cell A1. | write a value into a cell | ✅ |
| Highlight the row of cell {a cell} | Highlight the row of cell B7. | *(defined below, in English)* | ✅ |
 
`english.vla` becomes a **build artifact** generated from that sheet —
still the file of record, still diffable, still the thing CI checks, but no
longer the thing a human edits. Three consequences fall out immediately: the
`test:` column is filled in by the same act that writes the rule; the "Proven"
column is live, because the loader already refuses to load a vocabulary whose
tests fail; and **LE.1's sentence palette is the read-only view of the identical
table** — one artifact, two directions, and the project dogfoods its own claim
that a spreadsheet is a reasonable place to keep serious work.
 
## II.4 — Summary: where to put the cut
 
The interval between `instructions.txt` and `english.vla` is dense, and it fills
in this order:
 
```
instructions.txt                        a program: sentences, no holes
  │
  ├─ To {phrase}: …                a definition: sentences with holes,
  │                                body in English            ← Cut A
  ├─ authored by example           sentences in, pattern out  ← Cut B
  │
  ├─ english.vla (post-F.1)  patterns → named verbs, no dots
  │
  └─ kernel.vocab / prelude.vla    named verbs → VLA → VBA
```
 
Each layer is written in the layer above's output and read by one more person
than the layer below it. The Prolog-to-Lisp notation does not disappear — it
descends, becoming a compilation target rather than a human surface, which is
what happened to assembly and is the normal fate of a notation that was always
addressed to a machine.
 
And the property worth keeping from §−1.5: **you never reach the bottom of the
interval, and that is the design succeeding, not failing.** There will always be
a kernel that a programmer maintains. The measure of progress is not that the
kernel vanishes; it is that the boundary keeps moving downward, one bisection at
a time, and that each move is made by somebody who was, the week before, a user.
 
---
 
# PART III — THE HORIZON
 
*Assume everything in `BETA_ROADMAP.md` is implemented. Not "shipped and adopted" —
just built. What becomes possible that is not possible now, and what should the
freed hours be spent on?*
 
*A note on how this part is written: each horizon names the roadmap items it
actually depends on, so that the daydreaming stays load-bearing. A horizon with
no dependency line is a wish. A horizon with one is a plan you have not
scheduled yet.*
 
## III.0 — The premise: a freed hour is a budget, not a saving
 
The framing worth abandoning first is "saving time." Time saved from an SOP does
not return to the person as leisure; in every organization that has ever
automated anything, it returns as **capacity**, and capacity is spent on whatever
the tools make cheapest. The interesting question is therefore not *how much time
does Frazaro save* but **what does Frazaro make cheap that used to be
impossible**, because that is what the recovered hours will be spent on whether
anyone plans it or not.
 
Three things become cheap that were previously not merely expensive but
unavailable at any price to the person who needed them:
 
1. **Writing down a procedure in a form that runs.** Not describing it. Not
   filing a ticket about it. Writing it.
2. **Reading someone else's automation.** Today this requires a programmer. It
   is the reason most spreadsheet automation is unauditable and therefore
   untrusted and therefore not done.
3. **Changing a procedure and seeing what changed.** A one-line English diff a
   manager can approve is a different object from a VBA patch.
 
Everything below is a consequence of one of those three.
 
## III.1 — The near horizon (things a beta user would notice within a month)
 
### The SOP stops being a document *about* a procedure and becomes the procedure
 
Every regulated organization on earth maintains two artifacts that are supposed
to correspond and never do: the written SOP, and whatever people actually click.
Drift between them is a permanent, unfixable, universally-tolerated cost, and it
is the reason audits are expensive.
 
Frazaro collapses them. The procedure is written in the language the procedure
was already written in, and it *runs*. The compliance artifact and the
executable are the same file. The reviewer who signs it is reading the thing
that ran, not a description of the thing that ran.
 
That is a stronger claim than "faster spreadsheets," and it is worth noticing
that the roadmap already contains it without saying so: it is what LE.6 (worked
SOP templates) and DO.6 (the SOP cookbook) are *for*.
 
> *Depends on:* Grammar breadth, LE.6, DO.6, and — critically — the
> English-rendering direction of Part II.1, so that a program written by one
> person is readable by the person who signs it.
 
### Procedures acquire version control, and the diffs are in English
 
```
- Copy values of range A1:D50 into tab Archive.
+ Copy values of range A1:F50 into tab Archive.
```
 
A manager can approve that. A manager cannot approve a VBA diff, which is why
process changes in most organizations are approved on the basis of a verbal
summary of a change nobody in the approval chain has read.
 
This is a genuinely new object: **a procedural changelog that non-technical
reviewers can actually review.** It is nearly free once programs are text and
sentences are stable — which is to say, once CO.1 and CO.2 exist.
 
> *Depends on:* CO.1, CO.2, and the surface stability they promise.
 
### The spreadsheet starts arguing with you
 
**V.1 — `verify:` rows** is filed in Tranche 9 as a UI item, and the roadmap's
own note is right that it may be "the most Frazaro-shaped feature on this entire
roadmap." It is undervalued even by that description.
 
The famous spreadsheet catastrophes — the ones with names, the ones that moved
national policy and cost banks billions — were not caused by people who could not
write formulas. They were caused by spreadsheets that had **no way to say what
they assumed.** A range that silently stopped covering all the rows. A column
that quietly became text. There was no place to write *"the total of the
categories must equal the grand total"* in a way the file itself would check.
 
`verify:` is that place, in English, checked on every run:
 
```
Verify: total of column D equals grand total.
Verify: every row of range A2:A500 is not empty.
Verify: count of found-items is at least 1.
```
 
This is the project's own doctrine — *proof accompanies capability*, predictions
before runs — handed to the user. It is also, commercially, the easiest thing in
this entire document to explain to a CFO.
 
> *Depends on:* V.1, plus a small grammar section for assertion sentences.
 
### Pointing Frazaro at the existing mess
 
The compiler runs English → VLA → VBA. Nothing prevents a second front end that
reads *recorded macros* and emits sentences. Most of the world's spreadsheet
automation is a recorded macro full of `.Select` that nobody dares touch because
the author left in 2019.
 
"Show me what this macro does, in English" is a smaller feature than it sounds
(the object-model vocabulary is already mapped in `english.vla`, in the
direction you need) and it is an extraordinary acquisition instrument: it meets
users at the exact moment they are most afraid, and it converts their existing
liability into your input format.
 
> *Depends on:* the rendering direction of Part II.1 plus a VBA reader. PF.3's
> "never emit `.Select`" is the same knowledge pointed the other way.
 
## III.2 — The middle horizon (the flows that change what a job is)
 
### The house dialect, and a new role: the phrasebook author
 
Once Part II's Cut A exists, an organization accretes a vocabulary the way it
accretes a chart of accounts. "Close out the month." "Run the variance pack."
"Reconcile the intercompany file." Each is a verb somebody defined, in English,
in terms of verbs somebody else defined.
 
This is the moat the roadmap already identified — *"nobody can clone a curated
dialect plus a corpus of tested phrasings plus the judgment that shaped them"* —
but it is more interesting as a **social** object than a defensive one. It
creates a role that does not currently exist: the person who tends the
organization's procedural vocabulary. Not IT. Not the analyst. Something closer
to a librarian or a standards editor — and, notably, a role that the best
process-knowledgeable clerk in the department is *already* qualified for.
 
The deep organizational win: **the automation is written by the person who knows
the process.** The requirements-translation loop, where a domain expert explains
a procedure to a developer who misunderstands it in a novel way, is where most
enterprise automation dies. Frazaro's mission statement is aimed at exactly that
loop, and this is the shape of its removal.
 
> *Depends on:* GO.1 (precedence), F.4 (collision), Part II Cut A, DO.4.
 
### You can talk to it
 
A constrained grammar with a published sentence catalogue is the thing speech
recognition has always wanted and almost never gets. General code is
undictatable — punctuation, casing, identifiers. Frazaro sentences are English
with a known vocabulary and a known shape, which means recognition can be
*constrained by the grammar itself*: the recognizer only has to distinguish
between sentences the corpus accepts.
 
That opens work that has never been open:
 
- Automation authored on a phone, on a train, by someone who has never opened a
  code editor.
- **Accessibility as a first-class capability rather than a compliance item.**
  Tranche 13 is thin and honest about it, but a language dictatable end-to-end is
  a genuinely important thing for users with motor impairments, and there is
  nothing else in this category.
- Hands-busy work: the lab bench, the warehouse floor, the shop counter, where
  the person who knows the procedure has never been within reach of a keyboard
  while performing it.
 
> *Depends on:* LE.1 (the palette is the recognizer's grammar), LE.3
> (autocomplete is the same constraint), LX.8 (refusals that teach, because
> dictation misfires constantly).
 
### The program with no source language
 
This is the mission's real payoff and it is bigger than "Spanish users get a
Spanish product."
 
Once **F.2** makes English emit *forms* rather than strings, and once the
rendering direction exists, a program is no longer stored in a language. It is
stored as forms and **rendered in the reader's language on arrival.** The clerk
in Guadalajara writes it in Spanish. The manager in Osaka opens the same workbook
and reads it in Japanese. Neither of them is reading a translation of a comment;
they are reading the program, rendered.
 
Nothing in mainstream programming works this way, and the reason is instructive:
general-purpose languages have unbounded identifier vocabularies, so their source
cannot be re-rendered. A **curated dialect can be**, because its vocabulary is
finite, tested, and owned. The constraint that makes Frazaro learnable is
precisely the constraint that makes it translatable. That is not a coincidence;
it is the same property twice.
 
Consequences worth savouring: cross-border process review without translation
services; a multinational's SOPs maintained once and read everywhere; a
contributor in one language improving a phrasebook that speakers of another
language immediately benefit from.
 
> *Depends on:* all of Tranche 1 (this is the mission), F.2, LX.4, LX.10 as the
> proof, plus rendering. This is the horizon that justifies the elevation of
> Linguistics — and it is worth noting that it justifies it *as a destination*,
> not as a prerequisite, which is Part I's entire quarrel with the tranche's
> position rather than its content.
 
### The honest AI bridge
 
The obvious 2026 question — *why not just ask a model to write the VBA?* — has an
answer, and Frazaro is unusually well-shaped to be it.
 
An LLM writing VBA produces code that the person who asked for it **cannot
read**, cannot verify, and must either trust or discard. An LLM writing *Frazaro
sentences* produces something entirely different:
 
- Its output is in a language with a **published, finite grammar**, so
  malformed output is refused mechanically rather than discovered in production.
- Its output is **readable by the user who requested it** — the whole point.
- Its output is **testable**, because the corpus protocol already demands that
  every capability arrive with its proof.
- Its errors are **refusals with directions** (LX.8), not runtime explosions.
 
Frazaro becomes the safe target language for AI-generated spreadsheet
automation: the layer that makes a model's output auditable by the person who
has to sign it. That is a strategically excellent place to be standing, and it
inverts the usual anxiety — the model does not replace the language; the
language is what makes the model usable in a regulated spreadsheet at all.
 
> *Depends on:* LX.7 (a spec a model can be constrained to), LE.1 (the catalogue
> as the constraint), F.4, and the refusal quality of LX.2/LX.8.
 
## III.3 — The far horizon (the parts worth daydreaming about)
 
### Every user becomes a macro author, and the language grows toward them
 
The stairwell of §−1.5, in product form. Hart's memo observed that a finished
interpreter is an invitation and that every user is a potential author of the
next extension. Frazaro's phrasebook space has the same density property: between
the language and any user there is room for one more definition, and the one who
writes it is always somebody who was, last week, only reading.
 
What that looks like in the field: a bookkeeper defines *"reconcile the card
file"* on a Tuesday. Her colleague uses it without knowing it was defined. Six
months later it is in the department phrasebook. Two years later she has left,
and the verb is still running every Tuesday, expanding at call sites she never
imagined, in the hands of people who never met her. That is a metametamacro with
an expansion count, authored by someone who would not describe herself as
technical, and it is the most durable thing she made.
 
### A procedural commons
 
Phrasebooks are shareable in a way software libraries are not, because they are
*readable by the people who need them*. A parish-accounting phrasebook. A
school-timetabling phrasebook. A clinical-trial data-entry phrasebook. A
small-manufacturer inventory phrasebook. Each is a few hundred sentences, each
encodes a trade's actual vocabulary, each maintained by practitioners of that
trade rather than by engineers.
 
The interesting artifact at the end of that road is a **commons of procedure** —
the accumulated, tested, executable working knowledge of trades that have never
had a way to write their knowledge down in runnable form. Nonprofits and small
organizations are the obvious beneficiaries: they have the procedures and have
never had the programmers.
 
> *This is what GO.1, GO.3 and DO.4 are for. Thin today, correctly.*
 
### The corpus as a linguistic artifact
 
Somewhere past ten thousand users there exists something no one currently has: a
**tested, curated, empirically-grounded register of business English**, with
every accepted phrasing carrying its proof and its history of why it was
accepted. Not a style guide — a corpus with executable semantics and a change
log of adjudications.
 
That is a research asset in its own right, of interest to linguists, to
controlled-natural-language work, to anyone studying how a constrained language
evolves under real use. And the ledgers already being written — the ones that
record *why* "Store" beat "Remember under" — are its provenance. The project is
accidentally producing a primary source, and it is producing it well.
 
### Teaching, without the syntax tax
 
Every attempt to teach programming spends its first weeks on punctuation. A
language whose refusals are sentences, whose vocabulary is browsable, and whose
host is the application students will actually use in their working lives is a
better first programming environment than most things built specifically to be
one. Loops, conditionals, variables, functions, and assertions — all present in
`instructions.txt` today, all sayable without a single character of ceremony.
 
The refusal is the tutor. That is not a metaphor: LX.8's style guide (*name the
problem, teach the fix, never blame*) is a pedagogy, written down as an interface
policy, three tranches away from being the most-read teaching prose in the
product.
 
### Excel is only the first host
 
VLA is the target-independent middle. The same grammar over documents, over
mail, over a database, over the browser. The mission sentence — *whatever
language a person prefers to think in should be the same language their top-level
source code is written in* — says nothing about spreadsheets. Spreadsheets are
where the people are, and where the pain is, and where the proof will come from.
 
> *This is the horizon Part I.6 asks you to protect with one sentence today.*
 
## III.4 — What to keep intact so the horizon stays reachable
 
The horizons above are not equally fragile. Four of them share a small set of
load-bearing properties, and those properties are cheap now and unbuyable later.
This is the practical residue of Part III, and it agrees with Part I:
 
| Property | Guards which horizon | Cost today |
|---|---|---|
| **Programs are forms, not text** (F.2) | The no-source-language program; the AI bridge; rendering | Weeks now, a rewrite later |
| **Templates never touch the backend** (F.1) | Backend plurality; other hosts; phrasebook readability | One sentence |
| **The vocabulary is finite, published, and owned** (LX.7, LE.1) | Dictation; the AI bridge; translation; teaching | Discipline, not code |
| **Rendering exists in both directions** | Translation; legacy import; readable diffs; the audit story | Not currently on the roadmap at all |
 
That last row is the one finding in Part III that is genuinely new: **the
English-rendering direction — forms back to sentences — is a dependency of four
separate horizons and appears nowhere in `BETA_ROADMAP.md`.** It is partially implied
by AS.4's round-trip property test, which is filed as an assurance item rather
than as a capability. It deserves its own line, in Grammar or in Interface, with
its reasons attached.
 
## III.5 — The daydream, stated plainly
 
Here is the picture worth keeping on the wall for the hard parts.
 
A woman who has worked in the same finance office for eleven years opens a
workbook on a Tuesday morning. She does not open a code editor, because there
isn't one. She types six sentences in her own language — the language she has
been using in emails about this exact procedure for a decade — and the month-end
close runs. It runs on her machine, in her locale, on her data, with the numbers
her department actually cares about, and it checks itself: three `verify:` lines
she wrote herself, which will refuse in words on the morning that someone
upstream changes a column.
 
The procedure she wrote is the procedure the auditor reads. There is no second
document. There is no drift.
 
Somewhere in her six sentences is a verb her colleague defined last spring. She
has never seen its definition and never will. It expands anyway, correctly, the
same way every time.
 
And the hour she used to spend on this — she does not get it back as leisure.
She spends it teaching the machine the *next* procedure, the one nobody has ever
automated because it was never worth a developer's week. Which is the whole
thing, really: not that the clicking stops, but that the ceiling on what is worth
automating drops through the floor, and the people who know how the work is
actually done get to do the automating.
 
That is the horizon. It is reachable from where this project already stands, and
the shortest path to it starts — this is Part I's entire argument, arriving at
Part III's conclusion by a different road — with **one real person, one real
SOP, this month.**