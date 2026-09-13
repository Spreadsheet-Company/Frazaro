# THE ID REGISTRY

*SD-9 ("IDs are never reused across documents, and a retired ID is never
re-minted") made mechanical, per F.12. The sibling of REBUILD.md's R11 +
Appendix D, which does the identical job one level down — R11 is names for
`.bas` modules; this is names for roadmap items. Same failure class, same
fix: a rule, one audit against it, and a check that runs again forever.*

---

## The rule

Two IDs collide if they are the same after **normalizing**: strip every `.`
and `-`, uppercase what remains. `F1` and `F.1` normalize to the same string
and are therefore the same identity question, whatever the punctuation
suggests — this is the exact shape of SD-9's founding incident (Alpha 1's
interpreter mode, `F1`, collided with this file's grammar/emitter ABI,
`F.1`, during promotion into strategy, and the item with the homograph is the
one that quietly failed to arrive).

1. **One namespace.** An ID is minted once, in one document, and means one
   thing everywhere it is then cited.
2. **A retired ID is never re-minted.** Not reused for an unrelated item, not
   reissued after a rename — retired means permanently out of circulation.
   The register below is the retired list; it is hand-maintained because
   retirement is a judgment call, not something a script can infer.
3. **New IDs use the family's established form.** Most families are dotted
   numeric (`F.13` follows `F.12`); a few are word-suffixed (`G-STRUCT`,
   `P-PROBE`, `L-TIER2`) and a new member of those families follows the same
   pattern, not a numeric one. Never introduce a bare numeric form (`F13`)
   into a family that has already gone dotted — that is exactly how the
   founding incident happened.
4. **Check before minting, and check at version-close.** Run
   [`tools/check_id_registry.ps1`](../tools/check_id_registry.ps1). It prints
   the next free ID per family and fails (non-zero exit) if a governed
   collision, a duplicate definition, or a re-minted retired ID exists.

## Governed set, and why it stops there

**Governed:** `docs/BETA_ROADMAP.md` + `docs/ALPHA*_ROADMAP.md` (today, just
`ALPHA6_ROADMAP.md`; the glob picks up `ALPHA7_ROADMAP.md` unmodified when
this version closes). This is the promotion path SD-9 describes in its own
telling — a ledger mints an item, the strategy file carries it forward or
promotes it — so it is the only place a new ID gets minted, and the only
place this script treats a collision as an error.

**Advisory, not governed:** `docs/REBUILD.md`, `docs/LESSONS.md`,
`docs/AUDIT.md`, `docs/PROJECT_BRIEF.md`. Each keeps its own internal
numbering for its own purpose — REBUILD.md's `R1`–`R11` are naming/lint
rules and its plate/layer steps are a target-architecture sketch, not
backlog items; LESSONS.md and AUDIT.md number findings from sessions that
predate this file. Renumbering any of them is out of scope for an `~hours`
item and is not what SD-9 asks for — but BETA_ROADMAP.md's own prose already
cites REBUILD.md IDs inline (`R7`, `R9`, `R10`, `S3.1`, ...) in the same
breath as its own items, so a token free in the governed set can still read
as ambiguous next to an advisory one. The script reports that overlap so it
is a documented choice, the same discipline Appendix D applied to module
names: audit once, record the reasons, don't re-litigate it by hand every
time.

## Current state — measured 2026-08-28 (second refresh, same day)

Ran clean: **0 governed collisions, 0 duplicate definitions, 0 retired IDs
re-minted.** SD-9's discipline has held across every item since the founding
incident. *Refreshed again, same day, after `SIG.0`–`SIG.5` plus `F.15`,
`LX.12`, `IN.14`, and `AS.9` were minted* — all six checked clean before
landing in BETA_ROADMAP.md. This mint again coexisted with a concurrent
session's own in-flight, uncommitted work (a G3/`mSigOwner` follow-up inside
`F.4`'s own entry, not itself a new ID mint) — the two edit streams were kept
apart at the git-hunk level too, not just the ID level, so the concurrent
session's own paragraph stayed unstaged for its own commit.

High-water mark per prefix family (next free ID), from the script's own
output:

| Prefix | Highest today | Next free |
|---|---|---|
| AC | AC.4 | AC.5 |
| AS | AS.9 | AS.10 |
| CO | CO.5 | CO.6 |
| DI | DI.5 | DI.6 |
| DO | DO.6 | DO.7 |
| DR | DR2 (bare) | DR3 |
| EN | EN.8 | EN.9 |
| F | F.15 | F.16 |
| G (bare numeric sub-family) | G11r | G12 |
| GO | GO.5 | GO.6 |
| IN | IN.14 | IN.15 |
| IO | IO.6 | IO.7 |
| L | L17 (bare) | L18 |
| LE | LE.10 | LE.11 |
| LX | LX.12 | LX.13 |
| PF | PF.6 | PF.7 |
| PI | PI.7 | PI.8 |
| S | S6 (bare) | S7 |
| SEC | SEC.6 | SEC.7 |
| SIG | SIG.8 | SIG.9 (SIG.6, SIG.7 reserved — see below) |
| U | U.18 | U.19 |
| V | V8 (bare) | V9 |

Word-suffix families in use (next member follows the same pattern, not a
number): `G-` (FILES/FORMAT/FORMULA/PATH/PIVOT/RENDER/ROWLOOP/SORTFILTER/
STRUCT/TABLES/TABS/TAIL), `L-` (FILE-HELPERS/PIVOT-HELPERS/SHEET-HELPERS/
TIER2/TIER3), `P-` (BULK/DICT/INCLUDE/KW/NTH/PRED/PRELUDE/PROBE/PROF/QUOTE/
SUBDOC/TOK/TRYELSE/WHENLET), `SD-` (numeric, governed separately as the
standing-decision register — SD-16 is next, after SD-15 (dynamic-dispatch
capability doctrine — no reach beyond the native tier without a declared,
granted capability; default deny), minted while scoping `SEC.*`).

**`SEC` is a wholly new family, minted the same day as this second refresh's
own `SIG` mint, both declared in the prefix inventory at the moment of
minting** (`BETA_ROADMAP.md`'s "HOW THIS ROADMAP IS ORDERED" section) rather
than found missing by a later audit — the L./V./DR gap's own lesson, applied
instead of repeated twice. Six items, `SEC.0`–`SEC.6`, under a new sixth
department, 🛡 THE ADVERSARY. `SD-15` (dynamic dispatch beyond the native
tier requires a declared capability, default deny) is `SEC.1`'s decision
half, appended to the standing-decision register in the same pass.

**`SIG` is this refresh's own new family: six items, `SIG.0`–`SIG.5`, under
a new seventh department, 🖋 THE SIGNATORY** ("what has to be true before
someone with budget authority can say yes") — the second department
`docs/CONSULTANT.md`'s own addendum proposed alongside the Adversary, in the
same breath. The remaining, non-department-shaped findings from that same
assessment landed as new items under their own existing families instead of
a new one: `F.15` (Fortifications — the two hardcoded module manifests),
`LX.12` (Linguistics — UI chrome outside `LX.2`'s catalogue), `IN.14` (the
interpreter — cancel/`DoEvents` in the execution loop), `AS.9` (Assurance —
coverage against real external text), plus two hydrations of already-minted
items with no new ID (`IO.4`, `G-PATH`) and two `U.*` mints (`U.17`, `U.18`,
already reflected in the table above, for snapshot-before-run and a
user-facing per-run log).

**`SIG.8` minted 2026-09-12 (owner), and `SIG.6`/`SIG.7` deliberately
skipped rather than spent.** The new item is the signing certificate as a
separately published artifact, so an IT department can deploy Frazaro as a
trusted publisher — scoped while closing `SIG.1`, which found that the
certificate exists but is distributed nowhere. It did **not** take the
family's nominal next free slot: `SIG.6` (one practicing auditor's verbatim
reaction, proposed by `ADVOCATUS.md`) is already cited as if real in
`VENTURE.md` §6.3 and §7 and in `MARKETING.md` §3.10, and `SIG.7` (the win
condition with an expiry, proposed by `VIABILITY.md`) is named in its own
review — both recorded as unminted candidates in `CONTEMPLATIONS.md`'s
collection table of 2026-09-09. Minting either number for a different item
would have resolved an outward-facing unbound symbol to the wrong
definition, which is the founding incident's own failure mode (SD-9) one
level up. They stay reserved for the items that already bear their names;
whoever collects them mints them as cited or declines them in writing.

**IN advanced between the 2026-08-18 and 2026-08-19 snapshots (recorded
then, left as-is here per this file's own nothing-is-pruned convention):**
`IN.9` → `IN.10` (user-procedure call/return, BETA_ROADMAP.md,
adjudication/scoping only — see the item's own "buy the decision now, defer
the artifact" framing). Several other families (`F`, `IN`, `G-`) advanced
further between the 2026-08-19 and 2026-08-28 snapshots via a different,
concurrent session's own grammar/interpreter work, not itemized here in
either refresh — each refresh's own scope is its own mint; the table itself
already reflects the current high-water marks regardless of who minted them.

**A second finding, same shape as the L./V./DR gap the founding audit
closed, surfaced by this refresh rather than by that one:** the `L`, `V`,
and `S` rows above jumped far past their prior snapshot (`L.11`→`L17`,
`V.1`→`V8`, and `S` is a wholly new row at `S6`) with no new item minted in
either quantity — because the **governed set includes every historical
`ALPHAn_ROADMAP.md` the glob matches (1 through 6 today), not only the
currently-open ledger**, and closed ledgers `ALPHA4_ROADMAP.md`/
`ALPHA5_ROADMAP.md` already carried bare-numeric tokens (`S6`, `L14`-`L17`,
`V2`-`V8`) in that shape before this file was ever written. Those are each
ledger's own internal sequencing, the same kind of thing REBUILD.md's `R1`-
`R11` is for that document — not draws against the rotating `L.`/`V.`
*dotted* families this file's own prefix inventory declares — but the
script's prefix grouping keys on the leading letters only, so a bare `V8`
and a dotted `V.1` land in the same high-water row and the bare form wins
on magnitude. **Not fixed by this pass** (the script's governed-set scope
and prefix-grouping logic are a design question, not a snapshot-staleness
one, and out of scope for a table refresh) — recorded here so the next
person minting a `V.*` or `L.*` ID reads `V9`/`L18` as noise from old,
closed ledgers, not as this file's own count, and reads the dotted form
(`V.2`, `L.12`) as the one actually free in the *live* rotating family.

**Three prefixes the founding audit found live but undeclared:** `L.` and
`V.` were already minted (`L.11`, `V.1`) and cited throughout BETA_ROADMAP.md,
but missing from this file's own prefix inventory (the "ordering" section,
now fixed). `DR` (`DR1`, `DR2`, both on `ALPHA6_ROADMAP.md`, DR2 currently
⛔ parked) was never declared anywhere. None of the three were ever
actually ambiguous — that was a documentation gap the audit closed, not a
collision the audit found.

**Spot-checked, not exhaustively reviewed:** the advisory-overlap output is
large (most of BETA_ROADMAP.md's own items are, unsurprisingly, *also* mentioned
in REBUILD.md and AUDIT.md's prose — that is normal cross-referencing, not a
collision). One family was checked by hand: the bare `G6`/`G7`/`G8`/`G9`/
`G11r` tokens that PROJECT_BRIEF.md and both ledgers share are the *same*
grammar items in both places (PROJECT_BRIEF.md line 53: "bucket 3 (Grammar)
in ledger order G.1 G11r → G.2 G7 → ..." — its own `G.1`–`G.5` is a
*sequencing* label over the same bare IDs, not a second meaning for them).
The remaining advisory-overlap entries were not individually verified; that
is future vigilance the script now makes cheap, not a claim this pass
makes about every row.

## Retired — never re-mint

| ID | Reason |
|---|---|
| `F1` (bare) | Alpha 1's interpreter-mode item. Collided with this file's `F.1` (the grammar/emitter ABI) during promotion into strategy — SD-9's founding incident, and the reason SD-9 and this registry exist. `IN.*` is where that work lives now; REBUILD.md still glosses `IN.*` as "Alpha 1's `F1`" for historical continuity, which is a citation of the retired form, not a re-mint of it. |

## What this pass did not do

Did not renumber REBUILD.md, LESSONS.md, AUDIT.md, or PROJECT_BRIEF.md —
out of scope, per the governed/advisory split above. Did not build fuzzy
meaning-comparison between documents; the governed-collision check is exact
and narrow by design (same normalized token, different punctuation, in the
one namespace that mints IDs), which is precisely SD-9's failure shape and
nothing broader was promised. Did not wire the script into an automated gate
— there is no CI runner in this project (`AS.6`'s "in CI" already means "in
`VlaSelfTest`," run by the human at the Windows box); running this script is
now a version-close step the same way running `VlaSelfTest` is.
