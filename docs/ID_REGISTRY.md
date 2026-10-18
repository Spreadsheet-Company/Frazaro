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

**Governed:** `docs/BETA_ROADMAP.md` + `docs/BETA_REARVIEW.md` + `docs/ALPHA*_ROADMAP.md` (today, just
`ALPHA6_ROADMAP.md`; the glob picks up `ALPHA7_ROADMAP.md` unmodified when
this version closes). This is the promotion path SD-9 describes in its own
telling — a ledger mints an item, the strategy file carries it forward or
promotes it — so it is the only place a new ID gets minted, and the only
place this script treats a collision as an error.

Since 2026-09-27 the roadmap holds open items only, and each closed item's
one-paragraph entry moves to THE CLOSED LEDGER at the end of
`docs/BETA_REARVIEW.md`, where its full build record already lives. The check
reads that ledger as its own mirrored file, so the pair is reported as
MIRRORED (informational), never as a DUPLICATE DEFINITION - the same
relationship the two files had before the rename.

**Advisory, not governed:** `docs/REBUILD.md`, `docs/LESSONS.md`,
`docs/AUDIT.md`, `docs/PROJECT_BRIEF.md`. Each keeps its own internal
numbering for its own purpose — REBUILD.md's `R1`–`R11` are naming/lint
rules and its plate/layer steps are a target-architecture sketch, not
backlog items; LESSONS.md and AUDIT.md number findings from sessions that
predate this file. Renumbering any of them is out of scope for an `~hours`
item and is not what SD-9 asks for — but BETA_REARVIEW.md's own prose already
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
landing in BETA_REARVIEW.md. This mint again coexisted with a concurrent
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
| AXM | AXM.13 | AXM.14 |
| CO | CO.5 | CO.6 |
| DI | DI.5 | DI.6 |
| DO | DO.6 | DO.7 |
| DR | DR2 (bare) | DR3 |
| EN | EN.8 | EN.9 |
| F | F.15 | F.16 |
| G (bare numeric sub-family) | G11r | G12 |
| GO | GO.5 | GO.6 |
| IN | IN.16 | IN.17 |
| IO | IO.6 | IO.7 |
| KERNEL | KERNEL.48 | KERNEL.49 |
| L | L17 (bare) | L18 |
| LE | LE.10 | LE.11 |
| LX | LX.12 | LX.13 |
| PF | PF.6 | PF.7 |
| PI | PI.7 | PI.8 |
| PORT | PORT.13 | PORT.14 |
| S | S6 (bare) | S7 |
| SEC | SEC.6 | SEC.7 |
| SIG | SIG.8 | SIG.9 (SIG.6, SIG.7 reserved — see below) |
| U | U.24 | U.25 |
| V | V8 (bare) | V9 |
| VLA | VLA.4 | VLA.5 |

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
minting** (`BETA_REARVIEW.md`'s "HOW THIS ROADMAP IS ORDERED" section) rather
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

**`AXM` is a new family, minted 2026-09-30 at the owner's call: twelve
items, `AXM.1`–`AXM.12`, under a new line, 🗣🔧🪟 LANGUAGE + MACHINE +
PRODUCT · THE SINGULARITY LINE**, the steps of `docs/SINGULARITY.md` picked
up in the order that each makes the next cheap, with every engine behind its
microscope. Declared in the prefix inventory at the moment of minting, as
`SEC` and `SIG` were. The doc's own handles (Stage 3, step 3.2; `REFLECT`,
`ABDUCE`) stay handles: only a filed step carries an ID, and `SINGULARITY.md`
§12 maps each filed step to its `AXM` number. `METAPROOF.4`, `V.1`, `U.18`
and `LE.7` are steps of the same road and keep their own families and
places.

**`AXM.13` minted 2026-10-03 at the owner's call**: `AXM.1`'s run on a real
fifty-sheet model, taken out of `AXM.1` when it closed as the instrument
with no real model or manager's questions to hand. It is placed after
`AXM.7`, and `AXM.8` waits on it.

**`KERNEL` is a new family, minted 2026-10-06 at the owner's call:
twenty-one items, `KERNEL.1`–`KERNEL.21`, under a new line, 🔧🪟🌍 MACHINE +
PRODUCT + COMMONS · THE KERNEL LINE**, the Frazaro spreadsheet (working
title CALLOSUM) built from `web/CALLOSUM.md`'s scope as a kernel with five
seams, in the order that each item makes the next cheap, the instrument
before the operation. Declared in the prefix inventory at the moment of
minting, as `AXM` was. CALLOSUM's own handles (slice 2, decision 3, the
Commandments' numbers) stay handles; CALLOSUM §14 maps each to its `KERNEL`
number. `PORT.9`–`PORT.11`, `G-USE`, `AXM.2`, `AXM.3` and `AXM.9` are steps
of the same road and keep their own families and places. `SD-20`–`SD-37`,
the eighteen commandments blessed the same day, are standing decisions in
the register, not items, and take no row here. The `PORT` family (minted
2026-10-01, `PORT.1`–`PORT.11`) had no row in the inventory above until this
pass; it has one now. `PORT.12` was minted 2026-10-08: the cut of `vla-lang`,
the language as a crate of its own before `KERNEL.7`, the boundary decided the
same day in Alonzo's sessions and scoped in `BETA_REARVIEW.md`, Part B;
`PORT.13` is next free. `KERNEL.22` was minted 2026-10-09 for the scoping
of the handle and the four calls, the language's half of Alonzo's
`ENGINE.1` (`BETA_REARVIEW.md`, under the KERNEL line, after `KERNEL.7`'s
record), the owner choosing it over `PORT.13` the same day as that
scoping's first decision. `KERNEL.23` was minted 2026-10-09 for the
machine's refusals, the texts Alonzo's `ENGINE.1` found opening with the
row's line, at the owner's approval of that scoping the same day.
`KERNEL.24` was minted 2026-10-09 for the plane kept at the frame's end,
the language's half of Alonzo's `ENGINE.2`, at the owner's approval of its
scoping the same day; the Alonzo session's note had named it `KERNEL.23`,
which the refusals had taken an hour before. `KERNEL.25` was minted
2026-10-09 at its scoping, for the evaluator's hot path made cheaper bit
for bit, the seven fixes Alonzo's `CART.1` measured (its decision 9).
`KERNEL.26` was minted 2026-10-10 at the owner's word, for the two
decisions `KERNEL.20`'s third slice left to the owner over its frame of
19 ms, the language's own functions computed by a run and the rows' values
kept as arrays, filed as one item to be adjudicated together. `KERNEL.27`
was minted 2026-10-10 at the owner's word, for the two further gains
`KERNEL.26`'s profile of its own frame named, the twins' cells made only
when read and a part repeated in one formula computed once, filed as one
item. `KERNEL.28` was minted 2026-10-11 at the owner's word, for the
language's remaining functions computed by a run, `ROW`, `COLUMN`, `MOD`,
`INT`, `ABS`, `SIN` and `CHOOSE`, the ask of Alonzo's `CART.2` (its decision
14), and scoped the same day. `KERNEL.29` was minted 2026-10-12 at the
owner's word, for `LET` and `LAMBDA`, the slice of their own `KERNEL.8`
names, with a parity floor, a named function as fast as the formula it
stands for, and three decisions for its scoping, recursion, functions as
values and visibility, from the owner's requirement, given in an Alonzo
session, that a function written in formulas be first-class. `KERNEL.30`
and `KERNEL.31` were minted 2026-10-12 at the owner's word, from the speed
half of an Alonzo session's audit of the architecture, after `KERNEL.28`
had closed: `KERNEL.30` for the shapes printed, each formula shape of a
plan as rows, its cells, its width at each level of the order and the
reads it makes; `KERNEL.31` for a sheet a range reads this frame kept as
an array, as a `.last` twin is, so that a reduction over the frame's own
cells walks an array. The one-cell-wide chain the same audit named was
left unminted, `KERNEL.28` having brought the running balance level with
Excel's twenty threads. `KERNEL.32` was minted 2026-10-12 at the owner's
word, from `KERNEL.20`'s close, for the doors computing through the
machine, the larger step behind the line's last clause, the view's window
and its cone, which the close ruled out. `KERNEL.33` was minted 2026-10-12
at the owner's word, the three changes Alonzo's playtest pass of Sand asked
of the language filed as one item: the plan's edges found once a shape, a
column a reference holds read once a row, and a row of single cells reached
as one; scoped, built, owner-tested and committed the same day. `KERNEL.34`,
`KERNEL.35` and `KERNEL.36` were minted 2026-10-12 at the owner's word,
from the measurement asked with 0.9.2's preparation of what costs every one
of Alonzo's cartridges and which items move it: `KERNEL.34` for a formula's
write in time with its cells, a fill's write being quadratic today;
`KERNEL.35` for a sheet's readers found by the rows a changed cell stands
on, so that a grid written one fill a row is marked as cheaply as one fill;
`KERNEL.36` for a frame's end copying its changed cells however many they
are. The fourth cost the measurement named, the evaluator's own, 80 to 83%
of a steady frame with a Life written by hand twelve times faster than the
engine's, is the VLA-to-WebAssembly emitter of Alonzo's `CHARTER.md` §4,
left unminted. `KERNEL.38` and `KERNEL.39` were minted 2026-10-12 at the
owner's word, from the findings of `KERNEL.35`'s scoping: `KERNEL.38` for a
fill's shape found by its formula's R1C1 text, as a plain formula's is, so
that one formula written as many fills loads, marks and runs as one fill,
which also answers the scoping's finding that a grid spelled by columns
runs nothing together; `KERNEL.39` for a part the same at every cell
computed once a frame for every shape that holds it. `KERNEL.37` is held
for the `values` projection Alonzo's `ENGINE.8` asks of the language for
its second slice, the view record less its formula rows, named by that
number in Alonzo's roadmap when it was this registry's next free, its
prompt with the owner to relay. `KERNEL.40` is next free. `KERNEL.37` was
minted 2026-10-12 at the owner's word, the prompt relayed, for that
projection, `values`, a window's grid record less its formula rows, so
that an engine's host draws a grid every frame without moving every
formula's text; `KERNEL.40` stays next free. `KERNEL.40` and `KERNEL.41`
were minted 2026-10-12 at the owner's word, from the findings of
`KERNEL.36`'s scoping: `KERNEL.40` for a twin cell a formula's first value
gains, kept in its array and made a cell only when a reader asks, so that a
grid's first frame ends without making every twin again; `KERNEL.41` for a
copy into a twin read through its cells that walks the changed cells in the
twin's order, so that a frame's end copies at any share on every sheet.
`KERNEL.42` is next free. `KERNEL.42` was minted 2026-10-15 at the owner's
word, from Alonzo's `ENGINE.17`, for a window's changes, the cells of a
window whose values changed since a frame, so that an engine's host skips
the view of a still window and repaints only the cells that changed.
`KERNEL.43` is next free. `KERNEL.43`, `KERNEL.44` and `KERNEL.45` were
minted 2026-10-16 at the owner's word of 2026-10-15, from the measurement of
a frame's formulas over 0.9.4 that scoped them: `KERNEL.43` for a run's
lanes typed end to end, a twin's numbers kept beside its slots and read by
stride, so that a fill's frame costs a few nanoseconds a cell; `KERNEL.44`
for a cell alone lowered once to a register program; `KERNEL.45` for a shape
whose changed reads reach most of it computed whole and its readers not
marked. The VLA-to-WebAssembly emitter Alonzo's charter recommends was
measured beside them and not minted. `KERNEL.46` is next free.
`KERNEL.46` was minted 2026-10-16 at the owner's word, from Alonzo's
`ENGINE.7` (decision 4), for a write that empties a cell, a row the
machine's write takes beside `cell`, `formula` and `derived`, so that a host
gives a cell back to no value, as a rebuilt cartridge that no longer writes
it asks. `KERNEL.47` is next free.
`KERNEL.47` was minted 2026-10-17 at the owner's word, from Buffet's
`POT.10` hand test of 2026-10-16 and its ledger's row in `MISE.6`, for a
number format in the record, Excel's format codes beside `general` and
`text`, so that a cell holding 1 under `0.00` draws 1.00, right-aligned as a
number. `KERNEL.48` is next free.
`KERNEL.48` was minted 2026-10-18 at the owner's word, from Alonzo's
`CART.12` scoping (the large board at rate zero, a cascade settled inside an
edit), for a frame's own count, the formula cells whose values the last
complete frame changed, read in constant time on `vla-lang`'s `Machine`, so
that a host steps again while a cascade still moves. `KERNEL.49` is next
free.
`PORT.13` was minted 2026-10-10 at the owner's word, for the
writer's fills Excel cannot hold as one shared formula, found when Excel
refused the Life fixture of `KERNEL.20`'s sixth slice; `PORT.14` is next
free.

The `VLA` family had its row in the inventory above and no sentence here
until 2026-10-12, when the owner asked where `VLA.3` was minted. Its items
are the language's own, in both its implementations (`vla-lang` and the
VBA reference) and the treaty between them, and a closed one leaves
`BETA_ROADMAP.md` for the closed ledger as every item does. `VLA.1` was
minted 2026-10-10 at the owner's word, the day `vla-lang` 0.8.1 went to
crates.io, as the family's first item: the language tests itself, its
fixtures travelling with the crate, its laws holding on generated programs
and text a stranger wrote refused by name, never a crash; scoped, built,
owner-tested and committed the same day. `VLA.2` was minted 2026-10-10 at
the owner's word, the day `VLA.1` closed, from that item's A/B: the port's
expander sharing what the reference shares, so that a macro walking a list
costs its length, not its length squared; scoped, built, owner-tested and
committed 2026-10-12. `VLA.3` was minted
2026-10-10 at the owner's word, on Alonzo's request after its `SPEC.3`,
whose §12 named it the cartridge pass's next step: the expand-time object
model, a template macro reaching a column, an address and a text's
characters as it reaches a row; scoped, built, owner-tested and committed
the same day (`494100a`), and released in 0.9.0. `VLA.4` was minted
2026-10-17 at the owner's word, approving a scoping asked for in chat: the
formula syntax, a VLA expression written as a formula's text and read back,
the expander's `formula-expand` its writer at expand time, so that a
cartridge composes a formula from expressions; it takes the number two
records had given an unminted candidate for the same need (a macro's text
inside `&expand`, `KERNEL.28`'s note and Alonzo's `CART.2` decision 14),
and folds in what remained of `F.7`; scoped, built, owner-tested and
committed the same day, `F.7` closing with it. `VLA.5` is next free.

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
`IN.9` → `IN.10` (user-procedure call/return, BETA_REARVIEW.md,
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
`V.` were already minted (`L.11`, `V.1`) and cited throughout BETA_REARVIEW.md,
but missing from this file's own prefix inventory (the "ordering" section,
now fixed). `DR` (`DR1`, `DR2`, both on `ALPHA6_ROADMAP.md`, DR2 currently
⛔ parked) was never declared anywhere. None of the three were ever
actually ambiguous — that was a documentation gap the audit closed, not a
collision the audit found.

**Spot-checked, not exhaustively reviewed:** the advisory-overlap output is
large (most of BETA_REARVIEW.md's own items are, unsurprisingly, *also* mentioned
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
