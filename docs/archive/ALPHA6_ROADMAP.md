# ALPHA 6 ROADMAP — the live ledger

*Terse by design. Alpha 5 is closed history (every item ✅ or explicitly parked
there); its records hold the reasoning that got us here. New passes append their
records HERE, but the item lines below stay one breath long. Checkmarks come only
from the owner's explicit word. Baselines at open: self-test **557/0**, counters
**102 rules / 16 macros / 199 tests (5 expected fails)**, corpus family
`english.txt` + `spreadsheet.vocab` + `hello.vla`, stamps core L18.1 / English
G12.1 / IDE G12.0 / Tests L18.0 / DevRig L8.2 / Runtime S5.2 / Build S5.4.*

---

## Tranche 1 — OPTIMIZATION *(all gated on P-PROF's before-numbers: measure first, the standing discipline; the sub-dial goes first whenever this tranche opens)*

- ⬜ **P-PROF** — the sub-dial: per-phase timing (tokenize / parse / expand / emit) behind one switch; every later P-item cites its numbers.
- ⬜ **P-NTH** — `Nth`'s Collection walk on hot paths → indexed access or cached cursors, goldens as witness.
- ⬜ **P-TOK** — tokenizer pass: single-scan buffer, no per-char concatenation.
- ⬜ **P-DICT** — the recorded transpile lead: macro/registry lookups from `Collection` error-traps to a real keyed store.
- ⬜ **P-PRELUDE** — parse the prelude once, reuse the forms (today it re-tokenizes per transpile).
- ⬜ **P-PROBE** — load-time probe transpiles batched/cached; Reload latency at 1,000-rule scale.
- ⬜ **P-BULK** — array-slab runtime helpers (read/write ranges in one COM call); quote literals are the calling convention.

## Tranche 2 — VLA *(the middle layer grows only as the Grammar Tranche demands)*

- ⬜ **L.11** — runtime `'@doc:` annotations surfacing through apropos.
- ⬜ **L-PIVOT-HELPERS** — Tier-2 runtime verbs for pivot creation/refresh (PivotCache/PivotTable plumbing behind honest one-call helpers), introduced with the pivot grammar pass, not before.
- ⬜ **L-FILE-HELPERS** — CSV/text import verbs (open-read-splice into a tab; QueryTables or plain I/O — adjudicate on the ledger), introduced with the file grammar pass.
- ⬜ **L-SHEET-HELPERS** — tab create/copy/move/rename verbs beyond today's Add, as the tab grammar pass requires.
- ⬜ **L0.2** — multi-row raw forms in the IDE (continuation while parens unbalanced; the L12 balance counter powers it).
- ⛔ **DR2** — the reload consolidation, parked after DR1's crash; staging hypotheses recorded on Alpha 5 (compile via `Application.OnTime` after return, or drop the compile step).

## Tranche 3 — GRAMMAR *(the session's main course; every pass predicts golden shapes exactly)*

- ⬜ **G11r** — functor carry-through revision (definition on the Alpha 5 ledger).
- ⬜ **G7** — slot defaults.
- ⬜ **G8** — number words and ordinals.
- ⬜ **G6** — list-valued slots (`{r:cell+}`) — the enabler for the sentence passes below.
- 🔒 **G9** — conjoined predicates — STAYS GATED on pilot evidence; does not open without the owner's word.
- ⬜ **G-PIVOT** — "Create pivot [table] from range A with rows of B, C, D and columns of E, F, G and filters of H." — design on the ledger FIRST (the heaviest sentence the grammar has attempted; leans on G6's lists and L-PIVOT-HELPERS).
- ⬜ **G-TABS** — "Create new tab called X after|before tab Y." and siblings (copy/move/rename).
- ⬜ **G-FILES** — "Copy contents|values of 'file.csv' into tab Y." and siblings (leans on L-FILE-HELPERS; path quoting rides G12's period rule).
- ⬜ **G-TABLES** — table-operation sentences (sort/filter/total a range as a unit) as the pilot's workload demands.

## Tranche 4 — INTERFACE *(in the owner's stated order)*

- ⬜ **U.15** — lazy vocabulary self-heal after state wipes (design on Alpha 5).
- ⬜ **U1** · ⬜ **U3** · ⬜ **U2** · ⬜ **U5** — the panel wave (definitions on their original ledgers).
- ⬜ **V.1** — `verify:` rows (design sketch on the Alpha 5 ledger — "the users' harness").
- ⬜ **U.12** — the IDE apropos surface (inherits three tiers + worksheet functions).
- ⬜ **U.14** — `VlaTryTranspile` (design on Alpha 5; DevRig ships → manual VBE replace first).
- ⬜ **U.13** · ⬜ **U.11** · ⬜ **U.6** · ⬜ **U.7** · ⬜ **U.8** — exploration rig and the remaining UI items, as ordered.
