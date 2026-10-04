# SINGULARITY - the Analyst ex Machina

*The summit, written down so it can steer. Frazaro's first gear is the imperative sentence: an analyst tells a spreadsheet what to do, and it does exactly that or refuses in words. The second gear is the logic engines: `SQL`, `DATALOG`, `PROLOG` and `OPTIMIZE`, questions and constraints over tables, answered in cells. This document is about the third gear, and about the far end of the road it opens: a spreadsheet that answers a manager the way an analyst would. The owner's name for that end state is the Spreadsheet Singularity; its working title is the Analyst ex Machina. It is treated here as an eventuality, so the road to it can be laid out stage by stage, and each stage can be measured.*

*What this document is not. It is not a commitment, and nothing in it is an item. The stages and steps below carry handles (Stage 3, step 3.2), never IDs; when a step is picked up it is filed in `BETA_ROADMAP.md` under the department that owns it, gets an ID from that department's family, and takes its line in `ID_REGISTRY.md`. Sizes are guesses in `~days`, `~weeks` and `~quarters`. Every standing decision in the register stands, and two of them are load-bearing here: `SD-13` (no outbound network call, ever) and the owner's veto of `GENSYM`, in these words: "determinism is the Chinese wall between this project and the Wild West of LLMs." The Analyst ex Machina is deterministic all the way down. It answers what it can prove and refuses the rest in words. It never guesses.*

*House rules, as for every ledger in `docs/`: append-only; a correction is added and dated, never written over; every figure is a prediction or a measurement and says which; where this file and the code disagree, the code is right.*

---

## 1. The bet, stated so it can be lost

**The claim.** A manager who today messages an analyst for answers about a workbook ("are there any unpaid invoices left?", "why did the accrual go up?", "what if we book the rebate in September?") can get the same answers from the workbook itself, in the same words, with a proof attached, and can tell the difference only by the speed. That is the Imitation Game, played over a spreadsheet: Turing's test with the deception removed, since the judge is allowed to know which column the machine wrote and is asked only whether the answers are the ones the analyst would have given.

**The premise that makes it winnable.** An analyst's replies at a month-end close are not free conversation. They are drawn from a finite repertoire: the same twenty controls, the same fifty questions, the same six kinds of explanation, month after month, in a vocabulary the company already owns. Free chat has no floor; a close has one. A finite grammar can cover a finite repertoire completely, and every question outside it can be refused by name, which is a reply an analyst on a chat also gives ("I'd need to know what you mean by stale"). The singularity is not a machine that answers anything. It is a machine that answers everything the close asks, proves each answer, and names exactly what it would need for the rest.

**The premise that makes it Frazaro's.** Winograd and Flores built The Coordinator on speech acts in 1987 and it failed, because it asked people to classify their own messages as requests, promises and assertions before sending them (Suchman's critique, 1994). Here the form carries the act: a sentence ending in `?` is a question, one containing `should` or `must` is a control, one ending in `.` is a command, one beginning `If ... were` is a hypothesis. Nobody classifies anything. And every act is checked before it runs, refused in words when malformed, and proven by a corpus, exactly as the first gear's sentences are today.

**What the singularity is not.** No statistical model takes part at any stage. No text leaves the machine. Nothing is inferred from what a user "probably meant". The chat window is a sheet: the manager writes questions in column B of a Q&A sheet, the answer arrives from column C, and the conversation is a workbook that can be saved, diffed, audited and replayed. Where an AI belongs in this picture is already decided by `LE.7`: a model may draft sentences in the published grammar, and Frazaro checks, refuses or answers them. The model proposes; Frazaro disposes. That boundary is the reason a boardroom can trust the answers, and it does not move.

## 2. The frame: speech acts, not moods

The owner's observation was that grammatical moods map onto corporate roles: the imperative belongs to the analyst, the declarative and interrogative to the manager. The sharper frame is Austin and Searle's taxonomy of speech acts, because "should" is not a mood but a modality, and because the taxonomy has a slot the moods lack: the declaration, a sentence that changes the record by being uttered. "Approved." is one. So is "This close is final."

| Act | Who | Sentence | Answer | Engine |
|---|---|---|---|---|
| Directive | analyst | `Delete blank rows in range A1:F200.` | the effect, or a refusal | the first gear, shipped |
| Control (deontic) | manager | `No cell in column G of sheet Invoices should contain "Unpaid".` | the violating rows, in column C: red for `must`, yellow for `should`, green when none | `DATALOG` with `not` |
| Definition | manager | `An invoice is overdue when its Due Date is before today and its Status is "Unpaid".` | a relation later sentences use | `Define`, promoted to rules |
| Reconciliation | controller | `Every row of Table Bank should match a row of Table Ledger on Amount and Date.` | the unmatched rows | anti-join |
| Question | executive | `How many invoices are overdue?` | a value in column C of the Q&A sheet, or a spill from it, with its proof | `DATALOG`, `PROLOG`, `SQL` |
| Audit | internal audit, IT | `Which formula columns of sheet Model contain a typed-over constant?` | the plugs | a scan over the workbook's own relations |
| Change | executive | `What changed in sheet Summary since "Close 2026-08.xlsx"?` | the changed inputs in the cone | closure over `refers` |
| Why | executive | `Why did cell B9 of sheet Summary change since "Close 2026-08.xlsx"?` | ranked explanations with evidence | `ABDUCE` (section 8) |
| Hypothesis (irrealis) | executive | `If cell B2 of sheet Model were 100, what would cell D9 be?` | a value from a copy, then discarded | `WHATIF` |
| Abductive goal | executive | `What would cell B2 have to be for cell D9 to be 0?` | one value, or "no such value" | `GOAL` |
| Human hypothesis | manager | `Could cell C5 be in thousands?` | tested: what would follow, and whether it does | deduction plus a check |
| Narration | anyone | `What did you do to sheet Summary today?` | the run, in the past tense | the run log |
| Declaration | CFO | `Sign off on sheet Summary.` | a line in the record: signed as whom, when, which hash | the run log |
| Topic | anyone | `About sheet Invoices:` | nothing; the sentences below it read shorter | the parser's scope |

Eight of these rows are the duties of an analyst on a chat: prepare, check, answer, explain, remember, rehearse, teach (a refusal with directions is teaching) and sign. The machine that does all eight, deterministically, is the Analyst ex Machina.

## 3. Here: `0.7.1`, measured

What stands at the release this document was written against, so each stage below can say what it adds rather than what it imagines.

- **The first gear.** 239 registered sentence rules in `english.vla` (`english_coverage.txt`: `Rules: 239`), each with at least one proof; 567 message ids in `VLA_Messages.bas`, every refusal among them. A program is checked row by row before anything runs, and a run takes an Undo snapshot first.
- **The second gear.** Four engines as worksheet functions, taking their tables as arguments: `SQL`, `DATALOG`, `PROLOG`, `OPTIMIZE` (with `OPTIMIZE_STATUS` and `OPTIMIZE_VIOLATIONS`, which names the rows that break a constraint). `DATALOG`'s tests are 98 proofs in `scripts/proofs/datalog.vla`, 23 of them checked by clingo, a solver of another lineage. `DATALOG` has closure proofs ("a closure over a three-link chain answers all six pairs"), `not`, arithmetic through `let`, text tests, and `count`, `sum` and `textjoin` comparing text exactly. `OPTIMIZE` grounds through `DATALOG`'s fixpoint and searches with an objective ("minimise overtime", `OPTIMIZATION.md`).
- **The interrogative, already present as a slot.** `english.vla` rule 4661: `Show in cell E2 who can-cover "Night" by applying the rules in H2:H4 to the data tables Staff, Shifts, and Leave.` The `{q:question}` sub-grammar is `G-PROLOG`'s, with the answer shapes who, whether, which, how many, each, none, every, alone. A question stands today inside an imperative that names where the answer goes, which rules apply and which tables are in scope. Section 5's Stage 2 makes those three implicit.
- **Rendering.** `EnglishRenderForm` (`VLA_SentenceEngine.bas`) turns a VLA form back into a sentence through the loaded phrasebook. It has no user-facing job yet.
- **Definitions.** `Define hot-pink as "#FF69B4".` exists at top level, for aliases only.
- **Memory.** `TakeRunSnapshot` keeps one Undo snapshot per program (`VLAu_` sheets); `VLA_EventSink.cls` already receives `SheetChange`; `U.18` (a per-run log a user can read) and `V.1` (`verify:` rows) are filed and open; `SEC.18` (an effect class for every reachable form) is filed and open.
- **What does not exist.** Nothing in `src/` walks a formula's precedents. Excel's own `Precedents` property stops at the sheet boundary, so a walker has to read the formula text. No relation reflects the workbook to the engines. No snapshot is compared to another. No sentence ends in `?`.
- **The verification floor.** 30 static checks in `tools/`; the pure suite at 1564, the host suite at 244, `TestDSLs` at 2330, `VerifyReports` at 334 on both backends (the counts in commit `085710a`'s message).

## 4. There: a Monday at the singularity

*A narrative, not a specification. Every sentence in it is meant to be one the grammar of a later stage accepts.*

At 7:00 the close workbook opens and its Controls sheet runs before anyone has read a number. Twenty-two rows, each a sentence the controller wrote last spring and has not touched since: `No cell in column G of sheet Invoices should contain "Unpaid".` `Cell B9 of sheet Summary must equal cell C4 of sheet GL.` `Every row of Table Bank should match a row of Table Ledger on Amount and Date.` Twenty are green. Row 14 is yellow, since it says `should`, and column C beside it does not say "failed"; it says `3 rows of Table Bank match no row of Table Ledger: rows 41, 42 and 97.` Row 19 is yellow too: `The total of column F of sheet Sales is 11.8% above the same column of "Close 2026-08.xlsx"; the tolerance is 10%.` No `must` failed, so the run reached its end; had cell B9 not equalled GL's C4, row 8 would be red, worded as a refusal is, and the sign-off below would refuse while it stood. One dialog at the end lists the two yellow rows, and the record keeps the same list.

At 8:15 the manager opens the Q&A sheet and writes in column B, under the topic line `About sheet Summary:`. `Why did cell B9 change since "Close 2026-08.xlsx"?` Column C answers in three sentences, and they are Peirce's three: `B9 rose by 6,300. If cell C5 of sheet Model were ten times its August value, that would account for the whole change: C5 was 700 and is 7,000, and every other input B9 depends on is unchanged. If so, cell F12 should also be 6,300 above August: it is.` The manager writes `Could cell C5 be in thousands?` and reads: `If C5 were in thousands, D9 would be 1,240; it is 1,240,000. So no.` Then: `Which formula columns of sheet Model contain a typed-over constant?` and reads `Column D of sheet Model: cell D17 holds 4,500 where D16 and D18 hold formulas.`

At 9:00 the executive asks the only question executives ask: `If cell B2 of sheet Model were 100, what would cell D9 be?` The answer comes from a copy that is deleted before the cell is written. At 9:04, `Sign off on sheet Summary.` writes one line to the record: signed as whom, at what time, against which hash of the workbook, with the Controls sheet's result beside it. The analyst, who built the model, spent the morning on the two rows that were not green.

## 5. The road, in seven stages

Each stage names what becomes sayable, the engine under it, the proof shape that pins it, what it stands on, the catch, and a size. The stages are ordered by dependency, not by value; Stage 1 is where the value starts.

### Stage 0 - The substrate: the workbook as relations

*The third gear is the second gear pointed at the workbook itself. Everything above depends on the workbook being visible to the engines as data.*

- **0.1 `REFLECT`: the workbook's own relations.** Eight relations any engine can take as a table: `cell(sheet, addr, value)`, `formula(sheet, addr, text)`, `refers(from, to)`, `name(name, refersto)`, `sheet(name, hidden)`, `table(name, sheet, range)`, `changed(addr, old, new)` and `ran(program, row, addr, when)`. Lazy and anchored: a question about `B9` materializes `B9`'s cone, never the workbook; a whole-sheet scan reads the used range as one array. *Proof shape:* fixture workbooks in `examples/` with the expected relations written down beside them. *The catch:* a million-cell model cannot be a million facts per question; the measurement item comes first: how long does reading every formula of a fifty-sheet model take, and how big is a typical cone. `~weeks`
- **0.2 The formula-reference reader.** A tokenizer for the references inside a formula (`A1`, `$A$1`, `Sheet!A1:B2`, names, structured references, external links), which is what `refers` is built from and what "read this formula to me" needs too. A parser for references, not yet for the whole formula language. *The catch:* implicit intersection and spilled ranges; references inside `INDIRECT` are refused by name as unreadable. `~weeks`
- **0.3 `DIFF`: `changed(addr, old, new)`.** Against the Undo snapshot, against a saved copy, or against another file opened read-only (`G-FILES`). One relation, three sources. *The catch:* a cross-file diff needs the same sheet names and shapes; a renamed sheet is reported, not matched. `~days` after `G-FILES`
- **0.4 Scope and placement.** The three things a standalone question leaves implicit: where the answer goes, which rules apply (the Rules sheet, or a `By applying the rules in H2:H4:` topic line) and which tables are in scope (the workbook's Tables, unless a topic line narrows them). *Placement, decided 2026-09-29 (the owner's call):* questions live on a **Q&A sheet**, a document type of its own that keeps the Frazaro tab's convention, the sentence in column B and its result in column C, and lets an answer spill rightward from column C. Not on the Frazaro tab itself, whose columns D onward are the scratch space for building the sentences in column B; and never downward, where a spill collides with the next question's row. A one-cell answer (a count, a yes or no, one value) sits in column C alone. `SD-19` applies to the rest: the convention is written down once, and a sentence can always name the choice explicitly. `~days`, mostly deciding
- **0.5 Topic sentences instead of pronouns.** `SD-16` forbids anaphora, and the singularity does not need it: `About sheet Invoices:` sets a scope that the sentences under it read against, until the next topic line or a blank row. Deixis to a declared topic, not reference to earlier text; deterministic, and visible on the sheet. `~days`

### Stage 1 - Assertives: the Controls sheet

*`V.1` re-scoped. `verify:` rows are not a feature of the analyst's program; they are a document type for managers, and it reuses Check, Run and column C unchanged.*

- **1.1 `must` and `should`.** Two words, two severities, decided 2026-09-29 (the owner's call). A failed `must` is a red cell in column C, worded as a refusal is, and the run counts as failed: `Sign off` refuses while any red stands. A failed `should` is a yellow cell in column C with the violations named, and the run goes on; when it ends, one dialog lists every yellow row of the run, and the record (1.6) keeps the same list. A Controls sheet keeps the Frazaro tab's convention: the control in column B, its result in column C, green when nothing violates it. `should` will be in every control sheet forever, so each spelling gets its `GRAMMAR_SINCE.md` row from the first slice. `~days`
- **1.2 Quantifiers.** `no`, `every`, `some`, `at most N`, `at least N`, `exactly one`, over rows of a range or Table, with the row-level `cond` sub-grammar the first gear already has. *Proof shape:* `(test-violations "No cell in column G of sheet Invoices should contain \"Unpaid\"." (rows 5 9 14))` in a proof file beside a fixture. `~weeks`
- **1.3 Tie-outs and tolerances.** `must equal`, `should be within 5% of`, `should be within 100 of`, between two cells, two totals, or a total and last month's file. *The catch:* a tolerance is a number, never a feeling; "roughly" is refused. `~days`
- **1.4 Reconciliation.** `Every row of Table Bank should match a row of Table Ledger on Amount and Date.` An anti-join; the answer is the unmatched rows on each side. *Pays into:* every month-end close. `~days` after 1.2
- **1.5 Definitions, promoted.** `An invoice is overdue when its Due Date is before today and its Status is "Unpaid".` becomes a `DATALOG` rule over the Table's rows, and `overdue` is then a word every later control and question may use. The company's glossary becomes executable, one sentence per term, on a Definitions sheet. *The catch:* a definition that names a column the Table lacks is refused at Check, in words; a definition may not redefine a built-in word (`SD-4`). `~weeks`
- **1.6 The record.** `U.18`'s run log, with a line per control result, and `Sign off on sheet Summary.` as the declaration that closes it: signed *as* whom (`Application.UserName` is a claim, not a proof, and the record says so), when, against which hash. Evidence for the auditor, produced by the act itself. `~days` after `U.18`

### Stage 2 - Questions

- **2.1 The standalone question.** A sentence ending in `?` is `G-PROLOG`'s `{q:question}` with Stage 0.4's conventions supplying the rest. The terminator carries the act. `~days` after 0.4
- **2.2 Questions over Tables.** `How many invoices are overdue?` `Which customers owe more than 10,000?` `Who approved invoice 1042?` `What is the total of column F where Region is "West"?` The wh-words, comparatives and aggregates, over Tables and defined terms. `~weeks`
- **2.3 Questions over the workbook: the audit list.** `Which formula columns contain a typed-over constant?` `Which formulas are inconsistent with their neighbours?` `Which named ranges are unused?` `Which cells are referenced but empty?` `Which totals do not foot?` `Which sheets are hidden?` `Which formulas reference another workbook?` Each is one query over `REFLECT`. This is the checklist internal audit runs by eye and end-user-computing tools charge for; the plug question alone earns the stage. *Pays into:* `IT_REVIEW.md`, which gains a sentence: the workbook can be asked where its risks are. `~weeks`
- **2.4 Why-provenance of an answer.** `Why do you say Bob can cover Night?` answers with the derivation: the rules and the rows that support the answer. Built by backward reconstruction over the computed fixpoint, under a depth budget, so the forward pass allocates nothing per tuple and `check_datalog_per_tuple_alloc` stays true. *The catch:* a derivation over recursion is a tree, and the answer prints one branch and says how many more there are. `~weeks`
- **2.5 Superlatives and top-N.** `Which five customers owe the most?` `Which region grew fastest?` (arithmetic, with the method named). `~days`

### Stage 3 - Change

- **3.1 What changed.** `What changed in sheet Summary since "Close 2026-08.xlsx"?` A spill of `changed`, filtered to a sheet, a range or a Table. `~days` after 0.3
- **3.2 Cause.** Three rules over `REFLECT` and `DIFF`: `input(X)` holds for a cell with no formula; `feeds(X, Y)` is the closure of `refers`; `cause(Y, X)` holds when `feeds(X, Y)`, `changed(X)` and `input(X)`. `Which inputs changed that cell B9 depends on?` is `cause(B9, X)?`, and closure is what `DATALOG` is for. `~days` after 0.1
- **3.3 `VARIANCE`: the arithmetic of a change.** For a sum, each component's delta, exactly. For a product, sequential substitution in a fixed, named order (price, then volume, then mix), which is the method every FP&A team applies by hand. The answer is the top contributors to the change, with the residual named, never hidden. *Proof shape:* a fixture model with a known decomposition. *The catch:* the order is a convention, so the sentence may name it and the record always states it. `~weeks`
- **3.4 `TEMPORAL`: the vocabulary of periods.** `last month`, `since August`, `as of the 15th`, `quarter to date`, over a calendar relation, and `during`, `before`, `overlaps` for intervals (Allen's algebra, the finite part). A snapshot series, one per close, gives `What was cell B9 on the 15th?` a source. `~weeks`

### Stage 4 - Abduction

*Section 8 in full. The stage that makes the machine a detective.*

- **4.1 `ABDUCE`: the fifth engine.** Given the rules, an observation (a control that failed, a number that changed, a total that does not foot), a set of abducible facts, and the controls as integrity constraints, find the smallest sets of abducible facts that would make the observation a matter of course. `OPTIMIZE`'s search with a different objective: fewest abduced facts, ties broken in a fixed order. clingo answers the same question natively (choice rules and `#minimize`), so the oracle exists on day one. `~quarters`
- **4.2 The explanation phrasebook.** The abducibles are not invented; they are a catalogue of explanation schemas in a file, each with a signature, a prediction and a test (section 8.3). A header file, in the house sense: a finance edition, an operations edition, an org's own. `~weeks` for the first fifteen
- **4.3 The why-question.** `Why did cell B9 change?` `Why is cell D9 negative?` `Why did the control on row 14 fail?` answered as hypothesis, prediction and test, in that order, in words. `~weeks` after 4.1
- **4.4 The human hypothesis.** `Could cell C5 be in thousands?` `Suppose row 14 of Table Sales is a duplicate.` `Rule out a unit error in column C.` The manager abduces; the machine deduces what would follow and tests it. Half of detective work is this half, and it needs no search at all. `~days` after 5.1

### Stage 5 - Hypotheticals

- **5.1 `WHATIF`.** `If cell B2 of sheet Model were 100, what would cell D9 be?` Copy, set, recalculate, read, delete. Deterministic; nothing is written to the model. *The catch:* a workbook whose formulas read the clock, a random number or an external link cannot be rehearsed faithfully and is refused by name, the same rule Spitball 1 gives the time scrubber. `~days`
- **5.2 `GOAL`.** `What would cell B2 have to be for cell D9 to be 0?` Bisection over a monotone formula, with monotonicity checked rather than assumed; Excel's own Goal Seek as a second opinion; `OPTIMIZE` when the unknowns are whole numbers. The answer is one value, or "no such value between 0 and 1,000,000", never an approximation without its tolerance. `~weeks`
- **5.3 Scenarios as sentences.** `Under the scenario "Rebate in September":` as a topic line, with the changed inputs listed once and every question below it answered on the copy. `~weeks`

### Stage 6 - The unanticipated question, and the game itself

- **6.1 The gap ledger.** Every refused question is written to a local ledger with the point where the reading stopped, clustered by the rule it nearly matched. `SD-7` says no grammar section is scheduled without a sentence that needs it; this is where those sentences come from. Local, visible, and nothing leaves the machine. `~days`
- **6.2 Definitions by the manager.** The commonest refusal at a close is a word the workbook does not know (`stale`, `plug`, `out of policy`). The refusal names the fix: one Definition sentence, written by the person who owns the word. `~days` after 1.5
- **6.3 The drafting bridge, `LE.7`.** For the question the grammar cannot yet say, a model may draft candidate sentences in the published grammar; Frazaro checks them, refuses the malformed, answers the rest. The model never sees the data unless the person shows it, and never answers anything itself. This is the only place an AI stands in the picture, and it stands outside the wall. `~weeks`, filed
- **6.4 The Imitation corpus.** The questions corpus is to the third gear what `instructions.txt` is to the first: real close questions and controls over fixture workbooks, each with the answer an analyst gave, as proofs. Started now, with the owner's own repertoire, and grown from the gap ledger. *Floors:* a count that never goes down, as `check_proofs.ps1` holds `DATALOG`'s. `~days` to start; never finished
- **6.5 The game.** The protocol is in section 7. Passing it is the summit.

## 6. The missing pieces, in two tables

### Engines

| Engine | Answers | Reuses | Oracle |
|---|---|---|---|
| `REFLECT` | the workbook as relations | `VLA_Relation`'s tables from ranges | fixture workbooks with expected relations |
| `DIFF` | `changed(addr, old, new)` | Undo snapshots, `G-FILES` | two fixtures and their known delta |
| `VARIANCE` | a change decomposed into contributions | `refers`, arithmetic | a model with a known decomposition |
| `ABDUCE` | minimal explanations of an observation | `OPTIMIZE`'s search, `DATALOG`'s grounding | clingo, choice rules and `#minimize` |
| `WHATIF` | a value under a hypothesis | the snapshot machinery | the model itself, on a copy |
| `GOAL` | the input that yields a target | bisection, Goal Seek, `OPTIMIZE` | Goal Seek as the second opinion |
| `TEMPORAL` | periods, intervals, as-of | a calendar relation, a snapshot series | proofs over a fixed calendar |
| `PROVENANCE` | why an answer holds | backward reconstruction over the fixpoint | clingo's own justification, where it can |
| `UNITS` | which cells mix units, thousands against ones | Spitball 33's unit definitions | fixtures with planted unit errors |
| `MATCH` | which names are the same name | `SameText`, named normalization rules, a stated edit-distance threshold | a fixture of known pairs |

### Grammar

| Family | Words | First needed by |
|---|---|---|
| Deontic | `must`, `should`, `never` | Stage 1 |
| Quantifiers | `no`, `every`, `some`, `at most`, `at least`, `exactly one` | Stage 1 |
| Tolerances | `equal`, `within N of`, `within N% of` | Stage 1 |
| Definitional | `An X is Y when ...`, `X means ...` | Stage 1 |
| Performative | `Sign off on`, `Mark ... as final`, `Approve` | Stage 1 |
| Interrogative | `which`, `who`, `how many`, `whether`, `what is`, the terminator `?` | Stage 2 |
| Comparatives and superlatives | `more than`, `at least`, `the most`, `the five largest` | Stage 2 |
| Temporal | `since`, `last month`, `as of`, `during`, `quarter to date` | Stage 3 |
| Causal | `why`, `what explains`, `which inputs changed` | Stages 3 and 4 |
| Hypothetical | `if ... were`, `suppose`, `could ... be`, `rule out` | Stages 4 and 5 |
| Abductive goal | `what would ... have to be` | Stage 5 |
| Topic | `About ...:`, `Under the scenario ...:` | Stage 0 |

Every family is finite, every word gets a `GRAMMAR_SINCE.md` row when it ships, and `SD-4` holds: a shipped spelling keeps its meaning.

## 7. The Imitation Game, as a protocol

*Defined so it can be failed. Turing's game was about deception; this one is about equivalence, and the judge knows which column is the machine's.*

1. **The board.** One workbook, real in shape and synthetic in content, with a Controls sheet and a Q&A sheet. A human analyst who knows the workbook and the company's vocabulary.
2. **The questions.** Column B of the Q&A sheet is filled by a manager who has not read this document, with what they would have asked on a chat that morning. Nothing is rehearsed and nothing is pre-filtered.
3. **The two sheets.** The analyst answers each question on a copy of the Q&A sheet, from the workbook alone, in the time they would normally take. Frazaro answers on the original, from column C. The judge reads the two sheets row by row.
4. **Scoring.** A row scores for the machine when its answer is the analyst's, or when its refusal names exactly what it would need (a definition, a table, a rule) and the analyst, given that, agrees the question was unanswerable from the workbook as it stood. A row scores against it when it answers wrongly, or refuses a question the analyst answered from the workbook alone.
5. **The levels.** Level 1: the Controls sheet scores every control the analyst runs by hand. Level 2: every recurring question. Level 3: every what-changed question. Level 4: every why-question, with the analyst agreeing the explanation is the one they would have given, or a better one. Level 5: every what-if and what-would-have-to-be. Level 6: the unanticipated questions, after one round of definitions written by the manager.
6. **The pass.** Level 6 at one close, then again at the next close with the same sheets and a different manager. The judge may notice the speed. That is allowed.

The corpus that this protocol produces is the third gear's goldens. A level passed is a floor that never goes down.

## 8. Abduction, at length

*The owner's pet, and the engine that turns answers into detective work.*

### 8.1 Peirce's formula, in cells

Peirce's syllogism of abduction: *the surprising fact C is observed; but if A were true, C would be a matter of course; hence there is reason to suspect that A is true.* In a workbook the surprising fact is always one of a few kinds: a number that changed, a control that failed, a total that does not foot, a sign that flipped, a cell that is empty where it was not. The hypothesis A is drawn from a finite catalogue of the things that happen to spreadsheets. And Peirce's method is not abduction alone but a triad: abduction proposes A, deduction works out what else must be true if A holds, and induction checks that prediction against the facts. The three sentences of a why-answer are that triad, in that order:

> B9 rose by 6,300. *(the surprising fact)*
> If cell C5 of sheet Model were ten times its August value, that would account for the whole change: C5 was 700 and is 7,000, and every other input B9 depends on is unchanged. *(abduction: the hypothesis, with its evidence)*
> If so, cell F12 should also be 6,300 above August: it is. *(deduction, then induction: the prediction, and its test)*

Eco and Sebeok's *The Sign of Three* (1983) made the case that Dupin and Holmes reason exactly this way, and that "deduction" in the stories is Peirce's abduction wearing a deerstalker. The register of the answer can borrow the method without the costume; the costume can be an edition.

### 8.2 The engine: abductive logic programming, over `OPTIMIZE`

The formal shape is Kakas, Kowalski and Toni's abductive logic programming (1992): a program (the `DATALOG` rules, including the definitions and `REFLECT`), a set of abducible predicates (the explanation schemas, section 8.3), integrity constraints (the controls), and an observation (the goal). An explanation is a set of abducible facts that, added to the program, entails the observation and violates no constraint; a *minimal* explanation adds nothing it does not need. That is a search with an objective, which is what `OPTIMIZE` already is: ground through `DATALOG`'s fixpoint, search over the abducible atoms, minimize their count, break ties in a fixed order so the answer is the same on every run. clingo states the same problem in three lines (a choice rule over the abducibles, a constraint that the observation be explained, `#minimize` over the chosen facts), so the second-lineage oracle that checks `DATALOG`'s proofs today can check `ABDUCE`'s from the first proof.

Ranking without statistics: fewest abduced facts first (parsimony); among equals, the hypothesis that accounts for more of the observed delta (coverage, which is arithmetic, exact); among those, the catalogue's own order. No probabilities, no priors, no learning. A ranking a person can re-derive by hand is the only kind the record can carry.

### 8.3 The explanation phrasebook

An abducible is not a bare predicate; it is a schema with three parts, written in a file the way sentence rules are:

- a **signature**: the query that finds its evidence (`changed(X)` and `input(X)` and `feeds(X, Y)`);
- a **prediction**: what else holds if the schema explains the fact (a sibling total moves by the same amount);
- a **test**: the query that checks the prediction against the workbook now.

The first catalogue, finance-flavoured, fifteen entries the owner will recognize from a decade of closes: an input changed; a formula changed; a row was added or removed inside a summed range; a summed range stops one row short (the total that does not foot); a typed-over constant in a formula column (the plug); a factor of a thousand or a million (thousands against ones); a factor of a hundred (a percentage entered as a whole number); a sign flipped; a row duplicated; a row missing against last month; a filter left on, or rows hidden; a reference shifted by an inserted row; a date rolled into the wrong period; a rounding difference at a stated tolerance; text stored as a number, or a number as text; a stale external link; a circular reference. Each carries its own proof: a fixture with the defect planted, and the expected explanation. An org adds its own schemas ("the rebate is booked a month late") in its own file, under `SEC.2`'s consent, since a schema is a rule.

Abduction cannot invent a schema. A fact no schema explains gets the honest answer: `No explanation in the catalogue accounts for the change in B9; the inputs that changed are C5, C9 and E2.` That sentence is the gap ledger's business, and the next schema's.

### 8.4 The detective's dialogue

Half of detective work is the machine proposing; the other half is the person proposing and the machine testing. `Could cell C5 be in thousands?` is a hypothesis handed over, and needs no search: deduce what would follow (`D9 would be 1,240`), test it (`it is 1,240,000`), report (`So no.`). `Suppose row 14 of Table Sales is a duplicate.` re-runs the controls and the questions under that supposition on a copy. `Rule out a unit error in column C.` runs one schema's test and reports the result. Sherlock's `When you have eliminated the impossible` is an integrity constraint; a schema whose prediction fails is eliminated, in words, and the record keeps the elimination beside the conclusion. Every step is a sentence, every sentence is checked, and the whole investigation is a sheet that can be replayed next month against the same board.

### 8.5 What abduction here is not

It is not diagnosis by likelihood. It never says "probably". It says which of a finite set of explanations fit, in an order a person can re-derive, with the evidence and the test beside each, and it says when none fit. That is less than a human detective can do and exactly as much as a record can carry.

## 9. The ceiling, honestly

- **The unanticipated question.** A finite grammar answers the repertoire and refuses the rest. The refusal is the floor of honesty, and the gap ledger, the manager's definitions and `LE.7` are the three ways the floor rises. It never becomes a chatbot, by decision.
- **Judgment.** An analyst knows the rebate is booked a month late because someone told them at lunch. The machine knows it when someone writes the schema. Institutional knowledge enters through sentences or not at all, which is the point, and also the limit.
- **Wrong but consistent data.** A control over a Table whose rows are wrong passes green. Controls detect; they do not prevent. What becomes checkable is the stated expectation, which no one writes down today.
- **Identity.** `Sign off` records a claim of identity, not a proof of one. The record says "signed as", never "signed by", until the host offers something stronger.
- **Size.** A cone can be large and a workbook can be a million cells. Every engine above is anchored at a cell, budgeted, and interruptible, and each ships behind a measurement item, as `OPTIMIZE` did. Numbers here are predictions until the measurement replaces them.
- **Recursion and provenance.** A derivation over a recursive rule is a tree that can be wide; the answer prints one branch and counts the rest. Per-tuple recording during the forward pass is refused by the project's own check, so provenance is reconstructed backward, on demand, under a budget.

## 10. Order of work

By dependency and by value, both:

1. **Measure first.** Read every formula of a real fifty-sheet model into `REFLECT`'s relations and time it; size ten typical cones. This is Stage 0.1's measurement item, and it decides the engine's shape before a line of it is written.
2. **Stage 1 before Stage 0 is complete.** The Controls sheet needs quantifiers and the row-level `cond`, not `refers`. `V.1`, re-scoped as the manager's document type, is the first thing a boardroom sees, and it is mostly wiring over Check and column C.
3. **Stage 0.2, the reference reader**, because it unlocks Stage 2.3's audit list, all of Stage 3, and "read this formula to me" at once.
4. **Stage 2.3, the audit list**, with the plug question first. One query, no new engine, and the sentence internal audit has wanted for twenty years.
5. **Stage 3.2, cause**, three rules and a closure `DATALOG` already proves.
6. **Then `ABDUCE`**, with its first fifteen schemas and clingo beside it from the first proof.
7. **The Imitation corpus starts on day one**, with the owner's own close repertoire, and every stage above adds its proofs to it.

## 11. The numbers, and how to re-run them

- Registered rules: `scripts/polyglotta/english_coverage.txt`, line 2 (`Rules: 239`).
- Message ids: `powershell -File tools\run_checks.ps1 -Filter check_proofs -ShowAll` prints `message ids read from VLA_Messages.bas: 567`.
- Proofs: the same run prints `98 proof(s), floor 98` and `23 proof(s) exported ... floor 23`.
- Static checks: `powershell -File tools\run_checks.ps1` (30 at `0.7.1`).
- Suite counts: the `Owner-verified live` paragraph of the latest item commit (`git log -1 --format=%B` on `085710a`: pure 1564, host 244, `TestDSLs` 2330, `VerifyReports` 334).
- The question slot: `grep -n "{q:question}" scripts/polyglotta/english.vla` (rule at line 4661).
- The renderer: `grep -n "Public Function EnglishRenderForm" src/VLA_SentenceEngine.bas`.
- Nothing walks precedents: `grep -rn -i "precedents" src/` returns only the English word in comments.

---

## 12. Filed, 2026-09-30

*Appended under the house rules above: the steps picked up, and where they went. Handles stay handles; a filed step carries an ID.*

The steps below were filed on `BETA_ROADMAP.md` under a new line, 🗣🔧🪟 LANGUAGE + MACHINE + PRODUCT · THE SINGULARITY LINE, family `AXM` (Analyst ex Machina), in the order that each makes the next cheap, with every engine behind its microscope; the reasoning and a fan-out table are in `BETA_REARVIEW.md` under the same heading. `METAPROOF.4` (the METAMETAMACRO line: `(tables …)`, a proof handing its program table arguments) is the first step of that order and keeps its ID, so the third gear's corpus has its notation before its first proof.

| Step here | Filed as |
|---|---|
| 0.1's measurement item | `AXM.1` |
| 1.1, with `V.1` re-scoped; 1.2, 1.3 and 1.4 as its later slices | `AXM.2` |
| answers as sentences, implicit throughout (`EnglishRenderForm`'s first job) | `AXM.3` |
| 1.6, after `U.18` | `AXM.4` |
| 1.5 | `AXM.5` |
| 6.1 and 6.4 | `AXM.6` |
| 0.2 | `AXM.7` |
| 0.1 | `AXM.8` |
| 0.4, 0.5, 2.1 and 2.2 | `AXM.9` |
| 2.3 | `AXM.10` |
| 0.3, 3.1 and 3.2 | `AXM.11` |
| the corpus before 4.1, and 4.2's first fifteen schemas as prose | `AXM.12` |

Not filed, still handles: 2.4, 2.5, 3.3, 3.4, 4.1, 4.3, 4.4, 5.1, 5.2, 5.3 and 6.5. 6.3 is `LE.7` already.

**Two corrections to §3, dated 2026-09-30.** "Memory" omits one thing that stands: `VLA_Log` (`LogParseFailure`, `VLA_IDE.bas`) already keeps every misunderstood sentence with its time, row, text and message on a very-hidden sheet in the user's own workbook, local only, and Copy Feedback exports it. That is Stage 6.1's gap ledger without the clustering, which is why 6.1 is filed inside `AXM.6` at `~days` rather than built from nothing. And "no sentence ends in `?`" understates it: the tokenizer folds `?` and `!` into `.` (`EnTokenize`, `VLA_SentenceEngine.bas`), so a sentence ending in `?` is read today as if it ended in `.`, which is why the act seam (`AXM.2`) comes before any question.

**0.1's measurement item, dated 2026-10-03.** It is now two items. `AXM.1`, the instrument, closed with both its controls passing. `AXM.13` is its run on a real fifty-sheet model, and it waits on a reader that sizes true cones: `AXM.7`, or the Rust port's `PORT.8`, `REFLECT` over a file. No measurement has replaced a size prediction in this document yet: Stage 0.1's catch, §9's "Size" and §10's first step stand as written.
