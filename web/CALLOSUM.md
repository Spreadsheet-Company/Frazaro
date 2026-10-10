# CALLOSUM - the bridge between mind and machine

*The corpus callosum is the bundle of fibres that joins the two hemispheres of a brain, so that what one half sees the other half can name. This file is named for the bridge Frazaro is building between the mind that writes a sentence and the machine that holds a grid. It is the record of one sitting, 2026-10-05, held while another session prepared `0.8.0`. The owner asked three questions in turn. First, from ten thousand feet: at the rate of the last two months, what will be added to Frazaro in a year, and in ten; what the reflect feature in particular makes possible ("the spreadsheet is not currently accessible trivially to itself; this is Lisp in the machine, Lisp ex Machina"); and what tectonic shifts lie beyond the roadmap and `HORIZON.md`. Second, to scope the Frazaro spreadsheet: a two-pane WebAssembly application in Elm's shape, a grid that is only a view of an underlying source, a file of ten thousand English sentences pasted in to generate a workbook of arbitrary complexity, with the ten-thousand-and-first sentence modifying or querying the live model. Third, that the session live here, under this name, with the owner's thesis as its first line:*

> *This tool is the new bridge between mind and machine, with natural language macros being the first in a long line of semantic abstractions to democratize automation for the everyperson.* (the owner, 2026-10-05)

*What this document is not. It is not a commitment, nothing in it is an item, and nothing in it carries a handle. An idea here that is ever picked up goes through `SINGULARITY.md` if it is third-gear work, or straight onto `BETA_ROADMAP.md` under the department that owns it, and takes an ID there. Every standing decision in the register stands; the two that are load-bearing here are `SD-13` (no outbound network call, ever) and the owner's veto of `GENSYM` ("determinism is the Chinese wall between this project and the Wild West of LLMs"). The spreadsheet described below is deterministic all the way down, and the wall is not moved by any row of it.*

*House rules, as for every ledger in `docs/`: append-only; a correction is added and dated, never written over; where this file and the code disagree, the code is right; where this file and `HORIZON.md` section 12 disagree, this file is later and says so by date in section 7, and `HORIZON.md` carries a dated pointer here. Marks follow `HORIZON.md`: an idea past the roadmap is **math exists** (the result is published and the only work is building it in), **engineering** (nothing new has to be true, but someone has to build it) or **fiction** (it needs the world to change); an argument about the present is **exhibit** (a dated, public, checkable fact), **shelf** (a claim a document in this repository already makes, cited) or **recommendation** (the sitting's advice, the owner's to take or leave). Every figure is a **measurement**, dated, with its method in section 10, or a **prediction** in the house's `~` form. Sizes are `~days`, `~weeks` and `~quarter`.*

---

## 1. The thesis: the bridge, and the line of abstractions behind it

The owner's sentence names a line, and the project already has most of it. Each abstraction below is *semantic* rather than syntactic: it names what is meant, not how the host does it, and a proof holds the meaning still (`SD-4`). What each one democratizes is the thing a person previously needed a specialist for.

| Abstraction | What a person writes | What it democratizes | Where it stands today |
|---|---|---|---|
| the sentence (the first gear) | `Put formula "=B2*2" into cell B3.` | automation: a macro with no VBA | shipped; 240 rules in `english.vla`, each with a proof |
| the question (the second gear) | `SQL`, `DATALOG`, `PROLOG`, `OPTIMIZE` in cells; a row ending in `?` to come | querying and logic without a query language in front of the table | shipped as worksheet functions; the standalone question is `AXM.9` |
| the control (the third gear) | `No cell in column G of sheet Invoices should contain "Unpaid".` | audit: the stated expectation, checked | `AXM.2` |
| reflection | `frazaro reflect model.xlsx`; `Which formula columns contain a typed-over constant?` | the workbook as data to its own owner | shipped as a reader over files; `AXM.8` inside the host |
| the library | `Use month_end_controls.txt with ledger: Table GL, bank: Table Bank, and tolerance: 10%.` | composition: a paragraph written once, used everywhere | `G-USE`, scoped in `HORIZON.md` §12.9 |
| the self-describing workbook | the `Frazaro` sheet, the `Frazaro.Build` stamp, `frazaro rebuild` | provenance: a file that proves its own build | shipped; §12.6's hidden Tables open |
| the live view | a grid that is a view of the sentences, and nothing else | the IDE: the spreadsheet as the editor of its own source | this file, sections 6 to 8 |

The web's lesson, which `HORIZON.md` §11.3 already draws, holds at every row: the person *writing* needs Frazaro, and the person *reading* needs only the spreadsheet program already on the machine. A bridge is crossed in both directions, and the far bank must need nothing installed. (*shelf*)

## 2. The rate, measured

What the tree says on 2026-10-05 (*measurement*; the method is in section 10):

| Measure | Value |
|---|---|
| Commits in this repository since its first, 2026-09-04 | 205 |
| Static checks in `tools/`, `0.7.1` to today | 30 to 47 |
| Conformance oracles in the treaty | 1 to 10 |
| Rust in `core/src/` and `cli/src/`, written 2026-10-01 to 10-05 | 29,484 lines |
| The VBA reference, `src/*.bas` | 99,045 lines |

Two honest things about that rate. The Rust number is a port: every slice had a golden to hit and OOXML's shape was given, so a day bought a slice. The writer and the reader were new ground, but the file format was still given. The engines of the Singularity line have neither a reference nor a format; they are design-bound and corpus-bound. And the roadmap names two constraints at this version, throughput and learning. Throughput was the binding one these two months. The grammar grows by sentences per refusal, and the refusals that matter are typed by people who did not write the grammar, so within a year the binding constraint is learning. Every forecast below is read against that.

## 3. A year, and ten

**October 2027** (*prediction*), in the order the road makes each cheap:

- **The third gear through Stage 3, in both implementations.** Controls, definitions, standalone questions, the audit list, diff and cause: `AXM.2` through `AXM.11`. The core already reads, diffs and audits; the VBA already has the grammar and the engines. The work is the seam.
- **The why-corpus, then the first `ABDUCE`.** Twenty explanations by hand (`AXM.12`), then the engine over `OPTIMIZE`'s search with clingo beside it. The bet is on the engine's first version, not on Level 4 of the Imitation Game.
- **The Imitation Game played at Level 2 or 3 on a fixture workbook.** Level 4 needs explanations an analyst agrees with, which needs closes the author did not write.
- **Recalculation over a declared subset of the formula language.** Section 7's third decision; the piece that lets the web page answer a question instead of printing a relation.
- **Libraries of sentences.** `G-USE`, the linker, shippable in the VBA host first.
- **The self-describing workbook completed.** The hidden Tables of `HORIZON.md` §12.6, which `BETA_REARVIEW.md` already decided are the reader's relations written at build time.
- **A second edition and new audiences.** Spanish through the EDITION line; the crates, the Store and GitHub Pages carrying the doors to people who install nothing.

**2036** (*prediction*), when the rate stops being one person's and becomes the language's:

- **The grammar's growth flattens.** Church's bequest in `HORIZON.md` §3: a few thousand rules, and the gap ledger's rate falling quarter on quarter. The finite close is the bet, and ten years is when it is won or lost.
- **A second implementation by someone else.** The treaty, the goldens and the proofs are the standard. The day another team's door passes `tools/prove.ps1`, Frazaro is a language in the sense SQL is. The suite is one person's work; the second team is not.
- **The estate, the record and merge.** One workbook to all of them; the past replayed under a new control; two analysts' edits merged by relation (section 5). All deterministic, all buildable alone.
- **The induction family, partly.** Definitions and controls induced from the record by enumeration and exact coverage (`HORIZON.md` §9.3). Induced grammar waits on a refusal corpus large enough to cluster.
- **One record outside finance.** The engines take any table; the first phrasebook is whichever rule-system a user already keeps in a grid.
- **Not alone.** The transactional host, proof-carrying commerce's adoption, the verified kernel. Each needs a host built around the engine, a counterparty, or a proof team.
- **The role.** `HORIZON.md` §6 already says the heaviest users of a deterministic spreadsheet will be machines. The ten-year product is the checker; the add-in is how it got its first thousand sentences. (*shelf*)

## 4. Reflect, and the Lisp in the machine

**What reflect is today**, so the claim stands on something: a reader from a file to S-expressions in a fixed order; six relations in the golden (`sheet`, `name`, `table`, `cell`, `formula`, `refers`) plus `changed` through `diff`; the six audit walks; counts and cones; two formats; three oracles. Reading a workbook is not new: Python libraries read formulas, Excel has had `FORMULATEXT` since 2013, and the end-user-computing tools walk precedents for a fee (*exhibit*). Two things are new. The output is a relation set in the engines' own notation, held to goldens, with every formula in R1C1 normal form relative to its cell and every unreadable reference refused by name. And the same core that reads also writes. That second fact is where the revolution is, and it is one arrow short.

**The missing arrow** is the one from relations back to a workbook. The writer takes sentences; the reader prints relations. Add a build that takes a relation file, and `reflect` of `build` of a relation set is that set, on the subset the writer covers. At that point the `.xlsx` is a print representation of an S-expression, and "the workbook is a list" stops being a figure of speech. In Lisp's terms the project has `read` and `print` and lacks the identity `(read (print x)) = x`. *Engineering*, small; the first slice to name if the rest of this section is to become literal. (*recommendation*)

**The reflective tower.** Brian Cantwell Smith's 3-Lisp (1982), Pattie Maes' computational reflection (1987) and Kiczales' metaobject protocol (1991) split reflection two ways. Introspection reads the system's own representation; intercession changes the system by changing it. Structural reflection is about the program; behavioural reflection is about its running. Frazaro sits in one cell of that grid today: reflect is structural introspection; the run log and provenance are behavioural introspection; build and rebuild are structural intercession over sentences but not yet over relations; `HORIZON.md`'s item 15, the keystroke refused, is behavioural intercession, and it is fiction until a host is built around the engine. (*exhibit* for the lineage; *shelf* for where the project stands)

**Causal connection.** The property Excel lacks and a Lisp has is what Maes named causal connection: the self-representation and the system are linked, so that changing one changes the other. Reflect today is a snapshot; edit a `(cell ...)` row and nothing happens. `CONTEMPLATIONS.md`'s second contemplation already asks how to make the sheet a view and not the truth. The tectonic form of that question: the relation set becomes the truth, the grid a materialized view of it, formulas become rules, an edit becomes an asserted fact, and the record is the extensional database with the grid as its intensional one. The spreadsheet becomes a deductive database that happens to print as a grid. *Math exists* (materialized Datalog views; DBSP, 2022, for keeping them incremental); *engineering*, large; the strain is the host, as `HORIZON.md` §9.4 names it.

**What expands** once the workbook is a relation set with causal connection, concretely:

1. **Controls over shape, not only values.** `No formula in column D of sheet Model should refer outside the sheet.` The audit list of `AXM.10` is a fixed list a vendor wrote; this is a grammar the controller writes. Meta-controls. *Engineering*: Stage 1's quantifier grammar over the reflect relations.
2. **Refactoring as rules.** Rename a sheet, move a column, replace forty formulas with one: each a rule over `formula` and `refers` that yields a new relation set, built, then diffed against the old to prove nothing else moved. `defmacro` for grids, with `macroexpand` as the diff. *Engineering*: the writer's mover (`shift_a1_references`) already shifts references.
3. **Alpha-equivalence and content addressing.** R1C1 already names a formula independent of its cell. Two workbooks are the same program when their normalized relation sets match up to sheet renaming. Hash every formula by its normal form, as Unison (2019) hashes every definition and as de Bruijn (1972) named variables by position, and an estate of four thousand models collapses to a few hundred distinct formula shapes, which is the size of the audit. External links by content hash never break when a file moves. *Engineering*: the normal form and the hash are built.
4. **Decompilation.** The inverse of build over a workbook nobody built: the shortest program in the published grammar whose build has these relations, by enumeration smallest-first and exact match, with what remains named as hand-made cells. Flash Fill (Gulwani, 2011) is deterministic enumerative synthesis and Excel has shipped it since 2013 (*exhibit*). The installed base becomes source code: *this workbook is twenty-three sentences and forty hand cells*. And the length is a quality metric, minimum description length (Rissanen, 1978), a practical Kolmogorov complexity. *Math exists*; *engineering*, large and partial by design.
5. **Metacircularity without `eval`.** The engines are cells, reflect sees cells, so `DATALOG` can query the `DATALOG` calls a workbook contains: which rules load where, which Tables feed which engine. The workbook audits its own logic programs. Nothing evaluates at run time. Frazaro is a Lisp with its reader, printer and expander kept and its `eval` deleted, which is the staged discipline of MetaML (Taha and Sheard, 1997) and of Scheme's `syntax-rules`. That is the honest "Lisp ex Machina": a two-stage Lisp whose second stage is the host's recalculation. (*shelf*: `SD-15` and the GENSYM veto are why there is no `eval`)
6. **Curry-Howard on the Controls sheet.** A control is a proposition, a green cell is an inhabited type, and the derivation provenance prints is the proof term. Reflect supplies the context the proof is checked in. One sentence, because the shelf has the rest.

## 5. Seven tectonic moves that are on no shelf

*Each was grepped for in `docs/` on 2026-10-05 and found absent, except where a row says otherwise. Each names its lineage and what in the repository it stands on (section 9).*

1. **Version control for spreadsheets.** `diff` exists; merge does not. Three-way merge over the relation set, with conflicts refused by name, and branches of the record instead of copies of the file. Darcs (2003) and its patch theory are the lineage. Excel's co-authoring is last-writer-wins at the cell, with no history and no branch (*exhibit*). *Math exists*; *engineering*. Stands on `core/src/reflect/diff.rs` and the R1C1 normal form.
2. **Multi-way constraints: a spreadsheet one can tell the answer.** A cell is a one-way function and `GOAL` is a one-way reversal. Sketchpad (1963), ThingLab (1979), Cassowary (2001) and Hyvönen and De Pascale's interval computations on the spreadsheet (1996) solve a relation in whichever direction is typed. The plug is a human doing this by hand and leaving no trace. `Cells D9 and B2 are tied by the formula in D9; type either.` *Math exists*; deterministic under a fixed propagation order; `OPTIMIZE` is the solver substrate.
3. **Falsification as a sentence.** QuickCheck (Claessen and Hughes, 2000) with a seed stated in the sentence, so it is a computation and not a sample: `Try ten thousand inputs within the ranges on sheet Assumptions against every control.` finds the counterexample and shrinks it to the smallest. Mutation testing (DeMillo, Lipton and Sayward, 1978): plant each of the fifteen explanation schemas' defects in a copy and watch which controls go red, so a control that never goes red is dead and the Controls sheet gets a coverage number. This is the one-year form of `HORIZON.md` §9.1's items 1 and 3, by search rather than proof. *Math exists*; *engineering*, small; the seeded dice are `SPITBALLS.md`'s Spitball 34.
4. **The sentence recorder.** Excel's macro recorder (1993) records VBA. This one records sentences in the published grammar: drag a column and the Frazaro tab writes `Move column F of sheet Sales before column B.` Direct manipulation (Hutchins, Hollan and Norman, 1985) transcribed into a language whose every sentence keeps its meaning, so the recording is a program and not a log; programming by demonstration (Cypher, 1993) with the demonstration's meaning fixed. The full round trip is a lens (Foster, Pierce and others, 2007; Sketch-n-Sketch, 2016): edit the output and the sentence updates. `BETA_REARVIEW.md` notes that bidirectionality this far is a separate decision, not bought (*shelf*). *Engineering* for the recorder; the lens is *math exists*, *engineering* hard. Stands on `VLA_EventSink.cls`, which already receives sheet changes, and on the renderer. Section 7 makes the recorder the Frazaro spreadsheet's formula bar.
5. **Modelling standards as executable files.** FAST, ICAEW's Twenty Principles and SMART are prose rules about model structure. As a Controls sheet over reflect, each becomes a library: `Use fast_standard.txt with model: this workbook.` Catala is the lineage `CONTEMPLATIONS.md`'s fourth contemplation names for legislation; this is the same move for the standards bodies, and `MARKETING.md` already names the crowd (*shelf*). *Engineering*, small; adoption *fiction*. "FAST-compliant" becomes a green column instead of a certificate.
6. **An open semantics for the formula language, with the core as its reference.** Sestoft's Corecalc and Funcalc (2014), OASIS OpenFormula and ISO/IEC 29500 are the lineage. The roadmap names `run`; the tectonic framing is that a workbook's meaning becomes defined by a specification and a conformance suite rather than by one vendor's behaviour. The SQL of spreadsheets needs its PostgreSQL. It also unlocks interval arithmetic (Moore, 1966): `Cell D9 lies between -12 and 4,400 for every input within sheet Assumptions.` by exact bounds, no solver, with over-approximation named. *Math exists*; *engineering*, large. Section 7's third decision is its first step.
7. **A bill of materials for the estate.** `rebuild` is reproducible builds for one workbook. The estate's version records which workbooks were built from which libraries at which versions, and which carry a stale one; with content addressing (section 4, item 3) it yields the industry's first advisory for a formula: *this library's third version had a wrong tolerance, and these two hundred workbooks carry it*. Reproducible Builds (2013) and SLSA are the lineage. *Engineering*.

**One wall question**, named so it is not a surprise: FP&A lives on sensitivity analysis and Monte Carlo. The deterministic answer is exhaustive enumeration over a stated grid of inputs (a tornado chart as a sentence) and a seeded sample labelled as a sample, with the interpretation left to the person. Where that line falls is the owner's.

## 6. The Frazaro spreadsheet: why the page is already Elm, and where it is not

Elm's architecture (Czaplicki, 2012; the grandparent is Smalltalk-80's model-view-controller, Reenskaug, 1979) is four things: a model, messages, a pure update, and a view that is a function of the model and nothing else. The web page has three of them today (*shelf*: `web/README.md`, `web/index.template.html`). The model is the rows of sentences plus the picked phrasebooks, the only state the page keeps. A message is an edit to a row. The update is translate and build, pure, bytes in and bytes out, in the core, run on every change after a 120 ms pause. The view is where it stops: VLA text, VBA text, a hash and a download button, and no grid.

| Elm | The Frazaro spreadsheet |
|---|---|
| Model | the program text, the phrasebooks, the folder of libraries and data |
| Msg | a sentence typed, or a hand edit the grid turns into a sentence |
| update | translate, then build the sheet model; pure, in the core |
| view | a window of cells with values, formulas, styles, and the row that wrote each |
| Cmd | download the `.xlsx`; read a picked file or folder |
| Sub | none: no clock, no socket; `SD-13` as a type |
| no runtime exceptions | refusals in words at Check, with the row |
| the time-travelling debugger | Spitball 1 for free: the view of the program cut at sentence N |

The discipline that matters most in Elm is that nothing edits the view. That is the decision that makes "the spreadsheet is only a view" true instead of a slogan. The formula bar and the grid become message sources that emit sentences: type a number into a cell and the program gains `Put 5 into cell B2.` at its end; edit a formula in the bar and it gains a `Put formula` sentence. The program stays the single source, `rebuild` still says yes, and the append-only record of `HORIZON.md` §9.2 (item 7) arrives at no cost, because the program is a log of acts. The in-cell formula bar is not removed; it is demoted, in the owner's words, to "hand modification of macro output", and every hand modification is itself a sentence.

The missing value view is a deliberate absence. `HORIZON.md` §12.2 says the core contains no calc engine on purpose, and foresees the amendment: "an open engine behind a declared subset that refuses by name". That decision was right for the generator and is the one this product amends (section 7, decision 3).

## 7. The shape, and the decisions

```
  left pane: the controller                     right pane: the view
  the program, one sentence a row;              a viewport over the sheet model:
  a Use line's expansion beneath it;            values, formulas, styles, selection,
  refusals and answers in column C              precedents lit from refers, a formula bar
  (today's Frazaro sheet, grown into an editor) that writes sentences, never cells
                   |                                          ^
                   v                                          |
          frazaro-core, pure, wasm32, empty import section     |
          translate -> build the Workbook struct -> recalc over the declared subset
          -> the view record for one window --------------------+
          -> frazaro_build_xlsx on Download only
          -> reflect / audit / diff / cone over the Workbook in memory
```

Closing the left pane leaves the grid full-screen, as Excel and Sheets appear today. Opening it is "view source". The room is `HORIZON.md` §11.4's, unchanged: the sentence in column B, its result in column C; the left pane *is* the Frazaro sheet, and the right pane is the rest of the workbook.

The decisions, each with what it keeps (*recommendation* throughout, with the measurements that ground each):

1. **The core stays a pure function over bytes.** No model handle in the module's memory. A view call takes the program, the books, a sheet and a window, and answers with that window's cells. Retranslating a 10,000-line program costs 105 ms natively (*measurement*, section 10), so purity is affordable, and the page already works this way. A stateful handle is the escape hatch, taken only when a measurement at real sizes says so. *Amended 2026-10-09 for an engine's door alone (`KERNEL.22`; `Alonzo/SPEC.md` §4.1): the measurement came, the view record of the 10,000-line program printing in 812 to 893 ms natively, and a frame of 64,000 cells thirty times a second cannot rebuild its grid from text; so the language's machine (`vla_lang::machine`) keeps a grid under a handle for an engine's four calls, load, write, step and view. Frazaro's own doors, the page and the command line, keep the pure function.*
2. **The view is drawn from the sheet model, never from the `.xlsx` bytes.** Zip and XML are for Download. Of the 363 ms a 10,000-line build costs natively, the model's share is what the view pays.
3. **Values come from a recalculation engine over a declared subset.** *Amends `HORIZON.md` §12.2, "What the core does not contain, on purpose: a calc engine", 2026-10-05.* Its oracle exists today: the cached values Excel saved into `scripts/reflect/saved.xlsx`, and every fixture the owner saves in Excel after it. The subset is chosen by measurement, not taste: in the Enron corpus, 76% of spreadsheets use the same 15 functions and only 134 functions appear at all (Hermans and Murphy-Hill, 2015; *exhibit*). A formula outside the subset shows its text and `not computed here`, with the function named, which is the honest static label §12.6 already requires. *Math exists* (ISO/IEC 29500 and OpenFormula publish the semantics); *engineering*, and the long pole.
4. **Every edit is a sentence, in the published grammar.** `SD-4` holds on the recorded sentence. A later sentence over a cell an earlier one wrote is shown as an override, which is what a plug is, and `AXM.2`'s controls can refuse it.
5. **The grid is a viewport, not an Excel clone.** *Amends `HORIZON.md` §12.2, "a full grid is a commodity to adopt if a door ever needs one, never to build", 2026-10-05, to: a viewport is small enough to write; a grid is a commodity.* Scroll, select, resize columns, freeze headers, a formula bar, precedents and dependents lit from the cone walk. No fill handle, no drag-move, no in-cell editing beyond the bar; those belong to the Excel door (decision 9), where Excel is the view. If a measurement ever wants the commodity: Univer is Apache-2.0 and could be inlined; HyperFormula is GPLv3 or commercial and could not (*exhibit*, section 11).
6. **"Nothing executed" stays true, with one sentence amended.** The module computes formulas as a calculator does and runs no code of the user's. `tools/check_web_offline.ps1` is unchanged, and recalculation needs no imports, so `tools/check_core_imports.ps1`'s zero stays zero.
7. **`SD-18` stays, and the treaty gains an oracle.** VBA stays the reference for the language. Recalculation has no VBA reference, so its oracle is Excel's own saved values, as the writer's oracle was the owner opening the file; the amendment to `conformance/README.md` names it when the slice lands.
8. **The project folder is local and never fetched.** In Chromium the File System Access API reads and writes a folder the user picked; elsewhere a folder input reads and Download writes; the Tauri door has real files. Libraries resolve beside the program, as §12.9 says, and a data file is read by the reader into Tables. No `fetch` anywhere, which the offline check already forbids.
9. **Two renderings of one view.** The page's viewport, and Excel itself through the Office.js task pane of `PORT.11`, where the left pane is Frazaro's and the grid is Excel's. That is the owner's "VS Code with Excel as the resulting visualization", taken literally, and it needs no viewport at all.

## 8. The slices, each with its oracle

| # | Slice | Ships as | Oracle | Size |
|---|---|---|---|---|
| 1 | **The view record and the viewport.** A view call over the built sheet model; a canvas viewport; clicking a cell selects its sentence, selecting a sentence lights its cells | the page, two panes, values for literals and folded formulas | a view golden per build fixture (`frazaro view fixture.txt --sheet Model --window A1:F20`), and the model's relations equal to the built file's, a free oracle | `~weeks` |
| 2 | **Recalculation over the declared subset.** A parser for operators and calls over the reference scanner that exists; a dependency graph from `refers`; topological evaluation; cycles and the unreadable refused by name; error values; the Enron 15, then the owner's list, then spills (`SEQUENCE` and `FILTER` are already in the fixtures) | live values; `frazaro calc <file.xlsx>` printing computed against cached | Excel's saved values; a `check_recalc_golden.ps1` with floors; a function enters the subset with its fixture | `~weeks` to the 15 with errors; `~quarter` with dates, text and spills |
| 3 | **The hand edit as a sentence.** The formula bar and the grid emit `Put` sentences; overrides shown | the demoted formula bar | the round trip: the recorded sentence rebuilds the same cell; the build goldens unchanged | `~days` |
| 4 | **Reflection over the model, in the page.** `reflect`, `audit`, `diff` and the cone over the Workbook in memory; the hidden Tables of §12.6 written by the build | precedents lit; the audit of the program being typed | reflect of the model equals reflect of its file | `~days` |
| 5 | **The engines in the core and the question act.** `PORT.9` and `AXM.9`: a row ending in `?` answers in column C from the live model; `AXM.2`'s `should` and `must` colour cells | the Q&A and Controls sheets as the left pane | `scripts/proofs/datalog.vla` with clingo; the Imitation corpus | `~quarter`, on the roadmap |
| 6 | **The interpreter over the sheet model.** `PORT.10`: `Make cell A1 bold.` stops being refused by the writer; formats, inserts, deletes and sorts happen in memory | the first gear, live | the interpreter golden; `VerifyReports`' lines | `~weeks`, on the roadmap |
| 7 | **The project and the editor.** `G-USE` libraries with the expansion shown beneath the `Use` line; data files as Tables; one document with per-row status, the gap ledger inline, completion from the pattern table (the patterns are data; `DidYouMean` exists); virtualized rows, since ten thousand input elements is the wrong structure | the IDE half of the picture | the build golden of a program using a library; the stamp hashing every file used | `~weeks` for the loader, `~weeks` for the editor |
| 8 | **The doors.** `PORT.11`: Tauri with real files and a signed installer; the Office.js task pane where Excel is the grid | one per audience | the same suite through each | `~days` each |

**The order.** Slices 1 to 4 are the Frazaro spreadsheet: a grid showing the values of a 10,000-sentence program, with a hand edit becoming a sentence, in a few weeks plus the recalculation's first fortnight. Slices 5 and 6 are already filed and make the pane answer and act. Slice 7 is the IDE. Slice 8 is distribution. At the measured rate the whole is about three quarters of one person's work, with two long poles, recalculation and the engines, one of which is already scheduled. Slices 1 to 4 come first because they put the whole picture on a screen, and nothing in the later slices is blocked by doing them in that order. (*prediction*; *recommendation*)

**The risk ledger.**

- **Excel's semantics.** `LAMBDA` and `LET`, implicit intersection, date serials, text-to-number coercion, locale formats. Each is refused by name until its fixture lands; the oracle is Excel's own values, so no guess ever ships as a value.
- **Memory at scale.** The Workbook is an ordered map of cells. A million-cell data sheet in the module's memory is a measurement item before it is a promise, as `AXM.13` is for the reader.
- **The browser's files.** Folder access is Chromium's today; Firefox and Safari read and download. The Tauri door exists for exactly this.
- **Expectations.** Users will ask for Excel's grid. The answer is the Excel door, not a clone, and the viewport's job is to stay small.
- **Size.** Recalculation adds a few hundred kilobytes of wasm (*prediction*). Still one file.
- **The page's claim.** `web/README.md` says the page "executes nothing". It stays true in the sense that matters, and the sentence is reworded when slice 2 lands, as decision 6 says.

## 9. Seeds: what each idea stands on today

*So the file can be checked against the code, as the house rules require. Each row names what exists at the time of writing, 2026-10-05; a later reader verifies before building on it.*

| Idea | Stands on today | Where |
|---|---|---|
| the view record (8.1) | every ABI answer is one record in the module's memory; the writer's walker knows the sentence that wrote each cell, since its refusals quote it | `core/src/abi.rs`, `core/src/build.rs` |
| recalculation (7.3, 8.2) | the reference scanner and the R1C1 renderer; the table of the format's future functions and of the functions that return an array; a host's cached values in a saved fixture; `fullCalcOnLoad` | `core/src/refers.rs`, `core/src/sheet/xlfn.rs`, `scripts/reflect/saved.xlsx`, `core/src/sheet/ooxml.rs` |
| the hand edit as a sentence (7.4, 8.3; 5.4) | the event sink receives sheet changes; the page keeps only the rows | `src/VLA_EventSink.cls`, `web/index.template.html` |
| reflection over the model (8.4) | the rear-view's decision that `_frazaro_refers` and `_frazaro_cells` are the reader's relations produced at build time over the writer's own model | `BETA_REARVIEW.md`, `PORT.8` slice 8e's scoping |
| the engines live (8.5) | the four engines in VBA; the question slot inside an imperative | `src/VLA_Datalog.bas` and kin; `english.vla`'s `{q:question}` |
| the interpreter live (8.6) | the Interpret backend and its bounded built-ins | `src/VLA_Interpreter.bas`, `TryEvalBuiltin` |
| libraries (8.7) | the design, with named slots, structural substitution and the visible expansion | `HORIZON.md` §12.9; `G-USE` |
| the missing arrow (4) | the reader prints relations; the writer takes sentences; the sheet model is one struct both could share | `core/src/reflect/`, `core/src/build.rs`, `core/src/sheet/mod.rs` |
| content addressing (4.3) | the R1C1 normal form; SHA-256 from nothing | `core/src/refers.rs`, `core/src/sha256.rs` |
| decompilation (4.4) | the writer's static subset is the target language; the build goldens are the first test pairs | `core/src/build.rs`, `scripts/build/` |
| merge (5.1) | two files compared one sheet pair at a time; a model's parts copied byte for byte around new sheets | `core/src/reflect/diff.rs`, `core/src/sheet/merge.rs` |
| constraints (5.2) | a search with an objective over grounded atoms | `src/VLA_OptimizeSearch.bas`; `SINGULARITY.md`'s `GOAL` |
| falsification (5.3) | the planted-defect fixtures of the explanation schemas; seeded dice as a spitball | `SINGULARITY.md` §8.3; `SPITBALLS.md` Spitball 34 |
| standards as files (5.5) | the audit list's six walks; the library design | `core/src/reflect/audit.rs`; `HORIZON.md` §12.9 |
| the open semantics (5.6) | `run` named as not in this version; the future-function table | `cli/README.md`; `core/src/sheet/xlfn.rs` |
| the bill of materials (5.7) | the stamp hashes the program, the prelude and every phrasebook; `rebuild` compares whole | `core/src/build.rs` (`Stamp`) |

## 10. The numbers, and how to re-run them

*Measured 2026-10-05 on the owner's Windows machine; the release door was built that evening with `cargo build --release -p frazaro`. Native timings; the browser's module runs slower by a factor the first slice measures.*

- Commits: `git log --oneline | wc -l` (205); the first commit's date: `git log --format=%ad --date=short | tail -1` (2026-09-04).
- Checks: `(Get-ChildItem tools/check_*.ps1).Count` (47); the count at `0.7.1` is `SINGULARITY.md` §3's (30).
- Oracles: the amendments of `conformance/README.md` (oracle 10 is the audit golden, 2026-10-05).
- Lines: `find core/src cli/src -name '*.rs' | xargs cat | wc -l` (29,484); `cat src/*.bas | wc -l` (99,045).
- Translate the corpus, 1,444 lines: `target/release/frazaro translate-vla scripts/instructions.txt --prelude scripts/prelude.vla --phrasebook scripts/polyglotta/english.vla`, timed in a shell: 68 ms on the second run (124 ms cold).
- A synthetic program of N lines, half values and half formulas, written to a scratch file:

  ```
  awk -v n=10000 'BEGIN{for(i=1;i<=n;i++){ if(i%2) printf "Put %d into cell A%d.\n", i, i; else printf "Put formula \"=A%d*2\" into cell B%d.\n", i-1, i }}' > big_10000.txt
  ```

  Translate: 66 ms at 1,000 lines, 105 ms at 10,000. Build to `.xlsx` with `--out ... --replace`: 91 ms at 1,000, 363 ms at 10,000. `frazaro reflect` of the 10,000-line build: 35,003 rows in 357 ms.
- The module: `ls -l target/wasm32-unknown-unknown/release/frazaro_core.wasm` (958,023 bytes); the page: `ls -l web/index.html` (1,618,530 bytes).
- The page's re-run delay: `schedule()` in `web/index.template.html`, 120 ms.
- The Enron figures: Hermans and Murphy-Hill, "Enron versus EUSES: A Comparison of Two Spreadsheet Corpora", 2015 (section 11).

## 11. Sources for the outside facts

- Hermans, F. and Murphy-Hill, E., "Enron versus EUSES: A Comparison of Two Spreadsheet Corpora", 2015. https://arxiv.org/pdf/1503.04055
- Czaplicki, E., Elm, 2012; the Elm Architecture is described in the language's guide. https://guide.elm-lang.org/architecture/
- Reenskaug, T., the model-view-controller notes, Xerox PARC, 1979.
- Smith, B. C., "Reflection and Semantics in a Procedural Language", MIT, 1982. Maes, P., "Concepts and Experiments in Computational Reflection", OOPSLA, 1987. Kiczales, G., des Rivières, J. and Bobrow, D. G., *The Art of the Metaobject Protocol*, 1991.
- Taha, W. and Sheard, T., "Multi-Stage Programming with Explicit Annotations" (MetaML), 1997.
- Budiu, M. and others, "DBSP: Automatic Incremental View Maintenance for Rich Query Languages", 2022.
- Sutherland, I., *Sketchpad*, 1963. Borning, A., *ThingLab*, 1979. Badros, G., Borning, A. and Stuckey, P., "The Cassowary Linear Arithmetic Constraint Solving Algorithm", 2001. Hyvönen, E. and De Pascale, S., "Interval Computations on the Spreadsheet", 1996.
- Claessen, K. and Hughes, J., "QuickCheck", ICFP, 2000. DeMillo, R., Lipton, R. and Sayward, F., "Hints on Test Data Selection", 1978.
- Hutchins, E., Hollan, J. and Norman, D., "Direct Manipulation Interfaces", 1985. Cypher, A. (ed.), *Watch What I Do: Programming by Demonstration*, 1993. Foster, J. N., Greenwald, M., Moore, J., Pierce, B. and Schmitt, A., "Combinators for Bidirectional Tree Transformations", 2007. Chugh, R. and others, *Sketch-n-Sketch*, 2016.
- Roundy, D., Darcs, 2003. Unison, the content-addressed language, 2019. de Bruijn, N. G., "Lambda Calculus Notation with Nameless Dummies", 1972.
- Gulwani, S., "Automating String Processing in Spreadsheets Using Input-Output Examples", POPL, 2011. Rissanen, J., "Modeling by Shortest Data Description", 1978.
- Sestoft, P., *Spreadsheet Implementation Technology*, MIT Press, 2014. OASIS, OpenFormula (ODF 1.2 part 2), 2011. ISO/IEC 29500. Moore, R. E., *Interval Analysis*, 1966.
- The Reproducible Builds project, 2013 onward; SLSA, the supply-chain levels for software artifacts.
- Merigoux, D. and others, Catala, 2021.
- IronCalc, a Rust spreadsheet engine, MIT/Apache-2.0. https://github.com/ironcalc/IronCalc
- Univer, a TypeScript spreadsheet SDK, Apache-2.0. https://github.com/dream-num/univer
- HyperFormula, licensing: GPLv3 or commercial. https://hyperformula.handsontable.com/docs/guide/licensing.html

---

*Appended 2026-10-06, under the house rules above. The owner's fourth question of the sitting, asked the next day: how much human drudgery the spreadsheet scoped here would replace; what libraries of sentences, self-reflecting workbooks and proofs all the way down would bring to industry that ends whole classes of the headaches Excel-as-is causes; and, in one sentence, to sell the roadmap to its own ideator, not shyly, showing the corporate world if Frazaro were completed and adopted as scoped, with thousands of phrasebook authors and library writers, dozens of languages, and a WebAssembly interface that calculates faster than Excel once the optimizations land. The owner asked that the entry take the optimistic view of corporate drudgery's future, and gave it its epigraph. Marks as above: every figure in 12.2 is an* **exhibit** *with its source beside it; the tables of 12.3 and the quarter-end of 12.5 are* **predictions** *composed from scoped items, as the Mondays of `SINGULARITY.md` and `HORIZON.md` are; the sentence of 12.1 is a pitch and says so.*

> *Just because analysts of the past clicked, typed, and suffered need not mean that analysts of the future must needs suffer the self-same fate.* (the owner, 2026-10-06)

## 12. The future of corporate drudgery, optimistically: the roadmap pitched to its ideator

### 12.1 The sentence

**The spreadsheet is the only programming medium on Earth whose programs have no source code, and Frazaro gives them one: a page of sentences anyone in the company can read, in their own language, each proven to do what it says, so that every hour finance spends re-checking, re-explaining and rebuilding what the grid forgot becomes a question the workbook answers about itself.**

That is the strike at the heart, and it is why the roadmap is not an increment. VisiCalc's leap was dependency (section 2 of `HORIZON.md`): a cell is a named expression, and the sheet is one enormous recursive binding. It stopped there. The grid is the program, and the program has no source, no types, no tests, no history, and no reader but the person who typed it. Every headache in 12.3 is a symptom of that one missing thing. The three abstractions the owner named are the three forms the supply takes. Libraries are the source made reusable. Reflection is the source made readable by the program itself. Proofs all the way down are the source made trustworthy by someone who did not write it.

### 12.2 The exhibits

| Exhibit | Figure | Source |
|---|---|---|
| Audited spreadsheets with at least one error, Panko's syntheses | 88% to 94% | Panko, "What We Know About Spreadsheet Errors" and later syntheses; https://arxiv.org/pdf/0801.1514.pdf |
| Formula cells in error, across seven rigorous audits | about 1% to 5% | Campbell summarizing Panko, Society of Actuaries, 2010; https://www.soa.org/globalassets/assets/library/journals/actuarial-practice-forum/2010/february/apf-2010-02-campbell.pdf |
| Monthly close, median over 3,303 companies | 8 days | APQC Open Standards Benchmarking; https://www.apqc.org/what-we-do/benchmarking/open-standards-benchmarking/measures/cycle-time-days-finance-shared |
| Monthly close, top performers against the bottom quartile | 4.8 days against 10 or more | CFO.com, "Metric of the Month: Cycle Time for Monthly Close"; https://www.cfo.com/news/metric-of-the-month-cycle-time-for-monthly-close/ |
| JPMorgan's Chief Investment Office VaR model, 2012: copy-paste between spreadsheets, a SUM where an AVERAGE belonged | $6.2 billion | the public reports collected by EuSpRIG; https://handsontable.com/blog/?p=2566 |
| Public Health England, 2020: cases lost to the 65,536-row limit of an old file format | 15,841 | https://ia.acs.org.au/article/2020/excel-causes-uk-covid-confusion.html |
| Norway's sovereign wealth fund, 2023: one date typed a month off | $92 million | https://it.slashdot.org/story/24/02/12/211251/the-norwegian-sovereign-wealth-funds-92-million-excel-error |
| People who use Excel | about 750 million | Statista; https://www.statista.com/statistics/983321/worldwide-microsoft-excel-users/ |

### 12.3 The headaches, by class

| Today | What ends it | Which abstraction |
|---|---|---|
| **The close as checking.** Tie-outs by eye, reconciliations by lookup and highlight, redone every month by hand | The Controls sheet runs at open. A reconciliation is one sentence. The yellow rows are the whole morning's job | controls, standing orders (`AXM.2`; `HORIZON.md` §9.4 item 16) |
| **The plug and the typed-over constant**, discovered at audit or never | `No formula column should contain a typed constant.` runs every morning. A plug is a sentence with an expiry, red the day it lapses | controls, reflect (`AXM.10`; `HORIZON.md` §5 item 1) |
| **Why did B9 change.** A morning tracing precedents across sheets with a highlighter | Cause, variance and abduction answer in three sentences in column C, with the prediction that tested them | reflect, the engines (`AXM.11`, `AXM.12`; `SINGULARITY.md` §8) |
| **Copy-paste as the integration layer** between workbooks, which is how the Whale swam | There is no copy. A model is built from sentences, a paragraph is a `Use` line, and a table arrives as a file carrying its own derivation | libraries, build (`G-USE`; `HORIZON.md` §5 item 2) |
| **Version chaos**, the file named final twice | The program is the source. Diff, merge and rebuild work on relations, and a workbook proves which sentences built it | the self-describing workbook (`HORIZON.md` §12.6; section 5 item 1 here) |
| **The model nobody understands** after its author leaves | View source. The model rendered as a book, inputs first with their provenance, every formula in words. Decompilation for the legacy estate | reflect, render, decompile (`AXM.3`; `HORIZON.md` §9.5 item 22; section 4 item 4 here) |
| **Silent truncation and type traps**: the row limit, the wrong month, text that looks like a number | A sentence the sheet cannot hold is refused by name before anything runs. A dimension is a type, and September cannot be added to a dollar | refusals, units (`SD-2`; `HORIZON.md` §5 item 5) |
| **Model-risk compliance by inventory and screenshot** | The record, the stamp, the audit list and rebuild produce the evidence by the act itself | the record, reflect (`AXM.4`; `HORIZON.md` §12.6) |
| **English as the prerequisite for logic** | Dozens of phrasebooks, one meaning. The Tokyo controller's controls run unchanged in São Paulo, and the goldens hold it | editions (the EDITION line; `HORIZON.md` §3, Quine) |
| **Learning Excel, recording VBA** | The sentence recorder. One's own clicks come back as sentences one can read, keep and prove | the recorder (section 5 item 4 here) |
| **The forty-second recalculation** and manual calculation mode | Incremental, vectorized, content-addressed evaluation, with values folded at build time | the core (12.6 below) |

### 12.4 Drudgery, counted honestly

No percentage for the drudgery removed can be sourced, so here is the structure instead. A close holds three kinds of work: waiting for data, checking and explaining, and deciding. Only the middle kind is drudgery, and it is exactly the kind that exists because the grid forgets what it meant. The prediction is that the gap in 12.2 between the top performers and the median is almost entirely the middle kind, and that the gap between the median and the bottom quartile is the same kind again. What remains afterwards is what no machine should do and what `HORIZON.md` §9.8 puts in neither set: deciding what should be true, writing the first sentence, and knowing the thing said at lunch.

### 12.5 A quarter-end, if it works

*Every sentence below is supplied by a scoped item; the whole is a prediction.*

Quarter-end, at a company that adopted it. Nothing closes on the first morning, because nothing opened. The record has been accepting acts all quarter and refusing the ones that would have turned a `must` red. At seven the standing order runs the Controls sheet, forty rows a controller wrote once in sentences, eleven of them from a library the modelling standards body publishes as a file. Thirty-eight are green. Row 14 is yellow and names the three bank rows that match no ledger row. The counterparty's proof of its side arrived as a file at twenty to seven, and nobody here has seen their ledger.

The analyst, who used to spend the first three days of every close reconciling by lookup, spends the morning on row 14 and on one question the manager typed: why did Summary's B9 rise? Column C answers in Peirce's three sentences and offers a fourth, a schema the catalogue lacked, which she adopts with her name on the line. The model is live under her hands. A sentence typed at a quarter past eight shows its consequence in the grid before she lifts her finger, because the engine recomputes only the cone that changed.

The bookkeeper in Lagos writes the same controls in Yoruba against the same engine, and the goldens hold that they mean what the Tokyo controller's mean. The new hire reads the model as a book and understands in an afternoon what used to take a quarter. The auditor in Lisbon disputes cell D17 from her own seat, and the model answers with the plug sentence, its expiry, and the bank line that will replace it. IT keeps no spreadsheet inventory, because every workbook in the estate carries its stamp, and "which models still use last year's tax rate" is a query answered in seconds rather than a survey answered in weeks. The CFO signs off at four minutes past nine against a hash a second person can re-derive, and every number in the board pack carries its derivation back to the bank line it came from. The quarter closes in the time it takes the data to arrive, because the checking was never a phase. It was the standing state of the record.

### 12.6 On speed, said plainly

Excel's recalculation engine is good, and the claim to beat it is not about raw arithmetic. It is about three things the sentence level knows and the grid forgets. A column of forty formulas is one sentence, so it compiles once and evaluates as a vector. Nothing that did not change is recomputed, because the dependency relation is data and the fixpoints are maintained incrementally (DBSP, 2022; section 4). Identical formula shapes are computed once by content address, across a model and across an estate (section 4, item 3). Add values folded at build time, so a workbook arrives already computed, and a view that evaluates only the window and its cone (section 7, decisions 1 and 2). Determinism is what makes every one of those sound, since a memo is only honest when the same inputs always give the same answer; that is the wall paying for the speed. "The fastest spreadsheet on the planet" is a *prediction* until the recalculation slice (section 8, slice 2) measures it, and it is a prediction the architecture earns.

### 12.7 The analyst of the future

The optimistic view, which the owner asked for and the evidence permits. The drudgery in 12.3 was never the job; it was the tax the medium levied on the job, and 12.2 is the receipt. The analyst of the future does what the analyst of the past was hired to do and rarely had the hours for: writes the controls, names the definitions, reads the model aloud to the board, asks why and gets an answer with its test attached, and spends the close on the two rows that were not green. `HORIZON.md` §5 item 9 calls that job lexicographer; `SINGULARITY.md` §2 calls it eight duties and gives seven of them to the machine. Either way the clicking is gone, the typing is one sentence where it was forty cells, and the suffering, which was never anyone's purpose, has no seat left. With the commons of thousands of phrasebook authors, the sentences arrive faster and in more languages; without it, the engine does the same work in fewer tongues, since every row of 12.3 is the core's and the commons only widens who can write. The abacus moved beads and the spreadsheet moved numbers. The thing scoped in this file moves the work itself, out of the hands and into the record, where it can be read, proven and kept.

*Sources for this section's figures are in the third column of 12.2; the rest of the file's sources are in section 11.*

---

*Appended 2026-10-06, under the house rules above. The owner read the sitting's assessment of the two-pane model, its strengths, its weaknesses and ten recommendations for the synthesis between spreadsheet and spreadsheeter, and blessed the ten: "I hereby bless your 10 syntheses as the inaugural 10 Frazaro Commandments, and should be blessed as the first standing decisions which guide Frazaro-as-spreadsheet into the unknowable future of spreadsheetery." They are minted as `SD-20` to `SD-29` in `BETA_ROADMAP.md`'s standing-decision register, with their reasons in `BETA_REARVIEW.md`; this section holds each one's text and lineage, in the order given, and after them the further commandments the sitting proposes, which stand as* **recommendations** *until the owner blesses, amends or declines them. The owner's frame for the whole: Frazaro as the Magellan of a New World in which the spreadsheet is an IDE, and citizen developers are promoted to developers de facto, because the main entry point to automating a spreadsheet is a sentence.*

## 13. The Commandments: the standing decisions of the Frazaro spreadsheet

### 13.1 Why a register of its own

The assessment that produced them, in two sentences. The two-pane model is viable and better supported by prior art than it looks: Calculation View, at Microsoft Research in 2018, put a textual pane beside Excel's grid with edits flowing both ways and found in a user study that it significantly improved authoring and debugging. Its one real danger is making the pane primary and the grid secondary, which would trade away the property Nardi and Miller found made spreadsheets win: program and output as one object, tinkered with locally, with immediate feedback. The commandments are the conditions under which the model keeps that property. They bind every door that shows a grid, the page's viewport and the Excel task pane alike, and they leave the language, the engines and `SD-1` to `SD-19` untouched.

### 13.2 The ten, blessed

1. **`SD-20`. The grid is the pointing device for the grammar.** Select a range, a column, a sheet or a Table, and the next sentence in the pane takes it as its noun, the slot filled as the grammar spells it. The grid supplies the nouns and the pane the verbs. *Lineage:* Bolt's "Put-That-There" (1980), pointing plus speech, and the language's first verb is Put. Spatial deixis to a visible selection is not anaphora, so `SD-16` holds.
2. **`SD-21`. Direct manipulation stays first-class, and every manipulation is transcribed.** A value typed, a formula edited, a column moved, a range formatted: each is accepted as the spreadsheet has always accepted it, and each becomes a sentence appended to the program. The program stays the single source without the grid becoming read-only. *Lineage:* Nardi and Miller (1990, 1991); Sketch-n-Sketch (2016); Excel's macro recorder (1993), which recorded VBA where this records sentences, with `SD-4` making the recording a program rather than a log.
3. **`SD-22`. The source is kept apart from the changes, and the changes condense.** Two layers: the declarative program and the log of transcribed edits. A condense act rewrites the log into the shortest program whose build has the same relations, overrides collapsed, shown as a diff before it is accepted. *Lineage:* Smalltalk-80's changes file and condense-changes; git's squash; section 4's decompiler pointed at one's own history. It also teaches: the author sees what their clicks meant.
4. **`SD-23`. One model, many projections.** The grid, the sentences, the dependency cone as a diagram, the model as a book, the controls dashboard and the diff are each a pure function of one model; none is the truth, and the pane is never the only view. *Lineage:* Calculation View (2018); Simonyi's intentional programming, one intention tree with many projections, from the author of Multiplan. Elm's single view is right for the engine and wrong for the interface, which also needs state the program never sees: selection, scroll, a half-typed edit.
5. **`SD-24`. The formula bar shows value, formula and the sentence that wrote the cell.** Three lines for the selected cell, the third naming the program's row; click it to jump. Editing the formula rewrites the `Put formula` sentence that wrote the cell, and that one shape is the whole backward arrow. *Lineage:* lenses (Foster, Pierce and others, 2007), restricted to one shape; the general lens stays a separate decision not bought.
6. **`SD-25`. A refusal lands in both panes.** The row turns red with the refusal in words, and the cells its sentence names carry a marker that leads back to the row. Evaluating the program's prefix, Spitball 1's time scrubber, is the debugger. *Lineage:* Green and Petre's cognitive dimensions (1996), where hidden dependencies and role-expressiveness are the spreadsheet's worst scores; a refusal visible in one pane only would hide its consequence from the person looking at the other.
7. **`SD-26`. Auditing and review are modes of the pane.** In audit mode the pane holds the findings and the controls, each linked to its cells; in review mode it holds the sentence diff of two versions, with the grid coloured by the relation diff. The auditor reads ten changed sentences instead of four thousand changed cells, which Excel cannot offer because a workbook has no source. *Lineage:* code review as practised over source control.
8. **`SD-27`. Naming may follow building.** `Name this sheet Model.` and `Call this range Sales.` may come after the sentences that built the thing named, and condensation rewrites the earlier sentences to use the name. *Lineage:* premature commitment, in Green and Petre's terms, is what a grammar pays when it demands the name before the layout exists; spreadsheet authors discover structure by building.
9. **`SD-28`. The grid comes first for readers.** The pane is closed by default for a reader, and an explain-this-cell act opens it at the sentence that wrote the cell. The split view is for authors and auditors. *Lineage:* two in five of the Enron corpus's spreadsheets hold no formula; readers outnumber writers, and progressive disclosure keeps the spreadsheet a spreadsheet for them.
10. **`SD-29`. Measure before building.** Before the viewport, or any later change to the interface, a small study in the Imitation Game's spirit runs construction and audit tasks on the existing page beside Excel against Excel alone, with time, errors and comprehension questions in Hendry and Green's shape (1994): choose a cell, explain its formula, find its inputs. *Lineage:* the house rule for engines (`AXM.1`, `OPTIMIZE.0.1`), owed to the interface too; the viewport is weeks and the study is days.

### 13.3 Proposed, not yet blessed

*Recommendations, each a sentence and its reason. The owner blesses, amends or declines; a blessed one takes the next number in the register and moves up into 13.2 by a dated line.*

- **Nothing fails silently, and nothing succeeds invisibly.** Every act either shows its effect on the grid or refuses in words, and a clean result is a sentence, never an empty cell. *Reason:* `SD-2` governs the form of a refusal; this governs the presence of an answer, and `LX.15` already made a clean lint report a sentence rather than silence.
- **Speak in the user's names.** The recorder and the condense act prefer a Table, a defined name or a header to an address whenever one exists, so a transcribed sentence reads as the author would have written it. *Reason:* addresses are the machine's names and headers are the company's, and a program in the company's names survives a moved column. *Lineage:* Codd's relations named by attribute; Excel's structured references.
- **Appearance is declared, not painted.** Formatting sentences live apart from logic sentences in the program, and a palette is a library. *Reason:* `HORIZON.md` §11.3 already reads definitions and controls as the CSS of the workbook; keeping the look in its own paragraph keeps the logic legible and the look reusable. *Lineage:* `SPITBALLS.md`'s Spitballs 5 and 6, palettes as header files.
- **No blank page.** The first thing a newcomer sees is a sentence already doing something in a cell; the empty workbook is never the first screen. *Reason:* people learn a language from an example that works, the page opens that way today, and the spreadsheet should too. *Lineage:* Papert's Logo; the onboarding samples of `LE.6`.
- **The grid never freezes.** Every engine is interruptible and, when stopped or still working, says how far it got in a sentence. *Reason:* `SINGULARITY.md` §9 already makes every engine anchored, budgeted and interruptible; a live grid turns that from an engine property into a promise the person can see.
- **The program is the accessible workbook.** Whatever the grid shows, the sentences say, so a keyboard and a screen reader reach everything the mouse does, and a blind analyst reads the model as prose. *Reason:* a grid is hostile to a screen reader and a program of sentences is not; the two-pane model makes accessibility a consequence of the architecture rather than a feature. *Pays into:* the ACCESSIBILITY department.
- **No lock-in by construction.** Nothing the spreadsheet makes needs the spreadsheet to open: the built workbook runs in any host with no add-in, and view-source is for those who have Frazaro. *Reason:* `HORIZON.md` §12.6 claims it for the generator, and the spreadsheet must keep the claim as it grows, since a live tool is always tempted toward a private format.
- **A refusal is a lesson.** It names what was meant, what was wrong, and the nearest sentence that works. *Reason:* `SINGULARITY.md` §2 counts teaching among the analyst's eight duties, "a refusal with directions is teaching", and `DidYouMean` already computes the nearest rules; the commandment makes the lesson owed rather than optional.

*The sitting's note on the frame.* Magellan did not finish the voyage; Elcano brought the Victoria home. The conformance treaty, the goldens and these registers exist so that the standard outlives any one captain, which `CONTINUITY.md` asks for in other words.

### 13.4 Blessed, later the same day

*Appended 2026-10-06. The owner blessed the eight proposals of 13.3 ("Sure, why not: I bless the additional 8."). They are `SD-30` to `SD-37` in the register, in 13.3's order: `SD-30` nothing fails silently, and nothing succeeds invisibly; `SD-31` speak in the user's names; `SD-32` appearance is declared, not painted; `SD-33` no blank page; `SD-34` the grid never freezes; `SD-35` the program is the accessible workbook; `SD-36` no lock-in by construction; `SD-37` a refusal is a lesson. Their reasons are in `BETA_REARVIEW.md`'s register. 13.3 stands as written, under the house rules; the Frazaro spreadsheet now has eighteen commandments, and the next proposal takes `SD-38`.*

---

*Appended 2026-10-06, under the house rules above. The owner asked for "a comprehensive KERNEL.* series of steps to the roadmap to guide the development of CALLOSUM (working title of the WASM Frazaro spreadsheet) from idea to reality", after the sitting's answer to how a minimum viable two-pane spreadsheet should be constructed so that extensibility is the core primitive: a kernel holding mechanism only, five seams as the only entrances, and a distro as a folder the build tools bake into one artifact per door. The steps were filed on `BETA_ROADMAP.md` under a new line, 🔧🪟🌍 MACHINE + PRODUCT + COMMONS · THE KERNEL LINE, family `KERNEL`, twenty-one items; the reasoning, the fan-out table and the two tracks are in `BETA_REARVIEW.md` under the same heading, and the family is declared in `ID_REGISTRY.md`. Handles here stay handles; a filed step carries its ID.*

## 14. Filed, 2026-10-06: the KERNEL line

### 14.1 The kernel and its seams, in the house's terms

The kernel holds mechanism only: forms and their expansion, the emitters, the sheet model, recalculation over a declared subset, the relation set and the ABI. It holds no English, no message text, no default, no chrome, no format beyond a trait and no door; all of those are data or implementations of a seam. Five seams are the only entrances, each a trait or a data file with one implementation and an oracle today: sentences (a phrasebook with proofs; `frazaro prove`), paragraphs (a library, `G-USE`; the build golden and the stamp), engines (tables in, a table out, a head-table row, a proof-file kind; proof files with clingo beside them), formats and hosts (the `Source` trait, a `Sink` beside it, a host profile per door; reflect and build goldens), and projections (a pure function from model and window to a record; a view golden). The rule that follows: the kernel grows a seam, never a feature. A distro is a folder of phrasebooks, libraries, examples, a palette, an edition's chrome, a README and default `Use` lines, baked by the existing build tools into one page, one add-in, one command-line door, one desktop shell; the registry is a repository the user clones, as CTAN is for TeX, since nothing is fetched (`SD-13`), and the stamp's hashes are the lockfile. In Linux's terms: the head table is the syscall table, `SD-4` is "we don't break userspace", the treaty is POSIX, `OUTPUT-EXCEPTION.md` is the syscall note, a library is a userland package, a phrasebook with raw forms or an engine is a signed module, and the consent gate, `SD-15` and provenance are the signing. What is not built, and why: a runtime plugin loader (`SD-15`), a network registry (`SD-13`), a scripting language inside the language (§12.9's brevity, never power), and a second host before a second host exists (`SD-11`).

### 14.2 The map from this file to the items

| Here | Filed as |
|---|---|
| the kernel's boundary and the five seams (14.1) | `KERNEL.1` |
| the distro folder (14.1; §7 decision 8's folder, generalized) | `KERNEL.2` |
| `SD-29`, the study before the viewport | `KERNEL.3` |
| §8 slice 1, the view record (§7 decisions 1 and 2) | `KERNEL.4` |
| §8 slice 1, the viewport (§7 decision 5; `SD-20`, `SD-24`, `SD-25`, `SD-28`) | `KERNEL.5` |
| §8 slice 2, recalculation: the subset chosen by measurement (§7 decision 3's Enron figure) | `KERNEL.6` |
| §8 slice 2, recalculation's floor (§7 decisions 3, 6 and 7; the treaty's next oracle) | `KERNEL.7` |
| §8 slice 2, dates, text, lookups and spills | `KERNEL.8` |
| §8 slice 3, the hand edit as a sentence (`SD-21`, `SD-31`) | `KERNEL.9` |
| `SD-22` and `SD-27`, condense and late naming; §4 item 4's search over one's own history | `KERNEL.10` |
| §8 slice 4, reflection over the model and §12.6's hidden Tables | `KERNEL.11` |
| `SD-23`, projections as a seam | `KERNEL.12` |
| `SD-26`, audit and review modes; §5 item 1's review half | `KERNEL.13` |
| §8 slice 7, the project folder and the editor (`SD-33`, `SD-37`) | `KERNEL.14` |
| §8 slice 5, the engines live (`SD-34`), after `PORT.9`, `AXM.2`, `AXM.9` | `KERNEL.15` |
| §8 slice 6, the interpreter live, after `PORT.10` | `KERNEL.16` |
| `SD-35`, accessibility | `KERNEL.17` |
| describe, the kernel's `lsmod` (14.1) | `KERNEL.18` |
| §8 slice 8 and §7 decision 9, the pane in the other doors, after `PORT.11` | `KERNEL.19` |
| §12.6, speed said plainly, measured | `KERNEL.20` |
| §5 item 1, merge by relation | `KERNEL.21` |

Not filed, still handles: §4's missing arrow (a build from a relation file), §4 item 3's content addressing beyond what `KERNEL.20` needs, §4 item 4's decompilation of a workbook nobody built, §5 items 2 to 7 (constraints, falsification, the sentence recorder's general lens, standards as files, the open semantics beyond the declared subset, the bill of materials), and `HORIZON.md` §9.1's for-all-inputs line. Each waits for the measurement `KERNEL.8` makes possible, or for a sentence that needs it (`SD-7`). The eighteen commandments are standing decisions, not items, and bind the line from above.

---

*Appended 2026-10-08, under the house rules above. `KERNEL.3`, the study before the viewport (`SD-29`, §13.2 item 10; §14.2), is built as an instrument: `docs/PROTOCOL.md` holds the fixture with its five planted defects in two variants (`scripts/study/`, written by `tools/build_study_fixture.ps1`), the task sheet in two halves, the protocol that runs them on this page beside Excel against Excel alone, the renderer benchmark (`tools/bench_view.ps1`, two throwaway renderers of a real view record of the 10,000-line program of §10, canvas against a virtualized DOM, at a normal window and a full-screen one), the decision rule written before the first run (`PROTOCOL.md` §8: canvas or DOM by the 16 ms frame and crispness; the formula bar's first line and the pane's default for an author answered provisionally, with the measurement named that would overturn each), and the result table, blank. The running is the owner's. **This paragraph is the slot:** when the owner has run the study and the benchmark, the filled tables of `PROTOCOL.md` §9.2 and §9.3 are appended here as a dated measurement, with the machine, the browser and the door named, and §9.1's pilot rows beside them marked as the rehearsal they are; the renderer the rule chooses is then written into `KERNEL.5`'s roadmap entry. Until that entry exists, no number of the study is a measurement, and §8's "a canvas viewport" remains the sitting's recommendation, not a result.*

*Measured 2026-10-08, the benchmark half of `KERNEL.3` (`docs/PROTOCOL.md` §9.2's entry; the owner's run in fullscreen, Chrome 154 on a 2560 by 1440 CSS screen at a ratio of 1.5, 60 Hz): **canvas**. Over a dense record at full screen, a virtualized DOM renderer holds a 60 Hz frame at the default cell size (1,846 cells with text, a full rewrite at 14.3 ms of 16.7) and fails one size down (2,759 cells, 24.9 ms); the canvas renderer holds through the densest legible size (4,641 cells, 13.1 ms) and fails only at 7 px text (7,436 cells, 22.5 ms, no frame dropped). A full rewrite rises at 8.2 ms per thousand cells for the DOM and 3.1 for the canvas, a three-row scroll step at 3.9 and 1.8; the DOM's cost is the browser's paint walk over every element on screen, whatever the renderer recycles, its script's own part being 0.3 ms at every size. §7 decision 5's "a canvas viewport" is now a measurement, not a recommendation, and `KERNEL.5`'s roadmap entry names it. The study's human half, §9.1's pilot rows and the two provisional answers of §8, is still to run, and the slot above stays open for it.*
