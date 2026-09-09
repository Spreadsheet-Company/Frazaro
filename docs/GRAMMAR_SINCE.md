# THE GRAMMAR SINCE-LEDGER

*CO.6. One question, answered per form: **which release did this first
work in?** `CO.4` made versions orderable and `SD-14` says what each number
means; neither says when any particular thing you can say became sayable.
Without that, an author writing `F.10`'s `requires: version:X` is guessing
at X, and a version number alone is a loose gate — `SD-14`'s `MINOR` fires
for a new ribbon button (`GO.6`) exactly as readily as for a new grammar
rule, so "this build is ≥ 0.6.0" implies nothing about any specific form.
This file is what makes the gate tight.*

---

## The rule

1. **It records when a form first WORKED, not when it was first spelled.**
   These come apart, and the founding case is in the seed below.
   `paint cell {r:text}` shipped in `0.5.0` and `0.5.1` and never worked
   once: it called `vlacolr`, a name that resolves nowhere in the shipped
   modules, and carried no cell slot at all, so it could not have painted
   a cell even with the spelling fixed (`AS.1`'s own finding, `0.5.2`).
   It is therefore recorded at `0.5.2` — the release where the corrected
   `paint cell {r:cell} {e:expr}` first did something — and not at `0.5.0`.
   Recording the earlier date would tell an author their phrasebook runs
   on `0.5.0`, which is false, and this file exists to answer exactly
   that question. `SD-4` is untouched by this: it governs what a shipped
   spelling may be changed to, which is a different question from what a
   compatibility check should report.

2. **Append-only.** Once a row carries a `since:`, that value never
   changes — `SD-9`'s never-re-mint discipline at form scale. A wrong
   `since:` is not a cosmetic error: it is frozen, and it makes a
   `requires:` gate reject builds that would have run the phrasebook
   fine. This is why the seed below was verified against source history
   rather than taken from the generated artifact (see *How this was
   seeded*).

3. **A retired form keeps its row** and gains an `until:`. Retirement
   itself is `CO.1`'s refusal path; this file only records the dates.

## How this was seeded, and how to reproduce it

Two inventories, dated by presence at each release tag (`v0.5.0`,
`v0.5.1`), which are immutable — a GitHub ruleset forbids moving or
deleting `v*`, so the mapping cannot silently shift underneath this file.

- **Phrasebook rules** — the inventory comes from
  `scripts/polyglotta/english_expanded.vla`, the generated artifact
  `check_rule_coverage.ps1` already reads *instead of* `english.vla`,
  because the source undercounts: three rules here
  (`hide the total row of table`, `show the total row of table`,
  `set style of table`) are generator-emitted and appear in no source
  file at all.
- **Core dispatch arms** — from
  `tools/check_emitter_coverage.ps1 -ListArms`, which is that script's
  own `Get-CaseArmGroups` over its own six named dispatch functions. The
  switch was added for this file rather than copying the parsing into a
  seeder, because a second copy of subtle parsing is the divergence this
  project keeps paying for. A comma-grouped arm (`Case "exit-for",
  "exit-do"`) stays one row joined by `|`, matching how that script
  already reports it: this ledger dates a code path, not a spelling.
  `-SourceDir` points the same parser at a tag extracted to a temp
  directory, which is how the historical columns were taken.

**Two traps found while seeding, both of which would have frozen wrong
dates into an append-only file:**

- **The generated rule artifact is not safe for historical seeding.** At
  `v0.5.0` `english_expanded.vla` carried 147 rules and at `v0.5.1` it
  carried 150, which reads as three rules added in `0.5.1`. It is not:
  `english.vla`'s own rule membership is **identical** between those two
  tags. The `v0.5.0` artifact was simply stale — `0.5.1`'s release notes
  record both the re-export and the staleness stamp that had been
  ignoring whitespace. Seeding from `git show <tag>:…english_expanded.vla`
  would have stamped `show every …` / `show everyone who reports to …`
  as `since: 0.5.1` when they were sayable in `0.5.0`, permanently.
  **The inventory comes from the artifact; the dates come from source.**
- **Source alone undercounts**, per the generator-emitted three above —
  which is why the inventory cannot come from source either.

## Maintenance

This file is maintained by hand at release time: **adding a form means
adding its row in the same commit.** That discipline is no longer
enforced by review alone —
[`tools/check_grammar_since.ps1`](../tools/check_grammar_since.ps1)
(`F.10`, which is this ledger's first real consumer) fails when a live
form has no row here, so a new form cannot reach a release undated.

It takes both inventories from the scripts that already own that
parsing rather than re-deriving either: `check_rule_coverage.ps1
-ListRules` for the 150 phrase rules and `check_emitter_coverage.ps1
-ListArms` for the 128 dispatch arms. `-ListArms` was added for the
seed above; `-ListRules` was added for the checker, and both for the
same stated reason — a second copy of subtle parsing is the divergence
this project keeps paying for, and the rule-pattern reader had already
been silently broken for two weeks by exactly that class of drift.
Each script's default output was diffed byte-for-byte against a
pre-change baseline when its switch was added.

**A new ARM is not the same as a new FORM** — `IN.15`, `0.5.3`, is the
case that shows it, and it is rule 1 read in the other direction.
Sixteen `TryRuntimeHelper` arms (`vlacolor`, `vladictget`, and the
fourteen pivot helpers) appeared in the live inventory that release and
are dated **`0.5.0`**, not `0.5.3`. They are dated to the release they
first *worked* in, and every one of them worked in `0.5.0`: they were
reachable then through `TryRuntimeHelper`'s generic `Application.Run`
tier, which returns a helper's value perfectly well. What `IN.15`
changed is where a *refusal* goes — `Application.Run` never propagated
one — so what gained a date here is the dispatch mechanism, not the
sayability of the form. Dating them `0.5.3` would tell an author their
phrasebook needs `0.5.3` when it ran on `0.5.0`, which is precisely the
frozen-wrong-`since:` failure rule 2 exists to prevent. Verified, not
assumed: all sixteen are present as `Public` procedures in
`VLA_Runtime.bas` at the immutable `v0.5.0` tag.

**What the ratchet does not do**, because nothing can: check that a
date is *correct*. That is why the seed was verified against source
history rather than the generated artifact, and why rule 2 above makes
these rows append-only. It answers one question — does every live form
have a row — and its baseline is `0` undated, a ceiling with no honest
reason to rise. Rows with no live form are *reported and not failed
on*: a retired form keeps its row and gains an `until:` (rule 3), so an
extra row is legal.

---

## Snapshot

Taken 2026-09-07, during `0.5.2`'s development — `0.5.0` and `0.5.1` were
the tagged releases at the time.
**150 phrasebook rules, 128 core dispatch arms, 278 rows.** Every entry
below was present at `v0.5.0` except the one marked `0.5.2`.

### Phrasebook rules

```
0.5.0  add [new] sheet called {s:sheet}
0.5.0  add a row to table {n:text}
0.5.0  add border to range {r:range}
0.5.0  add {d:rows|columns|filters} of {f:text-list} to pivot {n:text}
0.5.0  add {f:text-list} to pivot {n:text} as {d:sum|count|average}
0.5.0  align cell|range {r:range} {d:left|right|center/ed}
0.5.0  ask {q:expr} and put answer into {v:var}
0.5.0  band every other row of {r:range} {n:expr}
0.5.0  center cell|range {r:range}
0.5.0  change the source of pivot {n:text} to {r:range}
0.5.0  clear cell in column {c:column} row {n:expr}
0.5.0  clear cell|range {r:range}
0.5.0  clear color of cell {r:cell}
0.5.0  clear everything from {r:range}
0.5.0  clear formatting of range {r:range}
0.5.0  clear pivot {n:text}
0.5.0  clear status bar
0.5.0  close this workbook
0.5.0  collapse {f:text-list} in pivot {n:text}
0.5.0  convert range {r:range} to values
0.5.0  copy cell {a:cell} to cell {b:cell}
0.5.0  copy column widths of {a:range} to {b:range}
0.5.0  copy formatting of {a:range} to {b:range}
0.5.0  copy formulas of {a:range} to {b:cell}
0.5.0  copy range {a:range} to range {b:range}
0.5.0  copy range {a:range} to {b:cell} transposed
0.5.0  copy range {r:range} to sheet {s:sheet}
0.5.0  copy row {r:expr} to row {n:expr}
0.5.0  delete blank rows in {r:range}
0.5.0  delete column {c:column}
0.5.0  delete pivot {n:text}
0.5.0  delete row {n:expr}
0.5.0  delete row {r:expr} of table {n:text}
0.5.0  delete rows {a:expr} through {b:expr}
0.5.0  delete sheet {s:sheet}
0.5.0  delete {n:expr} row
0.5.0  delete {r:range} and shift cells left
0.5.0  delete {r:range} and shift cells up
0.5.0  email {who:expr} with subject {s:expr} and message {m:expr}
0.5.0  email {who:expr} with subject {s:expr} and message {m:expr} and attachment {p:expr}
0.5.0  expand {f:text-list} in pivot {n:text}
0.5.0  export this sheet as pdf {p:expr}
0.5.0  fill {d:down|right} range {r:range}
0.5.0  fill {r:range} with a growth series starting at {n:expr}
0.5.0  fill {r:range} with a growth series starting at {n:expr} with step {s:expr}
0.5.0  fill {r:range} with a series starting at {n:expr}
0.5.0  fill {r:range} with a series starting at {n:expr} with step {s:expr}
0.5.0  fit all columns
0.5.0  fit column {c:column}
0.5.0  fit row {r:expr}
0.5.0  format cell {r:cell} as {d:currency|percent|date}
0.5.0  freeze the first {n:expr} rows
0.5.0  freeze top row
0.5.0  go to sheet {s:sheet}
0.5.0  group rows {a:expr} through {b:expr}
0.5.0  hide the total row of table {n:text}
0.5.0  insert column before {c:column}
0.5.0  insert row [at] {n:expr=1}
0.5.0  insert {n:expr} rows at row {r:expr}
0.5.0  keep only rows of range {r:range} where column {f:expr} is {v:expr}
0.5.0  make a button [called] {c:text} at cell {r:cell}
0.5.0  make a pivot table from {r:range} at {b:cell} called {n:text}
0.5.0  make a pivot table from {r:range} at {b:cell} called {n:text} with rows of {a:text-list} and columns of {c:text-list} and values of {v:text-list}
0.5.0  make cell in column {c:column} row {n:expr} bold
0.5.0  make cell {r:cell} {d:bold|italic}
0.5.0  make cell {r:cell} {d:red|yellow|black|blue|cyan|green|magenta|white}
0.5.0  make row {r:expr} a header row
0.5.0  make {a:range} look like {b:range}
0.5.0  move column {c:column} before column {d:column}
0.5.0  move range {a:range} to {b:cell}
0.5.0  name range {r:range} as {n:text}
0.5.0  open workbook {p:expr}
0.5.2  paint cell {r:cell} {e:expr}
0.5.0  paste values of range {a:range} into range {b:range}
0.5.0  print this sheet
0.5.0  put formula {f:text} into|in cell {r:cell}
0.5.0  put sum of range {r:range} into|in cell {c:cell}
0.5.0  put today into|in cell {r:cell}
0.5.0  put {e:expr} in status bar
0.5.0  put {e:expr} into|in cell {r:cell}
0.5.0  put {e:expr} into|in cell {r:cell} of sheet {s:sheet}
0.5.0  put {e:expr} into|in column number {c:expr} row {n:expr}
0.5.0  put {e:expr} into|in column {c:column} row {n:expr}
0.5.0  put {e:expr} into|in range {r:range}
0.5.0  refresh every pivot table
0.5.0  refresh everything
0.5.0  refresh pivot {n:text}
0.5.0  remember range {r:range} as {v:var}
0.5.0  remember rows {a:expr} to {b:expr} of column {c:column} as {v:var}
0.5.0  remove duplicates from range {r:range}
0.5.0  remove duplicates from range {r:range} by column {k:expr}
0.5.0  remove trailing empty rows and columns
0.5.0  remove {f:text-list} from pivot {n:text}
0.5.0  rename pivot {n:text} to {m:text}
0.5.0  replace {a:expr} with {b:expr} in column {c:column}
0.5.0  replace {a:expr} with {b:expr} in range {r:range}
0.5.0  save a copy as {p:expr}
0.5.0  save this workbook
0.5.0  save this workbook as {p:expr}
0.5.0  set fill-color of cell {r:cell} to {e:expr}
0.5.0  set font size of cell {r:cell} to {n:expr}
0.5.0  set font-color of cell {r:cell} to {e:expr}
0.5.0  set height of row {n:expr} to {h:expr}
0.5.0  set style of table {n:text} to {s:text}
0.5.0  set tab-color of sheet {s:sheet} to {e:expr}
0.5.0  set width of column {c:column} to {w:expr}
0.5.0  set {v:var} to average of range {r:range}
0.5.0  set {v:var} to cell {r:cell} of sheet {s:sheet}
0.5.0  set {v:var} to column {c:text} of table {n:text}
0.5.0  set {v:var} to count of range {c:range} matching {e:expr}
0.5.0  set {v:var} to first {n:expr} letters of {t:expr}
0.5.0  set {v:var} to last filled row of column {c:column}
0.5.0  set {v:var} to last {n:expr} letters of {t:expr}
0.5.0  set {v:var} to lookup of {e:expr} in range {r:range} column {k:expr}
0.5.0  set {v:var} to position of {a:expr} in {t:expr}
0.5.0  set {v:var} to row of {e:expr} in column {c:column}
0.5.0  set {v:var} to sum of range {r:range}
0.5.0  set {v:var} to sum of range {r:range} where range {c:range} matches {e:expr}
0.5.0  set {v:var} to trimmed {t:expr}
0.5.0  set {v:var} to {e:expr} rounded to {n:expr} decimals
0.5.0  set {v:var} to {t:expr} with {a:expr} replaced by {b:expr}
0.5.0  show all rows
0.5.0  show every {shown:text} in {table:text} with {filtercol:text} over {val:expr} as {alias:text} in cell {r:cell}
0.5.0  show every {table:text} {shown:text} whose {filtercol:text} is over {val:expr} as {alias:text} in cell {r:cell}
0.5.0  show everyone who reports to {person:expr} directly or not in {table:text} as {alias:text} in cell {r:cell}
0.5.0  show pivot {n:text} in {d:compact|tabular|outline} form
0.5.0  show the total row of table {n:text}
0.5.0  sort range {r:range} by column {k:cell} [ascending]
0.5.0  sort range {r:range} by column {k:cell} descending
0.5.0  sort {f:text} in pivot {n:text} {d:ascending|descending}
0.5.0  sort {f:text} in pivot {n:text} {d:ascending|descending} by {v:text}
0.5.0  stamp {e:expr} into|in cell {r:cell}
0.5.0  store {e:expr} at|under [key] {k:expr} in {d:var}
0.5.0  turn off screen updating
0.5.0  turn on screen updating
0.5.0  turn table {n:text} back into a range
0.5.0  turn {r:range} into a table called {n:text}
0.5.0  unfreeze panes
0.5.0  ungroup rows {a:expr} through {b:expr}
0.5.0  unhide all rows and columns
0.5.0  wait {n:expr} seconds
0.5.0  work on sheet {s:sheet}
0.5.0  {d:add|remove} a blank row after {f:text-list} in pivot {n:text}
0.5.0  {d:hide|show} subtotals for {f:text-list} in pivot {n:text}
0.5.0  {d:hide|unhide} column {c:column}
0.5.0  {d:hide|unhide} row {n:expr}
0.5.0  {d:merge|unmerge} range {r:range}
0.5.0  {d:protect|unprotect} sheet {s:sheet} with password {p:expr}
0.5.0  {d:protect|unprotect} this sheet with password {p:expr}
0.5.0  {d:wrap|unwrap} text in range {r:range}
```

### Core dispatch arms

*Backend function, then the arm. All 128 are unchanged across `v0.5.0`,
`v0.5.1` and `0.5.2` — verified by running the same parser against each
tag, not assumed.*

```
0.5.0  EmitExpr           + | * | & | / | \ | = | <> | < | > | <= | >= | and | or | xor | mod | is | like | imp | eqv
0.5.0  EmitExpr           -
0.5.0  EmitExpr           .
0.5.0  EmitExpr           array
0.5.0  EmitExpr           empty | Empty
0.5.0  EmitExpr           false | False
0.5.0  EmitExpr           include
0.5.0  EmitExpr           interpolate
0.5.0  EmitExpr           new
0.5.0  EmitExpr           not
0.5.0  EmitExpr           nothing | Nothing
0.5.0  EmitExpr           null | Null
0.5.0  EmitExpr           quote
0.5.0  EmitExpr           true | True
0.5.0  EmitStmt           .
0.5.0  EmitStmt           at-line
0.5.0  EmitStmt           begin
0.5.0  EmitStmt           call
0.5.0  EmitStmt           const
0.5.0  EmitStmt           debug-print
0.5.0  EmitStmt           deflambda
0.5.0  EmitStmt           dim
0.5.0  EmitStmt           do-until
0.5.0  EmitStmt           doc
0.5.0  EmitStmt           exit-do
0.5.0  EmitStmt           exit-for
0.5.0  EmitStmt           exit-function
0.5.0  EmitStmt           exit-sub
0.5.0  EmitStmt           for
0.5.0  EmitStmt           for-each
0.5.0  EmitStmt           for-each-row
0.5.0  EmitStmt           gen-row
0.5.0  EmitStmt           goto
0.5.0  EmitStmt           if
0.5.0  EmitStmt           include
0.5.0  EmitStmt           label
0.5.0  EmitStmt           make-button
0.5.0  EmitStmt           obj-set!
0.5.0  EmitStmt           on-error
0.5.0  EmitStmt           quote
0.5.0  EmitStmt           raw
0.5.0  EmitStmt           redim
0.5.0  EmitStmt           resume
0.5.0  EmitStmt           return
0.5.0  EmitStmt           select
0.5.0  EmitStmt           set!
0.5.0  EmitStmt           then | else | elseif | case | case-else
0.5.0  EmitStmt           type | enum
0.5.0  EmitStmt           while
0.5.0  EmitStmt           with
0.5.0  EmitTop            begin
0.5.0  EmitTop            const
0.5.0  EmitTop            deflambda
0.5.0  EmitTop            dim
0.5.0  EmitTop            enum
0.5.0  EmitTop            function
0.5.0  EmitTop            include
0.5.0  EmitTop            public | private
0.5.0  EmitTop            raw
0.5.0  EmitTop            sub
0.5.0  EmitTop            type
0.5.0  EvalDynamicHead    application.worksheetfunction.average
0.5.0  EvalDynamicHead    application.worksheetfunction.countif
0.5.0  EvalDynamicHead    application.worksheetfunction.max
0.5.0  EvalDynamicHead    application.worksheetfunction.min
0.5.0  EvalDynamicHead    application.worksheetfunction.sum
0.5.0  EvalDynamicHead    application.worksheetfunction.sumif
0.5.0  EvalDynamicHead    application.worksheetfunction.vlookup
0.5.0  EvalDynamicHead    cells
0.5.0  EvalDynamicHead    columns
0.5.0  EvalDynamicHead    make-button
0.5.0  EvalDynamicHead    range
0.5.0  EvalDynamicHead    rows
0.5.0  EvalDynamicHead    workbooks
0.5.0  EvalDynamicHead    worksheets
0.5.0  EvalExpr           * | & | / | \ | = | <> | < | > | <= | >= | and | or | xor | mod | is | like | imp | eqv
0.5.0  EvalExpr           +
0.5.0  EvalExpr           -
0.5.0  EvalExpr           .
0.5.0  EvalExpr           array
0.5.0  EvalExpr           collection
0.5.0  EvalExpr           description
0.5.0  EvalExpr           false
0.5.0  EvalExpr           interpolate
0.5.0  EvalExpr           new
0.5.0  EvalExpr           not
0.5.0  EvalExpr           number
0.5.0  EvalExpr           quote
0.5.0  EvalExpr           source
0.5.0  EvalExpr           true
0.5.0  ExecStmt           .
0.5.0  ExecStmt           begin
0.5.0  ExecStmt           const
0.5.0  ExecStmt           debug-print
0.5.0  ExecStmt           dim
0.5.0  ExecStmt           do-until
0.5.0  ExecStmt           exit-for | exit-do
0.5.0  ExecStmt           exit-sub | exit-function
0.5.0  ExecStmt           for
0.5.0  ExecStmt           for-each
0.5.0  ExecStmt           for-each-row
0.5.0  ExecStmt           goto
0.5.0  ExecStmt           if
0.5.0  ExecStmt           label
0.5.0  ExecStmt           obj-set!
0.5.0  ExecStmt           on-error
0.5.0  ExecStmt           resume
0.5.0  ExecStmt           return
0.5.0  ExecStmt           select
0.5.0  ExecStmt           set!
0.5.0  ExecStmt           while
0.5.0  TryEvalBuiltin     date
0.5.0  TryEvalBuiltin     inputbox
0.5.0  TryEvalBuiltin     instr
0.5.0  TryEvalBuiltin     isempty
0.5.0  TryEvalBuiltin     lcase
0.5.0  TryEvalBuiltin     left
0.5.0  TryEvalBuiltin     len
0.5.0  TryEvalBuiltin     msgbox
0.5.0  TryEvalBuiltin     now
0.5.0  TryEvalBuiltin     right
0.5.0  TryEvalBuiltin     round
0.5.0  TryEvalBuiltin     time
0.5.0  TryEvalBuiltin     trim
0.5.0  TryEvalBuiltin     ucase
0.5.0  TryRuntimeHelper   vlacheckrangename
0.5.0  TryRuntimeHelper   vlachecksheetabsent
0.5.0  TryRuntimeHelper   vlachecksheetname
0.5.0  TryRuntimeHelper   vlacolor
0.5.0  TryRuntimeHelper   vladictget
0.5.0  TryRuntimeHelper   vlafillseries
0.5.0  TryRuntimeHelper   vlafreezepanes
0.5.0  TryRuntimeHelper   vlapivotaddvalues
0.5.0  TryRuntimeHelper   vlapivotchangesource
0.5.0  TryRuntimeHelper   vlapivotclear
0.5.0  TryRuntimeHelper   vlapivotdelete
0.5.0  TryRuntimeHelper   vlapivotrefresh
0.5.0  TryRuntimeHelper   vlapivotrename
0.5.0  TryRuntimeHelper   vlapivotsetblankline
0.5.0  TryRuntimeHelper   vlapivotsetorientation
0.5.0  TryRuntimeHelper   vlapivotsetrowlayout
0.5.0  TryRuntimeHelper   vlapivotsetshowdetail
0.5.0  TryRuntimeHelper   vlapivotsetsubtotals
0.5.0  TryRuntimeHelper   vlapivotsort
0.5.0  TryRuntimeHelper   vlasendmail
```
