# Releases

*Newest first. `tools/release.ps1 -Version X.Y.Z` publishes the section headed `## X.Y.Z` as that release's notes and refuses to run without one, so the notes are written before the release, never after. Cadence: a `0.5.N` patch at the end of each working day, a `0.N.0` minor at the end of each week; security and safety fixes ride the patches, larger features the minors. Each section carries a short *Known open security items* block: the standing advice, what closed in that release, and a pointer to the authoritative list. It does NOT re-enumerate every open item — that list lives in `docs/BETA_ROADMAP1.md` (full, with dispositions) and `README.md` (plain words), which are edited once rather than copied into every release forever. Sections written before `0.5.3` keep their longer blocks as published; they are history, not a template.*

## 0.5.6

### What changed

- **Sorting, saying whether there is a header row.** Name the column by
  its letter, and say whether the range starts with a header row that
  should stay put:

  - `Sort range A1:C50 by column B with a header row.` — row 1 stays on
    top, the rest is sorted. `Sort range A2:C50 by column B without a
    header row.` sorts every row you named.
  - `descending` (or `ascending`) goes after the column: `Sort range A1:C50
    by column B descending with a header row.`
  - **Two columns:** `Sort range A1:C50 by column B then by column D with a
    header row.` — and to set their directions, name both: `… by column B
    descending then by column D ascending …`.
  - **The whole sheet:** `Sort this sheet by column B with a header row.`

  A sort that doesn't say "with" or "without" stops before it runs and asks.

- **Filtering.** Turn the filter buttons on, show only the rows you want,
  and put everything back:

  - `Add filters to range A1:D50.` puts the filter buttons on the header
    row. Saying it again leaves them on.
  - `Filter range A1:D50 to show rows where column C is "West".` — or
    `contains "we"`, or `is greater than 100`, or `is less than 100`.
    Rows that don't match are hidden, not deleted. A second condition on
    another column narrows the first: West rows *and* over 100.
  - `is 100` matches the number, however the cell shows it — `$100.00`
    included — and `"West"` matches West exactly, not Western. A `*` or
    `?` in what you type is just a character.
  - `Clear the filter conditions.` shows every row again and keeps the
    buttons. `Remove the filters.` takes the buttons away too.
  - `Copy only the visible cells of range A1:D50 to F1.` copies what a
    filter leaves showing — rows and columns that are hidden stay behind.

  Filters work one set per sheet, the way Excel's do: adding them to a
  second range while the first still has them stops and says so, and so
  does filtering an empty range. A filter sentence that stops leaves the
  sheet exactly as it was. In a sort or a filter, a column that isn't part
  of the range — `by column F` on `A1:C50` — stops and names the column.

- **Three sentences you may already use now have clearer twins.** They
  keep working exactly as before. `Sort range A1:C50 by column B1.` always
  keeps row 1 where it is, even when row 1 is data — the new sentences say
  `with a header row` or `without`. `Keep only rows of range A1:D50 where
  column 3 is "West".` counts columns from the start of the range and
  hides the other rows rather than deleting them — `Filter range … to show
  rows where column C is …` names the column by its letter and says it
  shows. `Show all rows.` clears filter conditions and does not unhide rows
  hidden with `Hide row` — `Clear the filter conditions.` says which it is.

- **Number formats, for a cell or a range.** Money, percentages, plain
  numbers, dates and times, each in one sentence:

  - **Money, with the currency named:** `Format range B2:B9 as dollars.` —
    or `euros`, or `pounds`. And Excel's accounting layout, with the sign
    at the left and negatives in brackets: `Format range B2:B9 as
    accounting in dollars.`
  - **Percentages and plain numbers:** `Format range C2:C9 as percent.`,
    `Format range D2:D9 as a number.`, and `… as a number with thousands
    separators.`
  - **Any of those with a set number of decimals:** `Format range B2:B9 as
    euros with 0 decimals.`, `Format cell C2 as percent with 1 decimal.`,
    `Format range D2:D9 as a number with 0 decimals and thousands
    separators.` (0 to 30 decimals.)
  - **Dates and times:** `Format range A2:A9 as a short date.` — or `a long
    date`, or `an ISO date` (2025-12-09) — and `Format range E2:E9 as a
    time.`
  - **Back to plain, or anything else:** `Format range A1:A9 as general.`
    clears the format; `Format range A1:A9 using "00000".` takes any Excel
    format code, written the way Excel's own Format Cells box writes it.
  - **Text, for what you type next:** `Format range A2:A9 as text for new
    entries.` Anything typed there afterwards stays exactly as typed —
    `007` keeps its zeros. A number already in the cell stays a number,
    which is what Excel's own Text format does, and the sentence says so.

- **Formats follow the person reading, and the currency follows the
  numbers.** The same workbook opened on computers set up for different
  countries shows dates in each country's own order and uses each
  country's own decimal point and thousands separator — so `as a short
  date` is 09/12/2025 on a UK machine and 12/9/2025 on a US one, both
  correct. The currency never changes that way: dollars stay dollars
  wherever the file is opened, because the currency is a fact about the
  numbers, not the computer. That is why the sentence names it.

  The older `Format cell A1 as currency.` and `… as date.` still work
  exactly as before, and it is worth knowing what "before" was: `as
  currency` has always shown *the currency of the computer opening the
  file* — a $ in the US, a £ in the UK, on the same figures. If your
  numbers are in one currency, say which: `as dollars`. `as date` keeps
  month, day, year in that order everywhere; `as a short date` follows
  the reader.

- **Formatting sentences now reach a whole range, not just one cell.**
  Five sentences only ever took a single cell — `Make cell A1 bold.` would
  refuse `A1:C3`. Each now has a range version, worded the same way with
  `range` in place of `cell`:

  - `Make range A1:C1 bold.` (and `italic`)
  - `Make range A1:C1 blue.` — any of the eight named colours
  - `Set font-color of range A1:C1 to "#FF0000".`
  - `Set fill-color of range A1:C1 to hot-pink.`
  - `Set font size of range A1:C1 to 14.`

  The cell versions are unchanged.

- **Twelve new ways to make a sheet look right.** Each new sentence
  below takes a cell or a range — `cell A1` or `range A1:C3`:

  - **Undo bold and italic:** `Make range A1:C1 not bold.`,
    `Make cell A1 not italic.`
  - **Underline and strike through:** `Underline range A1:C1.`,
    `Strike through cell A1.` — and the one sentence that takes marks off
    again: `Remove underline from …`, `Remove strikethrough from …`,
    `Remove borders from …`.
  - **Typeface:** `Set font of range A1:C1 to "Courier New".` Put a font
    name in quotes; a name with a space in it has to be.
  - **Vertical alignment:** `Align range A1:C1 to the top.` — or `middle`,
    or `bottom`. (Left, right and centre were already there.)
  - **Indent and rotate:** `Indent range A1:C1 by 2.` (0 to 15 levels),
    `Rotate text in range A1:C1 by 45 degrees.` (−90 to 90).
  - **Borders, three ways, each saying which lines it draws:**
    `Add a border around range B2:D4.` draws the outside box only;
    `Add a bottom border to range B5:D5.` draws one edge (`top`,
    `bottom`, `left` or `right`); `Add borders to every cell in range
    A1:C3.` draws every line, inside and out.
  - **Clear a fill:** `Clear fill-color of range A1:C1.`

- **Border colour, as part of drawing the border.** Say `colored` and a
  colour right after the word border: `Add a border colored red around
  range B2:D4.`, `Add a bottom border colored green to range B5:D5.`, `Add
  borders colored "#0000FF" to every cell in range A1:C3.` Any of the eight
  colour names works as it is — red, yellow, black, blue, cyan, green,
  magenta, white — and any other colour as a code in quotes, like
  `"#FF69B4"`; a name you have defined yourself, like `hot-pink`, goes in
  quotes here too. Each sentence colours only the lines it draws — the box
  around a range does not draw its inside lines.

- **Two sentences you may already use now have clearer twins.** `Add
  border to range A1:C3.` draws a line around *every cell*, though it reads
  like it draws one box around the range; `Clear color of cell A1.` clears
  the *fill* and leaves the text colour alone, though "color" could mean
  either. Both keep working exactly as before — nothing you have written
  changes. The new spellings say what they do: `Add borders to every cell
  in …` and `Clear fill-color of …`, and they are the ones the examples use
  from now on. This is a new standing rule for every sentence Frazaro adds:
  **where Excel has two neighbouring operations, the sentence names the one
  it performs.** (A plain "Set border-color of …" was held back under it:
  Excel's own way of doing that also draws lines you did not ask for.
  Border colour arrives instead as part of drawing the border — see above.)

- **Colour names in quotes now work wherever a colour code does.** `Set
  font-color of cell A2 to "red".` used to stop with *"'red' is not a
  color"*; it now means red, like any of the eight names. Nothing that
  worked before changes.

- **Six colours now work under Interpret.** `Make cell C4 black.` — and
  `blue`, `cyan`, `green`, `magenta` and `white` — has always translated
  and run with Run, but stopped with *"there is nothing stored at key
  'vbblack'"* under Interpret; only `red` and `yellow` worked there. All
  eight now work both ways.

- **Interpret can set five more formatting properties, each reviewed by
  name.** Interpret only touches the parts of Excel someone has read and
  listed, and refuses everything else in words. This release adds five —
  underline, strikethrough, vertical alignment, indent and text rotation —
  each of which changes how a cell looks and nothing else. No new
  *action* was added: the outline border is drawn as four single-edge
  lines rather than through a new Excel command.

  Sorting and filtering add four more, also by name: taking a sheet's
  filter buttons away (it can only take them away, never add them), and
  three that only find cells already on the sheet — a column of a range,
  the part of the sheet in use, and the visible cells of a range. Sorting
  and filtering themselves were already on the list; they now also take
  a second sort column and a second filter condition.

- **PROLOG can work with text.** A spreadsheet's whole subject is cell
  values, and until now PROLOG could not take one apart or put two
  together. Seven goals do that now:

  - `(atom-length Text N)` — how many characters. `(atom-length "hello" N)`
    gives `5`.
  - `(atom-concat A B Whole)` — join two pieces: `(atom-concat "Item " 42 X)`
    gives `Item 42`. Given the whole, it takes it apart instead:
    `(atom-concat "ID-" Rest "ID-42")` gives `Rest` = `42`,
    `(atom-concat Stem ".xlsx" "report.xlsx")` gives `report`, and with
    both halves left blank it lists every way to split.
  - `(sub-atom Text Before Length After Part)` — any piece, by position or
    by content. `(sub-atom "Frazaro" 2 3 After Part)` gives `aza`;
    `(sub-atom Email Before 1 After "@")` finds where the `@` is. Positions
    count from 0, the way Prolog counts them.
  - `(atom-number Text N)` — text to a number and back.
    `(atom-number "42" N)` gives the number `42`; text that is not a number
    simply does not match, so it doubles as the test "is this a number?".
  - `(upcase-atom Text Upper)` and `(downcase-atom Text Lower)` — case.
  - `(atomic-list-concat Parts Separator Whole)` — split and join.
    `(atomic-list-concat P "," "a,b,c")` gives the list `(list "a" "b" "c")`,
    ready for `length`, `member` and the other list goals; run the other
    way it joins a list. `(atomic-list-concat (list First Last) " "
    "Ada Lovelace")` fills in both names at once.

  The names are real Prolog's, with a hyphen where Prolog writes an
  underscore, the way `sum-list` already is. Type the underscore version
  — `atom_length`, `sub_atom` — and Frazaro tells you the hyphenated one
  instead of quietly finding nothing.

- **What a text goal hands back is text.** The result behaves exactly
  like the value of a text cell: `(atom-concat 4 2 X)` makes the text
  `"42"`, not the number 42, and to compare a result with something you
  typed, quote it — `(== X "abc")`. Compare it with an unquoted name and
  Frazaro stops and explains the difference rather than quietly answering
  no. The values you *give* a text goal are read as text whichever way
  they were written, so `(atom-concat ab c abc)` is true and
  `(downcase-atom "ENG" eng)` is true.

- **Changing case gives the same answer on every computer.** The usual
  way to change case in Excel's own language depends on the Windows
  language settings — a Turkish Windows lowercases `I` differently — so a
  cell's answer would have depended on whose machine computed it.
  `upcase-atom` and `downcase-atom` instead carry their own table: A–Z and
  the accented letters of Western and Central European languages,
  Spanish's `ñ`, `á` and `ü` among them, and Polish, Czech and Turkish
  letters too. Letters from other alphabets — Greek, Cyrillic and so on —
  are refused by name rather than passed through unchanged and looking
  finished. Characters with no case at all, like digits or `€`, pass
  through as they are. Four characters whose case is genuinely disputed
  — the micro sign `µ`, the Turkish dotless `ı` and dotted `İ`, and the
  old long `ſ` — are refused in the one direction where the authorities
  disagree.

- **Emoji and counting.** Excel counts an emoji as two characters and
  Prolog counts it as one. The goals that count or cut at a position —
  `atom-length`, `sub-atom`, and `atom-concat` when it lists every split —
  refuse such text and say why, rather than silently pick one of the two
  answers. The goals that do not count, like changing case, pass emoji
  through untouched.

- **Taking a long text apart in every possible way has a limit, and says
  so.** `(sub-atom Text B L A S)` with only the text given lists every
  piece of it: 105 pieces for a 13-character text, and 120 for 14, which
  is the whole of a query's work allowance. Past that it is refused by
  name *before* it starts, with a message saying to fill in more of the
  arguments. Anything more specific — a position, a length, the piece
  you are looking for, or a known start or end for `atom-concat` — is
  answered directly and works on a text of any length.

- **Every number PROLOG writes is now written one way.** Until now a
  number could get two different spellings depending on where it came
  from. The result of `(is ...)` followed the computer's regional
  settings, so on a Windows set to write `0,5` the next calculation
  refused it as not a number. A table cell holding `0.5`, or a total from
  `sum-list`, was written `.5` without its leading zero — and so it never
  matched a `0.5` you typed: `(sum-list (list 0.25 0.25) 0.5)` answered
  *no*. Both now read `0.5` on every machine.

- **`(length "hello" N)` now points you at `atom-length`.** It is the
  first thing most people type. It is still refused — a piece of text is
  not a list — but the message now names the goal you wanted, and
  `(append "ab" "c" X)` names `atom-concat` the same way.

- **`(whole 3)` now points at `(whole? 3)`.** It used to be an unknown
  name that quietly found nothing. The near-miss spellings Frazaro
  recognises — a hyphen for an underscore, a missing question mark — are
  now worked out automatically from the full list of names on every
  release, which is how this one was found missing.

- **PROLOG refuses Prolog's side-effect goals, for good, and says why.**
  `write`, `assert`, `random` and the rest of that family used to be
  unknown names here, and until this release an unknown name quietly found nothing — so the
  line everyone types while debugging, `(query (p X) (write X))`, answered
  with an empty column and looked like "no match". Each is now refused
  with a message giving the reason and what to do instead:

  - **Changing the program while it answers** — `assert`, `asserta`,
    `assertz`, `retract`, `retractall`, `abolish`. Excel recalculates a
    formula whenever it chooses, as often as it chooses, so an answer that
    depended on how many times it had run would be wrong in a way no one
    could see. To collect answers use `(findall X Goal Bag)`; to count
    them, `(length Bag N)`.
  - **Printing** — `write`, `writeln`, `print`, `nl`, `format`, `writeq`,
    `write_canonical`, `write_term`. A formula in a cell has nowhere to
    print; what `PROLOG` returns *is* its output. Put the variable in the
    query and it fills a column, one row per answer.
  - **Remembering a value between goals** — `b_setval`, `b_getval`,
    `nb_setval`, `nb_getval`, `gensym`. In Prolog a value kept this
    way outlives the query, and a formula must not carry anything from
    one recalculation to the next — a counter would count how often
    Excel had recalculated. Pass the value along as an argument instead.
  - **The clock and the dice** — `random`, `random_between`,
    `random_member`, `random_permutation`, `get_time`, and `random`,
    `random_float` and `cputime` inside arithmetic, as in
    `(is X (random 10))`. Their answer would change every time Excel
    recalculated. Excel's own `RAND()`, `RANDBETWEEN()` and `NOW()` do
    this in the open: put one in a table you pass to `PROLOG`.
  - **Reading from outside** — `read`, `read_term`, `consult`, `halt`.
    `PROLOG` reads only its own program and the tables you give it.

  Written bare, the way Prolog writes them — `nl`, `halt` — they are
  refused the same way, and the hyphenated spellings (`random-between`,
  `get-time`, …) are refused too. A program using `format` with one
  argument in one place and two in another is told about `format`, not
  about its arities. A rule that contains one of these goals is refused
  only when it actually runs, so a rule you never call does no harm.
  None of these is waiting to be built.

- **More reserved words.** Fifty-four names join the reserved set: the
  seven text goals, their seven underscore spellings, `whole`, the
  twenty-eight goals above and eleven hyphenated spellings of them. If a
  knowledge base of yours uses one of those as a predicate name it will
  need renaming; `write`, `read`, `print`, `format` and `whole` are
  the plausible ones. `tab` and `flag` are deliberately **not** reserved:
  a sheet tab and a flagged order are ordinary things to keep facts about.
  The same list applies to the name of a table you pass to `PROLOG`, in
  any capitalisation: a Table named `Write` or `Between` is refused as
  a reserved word rather than loaded, so rename it.

- **A predicate nothing defines now stops the formula and says so.** A
  misspelled name used to find nothing, quietly:
  `=PROLOG("(fact (parent tom bob)) (query (parnet tom X))")` answered
  with an empty column and looked like "no match". Now it says
  `'parnet' is called by this query` but nothing defines it, and where a
  definition could come from — a `(fact ...)`, a `(rule ...)` or a table
  you pass in. The silence was worst where it was wrong rather than
  empty: `(not (parnet tom bob))` answered TRUE, and a `findall` over a
  misspelled goal collected nothing and counted it as zero. Both now stop.

  PROLOG checks this before it answers anything, and only for what your
  query can reach — the predicates it calls, and the ones those rules
  call. A rule nothing calls is not examined, so one cell of rules can
  serve several formulas that pass different tables. Inside what the
  query can reach, it is checked even on a branch today's data never
  takes — the else-part of an `(if ...)` that never runs — because that
  formula would otherwise break the day the data changed. A table you
  pass that has no rows is not undefined: it is there and empty, and a
  query over it returns no rows, as before.

  If a program of yours used an undefined name on purpose, as a goal
  that always fails, write one that matches nothing instead: `(= a b)`
  fails every time.

- **Cut in parentheses is refused.** Cut is written on its own, between
  the goals: `(rule (first X) (p X) !)`. Written `(!)`, it used to be a
  goal nothing could run, which quietly failed — so the rule never
  answered. It now stops and says how to write it.

- **A spelling mistake is reported as a spelling mistake.** Using a
  Prolog spelling at two sizes — `(atom X)` in one place and
  `(atom X Y)` in another, or `(integer? X)` beside `(integer? X Y)`
  — used to be blamed on the arguments: "used with 1 argument(s) in one
  place and 2 in another". It now gets the message it gets when written
  once: that `(atom ...)` is written `(atom? ...)`, or that
  `integer?` is not available and what to ask instead.

### Known open security items

**Closed this release:** none — 0.5.6 is a feature release. Three changes
touch the security surface without closing an item, listed so none is a
surprise:

- *Interpret* can now set five more cell properties — underline,
  strikethrough, vertical alignment, indent and text rotation. Interpret
  only reaches the parts of Excel on a list someone has read and named
  (`SEC.1`, `docs/THREAT_MODEL.md` §1.1); these five were added to that
  list by name, each changes only how a cell looks, and anything the list
  does not name is still refused in words. No new *action* was added.
- *Interpret* can now do four more things for sorting and filtering, each
  added to that same list by name: take a sheet's filter buttons away
  (Excel lets a program only take them away, never add them), and find
  cells already on the sheet — a column of a range, the part of the sheet
  in use, and a range's visible cells. Sorting and filtering were already
  on the list; they now also take a second column and a second condition.
  Nothing here reaches outside the workbook.
- PROLOG now refuses, by name, the Prolog goals that would read from
  outside its own program and tables — `read`, `read_term`, `consult` —
  along with those that print, keep state between answers, or depend on
  the clock. None of them ever ran here; they used to be unknown names
  that quietly found nothing, and now they say why they will not.

**Still open:** `SEC.3`, `SEC.7`, and two from the 2026-09-08 code
review — `SEC.10` and `SEC.15`. In plain words:
the remembered raw-VBA consent record still lives inside the workbook
(`SEC.10`);
formulas a program writes are not screened for functions that reach the
network (`SEC.15`); and effects like sending mail still run without a
permission prompt (`SEC.7`). `SEC.8` narrows that last one — it gates on
where the *workbook* came from — but does not close it: a phrasebook loaded
into a workbook of your own still reaches those verbs unprompted.

**Assessed and accepted, not fixed:** `SEC.12`, `SEC.14`, `SEC.16` and
`SEC.17`. Each needs a precondition an ordinary install does not meet —
mostly an Excel setting that ships off and that Frazaro never asks you to
turn on. The reasoning for each, and what would reopen it, is written down
rather than left implied.

The authoritative lists, kept current in one place instead of copied into
every release: [`README.md`](../README.md) in plain words, and
[`docs/BETA_ROADMAP1.md`](BETA_ROADMAP1.md) with the file, line, fix and
disposition for each.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from — and treat a workbook someone sent you the
same way before you press Interpret.** Frazaro makes no network call and
does not update itself; check the README's *Known open security items*
when you return for a newer build. Vulnerability reports:
`docs/SECURITY.md`. Everything else: `docs/SUPPORT.md`.

## 0.5.5

### What changed

- **PROLOG's type tests are finished.** `0.5.4` gave six — `var?`,
  `nonvar?`, `atom?`, `number?`, `atomic?` and `compound?`. Three more
  join them, and they are the three a real program actually reaches for.

  **`(ground? T)`** succeeds when `T` is completely filled in — no blanks
  anywhere inside it, however deep. `(ground? (order alice Qty))` fails
  while `Qty` is still unfilled and succeeds once it isn't. *Completely*
  is the word to read twice: one blank anywhere inside, at any depth, is
  enough to fail. This is the guard to write before anything that needs a
  finished value: printing it, storing it, comparing it.

  **`(callable? T)`** succeeds when `T` is a name or a structure —
  anything that could stand where a goal stands. A number can't, so
  `(callable? 42)` fails where `(nonvar? 42)` succeeds. That is the whole
  difference between the two, and it is the reason `callable?` exists.

  **`(is-list? L)`** succeeds when `L` really is a list — a chain that
  ends properly rather than trailing off. `(is-list? (list a b c))`
  succeeds; `(is-list? (cons a b))`, which ends in `b` instead of the
  empty list, does not; and neither does a list that is still half
  unfilled.

- **A type test answers; it never stops the query.** This is worth
  stating because the list *operations* behave differently on purpose. If
  you hand `(sum-list L N)` something that is not a list, it stops and
  says so — it cannot do its job without one. `(is-list? L)` on the same
  value simply answers *no*. That is what makes it usable as a guard:

  ```
  (fact (box (list 1 2 3)))
  (fact (box plain))
  (rule (total B N) (is-list? B) (sum-list B N))
  (query (box B) (total B N))
  ```

  gives one row — `(list 1 2 3)` and `6`. The `plain` row is skipped
  quietly instead of stopping the whole query, which is exactly what
  happens if you delete the `(is-list? B)` guard and run it again.

- **They see what a value *is*, not how it was written.** After
  `(= X 1)`, `(ground? X)` succeeds — it follows the value X was given
  rather than reading the letter `X`. The same is true all the way down:
  a list whose tail was filled in somewhere else is still a list.

- **`(integer? X)` and `(float? X)` are reserved, and Frazaro tells you
  why rather than failing quietly.** PROLOG has one kind of number
  today. `3` and `3.0` are not merely equal, they are the same value
  written down, so there is no honest answer to "is this one an integer".
  Rather than ship a test that would give a confident wrong answer — and
  one whose meaning would have to change later — the two names are
  reserved now and refused with an explanation pointing at `(number? X)`,
  the test that does exist.

  **That question is now settled, in the same release.** One kind of
  number is permanent, not a gap waiting to be filled, so `integer?` and
  `float?` stay refused for good rather than pending. What a person
  actually wanted when they reached for `integer?` — "has this got a
  fractional part?" — ships instead as **`(whole? X)`**, under a name
  that promises no type. `(whole? 3)` and `(whole? 3.0)` both succeed;
  `(whole? 3.5)` fails. The refusal now points at it by name.

- **If you write the Prolog spelling, Frazaro still tells you.**
  `(ground X)`, `(callable X)` and `(is_list X)` each stop and name the
  spelling to use instead. Note the last one: Prolog writes `is_list`
  with an underscore, Frazaro writes `is-list?` with a hyphen and a
  question mark, and typing the Prolog form gets you the Frazaro form
  rather than silence.

- **Get the spelling slightly wrong and Frazaro says so.** `sum-list` is
  written with a hyphen; Prolog writes `sum_list` with an underscore. A
  name that asks a question ends in `?`; `is-list` without one is one
  character short of `is-list?`. Each of those used to be an unknown
  predicate — which finds nothing and explains nothing — and now stops
  and names the spelling to use.

  The three it catches (`is-list`, `is_list?`, `sum_list`) are not a list
  someone thought of. They are every spelling you can reach from a real
  name by swapping a hyphen for an underscore or dropping a question
  mark, worked out mechanically over the whole reserved set, which is
  what makes the set complete rather than a guess.

- **PROLOG can say "or", and "if this then that, otherwise the other".**
  Until now the only way to say *or* was to write the same rule twice,
  which works for a whole rule and not at all for one step in the middle
  of a longer one. Two new forms fix that.

  **`(or A B ...)`** succeeds if any of its goals does, trying them left
  to right. Two or more, as many as you like:

  ```
  (rule (contact P) (or (employee P) (contractor P) (visitor P)))
  ```

  **`(if C T E)`** proves `C`; if that works it commits to the first way
  it worked and goes on to `T`, and if it doesn't it does `E` instead.
  The else-goal may be left off — `(if C T)` simply fails when `C` does.

  ```
  (query (if (member X L) (found X) (missing X)))
  ```

  Two things about `(if ...)` are worth knowing before you rely on it.
  It **commits**: once the condition succeeds one way, the other ways it
  might have succeeded are not tried. And the else-goal runs **only when
  the condition never succeeded at all** — if the condition works and the
  then-goal then fails, the whole thing fails rather than falling through
  to the else. Both are what Prolog does, and both are easy to assume
  backwards.

- **Which blanks come back as columns, when the answer had a choice in
  it.** Worth reading once, because the rule is not what you might guess
  and it is the one thing about these forms likely to look like a bug.

  A blank becomes a column of the answer **only if every way of
  succeeding fills it in**. Ask

  ```
  (query (or (p X) (q Y)))
  ```

  and you get a plain `TRUE` or `FALSE`, not two columns. That is
  deliberate: `X` is filled in only when the first branch is the one that
  worked and `Y` only when the second is, so on any given answer one of
  them is still blank — and a column of blanks reported as though it held
  values is worse than no column. Ask about a blank **both** branches
  fill in

  ```
  (query (or (p X) (q X)))
  ```

  and `X` comes back as a real column, one row per branch that succeeded.

  `(if C T E)` follows the same rule with one wrinkle worth knowing: a
  blank filled in by the **condition** counts as filled for the
  then-branch, because the condition's answers carry forward into it —
  but not for the else-branch, which is only ever reached when the
  condition failed and filled in nothing. So in

  ```
  (query (if (link W M) (tag M T) (nope T)))
  ```

  `T` is a column and `W` and `M` are not. Leave the else off, and there
  is only one way to succeed, so all three come back.

- **Why `or` and `if` rather than Prolog's `;` and `->`.** Real Prolog
  writes these `;` and `->`. **`;` is not available here and cannot be**:
  Frazaro's reader has always treated `;` as a comment that runs to the
  end of the line, long before PROLOG existed. That is not a small
  inconvenience — a `;` written mid-rule quietly deletes the rest of that
  line, and on the usual multi-line layout the rule still parses. You get
  a *different rule* rather than an error, and nothing tells you. So `;`
  had to go, and `or` and `if` are the words Frazaro already uses for
  these ideas everywhere else.

  Typing `(-> ...)` or `(\+ ...)` — Prolog's arrow and its negation —
  now gets you a short note naming the form Frazaro uses instead, rather
  than silence. The `\+` one matters most: left unrecognised it would
  have quietly *failed*, and a negation that fails looks exactly like a
  negation that worked.

- **A cut (`!`) inside one rule no longer cancels a cut in the rule that
  called it.** A real bug, present since cut shipped in `0.5.0` and found
  while building the above. If a rule used `!`, and a rule it called
  *also* used `!`, the inner one silently cancelled the outer one — so
  the outer rule kept offering answers it had been told to stop at. Rare,
  because it needs cuts at two levels at once, but wrong whenever it
  happened, and wrong quietly.

- **More reserved words.** Eighteen names join the reserved set, which is
  what lets Frazaro's advice about them always be right: `callable?`,
  `is-list?`, `ground?`, `integer?` and `float?`; the Prolog spellings
  `callable`, `is_list`, `ground`, `integer` and `float`; the three
  near-misses above; `whole?`, the whole-number test described
  earlier; and `or`, `if`, `->` and `\+` from the two control forms. If a
  knowledge base of yours uses one of those as a predicate name, it will
  need renaming — `ground`, `integer`, `float`, `or` and `if` are the
  plausible ones.

  `cons` and `nil` are still **not** reserved, unchanged from `0.5.4`: a
  predicate of your own may still be called that. Nor is `nth1`, which
  looks like a near-miss for `nth` and is not one — Prolog's `nth0` and
  `nth1` count from different places, so guessing which you meant would
  be worse than saying nothing.

- **Arithmetic grew from four operators to seventeen, in all three
  query languages at once.** Until now PROLOG, SQL and DATALOG could add,
  subtract, multiply and divide, and nothing else — no remainder, no
  integer division, no rounding. There is now `mod`, `rem`, `//`
  (integer division), `min`, `max`, `**`, `abs`, `sign`, `sqrt`,
  `truncate`, `round`, `ceiling` and `floor`.

  In PROLOG they are written the way the existing four are:
  `(is X (mod 7 3))`, `(is X (round 2.5))`. In DATALOG they go inside
  `let`: `(let S (abs A))`. SQL spells them as functions —
  `MOD(Salary, 7)`, `ABS(Salary - 95000)`, `ROUND(x)`, `POWER(2, 10)`,
  `DIV(a, b)`. The scalar two-argument minimum and maximum are
  **`LEAST`** and **`GREATEST`** in SQL, not `MIN`/`MAX`: those two are
  already aggregate functions there, and quietly changing what `MIN(x)`
  means would have altered the meaning of queries people have already
  written.

- **`mod` and `rem` both ship, because they disagree and you should not
  have to guess which one you got.** They are identical whenever the two
  numbers share a sign. When the signs differ they part company:
  `(mod -7 3)` is `2` and `(rem -7 3)` is `-1`. `mod` follows the sign of
  the number you divide *by*; `rem` follows the sign of the number you
  divide *into*. Shipping only one under the name `mod` would have been
  right half the time and silently wrong the other half.

- **`round` rounds halves away from zero.** `ROUND(2.5)` is `3`, not `2`.
  Worth stating because the rounding built into the underlying platform
  rounds halves to the nearest *even* number, which would have made
  `ROUND(2.5)` and `ROUND(3.5)` both `4`.

- **Arithmetic that has no answer now says so instead of guessing.**
  The square root of a negative, a negative number raised to a fractional
  power, a remainder or integer division by zero, and a result too large
  to represent are each refused by name, in whichever of the three
  languages you were writing. Previously several of these could surface
  as a raw platform error rather than an explanation.

- **A new release check: `tools/check_prolog_arith_operators.ps1`.** The
  code that actually performs arithmetic is shared by all three query
  languages, and each language separately decides which operators it will
  accept. That is an arrangement where an operator can be added in one
  place and forgotten in another — and the result would not be an error
  message but a *wrong number in a cell*, which looks like an answer.
  This check holds the one operator list, both refusal messages that
  advertise it, both computing functions, and SQL's own function-name
  map against each other, and fails the build if any of them disagree.

- **A new release check: `tools/check_test_assertion_safety.ps1`.** VBA
  evaluates *every* part of an `And`, even when an earlier part has
  already answered no. So a test written as one expression — check the
  answer has three rows, then read the third row — still reads that third
  row when there are only two, and stops with a raw error instead of
  reporting the failure. Harmless every day the test passes; on the one
  day it finds something, the run dies at the exact moment it was about
  to say what broke, and every later test never runs.

  This is invisible in a passing test suite by construction, so nothing
  but a static check can find it. **165 assertions across all four test
  files were carrying it**, and all 165 are fixed. The check now runs at
  every release and holds each file at zero — and fails just as loudly if
  a file comes in *under* its recorded number without the number being
  lowered too, so the ground gained cannot quietly be given back.

  It also catches the half-fixed version, which is the one that would
  have bitten: a test can be repaired where it checks the answer and left
  reading past the end of it where it *reports* the answer — and the
  report is built every single time, pass or fail.

  And it catches the opposite mistake: a test that expects an answer
  Frazaro can no longer give. When a value's printed form changes, every
  test naming the old form has to change with it — that happened once and
  two were missed by hand. The check now knows which forms are retired,
  so it lists them instead of leaving it to memory.

  Nothing here changes what Frazaro does. It changes what happens on the
  day something else goes wrong: you get told, instead of the report
  being destroyed by the test that was about to make it. The rewrite was
  checked against VBA's own rules for every shape a result can take,
  proving each replacement means exactly what it replaced, and both the
  check and that proof were mutation-tested rather than assumed to work.

### Known open security items

**Closed this release:** none — 0.5.5 is a feature release and touches no
security item.

**Still open:** `SEC.3`, `SEC.7`, and two from the 2026-09-08 code
review — `SEC.10` and `SEC.15`. In plain words:
the remembered raw-VBA consent record still lives inside the workbook
(`SEC.10`);
formulas a program writes are not screened for functions that reach the
network (`SEC.15`); and effects like sending mail still run without a
permission prompt (`SEC.7`). `SEC.8` narrows that last one — it gates on
where the *workbook* came from — but does not close it: a phrasebook loaded
into a workbook of your own still reaches those verbs unprompted.

**Assessed and accepted, not fixed:** `SEC.12`, `SEC.14`, `SEC.16` and
`SEC.17`. Each needs a precondition an ordinary install does not meet —
mostly an Excel setting that ships off and that Frazaro never asks you to
turn on. The reasoning for each, and what would reopen it, is written down
rather than left implied.

The authoritative lists, kept current in one place instead of copied into
every release: [`README.md`](../README.md) in plain words, and
[`docs/BETA_ROADMAP1.md`](BETA_ROADMAP1.md) with the file, line, fix and
disposition for each.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from — and treat a workbook someone sent you the
same way before you press Interpret.** Frazaro makes no network call and
does not update itself; check the README's *Known open security items*
when you return for a newer build. Vulnerability reports:
`docs/SECURITY.md`. Everything else: `docs/SUPPORT.md`.

## 0.5.4

### What changed

- **PROLOG can now ask what kind of thing a value is.** Six new goals —
  `var?`, `nonvar?`, `atom?`, `number?`, `atomic?` and `compound?` — each
  take one term and succeed or fail depending on its shape.
  `(number? Salary)` succeeds when `Salary` has been bound to a number,
  `(var? X)` succeeds only while `X` is still unfilled, `(compound? T)`
  when `T` is a structure like `(name Alice)` rather than a single value.
  They are the ordinary Prolog guards, and until now a rule had no way to
  check that what it received was the kind of thing it knew how to
  handle.

  They never change anything — a type test looks at a value and answers.
  It cannot fill in a blank, so it adds no column to your results.

  **The question mark is deliberate, and it is not decoration.** In
  Prolog you write `atom(X)`, but Frazaro's PROLOG is written in
  parentheses, where `(atom bob)` looks exactly like a piece of *data* —
  `(color red)`, `(name Alice)` and `(f a)` are all ordinary data written
  in that same shape. Prolog doesn't have that problem because its
  grammar tells a question apart from a value; ours doesn't, so the name
  has to. The question mark also matches what Frazaro already does
  elsewhere: its macro layer has long spelled its questions `null?`,
  `eq?` and `equal?` beside `car`, `cdr`, `cons` and `list`. A name that
  asks something ends in `?`; a name that produces something doesn't —
  which is why `between` below has none.

  **If you write the Prolog spelling, Frazaro tells you.** `(atom X)`
  doesn't quietly find nothing; it stops and says that this type test is
  spelled `(atom? X)`. All six work that way, and none of the six bare
  names can be used for a predicate of your own, so the advice can never
  be wrong.

- **`(between Low High X)` counts.** With `X` unfilled it produces every
  whole number from `Low` to `High` in turn, so a query can range over
  1 to 10 without ten hand-written facts. With `X` already filled it
  tests instead: `(between 1 10 5)` simply succeeds. Both bounds may be
  arithmetic, so `(between 1 (+ N 1) X)` works.

  If `Low` is greater than `High` the range is empty and the goal finds
  nothing — that is an ordinary "no rows", not an error, so a range
  computed from your data can safely come out empty.

  **A limit worth knowing before you meet it: the range can span at most
  119 values.** A whole query gets a fixed budget of resolution steps,
  and each generated value spends one. Ask for more and Frazaro refuses
  up front and tells you the range was too wide — deliberately, so that
  a wide range never gets reported as the *other* thing that exhausts
  that budget, a rule that calls itself forever. Those are different
  mistakes and now have different messages. Testing a single value is
  not affected, so `(between 1 1000000 5)` succeeds normally.

- **What the two are for, together.** A rule can now check what it was
  handed before doing arithmetic on it, and a query can feed it a range
  instead of a table:

  ```
  (rule (halved N H) (number? N) (is H (/ N 2)))
  (query (between 1 5 X) (halved X Y))
  ```

  That spills five rows — 1 through 5 beside their halves. The guard is
  what makes the rule safe to call with anything at all: without
  `(number? N)`, asking for `(halved eng H)` stops the entire query with
  an arithmetic refusal, because `eng` cannot be divided. With the guard,
  that call simply doesn't match, and the rest of the query carries on.
  That is what a type test is *for*, and until this release there was no
  way to write one.

- **Text that looks like a number is still text, and now says so.** A
  cell containing `42` typed as text is not the number 42 in PROLOG, and
  never has been — `=` and `==` have always treated the two as different
  values. The new goals follow that same rule rather than inventing a
  second one: `(atom? "42")` succeeds and `(number? "42")` does not, and
  `(between 1 10 "5")` is refused with a message that says plainly that
  a text cell reading 5 is not the number 5.

- **Comparing a text cell to an unquoted name no longer answers
  silently.** This is the one behaviour change in this release that can
  affect a program that used to run.

  A text cell reading `eng` and the bare word `eng` written in a query
  are different values in PROLOG — that has always been true, and it is
  still true. The problem was that nothing said so. Asking
  `(\= Dept eng)` — "is this department something other than eng" —
  answered **yes for every row of the table**, because the cell's value
  and the word you typed were never the same thing to begin with. The
  answer was correct about the values and useless as an answer, and
  there was no way to tell that had happened.

  Those four comparisons — `=`, `\=`, `==`, `\==` — now stop and explain
  when the only difference between the two things being compared is the
  quoting:

  ```
  the text "eng" and the name eng are different things in PROLOG - a
  text cell keeps its quotes, so if you meant the cell's own value
  write it as "eng".
  ```

  It is deliberately narrow. Two values that genuinely differ still
  compare quietly, exactly as before, so ordinary filtering is
  untouched. And it applies **only** where you explicitly asked whether
  two things are the same — never while Frazaro is matching rules
  against facts, because a query that finds a real answer must never be
  stopped by a near-miss against some *other* fact it was not asking
  about.

  The same message appears for a text `42` against a numeric `42`, where
  it says "number" rather than "name".

- **And a query that finds nothing at all now tells you why, if the
  reason is that quoting.** The same mistake shows up a second way: not
  as a comparison that answers oddly, but as a query that just comes back
  empty. `(query (emp N eng))` against a table whose department column is
  text used to return nothing and say nothing, and an empty result is the
  hardest thing to debug because it looks the same whatever caused it.

  Frazaro now checks, **only when a query found no rows whatsoever**,
  whether one of the things you asked for differs from something it
  stored by nothing but the quoting — and if so, says that instead of
  handing back a blank.

  It runs after the search, never during, so it cannot affect a query
  that works: if your query finds even one row, this check does not
  happen at all. Adding a fact that matches makes the message disappear.
  And an empty result with no such near-miss in it stays exactly as it
  was — empty, and silent, because sometimes the honest answer is that
  there is nothing there.

  Two cases it does not catch, worth knowing so the silence is not
  mistaken for a clean bill of health: a mismatch that only exists inside
  a rule's body, and one that only appears after an earlier part of the
  query has filled in a variable. Both need bookkeeping during the search
  that would slow down every query to help a few, so they are left for a
  later release.

- **A query no longer loses a column when your data is shaped like a
  goal.** If a fact stored a value such as `(not bob)` or `(atom? bob)` —
  ordinary data that happens to be written in the same shape as a
  Frazaro goal — a query reading it found the value correctly and then
  left it out of the result. One column where two were due.

  This one is worth calling out because of how it failed: nothing was
  reported, and the answer that came back was not empty or wrong-looking,
  just *narrower* than it should have been. A missing row is obvious; a
  missing column is easy to read straight past. Fixed, and both columns
  now appear.

- **A release check learned to look both ways.** PROLOG refuses to let
  you define a predicate with a reserved name, and a static check has
  been holding three places to agreeing on which names those are. It
  turned out to be checking only one direction — it would catch a name
  that was reserved but did nothing, and miss a name the solver acted on
  without ever reserving it, which would let you define a predicate that
  is then silently ignored. Nothing shipped had that fault; the check
  simply could not have caught it. It now checks both directions, and
  was verified by deliberately introducing the fault to confirm it is
  caught.

- **PROLOG has lists, and `findall` stops being a dead end.** Until now
  `findall` handed you a bag of answers and there was nothing you could
  do with it. You could put it in a cell and read it. You could not count
  it, add it up, pick the third one out of it, or ask whether something
  was in it. That was the largest missing piece in the engine, and it is
  the piece that made `findall` — the goal that gathers everything
  matching a pattern — worth much less than it looks.

  Six new goals: `(length L N)`, `(member X L)`, `(nth N L X)`,
  `(append A B C)`, `(reverse L R)` and `(sum-list L N)`. The one most
  people will reach for first:

  ```
  (fact (sale north 10)) (fact (sale north 20)) (fact (sale south 30))
  (query (findall Amount (sale north Amount) Bag) (sum-list Bag Total))
  ```

  That gathers every northern sale and totals it — 30 — in one query,
  which simply could not be written before.

- **Most of these goals run in more than one direction.** That is the
  part worth knowing, because it is what makes them worth having rather
  than being six functions with awkward spelling.

  `(member X L)` **tests** membership when `X` is already filled in, and
  **produces one row per element** when it is not. `(nth N L X)` picks
  the `N`th element when you name a position, and walks the whole list,
  position beside element, when you leave `N` blank. `(length L N)` and
  `(sum-list L N)` compute a count or a total when `N` is blank and
  **check** one when it is not.

  `(append A B C)` is the one that goes furthest. Given two lists it
  joins them. Given the *result* and either half, it hands back the
  other half — so `(append (list a) Rest Whole)` drops a known prefix,
  and `(append Front (list z) Whole)` drops a known suffix.
  Given only the result, it produces **every way the list could be split
  in two**, one row per split, both ends included. None of that is extra
  machinery; it is the same goal read in different directions, which is
  what a logic language is for.

  What it will not do is work backwards from nothing. `(append A B C)`
  with all three blank has nothing to join and nothing to split, and is
  refused by name rather than left to find nothing.

- **How a list is written, and why it is not `[a, b, c]`.** Write a list
  as `(list a b c)`. The empty list is `nil`.

  Square brackets are not available, and the reason is worth knowing
  rather than guessing at: Frazaro's reader treats `[`, `]` and `|` as
  ordinary letters, exactly like `a` or `x`. Writing `[H|T]` would not
  produce a list — it would produce one long word spelled `[H|T]`.
  Teaching the reader about brackets means changing it for all five of
  Frazaro's languages at once, which is a far larger change than this
  one.

  Underneath, a list is built from pairs — `(cons Head Tail)` — and
  `(list a b c)` is shorthand for `(cons a (cons b (cons c nil)))`. You
  can write either; they are the same list, not two kinds of list. The
  long form is what makes **taking a list apart** possible, because a
  pair is an ordinary term and matching is what this language already
  does:

  ```
  (query (findall X (color X) Bag) (= Bag (cons First Rest)))
  ```

  `First` is the first colour and `Rest` is everything after it. The same
  works in a rule's head, so a rule can be written to match only
  non-empty lists. That is the whole trade: `(list …)` to write one down,
  `(cons …)` to pull one apart.

  One thing `(list …)` is not is a goal. It builds a value, so it belongs
  in an argument. Writing `(query (list a b))` on its own stops and says
  so rather than quietly finding nothing.

- **This changes what a `findall` bag looks like in a cell.** A bag of
  three colours used to display as `(red green blue)`. It now displays as
  `(list red green blue)`. If you have a sheet that reads a bag as text,
  that text has changed.

  The old display was shorter by two characters and was also a small lie:
  you could not paste `(red green blue)` back into a program and get the
  same list — Frazaro would read it as *a thing called red with two
  parts*. The new display reads back as exactly the list it prints, which
  is the property the old one lacked.

  A list is only ever displayed this way when it really is a list. A pair
  whose tail is not a list — `(cons a b)`, which is a perfectly legal
  term and not a list at all — displays as itself, so the display can
  never make something look like a list that is not one.

- **The empty list is `nil`, and that fixed an old wrong answer.**
  A `findall` that finds nothing used to hand back `()`, and asking
  `(compound? Bag)` about it answered *yes* — because internally it was
  an empty structure. Standard Prolog says the empty list is a simple
  value, not a structure. Now that it is `nil`, `(compound? Bag)` answers
  *no* and `(atomic? Bag)` answers *yes*, which is the standard answer.
  Nothing in the type tests changed to make that happen.

- **These goals do their work in one step, on purpose.** A whole query
  gets a fixed budget of resolution steps. Written the traditional way —
  as Prolog rules that call themselves — `length` and `append` would
  spend one step per element and run out of budget on a list of about a
  hundred, which is exactly the size a real `findall` bag reaches. So
  they are built into the engine instead: counting, reversing and summing
  a list cost one step no matter how long it is. `member`, `nth` and
  `append` still cost one step per answer they hand back, because each of
  those really is a separate answer.

- **When something is not a list, Frazaro says so rather than finding
  nothing.** `(length (color red) N)` does not quietly return no rows; it
  stops and tells you it needs a list, and shows you what a list looks
  like. That includes a half-built list — `(cons a T)` where `T` has not
  been filled in yet — which some Prologs would complete for you and
  Frazaro deliberately refuses, because a query that quietly finds
  nothing is indistinguishable from one that correctly found nothing.

  Seven names are now reserved and cannot be used for your own
  predicates: `length`, `member`, `nth`, `append`, `reverse`,
  `sum-list` — and `list`, which is reserved because it is the shorthand
  marker rather than because it is a goal. `member` is the one worth
  flagging, since it is a natural name for a fact about people.

  `cons` and `nil` are **not** reserved. They are ways of writing data,
  so a predicate of your own may still be called `cons`.

### Known open security items

**Closed this release:** none — 0.5.4 is a feature release and touches no
security item.

**Still open:** `SEC.3`, `SEC.7`, and two from the 2026-09-08 code
review — `SEC.10` and `SEC.15`. In plain words:
the remembered raw-VBA consent record still lives inside the workbook
(`SEC.10`);
formulas a program writes are not screened for functions that reach the
network (`SEC.15`); and effects like sending mail still run without a
permission prompt (`SEC.7`). `SEC.8` narrows that last one — it gates on
where the *workbook* came from — but does not close it: a phrasebook loaded
into a workbook of your own still reaches those verbs unprompted.

**Assessed and accepted, not fixed:** `SEC.12`, `SEC.14`, `SEC.16` and
`SEC.17`. Each needs a precondition an ordinary install does not meet —
mostly an Excel setting that ships off and that Frazaro never asks you to
turn on. The reasoning for each, and what would reopen it, is written down
rather than left implied.

The authoritative lists, kept current in one place instead of copied into
every release: [`README.md`](../README.md) in plain words, and
[`docs/BETA_ROADMAP1.md`](BETA_ROADMAP1.md) with the file, line, fix and
disposition for each.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from — and treat a workbook someone sent you the
same way before you press Interpret.** Frazaro makes no network call and
does not update itself; check the README's *Known open security items*
when you return for a newer build. Vulnerability reports:
`docs/SECURITY.md`. Everything else: `docs/SUPPORT.md`.

## 0.5.3

### What changed

- **`SEC.8` — a workbook from the internet can no longer reach outside
  itself.** Office has blocked macros in files that came from the
  internet since 2022. That block reads the Mark-of-the-Web that Windows
  puts on a downloaded or emailed file, and it applies to macros — VBA
  code. A Frazaro program is not a macro: it lives in the cells. So a
  plain `.xlsx` someone mails you could carry a complete program, and
  after the ordinary "Enable Editing" click Frazaro would run it with the
  add-in's own privileges. Office's protection had never applied to it,
  because from Office's point of view there was nothing there to protect
  you from.

  Frazaro now reads that same mark itself, and refuses the things that
  reach outside the workbook: opening or saving a file, saving a copy,
  printing, exporting a PDF, composing an Outlook email, and
  password-protecting or unprotecting a sheet. Everything else runs
  normally — reading and writing cells, formatting, formulas, loops, the
  whole ordinary business of a program. A workbook from the internet
  still opens, still calculates, and can still be edited exactly as
  before. It just cannot reach off the page.

  The refusal names the thing the program tried to do, so you find out
  what a workbook wanted rather than only that something was blocked.

  **Clicking "Enable Editing" does not grant this.** Excel shows its own
  yellow banner on a file from the internet — that bar is Excel's, not
  Frazaro's, and it governs whether you can type in the sheet. Frazaro's
  check is separate and happens later, when a program actually tries to
  reach outside the workbook. Enabling editing leaves the file's origin
  mark exactly where it was, which is why the two decisions stay
  independent: you can read and edit a workbook someone sent you without
  also agreeing that its program may email, print, or save files. Only
  unblocking the file in Windows does that.

  **To allow it, you unblock the file in Windows** — close the workbook,
  right-click the file in File Explorer, choose Properties, tick
  *Unblock*, then reopen it. Frazaro deliberately offers no button of its
  own for this. That is the whole point of the design: a workbook that
  arrives from outside must not be able to carry its own permission slip,
  and anything Frazaro stored could be forged or shipped inside the
  workbook itself. Windows' own checkbox is the one channel a mailed file
  cannot reach.

  Two limits worth knowing, both deliberate. A workbook opened straight
  from a `https://` SharePoint address refuses external effect too:
  Windows cannot record an origin mark in that kind of location, so
  Frazaro cannot tell where the file came from, and it does not guess.
  (A OneDrive folder that syncs to your own disk is an ordinary local
  folder and is not affected — that is the common case.) The workaround
  is to save a local copy. And a file with no mark at all is treated as
  local, because an ordinary file you made yourself has no mark either
  and the two are genuinely indistinguishable — the protection comes from
  the mark being *present*, never from it being absent.

- **A new release check: `tools/check_sec8_provenance_gate.ps1`.** The
  self-test suite does no Office automation and cannot manufacture a file
  that claims to be from the internet, so it can check every part of the
  *decision* but never the call sites. This static check runs at every
  release and fails it if any of the nine external-effect sites loses its
  guard, or if the one line that reads the workbook's origin goes
  missing. It is mutation-tested in both directions rather than assumed
  to work, and it found a site the audit had missed on its first run.
  What it cannot do is prove the guard *functions* against a real marked
  workbook — only Excel can do that, so that rests on a live test.

- **Email problems now tell you what went wrong instead of dropping you
  into the code editor.** If Outlook was not available, or an attachment
  path did not exist, an email sentence did not report that — it stopped
  the program with Visual Basic's own `Run-time error '5'` dialog, the one
  with a *Debug* button. The message Frazaro meant to show you ("Could not
  start Outlook to create the email") existed the whole time and never got
  the chance to appear. This affected every release that has had the email
  sentence, and needing no Outlook installed is all it took to trigger it.

  The cause is a VBA quirk this project had already documented and proved
  with a standalone test: the mechanism the interpreter used to reach that
  particular helper does not carry an error back to the code that would
  have caught and displayed it. The email helper now uses a direct call
  instead, so its messages arrive the way every other refusal in Frazaro
  does.

  **Not fixed in this release, and named so it is not mistaken for
  shipped:** eight other helpers still reach you the same wrong way when
  they fail — an unrecognised colour, a missing key, and six pivot-table
  and worksheet operations with an unrecognised option. Those show
  Visual Basic's dialog rather than a Frazaro message. Only the email
  helper is fixed here, because only it was in the way of `SEC.8`; the
  rest are recorded in
  [`docs/BETA_ROADMAP1.md`](BETA_ROADMAP1.md) with the full list rather
  than fixed in a hurry alongside a security change.

- **A build defect found by a new release check: `VlaSlice` was never
  shipped.** Frazaro keeps two lists of modules — one for what a built
  add-in contains, one for what the development workbook reloads — and
  nothing had ever checked that they agree. `VlaSlice`, a small internal
  class the language core names directly, was in the second list and not
  the first, so a freshly built add-in would not have contained it at
  all. It has been in that state since `0.5.0`. It is now in both lists.

  This is the ninth time this project has made the same mistake, and the
  first time a machine caught it rather than a person hitting the
  resulting error. `tools/check_devrig_mods_parity.ps1` now runs at every
  release and fails it if either list is missing something the other has,
  in either direction — the more dangerous direction being a module the
  development workbook compiles happily while a built add-in would not,
  which is exactly what stayed hidden here. Modules that genuinely should
  not ship (the test suites, the builder itself) are listed in the check
  with a reason beside each. It is mutation-tested in both directions.

- **`SEC.13` — Word documents are no longer opened with macros enabled.**
  *Import Program File…* accepts Word documents, and it opened them
  through Word automation without setting Word's own
  `AutomationSecurity`. Word's default in that mode is *Low*, so a
  `.docm` carrying an `AutoOpen` or `Document_Open` macro ran that macro
  silently the instant Frazaro read the text out — no Trust Center
  prompt, because a document opened by automation does not get one. This
  was a live, one-click path in every shipped edition, on the menu and
  the ribbon both, and worse than the audit first recorded: the import
  router matches `.docm` by name, not only `.docx`/`.doc`.

  Frazaro now force-disables macros for the duration of the read, and the
  guard is deliberately set *after* the error handler is armed, so if it
  cannot be set the import refuses rather than opening the document
  unguarded. Because Frazaro reuses a copy of Word you already have open
  rather than always starting its own, it captures your Word's previous
  setting and puts it back afterwards — an application it does not own is
  handed back as it was found. If that restore should ever fail it stays
  quiet on purpose: the failure leaves Word *more* cautious than before,
  never less, and a Word restart clears it.

  Word documents still import exactly as they did — the text, the
  typography cleanup, everything. The only thing that changed is that a
  document's own macros no longer get to run on the way in.

  If you followed the previous advice in the README (*"import only Word
  files you wrote yourself, or paste the text instead"*), you no longer
  need to. That advice has been removed.

  **Not in scope, and named so it is not mistaken for shipped:** Frazaro
  still does not *ask* you before opening a document you point it at.
  Consent prompts are `SEC.7`/`SEC.8`'s subject and remain open.

- **Importing a Word document can no longer appear to hang Excel.** When
  Frazaro starts its own copy of Word to read a document (rather than
  reusing one you already have open), that copy is invisible. If Word
  decided to *ask* you something about the file rather than simply fail
  — a damaged document, a file-conversion prompt — the question appeared
  on a window you could not see or click, and Excel looked frozen with no
  way forward but Task Manager. Frazaro now tells the copy it starts not
  to raise alerts, so a document Word dislikes comes back as the ordinary
  refusal message instead of a hidden prompt.

  A copy of Word you already had open is deliberately left alone. Its
  windows are visible, so its questions are answerable — and silencing
  alerts in an application Frazaro did not start would cost you warnings
  you should see. That is the opposite direction of failure from the
  macro guard above, which is why the two are treated differently.

  Found while writing the live test for `SEC.13` rather than from a
  report: it never actually fired during testing. It is fixed as a
  hazard, not as an observed fault.

- **A new release check: `tools/check_word_automation_security.ps1`.**
  The self-test suite does no Office automation at all, so nothing in it
  could ever have caught this or its return. A static check now runs at
  every release and fails it if any code that opens a Word document does
  not force-disable macros first, in that same procedure, before the
  open. It is mutation-tested in both directions rather than assumed to
  work. What it cannot do is prove the guard *functions* — only Word can
  do that, so that rests on a live test, with a reproducible fixture
  recipe recorded in `tools/sec13_word_fixture.md`.

- **`SEC.11` — the fingerprint that remembers your consent is now a real
  one.** When a phrasebook contains `raw` — VBA written directly into a
  grammar file — Frazaro asks before loading it, and can remember your
  answer. What it remembers is tied to a fingerprint of that file's exact
  contents, so that editing the phrasebook at all asks you again.

  The fingerprint was too weak for the job. It was a simple arithmetic
  checksum, and checksums of that kind can be *aimed*: someone who wanted
  a different phrasebook to carry your fingerprint could adjust a couple
  of characters inside a comment until the numbers matched. Your "yes" to
  a file you had read could then have been silently inherited by a file
  you had never seen. It is now SHA-256, the same standard used for
  software signatures, where aiming at an existing fingerprint is not
  something anyone knows how to do.

  **You will be asked once more for phrasebooks you had already
  approved.** Old answers were filed under the old fingerprint and cannot
  be matched to the new one. Being asked again is the safe direction to
  fail, so nothing tries to convert them; the old entries are simply left
  alone and ignored.

  Two things deliberately did *not* change. The fingerprint still ignores
  spaces, tabs and line breaks, so re-indenting a phrasebook, or opening
  one that was saved on a different operating system, still counts as the
  same file rather than sending you back through the question. Changing
  what a phrasebook *says*, by even one character, still does.

  And exported *Expanded Phrasebook* files you already have keep
  working. The freshness check reads the older form, confirms the file is
  current, and tells you the next export will upgrade it. A re-exported
  file carries the new form, written out as `sha256:` followed by the
  digest so the two generations can never be mistaken for one another.
  Nothing you already have needs regenerating.

  Built without depending on anything being installed. The obvious route
  was to borrow Windows' own cryptography through .NET. Measured on the
  development machine, that turned out to fail outright — the component
  is registered against a version of .NET that Windows 11 no longer
  installs by default — so a phrasebook would have become unloadable on
  an ordinary machine. Frazaro now computes SHA-256 itself, in about 150
  lines, which works the same everywhere and can be checked completely by
  the self-test suite against the published standard test values.

- **A new release check: `tools/check_hash_twin.ps1`.** The fingerprint is
  computed in two places — in Frazaro itself, and in the release script
  that verifies an exported phrasebook is current. The two must agree
  exactly, and until now the only thing keeping them in step was a comment
  saying so. This project has been bitten nine times by that same shape.
  Both sides are now pinned to the same published test values, so either
  one drifting fails a release rather than being noticed later. The two
  were also checked against each other on the real 89,446-byte English
  phrasebook and produce the identical fingerprint.

- **`SEC.9` — a phrasebook has to be approved on this computer before it
  can replace the built-in grammar.** A phrasebook decides what your
  sentences *mean*. Frazaro used to prefer a grammar file sitting next to
  the workbook you had open over its own built-in copy, and it would
  silently reload any phrasebook path a workbook remembered — including
  paths on other machines. Between them, a workbook could quietly decide
  what every sentence in it did.

  This was not theoretical. On 2026-09-08 an old `english.vla` left over in
  a Downloads folder was picked up ahead of the add-in's own copy, simply
  because the workbook being opened was in that folder too. It failed
  noisily, but only by luck: that copy was old enough not to know a word the
  program used. A phrasebook that was merely *different* rather than *older*
  would have loaded without a word and changed what the program did.

  Frazaro now asks, once, naming the file, before using a phrasebook that
  did not come with it. Your answer is remembered **on this computer**,
  not inside the workbook — a workbook someone sends you cannot arrive with
  its own permission already granted, because the permission was never
  something a workbook could carry. Choosing a phrasebook yourself through
  *Load Phrasebook* counts as the answer, so picking a file never asks you
  about it a second time.

  Saying no is safe and is not a dead end: Frazaro simply uses its own
  built-in grammar, which is complete. The answer is remembered either way,
  so you are not asked again on every command — and if the file itself
  changes later, you are asked again, because the answer was about the file
  you were shown, not about its name.

  **Network and web locations are described, never quietly contacted.** If a
  workbook asks for a phrasebook on a network share, the question says so
  and warns that opening it would hand your Windows sign-in to that server.
  Nothing touches the location until you say yes — not even a check for
  whether the file exists, which is itself enough to leak that sign-in.

  **Changing your mind: *Forget Phrasebook Approvals*,** in the Utilities
  group. It lists every answer this computer has recorded and clears them
  all, so the next time each phrasebook is used you are asked again. Your
  answers are otherwise kept for good, which is deliberate — a phrasebook
  you declined stays declined rather than asking you the same question
  every time you open Excel, because a dialog that keeps reappearing is one
  people learn to click through without reading.

  Your recorded answers also appear in *Copy Diagnostic Report*, so if a
  sentence stops being understood you can see whether a phrasebook it
  needed was declined. That matters because the symptom on its own is
  indistinguishable from a typo: without it, Frazaro would just say it does
  not understand the sentence and never mention the phrasebook.

- **A new release check: `tools/check_sec9_phrasebook_gate.ps1`.** The
  self-test suite can check the whole decision — whether a path is remote,
  whether it belongs to Frazaro itself — but it cannot click a dialog or
  read the saved answers, so the places that *call* the check are where this
  could quietly stop working. A static check now runs at every release and
  fails it if either loader loses its gate, if the two decision functions
  are removed, or if the approval is ever moved to run after the
  file-existence probe rather than before it. That last one is the one that
  matters: reordering those two lines would restore the credential leak
  while leaving every visible behaviour, and every test, exactly as it was.
  It is mutation-tested in both directions rather than assumed to work.

- **Sixteen things a program can get wrong now say so, instead of dropping
  you into the code editor.** Type a colour Frazaro does not recognise, ask
  for a key that was never stored, or name a pivot table that is not there,
  and until now Excel's own *Run-time error '5'* box appeared, with a
  **Debug** button that opened the VBA editor at a line of Frazaro's
  internals. Frazaro had written a perfectly clear explanation for each of
  these — it just never reached the screen. You now get the ordinary Frazaro
  message saying what was wrong with what you wrote: *'notacolor' is not a
  color - use "#RRGGBB", like "#FF69B4"*, or *there is nothing stored at key
  'no-such-key'*, or *no pivot table named 'nosuchpivot' in this workbook.*

  The cause was one mechanism, not sixteen separate bugs. Frazaro reaches
  most of its built-in helpers through a general-purpose Excel facility
  that, it turns out, does not carry an error back to the code that asked
  for it. Any helper whose job includes saying no was therefore unable to
  say no. The helpers that only compute something were unaffected, which is
  why this took so long to notice: the failure was invisible until you made
  a mistake.

  Eight of the sixteen were found by hand. The other eight were found by the
  release check below, and had been missed — they refuse a misspelled pivot
  table name through a shared piece of code rather than in their own, which
  is exactly the kind of thing a person reading down a list does not see.

- **A second fix underneath it: Frazaro no longer decides whether something
  is one of its helpers by trying it and seeing what happens.** It now checks
  the name against the list of helpers first. The old approach could not tell
  "that is not a helper at all" apart from "that is a helper, and it is
  refusing" — so a deliberate refusal could have been reported as *"'vlacolor'
  is not a form..."*, naming the wrong problem with complete confidence. Being
  told the wrong thing firmly is worse than being told nothing, and this
  removes the possibility rather than making it less likely. A genuinely
  unknown name still gets exactly the same "not a form, place helper, dotted
  global, built-in, or VLA_Runtime helper" message it always did.

- **A new release check: `tools/check_runtime_raise_dispatch.ps1`.** The
  boundary this fix relies on — a helper that can refuse must be called the
  direct way — was being kept in someone's head, and had already slipped
  three times, each time found by a user hitting the crash. The check now
  works it out from the source: it reads every helper, follows the shared
  code they call, and fails the release if any helper that can refuse is
  still reached the broken way. It was written before the fix, so its first
  run listed exactly the work to do, and it is mutation-tested in both
  directions rather than assumed to work. It is what found the eight the
  hand count missed.

- **`PROLOG.7` — you can now compare two numbers in a rule.** PROLOG could
  already *calculate* — `(is Total (+ X Y))` works out a value and gives it
  a name. What it could not do was *test* one number against another. There
  was no way to write "salary over 80000", so the staffing example in the
  README had to match an exact department instead of a threshold, which is
  not what anyone actually wants to ask.

  The six comparisons now work as goals in their own right: `<`, `>`, `=<`,
  `>=`, `=:=` (equal) and `=\=` (not equal). You write them like any other
  goal in a rule body or a query:

  ```
  (rule (well-paid Name) (employee Name Salary) (> Salary 80000))
  ```

  A comparison is a question, not a calculation. It either succeeds and the
  search carries on, or it fails and that row simply isn't in the answer —
  it never invents a value or fills in a variable. Both sides can be
  arithmetic, so `(> (+ Base Bonus) 80000)` is fine.

  **A refusal now names the form you actually wrote.** Comparisons share
  the machinery that `(is ...)` uses to work out each side, and that shared
  code used to say "(is ...)" in every complaint it made. So a program
  containing `(> Salary 80000)` and no `(is ...)` at all could be told its
  problem was with `(is ...)` — pointing at a form the author never typed.
  Five separate messages did this. They now name whichever form is really
  running, and a check runs on every release to keep it that way. Nothing
  about the wording changed for `(is ...)` itself.

  The six symbols are now reserved, so they can't be used as your own
  predicate names — the same rule that already applies to `is`, `not`,
  `findall` and `!`. Unification (`=`) and structural equality (`==`) are a
  separate item, built next (below); `=:=` here is numeric equality only.

- **`PROLOG.8` — you can now match two things against each other in a
  rule.** Until now a variable only ever picked up a value one way: by
  lining a goal up against a stored fact. There was no way to say, in the
  middle of a rule, "let X be this" or "only if these two are different".
  Four goals now do that:

  ```
  (rule (corner C) (= C (point 0 0)))
  (rule (rival A B) (player A) (player B) (\= A B))
  ```

  `=` **matches, and fills in blanks while it does.** `(= X 1)` gives X the
  value 1. It works in both directions and reaches inside structures, so
  `(= (point X 7) (point 3 Y))` sets X to 3 and Y to 7 in one step.

  `\=` **succeeds when two things do NOT match** — the usual way to say
  "these must be different". It fills nothing in, whichever way it turns
  out.

  `==` **and** `\==` **ask a stricter question: are these already the same
  thing?** They never fill anything in. `(== X 1)` is false when X is still
  blank — X *could* become 1, but it isn't 1 yet — where `(= X 1)` would
  make it so. That difference is the whole reason both exist: use `=` when
  you want to set something up, `==` when you want to check it without
  changing anything.

  **These are not the number comparisons.** `(= 2.0 2)` is false: `=`
  matches things as written, and `2.0` is not written the same way as `2`.
  If you mean "the same number", that is `=:=` from `PROLOG.7`. The same
  goes for `==`.

  **Refusals name the goal you actually wrote.** Writing `(\= 1)` tells you
  about `(\= ...)` — not about `(= ...)`, `(not ...)` or `(is ...)`. `\=`
  means the same thing as "not equal to", and the shortest way to build it
  would have been to quietly rewrite it into `(not (= ...))`, but then your
  mistake would have been reported against a form you never typed. It is
  built directly instead, so the message can name what you wrote. This is
  the same principle `PROLOG.7` applied to the shared arithmetic messages.

  **Matching text that came from a table: quote it.** A word in a cell is
  stored as *text*, and a bare word in a query is a *name* — they are not
  the same thing, so they do not match. If your table has a `Dept` column
  reading `eng`, write `(== D "eng")`, with the quotes. Without them the
  test quietly succeeds on every row for `\=`, and on none for `==`, and
  both answers are technically right — you asked about a name, not about
  the text in the cell. Numbers need no quotes; this applies to text only.
  The rule is not new — it is what stops a cell reading `Alice` being
  mistaken for a blank waiting to be filled in — but these four goals are
  the first ones that make comparing against a written-out value an
  everyday thing to do, so it is worth saying here.

  All four symbols are now reserved and can't be used as your own predicate
  names, and the refusal that says so lists them. A release check now holds
  three things to agreeing about that list: what is refused as a name, what
  the solver actually runs, and what the message tells you — so a name
  can't end up forbidden but inert, or advertised but not really reserved.

### Known open security items

**Closed this release:** `SEC.8`, `SEC.9`, `SEC.11` and `SEC.13` — see above.

**Still open:** `SEC.3`, `SEC.7`, and two from the 2026-09-08 code
review — `SEC.10` and `SEC.15`. In plain words:
the remembered raw-VBA consent record still lives inside the workbook
(`SEC.10`);
formulas a program writes are not screened for functions that reach the
network (`SEC.15`); and effects like sending mail still run without a
permission prompt (`SEC.7`). `SEC.8` narrows that last one — it gates on
where the *workbook* came from — but does not close it: a phrasebook loaded
into a workbook of your own still reaches those verbs unprompted.

**Assessed and accepted, not fixed:** `SEC.12`, `SEC.14`, `SEC.16` and
`SEC.17`. Each needs a precondition an ordinary install does not meet —
mostly an Excel setting that ships off and that Frazaro never asks you to
turn on. The reasoning for each, and what would reopen it, is written down
rather than left implied.

The authoritative lists, kept current in one place instead of copied into
every release: [`README.md`](../README.md) in plain words, and
[`docs/BETA_ROADMAP1.md`](BETA_ROADMAP1.md) with the file, line, fix and
disposition for each.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from — and treat a workbook someone sent you the
same way before you press Interpret.** Frazaro makes no network call and
does not update itself; check the README's *Known open security items*
when you return for a newer build. Vulnerability reports:
`docs/SECURITY.md`. Everything else: `docs/SUPPORT.md`.

## 0.5.2

### What changed

- **`SEC.1` — the `CallByName` reflection fallback removed from dynamic
  dispatch, built and owner-verified live.** Previously, any `.`-member
  name the interpreter didn't recognize natively fell through to VBA's
  own `CallByName`, reaching the real Excel object model with whatever
  name and arguments a phrasebook rule supplied — a heuristic, not a
  capability gate. A full-repo census found 25 members reached only
  through it, well beyond the G-TABLES surface this item's own design
  anticipated: plain cell `.value` reads, cross-sheet `.range` lookups,
  the `listobjects`/`listrows`/`listcolumns` chain, and several
  housekeeping macros. Each got its own fixed, audited native case before
  the fallback itself came out, so anything not on that list now refuses
  in words instead of reaching arbitrary late-bound dispatch. Two
  corrections found live, not assumed clean: a corpus-only first census
  missed `.value`/`.range` entirely (caught only because the host-test
  suite reads them, not the shipped phrasebook text), and a second pass
  then found every property already native for *writing* (`name`,
  `bold`, `size`, and 17 others) had never been native for *reading* —
  fixed by mirroring the write-side list into the read side in one pass
  rather than chasing each individually. `VLA_SELF-TESTS` pure 947/947,
  host 143/143; `VerifyReports` emitter 141/141, interpreter 141/141,
  both live in Excel. G-PIVOT never used this mechanism at all (it
  dispatches through dedicated runtime helpers); keyword-argument calls
  already had no fallback to begin with. Permissioned, declared
  capabilities with real external effect (`vlasendmail` today) are a
  separate, still-open item — see `SEC.7`. Full mechanism:
  `docs/BETA_ROADMAP1.md`'s own SEC.1 entry.
- **`SEC.2` — `raw` behind explicit, per-phrasebook consent, built and
  owner-verified live.** A phrasebook using `raw` (literal VBA,
  previously unconsented) now shows a modal, naming the file, before it
  loads from disk; declining refuses the whole phrasebook, not just the
  `raw`-bearing rules. Two remembered scopes, an explicit choice rather
  than a silent default: *this workbook only* (safer — forging it needs
  write access to that one file) or *every workbook on this device*
  (more convenient, a wider target, named as such in the prompt
  itself). Gated at the file-path loader specifically, not the shared
  `EnglishLoadVocabularyText` primitive `VLA_Browser.bas`'s
  already-shipped host-free translate API calls directly and documents
  as never showing a dialog — an early draft got this wrong and only
  passed the purity ratchet on a technicality, caught before it
  shipped. No test-bypass toggle anywhere in the mechanism, by design.
  Full mechanism: `docs/BETA_ROADMAP1.md`'s own SEC.2 entry.
- **`GO.6` — a working Load Phrasebook button, owner live-tested.**
  Found live while hand-verifying `SEC.2`: the only mechanism that
  technically loaded an external phrasebook (`VLA_IDE.IdeVocabPath`'s
  four candidate paths) was undocumented, built for internal
  edition/dev purposes, and REPLACED the base corpus by exact filename
  match rather than adding to it — no ribbon command existed for an
  org admin or community contributor to load their own. The new "Load
  Phrasebook" button calls `EnglishLoadVocabulary` directly, so it
  inherits `SEC.2`'s raw-consent gate and `G3`'s same-shape-collision
  refusal automatically, no second loading path. ADDS rather than
  replaces — `G3`'s existing cross-file override mechanism already
  resolves collisions between sources, `GO.1`'s ratified precedence,
  no interpreter change needed — persists per workbook (one
  `VLA_LoadedPhrasebooks` custom document property, an unbounded
  vbLf-joined list; no cap, the same as a source file's own import
  statements), and shows what's currently loaded after every load
  (`EnglishLoadedSourcesReport`). A moved or deleted remembered
  phrasebook is skipped with a note rather than blocking every other
  command; a genuine content collision still refuses, unchanged. Full
  mechanism: `docs/BETA_ROADMAP1.md`'s own GO.6 entry.
- **The one `AS.1` gap closed: `paint cell {r:text}`.**
  `check_rule_coverage.ps1`'s first real report (after `GEXPANDERLINT.0`
  verticalized the phrasebook artifact) found a rule with zero
  test-success proofs. It had never worked: the rule called `vlacolr`
  (no "o"), a name that resolves nowhere in the shipped modules, and
  had no cell/range slot in its pattern at all, so it could never have
  painted a cell even with the spelling fixed. Corrected to the
  `set-fill-color` idiom every sibling color rule already uses
  (`pirate.vla`'s own "paint cell" rule confirmed the intended
  semantics) and given its missing test. 150/150 phrasebook rules now
  carry at least one proof.
- **`CO.4` — versions are now something Frazaro can compare, not just
  print.** Groundwork, with no button and nothing new to say yet: it is
  what a future phrasebook will be checked against when it declares a
  minimum version (`requires: version:0.5.1`), and what a generated
  module's own version stamp will mean once it carries one. The
  decision behind it is the substance. **The grammar's compatibility
  version is the release version you already see** — the one in Copy
  Feedback and in Add/Remove Programs — under the `MAJOR.MINOR.PATCH`
  rules this project already wrote down (`SD-14`). There is no second,
  separate "grammar version" to learn, and deliberately so: `PATCH` is
  *defined* as no observable change to what your existing sentences do,
  `MINOR` means something new became sayable, and `MAJOR` means a
  sentence that once shipped changed meaning or a program could behave
  differently after upgrading. Those rules already say everything a
  compatibility check needs, so inventing a parallel number would only
  create two things to keep in step. Scoping also corrected a
  long-standing internal misreading: `VLA_CORE_VERSION` looked like a
  grammar version and is not one — it, and 23 sibling constants, name
  the work item that last touched each module, and three modules
  legitimately share one value today. It is unchanged, and no
  compatibility check reads it. Ordering is plain numeric
  `MAJOR.MINOR.PATCH`, so `0.9.0` correctly sorts *before* `0.10.0`
  rather than after it the way plain text comparison would; a version
  that cannot be read (`banana`, or a `-beta` suffix, which this
  project has never used) is refused in words rather than being quietly
  treated as either satisfied or unmet, since both hide the typo.
- **`CO.6` — a written record of when each thing you can say became
  sayable.** New file, and unlike `CO.4` above this one is meant to be
  read: [`docs/GRAMMAR_SINCE.md`](GRAMMAR_SINCE.md) lists every phrasebook
  rule and every core form with the release it first **worked** in — 278
  entries, almost all of them `0.5.0`. It answers the question `CO.4`
  leaves open. Knowing that versions compare correctly does not tell you
  *which* version you need, and the version number alone is a loose
  answer: a release can go up because a ribbon button was added, which
  tells you nothing about whether a particular sentence will work. This
  file is the precise answer, and the thing a future `requires:`
  declaration will be checked against.

  **It records when a form first worked, not when it was first spelled** —
  a distinction with a real case in it. `paint cell` appears here at
  `0.5.2`, not `0.5.0`, even though the words shipped in `0.5.0`: as the
  `AS.1` entry above describes, that rule never once painted a cell. Dating
  it `0.5.0` would tell you a phrasebook using it runs on `0.5.0`, which is
  false. Entries are never edited afterwards, so a date that went in wrong
  would stay wrong — which is why the seed was checked against the actual
  release tags rather than taken from the generated phrasebook artifact.
  That check earned its keep: the artifact looked like it gained three
  rules in `0.5.1`, and it had not — those three were already sayable in
  `0.5.0` and the artifact had simply been stale, the same staleness the
  `0.5.1` notes below record fixing.

- **`F.10` — a phrasebook can now state what it needs, and is refused
  politely when it doesn't have it.** Write a line like
  `(requires-version "0.5.2")` at the top of a phrasebook and Frazaro
  checks it *before* loading anything from that file. If the build is too
  old you get a plain sentence — a phrasebook asking for `0.9.0` on this
  build is refused with "this phrasebook needs Frazaro 0.9.0 or newer;
  this is 0.5.2" — instead of rules that load and then mysteriously do
  the wrong thing. This is what
  [`docs/GRAMMAR_SINCE.md`](GRAMMAR_SINCE.md) above exists to be checked
  against: it tells a phrasebook author which version number to write.

  Nothing loads part-way. The check happens before the first rule is
  registered, so a phrasebook is either fully in or fully refused —
  and it does not matter where in the file the line sits.

  Two other kinds of requirement are understood and both currently
  **refuse**, on purpose rather than by omission.
  `(requires-capability "...")` is the permission system that is not
  built yet (`SEC.7`, below): since nothing can grant a capability, the
  honest answer to a phrasebook asking for one is no. `(requires-form
  "...")` — needing one specific sentence rather than a whole version —
  is understood but cannot be checked yet, because the record above
  lives in the source repository and is not carried inside Frazaro
  itself. A requirement Frazaro doesn't recognise at all is also
  refused: it cannot confirm the requirement is met, so it does not
  pretend to. Note this is a phrasebook *declaring* what it needs; it is
  not yet a restriction on what the verbs themselves may do — see the
  `SEC.7` item below, which is unchanged by this release.

  Because a version requirement is only as good as the record it is
  written against, the release checks now also refuse to let a form
  reach a release with no row in
  [`docs/GRAMMAR_SINCE.md`](GRAMMAR_SINCE.md). If you are writing a
  phrasebook, that means the version number you look up there covers
  every sentence Frazaro understands, not just the ones someone
  remembered to record.

### Known open security items

- **SEC.3** — generated code does not yet carry phrasebook provenance.
- **SEC.7** — a small set of verbs with real external effect (`vlasendmail`
  today) still runs with no permission check: a phrasebook you load can
  send mail with no consent prompt. `SEC.8` above narrows this but does
  not close it — it gates on where the *workbook* came from, not on what
  a phrasebook asked permission to do, so a phrasebook loaded into a
  workbook of your own still reaches these verbs unprompted.

`SEC.1`, `SEC.2` closed this release — see above.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from.** Frazaro makes no network call and does not
update itself; check the README's *Known open security items* when you
return for a newer build. Vulnerability reports: `docs/SECURITY.md`.
Everything else: `docs/SUPPORT.md`.

## 0.5.1 — 2026-09-07

**Pre-flight for the corporate push.** The first release cut by
`tools/release.ps1` rather than by hand, and the day the static ratchets
went green again. Nothing here changes what a sentence means; every item
is the machinery that keeps the next four weeks honest.

### What changed

- **The static ratchets pass, and are now run on every push.** `0.5.0`
  shipped past a red `check_raise_ratchet.ps1` (five raw `Err.Raise` sites
  had accumulated since the ceilings were set) and a `check_id_registry.ps1`
  that misread four mid-sentence bold ids as duplicate definitions. The
  four genuinely raw refusals in `VLA_Relation`/`VLA_Sql` now carry
  catalogue ids (`relation-table-noncontiguous-areas`,
  `sql-internal-*`); the one re-raise with no id to carry is documented as
  the eighth site of `VLA`'s ceiling; the id registry counts a bold id as
  a definition only when no prose precedes it on its line.
- **Three ribbon buttons find the phrasebook again** (EDITION-VOCABPATH).
  *Translate to VLA / to VBA*, *Export Expanded Phrasebook* and *Phrasebook
  Test Coverage* had reported "vocabulary file on disk... none found" in
  the dev workbook since the phrasebook moved under `scripts/polyglotta/`.
  They now search that folder as a last candidate; Check/Compile's own
  embedded-chain loading is byte-for-byte unchanged.
- **The expanded-phrasebook staleness stamp ignores whitespace.** It was
  the source file's byte size, so a Lint VLA re-indent or a line-ending
  change made `check_rule_coverage.ps1` demand a re-export, and Beta's
  history holds four commits that changed nothing but that number. It is
  now a hash of the source's non-whitespace bytes, computed identically
  in VBA and PowerShell; only a token change counts as stale.
- **`scripts/english_expanded.vla` re-exported** against the current
  corpus, so `check_rule_coverage.ps1` reports on the phrasebook that
  actually ships. It now lives beside its source under `scripts/polyglotta/`.
- **`release.ps1` refuses the placeholder** release-notes section instead
  of publishing it.

### Known open security items

Unchanged from `0.5.0`, and still the first engineering item in the queue:

- **SEC.1** — dynamic dispatch is not yet capability-gated: a phrasebook you
  load can reach roughly what a macro in a workbook you open could reach.
- **SEC.2** — `raw` phrasebook rules (literal VBA) run without a consent
  prompt. Refused on the default Interpret path; executes only through
  Compile/Export, behind Excel's own VBA-project trust prompt.
- **SEC.3** — generated code does not yet carry phrasebook provenance.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from.** Frazaro makes no network call and does not
update itself; check the README's *Known open security items* when you
return for a newer build. Vulnerability reports: `docs/SECURITY.md`.
Everything else: `docs/SUPPORT.md`.

## 0.5.0 — 2026-09-05

**A phrasebook for your workbook.** Frazaro lets people who are not
programmers automate Excel by writing English sentences that are *checked*
before they run, refuse with an explanation when they don't parse, and mean
exactly one thing when they do. This is a known-unfinished program, released
unfinished on purpose, because the only way to learn which sentences real
people reach for is to let real people reach for them. The open items are
counted, in order, in `docs/BETA_ROADMAP2.md`.

### Download

| File | Choose this if |
|---|---|
| `Frazaro_English.xlam` | The universal default. An Office document, not a program: works wherever Excel does, including managed machines. |
| `Frazaro_Espanol.xlam` | The Spanish edition (bilingual: every English sentence still works). |

Register the add-in once from inside the product (**Register for auto-load**
on the Frazaro ribbon) or through Excel's Add-ins dialog. Removal is one
button, **Uninstall Frazaro**. Full instructions: `docs/DEPLOY.md`. The
Windows installer is not part of this release; it returns in the next one.

### What works, owner-verified

The full English → VLA → interpret/compile pipeline; the worksheet IDE
(Check / Run / Undo / Known Sentences); control flow, value actions,
lookups, list and range operations; the four-layer error model; `=SQL()`,
`=DATALOG()` and `=PROLOG()` worksheet functions over your own tables. All
under a golden-file, self-test and dual-backend-parity discipline.

### Honest limits

Windows desktop Excel is the platform; Mac runs the interpreted path but
cannot export VBA. Grammar coverage is growing weekly and is not yet
complete for any one profession's vocabulary. No external users yet: the
project is looking for its first pilot.

### Known open security items

Frazaro's own threat model (`docs/THREAT_MODEL.md`) is public, and three of
its findings are open in this release:

- **SEC.1** — dynamic dispatch is not yet capability-gated: a phrasebook you
  load can reach roughly what a macro in a workbook you open could reach.
- **SEC.2** — `raw` phrasebook rules (literal VBA) run without a consent
  prompt. Refused on the default Interpret path; executes only through
  Compile/Export, behind Excel's own VBA-project trust prompt. First item
  after this release.
- **SEC.3** — generated code does not yet carry phrasebook provenance.

Until these close: **load phrasebooks only from people you would accept a
macro-enabled workbook from.** The phrasebooks embedded in these downloads
are audited at build time. Frazaro makes no network call and does not update
itself; check the README's *Known open security items* when you return for
a newer build. Vulnerability reports: `docs/SECURITY.md`. Everything else:
`docs/SUPPORT.md`.

### Licence

Engine Apache-2.0; phrasebooks MPL-2.0 (a phrasebook you write in your own
file is yours); the one module Frazaro copies into your workbook is 0BSD;
everything Frazaro generates from your sentences is yours outright
(`OUTPUT-EXCEPTION.md`). No warranty: beta software, as-is. Take a backup
before you run a program you have not run before.
