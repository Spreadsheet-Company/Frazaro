# vla-lang

**The VLA language with nothing around it: forms, their expansion, the
grid, the view record, and the machine that steps a grid.**

Frazaro is a language for spreadsheet work: you write the steps of a
procedure as English sentences, and Frazaro checks every sentence, refuses
the ones it cannot read, in words, and carries out the rest. VLA is the
Lisp beneath that English, the form every sentence translates into before
anything is emitted or built. This crate is that language alone: the reader
and the printer of its forms, the macro expander, the head table that names
the core forms, the refusal catalogue's mechanism with the language's own
refusals, the sheet model with its A1 coordinates and the references read
out of a formula's text, the view record, one window of a sheet as the
lines a viewport draws from, and recalculation's mechanism, the formulas
of a grid computed in dependency order with a budget in cells. No English,
no file format, no clock, and no dependency. Built for WebAssembly, the
module's import section is empty.

[`frazaro-core`](https://crates.io/crates/frazaro-core) stands on this crate
and adds the bridges: the English engine with a proof per sentence, the VBA
and formula emitters, the OOXML and OpenDocument writer and readers.
[`frazaro`](https://crates.io/crates/frazaro) is the command line over both.
An engine that runs a grid as a game stands on this crate and never on the
core, which is why the language is a crate of its own. The Excel add-in in
the repository is the reference implementation; this crate is held to the
same corpus of tests, so a form means the same thing here as there and is
refused in the same words.

## What it does

- **Reads and prints VLA.** `reader::read_forms` takes a text to its forms,
  each a symbol, a string, a number or a list with the line it came from;
  `reader::each_form` gives the same forms one at a time, for a text too long
  to hold as forms, refusing what `read_forms` refuses before the first; and
  `printer::write_datum` and `printer::write_pretty` write a form back, as
  the add-in writes its goldens. A string is written `"..."`, a quote and a
  backslash escaped, or raw, `#"..."#`, as it stands, more `#` where its text
  holds a quote and a `#`, a line break inside one read as a line feed; the
  printer writes every string escaped. `reader::list_heads` gives each list's
  head as the reader reads it, for a tool that must read a text as the
  reader does before anything is parsed.
- **Expands macros.** `expand::expand_text` takes a program's text and a
  prelude's, defines the prelude's macros and expands every form to a
  fixpoint, refusing a malformed definition or a runaway expansion by name.
  The prelude is a library the caller passes: the repository's
  [scripts/prelude.vla](https://raw.githubusercontent.com/Spreadsheet-Company/Frazaro/main/scripts/prelude.vla)
  is the standard one, and a program may carry its own definitions.
  Twenty-nine primitives compute while a macro expands, over the data
  written in the source: the list operations (`car`, `cdr`, `cons`,
  `quote-if`, `cond` and their kin, and `list?`, which tells a word from a
  list and refuses neither), arithmetic and comparison (`+expand`,
  `-expand`, `*expand`, `quotient-expand`, `mod-expand`, `=expand` and the
  rest), text joined or split into its characters (`&expand`,
  `characters-expand`), a cell's address written from a row and a column
  or read back (`address-expand`, `row-expand`, `column-expand`,
  `column-letters-expand`), and the entry of a list found by its head
  (`assoc-expand`). A name that ends in `expand` is computed while the
  macro expands; each refuses by name what it cannot compute. A loop,
  `each`, writes its forms once for each element of a list, and
  `each-index` with a count beside the element, from a start the author
  names, so that a macro says what its rows are and not how to walk to them;
  and a list of names where the name stands, `((address value) entries)`,
  names each part of an entry, so that no part is read by its position:

  ```lisp
  (defmacro (cells sheet & entries)
    "Puts values in cells of one sheet. Each entry is a cell's address, or a range of cells, and the value that goes there."
    (each ((address value) entries)
      (cell sheet address value)))
  ```

  `(cells "S" ("A1" 1) ("B1" 2))` expands to `(begin (cell "S" "A1" 1)
  (cell "S" "B1" 2))`: a loop standing alone is one `begin`, a loop that is
  an element of a list has its results spliced in, and a primitive reads a
  loop as the list of its results. A define, `(define stop State!B12)`,
  gives a form a word in every form of a program after its own: the
  expander puts the form where the word stands as a value, an expand-time
  primitive's argument among them, never at a list's head or inside a quote
  or a text, and refuses by name a word that already means something (a
  reference, a number, a form of the language, a macro's name), a word used
  in the form that defines it, and a ring of words.

  A macro may take named pieces after its positional parameters, so that a
  caller names only what differs, in any order, and a piece can be added
  with a default without moving a call:

  ```lisp
  (defmacro (rule range (sheet "Screen") (text "=1"))
    "a formula row, on the Screen and of the plain text unless the call says"
    (formula sheet range text))
  ```

  `(rule "B2:K11" (text "=2") (sheet "Tower"))` expands to `(formula "Tower"
  "B2:K11" "=2")`, and so does the call in the other order. `(sheet)`, the
  word alone, is a piece the call must give, and a default may name the
  parameters declared before it. An argument after the positional ones that
  is none of the pieces, a piece holding other than one form or given twice,
  a piece with no default left out, and a positional argument written as one
  of the pieces are each refused by name, with the call's line where its
  author wrote it. `Expander::get_macro` answers a macro's `pieces` as
  written.

  The word `self` names the place a verb writes, so that a rule reads as its
  author thinks and a verb can be an expression:

  ```lisp
  (defmacro (rule sheet range expression)
    "one formula row, self its range's first cell"
    (bind-expand ((self (ref sheet (first-cell-expand (ref sheet range)))))
      (formula sheet range (formula-expand expression))))

  (defmacro (running event)
    "its own last value plus an event"
    (+ (last self) event))

  (rule "Tally" "B1:E10" (running Input!B16))
  ```

  The last form expands to `(formula "Tally" "B1:E10"
  "=Tally.last!B1+Input!B16")`. `bind-expand` binds the word to one cell on a
  sheet for every form the expander meets in its body, computing the cell under
  any binding outside it, and an inner binding hides an outer. The expander
  reads the word where it stands as a value, as a define's word; a macro's
  template is still a function of its template and its arguments, and a
  primitive that reads a word's text refuses this one by name. `(offset-expand
  reference rows columns)` and `(offset-expand reference rows columns height
  width)` are Excel's `OFFSET` at expand time, so `(offset-expand self -1 -1 3
  3)` is the square around the cell: `(bind-expand ((self Screen!B2))
  (formula-expand (- (sum (last (offset-expand self -1 -1 3 3))) (last self))))`
  writes `=SUM(Screen.last!A1:C3)-Screen.last!B2`. With no binding around it the
  word is refused by name where it would be read, in a formula's expression and
  wherever one reference is read, and is left alone everywhere else, so a
  running program's variable of that name is a variable.

  A function computes a value while a program expands, as a primitive does,
  so that a helper that walks a list answers what it found:

  ```lisp
  (defexpand (keep-big xs)
    "the numbers of a list greater than 2"
    (quote-if (null? xs)
      (quote ())
      (quote-if (>expand (car xs) 2)
        (cons (car xs) (keep-big (cdr xs)))
        (keep-big (cdr xs)))))
  ```

  `(cell "S" "A1" (keep-big (quote (1 5 2 7))))` expands to `(cell "S" "A1"
  (5 7))`. A call computes wherever a primitive computes, at the top of a
  program, in a template and in a primitive's argument, its arguments values
  and its body one form read as a template is, and its value is expanded
  again as any form is. `(let-expand ((name value) ...) form ...)` names
  values for its forms. A call that is its function's value (its body, the
  branch a `quote-if` or `cond` chose, or a `let-expand`'s form) loops and
  nests nothing; any other call nests at most `FUNCTION_DEPTH_LIMIT` deep,
  14, and one value is worked out in at most `FUNCTION_FUEL` calls, 5,000,
  each refused by name past it. A refusal inside a working-out names the
  calls it stands in, `(computing (outer (a)) > (inner (a)))`. A primitive
  reads a call of a primitive or a function that a parameter, a loop's name
  or a `let-expand`'s name hands it as that call's value. A word names one
  thing, a macro or a function, and a second function of one name is
  refused. `Expander::define` takes a definition of either kind and answers
  whether a form was one.

  The expander shares what it walks, as the add-in does: a list is held
  once, its `cdr` a view of its items and a parameter's value the argument
  itself, so that a macro walking a list, or a loop over one, costs its
  length and not its square, which `tests/walks.rs` holds by counting the
  allocations and their bytes.
- **Writes formulas from expressions.** `formula_syntax::write` writes an
  expression the expander has finished with as a formula's text in Excel's
  own spelling, `(if (> A1 3) "big" "small")` as `=IF(A1>3,"big","small")`:
  no space outside a text, the fewest parentheses Excel's precedence allows,
  a sheet quoted only where it must be; and `formula_syntax::read` reads a
  formula back as that expression, for tools and tests. The expander's
  `formula-expand` is the writer at expand time, its argument expanded
  first, so a macro or a define inside it has done its work. `(last X)` is
  a reference's twin, the sheet's name with `.last` after it: after
  `(define stop State!B12)`, `(formula-expand (if (= (last stop) 1) 0 1))`
  writes `=IF(State.last!B12=1,0,1)`. `(absolute X)`, `(absolute-row X)` and
  `(absolute-column X)` write a reference with the `$` marks each adds, and
  never take one away, so they and `last` stand in either order: `(last
  (absolute stop))` writes `State.last!$B$12`. And `sheet-expand` and
  `range-expand` give a reference's sheet and its cell or range as the two
  texts `ref` takes, so a macro can write a row at a word, and
  `first-cell-expand` its first cell, a range's top-left, so a verb can twin
  the cell its fill is written for. `(funcall f a ...)` calls what `f`
  answers with its arguments, `(funcall (if c SQ PLUS1) 5)` writing
  `=IF(c,SQ,PLUS1)(5)`, and the reader reads a call in place back as one.
- **Names the core forms.** `headtable::row` and `headtable::resolve_head_alias`
  answer for every form the language has, its aliases and its arity, from
  the table the add-in's own code exports.
- **Holds the grid.** `sheet::Workbook`, `sheet::Sheet`, `sheet::Cell` and
  `sheet::Content`: sheets of cells, each a text, a number, a truth value or
  a formula's text, with column widths, a style table and the row of the
  sentence that wrote each cell. `Sheet::set_formula` fills a range with one
  formula as Excel fills it, the references moved cell by cell.
- **Draws a number under a format's code.** `numfmt::Code::parse` takes a
  code of Excel's, such as `0.00` or `▲0.00;▼0.00;=0.00`, its sections,
  digit placeholders, thousands, thousandths, percents and literals, and
  names whatever else a code holds that this version does not draw;
  `Code::text` gives the text a value draws as under it, as Excel's
  `Range.Text` does: the number's fifteen significant digits, rounded half
  away from zero to the code's decimals; and a date's and a time's parts,
  the year, the month and the day by number or English name, the hour, the
  minute, the second and its decimals, an elapsed time and a twelve-hour
  clock, over Excel's 1900 calendar. A style's `sheet::NumFmt` is `General`,
  `Text` or a `Code`.
- **Says how a cell looks.** A style (`sheet::Style`) is a fill, a number
  format and a wrap, and what it says of its text's font, its alignment and
  its borders (`appearance`): a weight, a slant, a size in points from 1 to
  409.55 by twentieths, as Office keeps one, a color, an underline, a strike
  and a family, its name and the generic that draws where a machine lacks it;
  a horizontal and a
  vertical alignment; and each of a cell's four edges, a line of Excel's
  fourteen and a color, each said or not. The view record prints what a
  format says after its wrap, `(font bold 22 "FFFFFF")`,
  `(align center middle)` and `(border (left medium))`, and the loader reads
  it back. `Style::over` is the format a cell draws with under a look: the
  look's fill, number format and wrap, what it says of a font, an alignment
  and a border, and the cell's own format for the rest. A format of groups
  alone, `(format 64 (fill "FFE699") (font "800020"))`, says no more than its
  groups (`Style::partial`, `sheet::Unsaid`), so that a look of it draws the
  cell's own fill, number format and wrap for those it leaves unsaid.
- **Holds a conditional format.** `rule::Rule` is a row of its own,
  `(conditional "Board" 1 "B2:J10" "=Lit!A1=1" 64 stop)` or `(cell greater "=5")`
  in the formula's place, as a sheet holds it (`Sheet::rules`, `Sheet::put_rule`). A
  machine computes each rule's answer as a fill on a sheet of its own that no
  reference can spell, so a rule moves no value, and prints what applies at a cell,
  `(lit "Board" "B2" 1 3)`, by priority up to the first that stops. A color scale, a
  data bar, a set of icons, a rank, an average, a duplicate and a value alone, which say
  no formula, are rows of their own words (`rule::Uncomputed`), planned for nothing and
  drawing nothing, so that a workbook keeps them.
- **Holds a merged range.** `(merge "Board" "B2:D2")` is a row of its own, printed in
  every record of its sheet after its rules (`Sheet::merges`, `Sheet::put_merge`), of more
  than one cell, no two overlapping (`grid-merge-overlap`). It moves no value and no plan:
  a covered cell holds what its rows wrote and is read as any cell. Each of its cells
  prints what its first draws, and where a sheet's merged cells by its rules pass
  1,048,576 its first alone does; drawing the range as one is a page's.
- **Reads references.** `refers::scan` finds every reference inside a
  formula's text, cells, ranges, columns, rows, names, Tables, external
  books and 3D spans, and names the calls it cannot read; `refers::r1c1`
  renders a formula relative to its cell, `refers::shift_a1_references` is
  the mover a fill uses, and `refers::quote_sheet` is the one sheet-quoting
  rule.
- **Draws the view record.** `view::window_of` finds a sheet and a window
  in a model, and `view::view_text` prints that window as lines: every
  sheet, the window and the sheet's extent, its columns and formats, then
  each cell's value or formula, the text a number draws under its format's
  code, its format and the row of the sentence that wrote it. It is the
  first implementation of `projection::Projection`, the
  seam a second projection enters by. `view::view_text_valued` prints the
  same record with each formula's computed value after it. The second,
  `view::ValueGrid`, is the same record less its formula rows, the values a
  host draws every frame without moving every formula's text:
  `view::view_text_values` prints it. The fifth, `view::FillGrid`, is what a
  save prints, what the record says of each cell and no row a load does not
  read: a fill's formula row once over each of its rectangles, a sheet's
  formats and its value cells over theirs, a formula's values over ranges
  (`(value "S" "A1:D1" 1 2 3 4)`, `(value "S" "A1:Z40" 0)`) and its own `sheet`
  row alone: `view::view_text_fills` prints it.
- **Computes the grid.** `calc::Calc::new` reads every formula off a
  `Workbook`, builds the dependency graph from the references and finds the
  order, naming every cycle; `Calc::step` evaluates the next cells of that
  order up to a budget in cells and says how far it got, so an engine can
  run a grid at a frame rate and yield; `Calc::run` evaluates everything.
  The parser reads Excel's operators and calls over the reference scanner,
  with Excel's precedence and coercions and its eight error values,
  `#CALC!` among them, as a value kind, array constants (`{1,2;3,4}`) and
  `:` between two expressions that answer references among its forms; a
  comparison holds two numbers equal to fifteen significant digits equal,
  as Excel's does (`=0.1+0.2=0.3` is TRUE), where a lookup matches a number
  exactly, and orders two texts as Excel's save ranks them, case aside, an
  accent after the plain letter, `ß` and `ss` one, a comparison by size
  that a character it did not rank decides labelled. A range where one
  value is wanted is met at the formula's cell, Excel's implicit
  intersection, in a legacy formula, and is read whole, as an array, in a
  dynamic-array formula, which `Env::dynamic` tells, in a legacy array
  formula, which `Env::array_formula` tells, and inside a parameter its
  function reads as an array in any formula (`calc::Param`). An operator,
  or a function of one value, given an array computes at each of its
  values, Excel's lift (`calc::Array`, its values shared by a count); a
  cell holds the first value of an answer of many in a legacy formula and
  a legacy array formula, and labels it a spill in a dynamic one until
  spills come. Functions enter through `calc::Library`, the seam: the
  language registers the fourteen a game's rule needs on day one (`IF`,
  `AND`, `OR`, `NOT`, `SUM`, `MIN`, `MAX`, `ABS`, `INT`, `MOD`, `ROW`,
  `COLUMN`, `CHOOSE`, `SIN`), the lookups (`INDEX`, `MATCH`, `XMATCH`,
  `VLOOKUP`, `HLOOKUP`, `LOOKUP`, `XLOOKUP`), the tests a lookup's miss is
  read by (`IFERROR`, `IFNA`, `ISERROR`, `ISNA`), the criteria (`COUNTIF`,
  `SUMIF`), `SINGLE`, Excel's `@`, the text functions (`calc::texts`:
  `LEFT`, `MID`, `LEN`, `FIND`, `SUBSTITUTE`, `TEXT` and the rest, a text
  counted by its characters as Excel counts it), the date functions
  (`calc::dates`: `DATE`, `YEAR`, `MONTH`, `WEEKDAY`, `NETWORKDAYS`,
  `YEARFRAC` and the rest, over Excel's 1900 calendar, a text a locale
  reads labelled in arithmetic), the aggregates by criteria: `SUMIFS`,
  `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS` and `MINIFS`, `SUBTOTAL`
  by each of its numbers, a hidden row (`Sheet::rows`) passed by at 101 to
  111, the database functions (`DSUM`, `DGET` and the rest) and the
  aggregates `SUBTOTAL` names (`AVERAGE`, `COUNT`, `COUNTA`, `PRODUCT`,
  `STDEV`, `VAR` and their newer names), and the array functions
  (`calc::arrays`: `SEQUENCE`, `FILTER`, `SORT`, `UNIQUE`, `SUMPRODUCT`,
  `MMULT`, the shapers `TAKE`, `HSTACK`, `TOCOL` and the rest, `#CALC!`
  for an array with nothing in it), and a cell's formula read by its
  reference (`calc::formulas`: `ISFORMULA`, and `FORMULATEXT`, the text
  the view record prints, labelled where Excel's formula bar may spell it
  otherwise), the information functions (`calc::information`: `ISNUMBER`,
  `ISBLANK`, `ISREF`, `ERROR.TYPE`, `TYPE`, `NA` and the rest, what a value or a
  reference is), `ROUND` and `IFS`, and a cell's address as a text from its row and its
  column (`calc::texts`: `ADDRESS`, an address in R1C1 labelled), a boolean as a function
  (`calc::information`: `TRUE()` and `FALSE()`), and a mask read by one call
  (`calc::library`: `BITAND`, `BITOR` and `BITXOR`, over whole numbers below
  2^48). A crate above registers more, as
  `frazaro-core` registers Excel's library by measurement. A lookup or a
  criterion asked of one range again and again searches the range's
  index (`calc::Indexes`), built at its second probe in one evaluation and
  kept until a cell of the range changes, its answers the ones a walk of the
  range gives (`calc::index`); the `SUMIFS` family's ranges are grouped by
  the keys of its equalities, every group answering the formulas that ask
  it. A formula that calls anything else is not computed, and says so naming
  the function; a cell that reads it inherits the label.
  `Calc::first_refusal` is what an engine's load asks: a cycle, a function
  outside the library, or a construct the parser does not read, a defined
  name's text among them, each refused by name from the catalogue.
- **Computes functions written in formulas.** `LET` binds a name to an
  expression within a formula, the expression substituted where the name
  is read, so that a formula written with `LET` computes what the same
  formula written out computes; `LAMBDA` makes a function of named
  parameters, called where it is written (`=LAMBDA(x,x+1)(5)`), bound by
  `LET`, handed to `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL` or
  `MAKEARRAY`, or held by a defined name and called by its name from any
  formula, `ISOMITTED` telling an argument left out. A named function costs
  what the formula it stands for costs: a call of one that does not call
  itself, in a fill of sixteen cells or more, is written into the formula
  at the load, and a call at a place of its own is kept, its function's
  body held once as a program each call runs and the call followed through
  the cells its function reads, so that it computes when what it reads
  changes. A function that calls itself runs as calls, each counted, to a
  depth of 4,096 and a million calls a cell, `#NUM!` past either, a call
  too deep for the stack computed on a stack of the language's own, so that
  the depth is the same on every host; a formula whose value is a function
  answers `#CALC!`, a function where a number is wanted `#VALUE!` and a
  reference called `#REF!`, and the information functions read a function
  as a kind of its own (`TYPE` 128).
- **Runs a grid as a machine.** `machine::Machine` is a grid kept between
  calls, for a game engine that steps it at a frame rate: `Machine::load`
  reads the rows a view or a reflect prints back into a grid, the record's
  inverse, one form at a time, so that a load peaks at about what it holds,
  and a `function` row, `(function "SUMMA" "VIII SUMMA" "A3:A5" "J3")`, makes
  a room of cells a function: its inputs read row by row as its parameters,
  its output a cell or a range answered as an array, its formula cells
  compiled to the `LAMBDA` a person would write from them and every call
  computed as that `LAMBDA` is, a body that does not compile `#CALC!` at its
  calls and the census saying why;
  `write` takes `cell`, `formula` and `derived` rows, all or none, a
  `formula` row over one cell reading as Excel 365 reads the text typed where it
  ends in the word `dynamic`, `name` rows, a defined name's text, every
  formula reading the name planned again, through other names too, and
  computed at the next frame, as a load of the written grid would, and
  `function` rows, a function defined or its row replaced, compiled again at
  every write that changes a formula of its room;
  `step` computes every formula in dependency order with a budget in cells,
  yields and resumes, and at the end of each frame copies every sheet's
  values into its twin, `X.last`, a hidden sheet the next frame's formulas
  read the previous frame from (`=Screen.last!B2+1`); `view` answers the
  record with each formula's value, the values (the record less its formula
  rows, what a host draws), the changes (the values' rows of the cells the
  last frame changed, so that a host repaints those alone, or `whole` where
  the machine cannot say), the fills (what a save prints and the load reads
  as the grid), the plane, a byte a cell, or the names, each defined name's
  `name` row and each function's `function` row, which a save prints beside
  the fills. After each frame
  `changed` says how many formula cells it changed the value of, to the
  bit, `changed_on` how many of them on a sheet and `reached` how many it
  computed, so that a host stepping a grid until a cascade settles stops
  when nothing moves: a frame that changed nothing, with no write before
  it, is followed by frames that change nothing. Each
  formula is parsed once per shape (a filled formula, or cells whose R1C1
  text is the same) and evaluated at each cell's offset, and the order is
  found over each fill's rectangles and each formula alone and kept as
  blocks, a shape's cells on a sheet at a level, and as a list of its cells
  too where it holds at most 262,144, so that a large grid's plan holds no
  entry for a fill's cell but where the fill reads itself or a cycle may
  stand. Life on a plane of 320 by 200, 62,964 formulas, steps a frame,
  its clock advanced, in about 2.3 ms natively on the repository's bench
  machine (its floors, `tools/check_recalc_floors.ps1`). The grid is the
  machine's own: a filled formula held once, a rectangle and its text, cut
  into pieces by a write into part of it, and every value once, by place,
  nine bytes a cell, a twin keeping its own only where a formula reads it;
  `Machine::workbook` lends it as a workbook, and `load::Draft` holds the
  rows as records between `read_rows`, or `read_rows_each` over `each_form`,
  and `Machine::new`. `census` prints
  the plan as rows before a frame runs: each shape's cells, the levels of
  the order it spans, the runs it is computed in and the path each of its
  reads takes, so that a grid's cost is read ahead (`calc::Shapes::census`
  prints a plan a door made of a workbook the same way). `machine::Handles`
  keeps grids under handles from 1, never reused, sixteen at once. The C
  surface over the four calls, `vla_load`, `vla_write`, `vla_step`,
  `vla_view` and their companions, is compiled always and exported only
  under the `c-abi` feature, so that no module built on this crate exports
  a second door to its grid.
- **Refuses in words.** Every refusal is a `messages::Refusal`, a value
  carrying the catalogue's id, number, source and text, never a panic.

## Using it from Rust

```text
cargo add vla-lang
```

```rust
use vla_lang::{expand, printer, reader};

fn main() {
    let prelude = std::fs::read_to_string("prelude.vla").unwrap();
    let program = "(when (> 1 0) (debug-print \"hello\"))\n";

    // The program with its macros expanded, and how many applications it took.
    match expand::expand_text(program, &prelude, true) {
        Ok((expanded, fired)) => println!("{expanded}; {fired} applications"),
        Err(r) => eprintln!("{} ({})", r, r.id),
    }

    // The same text read as forms and printed again, one datum a line.
    for form in reader::read_forms(program).unwrap() {
        println!("{}", printer::write_datum(&form));
    }
}
```

## The contract

Bytes in, bytes out, nothing else. The crate has no dependencies, and
nothing in it reads a file, a clock, a socket or a random number at run
time. The two tables it needs, the head table and its half of the message
catalogue, are embedded at build time from `data/`, inside the crate,
exported from the add-in's own code and held to it by the repository's
checks. `vla_lang::VERSION` is the crate's version, which is the release's,
the add-in's and `frazaro-core`'s, one number for the language.

Every text the crate reads is answered or refused by name, never with a
panic, a stack overflow or a hang, so an engine can hand it a cartridge
from anyone. The limits are public constants, each refused by its own id:
- lists nested deeper than 64, in what the reader reads and in what the
  expander writes (`reader::FORM_DEPTH_LIMIT`);
- a formula nested deeper than Excel's 64 levels, or longer than its
  8,192 characters (`calc::formula::NESTING_LIMIT`, `LENGTH_LIMIT`);
- more than a million forms in what one macro application writes
  (`expand::EXPANSION_SIZE_LIMIT`), since the crate copies a macro's
  arguments into its template.

Real programs nest 10 deep at most and real formulas 6. Each stage is
tested at its limit on a 1 MiB stack, the tightest an engine is given.

```text
cargo build --release -p vla-lang --features c-abi --target wasm32-unknown-unknown
```

builds the crate as a WebAssembly module with the machine's C surface,
whose import section the repository's `tools/check_core_imports.ps1` holds
at zero entries, so that no engine built on it can reach a socket, a clock
or a file through the language, and whose exports
`tools/check_wasm_exports.ps1` holds to exactly its nine functions and its
memory.

## What comes next

Each is an item of the repository's roadmap (`docs/BETA_ROADMAP.md`), and
this page gains its section when the item closes:

- **Spills, and Excel 365's arrays whole.** A dynamic formula's answer of
  many values laid into the cells beside it, a legacy array formula over
  many cells computed at each of them, and a dynamic formula's shape
  lowered to its program, so that a workbook of Excel 365's arrays computes
  as Excel computes it, and as fast as its legacy twin (KERNEL.55).
- **A write that empties a cell.** A row the machine's write takes beside
  `cell`, `formula` and `derived`, so that a host gives a cell back to no
  value, as a rebuilt cartridge that no longer writes it asks (KERNEL.46).
- **Two of Excel's readings.** A formula's literal of sixteen digits or
  more, which Excel cuts to fifteen, read as Excel reads it or named where
  Excel would cut it (KERNEL.51); and a database function's computed
  criteria, a formula computed at each record, where it is labelled today
  (KERNEL.52).

## Where it comes from

The add-in in the repository's `src/` is the reference implementation of
the language, and the goldens it wrote under `scripts/` are what this crate
is held to, by the treaty in
[conformance/README.md](https://github.com/Spreadsheet-Company/Frazaro/blob/main/conformance/README.md):
the corpus program read, expanded and compiled byte for byte through the
core, and the reference goldens the add-in writes reproduced by this
crate's own tests. The crate follows the goldens and never leads them. It
was cut out of `frazaro-core` on 2026-10-08 (PORT.12) so that an engine
can stand on the language without the bridges.

The goldens and fixtures the tests read travel with the package, under
`testdata/`, so `cargo test` runs from the published crate. The repository
holds the copies to their sources byte for byte. The expander is held to
the add-in's expansion of 656 cases, every expand-time primitive, the loop
and their refusals among them (`tests/expand_golden.rs`). Beside the goldens, the
tests hold the language's laws over generated programs (`tests/laws.rs`)
and knock at every door with mutated text (`tests/strangers.rs`, and a
longer run with `--ignored`).

## Licence

Apache-2.0. The repository's `REUSE.toml` maps every file to its licence.
