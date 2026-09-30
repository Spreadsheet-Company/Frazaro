# Level 0b101 · Reference: An LLM's Guide to Frazaro

*(A bridge betwixt minds & machines)*

| | |
|---|---|
| **For** | A language model asked to draft Frazaro sentences, and the person who pastes this page into one. Also anyone who wants everything on one page. |
| **You need** | Nothing else. This page is written to stand alone. |
| **You will be able to** | Draft sentences that pass validation, and know what to do when they do not. |
| **Plan on** | It is a reference. Consult it; do not read it through. A model reads it through in a moment, which is the other reason it exists. |
| **Comes after** | [Level 4](4-wizard.md), or nothing at all. |
| **Written against** | Frazaro `0.7.0`. Sentence shapes that first work in `0.7.1` are marked `(0.7.1)`, and those that first work in `0.8.0` are marked `(0.8.0)`. |

---

## R0. What this page is

Frazaro contains no language model and makes no network call. It reads a
small, published English and refuses everything else. A language model
is good at the opposite thing: it reads anything, and guesses.

The two fit together in one arrangement only, and this page describes it.

```text
   a person describes the work, in any words
                  |
        a model drafts sentences ............ guided by this page
                  |
   Frazaro validates every sentence ......... deterministic; it refuses or accepts
                  |
        the person reads the sentences, and runs them
```

The model proposes. Frazaro decides. The person signs. Nothing the model
writes is trusted because the model wrote it; it is trusted when
**Validate Instructions** turns the row green and a person who
understands the work has read it.

The bridge is carried by hand. Frazaro does not call a model, and a model
cannot call Frazaro. A person copies text between them. A built-in bridge
is on the project's roadmap and is not built.

**To the person.** Paste this whole page into the model before your
request. Say which sheets, columns and tables exist. Paste back, word for
word, any message Frazaro shows beside a red row.

**To the model.** Sections R1 to R14 are the language. R15 is how to
answer. R1 is the part to obey when the rest is unclear.

---

## R1. The contract

1. **Write only sentences whose shape is on this page.** Every sentence
   you write must match a structural form in R4, or a shape in R7 or R8,
   with values filling its holes. Do not invent a sentence because it
   would be reasonable English. Reasonable English is refused.
2. **Frazaro is the judge.** You cannot validate a sentence. Say so. Ask
   for the verdict.
3. **When a sentence is refused, repair it from the refusal.** The
   message says how far the reading got and what was expected next
   (R10). Change that word. Do not rewrite the whole program.
4. **When no sentence exists for a step, say so.** Write the step as a
   `#` note, which Frazaro keeps and does not run, and tell the person it
   is for a human to do, or for a new sentence to be requested (R13).
   Never approximate a missing sentence with a nearby one that does
   something else.
5. **Where two sentences are neighbours, choose on purpose.** A border
   *around* a range is not borders on *every cell in* it. If the request
   does not say which, ask.
6. **State every assumption about the workbook**: sheet names, which
   column holds what, where the data ends, the names of Excel Tables.
7. **Do not describe Frazaro as AI.** It is not. Do not say a program
   "should work". Say it has not been validated yet.
8. **Prefer the plain program.** Name intermediate values. One sentence
   per row. Hyphenated names.

---

## R2. Words, numbers and marks

### The alphabet

A bare word is made of the letters A to Z, digits, hyphens and
underscores. Capitals are ignored in bare words. Anything else belongs
inside double quotes.

### Text

| Write | Notes |
|---|---|
| `"Missing receipt"` | Double quotes. Straight, not curly. |
| `"She said ""yes"""` | A quotation mark inside text is written twice. |
| `"All done."` | A period inside quotes is text. The sentence still needs its own: `Show "All done.".` |
| `"don't"` | An apostrophe is fine inside quotes, and refused outside them. |

Quoted text keeps its capitals and spaces exactly.

### Numbers

| Write | Means |
|---|---|
| `5` · `3.14` · `-1` | As written. |
| `1,000,000` | One million. Commas must group digits in threes exactly. `1,00` is refused. |
| `15%` · `15 percent` | 0.15 |
| `zero` … `twenty` | 0 … 20. Above twenty, use digits. |
| `first` … `twentieth` | 1 … 20, where a value is expected: `Delete the third row.` |

### References

| Kind | Write | Not |
|---|---|---|
| Cell | `cell B2` | `$B$2` |
| Cell on another sheet | `cell Data!B2` · `cell 'Q1 Data'!B2` | |
| Range | `range A1:C50` · `range A:C` · `range 1:5` · `range Data!A1:B10` | |
| Named range | `range "Q1 Totals"` | a bare name with a space |
| Column | `column C` · `column AA` | `column 3`, except in the few shapes of R7 whose column hole is an `expr` |
| Computed cell | `cell in column D row r` | |
| Computed cell, column by number | `cell in column number 3 row 2` | |
| Sheet | `sheet Data` · `sheet "Q1 Data"` | `sheet Q1 Data` |
| File path | `"C:\Reports\close.xlsx"` | any unquoted path |

A bare reference is folded to small letters. That is harmless for cells,
sheets, tables and pivots, which Excel matches without regard to
capitals. It is not harmless for text that is compared with data. **Quote
any text that must match a cell's contents.**

### Names

One word. It begins with a letter. Hyphens are welcome and recommended:
`last-row`, `order-total`, `receipted-total`.

These words cannot be names, since Excel's programming language reserves
them:

```text
and as boolean byref byval byte call case close const currency date declare
dim do double each else elseif empty end enum eqv erase error event exit false
for friend function get goto if imp implements in input integer is let like
lock long loop lset me mod new next not nothing null object on open option
optional or paramarray preserve print private property public put raiseevent
redim rem resume rset seek select set single static stop string sub then time
to true type typeof unlock until variant wend while with withevents write xor
```

Nor can a value word (`today`, `now`). A step may not be named with a
word that already means something when `of` follows it (`length`,
`count`, `sum`, `first`, `last`), and no name is clearer for being one.

**A hyphenated name always works.** `date` is refused; `invoice-date` is
fine. When in doubt, hyphenate.

Do not begin any name with `vla-`. That prefix is reserved.

**A name is never shaped like a cell** (from `0.7.1`). A word of one to
three letters and then one to seven digits, such as `a1`, `q1` or `fy24`, is
refused wherever a name is made or a value is read: `Set A1 to 5.` and
`Put A1 plus B1 into cell C1.` are both refused before anything runs. To
mean the cell, write `cell A1`; for a name, hyphenate it (`q1-total`). Four
letters and digits (`abcd1`) make an ordinary name. A sheet, table or pivot
may still be called `Q1`.

### Filler

| Word | Rule |
|---|---|
| `the`, `please` | May appear anywhere. They are dropped. |
| `a`, `an` | May appear directly before a fixed word of a shape: `Make row 1 a header row.` `Add a border around range A1:C3.` **Never directly before a value**, where they would be read as a name. |

### Marks

| Mark | Use |
|---|---|
| `.` | Ends every sentence. |
| `:` | At the end of a row, opens a block. |
| `,` | Separates the items of a list, and the two halves of a one-line form. |
| `#` | Begins a note. The rest of the row is for people. |
| `+ * / = < > & $ ; ( ) '` | **Refused** outside quotes. Operators are words. |

---

## R3. Layout

1. **One sentence per row.** It ends with a period.
2. **A row beginning with `#` is a note.** Frazaro keeps it and does not
   run it. A note never ends a block.
3. **A row ending with a colon opens a block.** The rows under it belong
   to it.
4. **Indentation means nothing to Frazaro.** Indent two spaces for the
   reader.
5. **A blank line closes every open block.**
6. **`Done.` closes only the innermost block.**
7. **`Otherwise`, `When it is` and `If that fails` continue a chain across
   the blank line** that closed the branch before.

Rules 5 and 6 are where drafts go wrong. Therefore:

> **Never put a blank line inside a block.** Use a `#` note for a visual
> break.
>
> **Always put a blank line after a block**, before the next sentence
> that is meant to stand outside it.
>
> **Use `Done.` only** when an inner block ends and the outer block has
> more sentences to come.

```text
Count r from 2 to last-row:
  Set on-hand to cell in column D row r.
  If on-hand is at most 10:
    Put "REORDER" into column H row r.
    Make cell in column H row r bold.
  Done.
  Put "checked" into column I row r.

Show "Finished".
```

### Where a program works

A program that never says otherwise works on a sheet named *Output*.
Begin with the sheet you mean:

```text
Work on sheet "Expenses".
```

`Work on sheet` makes the sheet if it does not exist. `Go to sheet` fails
if it does not exist. A program may not work on the *Frazaro* sheet
itself, nor on any sheet Frazaro keeps for its own use.

### Inside a document

In a Word document or a text file, only what lies between two tags is
read. Each tag stands alone on its row.

```text
<Frazaro week="Week 39" (automated by Accounts Payable) 0.7.0>
Put "Reimbursement total, {week}" into cell H10.
</Frazaro>
```

| In the opening tag | Is |
|---|---|
| `name="value"` or `name=value` | A value the section may use as `{name}`. Quotes are needed when the value has a space. |
| `(words in parentheses)` | A note for people. |
| `0.7.0` | The section needs that release or a later one. |
| `english`, `espanol` | The phrasebook the section is written for. |

`{name}` is filled in when the document loads. To write a brace itself,
double it. A document with no tag is read whole, and every row that is
not a note must be a sentence.

---

## R4. Structural sentences

These are built into the engine. `<value>` is R5. `<condition>` is R6.
`<sentence>` is any single sentence.

### Deciding

```text
If <condition>, <sentence>.

If <condition>:
  ...

Otherwise, if <condition>:
  ...

Otherwise:
  ...
```

```text
When <value> is <case>:
  ...

When it is <case> or <case>:
  ...

Otherwise:
  ...
```

### Repeating

```text
Repeat <value> times:                      counter runs from 1
Repeat until <condition>:
While <condition>:
Count <name> from <value> to <value>:
Count <name> from <value> to <value> step <value>:
Count <name> down from <value> to <value>:
Count <name> down from <value> to <value> step <value>:
For each <name> in <list or remembered range>:
For each <name> in keys of <lookup>:
For each pair in <lookup>:                 key of pair, value of pair
```

Each has a one-line form with a comma in place of the colon:

```text
Repeat 2 times, show counter.
Count k from 2 to 9, show k.
For each r in results, show r.
While countdown is greater than 0, decrease countdown by 1.
```

```text
Stop the loop.          leaves the innermost loop
Stop.                   ends the program
Done.                   closes the innermost block
```

Inside `Repeat … times`, the word `counter` is the count. For a loop
inside a loop, use `Count` and name the number.

### Naming

```text
Set <name> to <value>.
Create a number called <name>.
Create a text called <name>.
Create a value called <name>.
Create a list called <name>.
Create a lookup called <name>.
Define <name> as <fixed value>.
```

`Define` stands at the top level, outside any block. Its value is quoted
text, a number, or an earlier defined name, and can never be changed.

### Changing

```text
Increase <name> by <value>.
Decrease <name> by <value>.
Add <value> to <name>.
Grow <name> by <value>.
Shrink <name> by <value>.
Append <value> to <list>.
```

`Increase total by 10%.` grows the total by a tenth. A percentage mixed
into a longer amount is refused: `Increase total by 5 plus 5%.`

### Steps

```text
To <name>:
  ...

To <name>, with <detail> of <usual value> and <detail> of <usual value>:
  ...

To <name> of <detail>:
  Give back <value>.

To get <name> using <detail> of <usual value> and <detail> of <usual value>:
  Give back <value>.

To get <name>:
  Give back <value>.
```

Calls:

```text
Tidy-up.
Stamp with row-number of 2 and value of "beta".
Set landed-total to landed-cost of order-total.
Set full-commission to commission using sale of 2000 and rate of 10%.
Set vat to vat-rate.
```

Define a step at the top level, before the sentences that use it. A
detail may be left out of a call, and takes its usual value. `Give back`
belongs only inside a step that returns a value.

### Guarding

```text
Try:
  ...

If that fails:
  ...
```

The first sentence under `Try` that fails ends the attempt. Inside
`If that fails:`, `the problem` is what went wrong, in words.
`If that fails:` may be left out.

### Talking

```text
Show <value>.                 a message box; the run waits
Say <value>.                  the same
Log <value>.                  a line in the trace; the run does not wait
```

### Buttons and events

```text
When "<caption>" is clicked:
  ...

Make a button "<caption>" at cell D3.
```

```text
When the sheet changes:
  ...
```

Both handlers stand at the top level. `When the sheet changes:` may
appear once in a program and runs only under **Interpret and Run**. A
button is never made with `Create`.

### Libraries

```text
Use library "tools.vla".
```

### The middle language

A row that begins with `(` is taken as a form of VLA and not as a
sentence. Do not write one unless the person asks for VLA.
[Level 3](3-master.md) is its manual.

---

## R5. Values

A value is written in words. In order, from the loosest binding to the
tightest:

| Rank | Words | Means |
|---|---|---|
| 1 | `joined with` · `followed by` | text put end to end |
| 2 | `plus` · `minus` | |
| 3 | `times` · `multiplied by` · `divided by` | |
| 4 | `%` · `percent`, after a value | divided by 100 |

Within a rank, left to right. `2 plus 3 times 4` is 14.

**There are no parentheses.** Name the inner part first:

```text
Set subtotal to price plus shipping.
Set total to subtotal times quantity.
```

### What can stand as a value

| Write | Gives |
|---|---|
| `5` · `"text"` · a name | |
| `today` · `now` | the date; the date and time |
| `cell B2` · `value in cell B2` | what the cell holds |
| `cell in column D row r` | the same, with the row computed |
| `cell in column number 3 row 2` | the same, with the column as a number |
| `length of <value>` | the number of characters |
| `uppercase of <value>` · `lowercase of <value>` | |
| `absolute of <value>` | the number without its sign |
| `year of` · `month of` · `day of` · `hour of` · `minute of` | a part of a date or time |
| `sum of` · `average of` · `largest of` · `smallest of` · `count of` | over a remembered range, or a list |
| `first of <list>` · `last of <list>` · `item 3 of <list>` | one element |
| `keys of <lookup>` | its keys, to walk |
| `<lookup> for <key>` | the value stored under the key |
| `key of pair` · `value of pair` | inside `For each pair in …` |
| `<step> of <value>` | a step that gives back a value |
| `<step> using <detail> of <value> and …` | the same, with named details |

### Three things to know

**A bare range is not a value.** Remember it first:

```text
Remember range G2:G41 as revenues.
Set total-revenue to sum of revenues.
```

or use the sentence that takes a range: `Set grand to sum of range B2:B3.`

**`row` takes everything after it.** `cell in column C row r plus 1` is
the cell in row r+1. To add to the cell's value, name it first.

**There is no "minus" before a lone number.** Write `-1`, or `0 minus 1`.

---

## R6. Conditions

```text
<value> <test> <value>
<value> is empty
<value> is not empty
```

| Meaning | Tests |
|---|---|
| equal | `is` · `equals` · `is equal to` |
| not equal | `is not` · `does not equal` · `is not equal to` |
| more | `is greater than` · `is more than` |
| less | `is less than` |
| at least | `is at least` · `is greater than or equal to` |
| at most | `is at most` · `is less than or equal to` |
| text | `contains` · `does not contain` · `starts with` · `ends with` |
| multiples | `is divisible by` |

The text tests ignore capitals. `is empty` counts a cell holding only
spaces as empty.

`(0.7.1)` A condition may also ask a whole range:
`If range A1:D50 contains "x", …` and
`If column C does not contain "x", …`

### Joining

`and` and `or` are read **strictly left to right**, with no precedence
and no parentheses. `a or b and c` is *(a or b) and c*.

> Do not mix `and` with `or` in one condition. Put one test inside
> another, or compute the first part under a name.

---

## R7. Every sentence shape in the phrasebook

### How to read a shape

| Notation | Means | You write |
|---|---|---|
| `word` | That word. | the word |
| `a\|b` | Either word. | one of them |
| `{d:a\|b}` | Either word, and the choice matters. | one of them |
| `[word]` | The word may be left out. | it, or nothing |
| `stem/suffix` | Either spelling: `center/ed` is *center* or *centered*. | either |
| `{x:category}` | A hole. | a value of that category |
| `{n:expr=1}` | A hole that may be left out, and is then 1. | a value, or nothing |
| `(0.7.1)` | First works in release 0.7.1. | |
| `(0.8.0)` | First works in release 0.8.0. | |

Every sentence begins with a capital, by custom, and ends with a period,
by rule.

### What fills a hole

| Category | Write | Examples |
|---|---|---|
| `cell` | a cell reference, or a named cell in quotes | `B2` · `Data!B2` · `"Total"` |
| `range` | a range, a single cell, whole columns or rows, or a named range in quotes | `A1:C50` · `B2` · `A:C` · `"Q1 Totals"` |
| `column` | one to three letters | `C` · `AA` |
| `sheet` | one bare word, or a name in quotes | `Data` · `"Q1 Data"` |
| `color` | one of eight words, or a code in quotes | `red` `yellow` `black` `blue` `cyan` `green` `magenta` `white` · `"#FF69B4"` |
| `text` | one bare word or number, or anything in quotes | `Sales` · `"Net Revenue"` |
| `expr` | a value, as in R5 | `5` · `"text"` · `total plus 1` |
| `cond` | a condition, as in R6 | |
| `var` · `name` | one word | `last-row` |
| `text-list` and the other `-list` categories | items separated by commas | `Region, Product` · `Staff, Shifts, and Leave` |
| `clause` · `question` | R8 | |

**Lists.** Separate every item with a comma. `and` may follow the last
comma and nowhere else. `Staff and Shifts` is not a list of two.
`Staff, Shifts` and `Staff, and Shifts` both are.

**A hole of category `expr` that wants a colour** takes quoted text or a
defined name, not a bare colour word. `Make cell C4 red.` has a `color`
choice and takes the word. `Set fill-color of cell C4 to "#FF0000".` has
an `expr` hole and takes a quoted code.

**A hole of category `expr` that wants a file** takes a quoted path, or a
name that holds one.

### A. Cells and values

```text
clear cell in column {c:column} row {n:expr}
clear cell|range {r:range}
clear everything from {r:range}
convert range {r:range} to values
copy cell {a:cell} to cell {b:cell}
copy formulas of {a:range} to {b:cell}
copy range {a:range} to {b:cell} transposed
copy range {a:range} to range {b:range}
fill {d:down|right} range {r:range}
fill {r:range} with a growth series starting at {n:expr}
fill {r:range} with a growth series starting at {n:expr} with step {s:expr}
fill {r:range} with a series starting at {n:expr}
fill {r:range} with a series starting at {n:expr} with step {s:expr}
move range {a:range} to {b:cell}
name range {r:range} as {n:text}
paste values of range {a:range} into range {b:range}
put {e:expr} into|in cell {r:cell}
put {e:expr} into|in column {c:column} row {n:expr}
put {e:expr} into|in column number {c:expr} row {n:expr}
put {e:expr} into|in range {r:range}
put formula {f:text} into|in cell {r:cell}
put formula {f:text} into|in range {r:range}   (0.7.1)
put formula {f:text} into|in rows {a:expr} to|through {b:expr} of column {c:column}   (0.7.1)
put today into|in cell {r:cell}
remember range {r:range} as {v:var}
remember rows {a:expr} to {b:expr} of column {c:column} as {v:var}
stamp {e:expr} into|in cell {r:cell}
store {e:expr} at|under [key] {k:expr} in {d:var}
```

`clear cell|range` clears contents and keeps formatting.
`clear everything from` clears both. A formula is written in quotes, as
Excel writes it: `Put formula "=B2*2" into cell B3.` Into a range or into
rows, it is written for the first cell, and the cells below get it adjusted
as Fill Down would adjust it: `Put formula "=B2*C2" into range D2:D50.`
leaves D3 holding `=B3*C3`. When the last row is above the first, as with
no data rows, `into rows …` writes nothing.

### B. Reading and computing

```text
put {d:average|largest|smallest} of range {r:range} into|in cell {c:cell}   (0.7.1)
put median of range {r:range} into|in cell {c:cell}   (0.8.0)
put standard deviation of range {r:range} as {k:sample|population} into|in cell {c:cell}   (0.8.0)
put sum of range {r:range} into|in cell {c:cell}
set {v:var} to {d:largest|smallest} of range {r:range}   (0.7.1)
set {v:var} to {e:expr} rounded to {n:expr} decimals
set {v:var} to average of range {r:range}
set {v:var} to average of range {r:range} where range {c:range} matches {e:expr}   (0.8.0)
set {v:var} to cell {r:cell} of sheet {s:sheet}
set {v:var} to count of {k:empty|filled} cells in range {r:range}   (0.7.1)
set {v:var} to count of range {c:range} matching {e:expr}
set {v:var} to last filled row of column {c:column}
set {v:var} to lookup of {e:expr} in range {r:range} column {k:expr}
set {v:var} to median of range {r:range}   (0.8.0)
set {v:var} to row of {e:expr} in column {c:column}
set {v:var} to standard deviation of range {r:range} as {k:sample|population}   (0.8.0)
set {v:var} to sum of range {r:range}
set {v:var} to sum of range {r:range} where range {c:range} matches {e:expr}
```

`row of … in column …` gives 0 when nothing is found. `lookup of …`
stops the run when nothing is found; guard it with `Try:`.
`put average|largest|smallest of range … into cell …`, like `put sum`,
writes the number, not a live formula. `count of empty cells` counts the
cells that show nothing, and `count of filled cells` the cells that hold
anything; a formula that shows nothing (`=""`) counts as both, as it does
in Excel's own COUNTBLANK and COUNTA.

`median of …` is the middle value, or the average of the middle two.
`standard deviation of … as a sample` is Excel's STDEV.S, for rows that
are some of the cases; `as the population` is STDEV.P, for rows that are
all of them. A sentence that names neither is refused. `average of …
where …` leaves out a matching row whose value is empty, as Excel's
AVERAGEIF does, so it is not the sum divided by the count. In all three,
empty cells and text are skipped and a zero counts. A range with too few
numbers stops the run at the sentence, as a lookup that finds nothing
does: a median of no numbers, a sample of fewer than two, no matching
row.

### C. Text held in a name, found and counted

```text
set {v:var} to {t:expr} padded on {d:left|right} with {c:expr} to {n:expr} characters
set {v:var} to {t:expr} with {a:expr} replaced by {b:expr}
set {v:var} to {t:expr} with each word capitalized after any {k:space|non-letter}
set {v:var} to {t:expr} with extra spaces removed
set {v:var} to {t:expr} with non-printing characters removed
set {v:var} to column {c:column} as one list
set {v:var} to column {c:column} as one list separated by {s:expr}
set {v:var} to column of first cell in range {r:range} containing {t:expr}   (0.7.1)
set {v:var} to first {n:expr} characters of {t:expr}
set {v:var} to first {n:expr} letters of {t:expr}
set {v:var} to how many cells in column {c:column} contain {t:expr}   (0.7.1)
set {v:var} to how many cells in range {r:range} contain {t:expr}   (0.7.1)
set {v:var} to last {n:expr} characters of {t:expr}
set {v:var} to last {n:expr} letters of {t:expr}
set {v:var} to position of {a:expr} in {t:expr}
set {v:var} to range {r:range} as one list
set {v:var} to range {r:range} as one list separated by {s:expr}
set {v:var} to row of first cell in column {c:column} containing {t:expr}   (0.7.1)
set {v:var} to row of first cell in range {r:range} containing {t:expr}   (0.7.1)
set {v:var} to text {d:before|after} {a:expr} in {t:expr}
set {v:var} to text {d:before|after} last {a:expr} in {t:expr}
set {v:var} to trimmed {t:expr}
```

`after any space` makes `Don't Stop` and `O'neil`.
`after any non-letter` makes `Don'T Stop` and `O'Neil`. The closing words
are required, since the two differ.

### D. Text changed where it stands

```text
capitalize each word in cell|range {r:range} after any {k:space|non-letter}
capitalize each word in column {c:column} after any {k:space|non-letter}
make cell|range {r:range} {k:upper|lower} case
make cell|range {r:range} {k:uppercase|lowercase}
make column {c:column} {k:upper|lower} case
make column {c:column} {k:uppercase|lowercase}
remove extra spaces from cell|range {r:range}
remove extra spaces from column {c:column}
remove non-printing characters from cell|range {r:range}
remove non-printing characters from column {c:column}
replace {a:expr} with {b:expr} in column {c:column}
replace {a:expr} with {b:expr} in formulas of column {c:column}   (0.7.1)
replace {a:expr} with {b:expr} in formulas of range {r:range}   (0.7.1)
replace {a:expr} with {b:expr} in formulas on this sheet   (0.7.1)
replace {a:expr} with {b:expr} in range {r:range}
replace {a:expr} with {b:expr} on this sheet   (0.7.1)
split column {c:column} by {d:expr} as text   (0.7.1)
split column {c:column} by {d:expr} reading numbers   (0.7.1)
```

`replace … in column|range`: from `0.7.1` it changes values and never a
formula. In `0.7.0` it is Excel's own Replace, which reaches into
formulas. If the workbook's release is not known, say which behaviour
the person should expect.

### E. Fonts, fills and looks

```text
band every other row of {r:range} {n:expr}
clear color of cell {r:cell}
clear fill-color of cell|range {r:range}
clear formatting of range {r:range}
copy column widths of {a:range} to {b:range}
copy formatting of {a:range} to {b:range}
make {a:range} look like {b:range}
make cell {r:cell} {d:bold|italic}
make cell {r:cell} {d:red|yellow|black|blue|cyan|green|magenta|white}
make cell in column {c:column} row {n:expr} bold
make cell|range {r:range} not {d:bold|italic}
make range {r:range} {d:bold|italic}
make range {r:range} {d:red|yellow|black|blue|cyan|green|magenta|white}
make row {r:expr} a header row
paint cell {r:cell} {e:expr}
set fill-color of cell {r:cell} to {e:expr}
set fill-color of range {r:range} to {e:expr}
set font of cell|range {r:range} to {n:text}
set font size of cell {r:cell} to {n:expr}
set font size of range {r:range} to {n:expr}
set font-color of cell {r:cell} to {e:expr}
set font-color of range {r:range} to {e:expr}
strike through cell|range {r:range}
underline cell|range {r:range}
```

`make cell C4 red` colours the **fill**, not the text. For the text, use
`set font-color of …`.

### F. Number formats

```text
format cell {r:cell} as {d:currency|percent|date}
format cell|range {r:range} as {d:dollars|euros|pounds}
format cell|range {r:range} as {d:dollars|euros|pounds} with {n:expr} decimal/s
format cell|range {r:range} as {d:short|long|iso} date
format cell|range {r:range} as accounting in {d:dollars|euros|pounds}
format cell|range {r:range} as accounting in {d:dollars|euros|pounds} with {n:expr} decimal/s
format cell|range {r:range} as general
format cell|range {r:range} as number
format cell|range {r:range} as number with {n:expr} decimal/s
format cell|range {r:range} as number with {n:expr} decimal/s and thousands separators
format cell|range {r:range} as number with thousands separators
format cell|range {r:range} as percent with {n:expr} decimal/s
format cell|range {r:range} as text for new entries
format cell|range {r:range} as time
format cell|range {r:range} using {n:text}
format range {r:range} as percent
```

A format changes how a number is shown and never the number.

### G. Alignment, size and layout

```text
{d:merge|unmerge} range {r:range}
{d:wrap|unwrap} text in range {r:range}
align cell|range {r:range} {d:left|right|center/ed}
align cell|range {r:range} to {d:top|middle|bottom}
center cell|range {r:range}
fit all columns
fit column {c:column}
indent cell|range {r:range} by {n:expr}
rotate text in cell|range {r:range} by {n:expr} degrees
set height of row {n:expr} to {h:expr}
set width of column {c:column} to {w:expr}
```

### H. Borders

```text
add {d:top|bottom|left|right} border colored {c:color} to cell|range {r:range}
add {d:top|bottom|left|right} border to cell|range {r:range}
add border around cell|range {r:range}
add border colored {c:color} around cell|range {r:range}
add border to range {r:range}
add borders colored {c:color} to every cell in [range] {r:range}
add borders to every cell in [range] {r:range}
remove {d:underline|strikethrough|borders} from cell|range {r:range}
```

`add border to range` draws every cell's lines. It is the older
spelling. Prefer the two that say which: `around`, or `to every cell in`.

### I. Rows and columns

```text
{d:hide|unhide} column {c:column}
{d:hide|unhide} row {n:expr}
copy row {r:expr} to row {n:expr}
delete {n:expr} row
delete {r:range} and shift cells left
delete {r:range} and shift cells up
delete blank rows in {r:range}
delete column {c:column}
delete row {n:expr}
delete rows {a:expr} through {b:expr}
fit row {r:expr}
freeze the first {n:expr} rows
freeze top row
group rows {a:expr} through {b:expr}
insert {n:expr} rows at row {r:expr}
insert column before {c:column}
insert row [at] {n:expr=1}
move column {c:column} before column {d:column}
remove trailing empty rows and columns
unfreeze panes
ungroup rows {a:expr} through {b:expr}
unhide all rows and columns
```

`delete {n:expr} row` is for an ordinal: `Delete the third row.`

**`delete blank rows in` deletes every row that has an empty cell
anywhere in the range given**, and deletes the whole row of the sheet.
Give it the one column that decides: `Delete blank rows in A2:A500.` Say
so to the person when a draft contains it.

When deleting rows in a loop, count **down**, so that the rows not yet
visited keep their numbers.

### J. Sorting, filtering and duplicates

```text
add filter/s to range {r:range}
clear filter condition/s
copy only visible cells of [range] {a:range} to [cell] {b:cell}
filter [range] {r:range} to show rows where column {c:column} contains {v:expr}
filter [range] {r:range} to show rows where column {c:column} is {d:greater|less} than {v:expr}
filter [range] {r:range} to show rows where column {c:column} is {v:expr}
keep only rows of range {r:range} where column {f:expr} is {v:expr}
remove duplicates from range {r:range}
remove duplicates from range {r:range} by column {k:expr}
remove filter/s
show all rows
sort [range] {r:range} by column {c:column} {d:ascending|descending} {h:with|without} header [row]
sort [range] {r:range} by column {c:column} {d:ascending|descending} then [by] column {e:column} {f:ascending|descending} {h:with|without} header [row]
sort [range] {r:range} by column {c:column} {h:with|without} header [row]
sort [range] {r:range} by column {c:column} then [by] column {e:column} {h:with|without} header [row]
sort [this] sheet by column {c:column} {d:ascending|descending} {h:with|without} header [row]
sort [this] sheet by column {c:column} {h:with|without} header [row]
sort range {r:range} by column {k:cell} [ascending]
sort range {r:range} by column {k:cell} descending
```

**A sort must say `with a header row` or `without a header row`.** With
two keys, both directions are named or neither is. The last two shapes
are older: they take a *cell* in the key column, `by column B1`, and
assume a header row. Prefer the shapes that say.

A filter hides rows and deletes nothing. **`keep only rows …` is a filter
too**, whatever the words suggest: the other rows are hidden, and
`Show all rows.` brings them back. Its column is a number, the column's
position within the range: `Keep only rows of range "A1:D50" where column
3 is "West".` Prefer `filter … to show rows where …`, which says what it
does. `remove duplicates` does delete.

### K. Excel Tables

```text
add a row to table {n:text}
delete row {r:expr} of table {n:text}
hide the total row of table {n:text}
set {v:var} to column {c:text} of table {n:text}
set style of table {n:text} to {s:text}
show the total row of table {n:text}
turn {r:range} into a table called {n:text}
turn table {n:text} back into a range
```

### L. Pivot tables

```text
{d:add|remove} a blank row after {f:text-list} in pivot {n:text}
{d:hide|show} subtotals for {f:text-list} in pivot {n:text}
add {d:rows|columns|filters} of {f:text-list} to pivot {n:text}
add {f:text-list} to pivot {n:text} as {d:sum|count|average}
change the source of pivot {n:text} to {r:range}
clear pivot {n:text}
collapse {f:text-list} in pivot {n:text}
delete pivot {n:text}
expand {f:text-list} in pivot {n:text}
make a pivot table from {r:range} at {b:cell} called {n:text}
make a pivot table from {r:range} at {b:cell} called {n:text} with rows of {a:text-list} and columns of {c:text-list} and values of {v:text-list}
refresh every pivot table
refresh pivot {n:text}
remove {f:text-list} from pivot {n:text}
rename pivot {n:text} to {m:text}
show pivot {n:text} in {d:compact|tabular|outline} form
sort {f:text} in pivot {n:text} {d:ascending|descending}
sort {f:text} in pivot {n:text} {d:ascending|descending} by {v:text}
```

A field whose name has a space goes in quotes. A pivot table begins with
`Make`, never `Create`.

### M. Sheets, workbooks and files

```text
{d:protect|unprotect} sheet {s:sheet} with password {p:expr}
{d:protect|unprotect} this sheet with password {p:expr}
add [new] sheet called {s:sheet}
close this workbook
copy range {r:range} to sheet {s:sheet}
delete sheet {s:sheet}
export this sheet as pdf {p:expr}
go to sheet {s:sheet}
open workbook {p:expr}
print this sheet
put {e:expr} into|in cell {r:cell} of sheet {s:sheet}
refresh everything
save a copy as {p:expr}
save this workbook
save this workbook as {p:expr}
set tab-color of sheet {s:sheet} to {e:expr}
work on sheet {s:sheet}
```

### N. The person at the keyboard, and the outside world

```text
ask {q:expr} and put answer into {v:var}
clear status bar
email {who:expr} with subject {s:expr} and message {m:expr}
email {who:expr} with subject {s:expr} and message {m:expr} and attachment {p:expr}
make a button [called] {c:text} at cell {r:cell}
put {e:expr} in status bar
turn off screen updating
turn on screen updating
wait {n:expr} seconds
```

`email …` opens a draft for the person to read and send. Nothing is sent
silently.

Words put in the status bar, and screen updating turned off, last until
the run ends; the run then puts both back as it found them. A message the
person must read after the run is a `show`.

**Sentences that reach outside the workbook** are the ones that open,
save, print, export, email or protect. **Undo Last Run** does not reach
what they did. In a workbook that Windows has marked as downloaded, they
are refused until the file is unblocked. Tell the person when a draft
contains one.

### O. Questions of your tables, and logic

```text
optimize cell {r:cell}
show every {shown:text} in {table:text} with {filtercol:text} over {val:expr} as {alias:text} in cell {r:cell}
show every {table:text} {shown:text} whose {filtercol:text} is over {val:expr} as {alias:text} in cell {r:cell}
show everyone who reports to {person:expr} directly or not in {table:text} as {alias:text} in cell {r:cell}
show in cell {r:cell} {q:question} by applying the rules in {rules:range} to the data tables {t:text-list}
write in cell {r:cell} that {h:clause}
```

`optimize cell` runs only under **Interpret and Run**.

---

## R8. Logic sentences

A rule or a fact is written into a cell. A question is asked of a range
of such cells, and of Excel Tables. The answer is a live formula.

```text
Write in cell <cell> that <clause>.
Show in cell <cell> <question> by applying the rules in <range> to the data tables <list>.
```

### Words

| Kind | Is | Examples |
|---|---|---|
| **Role noun** | The variable. One word, not quoted. The same noun is the same thing throughout the sentence. | `person` · `shift` · `bill` |
| **Relation** | One word, its parts joined by hyphens. | `can-cover` · `reports-to` · `has-tier` |
| **Set** | One word. | `listed` · `covered` · `big` |
| **Constant, text** | In quotes. | `"Night"` · `"Alice"` |
| **Constant, number** | Bare. | `10000` |

**Text is always quoted. A number is never quoted.** `"10000"` is
refused. A bare `Night` is a role noun and not the text.

Two different things need two different nouns: `a person reports-to a
boss`.

### A clause

```text
<subject> is <set>
<subject> <relation> <object>
<subject> is <set> if <conditions>
<subject> <relation> <object> if <conditions>
<subject> <relation> <object> if <conditions>, otherwise <object> if <conditions>
```

With no `if`, the clause is a fact and every part of it is a constant:
`"B3" is flagged`.

Every role noun in the head must be bound by a condition.

### Conditions

Joined by `, and` or `and`. Each is one of these, and any may begin with
`not`.

| Shape | Example |
|---|---|
| A row of a table | `Staff lists the person as Name, the cert as Cert, and the level as Level` |
| A relation | `the person can-cover the shift` |
| A relation, followed along | `the person reports-to the boss directly or not` |
| Membership of a set | `the shift is listed` |
| A comparison | `the level is at least the min` · `the amount is greater than 10000` |
| A text test | `the code starts with "GL-4"` |

In a row of a table, the word after `as` is the column's heading.

Comparisons: `is at least` · `is at most` · `is greater than` ·
`is less than`. Text tests: `starts with` · `ends with` · `contains`.

**There is no equality.** To require a value, put the constant in the
row: `Roster lists the person as Name and "Ops" as Dept`.

A comparison or a text test binds nothing. Its role nouns must be bound
by a row or a relation in the same sentence.

### Questions

| Write | Answers |
|---|---|
| `who can-cover "Night"` | a list |
| `what "Bob" can-cover` | a list |
| `which bill is a violation` | a list |
| `which person can-cover which shift` | every pairing |
| `whether "Bob" can-cover "Night"` | TRUE or FALSE |
| `who reports-to "Alice" directly or not` | a list, following the chain |
| `how many people can-cover "Night"` | one number |
| `how many people can-cover each shift that is listed` | a count beside each member |
| `which shift that is listed is not covered` | what is outside |
| `whether every shift that is listed is covered` | TRUE or FALSE |
| `who alone can-cover "Night"` | the one, when there is exactly one |
| `who can-cover "Night" as one list` | the answers in one cell |

### A whole example

```text
Write in cell M2 that a person can-cover a shift if Shifts lists the shift as Shift, the cert as Needs, and the min as MinLevel, and Staff lists the person as Name, the cert as Cert, and the level as Level, and the level is at least the min, and not Leave lists the person as Name and the shift as Shift.
Write in cell M3 that a shift is listed if Shifts lists the shift as Shift.
Write in cell M4 that a shift is covered if the person can-cover the shift.

Show in cell A4 who can-cover "Night" by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
Show in cell F4 which shift that is listed is not covered by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
Show in cell H4 whether every shift that is listed is covered by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
```

### Rules to keep

1. **The tables must be real Excel Tables**, and the list names them as
   Excel does.
2. **The range of rules must include every cell that defines a relation
   the question reaches.** Two rule sentences with the same head are
   "or". A question whose range holds one and not the other is refused.
3. **One rule to a cell.**
4. **An `otherwise` chain lives in one sentence**, and its relation is
   written in no second cell.
5. An answer that is a list needs room to spill, and an Excel with
   dynamic arrays.

---

## R9. The worksheet functions

The sentences of R8 write these for you. A person may also type them as
formulas. Each takes its program as text, then the tables it reads.
Errors come back in the cell, in words: `#DATALOG! …`

Inside a formula, a quotation mark within the program is doubled, as
Excel requires.

### `=SQL(query, table, …)`

```text
=SQL("SELECT Name, Salary FROM staff WHERE Salary > 80000", Staff)
```

| Supported | Not supported, and refused in words |
|---|---|
| `SELECT`, `DISTINCT`, `AS` | `LEFT`, `RIGHT`, `FULL`, `OUTER`, `CROSS JOIN` |
| `FROM`, `JOIN … ON`, `INNER JOIN` | A table joined to itself |
| `WHERE`, `AND`, `OR`, `NOT`, comparisons | A subquery in `FROM`, `WHERE` or `SELECT` |
| `+ - * /`, and `ABS SIGN SQRT TRUNCATE FLOOR CEILING ROUND MOD REMAINDER DIV LEAST GREATEST POWER` | A computed column with no `AS` |
| `COUNT SUM MIN MAX AVG`, `GROUP BY`, `HAVING` | An aggregate in `WHERE` or `ON` |
| `ORDER BY`, `LIMIT` | |
| `UNION`, `UNION ALL`, `INTERSECT`, `EXCEPT` | |
| `WITH`, and `WITH RECURSIVE` as `base UNION ALL step` | |

Text in a query is in single quotes. Tables must be Excel Tables.

### `=DATALOG(program, table, …)`

```text
=DATALOG("(fact (parent tom bob)) (fact (parent bob liz))
          (rule (ancestor X Y) (parent X Y))
          (rule (ancestor X Z) (parent X Y) (ancestor Y Z))
          (query ancestor)")
```

| Form | Rule |
|---|---|
| `(fact (p a b))` | Every argument is a value. No variables. |
| `(rule (head X Y) body …)` | At least one body item. Every variable in the head appears in the body. |
| `(query name)` | The relation's rows. Exactly one query in a program. |
| `(query (p "a" "b"))` | TRUE or FALSE. Every value written in. |
| `(query (not (p "a" X)))` | TRUE when nothing matches. |
| `(headless)` | The rows without their heading. |
| A variable | A word beginning with a capital. |
| A constant | A word beginning with a small letter, a number, or quoted text. Text is compared exactly, capitals included. |
| A table | Named as Excel names it. Read by position, `(staff X Y Z)`, or by heading, `(staff (name X) (salary S))`. Not both in one atom. |

| In a rule's body | |
|---|---|
| `(p X Y)` | An atom. |
| `(not (p X))` | Every variable in it must be bound by an atom written earlier. |
| `(> S 80000)` · `<` `<=` `>=` `=` `<>` | A comparison, over bound values. |
| `(let Z (+ X Y))` | Arithmetic into a new variable. |
| `(count N (p X Y))` · `(sum S (p X W))` | An aggregate into a new variable. |
| `(textjoin Names ", " (p Name "Night"))` | The values joined in one text. |
| `(text-starts-with T S)` · `(text-ends-with T S)` · `(text-contains T S)` | Text tests. |

A predicate has one number of arguments everywhere. No predicate may be
nested inside another's argument. A predicate may not depend on itself
through `not`, `count`, `sum` or `textjoin`.

### `=PROLOG(program, table, …)`

```text
=PROLOG("(rule (can-cover Shift Who)
           (shifts (shift Shift) (needs Cert) (minlevel Min))
           (staff (name Who) (cert Cert) (level Level))
           (>= Level Min)
           (not (leave (name Who) (shift Shift))))
         (query (can-cover Shift Who))", Shifts, Staff, Leave)
```

Unification and backtracking. Goals include `is`, `not`, `findall`,
`between`, `or`, `if`, cut, the comparisons, list goals (`length`
`member` `nth` `append` `reverse` `sum-list`) and text goals. It can
answer what DATALOG cannot, and it can fail to finish, in which case it
stops at a ceiling and says so. **Prefer DATALOG** unless the question
needs lists, or arithmetic over a search.

### `=OPTIMIZE(program, table, …)`

Rules, as in DATALOG, and then choices and constraints.

```text
(choose-exactly N (assign S P) (free S P) (per (shifts S W D Slot N K)))
(choose-at-most 5 (assign S P) (free-week S P W) (per (person-week P W)))
(forbid (assign S P) (assign T P) (next S T))
(query assign)
```

| Form | |
|---|---|
| `choose-exactly` · `choose-at-least` · `choose-at-most` · `choose-between` · `choose-any` | How many rows to choose from a pool, for each group named by `(per …)`. |
| `require` · `forbid` | What must hold, and what may never all hold at once. |
| `(effort quick)` · `normal` · `thorough` · a number | How much work the search may do. Work, never seconds. |
| `prefer` · `avoid` · `minimize` · `maximize` · `fewest-changes-from` | **Refused in this release**, by name. |

`=OPTIMIZE_STATUS(…)` and `=OPTIMIZE_VIOLATIONS(…)`, given the same
arguments, say how the search ended and which rows broke which rule.
The answer is the first assignment that breaks no rule, in the tables'
own order. The engine is partly built. Do not promise a best answer.

---

## R10. Refusals

A refusal stands beside the row, in column C, on red. Nothing has run.
**Validate stops at the first row it cannot read.** Rows above are
green. Rows below have not been judged.

### The frame

> I understood '…' - then I expected … but found '…'. Did you mean: '…'

| Part | Use it to |
|---|---|
| `I understood '…'` | Keep these words. They are right. |
| `I expected …` | Find what belongs at that point. |
| `but found '…'` | Find the word to change. `the end of the sentence` means something is missing. |
| `Did you mean: '…'` | See the nearest shapes, in R7's notation. |

> Don't understand: '…' No loaded sentence starts with '…' …

Here the first word begins no sentence at all. Choose another verb from
R7.

### Repairs

| Frazaro says | Do |
|---|---|
| expected one of 'into'/'in' | Use that word. |
| expected a range (like A1:C50 …) | Write a range, or quote a named one. |
| expected a column letter (like C or AA) | Write the letter. |
| expected a cell (like B2 …) | Write a cell. |
| expected a color (like red or yellow …) | Use one of the eight words, or a quoted code. |
| expected one of 'with'/'without' | Say whether there is a header row. |
| expected 'after' | Add `after any space` or `after any non-letter`. |
| expected one of 'ascending'/'descending' | Name both directions, or neither. |
| I don't understand the character '…' | Replace the symbol with its word, or quote the text. |
| mixes % into a longer amount | Compute the amount under a name first. |
| '…' is a reserved word in Excel's programming language | Hyphenate the name. |
| '…' was given a fixed value by Define and cannot be changed | Use another name. |
| the action '…' has no parameter called '…' | Use a detail the step defines. The message lists them. |
| the action '…' requires '…' | Add `with <detail> of <value>`. |
| 'Give back' only makes sense inside a value-returning action | Move it into a `To … of …:` step. |
| 'Stop the loop.' only makes sense inside a loop | Use `Stop.`, or move it. |
| 'If that fails:' must start its own paragraph right after the Try block | Put a blank line before it, and a `Try:` block above. |
| a button isn't created with 'Create ... called ...' | `Make a button "<caption>" at cell <cell>.` |

### When the run stops

A sentence that validates can still fail while running: a sheet is
missing, a lookup finds nothing, a median has no numbers to work on.

> The run stopped at line 7: …
> Put back as they were before the run: …
> Anything it did anywhere else, like a file saved or an email drafted,
> stays as it is.

When one of Excel's own functions had no answer, the words after the
sentence are, for now, Excel's, naming the function: `Unable to get the
Median property of the WorksheetFunction class`.

Guard such a sentence with `Try:`, or make the program create what it
needs.

---

## R11. What looks right and is wrong

A model's habits from other languages are the usual cause.

| You are tempted to write | Write |
|---|---|
| `Set total to price * qty.` | `Set total to price times qty.` |
| `Set total to (a plus b) times c.` | Two sentences, with a name for `a plus b`. |
| `If total > 5, …` | `If total is greater than 5, …` |
| `If total = 5 then …` | `If total is 5, …` |
| `Else:` · `Else if …:` | `Otherwise:` · `Otherwise, if …:` |
| `End if.` · `End.` · `Next.` | A blank line, or `Done.` |
| `For r = 2 to 10:` | `Count r from 2 to 10:` |
| `Loop 5 times:` | `Repeat 5 times:` |
| `// note` · `' note` | `# note` |
| `Set order total to 5.` | `Set order-total to 5.` |
| `Set date to today.` | `Set run-date to today.` |
| `Work on sheet Q1 Data.` | `Work on sheet "Q1 Data".` |
| `Put 5 into B2.` | `Put 5 into cell B2.` |
| `Set A1 to 5.` | `Put 5 into cell A1.` |
| `Set total to A1 plus 1.` | `Set total to cell A1 plus 1.` |
| `Set q1 to 5.` | `Set q1-total to 5.` |
| `Put 5 into cell $B$2.` | `Put 5 into cell B2.` |
| `Set x to range A1:A9.` | `Remember range A1:A9 as x.` |
| `Set x to sum of A1:A9.` | `Set x to sum of range A1:A9.` |
| `Put formula "=B2*C2" into D2:D50.` | `Put formula "=B2*C2" into range D2:D50.` |
| `Set fill-color of cell C4 to red.` | `Make cell C4 red.` |
| `Make cell C4 orange.` | `Set fill-color of cell C4 to "#FFA500".` |
| `Sort range A1:C50 by column B.` | `Sort range A1:C50 by column B with a header row.` |
| `Capitalize each word in range A2:A50.` | `… after any space.` |
| `Set s to standard deviation of range B2:B50.` | `… as a sample.` or `… as the population.` |
| `… where range A2:A90 is "West".` | `… where range A2:A90 matches "West".` |
| `Create a button called "Go" at cell D3.` | `Make a button "Go" at cell D3.` |
| `Create a pivot table …` | `Make a pivot table from … at … called …` |
| `Save this workbook as C:\Out\a.xlsx.` | `Save this workbook as "C:\Out\a.xlsx".` |
| `… to the data tables Staff and Shifts.` | `… Staff, Shifts.` |
| `… has-level "3"` | `… has-level 3` |
| `… can-cover Night` | `… can-cover "Night"` |
| `… the person can cover the shift` | `… the person can-cover the shift` |
| `Show "Done."` | `Show "Done.".` |
| `Set x to minus 5.` | `Set x to -5.` |
| `Put a 5 into cell B2.` | `Put 5 into cell B2.` |
| A blank line inside a loop | A `#` note |
| Two sentences on one row | One row each |

And two that are accepted and mean something else:

| You wrote | It means | You may have meant |
|---|---|---|
| A sentence directly under a block, with no blank line | It is inside the block. | Put a blank line first. |
| `cell in column C row r plus 1` | The cell in row r+1. | Name the value, then add. |

---

## R12. Worked translations

Each of these is a shipped sample, and every sentence is the sample's
own. The second is abridged. The request is how a person might put it.
The answer is what runs.

### One

*"Every Monday the sales system exports last week's orders as a plain
sheet called Sales. Make it readable: header row, money in column G,
columns wide enough, header stays put when I scroll."*

```text
# Tidy the Sales Export
Work on sheet "Sales".
Make row 1 a header row.
Format range G2:G41 as dollars.
Fit all columns.
Freeze top row.
```

Assumptions to state: the sheet is named Sales; row 1 holds headings;
column G holds the amounts, in rows 2 to 41.

### Two

*"Go through the expense claims on the Expenses sheet. Amount is column
D, receipt is column E and says Yes or No. Anything over 500 needs a
manager. Anything with no receipt is a problem, and that matters more.
Put the finding in column F and tell me how many of each."*

```text
# Expense Report Audit
Work on sheet "Expenses".
Put "Check" into cell F1.
Set last-row to last filled row of column A.

# A later line overwrites an earlier one, so the most serious
# finding is the one that stays.
Count r from 2 to last-row:
  Set amount to cell in column D row r.
  Set receipt to cell in column E row r.
  Put "OK" into column F row r.
  If amount is greater than 500, put "Needs manager approval" into column F row r.
  If receipt is "No", put "Missing receipt" into column F row r.
  If receipt is "No", make cell in column B row r bold.

Set missing-receipts to count of range F2:F26 matching "Missing receipt".
Set need-approval to count of range F2:F26 matching "Needs manager approval".
Show "Expense check done: " joined with missing-receipts joined with " claims are missing a receipt, and " joined with need-approval joined with " need a manager's approval".
```

Things to notice. The length of the data is found, not assumed. Each
value read from the sheet is given a name before it is tested. "Matters
more" became an order of sentences, with a note that says so. The blank
line after the loop is what ends it.

### Three

*"This is our Friday procedure. Steps 1 and 2 can be automated. Step 3
is a person's job."*

```text
Weekly Expense Reimbursement
Owner: Accounts Payable. Performed every Friday afternoon.

PROCEDURE
1. Open the week's expense report.
2. Add up every claim that has a receipt, and write the total beside the report.
3. Send the total to payroll before 3pm.

<Frazaro week="Week 39" (steps 1 and 2, automated by Accounts Payable)>
Work on sheet "Expenses".
Set receipted-total to sum of range D2:D26 where range E2:E26 matches "Yes".
Put "Reimbursement total, {week}" into cell H10.
Put receipted-total into cell I10.
Format cell I10 as dollars.
Make range H10:I10 bold.
Show "{week} reimbursement total: " joined with receipted-total.
</Frazaro>
```

The procedure stays as its owner wrote it. Two tags were added. Step 3
is left as prose, outside the tags, where a person will read it.

### Four

*"Who can cover the night shift? People are in a table called Staff,
shifts in Shifts, and leave in Leave. Someone can cover a shift if they
hold the certificate it needs at the level it needs, and are not on
leave for it."*

The three rules and the first question of R8's whole example.

Assumptions to state: the three are Excel Tables; their column headings
are Shift, Needs and MinLevel; Name, Cert and Level; Name and Shift; the
text in the Shift column is `Night`, with that capital.

---

## R13. When no sentence exists

It will happen often. The grammar is small on purpose.

1. **Look again in R7** under another verb. *Highlight* is `make … yellow`.
   *Autofit* is `fit column`. *Total* is `sum of`. *Rename* a sheet has
   no sentence; *rename pivot* has.
2. **Compose it from sentences that exist.** Most tasks are a loop, a
   condition and a `Put`.
3. **If it cannot be said, do not approximate.** Write a note:

   ```text
   # NOT AUTOMATED: add a data-validation list to column C.
   # No sentence exists for this yet. Do it by hand, or request one.
   ```

4. **Tell the person what to do with the gap.**
   - Frazaro logs every sentence it could not read. **Copy Diagnostic
     Report** puts that log on the clipboard, to be pasted into an email
     to the address in [SUPPORT.md](../SUPPORT.md). That is how the
     language learns a sentence.
   - A team may teach Frazaro a sentence of its own in a phrasebook.
     [Level 2](2-journeyman.md) is the manual.

If the person asks you to draft a phrasebook rule, follow Level 2, write
a proof for every rule, and say plainly that the loader runs the proofs
and that you have not.

---

## R14. The layers beneath, in brief

| Layer | Is | Manual |
|---|---|---|
| **Phrasebook** | A `.vla` text file of rules. A rule is a pattern, a template and its proofs. It is loaded with **Load Phrasebook**, after which its sentences appear in **What can I say?** | [Level 2](2-journeyman.md) |
| **VLA** | The middle language every sentence is translated into. A Lisp that maps one to one onto VBA. | [Level 3](3-master.md) |
| **Interpreter** | The default way to run. It acts on the workbook and writes no code. | |
| **Compiler** | Writes readable VBA, each line tagged with its sentence. It needs Excel's trust in the VBA project. | |

```lisp
(english-vla
    "make cell {r:cell} {d:bold|italic}"
    (make-{d} (range {r})))
(test-success
    "Make cell A1 bold."
    (make-bold (range "a1")))
```

### The buttons, by their labels

| Group | Buttons |
|---|---|
| **IDE** | New Frazaro · New Named Frazaro · What can I say? · Load Phrasebook · Load Instructions · Reload Instructions |
| **Testing** | Validate Instructions · Undo Last Run |
| **Interpreter** | Interpret and Run · Interpret and Trace |
| **Compiler** | Compile and Run · Compile and Trace |
| **Auditing** | Show me the VBA · Translate File to VLA · Translate File to VBA · Export Expanded Phrasebook · Export Phrasebook Test Coverage · Lint VLA File |
| **Utilities** | Open CLI · Logic Engines · Register for Auto-Load · Uninstall Frazaro · Copy Diagnostic Report · Forget Phrasebook Approvals |

Use these labels exactly when telling a person what to press.

### The sheet

| Column | Holds |
|---|---|
| B | The sentences, one to a row. |
| C | The verdicts: `OK`, or a refusal. |

### What is true of every run

- The sheets a program names are copied first. **Undo Last Run** puts
  them back, and reaches the most recent run only.
- A sheet the run created, the sheet it works on included, is removed
  again by a stop or by Undo.
- Undo does not reach a file saved, an email drafted, or another
  workbook.
- A run that stops puts the sheets back and names the row.
- Excel's own settings a program changes (automatic calculation, the
  status bar, screen updating) are put back as the run found them when
  it ends, finished or stopped. They never outlast the run, so Undo has
  nothing to do with them. A setting meant to last is typed into
  **Open CLI**, or set in Excel.
- Nothing is sent anywhere. Frazaro makes no network call.

---

## R15. How to answer

### The shape of an answer

1. **The program**, in one fenced block, one sentence to a row. If the
   person's procedure lives in a document, include the two tags.
2. **Assumptions**, as a short list: sheets, columns, tables, where the
   data ends.
3. **Anything not automated**, and why.
4. **Any sentence that reaches outside the workbook**, named.
5. **Any shape marked `(0.7.1)`**, named, since an older copy will
   refuse it.
6. **What to do next**, in these words or near them:

   > Paste this into a *Frazaro* sheet, or load it with **Load
   > Instructions**, and press **Validate Instructions**. If a row turns
   > red, send me the message beside it, word for word.

### When a refusal comes back

1. Find the row.
2. Read the frame (R10). Keep what was understood.
3. Change what was found into what was expected.
4. Return the corrected row, and the rows that depend on it. Leave the
   rest alone.
5. If the same row is refused twice, stop guessing. Say that the
   sentence may not exist, and go to R13.

### Before you send

- [ ] Every sentence matches a shape in R4, R7 or R8.
- [ ] Every sentence ends with a period.
- [ ] Every block is followed by a blank line, and holds none.
- [ ] Every piece of text that must match data is quoted.
- [ ] Every name is one word, and none is reserved.
- [ ] No symbol stands outside quotes.
- [ ] No parentheses. No `a` or `an` before a value.
- [ ] Every list has its commas.
- [ ] The program says which sheet it works on.
- [ ] Every sort says whether there is a header row.
- [ ] Nothing was invented.

### What not to claim

| Do not say | Say |
|---|---|
| "This will work." | "This has not been validated. Frazaro will say." |
| "Frazaro's AI understands …" | "Frazaro matches sentence shapes. It has no model." |
| "I ran it." | You did not. |
| "Frazaro will figure out what you mean." | "Frazaro refuses what it cannot read, and says why." |

---

## R16. Version and provenance

| | |
|---|---|
| **Release described** | `0.7.0`: 223 sentence shapes, of which 220 are written out in the phrasebook and three are written by a generator. |
| **Marked `(0.7.1)`** | Sixteen shapes, and the two range conditions of R6. With them R7 lists 239, which is the number **Load Phrasebook** reports as *built-in vocabulary* in a copy that has them. |
| **Marked `(0.8.0)`** | Five shapes: a range's median and standard deviation, each set and put into a cell, and an average over the rows that match. With them R7 lists 244. |
| **Source of R7** | The project's own dated ledger of every shape, [GRAMMAR_SINCE.md](../GRAMMAR_SINCE.md), including three shapes that a generator writes and no file spells out. |
| **Source of R4, R5, R6, R8** | The engine's source, and the proofs the phrasebook runs every time it loads. |
| **Source of R12** | The shipped sample procedures. Every sentence is a sample's own. |
| **The authority** | **What can I say?**, in the copy of Frazaro at hand. It is generated from the grammar loaded at that moment and cannot be out of date. Where this page and that list disagree, the list is right. |
| **A shipped sentence keeps its meaning.** | A page like this one goes out of date by being incomplete. It does not go out of date by being wrong about a sentence it lists, with the one narrowing recorded under R7, part D, and one retirement recorded under R2: from `0.7.1` a name may not be shaped like a cell. |
| **Licence** | CC-BY-4.0, as for every document in this folder. |
| **Where to write** | `english@spreadsheet.company` |

---

Previous: [Level 4 · Wizard](4-wizard.md) ·
Back to: [Level 0 · Introduction](0-introduction.md)
