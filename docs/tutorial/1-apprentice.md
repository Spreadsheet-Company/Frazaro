# Level 1 · Apprentice: A Human's Guide to Programming in English

*(Do spreadsheets dream of electric paydays?)*

| | |
|---|---|
| **For** | The person who owns a procedure: an accountant, an analyst, an operations lead. No programming is assumed, and none is taught. |
| **You need** | Excel for Windows, the Frazaro add-in, and the [`examples`](../../examples/README.md) folder. Word, if your procedures live in `.docx` files. |
| **You will be able to** | Load a procedure, check it, run it, take the run back, write one of your own, and read a refusal without alarm. |
| **Plan on** | An afternoon, with the sample workbook open beside this page. |
| **Comes after** | [Level 0](0-introduction.md), which is optional. |
| **Written against** | Frazaro `0.7.0`. Sentences that arrive in `0.7.1` or `0.8.0` are marked. |

---

The subtitle asks a question, and a training manual should answer its own
questions. No. A spreadsheet dreams of nothing. It waits, with the patience
of furniture, for somebody to come and do Friday's reimbursement run by
hand, again.

This manual is about writing that procedure down once, in sentences you
could read aloud to a colleague, and letting Excel do it on Friday.

You will not learn to program here, in the sense of learning a notation.
You already know the notation. What you will learn is which English Frazaro
reads, how it tells you when it cannot read yours, and how to take a run
back when the result is not what you wanted.

## Contents

| Part | Lessons |
|---|---|
| **I. The bench** | 1 The bench · 2 The first run · 3 Anatomy of a sentence · 4 A program of your own |
| **II. Saying things** | 5 Values and names · 6 Arithmetic and text, in words · 7 Reading your data · 8 Making it look right · 9 Rows, columns, sheets and files · 10 Sorting, filtering, tables and pivots · 11 Tidying text |
| **III. Deciding and repeating** | 12 Deciding · 13 Repeating · 14 Lists and lookups · 15 Steps of your own · 16 Talking to the person at the keyboard · 17 When something might go wrong |
| **IV. The real thing** | 18 Reading a refusal · 19 Your SOP, as written · 20 Questions of your tables · 21 Running for real |
| **V. Practice** | Exercises · Answers · What an Apprentice can do · The one-page card |

---

# Part I. The bench

## Lesson 1. The bench

### Installing

1. Download
   [Frazaro_English.xlam](https://github.com/Spreadsheet-Company/Frazaro/releases/latest/download/Frazaro_English.xlam).
2. Before opening it, right-click the file, choose **Properties**, tick
   **Unblock** if the box is there, and press OK. Windows marks files that
   arrive from the internet, and Excel declines to open a marked add-in.
3. Open the file. Excel asks, once, whether to trust the publisher. Say yes.
4. A **Frazaro** tab appears on the ribbon. Press **Register for Auto-Load**
   if you want it there every time Excel starts.

That is the whole installation. It writes no code into your workbooks and
asks for no change to Excel's security settings. If your organization
manages your computer, [DEPLOY.md](../DEPLOY.md) and
[IT_REVIEW.md](../IT_REVIEW.md) are written for the people who will ask.

### The Frazaro tab

The tab has six groups. An Apprentice uses the buttons in bold.

| Group | Buttons | What they are for |
|---|---|---|
| **IDE** | **New Frazaro**, New Named Frazaro, **What can I say?**, Load Phrasebook, **Load Instructions**, **Reload Instructions** | Making a place to write, finding out what can be written, and bringing a procedure in from a file. |
| **Testing** | **Validate Instructions**, **Undo Last Run** | Checking before you run, and taking a run back after. |
| **Interpreter** | **Interpret and Run**, Interpret and Trace | Running. This is the ordinary way. |
| **Compiler** | Compile and Run, Compile and Trace | Running by way of generated VBA. Lesson 21. |
| **Auditing** | Show me the VBA, and the file tools | Looking underneath. Levels 2 and 3. |
| **Utilities** | Open CLI, Logic Engines, Register for Auto-Load, Uninstall Frazaro, **Copy Diagnostic Report**, Forget Phrasebook Approvals | Housekeeping. |

### The Frazaro sheet

A program lives on a worksheet named *Frazaro*. **New Frazaro** makes one;
so does **Load Instructions** the first time you use it.

| Column | Holds |
|---|---|
| **B** | Your sentences, one to a row. Row 5 of the sheet is line 5 of the program. |
| **C** | What Frazaro thinks of each row: a green **OK**, or a note in words on a red ground. |
| **D, E, F** | Yours. Put working values there if you like; a sentence can read them. |
| A | Hidden. It remembers which file the program was loaded from, so that **Reload Instructions** knows where to look. |

## Lesson 2. The first run

1. Open **`Frazaro Sample Data.xlsx`** from the `examples` folder.
2. On the **Frazaro** ribbon tab click **Load Instructions** and pick
   **`00 Weekly Expense Reimbursement.docx`**. If this computer has no Word,
   pick the `.txt` of the same name; they hold the same procedure.
3. Look at the *Frazaro* sheet. The procedure is there, one line to a row,
   and every row already has a verdict in column C: loading a file checks it
   at once.
4. Press **Validate Instructions** anyway. It is the button you will press
   most, and it never changes anything. Every row should say OK.
5. Press **Interpret and Run**.

**What you should see.** On the *Expenses* sheet, cell H10 reads
*Reimbursement total, Week 39*. Cell I10 holds 6,099.74, shown as dollars.
Both are bold. A message box reports the total.

6. Press **Undo Last Run**. Frazaro says what it put back. H10 and I10 are
   empty again.

Three things just happened that are worth noticing.

**The document was not rewritten for Frazaro's benefit.** It has a title, an
owner, a purpose and a revision history, none of which Frazaro read. It read
only what sits between `<Frazaro>` and `</Frazaro>`. The first row on the
sheet says so: which lines were read, and how many were left out. Lesson 19
is about those two tags.

**The check came before the run.** Had one sentence been unreadable, nothing
would have run.

**The run could be taken back.** Before a program's first sentence, Frazaro
copies the sheets the program names. **Undo Last Run** puts those copies
back. It covers the most recent run of that program.

## Lesson 3. Anatomy of a sentence

```text
Put 5 into cell B2.
```

A verb, a value, a place, and a period (full stop). Frazaro's sentences are commands, as
the steps of a procedure are.

| The rule | In practice |
|---|---|
| **A sentence ends with a period.** | One sentence, one instruction, one row. |
| **Capitals do not matter**, except inside quotes. | `put 5 into cell b2.` is the same sentence. |
| **"the" and "please" are ignored.** "a" and "an" are skipped wherever a grammar word is expected. | `Please make the cell A1 bold.` is `Make cell A1 bold.` |
| **References are written bare**, as a formula writes them. | `cell B2`, `range A1:C50`, `column C`, `sheet Data` |
| **Quotes mean "this exact text".** Double quotes, always. | `"Sales"`, `"Over 90 days"` |
| **A name with a space in it needs quotes.** | `sheet "Q1 Data"` |
| **Excel's own sheet prefix works too.** | `cell Data!B2`, `cell 'Q1 Data'!A1` |
| **A quote inside text is written twice.** | `"He said ""ok"""` |
| **Your own names are one word**, with hyphens if you like. | `total`, `last-row`, `order-total` |
| **Numbers are written as you would write them.** | `1,000,000` · `3.14` · `15%` · `10 percent` · `-1` |
| **Small numbers may be words.** Zero to twenty, and first to twentieth. | `Delete the third row.` |
| **A line starting with `#` is a note for people.** | `# Done by hand: confirm the bank feed.` |

Two habits from the formula bar do not carry over, and Frazaro says so the
moment you try.

```text
Set total to 5 + 3.
```

> I don't understand the character '+' - write 'plus' for addition

Arithmetic is written in words: `plus`, `minus`, `times`, `divided by`. A
dollar sign is likewise refused, with the advice to write the plain number
and format the cell. The reason is not fussiness. An early version skipped
characters it did not know, read `$5` as 5 and `5 + 3` as "5 3", and was
quietly wrong. It now refuses, which is louder and better.

A name may not be a word that Excel's own programming language keeps for
itself, such as `seek`, `date`, `stop` or `next`. You do not need the list:

> 'seek' is a reserved word in Excel's programming language and cannot be
> used as a name - a hyphenated name like 'seek-row' or 'my-seek' always
> works

Hyphenated names never collide, which is one reason the samples are full of
them.

## Lesson 4. A program of your own

1. With the **`Frazaro Sample Data.xlsx`** sample workbook open, press **New Frazaro**. 
   If a *Frazaro* sheet already holds sample 00, clear column B first, or use **New Named
   Frazaro** to give this program a sheet of its own.
2. Type these five sentences into column B, one to a row.

```text
Work on sheet "Sales".
Make row 1 a header row.
Format range G2:G41 as dollars.
Fit all columns.
Freeze top row.
```

3. Press **Validate Instructions**, then **Interpret and Run**, then look at
   the *Sales* sheet. A raw export has become something you could send.
4. Press **Undo Last Run** and it is a raw export again.

This is sample 01, *Tidy the Sales Export*. It is a complete program.

### Where a program works

`Work on sheet "Sales".` chooses the sheet that the following sentences act
on, and creates it if it is missing. A program that never says so works on a
sheet called **Output**, which Frazaro makes for it.

| Sentence | Does |
|---|---|
| `Work on sheet "Sales".` | Goes to the sheet, creating it first if need be. Safe to run twice. |
| `Go to sheet Data.` | Goes to the sheet. If it does not exist, the run stops and says so. |
| `Add sheet called "Report".` | Makes a new sheet. If one by that name exists, the run stops and says so. |

While a program runs, its own *Frazaro* sheet is protected, so a program
cannot overwrite its own sentences. For the same reason a program may not
name a Frazaro sheet as its workplace:

> 'Frazaro' is a Frazaro sheet - programs can't write where instructions or
> Frazaro's bookkeeping live

### Not sure how to say something?

Press **What can I say?**. Frazaro writes every sentence shape it currently
understands onto a sheet called *Phrasebook*, as an Excel Table with two
columns: the shape, and a worked example. Filter it by the word you are
reaching for. The list is generated from the grammar that is loaded at that
moment, so it cannot be out of date. The running version is on its first
line.

---

# Part II. Saying things

From here on, every lesson is a small vocabulary. Read the sentences aloud.
If one sounds like something you would say across a desk, that is the point.

## Lesson 5. Values and names

Two verbs carry most of the weight, and they are not interchangeable.

> **Put** places a value somewhere on a sheet. **Set** gives a value a name.

```text
Put 5 into cell B2.
Put "Test Report" into cell A1.
Put today into cell D1.
Put formula "=B2*2" into cell B3.
Put 0 into range A1:B5.
Put 5 into cell B2 of sheet Data.

Set total to 0.
Set budget to 1,000,000.
Set remote to cell B2 of sheet Data.
```

`into` may be written `in`. A name that has been Set can be used in any
later sentence: `Put total into cell B2.`

**From `0.7.1`, a name can never look like a cell.** A word of one to three
letters and then digits, such as `A1`, `q1` or `fy24`, is a cell. So
`Set A1 to 5.` is refused before anything runs, and the refusal says what to
write instead:

> 'a1' is shaped like a cell, so it cannot be a name. To mean cell A1, write
> the word cell in front of it: 'Put 5 into cell A1.' or 'Set total to cell
> A1 plus 1.' To keep a value under a name, choose one that is not shaped
> like a cell, like 'a1-total'.

To put 5 in the cell, write `Put 5 into cell A1.` To keep a quarter's figure
under a name, hyphenate it: `Set q1-total to 5.` Sheets, tables and pivots
may still be called `Q1`.

### A formula into many cells

**From `0.7.1`**, a formula goes into a whole range, or into rows the program
counts as it runs:

```text
Put formula "=B2*C2" into range D2:D50.
Set last-row to last filled row of column B.
Put formula "=B2*C2" into rows 2 to last-row of column D.
```

Write the formula for the first cell. The cells below get it adjusted the
way Fill Down adjusts it, so D3 holds `=B3*C3`. `through` may stand for
`to`. When there are no data rows, and the last row is above the first,
nothing is written, so a header row is never touched.

### Four more ways to name something

```text
Remember range B2:B4 as results.
Define hot-pink as "#FF69B4".
Create a number called total.
Create a text called label.
```

| Verb | Names | Notes |
|---|---|---|
| `Set` | a value | The value as it is at that moment. Setting the name again replaces it. |
| `Remember` | a range of cells | The cells themselves, not a copy of what is in them. |
| `Define` | a fixed value | Text in quotes, a number, or an earlier Define. It can never be changed afterwards, and it must stand at the top level of the program, outside any block. |
| `Create` | an empty number, text, value, list or lookup | Only needed when you want the name to exist before you first Set it. |

Trying to change a defined value is refused at validation:

> 'hot-pink' was given a fixed value by Define and cannot be changed - use a
> different name

### Changing a value

```text
Increase total by 5.
Decrease countdown by 1.
Grow total by 10%.
Shrink total by 25%.
```

`Increase` and `Decrease` add and subtract. `Grow` and `Shrink` multiply:
`Grow total by 10%.` makes it one tenth larger.

One idiom is honoured because everybody reaches for it. When the amount is
a plain share, a single number or name wearing `%`, then `Increase` and
`Decrease` mean growing and shrinking too:

```text
Increase growth-check by 50%.
Decrease growth-check by 75%.
```

What Frazaro will not do is guess when a percentage is blended into a
longer amount. `Increase total by 5 plus 5%.` has two readings, so it has
none:

> … mixes % into a longer amount - a plain share works ('Increase total by
> 10%.' grows it by 10%), and a plain number works, but a blend has two
> readings and I will not pick one.

Work the amount out first, give it a name, and increase by the name.

## Lesson 6. Arithmetic and text, in words

| You write | It means |
|---|---|
| `plus`, `minus` | add, subtract |
| `times`, `multiplied by` | multiply |
| `divided by` | divide |
| `joined with`, `followed by` | put two pieces of text end to end |
| `%` or `percent` after a value | that value divided by 100 |

```text
Put counter times 10 into column E row counter.
Set difference to debits minus credits.
Put north-revenue divided by total-revenue into cell C11.
Set thousand-check to 1,000 plus 500.
Set msg to "total is " joined with total.
```

**Order of working.** `%` binds tightest; then `times` and `divided by`;
then `plus` and `minus`; then `joined with`. Within a rank, left to right.
So `2 plus 3 times 4` is 14, as in school.

**There are no parentheses.** When you need a different order, work the
inner part out first and name it:

```text
Set subtotal to price plus shipping.
Set total to subtotal times quantity.
```

This costs a line and buys a name, and the name is usually worth having.

### Value words

These can stand anywhere a value can.

| Word | Gives |
|---|---|
| `today`, `now` | the date; the date and time |
| `length of code` | how many characters |
| `uppercase of name`, `lowercase of name` | the text in capitals, or in small letters |
| `absolute of difference` | the number without its sign |
| `year of`, `month of`, `day of`, `hour of`, `minute of` | that part of a date or time |
| `sum of`, `average of`, `largest of`, `smallest of`, `count of` | over a remembered range or a list |
| `first of`, `last of`, `item 3 of` | one element of a list or a remembered range |
| `cell B2`, `value in cell B2` | what the cell holds |
| `cell in column D row r` | what the cell holds, the row worked out as the program goes |
| `cell in column number 3 row 2` | the same, the column given as a number |

```text
Set n to length of code.
Set m to month of today.
Set shout to uppercase of name-part.
Set v to cell B2 plus 1.
```

A cell is always read with the word `cell` in front of it. A bare `B2` is
refused, as Lesson 5 explains: it could only be a name, and a name can never
look like a cell.

### Rounding

Rounding has a sentence of its own:

```text
Set price to total rounded to 2 decimals.
```

### One thing to know about computed rows

`row` takes everything that follows it as the row number.

```text
Set x to cell in column number 3 row 2 plus 1.
```

That is the cell in row three, which is what the sentence says when read
aloud. To add one to the *value*, name it first:
`Set v to cell in column number 3 row 2.` then `Increase v by 1.`

## Lesson 7. Reading your data

```text
Set grand to sum of range B2:B3.
Set avg to average of range A1:A9.
Set west-total to sum of range B2:B90 where range A2:A90 matches "West".
Set west-count to count of range A2:A90 matching "West".
Set price to lookup of part-code in range A2:C90 column 3.
Set found-row to row of "Widget" in column A.
Set last-row to last filled row of column A.
Set amount to cell in column D row r.
```

| Sentence | Good to know |
|---|---|
| `sum of range … where range … matches …` | The two ranges run side by side. A row counts when its cell in the second range matches. |
| `lookup of … in range … column 3` | Finds the value in the first column of the range and reads across to the third. If the value is not there, the run stops at that sentence. Lesson 17 shows how to catch it. |
| `row of "Widget" in column A` | Gives 0 when nothing is found, so test for it: `If found-row is 0, …` |
| `last filled row of column A` | The sentence that lets a program work on data whose length it does not know in advance. |

One cell is read the same way, with its word in front: `Set amount to cell
D2.` A bare `D2` is refused (Lesson 5).

A range you will use more than once is worth remembering by name. This is
the opening of sample 02:

```text
Work on sheet "Sales".
Remember range G2:G41 as revenues.
Set total-revenue to sum of revenues.
Set orders to count of revenues.
Set average-order to average of revenues.
Set largest-order to largest of revenues.
```

**From `0.7.1`**, finding and counting text anywhere in a range:

```text
Set r to the row of the first cell in range A1:D50 containing "total".
Set n to how many cells in range A2:A99 contain "late".
```

Both ignore capitals, and both give 0 when nothing matches.

**From `0.7.1`**, a range's largest, smallest and average, and a count of
its empty or filled cells, each in one sentence:

```text
Set top to largest of range B2:B50.
Set bottom to smallest of range B2:B50.
Put average of range B2:B50 into cell B51.
Set blanks to count of empty cells in range A2:A50.
Set used to count of filled cells in range A2:A50.
```

`Put average …` (like `largest`, `smallest` and the older `Put sum …`) puts
the number in the cell, not a formula, so it will not change when the data
does. For a cell that follows the data, write the formula:
`Put formula "=AVERAGE(B2:B50)" into cell B51.` A cell whose formula shows
nothing counts as empty and as filled both, as it does in Excel.

**From `0.8.0`**, a range's median and standard deviation, and an average
over the rows that match:

```text
Set middle to median of range B2:B50.
Set spread to standard deviation of range B2:B50 as a sample.
Put standard deviation of range B2:B50 as the population into cell B55.
Set west-average to average of range B2:B90 where range A2:A90 matches "West".
```

A standard deviation says which one it means: `as a sample` (Excel's
STDEV.S) when the rows are some of the cases, `as the population`
(STDEV.P) when they are all of them. Leave it out and the sentence is
refused, with both shown. The average leaves out a matching row that has
no value, as Excel's AVERAGEIF does. When there are too few numbers, a
median of an empty range for one, the run stops at that sentence, as a
lookup that finds nothing does.

## Lesson 8. Making it look right

```text
Make range A1:C1 bold.
Make cell A1 italic.
Make range B8:C8 not bold.
Underline range A10:C10.
Set font size of cell A1 to 18.
Set font of range A14:C14 to "Courier New".
Set fill-color of range A4:C4 to "#1F3A5F".
Set font-color of range A4:C4 to "#FFFFFF".
Center range A1:D1.
Align range A2:A9 left.
Align range A15:C15 to the top.
Wrap text in range A1:A10.
Merge range A1:D1.
Set width of column A to 18.
Set height of row 1 to 30.
Fit column A.
Fit all columns.
Freeze top row.
Freeze the first 3 rows.
Make row 1 a header row.
Band every other row of A5:C9 "#EEF3F8".
```

### Colour

There are two ways to say a colour, and each sentence takes one of them.

**By word.** `Make cell C4 red.`, `Make range A1:C1 blue.` and the
`colored` border sentences below take one of eight plain words: `red`,
`yellow`, `black`, `blue`, `cyan`, `green`, `magenta`, `white`.

**By value.** `Set fill-color of …` and `Set font-color of …` take a code in
quotes, such as `"#1F3A5F"`, or a name you have defined for one:

```text
Define hot-pink as "#FF69B4".
Set font-color of cell A6 to hot-pink.
```

The `examples` folder has a palette ready to copy, `joy.txt`: the fifteen
paints and four bases of a well-known television painter, each as a
`Define`. There are no mistakes in it, only happy accidents.

> **Mind the fill.** `Make cell C4 red.` colours the cell's *background*.
> To colour the text, say so: `Set font-color of cell C4 to "#FF0000".`

### Numbers, dates and money

```text
Format range A1:A9 as dollars.
Format range A1:A9 as euros with 0 decimals.
Format range B5:B10 as accounting in dollars.
Format range C5:C10 as percent with 1 decimal.
Format range A1:A9 as a number with 0 decimals and thousands separators.
Format cell B2 as a long date.
Format range A1:A9 as a short date.
Format range A1:A9 as an ISO date.
Format range A1:A9 as a time.
Format range A1:A9 as text for new entries.
Format range A1:A9 as general.
Format range A1:A9 using "00000".
```

The currency is named in the sentence: `dollars`, `euros` or `pounds`.
Dates and times follow the settings of the computer that reads them.

### Borders

```text
Add a border around range B2:D4.
Add borders to every cell in range A1:C3.
Add a bottom border to range B5:D5.
Add a bottom border colored green to range B5:D5.
Remove borders from range B2:D4.
```

`around` draws a box. `to every cell in` draws a grid. The sentence says
which, because English hears "add a border" as a box and Excel is happy to
give you a grid.

## Lesson 9. Rows, columns, sheets and files

```text
Delete row 5.
Delete the third row.
Delete rows 4 through 9.
Insert a row.
Insert 3 rows at row 5.
Insert column before C.
Hide column C.
Unhide column C.
Unhide all rows and columns.
Move column B before column D.
Group rows 2 through 10.
Delete blank rows in A1:D50.
Remove trailing empty rows and columns.

Copy range A1:D10 to range F1:I10.
Copy range A1:D10 to sheet Archive.
Copy range A1:A5 to C1 transposed.
Copy formatting of A1:A5 to C1:C5.
Move range A1:A5 to C1.
Convert range B2:B9 to values.
Fill down range C2:C90.
Fill A1:A10 with a series starting at 1.

Clear range A1:C10.
Clear everything from A1:D50.
Clear formatting of range A1:D50.

Delete sheet Old.
Set tab-color of sheet Data to "#FF69B4".
Protect sheet Data with password "abc".
Save this workbook.
```

`Clear range` empties the cells and keeps their formatting. `Clear
everything from` removes both. The sentence names which, as before.

> **`Delete blank rows in` means more than it sounds.** It deletes every
> row that has an empty cell *anywhere in the range you name*, and it
> deletes the whole row of the sheet. Over `A1:D50`, a row with a name, a
> date, an amount and no receipt is a row with a blank in it. Name the one
> column that decides: `Delete blank rows in A2:A50.`

### Sentences that reach outside the workbook

```text
Set report-path to "C:\Reports\Weekly.pdf".
Export this sheet as pdf report-path.
Set backup-path to "C:\Reports\Backup.xlsx".
Save a copy as backup-path.
Open workbook "C:\Reports\file.xlsx".
Print this sheet.
Email "boss@co.com" with subject "Report" and message "Attached.".
```

Four things to know.

- **A file path goes in quotes.** File names contain periods, and a period
  ends a sentence.
- **`Email` never sends.** It opens a draft in your own Outlook, for you to
  read and send.
- **`Print this sheet` opens the print preview.**
- **Undo does not reach these.** A saved file stays saved.

If Windows has marked a workbook as having come from the internet, Frazaro
lets a program read and edit it but refuses these outward-reaching
sentences. To allow them, unblock the file in Windows: right-click it,
**Properties**, **Unblock**. Frazaro has no button of its own for this, on
purpose, so that a workbook arriving from outside cannot carry its own
permission slip.

## Lesson 10. Sorting, filtering, tables and pivots

### Sorting

```text
Sort range A1:C50 by column B with a header row.
Sort range A2:C50 by column B descending without a header row.
Sort range A1:C50 by column B descending then by column C ascending with a header row.
Sort this sheet by column B descending with a header row.
```

Every sort says whether its first row is a header. Leave it out and the
sentence is refused:

```text
Sort range A1:C50 by column B.
```

> … then I expected one of 'with'/'without' but found the end of the
> sentence.

Excel will guess whether your first row is a header. Frazaro will not,
because a wrong guess sorts your headings into the middle of your data and
reports success. This is a general rule of the language: **where Excel
offers two readings, the sentence names the one it wants.**

### Filtering

```text
Add filters to range A1:D50.
Filter range A1:D50 to show rows where column C is "West".
Filter range A1:D50 to show rows where column D is greater than 100.
Filter range A1:D50 to show rows where column C contains "we".
Copy only the visible cells of range A1:D50 to F1.
Clear the filter conditions.
Remove the filters.
Remove duplicates from range A1:C50 by column 2.
```

`Clear the filter conditions.` shows every row again and leaves the filter
buttons. `Remove the filters.` takes the buttons away too.

### Tables

```text
Turn A1:D10 into a table called Sales.
Set style of table Sales to TableStyleMedium2.
Show the total row of table Sales.
Add a row to table Sales.
Delete row 1 of table Sales.
Set qty-values to column Qty of table Sales.
Turn table Sales back into a range.
```

### Pivot tables

This is the heart of sample 05. It reads as a pivot table is described to a
new colleague.

```text
Make a pivot table from A1:G41 at J3 called RegionPivot.
Add rows of Region to pivot RegionPivot.
Add columns of Channel to pivot RegionPivot.
Add Revenue to pivot RegionPivot as a sum.
Show pivot RegionPivot in tabular form.
Sort Region in pivot RegionPivot descending by Revenue.
```

Several fields are written as a list: `Add rows of Product, Rep to pivot
ProductPivot.` Three or more take the comma before "and", which is not a
point of style here but how Frazaro tells where the list ends:
`rows of Region, Product, and Channel`.

Also available: `Refresh pivot …`, `Refresh every pivot table.`,
`Collapse … in pivot …`, `Expand …`, `Hide subtotals for … in pivot …`,
`Remove … from pivot …`, `Rename pivot … to …`,
`Change the source of pivot … to …`, `Delete pivot …`.

## Lesson 11. Tidying text

### Where it stands

```text
Make range A2:A50 upper case.
Make column B lowercase.
Capitalize each word in range A2:A50 after any space.
Capitalize each word in column B after any non-letter.
Remove extra spaces from column A.
Remove non-printing characters from range A2:A50.
Replace "N/A" with 0 in column C.
```

The case and spacing sentences change text and leave everything else alone:
a number stays a number, and a formula stays a formula.

`Capitalize` makes you choose, and the choice is visible in names.

| The text | `after any space` | `after any non-letter` |
|---|---|---|
| `don't stop` | Don't Stop | Don'T Stop |
| `3rd quarter` | 3rd Quarter | 3Rd Quarter |
| `o'neil` | O'neil | O'Neil |

Neither column is right for every row, which is why the sentence will not
choose for you.

`Replace` finds text anywhere in a cell, in any mix of capitals. From
`0.7.1` it changes the values you typed and never the inside of a formula;
in `0.7.0` it is Excel's own Find and Replace, which reaches into formulas
too, so keep it away from columns that hold them.

### In a name

```text
Set part to the text before "-" in code.
Set rest to the text after ": " in cell B2.
Set extension to the text after the last "." in file-name.
Set prefix to the first 3 characters of code.
Set code to id padded on the left with "0" to 5 characters.
Set names to range A2:A9 as one list separated by "; ".
Set tidy to raw-name with extra spaces removed.
Set cleaned to msg with "x" replaced by "y".
```

With `code` holding `INV-ab-Cd`:

| Sentence | Gives |
|---|---|
| `the text before "-" in code` | `INV` |
| `the text after "-" in code` | `ab-Cd` |
| `the text after the last "-" in code` | `Cd` |
| `the first 3 characters of code` | `INV` |
| `the last 2 characters of code` | `Cd` |
| `42 padded on the left with "0" to 5 characters` | `00042` |

If the marker is not in the text at all, the run stops at that sentence
rather than hand you something empty and let you carry on.

**From `0.7.1`**: splitting a column, and replacing inside formulas.

```text
Split column C by "," as text.
Split column A by " - " reading numbers.
Replace "Sheet1" with "Data" in the formulas of range B2:B20.
Replace "N/A" with 0 on this sheet.
```

A split must say `as text` or `reading numbers`, never neither, and never
writes over a cell that already holds something.

---

# Part III. Deciding and repeating

## Lesson 12. Deciding

### The short form

A condition, a comma, and the sentence to run.

```text
If grand is at least 45, make cell C4 yellow.
If receipt is "No", put "Missing receipt" into column F row r.
If cell B2 is empty, show "blank".
If code starts with "AB", show "ab".
```

### The long form

A condition, a colon, and the sentences under it. A blank line separates
one branch from the next.

```text
If grand is greater than 100:
  Set verdict to "huge".

Otherwise, if grand is greater than 40:
  Set verdict to "solid".

Otherwise:
  Set verdict to "small".
```

### Choosing among values

```text
When region is "North":
  Set region-label to "cold".

When it is "South" or "East":
  Set region-label to "warm".

Otherwise:
  Set region-label to "unknown".
```

The first `When` names the thing being examined. The ones after it say `it`.

### What a condition can say

| Comparing | Ways to say it |
|---|---|
| equal | `is` · `equals` · `is equal to` |
| not equal | `is not` · `does not equal` · `is not equal to` |
| more | `is greater than` · `is more than` |
| less | `is less than` |
| at least, at most | `is at least` · `is at most` · `is greater than or equal to` · `is less than or equal to` |
| empty | `is empty` · `is not empty` |
| text | `contains` · `does not contain` · `starts with` · `ends with` |
| whole multiples | `is divisible by` |

`is empty` counts a cell holding only spaces as empty. The text tests
ignore capitals.

Conditions join with `and` and `or`:

```text
If amount is greater than 500 and receipt is "No", show "Escalate".
```

> **Mixing `and` with `or`.** They are read strictly left to right, and
> there are no parentheses. `a or b and c` is read as *(a or b) and c*. When
> you mean something else, put one test inside another, or work the first
> part out under a name.

## Lesson 13. Repeating

```text
Repeat 5 times:
  Increase total by counter.

Count r from 2 to last-row:
  Set amount to cell in column D row r.
  If amount is greater than 500, put "Needs manager approval" into column F row r.

Count stripe-row from 1 to f-last step 2:
  Make cell in column F row stripe-row bold.

Count back-row down from f-last to 1 step 2:
  Log "back-row " joined with back-row.

While countdown is greater than 0:
  Log "countdown " joined with countdown.
  Decrease countdown by 1.

Repeat until fuel is 0:
  Decrease fuel by 1.
```

| Loop | Use it when |
|---|---|
| `Repeat N times:` | You want something done N times. The word `counter` runs from 1 to N inside it. |
| `Count r from A to B:` | The number matters, and you want to name it. This is the loop for working down the rows of a sheet. Add `step 2` to skip, `down` to go backwards. |
| `While …:` | You want to carry on as long as something is true. |
| `Repeat until …:` | You want to carry on until something becomes true. |
| `For each x in …:` | You have a list or a remembered range. Lesson 14. |

Short loops have a one-line form, with a comma, like the short `If`:

```text
Count echo-row from 1 to f-last, log "echo " joined with echo-row.
```

`Stop the loop.` leaves the loop you are in.

```text
Count search-row from 1 to 10:
  If search-row is 3, stop the loop.
```

### How blocks begin and end

This is the one piece of layout that matters, so it has its own heading.

1. **A sentence ending in a colon opens a block.**
2. **The sentences under it belong to it.** Indenting them is for the
   reader. Frazaro does not look at it.
3. **A blank line closes every open block.**
4. **`Done.` closes only the innermost block**, so that the outer one can
   carry on.
5. **Chains survive blank lines.** `Otherwise`, `When it is` and
   `If that fails` pick up the chain they belong to, even across the blank
   line that ended the branch before.

Here is rule 4 earning its keep. The `Repeat` sits inside the `If`. `Done.`
closes the `Repeat`, the `If` continues with one more sentence, and the
blank line then ends that branch.

```text
If grand is greater than 40:
  Put "PASS" into cell C4.
  Repeat 2 times:
    Log "pass check " joined with counter.
  Done.
  Make cell C4 bold.

Otherwise:
  Put "CHECK" into cell C4.
```

And here, from sample 07, a block inside a block where nothing needs to
follow the inner one. A single blank line closes both.

```text
Count r from 2 to last-row:
  Set on-hand to cell in column D row r.
  Set reorder-point to cell in column E row r.
  Put "OK" into column H row r.
  If on-hand is at most reorder-point:
    Put "REORDER" into column H row r.
    Make cell in column H row r bold.

Set reorder-count to count of reorder-list.
```

> **Two mistakes nearly everyone makes once.**
>
> *A blank line inside a block, for air.* It ends the block. Use a `#` note
> if you want a visual break; a note does not end anything.
>
> *A sentence directly under a block, at the left margin, meant to come
> after it.* With no blank line between, it is inside the block, whatever
> the indentation suggests. What ends a block is the blank line.

One more. Inside a `Repeat`, the word `counter` belongs to that `Repeat`.
When you need a loop within a loop, use `Count` for the inner one and give
its number a name of its own.

## Lesson 14. Lists and lookups

A **list** collects values as a program goes, when it cannot know in
advance how many there will be.

```text
Create a list called found-items.
Repeat 4 times:
  If counter is greater than 2, append counter times 100 to found-items.

For each f in found-items, log "found " joined with f.
Set list-count to count of found-items.
Set pick-check to item 2 of found-items.
```

After that, the list holds 300 and 400, `list-count` is 2, and
`pick-check` is 400.

A **lookup** stores values under keys: prices by part number, names by
staff code.

```text
Create a lookup called prices.
Store 100 at key "ax-7" in prices.
Store 250 under "bx-2" in prices.
Set ax-price to prices for "ax-7".
Set price-sum to prices for "ax-7" plus prices for "bx-2".
If prices for "bx-2" is greater than 200, set price-verdict to "steep".
```

| About lookups | |
|---|---|
| Keys ignore capitals. | `"AX-7"` finds `"ax-7"`. |
| Storing at a key that exists replaces what was there. | |
| Reading a key that was never stored stops the run and names the key. | |

Walking a lookup:

```text
For each k in keys of prices:
  Set key-list to key-list joined with k.

For each pair in prices:
  If value of pair is greater than 200, set pair-trace to pair-trace joined with key of pair.
```

To walk the cells of a range, remember the range first:

```text
Remember range B2:B4 as results.
For each r in results, log r.
```

## Lesson 15. Steps of your own

Say a thing once, give it a name, and use the name. Steps are defined with
`To`, at the top level of the program, before the sentences that use them.

### A step that does things

```text
To tidy-up:
  Fit column A.
  Fit column B.
  Set font-color of cell A1 to hot-pink.

Tidy-up.
```

### A step with details that vary

Each detail is given a usual value with `of`. A call may leave out any
detail it is content with.

```text
To stamp, with row-number of 1 and value of "ok":
  Put value into column F row row-number.

Stamp.
Stamp with row-number of 2 and value of "beta".
```

### A step that gives back a value

```text
To landed-cost of amount:
  Give back amount times 1.05.

Set landed-total to landed-cost of order-total.
```

Once defined, `landed-cost of …` is a value like any other, and can stand
inside a condition: `If tax of 50 is greater than 3: …`

### A value step with several named details

```text
To get commission using sale of 1000 and rate of 5%:
  Give back sale times rate.

Set full-commission to commission using sale of 2000 and rate of 10%.
Set default-commission to get commission using sale of 600.
```

`full-commission` is 200. `default-commission` is 30, because the rate was
left out and took its usual value of 5%. The word `get` is optional and
reads well.

### A value step with no details at all

```text
To get vat-rate:
  Give back 20%.
```

After that, `vat-rate` is simply a value.

### What validation checks for you

Every call is compared with its definition before anything runs. The
refusal quotes your sentence back, and then:

> … the action 'stamp' has no parameter called 'colour'. Its parameters
> are: row-number, value

A misspelled detail is caught at your desk and not during the month-end
close.

## Lesson 16. Talking to the person at the keyboard

```text
Show total.
Show "Expense check done: " joined with missing-receipts joined with " claims are missing a receipt".
Log "grand is " joined with grand.
Ask "Your name?" and put answer into who.
Put "working" in status bar.
Clear status bar.
Wait 2 seconds.
```

| Verb | Does | Stops the run to wait? |
|---|---|---|
| `Show` (or `Say`) | Opens a message box. | Yes, until OK is pressed. |
| `Log` | Writes a line to the run's Trace log (and VBA's Immediate pane), for whoever maintains the procedure. | No. |
| `Ask` | Opens a box with a question and keeps the answer under a name. | Yes. |
| `Put … in status bar` | Writes at the bottom of the Excel window, until the run ends. | No. |

To read what a program logged, run it with **Interpret and Trace**. The
trace is written to a sheet named *Trace*.

> **End the sentence outside the quotes.** A message that ends in a full
> stop still needs the sentence's own:
> `Show "All done.".` It looks odd and it is correct.

### Buttons

A program can put a button on a sheet and say what pressing it does.

```text
When "Refresh" is clicked:
  Refresh everything.
  Show "Refreshed".

Make a button "Refresh" at cell D3.
```

The `When … is clicked:` part is a definition, like a step, and stands at
the top level. Run the program once with **Interpret and Run**; from then
until the workbook is closed, the button does what the sentences say.

## Lesson 17. When something might go wrong

Some sentences can fail for honest reasons. The sheet is not there yet. The
part number is not in the price list. `Try` lets a procedure say what to do
about it.

```text
Try:
  Go to sheet Nowhere-Land.
  Set rescue to "unreachable".

If that fails:
  Log "the problem was " joined with the problem.
  Set rescue to "rescued".
```

The first sentence under `Try` that fails ends the attempt. The sentences
after it in the block do not run. The sentences under `If that fails:` run
instead, and inside them `the problem` holds what went wrong, in words.

`If that fails:` is optional. Without it, a failure inside the `Try` simply
moves on. Sample 05 uses this to delete last week's pivot tables, which do
not exist the first time it runs:

```text
Try:
  Delete pivot RegionPivot.
  Delete pivot ProductPivot.

If that fails:
  Log "No earlier pivots to remove".
```

### When there is no Try

A sentence that fails outside any `Try` stops the run. Frazaro then does
four things, in this order: it puts the sheets back as they were before the
run, marks the row it stopped on, goes to that row, and tells you.

> The run stopped at line 7:
> Go to sheet Nowhere-Land.
>
> *(what went wrong, in words)*
>
> Put back as they were before the run: Output.
> Anything it did anywhere else, like a file saved or an email drafted,
> stays as it is.

The last line is there every time. It is the honest edge of Undo.

---

# Part IV. The real thing

## Lesson 18. Reading a refusal

A refusal is Frazaro working as designed. It appears in column C beside the
row, on a red ground, and nothing has run.

### The anatomy

Most refusals have the same three parts.

> I understood 'put 1' - then I expected one of 'into'/'in' but found
> 'onto'. Did you mean: …

1. **How far the reading got.** `'put 1'` was understood.
2. **What was expected next, and what was found.**
3. **Sentence shapes that start the same way**, up to three of them.

When nothing at all was understood, the refusal says that, and sends you to
the list:

> Don't understand: '…' No loaded sentence starts with 'expense' - press
> What can I say? on the Frazaro tab to see all … sentences this program
> understands.

### One at a time

**Validate stops at the first sentence it cannot read.** Every row above it
turns green, that row turns red, and the rows below wait. Fix it, validate
again, and the next one, if there is one, will show itself.

### A field guide

| You wrote | Frazaro says | Write instead |
|---|---|---|
| `Put 1 onto cell A2.` | … expected one of 'into'/'in' but found 'onto'. | `into` |
| `Sort range banana by column B1.` | … expected a range (like A1:C50, B2, or Data!A1:B10 - quotes for a named range) but found 'banana'. | A range, or the named range in quotes. |
| `Fit column banana.` | … expected a column letter (like C or AA) but found 'banana'. | `Fit column C.` |
| `Set total to 5 + 3.` | I don't understand the character '+' - write 'plus' for addition | `5 plus 3` |
| `Set price to $5.` | I don't understand the character '$' - write the plain number (or 'Format ... as currency.' for display) | `Set price to 5.` |
| `Set label to don't.` | I don't understand the character ''' … | `Set label to "don't".` |
| `Sort range A1:C50 by column B.` | … expected one of 'with'/'without' but found the end of the sentence. | `… with a header row.` |
| `Capitalize each word in range A2:A50.` | … then I expected 'after' but found the end of the sentence. | `… after any space.` |
| `Filter range A1:D50 where column C is "West".` | … then I expected 'to' … | `Filter range A1:D50 to show rows where …` |
| `Count seek from 1 to 10:` | 'seek' is a reserved word in Excel's programming language … | `seek-row` |
| `Set A1 to 5.` | 'a1' is shaped like a cell, so it cannot be a name. To mean cell A1, write the word cell in front of it … | `Put 5 into cell A1.` |
| `Set total to B2 plus 1.` | 'b2' is shaped like a cell, so it cannot be a name … | `Set total to cell B2 plus 1.` |
| `Give back total.` outside a step | 'Give back' only makes sense inside a value-returning action … | `Show total.`, or move it into a step. |
| `Stop the loop.` outside a loop | 'Stop the loop.' only makes sense inside a loop … | `Stop.` ends the program. |
| `Create a button called "Go" at cell D3.` | a button isn't created with 'Create ... called ...' - use 'Make a button "<caption>" at cell <cell>.' instead | `Make a button "Go" at cell D3.` |
| `If that fails:` with no `Try` before it | 'If that fails:' must start its own paragraph right after the Try block it rescues … | Put a `Try:` block above it. |

### When the refusal is wrong

Sometimes you wrote a perfectly reasonable sentence and Frazaro has not
learned it yet. That is not your mistake. Every sentence Frazaro could not
read is kept in a log inside your own workbook. **Copy Diagnostic Report**
puts that log on the clipboard: the version, the sentences it could not
read, and what it said about each. Paste it into an email to the address in
[SUPPORT.md](../SUPPORT.md). Sentences people actually reached for are how
the language learns new ones.

The report is copied to your clipboard and goes nowhere until you paste it.
Frazaro sends nothing by itself.

## Lesson 19. Your SOP, as written

Real procedures are not written for a computer. They have titles,
"Performed by" lines, and steps like "Send the total to payroll". You do not
have to rewrite them.

### Two lines

Put `<Frazaro>` on a line of its own above the part Frazaro should run, and
`</Frazaro>` on a line of its own below it. Everything outside stays exactly
as a person wrote it.

| The rule | |
|---|---|
| Each tag sits alone on its line. | Capitals do not matter. |
| A document may have several tagged sections. | They are read in order. |
| The last section may be left open. | It then runs to the end of the document. |
| A document with no tag at all is read from top to bottom. | Every line that is not a `#` note is then an instruction. |

### What the opening tag can carry

```text
<Frazaro week="Week 39" (steps 1 and 2, automated by Accounts Payable) 0.7.0>
```

| In the tag | It is | Example |
|---|---|---|
| `name=value` | A value the section can use as `{name}`. | `week="Week 39"`, `year=2026` |
| `(anything in parentheses)` | A note for people. Frazaro does not read it. | `(approved by the Controller)` |
| a version | This section needs that Frazaro or later. | `0.7.0` |
| a language | This section is written for that phrasebook. | `espanol` |

A value is filled in as the document loads, so the sheet shows the finished
sentence. Inside quoted text `{week}` becomes the text. Standing alone in a
sentence it becomes a value: `Put {year} into cell B1.` puts the number
2026. To write a brace itself, type it twice.

A value with a space in it needs its quotes: `week="Week 39"`. A single
word or number does not: `year=2026`.

If a tag asks for a newer Frazaro than the one installed, nothing loads:

> … this section needs Frazaro 0.7.1 or later, and this copy is Frazaro
> 0.7.0. Update Frazaro, then load the file again. Nothing was loaded, and
> the program on this sheet has not been changed.

### The practice SOP

Load **`Practice - Expense Reimbursement (as written).docx`**. It has no
tags, so every line is read as an instruction, and the prose lines are
refused. For each red row you have two repairs.

- **Put `#` in front of it.** It becomes a note: kept in the procedure,
  never run.
- **Rewrite it as an instruction.**

Or do neither: open the document, put the two tags around the
instructions, save, and press **Reload Instructions**. The red row itself
suggests this.

### Files

| File | |
|---|---|
| Word (`.docx`, `.doc`) | Read through Word itself, in a hidden copy that Frazaro starts and closes, with the document's macros switched off. Each paragraph becomes one row. Word's curly quotes and long dashes are straightened on the way in. Save the document before you reload it; unsaved changes are not seen. |
| Plain text (`.txt`) | Works everywhere and needs nothing extra. |
| PDF | Refused by name. A PDF stores marks on a page, not lines and paragraphs, and a converter has to guess where the blank lines were. A lost blank line changes what a procedure does without changing a word of it. Open the PDF in Word, save as `.docx`, **read the result**, then load that. |
| Pictures and scans | Not read. Where a step was a screenshot, write a `#` note saying what it showed. |

The daily loop, once a procedure lives in a file: edit, save,
**Reload Instructions**, **Interpret and Run**.

## Lesson 20. Questions of your tables

Some questions are awkward as formulas and easy as sentences. *Who reports
to Alice, directly or not? Which shifts can nobody cover?*

Frazaro answers these with worksheet functions that read your Excel Tables
and spill their answers as live formulas. You do not have to write the
functions. There are sentences that write them for you.

```text
Show every Name in Staffing with Salary over 80000 as Rich in cell E2.
Show everyone who reports to "Alice" directly or not in Reports as team in cell E2.
```

The first puts a formula in E2 that lists every Name in the *Staffing*
table whose Salary is over 80,000. `as Rich` gives the answer a name of its
own, which the formula uses for the rule it writes. The second follows a
chain of reports as far as it goes. Both answers are live: change the table
and they change with it.

### A policy, written as sentences

Sample 08 goes further. The staffing policy itself is written in sentences,
one rule to a cell, and the questions are asked of the rules.

```text
Write in cell M2 that a person can-cover a shift if Shifts lists the shift as Shift, the cert as Needs, and the min as MinLevel, and Staff lists the person as Name, the cert as Cert, and the level as Level, and the level is at least the min, and not Leave lists the person as Name and the shift as Shift.
Write in cell M3 that a shift is listed if Shifts lists the shift as Shift.
Write in cell M4 that a shift is covered if the person can-cover the shift.

Show in cell A4 who can-cover "Night" by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
Show in cell F4 which shift that is listed is not covered by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
Show in cell H4 whether every shift that is listed is covered by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
```

Read the first rule slowly. `a person` and `a shift` are what the rule is
about. `Shifts lists the shift as Shift` says that the *Shifts* table has a
column headed *Shift*, and that is where the shift is found. `the level is
at least the min` compares two things the tables supplied. `not Leave lists
…` rules out anyone on leave for that shift.

| A question beginning | Answers with |
|---|---|
| `who …`, `what …`, `which person …` | a list |
| `whether …` | TRUE or FALSE |
| `how many …` | a number |
| `… as one list` | everything in one cell, separated by commas |

Two things to get right.

- **Text you mean exactly goes in quotes**: `"Night"`, `"Alice"`. A bare
  word is folded to small letters on its way through, and a table's own
  text keeps its capitals, so `night` would not find `Night`.
- **A number is written bare.** `80000`, never `"80000"`.

Excel needs to be a version with dynamic arrays (Microsoft 365, or Excel
2021 and later) for an answer to spill. [Level 4](4-wizard.md) explains the
engines underneath.

## Lesson 21. Running for real

### Two ways to run

| | **Interpret and Run** | **Compile and Run** |
|---|---|---|
| What it does | Performs each sentence directly on the workbook. | Writes the program out as a VBA module in the workbook, then runs that. |
| Writes code into your file | No | Yes |
| Needs an Excel security setting | No | Yes: *Trust access to the VBA project object model*. It refuses in words when the setting is off. |
| Works on Excel for Mac | Yes | No |
| Use it | Every day. | When an auditor wants the artefact, or a workbook has to run where Frazaro is not installed. |

The two are held to the same answers. The project runs one body of programs
through both, compares the resulting sheets, and fails a test if they
differ.

**Show me the VBA** writes the VBA that Compile would produce onto a sheet
named *Generated VBA*, without changing anything else. Each line is tagged
with the sentence it came from. It needs no security setting, so anyone may
look.

### Seeing what happened

**Interpret and Trace** and **Compile and Trace** run the program and then
show a transcript on a sheet named *Trace*, which you can copy into an
email.

### What Undo covers

| Undo puts back | Undo does not reach |
|---|---|
| The Output sheet, and every sheet the program names | A workbook saved or exported to disk |
| Sheets the run deleted | An email draft it opened |
| It removes sheets the run created, the sheet it worked on included | Another workbook the program opened |

It covers the most recent run of that program, and each program in a
workbook has its own.

Excel's own settings are not Undo's business, because they never outlast
a run. Whatever a program did to automatic calculation, the status bar or
screen updating, the run puts back as it found it the moment it ends,
finished or stopped. To leave Excel in manual calculation, use Excel's
Formulas tab, or type `Turn off automatic calculation.` into **Open CLI**.

### Several programs in one workbook

**New Named Frazaro** asks for a short name and makes a sheet called
*Frazaro (name)*. Each has its own sentences, its own verdicts and its own
Undo. A button acts on the program whose sheet you are looking at.

### If a run will not stop

A `While` whose condition never becomes false will run until you stop it.
**Ctrl+Break** stops a running program.

---

# Part V. Practice

## Exercises

Use `Frazaro Sample Data.xlsx` and a new *Frazaro* sheet for each. Validate
before you run. The answers follow, and any answer that validates and does
the job is a right answer.

**1. A headline.** On a sheet called *Scratch*, put the words *Monthly
Report* into A1, make them bold at size 18, and put today's date into B1
shown as a long date.

**2. A total.** On the *Expenses* sheet, write the sum of every amount in
D2:D26 into cell I12, shown as dollars, with the label *All claims* beside
it in H12.

**3. A condition.** Extend exercise 2. If the total is over 7,000, show the
message *Over budget*; otherwise show *Within budget*.

**4. A loop.** On the *Expenses* sheet, for every claim from row 2 to the
last filled row, put *Large* into column G when the amount in column D is
greater than 500.

**5. A chain.** Start a value at 200. Grow it by 10%, increase it by 50%,
double it by adding 100 percent to it, and decrease it by 75%. Put the
result into cell A1. What is it?

**6. A step.** Define a step that gives back 8% of whatever amount it is
given. Use it to put the tax on 100 into cell B1.

**7. A rescue.** Working on a sheet called *Scratch*, try to go to a sheet
called *Archive-2019*. If that fails, come back and put the words
*No archive* into cell A1.

**8. A repair.** Each of these is refused. Say why, and mend it.

```text
Set total to 4 * 25.
Sort range A1:G41 by column G descending.
Put "Done" on cell A1.
Set next to 5.
```

**9. An SOP.** Take a procedure of your own, a real one. Put the two tags
around the three steps you most dislike doing by hand, write those steps as
sentences, and get every row green. You need not run it today.

## Answers

**1.**

```text
Work on sheet Scratch.
Put "Monthly Report" into cell A1.
Make cell A1 bold.
Set font size of cell A1 to 18.
Put today into cell B1.
Format cell B1 as a long date.
```

**2.**

```text
Work on sheet "Expenses".
Set all-claims to sum of range D2:D26.
Put "All claims" into cell H12.
Put all-claims into cell I12.
Format cell I12 as dollars.
```

The total is 7,060.84.

**3.** Add:

```text
If all-claims is greater than 7,000:
  Show "Over budget".

Otherwise:
  Show "Within budget".
```

It is over budget.

**4.**

```text
Work on sheet "Expenses".
Set last-row to last filled row of column A.
Count r from 2 to last-row:
  Set amount to cell in column D row r.
  If amount is greater than 500, put "Large" into column G row r.
```

Six claims are marked.

**5.**

```text
Create a number called growth-check.
Set growth-check to 200.
Grow growth-check by 10%.
Increase growth-check by 50%.
Add 100 percent to growth-check.
Decrease growth-check by 75%.
Put growth-check into cell A1.
```

165. That is 200 × 1.1 × 1.5 × 2 × 0.25.

**6.**

```text
To tax of amount:
  Give back amount times 0.08.

Put tax of 100 into cell B1.
```

B1 holds 8.

**7.**

```text
Work on sheet Scratch.
Try:
  Go to sheet Archive-2019.

If that fails:
  Go to sheet Scratch.
  Put "No archive" into cell A1.
```

**8.**

| Refused | Because | Mended |
|---|---|---|
| `Set total to 4 * 25.` | Arithmetic is written in words. | `Set total to 4 times 25.` |
| `Sort range A1:G41 by column G descending.` | A sort says whether it has a header row. | `Sort range A1:G41 by column G descending with a header row.` |
| `Put "Done" on cell A1.` | `Put` takes `into` or `in`. | `Put "Done" into cell A1.` |
| `Set next to 5.` | `next` is a reserved word. | `Set next-one to 5.` |

**9.** There is no answer in the back of the book for this one. If a
sentence you needed is missing, **Copy Diagnostic Report** and send it.

## What an Apprentice can do

You have finished this level when you can do each of these without looking
anything up.

- [ ] Load a procedure from Word or a text file, and reload it after an
      edit.
- [ ] Validate, run, and undo.
- [ ] Say where a program works, and what happens if it never says.
- [ ] Put a value in a cell and give a value a name, and say which verb
      does which.
- [ ] Write a condition, in the short form and the long.
- [ ] Walk down the rows of a sheet whose length you do not know.
- [ ] End a block on purpose, and end an inner block without ending the
      outer one.
- [ ] Define a step and use it.
- [ ] Guard a sentence that might fail.
- [ ] Read a refusal, find the word it stopped at, and mend the sentence.
- [ ] Put two tags around the working part of a real SOP.
- [ ] Find a sentence you do not know with **What can I say?**, and report
      one that does not exist.

## The one-page card

```text
LAYOUT
  One sentence per row, ending with a period.
  # starts a note.          A colon opens a block.
  A blank line closes every open block.    Done. closes one.
  <Frazaro> ... </Frazaro> fences the part that runs.

REFERENCES          cell B2   range A1:C50   column C   sheet Data
                    sheet "Q1 Data"   cell Data!B2   cell in column D row r
TEXT                "in double quotes"      a quote inside: ""
NAMES               one word, hyphens welcome: last-row, order-total
NUMBERS             1,000,000   3.14   15%   10 percent   the third row

VALUES              Put 5 into cell B2.          Set total to 0.
                    Remember range B2:B4 as results.
                    Define hot-pink as "#FF69B4".
                    Increase total by 5.         Grow total by 10%.
ARITHMETIC          plus  minus  times  divided by  joined with
                    No parentheses: name the inner part first.
READING             sum of   average of   count of   largest of   smallest of
                    last filled row of column A
                    lookup of code in range A2:C90 column 3

DECIDING            If total is greater than 5, show "big".
                    If ...:   Otherwise, if ...:   Otherwise:
                    When region is "North":   When it is "South" or "East":
CONDITIONS          is  is not  is greater than  is less than  is at least
                    is at most  is empty  contains  starts with  ends with
REPEATING           Repeat 5 times:            (counter runs 1 to 5)
                    Count r from 2 to last-row:
                    While ...:     Repeat until ...:     For each x in list:
                    Stop the loop.
STEPS               To tidy-up:                    Tidy-up.
                    To tax of amount:  Give back amount times 0.08.
TALKING             Show ...   Log ...   Ask "..." and put answer into who.
TROUBLE             Try:       If that fails:       the problem

BUTTONS             Load Instructions  ->  Validate Instructions
                    ->  Interpret and Run  ->  Undo Last Run
LOST?               What can I say?
```

---

Previous: [Level 0 · Introduction](0-introduction.md) ·
Next: [Level 2 · Journeyman: A Grammarian's Guide to Sentence Templates](2-journeyman.md)
