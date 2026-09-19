# Frazaro examples

**Eight real procedures, written in plain English, that run in Excel.** They
start with a five-sentence cleanup and end with a staffing policy you can
question. Each one is a standard operating procedure (SOP) of the kind people
already keep in Word: a sales summary, an expense audit, receivables aging, a
month-end close, a purchase order, a shift schedule. The difference is that Frazaro
can run these ones.

All of them work on one workbook, **`Frazaro Sample Data.xlsx`**, so there is
nothing to set up.

---

## Run your first one in two minutes

1. Open **`Frazaro Sample Data.xlsx`** with Frazaro installed.
2. On the **Frazaro** tab, press **Load Instructions** and pick
   **`01 Tidy the Sales Export.txt`**. The procedure appears on a new
   *Frazaro* sheet, one line per row.
3. Press **Validate Instructions**. Every line gets a green OK, or a note in
   plain words saying what to change. Nothing has run yet.
4. Press **Interpret and Run**, then look at the *Sales* sheet.

Changed your mind? **Undo Last Run** puts the workbook back the way it was.

This is the whole of sample 01:

```text
# 01 - Tidy the Sales Export
# Area: Sales. Works on the Sales sheet of Frazaro Sample Data.xlsx.

Work on sheet "Sales".
Make row 1 a header row.
Format range G2:G41 as dollars.
Fit all columns.
Freeze top row.
```

---

## The samples, easiest first

| # | Sample | Area | What it shows you |
|---|---|---|---|
| 01 | [Tidy the Sales Export](01%20Tidy%20the%20Sales%20Export.txt) | Sales | Your first program: five sentences that make a raw export readable. |
| 02 | [Weekly Sales Summary](02%20Weekly%20Sales%20Summary.docx) | Sales | Totals, totals by region, and a report built on its own sheet. |
| 03 | [Expense Report Audit](03%20Expense%20Report%20Audit.txt) | Finance | Checking every row with a loop, and a message when it's done. |
| 04 | [Accounts Receivable Aging](04%20Accounts%20Receivable%20Aging.docx) | Accounting | Aging buckets, sorting, percent of total, and a warning when policy is breached. |
| 05 | [Sales Pivot by Region](05%20Sales%20Pivot%20by%20Region.docx) | Sales | Pivot tables one plain sentence at a time, and re-running safely with `Try`. |
| 06 | [Month-End Close](06%20Month-End%20Close.docx) | Accounting | Manual steps kept as notes, a balance check with two outcomes, and locking the ledger. |
| 07 | [Inventory Reorder](07%20Inventory%20Reorder.docx) | Operations | A step of your own, lists, lookups by SKU, and a purchase order built line by line. |
| 08 | [Shift Coverage](08%20Shift%20Coverage.docx) | Operations | A staffing policy written as sentences, then questions whose answers update themselves. |
| + | [Practice: Expense Reimbursement (as written)](Practice%20-%20Expense%20Reimbursement%20(as%20written).docx) | Finance | Loading an SOP the way people really write them, and fixing it line by line. |

Each sample stands alone, so you can run them in any order. Sample 06, the
month-end close, shows the idea best:

```text
# Before you run this - three steps done by hand
# 1. Export the month's trial balance from the accounting system.
# 2. Paste it over the Ledger sheet, keeping the column headers in row 1.

Work on sheet "Ledger".
Replace "N/A" with 0 in range D2:E21.
Set debits to sum of range D2:D21.
Set credits to sum of range E2:E21.
Set difference to debits minus credits.
...
If unreconciled is 0, protect sheet "Ledger" with password "close".
```

The steps a person does by hand stay in the procedure as notes, and the
steps a computer should do are sentences it can check. It is all one
document, and anyone on the team can read it.

---

## How to read a procedure

- **A line that starts with `#` is a note for people.** Frazaro skips it. Use
  notes for headings, reasons, and the steps someone does by hand.
- **Every other line is an instruction.** It is checked before anything runs,
  and it means exactly one thing. If Frazaro can't read a line, it says so in
  words and nothing runs.
- **A line ending in `:` starts a block** (`Count r from 2 to last-row:`,
  `If difference is 0:`). The lines under it belong to it. **A blank line ends
  the block.** Indenting is only there to make it easier to read.
- **Quotes mean "this exact text":** `"Sales"`, `"Over 90 days"`, `"Night"`.
- **Names with hyphens are your own words for a value:** `total-revenue`,
  `last-row`, `order-total`. Give a value a name with `Set`, then use it
  anywhere.

---

## The practice SOP

Real procedures aren't written for a computer. They have titles, "Performed
by" lines and steps like "Send the total to payroll". Load
**`Practice - Expense Reimbursement (as written).docx`** and press **Validate
Instructions**. Frazaro flags the lines it can't run. For each one you have
two fixes:

- **Put `#` in front of it.** It becomes a note: kept in the procedure, and
  never run.
- **Rewrite it as an instruction.** For example, "Add up every claim that has a
  receipt" is already written just below it as
  `Set receipted-total to sum of range D2:D26 where range E2:E26 matches "Yes".`

When every line is green, run it. That's the whole process for your own SOPs
too.

---

## Bonus: Bob Ross's palette

**`joy.txt`** (or **`joy.docx`**) holds the fifteen paints and four bases from
*The Joy of Painting* as named colours, a small easter egg. Load it and press **Interpret and Run** to paint a
swatch of each one on a *Joy* sheet.

To use the palette in your own procedure, copy its `Define` lines to the top.
After that, a colour is just a word:

```text
Define phthalo-blue as "#0C0040".
Define cadmium-yellow as "#FFEC00".

Set fill-color of range A1:D1 to phthalo-blue.
Set font-color of range A1:D1 to cadmium-yellow.
```

`Define` works for any colour you like, not just his. Any `"#RRGGBB"` code can
have a name.

---

## Bringing your own SOP

- **Word (`.docx`) works now.** Frazaro reads the file through Word itself,
  with the document's own macros switched off. Each paragraph becomes one row.
  You need Word installed.
- **Plain text (`.txt`) works everywhere** and needs nothing extra.
- **PDF isn't read directly yet.** For now, open the PDF in Word (Word converts
  it), save it as `.docx`, and load that. A scanned PDF has no text for Word to
  recover.
- **Pictures and screenshots aren't read.** Where a step was a screenshot,
  write a `#` note saying what it showed, so the step doesn't quietly
  disappear.
- **Not sure how to say something?** Press **What can I say?** on the
  Frazaro tab to see the sentences Frazaro knows.

---

## For contributors

The `.docx` samples and the sample-data workbook are generated from
[`tools/build_examples.ps1`](../tools/build_examples.ps1), which holds every
SOP's text and every row of data. Git can't show changes inside a `.docx` or
`.xlsx`, so edit the script, then rebuild:

```text
powershell -File tools\build_examples.ps1
```

The rebuild is byte-for-byte identical when nothing changed. It uses no Office
automation, so it runs on any Windows machine with PowerShell 5.1. Every
instruction in the samples uses a sentence form the phrasebook proves when it
loads (`scripts/polyglotta/english.vla`) or the regression corpus runs
(`scripts/instructions.txt`).
