<#
build_examples.ps1 - LE.6's generator: the onboarding samples in examples/.

WHAT IT BUILDS, all into examples/:
  - Frazaro Sample Data.xlsx   one workbook of realistic data (a sales
                               export, an expense report, open invoices, a
                               trial balance, stock levels, and three Excel
                               Tables for a warehouse schedule), plus a Start
                               Here sheet. Every sample SOP works on it.
  - NN <Title>.txt / .docx     eight sample SOPs, easiest first, and one
                               practice SOP written the way a real one is.

WHY A SCRIPT AND NOT HAND-SAVED FILES: a .docx or .xlsx is a zip, so git can
show nothing about a change to one. The SOP text below IS the source; the
binaries are build output and are rebuilt byte-for-byte identically (fixed
zip timestamps, no creation dates), so an unchanged SOP never shows as a
changed file. Edit here, rerun, commit both.

NO OFFICE AUTOMATION. Both formats are written as their raw OOXML parts,
so this runs on any Windows box with PowerShell 5.1 and never launches Word
or Excel. PDFs are the one thing it cannot make; see examples/README.md.

THE SOP MARKUP, one line per program line:
  = text      the title        -> "# text", styled as a title in Word
  ~ text      the subtitle     -> "# text", styled as a subtitle
  ## text     a section        -> "# text", styled as a heading
  # text      a note           -> "# text", styled as a quiet note
  ? text      prose, unmarked  -> "text" with no #, styled as ordinary
                                  prose (the practice SOP only: these are
                                  the lines Frazaro is meant to refuse)
  (blank)     a blank line     -> ends every open block, as in any program
  text        an instruction   -> checked and run by Frazaro; leading
                                  spaces mark a block body (indentation is
                                  cosmetic - blank lines and Done. close
                                  blocks - so Word renders it as a margin)
Keep instruction lines ASCII: Word's own typography is undone on import,
but a .txt is read as-is.

Every instruction shape used below is one the phrasebook proves at load
(a test-success sentence in scripts/polyglotta/english.vla) or one the
regression corpus (scripts/instructions.txt) runs.

House style (tools/*.ps1): PowerShell 5.1, host-independent.
Usage:  powershell -File tools\build_examples.ps1
#>
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
$root = Split-Path -Parent $PSScriptRoot
$out  = Join-Path $root 'examples'
if (-not (Test-Path $out)) { New-Item -ItemType Directory -Path $out | Out-Null }
$inv  = [System.Globalization.CultureInfo]::InvariantCulture
$utf8 = New-Object System.Text.UTF8Encoding($false)
$ZipStamp = New-Object DateTimeOffset(2026, 9, 18, 0, 0, 0, [TimeSpan]::Zero)
$DataBook = 'Frazaro Sample Data.xlsx'

function Esc([string]$s) { [System.Security.SecurityElement]::Escape($s) }

function Write-Zip([string]$path, $parts) {
    if (Test-Path $path) { Remove-Item $path -Force }
    $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::CreateNew)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $parts.Keys) {
                $e = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                $e.LastWriteTime = $ZipStamp
                $bytes = $utf8.GetBytes([string]$parts[$name])
                $s = $e.Open(); $s.Write($bytes, 0, $bytes.Length); $s.Dispose()
            }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
}

$XmlHead = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n"

# =====================================================================
#  THE SOPs
# =====================================================================
$sops = @(
@{ Id = '01'; Title = 'Tidy the Sales Export'; Area = 'Sales'; Format = 'txt'
   Uses = 'Sales'; Learn = 'Your first program: five sentences that make a raw export readable.'
   Body = @'
= 01 - Tidy the Sales Export
~ Area: Sales. Works on the Sales sheet of Frazaro Sample Data.xlsx.
#
# Every Monday the sales system exports last week's orders as a plain,
# unformatted sheet. These five sentences make it readable.
#
# Lines that start with # are notes for people. Frazaro skips them.
# Every other line is an instruction, checked before anything runs.

Work on sheet "Sales".
Make row 1 a header row.
Format range G2:G41 as dollars.
Fit all columns.
Freeze top row.
'@ },

@{ Id = '02'; Title = 'Weekly Sales Summary'; Area = 'Sales'; Format = 'docx'
   Uses = 'Sales -> Weekly Summary'; Learn = 'Totals, totals by region, and building a report on its own sheet.'
   Body = @'
= Weekly Sales Summary
~ Sample 02 - Sales. Reads the Sales sheet, then builds a Weekly Summary sheet.
# Runs every Monday once the weekly export is in. Owner: Sales Operations.

## Step 1 - Read this week's numbers
# "Remember" gives a column of numbers a name; "Set" gives one number a name.
# Either name can be used in any step below.
Work on sheet "Sales".
Remember range G2:G41 as revenues.
Remember range F2:F41 as units-sold.
Set total-revenue to sum of revenues.
Set total-units to sum of units-sold.
Set orders to count of revenues.
Set average-order to average of revenues.
Set largest-order to largest of revenues.

## Step 2 - Split revenue by region
# "where ... matches" adds up only the rows for that region.
Set north-revenue to sum of range G2:G41 where range B2:B41 matches "North".
Set south-revenue to sum of range G2:G41 where range B2:B41 matches "South".
Set east-revenue to sum of range G2:G41 where range B2:B41 matches "East".
Set west-revenue to sum of range G2:G41 where range B2:B41 matches "West".

## Step 3 - Start the summary sheet
# Frazaro creates the sheet the first week and reuses it every week after.
Work on sheet "Weekly Summary".
Clear everything from A1:D20.
Put "Weekly Sales Summary" into cell A1.
Set font size of cell A1 to 18.
Make cell A1 bold.
Put "Prepared on" into cell A2.
Put today into cell B2.
Format cell B2 as a long date.

## Step 4 - The headline numbers
Put "Revenue" into cell A4.
Put total-revenue into cell B4.
Put "Orders" into cell A5.
Put orders into cell B5.
Put "Units sold" into cell A6.
Put total-units into cell B6.
Put "Average order" into cell A7.
Put average-order into cell B7.
Put "Largest order" into cell A8.
Put largest-order into cell B8.
Format cell B4 as dollars with 0 decimals.
Format range B5:B6 as a number with thousands separators.
Format range B7:B8 as dollars.
Make range A4:A8 bold.

## Step 5 - Revenue by region, largest first
Put "Region" into cell A10.
Put "Revenue" into cell B10.
Put "Share" into cell C10.
Put "North" into cell A11.
Put north-revenue into cell B11.
Put north-revenue divided by total-revenue into cell C11.
Put "South" into cell A12.
Put south-revenue into cell B12.
Put south-revenue divided by total-revenue into cell C12.
Put "East" into cell A13.
Put east-revenue into cell B13.
Put east-revenue divided by total-revenue into cell C13.
Put "West" into cell A14.
Put west-revenue into cell B14.
Put west-revenue divided by total-revenue into cell C14.
Sort range A10:C14 by column B descending with a header row.
Format range B11:B14 as dollars with 0 decimals.
Format range C11:C14 as percent with 1 decimal.
Set fill-color of range A10:C10 to "#1F3A5F".
Set font-color of range A10:C10 to "#FFFFFF".
Make range A10:C10 bold.
Band every other row of A11:C14 "#EEF3F8".
Set width of column A to 22.
Fit column B.
Set width of column C to 10.

# Done. The Weekly Summary sheet is ready to paste into Monday's email.
'@ },

@{ Id = '03'; Title = 'Expense Report Audit'; Area = 'Finance'; Format = 'txt'
   Uses = 'Expenses'; Learn = 'Checking every row with a loop, and a message when it is done.'
   Body = @'
= 03 - Expense Report Audit
~ Area: Finance. Works on the Expenses sheet of Frazaro Sample Data.xlsx.
#
# Policy: every claim needs a receipt, and any claim over $500 needs a
# manager's approval. This checks each claim, marks it in column F,
# totals spend by category, and tells you what it found.
#
# New in this one: "Count r from 2 to last-row" repeats the indented
# lines once for every claim, with r standing for the row number.

Work on sheet "Expenses".
Put "Check" into cell F1.
Set last-row to last filled row of column A.

# Check every claim. A later line overwrites an earlier one, so the
# most serious finding is the one that stays.
Count r from 2 to last-row:
  Set amount to cell in column D row r.
  Set receipt to cell in column E row r.
  Put "OK" into column F row r.
  If amount is greater than 500, put "Needs manager approval" into column F row r.
  If receipt is "No", put "Missing receipt" into column F row r.
  If receipt is "No", make cell in column B row r bold.

# Count what we found.
Set missing-receipts to count of range F2:F26 matching "Missing receipt".
Set need-approval to count of range F2:F26 matching "Needs manager approval".

# Spend by category, beside the claims.
Put "Category" into cell H1.
Put "Spend" into cell I1.
Put "Travel" into cell H2.
Set travel-spend to sum of range D2:D26 where range C2:C26 matches "Travel".
Put travel-spend into cell I2.
Put "Lodging" into cell H3.
Set lodging-spend to sum of range D2:D26 where range C2:C26 matches "Lodging".
Put lodging-spend into cell I3.
Put "Meals" into cell H4.
Set meals-spend to sum of range D2:D26 where range C2:C26 matches "Meals".
Put meals-spend into cell I4.
Put "Supplies" into cell H5.
Set supplies-spend to sum of range D2:D26 where range C2:C26 matches "Supplies".
Put supplies-spend into cell I5.
Put "Software" into cell H6.
Set software-spend to sum of range D2:D26 where range C2:C26 matches "Software".
Put software-spend into cell I6.
Put "Total" into cell H7.
Put sum of range I2:I6 into cell I7.

# Make it easy on the eyes.
Make row 1 a header row.
Make range H7:I7 bold.
Format range D2:D26 as dollars.
Format range I2:I7 as dollars.
Fit all columns.

Show "Expense check done: " joined with missing-receipts joined with " claims are missing a receipt, and " joined with need-approval joined with " need a manager's approval".
'@ },

@{ Id = '04'; Title = 'Accounts Receivable Aging'; Area = 'Accounting'; Format = 'docx'
   Uses = 'Invoices -> AR Aging'; Learn = 'Aging buckets, sorting, percent of total, and a warning when policy is breached.'
   Body = @'
= Accounts Receivable Aging
~ Sample 04 - Accounting. Reads the Invoices sheet, then builds an AR Aging sheet.
# Run on the first business day of the month, once the invoice export is in.

## Step 1 - Put every open invoice in an aging bucket
# Days Overdue (column E) comes straight from the billing system.
# Each check overwrites the one before, so an invoice lands in the oldest bucket it reaches.
Work on sheet "Invoices".
Put "Bucket" into cell F1.
Set last-row to last filled row of column A.
Count r from 2 to last-row:
  Set days to cell in column E row r.
  Put "Current" into column F row r.
  If days is greater than 0, put "1-30 days" into column F row r.
  If days is greater than 30, put "31-60 days" into column F row r.
  If days is greater than 60, put "61-90 days" into column F row r.
  If days is greater than 90, put "Over 90 days" into column F row r.
  If days is greater than 90, make cell in column B row r bold.

## Step 2 - Total each bucket
Set total-ar to sum of range D2:D31.
Set ar-current to sum of range D2:D31 where range F2:F31 matches "Current".
Set ar-under-thirty to sum of range D2:D31 where range F2:F31 matches "1-30 days".
Set ar-under-sixty to sum of range D2:D31 where range F2:F31 matches "31-60 days".
Set ar-under-ninety to sum of range D2:D31 where range F2:F31 matches "61-90 days".
Set ar-over-ninety to sum of range D2:D31 where range F2:F31 matches "Over 90 days".
Set over-ninety-count to count of range F2:F31 matching "Over 90 days".

## Step 3 - Put the oldest invoices at the top
Sort range A1:F31 by column E descending with a header row.
Make row 1 a header row.
Format range D2:D31 as dollars.
Fit all columns.

## Step 4 - Build the aging summary
Work on sheet "AR Aging".
Clear everything from A1:C15.
Put "Accounts Receivable Aging" into cell A1.
Set font size of cell A1 to 18.
Make cell A1 bold.
Put "As of" into cell A2.
Put today into cell B2.
Format cell B2 as a long date.
Put "Bucket" into cell A4.
Put "Amount" into cell B4.
Put "Share of AR" into cell C4.
Put "Current" into cell A5.
Put ar-current into cell B5.
Put ar-current divided by total-ar into cell C5.
Put "1-30 days" into cell A6.
Put ar-under-thirty into cell B6.
Put ar-under-thirty divided by total-ar into cell C6.
Put "31-60 days" into cell A7.
Put ar-under-sixty into cell B7.
Put ar-under-sixty divided by total-ar into cell C7.
Put "61-90 days" into cell A8.
Put ar-under-ninety into cell B8.
Put ar-under-ninety divided by total-ar into cell C8.
Put "Over 90 days" into cell A9.
Put ar-over-ninety into cell B9.
Put ar-over-ninety divided by total-ar into cell C9.
Put "Total" into cell A10.
Put total-ar into cell B10.
Put 1 into cell C10.
Set fill-color of range A4:C4 to "#1F3A5F".
Set font-color of range A4:C4 to "#FFFFFF".
Make range A4:C4 bold.
Band every other row of A5:C9 "#EEF3F8".
Add a top border to range A10:C10.
Make range A10:C10 bold.
Format range B5:B10 as accounting in dollars.
Format range C5:C10 as percent with 1 decimal.
Set width of column A to 18.
Fit column B.
Set width of column C to 14.

## Step 5 - Raise a flag if too much is overdue
# Company policy: no more than 15% of receivables may be over 90 days old.
Set over-ninety-share to ar-over-ninety divided by total-ar.
If over-ninety-share is greater than 0.15:
  Set fill-color of cell C9 to "#9E1B32".
  Set font-color of cell C9 to "#FFFFFF".
  Make cell C9 bold.
  Show "Heads up: " joined with over-ninety-count joined with " invoices are over 90 days old. Start collection calls from the top of the Invoices sheet".
'@ },

@{ Id = '05'; Title = 'Sales Pivot by Region'; Area = 'Sales'; Format = 'docx'
   Uses = 'Sales'; Learn = 'Pivot tables, one plain sentence at a time, and re-running safely with Try.'
   Body = @'
= Sales Pivot by Region
~ Sample 05 - Sales. Builds two pivot tables beside the data on the Sales sheet.
# The Friday version of sample 02: the same orders, sliced the ways a sales manager asks for.

## Step 1 - Remove last week's pivots, if there are any
# "Try" means: attempt these lines, and if they cannot be done, do the "If that fails" part instead of stopping.
# The first time you run this there is nothing to remove, and that is fine.
Work on sheet "Sales".
Try:
  Delete pivot RegionPivot.
  Delete pivot ProductPivot.

If that fails:
  Log "No earlier pivots to remove".

## Step 2 - Revenue by region and channel
Make a pivot table from A1:G41 at J3 called RegionPivot.
Add rows of Region to pivot RegionPivot.
Add columns of Channel to pivot RegionPivot.
Add Revenue to pivot RegionPivot as a sum.
Show pivot RegionPivot in tabular form.
Sort Region in pivot RegionPivot descending by Revenue.

## Step 3 - Units and revenue by product, with each rep underneath
# Collapsed, it reads as one line per product. Double-click a product in Excel to see its reps.
Make a pivot table from A1:G41 at J16 called ProductPivot.
Add rows of Product, Rep to pivot ProductPivot.
Add Units, Revenue to pivot ProductPivot as a sum.
Add filters of Region to pivot ProductPivot.
Collapse Product in pivot ProductPivot.
Fit all columns.

# Next Friday: paste the new export over A1:G41 and run this again.
'@ },

@{ Id = '06'; Title = 'Month-End Close'; Area = 'Accounting'; Format = 'docx'
   Uses = 'Ledger -> Close Summary'; Learn = 'Manual steps kept as notes, a balance check with two outcomes, and locking the ledger.'
   Body = @'
= Month-End Close
~ Sample 06 - Accounting. Checks the trial balance on the Ledger sheet, then writes a Close Summary.
# Owner: the Controller. Run on business day 3, after every sub-ledger has posted.

## Before you run this - three steps done by hand
# 1. Export the month's trial balance from the accounting system.
# 2. Paste it over the Ledger sheet, keeping the column headers in row 1.
# 3. Confirm the bank feed has synced through the last day of the month.
# Steps like these stay in the procedure as notes. Frazaro skips them; the people running the close do not.

## Step 1 - Unlock and clean the trial balance
# If an earlier run locked the ledger, unlock it first. Unlocking a sheet that is not locked does nothing.
# The password only guards against accidental edits. It is not security.
Unprotect sheet "Ledger" with password "close".
Work on sheet "Ledger".
Replace "N/A" with 0 in range D2:E21.
Format range D2:E21 as accounting in dollars.

## Step 2 - Does it balance?
Set debits to sum of range D2:D21.
Set credits to sum of range E2:E21.
Set difference to debits minus credits.
Set difference to difference rounded to 2 decimals.

## Step 3 - The month's result
Set revenue to sum of range E2:E21 where range C2:C21 matches "Revenue".
Set expenses to sum of range D2:D21 where range C2:C21 matches "Expense".
Set net-income to revenue minus expenses.

## Step 4 - Flag every account still waiting on a reconciliation
Put "Status" into cell G1.
Count r from 2 to 21:
  Put "Ready" into column G row r.
  Set reconciled to cell in column F row r.
  If reconciled is "No", put "Reconcile before close" into column G row r.
  If reconciled is "No", make cell in column B row r bold.

Set unreconciled to count of range F2:F21 matching "No".
Make row 1 a header row.
Fit all columns.

## Step 5 - Write the close summary
Work on sheet "Close Summary".
Clear everything from A1:C20.
Put "Month-End Close" into cell A1.
Set font size of cell A1 to 18.
Make cell A1 bold.
Put "Run on" into cell A2.
Put today into cell B2.
Format cell B2 as a long date.
Put "Check" into cell A4.
Put "Result" into cell B4.
Put "Trial balance" into cell A5.
Put "Total debits" into cell A6.
Put debits into cell B6.
Put "Total credits" into cell A7.
Put credits into cell B7.
Put "Revenue" into cell A8.
Put revenue into cell B8.
Put "Expenses" into cell A9.
Put expenses into cell B9.
Put "Net income" into cell A10.
Put net-income into cell B10.
Put "Accounts to reconcile" into cell A11.
Put unreconciled into cell B11.
Put "Bank feed synced" into cell A12.
Put "Confirm by hand" into cell B12.
Put "Controller sign-off" into cell A13.
Put "Waiting" into cell B13.
Format range B6:B10 as accounting in dollars.
Set fill-color of range A4:B4 to "#1F3A5F".
Set font-color of range A4:B4 to "#FFFFFF".
Make range A4:B4 bold.
Band every other row of A5:B13 "#EEF3F8".
Add a top border to range A10:B10.
Make range A10:B10 bold.
Set width of column A to 26.

# The verdict goes in last, so the row banding above cannot paint over its colour.
If difference is 0:
  Put "Balanced" into cell B5.
  Set fill-color of cell B5 to "#507D2A".

Otherwise:
  Put "Out of balance by " joined with difference into cell B5.
  Set fill-color of cell B5 to "#9E1B32".

Set font-color of cell B5 to "#FFFFFF".
Make cell B5 bold.
Fit column B.

## Step 6 - Lock the period, but only when it is clean
# A locked ledger cannot be changed by accident after the books are closed.
If unreconciled is 0, protect sheet "Ledger" with password "close".
If unreconciled is greater than 0, show "The close summary is ready, but " joined with unreconciled joined with " accounts still need reconciling. The ledger stays unlocked until they are done".
'@ },

@{ Id = '07'; Title = 'Inventory Reorder'; Area = 'Operations'; Format = 'docx'
   Uses = 'Inventory -> Purchase Order'; Learn = 'A step of your own, lists, lookups by SKU, and a purchase order built line by line.'
   Body = @'
= Inventory Reorder
~ Sample 07 - Operations. Checks stock on the Inventory sheet, then writes a Purchase Order.
# Run every Thursday morning, so orders reach suppliers before the weekend.

## Step 1 - Teach Frazaro one thing: what an order really costs
# Freight adds 5% to every order. Say it once here, then use it by name in Step 5.
To landed-cost of amount:
  Give back amount times 1.05.

## Step 2 - Find everything at or below its reorder point
Work on sheet "Inventory".
Create a list called reorder-list.
Create a number called order-total.
Set order-total to 0.
Set last-row to last filled row of column A.
Count r from 2 to last-row:
  Set on-hand to cell in column D row r.
  Set reorder-point to cell in column E row r.
  Put "OK" into column H row r.
  If on-hand is at most reorder-point:
    Put "REORDER" into column H row r.
    Make cell in column H row r bold.
    Set sku-code to cell in column A row r.
    Append sku-code to reorder-list.
    Increase order-total by cell in column F row r times cell in column G row r.

Set reorder-count to count of reorder-list.
Sort range A1:H21 by column H descending then by column D ascending with a header row.
Make row 1 a header row.
Format range G2:G21 as dollars.
Fit all columns.

## Step 3 - Start a fresh purchase order
Work on sheet "Purchase Order".
Clear everything from A1:F40.
Put "Purchase Order" into cell A1.
Set font size of cell A1 to 18.
Make cell A1 bold.
Put "Prepared on" into cell A2.
Put today into cell B2.
Format cell B2 as a long date.
Put "SKU" into cell A4.
Put "Item" into cell B4.
Put "Supplier" into cell C4.
Put "Quantity" into cell D4.
Put "Unit cost" into cell E4.
Put "Line total" into cell F4.
Set fill-color of range A4:F4 to "#1F3A5F".
Set font-color of range A4:F4 to "#FFFFFF".
Make range A4:F4 bold.
Create a number called po-row.
Set po-row to 5.

## Step 4 - One line for every item that needs reordering
# Each lookup finds the SKU on the Inventory sheet and reads one column of its row.
For each sku in reorder-list:
  Go to sheet "Inventory".
  Set item-name to lookup of sku in range A2:G21 column 2.
  Set supplier to lookup of sku in range A2:G21 column 3.
  Set quantity to lookup of sku in range A2:G21 column 6.
  Set unit-cost to lookup of sku in range A2:G21 column 7.
  Go to sheet "Purchase Order".
  Put sku into column A row po-row.
  Put item-name into column B row po-row.
  Put supplier into column C row po-row.
  Put quantity into column D row po-row.
  Put unit-cost into column E row po-row.
  Put quantity times unit-cost into column F row po-row.
  Increase po-row by 1.

## Step 5 - Totals, with freight
Set landed-total to landed-cost of order-total.
Set landed-total to landed-total rounded to 2 decimals.
Increase po-row by 1.
Put "Subtotal" into column E row po-row.
Put order-total into column F row po-row.
Increase po-row by 1.
Put "With 5% freight" into column E row po-row.
Put landed-total into column F row po-row.
Format range E5:F40 as dollars.
Fit all columns.

Show "Purchase order ready: " joined with reorder-count joined with " items to reorder, " joined with landed-total joined with " including freight".
'@ },

@{ Id = '08'; Title = 'Shift Coverage'; Area = 'Operations'; Format = 'docx'
   Uses = 'Staff, Shifts, Leave -> Coverage'; Learn = 'Writing a staffing policy as sentences, then asking it questions that update themselves.'
   Body = @'
= Shift Coverage
~ Sample 08 - Operations. Writes the staffing policy in plain English, then asks who can cover which shift.
# The warehouse schedule lives in three Excel Tables: Staff (who holds which certification, at what level),
# Shifts (what each shift needs), and Leave (who is off, and for which shift).
# Instead of a formula nobody can read, the policy is written as sentences in column M.
# Every answer below is live: change a Table and the answers change with it.

## Step 1 - A fresh Coverage sheet
Work on sheet "Coverage".
Clear everything from A1:N40.
Put "Shift Coverage" into cell A1.
Set font size of cell A1 to 18.
Make cell A1 bold.

## Step 2 - Write the policy, one rule per cell
# A person can cover a shift when they hold the certification it needs, at a high enough level, and are not on leave for it.
Put "The policy" into cell M1.
Make cell M1 bold.
Write in cell M2 that a person can-cover a shift if Shifts lists the shift as Shift, the cert as Needs, and the min as MinLevel, and Staff lists the person as Name, the cert as Cert, and the level as Level, and the level is at least the min, and not Leave lists the person as Name and the shift as Shift.
Write in cell M3 that a shift is listed if Shifts lists the shift as Shift.
Write in cell M4 that a shift is covered if the person can-cover the shift.
Set width of column M to 80.
Wrap text in range M2:M4.

## Step 3 - Ask the questions a scheduler asks every week
Put "Who can work Night?" into cell A3.
Show in cell A4 who can-cover "Night" by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
Put "How many can work each shift?" into cell C3.
Show in cell C4 how many people can-cover each shift that is listed by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
Put "Shifts nobody can cover" into cell F3.
Show in cell F4 which shift that is listed is not covered by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
Put "Is every shift covered?" into cell H3.
Show in cell H4 whether every shift that is listed is covered by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
Put "Only one person for Night" into cell J3.
Show in cell J4 who alone can-cover "Night" by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
Put "Everyone, and every shift they can cover" into cell A13.
Show in cell A14 which person can-cover which shift by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
Make range A3:J3 bold.
Make cell A13 bold.
Set width of column A to 26.
Set width of column C to 30.
Set width of column F to 26.
Set width of column H to 24.
Set width of column J to 26.

## Step 4 - Now change the data and watch
# Try these on the Staff, Shifts and Leave sheets, then come back here:
# - Delete Cara's Night row from the Leave table. Night gains a second person, so "Only one person for Night" empties.
# - Give Ben Forklift at level 3 in the Staff table. Stocktake is covered, and "Is every shift covered?" turns TRUE.
# - Change a rule in column M. The policy is the text you can read; there is no hidden version.
'@ },

@{ Id = 'Joy'; Title = 'The Joy of Painting Palette'; Area = 'Any sheet'; Format = 'txt+docx'; File = 'joy'
   Uses = 'Joy'; Learn = "Bob Ross's fifteen paints and four bases as named colours: copy the Define lines into any program."
   Body = @'
= The Joy of Painting Palette
~ Bob Ross's fifteen paints and four bases, as colours any Frazaro program can use by name.
# Load this and press Interpret and Run to paint a swatch of every colour on a Joy sheet.
# To use the palette in your own procedure, copy the Define lines below to the top of it.
# Then write, for example: Set fill-color of range A1:D1 to phthalo-blue.
#
# Hex codes follow the Bob Ross paintings dataset (github.com/jwilber/Bob_Ross_Paintings),
# with two exceptions it does not give: Indian Red is the standard #CD5C5C, and Liquid White is #FFFFFF.
# Liquid Clear is transparent, so its #FFFFFF is only a stand-in.
# These are the paints straight from the tube, so several are very dark.
# There are no mistakes, just happy accidents.

## The palette
Define alizarin-crimson as "#4E1500".
Define bright-red as "#DB0000".
Define burnt-umber as "#8A3324".
Define cadmium-yellow as "#FFEC00".
Define dark-sienna as "#5F2E1F".
Define indian-red as "#CD5C5C".
Define indian-yellow as "#FFB800".
Define midnight-black as "#000000".
Define phthalo-blue as "#0C0040".
Define phthalo-green as "#102E3C".
Define prussian-blue as "#021E44".
Define sap-green as "#0A3410".
Define titanium-white as "#FFFFFF".
Define van-dyke-brown as "#221B15".
Define yellow-ochre as "#C79B00".

## The bases, painted on before the colour goes on
Define liquid-white as "#FFFFFF".
Define liquid-clear as "#FFFFFF".
Define liquid-black as "#000000".
Define black-gesso as "#000000".

## The swatches
Work on sheet "Joy".
Clear everything from A1:C30.
Put "The Joy of Painting" into cell A1.
Set font size of cell A1 to 18.
Make cell A1 bold.
Put "Paint" into cell A3.
Put "Hex" into cell B3.
Put "Swatch" into cell C3.
Make range A3:C3 bold.
Add a bottom border to range A3:C3.
Put "Alizarin Crimson" into cell A4.
Put alizarin-crimson into cell B4.
Set fill-color of cell C4 to alizarin-crimson.
Put "Bright Red" into cell A5.
Put bright-red into cell B5.
Set fill-color of cell C5 to bright-red.
Put "Burnt Umber" into cell A6.
Put burnt-umber into cell B6.
Set fill-color of cell C6 to burnt-umber.
Put "Cadmium Yellow" into cell A7.
Put cadmium-yellow into cell B7.
Set fill-color of cell C7 to cadmium-yellow.
Put "Dark Sienna" into cell A8.
Put dark-sienna into cell B8.
Set fill-color of cell C8 to dark-sienna.
Put "Indian Red" into cell A9.
Put indian-red into cell B9.
Set fill-color of cell C9 to indian-red.
Put "Indian Yellow" into cell A10.
Put indian-yellow into cell B10.
Set fill-color of cell C10 to indian-yellow.
Put "Midnight Black" into cell A11.
Put midnight-black into cell B11.
Set fill-color of cell C11 to midnight-black.
Put "Phthalo Blue" into cell A12.
Put phthalo-blue into cell B12.
Set fill-color of cell C12 to phthalo-blue.
Put "Phthalo Green" into cell A13.
Put phthalo-green into cell B13.
Set fill-color of cell C13 to phthalo-green.
Put "Prussian Blue" into cell A14.
Put prussian-blue into cell B14.
Set fill-color of cell C14 to prussian-blue.
Put "Sap Green" into cell A15.
Put sap-green into cell B15.
Set fill-color of cell C15 to sap-green.
Put "Titanium White" into cell A16.
Put titanium-white into cell B16.
Set fill-color of cell C16 to titanium-white.
Add a border around range C16.
Put "Van Dyke Brown" into cell A17.
Put van-dyke-brown into cell B17.
Set fill-color of cell C17 to van-dyke-brown.
Put "Yellow Ochre" into cell A18.
Put yellow-ochre into cell B18.
Set fill-color of cell C18 to yellow-ochre.
Put "Bases" into cell A20.
Make cell A20 bold.
Add a bottom border to range A20:C20.
Put "Liquid White" into cell A21.
Put liquid-white into cell B21.
Set fill-color of cell C21 to liquid-white.
Add a border around range C21.
Put "Liquid Clear" into cell A22.
Put liquid-clear into cell B22.
Set fill-color of cell C22 to liquid-clear.
Add a border around range C22.
Put "Liquid Black" into cell A23.
Put liquid-black into cell B23.
Set fill-color of cell C23 to liquid-black.
Put "Black Gesso" into cell A24.
Put black-gesso into cell B24.
Set fill-color of cell C24 to black-gesso.
Set width of column A to 20.
Set width of column B to 12.
Set width of column C to 16.
'@ },

@{ Id = 'Practice'; Title = 'Expense Reimbursement (as written)'; Area = 'Finance'; Format = 'docx'
   Uses = 'Expenses'; Learn = 'Load an SOP the way people really write them, and fix it line by line.'
   Body = @'
? Expense Reimbursement Run
? Performed by Accounts Payable every Friday afternoon.
? Purpose: pay employees back for claims that have a receipt.
? Open the expense report for the week.
Work on sheet "Expenses".
? Add up every claim that has a receipt.
Set receipted-total to sum of range D2:D26 where range E2:E26 matches "Yes".
Put "Reimbursement total" into cell H10.
Put receipted-total into cell I10.
Format cell I10 as dollars.
Make range H10:I10 bold.
? Send the total to payroll before 3pm.
Show "This week's reimbursement total: " joined with receipted-total.
'@ }
)

# =====================================================================
#  .txt rendering
# =====================================================================
function Render-Txt($body) {
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($raw in ($body -split "`r?`n")) {
        if     ($raw -match '^= (.*)$')  { $lines.Add('# ' + $Matches[1]) }
        elseif ($raw -match '^~ (.*)$')  { $lines.Add('# ' + $Matches[1]) }
        elseif ($raw -match '^## (.*)$') { $lines.Add('# ' + $Matches[1].ToUpperInvariant()) }
        elseif ($raw -match '^\? (.*)$') { $lines.Add($Matches[1]) }
        else                             { $lines.Add($raw) }
    }
    ($lines -join "`r`n") + "`r`n"
}

# =====================================================================
#  .docx rendering
# =====================================================================
$C_NAVY = '1F3A5F'; $C_TEAL = '0F7B8A'; $C_GREY = '66727D'; $C_INK = '1E2329'; $C_HASH = 'B8C1C9'

$DocxStyles = $XmlHead + @"
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rPr><w:color w:val="$C_INK"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FzTitle"><w:name w:val="Frazaro Title"/><w:basedOn w:val="Normal"/><w:qFormat/>
 <w:pPr><w:spacing w:after="40"/></w:pPr><w:rPr><w:rFonts w:ascii="Segoe UI Light" w:hAnsi="Segoe UI Light" w:cs="Segoe UI Light"/><w:color w:val="$C_NAVY"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FzSubtitle"><w:name w:val="Frazaro Subtitle"/><w:basedOn w:val="Normal"/><w:qFormat/>
 <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="8" w:color="$C_TEAL"/></w:pBdr><w:spacing w:after="160"/></w:pPr><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:cs="Segoe UI"/><w:color w:val="$C_GREY"/><w:sz w:val="22"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FzHeading"><w:name w:val="Frazaro Section"/><w:basedOn w:val="Normal"/><w:qFormat/>
 <w:pPr><w:keepNext/><w:spacing w:before="200" w:after="80"/></w:pPr><w:rPr><w:rFonts w:ascii="Segoe UI Semibold" w:hAnsi="Segoe UI Semibold" w:cs="Segoe UI Semibold"/><w:color w:val="$C_TEAL"/><w:sz w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FzNote"><w:name w:val="Frazaro Note"/><w:basedOn w:val="Normal"/><w:qFormat/>
 <w:pPr><w:keepNext/><w:spacing w:after="40"/></w:pPr><w:rPr><w:i/><w:color w:val="$C_GREY"/><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FzStep"><w:name w:val="Frazaro Instruction"/><w:basedOn w:val="Normal"/><w:qFormat/>
 <w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="10" w:color="$C_TEAL"/></w:pBdr><w:ind w:left="240"/><w:spacing w:after="30"/></w:pPr><w:rPr><w:color w:val="$C_INK"/><w:sz w:val="23"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FzStep2"><w:name w:val="Frazaro Instruction, in a block"/><w:basedOn w:val="FzStep"/><w:qFormat/>
 <w:pPr><w:ind w:left="720"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FzStep3"><w:name w:val="Frazaro Instruction, in a block in a block"/><w:basedOn w:val="FzStep"/><w:qFormat/>
 <w:pPr><w:ind w:left="1200"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FzProse"><w:name w:val="Frazaro Unmarked Prose"/><w:basedOn w:val="Normal"/><w:qFormat/>
 <w:pPr><w:spacing w:after="60"/></w:pPr><w:rPr><w:sz w:val="23"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FzProseTitle"><w:name w:val="Frazaro Unmarked Title"/><w:basedOn w:val="Normal"/><w:qFormat/>
 <w:pPr><w:spacing w:after="120"/></w:pPr><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:cs="Cambria"/><w:b/><w:color w:val="$C_INK"/><w:sz w:val="40"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FzGap"><w:name w:val="Frazaro Blank Line"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="200" w:lineRule="exact"/></w:pPr></w:style>
<w:style w:type="character" w:customStyle="1" w:styleId="FzHash"><w:name w:val="Frazaro Hash"/><w:rPr><w:i w:val="0"/><w:color w:val="$C_HASH"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="right" w:pos="9640"/></w:tabs></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="right" w:pos="9640"/></w:tabs></w:pPr></w:style>
</w:styles>
"@

function Run([string]$text, [string]$rStyle) {
    $rp = ''
    if ($rStyle) { $rp = "<w:rPr><w:rStyle w:val=`"$rStyle`"/></w:rPr>" }
    "<w:r>$rp<w:t xml:space=`"preserve`">$(Esc $text)</w:t></w:r>"
}
function Para([string]$style, [string]$runs) { "<w:p><w:pPr><w:pStyle w:val=`"$style`"/></w:pPr>$runs</w:p>" }

function Render-DocxBody($body, [bool]$practice) {
    $ps = New-Object System.Collections.Generic.List[string]
    $first = $true
    foreach ($raw in ($body -split "`r?`n")) {
        if     ($raw -match '^= (.*)$')  { $ps.Add((Para 'FzTitle'    ((Run '# ' 'FzHash') + (Run $Matches[1])))) }
        elseif ($raw -match '^~ (.*)$')  { $ps.Add((Para 'FzSubtitle' ((Run '# ' 'FzHash') + (Run $Matches[1])))) }
        elseif ($raw -match '^## (.*)$') { $ps.Add((Para 'FzHeading'  ((Run '# ' 'FzHash') + (Run $Matches[1])))) }
        elseif ($raw -match '^#( (.*))?$') { $ps.Add((Para 'FzNote'   ((Run '# ' 'FzHash') + (Run $Matches[2])))) }
        elseif ($raw -match '^\? (.*)$') {
            $st = 'FzProse'; if ($first) { $st = 'FzProseTitle' }
            $ps.Add((Para $st (Run $Matches[1])))
        }
        elseif ($raw.Trim().Length -eq 0) { $ps.Add('<w:p><w:pPr><w:pStyle w:val="FzGap"/></w:pPr></w:p>') }
        else {
            $indent = $raw.Length - $raw.TrimStart(' ').Length
            $st = 'FzStep'; if ($indent -ge 4) { $st = 'FzStep3' } elseif ($indent -ge 2) { $st = 'FzStep2' }
            $ps.Add((Para $st (Run $raw.Trim())))
        }
        $first = $false
    }
    $ps -join "`n"
}

function Build-Docx($sop, [string]$path) {
    $practice = ($sop.Id -eq 'Practice')
    $w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $body = Render-DocxBody $sop.Body $practice
    $doc = $XmlHead + "<w:document $w><w:body>`n$body`n" +
        '<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/><w:footerReference w:type="default" r:id="rIdFtr"/>' +
        '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1300" w:right="1300" w:bottom="1300" w:left="1300" w:header="600" w:footer="600" w:gutter="0"/></w:sectPr>' +
        '</w:body></w:document>'

    $kicker = if ($practice) { 'FRAZARO  |  PRACTICE SOP' }
              elseif ($sop.Id -eq 'Joy') { 'FRAZARO  |  A PALETTE FOR ANY PROGRAM' }
              else { "FRAZARO  |  SAMPLE SOP $($sop.Id) OF 08" }
    $small = '<w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:cs="Segoe UI"/><w:spacing w:val="12"/><w:color w:val="' + $C_GREY + '"/><w:sz w:val="16"/></w:rPr>'
    $hdr = $XmlHead + "<w:hdr $w><w:p><w:pPr><w:pStyle w:val=`"Header`"/><w:pBdr><w:bottom w:val=`"single`" w:sz=`"4`" w:space=`"6`" w:color=`"$C_HASH`"/></w:pBdr></w:pPr>" +
        "<w:r>$small<w:t xml:space=`"preserve`">$(Esc $kicker)</w:t></w:r><w:r>$small<w:tab/><w:t>$(Esc $sop.Area.ToUpperInvariant())</w:t></w:r></w:p></w:hdr>"

    $foot = if ($practice) {
        'Written the way real SOPs are, on purpose. Load it, press Validate Instructions, then add # to each line Frazaro flags, or rewrite it.'
    } else {
        'Lines that start with # are notes for people. Every other line is an instruction Frazaro checks, then runs.'
    }
    $tiny = '<w:rPr><w:color w:val="' + $C_GREY + '"/><w:sz w:val="16"/></w:rPr>'
    $ftr = $XmlHead + "<w:ftr $w><w:p><w:pPr><w:pStyle w:val=`"Footer`"/><w:ind w:right=`"900`"/></w:pPr>" +
        "<w:r>$tiny<w:t xml:space=`"preserve`">$(Esc $foot)</w:t></w:r><w:r>$tiny<w:tab/></w:r>" +
        "<w:fldSimple w:instr=`" PAGE `"><w:r>$tiny<w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>"

    $parts = [ordered]@{}
    $parts['[Content_Types].xml'] = $XmlHead + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
        '</Types>'
    $parts['_rels/.rels'] = $XmlHead + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
        '</Relationships>'
    $parts['word/_rels/document.xml.rels'] = $XmlHead + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
        '<Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
        '</Relationships>'
    $parts['word/document.xml'] = $doc
    $parts['word/styles.xml']   = $DocxStyles
    $parts['word/header1.xml']  = $hdr
    $parts['word/footer1.xml']  = $ftr
    $parts['docProps/core.xml'] = $XmlHead + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">' +
        "<dc:title>$(Esc $sop.Title)</dc:title><dc:subject>Frazaro sample SOP</dc:subject><dc:creator>Spreadsheet Company</dc:creator></cp:coreProperties>"
    $parts['docProps/app.xml'] = $XmlHead + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application></Properties>'
    Write-Zip $path $parts
}

# =====================================================================
#  THE SAMPLE DATA
# =====================================================================
function Serial([string]$iso) { ([datetime]::ParseExact($iso, 'yyyy-MM-dd', $inv) - [datetime]'1899-12-30').Days }

# A tiny fixed LCG, so the sales export is identical on every build and
# every .NET version (System.Random's sequence is not a promise).
$script:seed = 20260907
# The high bits: a power-of-two LCG's low bits cycle with a tiny period.
function Next([int]$n) { $script:seed = [int](([long]$script:seed * 1103515245 + 12345) % 2147483648); [int](($script:seed -shr 16) % $n) }

$repsBy = @{ North = @('Priya Nair', 'Tom Becker'); South = @('Luis Ortega', 'Mei Chen')
             East  = @('Sam Okafor', 'Grace Kim');  West  = @('Omar Haddad', 'Jess Moreno') }
$regions  = @('North', 'South', 'East', 'West')
$products = @('Widget', 'Gadget', 'Gizmo', 'Sprocket')
$prices   = @{ Widget = 25; Gadget = 40; Gizmo = 65; Sprocket = 12 }
$channels = @('Online', 'Retail', 'Wholesale')
$days     = @('2026-09-07', '2026-09-08', '2026-09-09', '2026-09-10', '2026-09-11')
$sales = @(,@('Date', 'Region', 'Rep', 'Product', 'Channel', 'Units', 'Revenue'))
for ($i = 0; $i -lt 40; $i++) {
    $reg = $regions[(Next 4)]; $rep = $repsBy[$reg][(Next 2)]
    $prod = $products[(Next 4)]; $ch = $channels[(Next 3)]
    $units = 2 + (Next 38)
    if ($ch -eq 'Wholesale') { $units = $units * 3 }
    $sales += ,@(@{ D = (Serial $days[[math]::Floor($i / 8)]) }, $reg, $rep, $prod, $ch, $units, ($units * $prices[$prod]))
}

$expenses = @(,@('Date', 'Employee', 'Category', 'Amount', 'Receipt'))
$expRows = @(
    @('2026-09-01', 'Priya Nair',  'Travel',   842.60, 'Yes'), @('2026-09-01', 'Tom Becker',  'Meals',     64.20, 'Yes'),
    @('2026-09-02', 'Luis Ortega', 'Lodging',  389.00, 'Yes'), @('2026-09-02', 'Mei Chen',    'Supplies',  46.75, 'No'),
    @('2026-09-02', 'Sam Okafor',  'Software', 129.00, 'Yes'), @('2026-09-03', 'Grace Kim',   'Meals',     38.10, 'Yes'),
    @('2026-09-03', 'Omar Haddad', 'Travel',   612.40, 'No'),  @('2026-09-03', 'Jess Moreno', 'Lodging',  455.00, 'Yes'),
    @('2026-09-04', 'Priya Nair',  'Meals',     92.35, 'Yes'), @('2026-09-04', 'Tom Becker',  'Travel',   238.90, 'Yes'),
    @('2026-09-04', 'Luis Ortega', 'Supplies', 118.00, 'Yes'), @('2026-09-05', 'Mei Chen',    'Travel',   1180.00, 'Yes'),
    @('2026-09-05', 'Sam Okafor',  'Meals',     27.80, 'No'),  @('2026-09-08', 'Grace Kim',   'Software', 599.00, 'Yes'),
    @('2026-09-08', 'Omar Haddad', 'Lodging',  312.00, 'Yes'), @('2026-09-08', 'Jess Moreno', 'Meals',     71.45, 'Yes'),
    @('2026-09-09', 'Priya Nair',  'Supplies',  33.99, 'Yes'), @('2026-09-09', 'Tom Becker',  'Lodging',  528.00, 'Yes'),
    @('2026-09-09', 'Luis Ortega', 'Travel',   156.30, 'Yes'), @('2026-09-10', 'Mei Chen',    'Meals',     48.60, 'Yes'),
    @('2026-09-10', 'Sam Okafor',  'Travel',   274.15, 'No'),  @('2026-09-10', 'Grace Kim',   'Supplies',  89.00, 'Yes'),
    @('2026-09-11', 'Omar Haddad', 'Meals',     55.25, 'Yes'), @('2026-09-11', 'Jess Moreno', 'Software',  49.00, 'Yes'),
    @('2026-09-11', 'Priya Nair',  'Lodging',  610.00, 'Yes')
)
foreach ($r in $expRows) { $expenses += ,@(@{ D = (Serial $r[0]) }, $r[1], $r[2], $r[3], $r[4]) }

$invoices = @(,@('Invoice', 'Customer', 'Issued', 'Amount', 'Days Overdue'))
$invRows = @(
    @('Harbor Dental',        '2026-08-28', 4200.00,   0), @('Maple & Finch',         '2026-05-12', 9850.00, 112),
    @('Bluebird Bakery',      '2026-08-14',  780.50,   5), @('Kestrel Logistics',     '2026-07-02', 12400.00, 48),
    @('Orchard Street Cafe',  '2026-08-30',  615.00,   0), @('Summit Fitness',        '2026-06-03', 3320.00, 77),
    @('Tidewater Marine',     '2026-04-21', 7600.00, 133), @('Copperleaf Design',     '2026-08-05', 2150.00, 14),
    @('Juniper Health',       '2026-08-22', 5480.00,   0), @('Redwood Supply',        '2026-07-18', 1975.00, 32),
    @('Harbor Dental',        '2026-07-09', 3900.00,  41), @('Lantern Books',         '2026-08-19',  420.00,   0),
    @('Silverline Transport', '2026-05-30', 6240.00,  94), @('Maple & Finch',         '2026-08-11', 2280.00,  8),
    @('Pinecrest Clinic',     '2026-06-20', 4410.00,  60), @('Bluebird Bakery',       '2026-08-26',  365.00,   0),
    @('Kestrel Logistics',    '2026-08-01', 8800.00,  18), @('Granite Peak Outfitters','2026-06-11', 2760.00, 69),
    @('Orchard Street Cafe',  '2026-07-24',  540.00,  26), @('Tidewater Marine',      '2026-08-29', 3150.00,   0),
    @('Summit Fitness',       '2026-07-15', 1480.00,  35), @('Copperleaf Design',     '2026-04-30', 2890.00, 124),
    @('Juniper Health',       '2026-07-28', 6120.00,  22), @('Redwood Supply',        '2026-06-26', 3380.00, 55),
    @('Lantern Books',        '2026-08-07',  690.00,  12), @('Silverline Transport',  '2026-08-24', 4050.00,   0),
    @('Pinecrest Clinic',     '2026-08-17', 1720.00,   3), @('Harbor Dental',         '2026-06-08', 2340.00,  84),
    @('Granite Peak Outfitters','2026-08-27', 1260.00,  0), @('Maple & Finch',        '2026-07-05', 3470.00,  44)
)
$n = 1041
foreach ($r in $invRows) { $invoices += ,@("INV-$n", $r[0], @{ D = (Serial $r[1]) }, $r[2], $r[3]); $n++ }

# Balanced on purpose: debits and credits both total 660,400.00 once the
# "N/A" placeholders become 0 (sample 06 does that). Three accounts are
# unreconciled, so the close stops short of locking the ledger.
$ledger = @(,@('Code', 'Account', 'Type', 'Debit', 'Credit', 'Reconciled'))
$ledRows = @(
    @(1000, 'Cash',                     'Asset',     84250.00, 'N/A',     'Yes'),
    @(1100, 'Accounts Receivable',      'Asset',    126400.00, 'N/A',     'Yes'),
    @(1200, 'Inventory',                'Asset',     58900.00, 'N/A',     'No'),
    @(1300, 'Prepaid Expenses',         'Asset',      6000.00, 'N/A',     'Yes'),
    @(1500, 'Equipment',                'Asset',    145000.00, 'N/A',     'Yes'),
    @(1510, 'Accumulated Depreciation', 'Asset',    'N/A',      38500.00, 'Yes'),
    @(2000, 'Accounts Payable',         'Liability','N/A',      61300.00, 'No'),
    @(2100, 'Accrued Payroll',          'Liability','N/A',      18750.00, 'Yes'),
    @(2200, 'Sales Tax Payable',        'Liability','N/A',       9420.00, 'Yes'),
    @(2500, 'Loan Payable',             'Liability','N/A',      90000.00, 'Yes'),
    @(3000, 'Owner Equity',             'Equity',   'N/A',     150000.00, 'Yes'),
    @(3100, 'Retained Earnings',        'Equity',   'N/A',      38030.00, 'Yes'),
    @(4000, 'Product Revenue',          'Revenue',  'N/A',     212600.00, 'Yes'),
    @(4100, 'Service Revenue',          'Revenue',  'N/A',      41800.00, 'Yes'),
    @(5000, 'Cost of Goods Sold',       'Expense',  118300.00, 'N/A',     'Yes'),
    @(6000, 'Salaries',                 'Expense',   96500.00, 'N/A',     'Yes'),
    @(6100, 'Rent',                     'Expense',   12000.00, 'N/A',     'Yes'),
    @(6200, 'Utilities',                'Expense',    3150.00, 'N/A',     'No'),
    @(6300, 'Depreciation',             'Expense',    4200.00, 'N/A',     'Yes'),
    @(6400, 'Software Subscriptions',   'Expense',    5700.00, 'N/A',     'Yes')
)
foreach ($r in $ledRows) { $ledger += ,$r }

$inventory = @(,@('SKU', 'Item', 'Supplier', 'On Hand', 'Reorder Point', 'Reorder Qty', 'Unit Cost', 'Status'))
$invtRows = @(
    @('WH-1001', 'Pallet wrap, 18 in',       'Crestline Packaging', 42, 20, 60,  18.50),
    @('WH-1002', 'Corrugated box, medium',   'Crestline Packaging', 180, 250, 500, 1.15),
    @('WH-1003', 'Corrugated box, large',    'Crestline Packaging', 310, 200, 400, 1.60),
    @('WH-1004', 'Packing tape, 6 pack',     'Crestline Packaging', 14, 24, 48,  11.90),
    @('WH-1005', 'Label roll, 4x6',          'Northgate Office',    9,  12, 36,  22.00),
    @('WH-1006', 'Nitrile gloves, box',      'Safeguard Supply',    35, 30, 60,   8.40),
    @('WH-1007', 'Safety vest, hi-vis',      'Safeguard Supply',    22, 10, 20,   6.75),
    @('WH-1008', 'Forklift propane tank',    'Valley Gas',          4,  6,  12,  48.00),
    @('WH-1009', 'Pallet, standard',         'Timberline Pallets',  95, 80, 150, 12.25),
    @('WH-1010', 'Stretch film dispenser',   'Crestline Packaging', 7,  3,  6,   34.00),
    @('WH-1011', 'Bubble wrap, 12 in roll',  'Crestline Packaging', 16, 16, 30,  27.50),
    @('WH-1012', 'Box cutter',               'Northgate Office',    40, 25, 50,   3.20),
    @('WH-1013', 'First aid refill kit',     'Safeguard Supply',    2,  4,  8,   39.95),
    @('WH-1014', 'Barcode scanner battery',  'Northgate Office',    11, 8,  16,  29.00),
    @('WH-1015', 'Dock bumper',              'Timberline Pallets',  6,  4,  4,  115.00),
    @('WH-1016', 'Shelf label holder, 50',   'Northgate Office',    3,  5,  10,  17.80),
    @('WH-1017', 'Ear plugs, box of 200',    'Safeguard Supply',    18, 10, 20,  24.60),
    @('WH-1018', 'Void fill paper, bundle',  'Crestline Packaging', 55, 60, 120,  9.10),
    @('WH-1019', 'Floor tape, yellow',       'Safeguard Supply',    12, 8,  16,  14.25),
    @('WH-1020', 'Hand truck tire',          'Valley Gas',          5,  2,  4,   26.40)
)
foreach ($r in $invtRows) { $inventory += ,($r + @('')) }

$staff  = @(@('Name', 'Cert', 'Level'),
            @('Ana', 'Forklift', 3), @('Ana', 'Lead', 2), @('Ben', 'Forklift', 1), @('Cara', 'Safety', 2),
            @('Cara', 'Lead', 3), @('Dev', 'Forklift', 2), @('Eli', 'Safety', 1), @('Fay', 'Forklift', 2),
            @('Fay', 'Safety', 3), @('Gus', 'Lead', 1))
$shifts = @(@('Shift', 'Needs', 'MinLevel'),
            @('Early', 'Forklift', 1), @('Late', 'Forklift', 2), @('Night', 'Lead', 2),
            @('Weekend', 'Safety', 1), @('Stocktake', 'Forklift', 3))
$leave  = @(@('Name', 'Shift'), @('Ana', 'Stocktake'), @('Cara', 'Night'), @('Dev', 'Late'))

# =====================================================================
#  .xlsx rendering
# =====================================================================
# Cell styles (cellXfs index): 0 plain, 1 header (white on navy),
# 2 short date, 3 title, 4 subtitle, 5 bold navy, 6 big step number,
# 7 wrapped body, 8 banded row, 9 banded bold, 10 small grey note.
$XlsxStyles = $XmlHead + @"
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="7">
<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><sz val="24"/><color rgb="FF$C_NAVY"/><name val="Segoe UI Light"/><family val="2"/></font>
<font><sz val="12"/><color rgb="FF$C_GREY"/><name val="Segoe UI"/><family val="2"/></font>
<font><b/><sz val="11"/><color rgb="FF$C_NAVY"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><sz val="18"/><color rgb="FF$C_TEAL"/><name val="Segoe UI Semibold"/><family val="2"/></font>
<font><i/><sz val="10"/><color rgb="FF$C_GREY"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF$C_NAVY"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFEEF3F8"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="11">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment vertical="center"/></xf>
<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="4" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"@

function ColName([int]$i) { [string][char](64 + $i) }

function CellXml([string]$ref, $v, [int]$s) {
    $sa = ''; if ($s -gt 0) { $sa = " s=`"$s`"" }
    if ($v -is [hashtable]) { return "<c r=`"$ref`" s=`"2`"><v>$($v.D)</v></c>" }
    if ($null -eq $v -or ($v -is [string] -and $v.Length -eq 0)) { if ($s -gt 0) { return "<c r=`"$ref`"$sa/>" }; return '' }
    if ($v -is [int] -or $v -is [double] -or $v -is [decimal] -or $v -is [long]) {
        return "<c r=`"$ref`"$sa><v>$(([double]$v).ToString('R', $inv))</v></c>"
    }
    "<c r=`"$ref`"$sa t=`"inlineStr`"><is><t xml:space=`"preserve`">$(Esc ([string]$v))</t></is></c>"
}

# A sheet from rows; $opts: HeaderStyle, Widths, Heights (row->pt), Styles
# (ref->style for individual cells), Grid, Selected, Tab, Table.
function SheetXml($rows, $opts) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append($XmlHead + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    if ($opts.Tab) { [void]$sb.Append("<sheetPr><tabColor rgb=`"FF$($opts.Tab)`"/></sheetPr>") }
    $sv = '<sheetView workbookViewId="0"'
    if ($opts.Selected) { $sv += ' tabSelected="1"' }
    if ($opts.NoGrid)   { $sv += ' showGridLines="0"' }
    [void]$sb.Append("<sheetViews>$sv/></sheetViews><sheetFormatPr defaultRowHeight=`"15`"/>")
    if ($opts.Widths) {
        [void]$sb.Append('<cols>')
        for ($c = 0; $c -lt $opts.Widths.Count; $c++) {
            $w = $opts.Widths[$c]; $i = $c + 1
            [void]$sb.Append("<col min=`"$i`" max=`"$i`" width=`"$w`" customWidth=`"1`"/>")
        }
        [void]$sb.Append('</cols>')
    }
    [void]$sb.Append('<sheetData>')
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $rn = $r + 1
        $ht = ''
        if ($opts.Heights -and $opts.Heights.ContainsKey($rn)) { $ht = " ht=`"$($opts.Heights[$rn])`" customHeight=`"1`"" }
        [void]$sb.Append("<row r=`"$rn`"$ht>")
        $row = $rows[$r]
        for ($c = 0; $c -lt $row.Count; $c++) {
            $ref = (ColName ($c + 1)) + $rn
            $s = 0
            if ($r -eq 0 -and $opts.HeaderStyle) { $s = $opts.HeaderStyle }
            if ($opts.Styles -and $opts.Styles.ContainsKey($ref)) { $s = $opts.Styles[$ref] }
            [void]$sb.Append((CellXml $ref $row[$c] $s))
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')
    if ($opts.Table) { [void]$sb.Append('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>') }
    [void]$sb.Append('</worksheet>')
    $sb.ToString()
}

function TableXml([int]$id, [string]$name, $rows) {
    $ref = "A1:$(ColName $rows[0].Count)$($rows.Count)"
    $cols = ($rows[0] | ForEach-Object -Begin { $k = 0 } -Process { $k++; "<tableColumn id=`"$k`" name=`"$(Esc $_)`"/>" }) -join ''
    $XmlHead + "<table xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" id=`"$id`" name=`"$name`" displayName=`"$name`" ref=`"$ref`" totalsRowShown=`"0`">" +
        "<autoFilter ref=`"$ref`"/><tableColumns count=`"$($rows[0].Count)`">$cols</tableColumns>" +
        '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'
}

function Build-Workbook([string]$path) {
    # ---- Start Here ----
    $S = @{}
    $start = New-Object System.Collections.Generic.List[object]
    $start.Add(@()) # row 1: breathing room
    $start.Add(@('', 'Welcome to Frazaro')); $S['B2'] = 3
    $start.Add(@('', 'Automate this workbook in plain English. Eight sample procedures, from five sentences to a staffing policy.')); $S['B3'] = 4
    $start.Add(@())
    $start.Add(@('', 'How to run a sample')); $S['B5'] = 5
    $steps = @(
        'On the Frazaro tab of the ribbon, press Load Instructions.',
        'Pick a sample from the examples folder. Start with "01 Tidy the Sales Export.txt".',
        'Press Validate Instructions. Every line gets a green OK, or a note in words saying what to change. Nothing has run yet.',
        'Press Interpret and Run. Changed your mind? Undo Last Run puts the workbook back the way it was.'
    )
    for ($i = 0; $i -lt 4; $i++) {
        $rn = 6 + $i
        $start.Add(@('', ($i + 1), $steps[$i])); $S["B$rn"] = 6; $S["C$rn"] = 7
    }
    $start.Add(@())
    $start.Add(@('', 'The samples, easiest first')); $S['B11'] = 5
    $start.Add(@('', '#', 'Sample', 'Area', 'Works on', "What you'll learn"))
    foreach ($col in 'B', 'C', 'D', 'E', 'F') { $S["${col}12"] = 1 }
    $rn = 13
    foreach ($sop in $sops) {
        $file = if ($sop.Id -eq 'Practice') { "Practice - $($sop.Title).docx" }
                elseif ($sop.File) { "$($sop.File).txt or $($sop.File).docx" }
                else { "$($sop.Id) $($sop.Title).$($sop.Format)" }
        $label = if ($sop.Id -eq 'Practice') { '+' } elseif ($sop.Id -eq 'Joy') { '*' } else { [int]$sop.Id }
        $start.Add(@('', $label, $file, $sop.Area, $sop.Uses, $sop.Learn))
        $band = (($rn - 13) % 2 -eq 1)
        foreach ($col in 'B', 'C', 'D', 'E', 'F') { $S["$col$rn"] = $(if ($band) { 8 } else { 7 }) }
        if ($band) { $S["C$rn"] = 9 } else { $S["C$rn"] = 5 }
        $rn++
    }
    $start.Add(@())
    $notes = @(
        'Lines that start with # are notes for people. Every other line is an instruction Frazaro checks before anything runs.',
        'The .docx samples need Microsoft Word on this computer (Frazaro reads them through Word, with macros disabled). The .txt samples need nothing extra.',
        'Every sample works on this workbook. Sales is a raw export on purpose: sample 01 tidies it. Staff, Shifts and Leave are Excel Tables for sample 08.',
        'Save a copy before you start if you want to keep the untouched data - or just press Undo Last Run after each sample.'
    )
    foreach ($t in $notes) { $rn++; $start.Add(@('', '', $t)); $S["C$rn"] = 10 }
    $H = @{ 2 = 36; 3 = 22; 5 = 22 }
    for ($i = 6; $i -le 9; $i++) { $H[$i] = 30 }
    for ($i = 13; $i -lt 13 + $sops.Count; $i++) { $H[$i] = 32 }
    for ($i = 14 + $sops.Count; $i -le $rn; $i++) { $H[$i] = 28 }

    $sheets = @(
        @{ Name = 'Start Here'; Rows = $start; Opts = @{ Selected = $true; NoGrid = $true; Tab = $C_TEAL; Styles = $S; Heights = $H; Widths = @(3, 6, 44, 13, 26, 62) } },
        @{ Name = 'Sales';      Rows = $sales;     Opts = @{} },
        @{ Name = 'Expenses';   Rows = $expenses;  Opts = @{ HeaderStyle = 5; Widths = @(12, 16, 12, 11, 10, 24) } },
        @{ Name = 'Invoices';   Rows = $invoices;  Opts = @{ HeaderStyle = 5; Widths = @(11, 24, 12, 12, 13, 14) } },
        @{ Name = 'Ledger';     Rows = $ledger;    Opts = @{ HeaderStyle = 5; Widths = @(8, 26, 11, 14, 14, 12, 24) } },
        @{ Name = 'Inventory';  Rows = $inventory; Opts = @{ HeaderStyle = 5; Widths = @(10, 26, 22, 10, 14, 12, 11, 11) } },
        @{ Name = 'Staff';      Rows = $staff;     Opts = @{ Widths = @(12, 12, 10); Table = 'Staff' } },
        @{ Name = 'Shifts';     Rows = $shifts;    Opts = @{ Widths = @(14, 12, 12); Table = 'Shifts' } },
        @{ Name = 'Leave';      Rows = $leave;     Opts = @{ Widths = @(12, 14); Table = 'Leave' } }
    )

    $parts = [ordered]@{}
    $ct = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
          '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    $wbSheets = ''; $wbRels = ''; $tableId = 0
    for ($i = 0; $i -lt $sheets.Count; $i++) {
        $n = $i + 1; $sh = $sheets[$i]
        $ct += "<Override PartName=`"/xl/worksheets/sheet$n.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
        $wbSheets += "<sheet name=`"$(Esc $sh.Name)`" sheetId=`"$n`" r:id=`"rId$n`"/>"
        $wbRels += "<Relationship Id=`"rId$n`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$n.xml`"/>"
        $parts["xl/worksheets/sheet$n.xml"] = SheetXml $sh.Rows $sh.Opts
        if ($sh.Opts.Table) {
            $tableId++
            $ct += "<Override PartName=`"/xl/tables/table$tableId.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml`"/>"
            $parts["xl/worksheets/_rels/sheet$n.xml.rels"] = $XmlHead + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                "<Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table`" Target=`"../tables/table$tableId.xml`"/></Relationships>"
            $parts["xl/tables/table$tableId.xml"] = TableXml $tableId $sh.Opts.Table $sh.Rows
        }
    }
    $ct += '</Types>'
    $wbRels += "<Relationship Id=`"rId$($sheets.Count + 1)`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>"

    $final = [ordered]@{}
    $final['[Content_Types].xml'] = $XmlHead + $ct
    $final['_rels/.rels'] = $XmlHead + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
        '</Relationships>'
    $final['xl/workbook.xml'] = $XmlHead + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        "<bookViews><workbookView activeTab=`"0`"/></bookViews><sheets>$wbSheets</sheets></workbook>"
    $final['xl/_rels/workbook.xml.rels'] = $XmlHead + "<Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`">$wbRels</Relationships>"
    $final['xl/styles.xml'] = $XlsxStyles
    foreach ($k in $parts.Keys) { $final[$k] = $parts[$k] }
    $final['docProps/core.xml'] = $XmlHead + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">' +
        '<dc:title>Frazaro Sample Data</dc:title><dc:creator>Spreadsheet Company</dc:creator></cp:coreProperties>'
    $final['docProps/app.xml'] = $XmlHead + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>'
    Write-Zip $path $final
}

# =====================================================================
#  BUILD
# =====================================================================
foreach ($sop in $sops) {
    $name = if ($sop.Id -eq 'Practice') { "Practice - $($sop.Title)" }
            elseif ($sop.File) { $sop.File }
            else { "$($sop.Id) $($sop.Title)" }
    foreach ($fmt in ($sop.Format -split '\+')) {
        $path = Join-Path $out "$name.$fmt"
        # A file open in Word or Excel mid-test is skipped, not fatal: one
        # locked sample must not stop every other sample being written.
        try {
            if ($fmt -eq 'txt') {
                [System.IO.File]::WriteAllText($path, (Render-Txt $sop.Body), (New-Object System.Text.ASCIIEncoding))
            } else {
                Build-Docx $sop $path
            }
            Write-Host "  wrote $name.$fmt"
        } catch [System.IO.IOException] {
            Write-Host "  SKIPPED $name.$fmt - it is open (close it and rerun to rebuild it)"
        }
    }
}
# Last, and allowed to be skipped: the workbook is the file most likely to
# be open in Excel mid-test, and a locked data file must not stop the SOPs
# (the usual reason to rebuild) from being written.
try {
    Build-Workbook (Join-Path $out $DataBook)
    Write-Host "  wrote $DataBook"
} catch [System.IO.IOException] {
    Write-Host "  SKIPPED $DataBook - it is open (close it in Excel and rerun to rebuild it)"
}
Write-Host 'examples built.'
