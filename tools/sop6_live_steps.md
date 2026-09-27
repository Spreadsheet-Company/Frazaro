# SOP.6 - live steps

*The `<Frazaro>` tag. Eleven steps. Move this file to `archive/` once SOP.6
has landed.*

**Passed live 2026-09-26: all eleven steps**, over three runs (see the notes
at the end). Final suites: pure 1493/1493, host 202/202, `TestDSLs`
2299/2299, `VerifyReports` 283/283 on both backends.

**Before you start.** `VlaDevReload`, then `Debug > Compile VBAProject`. No
module is new: `VLA_IDE.bas`, `VLA_Messages.bas` and `VLA_Tests.bas` changed,
and nothing was added to or removed from either `mods` array. The G-TEXT
session has uncommitted work of its own in `VLA_Tests.bas`,
`VLA_SentenceEngine.bas` and elsewhere, and a reload picks that up too.

The examples were rebuilt for this item. `examples\Frazaro Sample Data.xlsx`
is not in git; the copy on this machine now has the new Start Here text and
is the one to open. If it was open in Excel during the rebuild, close it and
run `powershell -File tools\build_examples.ps1` first. Every fixture below
lives in `tools\sop6_fixtures\`.

Predicted before the pass, so you can hold me to it:

- `VlaSelfTest` **0 failed**, and **1460** from HEAD's recorded 1388 plus
  `TestIdeTags`'s **72**. The G-TEXT session's uncommitted `TestTextValues`
  adds its own on top (31 assertion sites when I counted, so about 1491 if
  its code is in when you reload). If the total differs but nothing FAILS,
  their count moved, not mine. All 72 `sop6:` lines were pre-run against a
  line-for-line PowerShell port that reads them out of `VLA_Tests.bas`: 72/72.
- `VlaSelfTestHost` and `VerifyReports`: untouched by this item.
- **The one I am least sure of is step 4**: Word's own typography on a tag
  typed by hand. The scan runs after `NormalizeProgramText`, which
  straightens curly quotes, but nothing can pin that without Word.

---

## 1. It compiles, and the version moved

In the Immediate window:

```
?VLA_IDE.VLA_IDE_VERSION
```

**Expected:** compiles with no error, and prints `SOP.6`.

## 2. The suite

```
?VlaSelfTest
```

**Expected:** 0 failed, and the total predicted above. Paste back the
summary line; it prints last.

## 3. Sample 00, from Word

Open `examples\Frazaro Sample Data.xlsx`. On the Frazaro tab press **New
Frazaro** if the workbook has no program sheet yet, then **Load
Instructions** and pick `examples\00 Weekly Expense Reimbursement.docx`.

**Expected**, column B, rows 1 to 9, exactly:

```
# From "00 Weekly Expense Reimbursement.docx": Frazaro read only the <Frazaro> section on lines 13-21, and left out the document's other 11 lines.
# The section on lines 13-21: week = "Week 39"; (steps 1 and 2, automated by Accounts Payable)
Work on sheet "Expenses".
Set receipted-total to sum of range D2:D26 where range E2:E26 matches "Yes".
Put "Reimbursement total, Week 39" into cell H10.
Put receipted-total into cell I10.
Format cell I10 as dollars.
Make range H10:I10 bold.
Show "Week 39 reimbursement total: " joined with receipted-total.
```

Every row gets a green OK. Then press **Interpret and Run**: on the
*Expenses* sheet, H10 reads `Reimbursement total, Week 39`, I10 reads
`$6,099.74` (21 receipted claims of 25), both bold, and a message says
`Week 39 reimbursement total: 6099.74`. **Undo Last Run** afterwards.

## 4. Word's own quotes, typed by hand

Open the same `.docx` in Word. In the tag line, delete `"Week 39"` and type
`"Week 40"` - Word turns the quotes curly as you type. Save (Frazaro reads
the file as saved, through a Word of its own, so this Word can stay open).
In Excel press **Reload Instructions**.

**Expected:** rows 2, 5 and 9 now say `Week 40`, and every row is green. Close Word without further changes; the next rebuild of the examples
restores the file.

## 5. Two sections, the second left open

**Load Instructions**, `tools\sop6_fixtures\sections.txt`.

**Expected**, column B, rows 1 to 10, exactly (row 5 and row 10 are empty):

```
# From "sections.txt": Frazaro read only the 2 <Frazaro> sections (on lines 3-7 and from line 9 to the end), and left out the document's other 3 lines.
Work on sheet "Ledger".
Replace "N/A" with 0 in range D2:E21.
Set debits to sum of range D2:D21.

# The section from line 9 to the end: (part 2)
Set credits to sum of range E2:E21.
Set difference to debits minus credits.
Show "Difference: " joined with difference.

```

Every instruction row is green: a name set in the first section is used in
the second. **Interpret and Run** shows `Difference: 0` (the Ledger balances
by design). **Undo Last Run** afterwards.

## 6. The point of pain teaches the tag

**Load Instructions**, `examples\Practice - Expense Reimbursement (as
written).docx` - a whole SOP with no tag.

**Expected:** one row near the top is red - the first line Frazaro cannot
read; the title has no full stop, so the sentence may be charged to row 1
or row 2 - and its note in column C ends with:

```
(Loading a whole SOP? Put <Frazaro> on a line of its own above the instructions and </Frazaro> below them, then press Reload Instructions: Frazaro reads only that part, and the rest of the document can stay exactly as it is.)
```

Then press **Validate Instructions**: the same row is red again, and this
time its note has no such sentence - only the Check right after a load
carries it.

## 7. A refused tag leaves the sheet alone

With step 6's program still on the sheet, **Load Instructions**,
`tools\sop6_fixtures\approved.txt`.

**Expected:** this message, and column B still holding the practice SOP's
lines, untouched:

```
approved.txt, line 2: Frazaro does not know the word 'approved' in a tag. A word on its own in a tag is for Frazaro - a version like 0.6.2, or a language like espanol. To keep it as a note for people, put it in parentheses: (approved). To make it a value the section can use, give it a name, like status=approved. Nothing was loaded, and the program on this sheet has not been changed.
```

## 8. A version newer than this Frazaro

First check the name the message will use:

```
?VLA_RELEASE_VERSION
```

It should print `0.6.2`; if a release has moved it, the message below names
the new one instead. **Load Instructions**, `tools\sop6_fixtures\version.txt`.

**Expected:**

```
version.txt, line 1: this section needs Frazaro 9.0.0 or later, and this copy is Frazaro 0.6.2. Update Frazaro, then load the file again. Nothing was loaded, and the program on this sheet has not been changed.
```

## 9. A language this Frazaro does not have

First check this workbook asks for no German phrasebook:

```
?VLA_IDE.LoadedPhrasebookPaths(ActiveWorkbook).Count
```

It should print `0`. If it does not, none of the paths may end in
`deutsche.vla` (`?VLA_IDE.LoadedPhrasebookPaths(ActiveWorkbook)(1)` and so
on), or this step's premise is gone. **Load Instructions**,
`tools\sop6_fixtures\deutsche.txt`.

**Expected:**

```
deutsche.txt, line 1: this section is written for the deutsche phrasebook, which this copy of Frazaro does not have. Add it with Load Phrasebook, or use the edition of Frazaro made for it. Nothing was loaded, and the program on this sheet has not been changed.
```

## 10. A tag's contents, accepted

**Load Instructions**, `tools\sop6_fixtures\accepted.txt`.

**Expected**, column B, rows 1 to 4, exactly, every row green:

```
# From "accepted.txt": Frazaro read only the <Frazaro> section on lines 1-4.
# The section on lines 1-4: Frazaro 0.6.2; english; (reviewed by Accounts Payable)
Work on sheet "Expenses".
Fit all columns.
```

## 11. The red row names the ribbon, never VBA

*Added after the first pass, 2026-09-26: step 6's note told a user to print
`?EnglishListPhrases` in the Immediate window, and named a "Known Sentences"
button the ribbon has not had since LE.1.* `VlaDevReload` and compile first
(`VLA_Messages.bas` and `VLA_SentenceEngine.bas` changed), then **Load
Instructions**, `examples\Practice - Expense Reimbursement (as
written).docx`.

**Expected:** the red row's note in column C reads, with the sentence count
your grammar has in place of 234:

```
Don't understand: 'expense reimbursement run performed by accounts payable every friday afternoon.' No loaded sentence starts with 'expense' - press What can I say? on the Frazaro tab to see all 234 sentences this program understands. (line 1) (Loading a whole SOP? Put <Frazaro> on a line of its own above the instructions and </Frazaro> below them, then press Reload Instructions: Frazaro reads only that part, and the rest of the document can stay exactly as it is.)
```

No "VBA", no "Immediate window", and no `?`.

---

**First pass, 2026-09-26:** steps 1 to 6 passed - pure 1493/1493, host
202/202, `SOP.6`, Week 39 filled in, Week 40 typed with Word's curly quotes,
both sections green and run, and the hint on the red row - and step 6 found
the VBA advice step 11 now checks. Steps 7 to 10 were not in the report.

**Second run, 2026-09-26:** pure 1491/1493 - two grammar pins (`u9`, `l8r`)
still asserted the old advice. Both are fixed (`VLA_Messages.bas`,
`VLA_Tests_Grammar.bas`), so with the reload for step 11, step 2 should read
pure **1493/1493** again, host 202/202, and the two lines should read
`PASS  u9: a same-path re-carry teaches Reload, not rename` and
`PASS  l8r: an operator in statement position refuses with directions, and
names no VBA`.

Paste back step 2's summary line, and for each other step either PASS or
what you saw instead.
