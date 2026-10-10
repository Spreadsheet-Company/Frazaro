# Level 0 · Introduction: Natural Language Macros

*(An abstract overview of Frazaro)*

| | |
|---|---|
| **For** | Anyone deciding whether to read further: the analyst, the team lead, the IT reviewer, the person who signs. |
| **You need** | Nothing. Excel can stay closed. |
| **You will be able to** | Say what Frazaro is, what it promises, where it stops, and which of the five manuals after this one is yours. |
| **Plan on** | Fifteen minutes. |
| **Written against** | Frazaro `0.7.0`. Sentences that arrive in `0.7.1` are marked where they appear. |

---

## Abstract

An abstract overview ought to begin with an abstract, so here is one.

> Frazaro is an add-in for Microsoft Excel that reads a procedure written as
> English sentences, checks every sentence against a published grammar,
> refuses in words the sentences it cannot read, and runs only when every sentence 
> translates successfully to VBA. The grammar is a data file called a *phrasebook*; 
> English is the first phrasebook and not the last. Each accepted sentence is 
> translated mechanically into a small middle language, VLA, which is either
> interpreted directly on the workbook or compiled into readable VBA for
> audit. No statistical model takes part at any stage: the same sentences
> produce the same result on every run, on the user's own machine, with no
> network connection. This document states the problem, describes the
> design, lists what the design guarantees and what it does not, and maps
> the five manuals that follow.
>
> **Keywords:** controlled natural language; spreadsheets; standard
> operating procedures; end-user programming; Lisp; refusal.

The rest of this page is the long form of that paragraph.

---

## 1. The problem: two documents that ought to be one

Every team that lives in spreadsheets keeps procedures. The month-end close,
the Friday reimbursement run, the aging report, the purchase order. Each
exists twice.

The first copy is the **procedure**: a Word document with a title, an owner,
numbered steps and a revision history. People can read it. Excel cannot.

The second copy, where there is one, is the **macro**: a VBA module somebody
wrote in 2019. Excel can run it. The person accountable for the result
cannot read it, and the person who could has changed jobs.

The two drift apart, quietly, and the auditor reads the first while Excel
runs the second.

A **natural language macro** is a single document that is both. It is
written in sentences a colleague can read aloud, and those same sentences
are what runs. Frazaro's founding sentence says it more formally:

> The language in which a person thinks (source thought) should be the same
> language in which a person works (source code).

## 2. A specimen

This is a real standard operating procedure, kept the way such things are
kept. Two lines were added to it, the ones in angle brackets, and Frazaro
reads only what sits between them.

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

It is not pseudocode. It is the first of the nine
[sample procedures](../../examples/README.md), and it runs.

| The sentence | What happens |
|---|---|
| `Work on sheet "Expenses".` | The Expenses sheet becomes the place the following sentences act on. It is created if it does not exist. |
| `Set receipted-total to sum of …` | Adds the amounts in column D whose receipt column says Yes, and gives the result a name. |
| `Put "Reimbursement total, {week}" into cell H10.` | Writes a label. `{week}` is filled in from the opening tag, so next week's run changes `Week 39` in one place. |
| `Put receipted-total into cell I10.` | Writes the number. |
| `Format cell I10 as dollars.` | Shows it as money. |
| `Make range H10:I10 bold.` | Makes the pair stand out. |
| `Show "…" joined with receipted-total.` | Tells the person running it what the total was. |

Step 3, sending the total to payroll before three o'clock, is still done by
a person. It stays in the document as prose, where a person will see it.

## 3. From sentence to spreadsheet

```text
   a Word document, a text file, or the sheet itself
                        |
                 the sentence engine ......... reads each row, word by word
     sentence shapes from the phrasebook, tried in a fixed order;
     the first shape that fits wins, and there is no second reading
                        |
                       VLA .................... one statement of meaning
                     /       \
            interpreter     compiler
            acts on the     writes readable VBA,
            live workbook   each line tagged with its sentence
                     \       /
          the same programs, both ways, must agree
```

Four things happen to a sentence, always in this order.

1. **It is loaded.** From a `.docx`, from a `.txt`, or typed straight into
   the *Frazaro* sheet, one sentence per row.
2. **It is validated.** The engine drops filler words such as "the" and
   "please", then tries its sentence shapes in order. Each shape has typed
   holes: a hole marked as a cell must look like a cell. A row that fits is
   marked OK in green. A row that does not fit gets a note beside it, in
   words. Nothing has run yet.
3. **It is translated.** The pieces of the sentence are copied into the
   shape's template, the way a form is filled in.

   ```text
   English   Make cell A1 bold.
   VLA       (make-bold (range "a1"))
   VBA       range("a1").font.bold = True
   ```

4. **It is run, one of two ways.** *Interpret* walks the VLA and acts on the
   workbook directly; no code is written into your file. *Compile* turns the
   VLA into a VBA module you can read, keep, and hand to an auditor. Before
   either, the sheets the program names are copied, so the run can be taken
   back.

## 4. Five promises, and how each is kept

A promise without a mechanism is a slogan, so each of these names the
mechanism and where you can check it.

| Promise | How it is kept | Check it yourself |
|---|---|---|
| **Checked before anything runs.** | *Validate Instructions* reads every sentence first. A misspelled step name or a range that is not a range is caught there, not halfway through a run. | Misspell a word in any sample and press **Validate Instructions**. |
| **One sentence, one meaning.** | Shapes are tried in a fixed order and the first match wins. Because the shapes are finite, the loader can decide whether two of them overlap, and it refuses a phrasebook where they do. A sentence that has shipped keeps its meaning. | [GRAMMAR_SINCE.md](../GRAMMAR_SINCE.md) dates every sentence shape. |
| **Refusals teach.** | A refused sentence is told how far the reading got, what was expected next, and what was found instead. | Level 1, Lesson 18. |
| **Nothing leaves the machine.** | Frazaro has no update check, no telemetry and no network call of any kind, by standing decision. A script scans the code for one before every release. | `tools/check_no_network.ps1`; [IT_REVIEW.md](../IT_REVIEW.md) §3. |
| **A run can be taken back.** | **Undo Last Run** restores the sheets the program named, including sheets the run created or deleted, and puts the tabs back in their order and visibility, with the sheet you were on in front. | Run any sample, then press **Undo Last Run**. |

The fifth promise has an edge worth knowing before you rely on it. Undo puts
back *sheets*: their contents, their names, their order and whether they
show. A file the program saved to disk, or an email draft it opened, is
outside any sheet and stays as it is. When a run stops partway, the message
says exactly this.

A real refusal, from the engine that answers questions about tables:

> fact 'parent' uses 'X', which looks like a variable (it starts with a
> capital letter) - facts must be fully specific; did you mean to write a
> rule instead?

## 5. What Frazaro is not

**It is not AI.** There is no model and no statistics. Frazaro can refuse a
sentence; it cannot guess at one. If you want a draft in your own words that
a person will then check, a prompt box is the right tool, and
[COMPARISON.md](../COMPARISON.md) says so without embarrassment.

**It is not a macro recorder.** A recorder writes down what your mouse did.
Frazaro reads what you meant, in sentences you wrote on purpose.

**It is not unlimited English.** The bet is the opposite one:

> A small, checked English is better than an unlimited, guessed one.

**It is not finished.** This is a beta, released unfinished on purpose,
because the only way to learn which sentences people reach for is to let
people reach for them. The present limits, stated plainly:

- Windows desktop Excel is the platform. Excel for Mac runs the interpreted
  path and cannot compile to VBA or read Word documents. Excel for the web
  has no VBA and cannot run the add-in.
- Grammar coverage grows every week and is not yet complete for any one
  profession's vocabulary.
- PDF files are refused by name. Open the PDF in Word, save it as `.docx`,
  read the result, then load that.
- It is an Excel macro add-in, so an organization's macro policy decides
  whether it may run.

## 6. What is in the box

| Thing | What it is for | Taught in |
|---|---|---|
| The **Frazaro** tab on the ribbon | Load, validate, run, trace, undo, and look up what can be said. | Level 1 |
| Nine sample procedures and one sample workbook | An expense run, a sales cleanup, a weekly summary, an audit with a loop, receivables aging, pivot tables, a month-end close, a purchase order, a staffing policy. | Level 1 |
| Two editions | `Frazaro_English.xlam` and `Frazaro_Espanol.xlam`. | Levels 1 and 2 |
| Phrasebooks | The files that define which sentences exist. You can write your own. | Level 2 |
| VLA | The middle language, and its macro system. | Level 3 |
| Worksheet functions `=SQL()`, `=DATALOG()`, `=PROLOG()`, `=OPTIMIZE()` | Questions asked of your Excel Tables, answered as live formulas. | Levels 1 and 4 |
| The CLI | A console that takes a sentence or a form and keeps a transcript. | Level 3 |

## 7. The six levels

The titles are the ones on the website. The last level is numbered `0b101`,
which is five, written the way its intended reader counts.

| Level | Manual | Written for | You finish able to |
|---|---|---|---|
| 0 | **Introduction: Natural Language Macros** (this page) | Everyone | Decide what to read next. |
| 1 | [**Apprentice: A Human's Guide to Programming in English**](1-apprentice.md) | The person who owns a procedure | Write, validate, run and undo a procedure of your own, and read a refusal without alarm. |
| 2 | [**Journeyman: A Grammarian's Guide to Sentence Templates**](2-journeyman.md) | The person who keeps a team's vocabulary | Teach Frazaro a new sentence, with its proof, in a phrasebook of your own. |
| 3 | [**Master: A Hacker's Guide to VLA**](3-master.md) | Developers, and the VBA-literate | Read and write the middle language, write macros, and write macros that write macros. |
| 4 | [**Wizard: A Sorcerer's Guide to Computational Grammar**](4-wizard.md) | Language designers, linguists, architects | Explain why the grammar is shaped as it is, and design a new section of it. |
| 0b101 | [**Reference: An LLM's Guide to Frazaro**](5-reference.md) | A language model, and the person pasting this into one | Draft sentences that pass validation, and know what to do when they do not. |

**By role.** An analyst needs Level 1 and nothing else. A team lead who
wants the team's own verbs needs Level 2 as well. A developer reviewing the
generated code, or extending the engine, reads Level 3. Level 4 is for the
reader who wants the argument rather than the instructions. Level 0b101 is a
reference, and is meant to be consulted rather than read through.

Each level assumes the ones before it and says so at the top.

## 8. For the person deciding: how hard is this to adopt?

A fair question, and the one these manuals are most often opened to answer.

**What a new user has to learn.** Six buttons. Four conventions: a sentence
ends with a period, text goes in double quotes, cell references are written
bare as a formula writes them, and a block ends at a blank line. How to read
the note beside a red row. The vocabulary itself does not have to be
memorized, because **What can I say?** lists every sentence shape currently
loaded, each with a worked example. Level 1 covers all of it; plan on an
afternoon with the sample workbook open.

**What a new user does not have to learn.** VBA. Parentheses. The Visual
Basic editor. No message Frazaro shows will send a user there, and a check
script refuses any message that tries.

**What IT has to decide.** Whether the macro policy allows the add-in. The
default way of running, *Interpret*, needs no access to the VBA project and
writes no code into a workbook. Only *Compile* needs Excel's *Trust access
to the VBA project object model* setting, and it refuses in words when the
setting is off. [IT_REVIEW.md](../IT_REVIEW.md) is the two-page summary:
what is installed, what is written and where, what can be reached, how it is
removed.

**What can go wrong, and what catches it.**

| If | Then |
|---|---|
| A sentence is misspelled or not in the grammar | Validation marks the row red, in words. Nothing runs. |
| A sentence is valid but fails while running (a sheet is missing, say) | The run stops, names the sentence and its row, and puts the sheets back as they were. |
| The result is not what was wanted | Click **Undo Last Run**. |
| A sentence runs and silently does the wrong thing | That is a bug, not an intended failure. [SUPPORT.md](../SUPPORT.md) says where to send it. |

**What is known to be open.** The project publishes its own threat model and
its own list of security items not yet closed, in plain words, in the
[README](../../README.md#known-open-security-items). The one line to carry
away: load phrasebooks only from people you would accept a macro-enabled
workbook from.

| A buyer's question | Where it is answered |
|---|---|
| How does it compare with Copilot, Office Scripts, Python in Excel? | [COMPARISON.md](../COMPARISON.md), one page, with what each does better. |
| What does it install, and how is it removed? | [IT_REVIEW.md](../IT_REVIEW.md), [DEPLOY.md](../DEPLOY.md) |
| What can a program or a phrasebook reach? | [THREAT_MODEL.md](../THREAT_MODEL.md) |
| Who owns what it generates? | You do. [OUTPUT-EXCEPTION.md](../../OUTPUT-EXCEPTION.md) |
| What is the licence? | Apache-2.0 for the engine, MPL-2.0 for the phrasebooks, 0BSD for the samples and the one module copied into a workbook, CC-BY-4.0 for documents such as this one. |
| Where do I report a problem, and what happens next? | [SUPPORT.md](../SUPPORT.md): `english@spreadsheet.company`, acknowledged within five business days. |
| Can we try it on one real procedure? | Yes, and that is the invitation: one person, one procedure, one workbook. The address above reaches a person. |

## 9. Glossary

**Backend.** One of the two ways a program runs: the interpreter or the
compiler. They are held to agreement by running the same programs through
both.

**Block.** A sentence ending in a colon, and the sentences under it. A blank
line closes every open block. `Done.` closes just the innermost one.

**Compile.** Turn a program into a VBA module and run that. For the auditor
who wants the artefact.

**Edition.** A build of the add-in with its phrasebooks embedded: English,
or Spanish layered on English.

**Hole** (also **slot**). The typed gap in a sentence shape. `{r:cell}` is a
hole named `r` that must look like a cell.

**Interpret.** Run a program directly on the workbook, writing no code. The
default.

**Pattern.** The words and holes of a sentence shape:
`make cell {r:cell} {d:bold|italic}`.

**Phrasebook.** A text file of rules. *Frazaro* is Esperanto for
"phrasebook": *frazo*, a phrase; *-ar-*, a collection of; *-o*, the ending
of a noun. A *vortaro* is a dictionary by the same construction.

**Prelude.** VLA's standard library, written in VLA.

**Proof.** A sentence shipped beside a rule, with the exact translation it
must produce, or the refusal it must draw. The loader runs every proof each
time a phrasebook loads.

**Refusal.** Frazaro declining to read or to run a sentence, in words. The
intended way to fail.

**Rule.** A pattern, a template and, by house custom, at least one proof.

**Section tag.** `<Frazaro>` and `</Frazaro>`, each on a line of its own.
Only what sits between them is read.

**SOP.** Standard operating procedure.

**Standing decision.** A numbered, dated design decision with its reasons
attached, kept in a register that is never pruned. Cited in these manuals as
`SD-4`, `SD-13` and so on.

**Template.** What a rule produces: a VLA form with the holes' values copied
in.

**VLA.** Visual Lisp for Applications. VBA wearing parentheses.

## 10. How these manuals are written

**Where the sentences come from.** Every Frazaro sentence in these manuals
is one the phrasebook proves when it loads, or one that the regression
corpus or a sample procedure runs. Where a manual composes an example of its
own, only the values in the holes differ from a proven sentence: the cell,
the name, the number, the quoted text.

**Which version.** These manuals describe `0.7.0`. A sentence that first
works in `0.7.1` carries that mark, and the opening tag can enforce it for
you: `<Frazaro 0.7.1>` refuses to load on an older copy, by name.

**Who is right when we disagree.** The product is. **What can I say?** is
generated from the grammar that is loaded at that moment, and so it cannot
be out of date. If a manual shows a sentence that your copy refuses, the
manual is wrong or the copy is old, and [SUPPORT.md](../SUPPORT.md) would
like to hear which.

**Conventions.**

```text
A Frazaro sentence, or several, is set like this.
```

> A refusal or other message from Frazaro is set like this.

**Bold** marks a button by its real label. `Code type` marks a sentence, a
name, or a file.

---

Next: [Level 1 · Apprentice: A Human's Guide to Programming in English](1-apprentice.md)
