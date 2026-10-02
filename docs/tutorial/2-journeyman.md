# Level 2 · Journeyman: A Grammarian's Guide to Sentence Templates

*(The student becomes the teacher)*

| | |
|---|---|
| **For** | Whoever keeps a team's vocabulary: the person colleagues come to when Frazaro lacks the sentence they wanted. |
| **You need** | Level 1, a plain text editor, and a willingness to look at parentheses without touching most of them. |
| **You will be able to** | Write a sentence rule with its proof, load it, layer it over the built-in grammar, and tell a good sentence from one that will be regretted. |
| **Plan on** | A morning for the mechanics. The judgement in Part IV takes longer, and is the part worth having. |
| **Comes after** | [Level 1](1-apprentice.md) |
| **Written against** | Frazaro `0.7.0` |

---

At Level 1 you learned to speak. At this level you decide what may be said,
which is an older trade and a more dangerous one.

A grammarian used to be someone who described how people already spoke.
Here you are closer to the other kind, the one who writes the rules down
first. Frazaro gives that person unusual power, since a rule you write is
obeyed exactly, and an unusual discipline to go with it: **a rule ships with
its proof, and a phrasebook that fails its own proofs does not load.** You
cannot teach what you cannot demonstrate.

## Contents

| Part | Sections |
|---|---|
| **I. One rule** | 1 What a rule is · 2 Your first phrasebook · 3 Loading it |
| **II. Patterns** | 4 Words and filler · 5 Holes · 6 Choices, options and spellings · 7 Lists |
| **III. Templates and proofs** | 8 Templates · 9 Macros a phrasebook carries · 10 Proofs · 11 Order and overlap |
| **IV. The craft** | 12 Auditioning a sentence · 13 Promises |
| **V. Layers and languages** | 14 Layers and overrides · 15 Another language · 16 Words for values · 17 Trust |
| **VI. Practice** | Tools · Exercises · Answers · What a Journeyman can do |

---

# Part I. One rule

## 1. What a rule is

This is the rule behind `Make cell A1 bold.`, exactly as it stands in the
English phrasebook,
[`english.vla`](../../scripts/polyglotta/english.vla).

```lisp
(english-vla
    "make cell {r:cell} {d:bold|italic}"
    (make-{d} (range {r})))

(test-success
    "Make cell A1 bold."
    (make-bold (range "a1")))

(test-success
    "Make cell A1 italic."
    (make-italic (range "a1")))
```

Three parts.

| Part | Here | Says |
|---|---|---|
| **Pattern** | `"make cell {r:cell} {d:bold\|italic}"` | What the sentence looks like. Words to be matched, and holes to be filled. |
| **Template** | `(make-{d} (range {r}))` | What the sentence means, as a form in VLA, the middle language. The holes' values are copied in. |
| **Proof** | the two `test-success` forms | A sentence, and the exact translation it must produce. |

Translating a sentence is filling in a form. `{r:cell}` matched `A1`, so
`{r}` in the template becomes `"a1"`. `{d:bold|italic}` matched `bold`, so
`make-{d}` becomes `make-bold`. Nothing appears in the translation that the
sentence did not say.

You do not need to know what `make-bold` does underneath. It is a named
action that the phrasebook already provides, and Section 8 lists the ones
you are most likely to want. Writing new ones is [Level 3](3-master.md).

## 2. Your first phrasebook

A phrasebook is a plain text file whose name ends in `.vla`. Make one
called `team.vla`, anywhere you like, with this in it.

```lisp
; team.vla - Accounts Payable's own sentences.
; A line that starts with a semicolon is a note for people.

(requires-version "0.7.0")

(english-vla
    "flag cell {r:cell}"
    (begin (make-bold (range {r}))
           (set-fill-color (range {r}) vbyellow)))

(test-success
    "Flag cell C4."
    (begin (make-bold (range "c4"))
           (set-fill-color (range "c4") vbyellow)))
```

One rule, one proof. The sentence `Flag cell C4.` will make the cell bold
and fill it yellow, which until now took two sentences.

(The rules in `team.vla`, and in the exercises at the end, are this
manual's own. They are not shipped sentences. The rules of this section
and of Section 9, and the answers to exercises 2, 3 and 7, have been
loaded into a live copy of Frazaro, where every proof passed. Every other
rule quoted in this manual is copied from the shipped phrasebook.)

Reading it from the top:

- `(requires-version "0.7.0")` says which Frazaro this file was written
  for. An older copy refuses the file by name, which is kinder than letting
  it fail on the first sentence that uses something new.
- `english-vla` says: a rule whose sentences are English and whose meaning
  is VLA.
- `(begin … …)` groups two actions into one template.
- `vbyellow` is one of Excel's eight built-in colour names.
- Spaces and line breaks inside the parentheses are yours to arrange.

## 3. Loading it

1. Open a workbook and press **Load Phrasebook** on the Frazaro tab.
2. Pick `team.vla`.

Frazaro reads the file, registers its rules on top of the built-in grammar,
runs every proof, and reports:

> Loaded: team.vla (1 rule added)
>
> built-in vocabulary - 223 rules
> C:\\...\team.vla - 1 rule

The lines under the first are every phrasebook now in force, and how many
rules each supplies. The count for the built-in vocabulary is the grammar
that came with your copy: 223 in `0.7.0`.

3. Press **What can I say?**. Your rule is in the list, with *flag* as its
   category, and its worked example is the sentence from your own proof.
4. On a *Frazaro* sheet, type `Flag cell C4.` and validate. Column C is green.

**The workbook remembers.** The path of the phrasebook is kept in the
workbook, and every later command in that workbook loads it again. You do
not press **Load Phrasebook** twice.

**The computer remembers separately.** Picking the file yourself counts as
approving it on this computer. When the file's contents change, even by a
comment, the approval no longer fits and Frazaro asks again, naming the
file. Section 17 explains why this is a feature and not a nuisance.

**If a proof fails, nothing loads.**

> team.vla line 10: test FAILED
>   sentence: Flag cell C4.
>   expected: …
>   got:      …

That is the whole file refused, not the one rule. A phrasebook is either
proven or absent.

---

# Part II. Patterns

A pattern is a line of words and holes, separated by single spaces.

## 4. Words and filler

A plain word in a pattern must appear in the sentence, in that place.
Capitals are ignored on both sides.

**Four words are filler, and may not be relied on**: `the`, `a`, `an` and
`please`. They are removed from every pattern when it is registered. Write
them in a pattern if it reads better to you. They will not be there.

In a sentence, `the` and `please` are dropped wherever they stand. `a` and
`an` are passed over wherever the pattern has a word next. Where the
pattern has a hole next, they are not passed over, and that is the one
consequence that catches people: the hole swallows the `a`.

| Pattern | Trouble |
|---|---|
| `send mail a {who:expr}` | The `a` is stripped from the pattern. In a sentence, the hole then takes the writer's `a` as its value, and everything after is one place out. |

This bit the Spanish phrasebook, where *a* is a preposition. The repair was
to say *para*. If a pattern of yours needs one of the four words to mean
something, choose another word.

**Numbers as words.** The words `zero` to `twenty` are read as digits, in
patterns and in sentences alike. A pattern containing `as one list` works
because both sides read `one` the same way.

## 5. Holes

A hole is written `{name:category}`. The name is how the template refers to
it. The category says what may fill it.

| Category | Accepts | Arrives in the template as | On a misfit, Frazaro expected… |
|---|---|---|---|
| `cell` | `B2`, `Data!B2`, or anything in quotes | text: `"b2"` | a cell (like B2 or Data!B2 - quotes for a named cell) |
| `range` | `A1:C50`, `B2`, `A:C`, `1:5`, `Data!A1:B10`, or anything in quotes | text: `"a1:c50"` | a range (like A1:C50, B2, or Data!A1:B10 - quotes for a named range) |
| `column` | one to three letters | text: `"c"` | a column letter (like C or AA) |
| `sheet` | any word, or a name in quotes | text: `"data"` | a sheet name (like Data) |
| `color` | one of the eight colour words, or a code in quotes | text: `"red"` | a color (like red or yellow - quotes for a code like "#FF69B4") |
| `text` | any one word or number, or anything in quotes | text | a reference (like B2 or "Sheet1") |
| `path` | a path in quotes, or a name that holds one | text, or the name | a quoted path … or a variable name |
| `expr` | a value: a number, quoted text, a name, arithmetic in words | a VLA expression | a value (like 5, "text", or total plus 1) |
| `cond` | a condition, as after `If` | a VLA condition | a condition (like total is greater than 5) |
| `var` | one word, which becomes a name the program may use | the bare name | a name (one word, like total) |
| `name` | one word | the bare name | a name (one word, like total) |

`{r}` with no category is a `name`.

Three things about how values travel.

**A bare reference is folded to small letters.** `A1` arrives as `"a1"`.
Excel does not mind.

**A quoted reference keeps its capitals and is never second-guessed.**
`range "Q1 Totals"` is accepted as a range without any check of its shape.
Quotes are the door for named ranges and for sheet names with spaces.

**A typed hole is what makes refusals specific.** With `{r:range}`,
`Sort range banana …` is refused at validation with a sentence about
ranges. With `{r:text}` it would be accepted, and fail later, while
running, in Excel's own words. Use the narrowest category that is true.

`var` differs from `name` in one respect: a word filling a `var` hole is
registered as a name the program has given a value, so that Frazaro
declares it. Use `var` where your sentence *creates or sets* a name, as in
`set {v:var} to …`.

### A hole with a usual value

```lisp
(english-vla
    "insert row [at] {n:expr=1}"
    (insert-row-at {n}))

(test-success
    "Insert row at 5."
    (insert-row-at 5))

(test-success
    "Insert a row."
    (insert-row-at 1))
```

`{n:expr=1}` may be left out of the sentence, and the template then
receives `1`. The text after `=` is copied into the template exactly as
written.

## 6. Choices, options and spellings

Four small notations, all found in the shipped phrasebook.

| Notation | Example | Means |
|---|---|---|
| `{d:left\|right}` | `align cell\|range {r:range} {d:left\|right\|center/ed}` | One of these words, **and the word that matched travels into the template** as `{d}`. |
| `into\|in` | `put {e:expr} into\|in cell {r:cell}` | One of these words. Nothing travels. |
| `[word]` | `add [new] sheet called {s:sheet}` | This word may be present or absent. It never causes a refusal. |
| `stem/suffix` | `center/ed`, `filter/s`, `decimal/s` | Either spelling: `center` or `centered`. In a choice, the stem is what travels. |

The rule to remember is short. **Braces capture; bare words only match.**

They combine. `[in|into]` is an optional choice. `at|under [key]` accepts
*at*, *at key*, *under* and *under key*, which is how `Store` reads:

```lisp
(english-vla
    "store {e:expr} at|under [key] {k:expr} in {d:var}"
    (vladictset {d} {k} {e}))
```

Four natural spellings, one rule, one meaning.

An option is one word. A branch has at most one slash.

### Why a captured word is useful

```lisp
(english-vla
    "{d:hide|unhide} column {c:column}"
    ({d}-column {c}))
```

The matched verb completes the name of the action. `Hide column C.` becomes
`(hide-column "c")` and `Unhide column C.` becomes `(unhide-column "c")`.
One rule covers both, and the template cannot disagree with the sentence
about which was meant.

## 7. Lists

Add `-list` to a reference category, and the hole takes one or more items
separated by commas.

```lisp
(english-vla
    "add {d:rows|columns|filters} of {f:text-list} to pivot {n:text}"
    …)
(test-success
    "Add rows of Region, Product to pivot SalesPivot."
    …)
```

The categories are `text-list`, `range-list`, `cell-list`, `column-list`,
`sheet-list` and `color-list`. The template receives
`(array "region" "product")`.

**The comma before "and" is required.** `Region, Product, and Channel` is a
list of three. In `rows of Region, Product and columns of Segment` the list
is two long and stops cleanly at *Product*, because `and` without a comma
before it belongs to the sentence and not to the list. The rule is what
lets a list sit in the middle of a sentence with no guessing about where it
ends.

---

# Part III. Templates and proofs

## 8. Templates

A template is one VLA form, or several. It is written with the hole names
in braces wherever the values should go.

The good news for a Journeyman is that most sentences you will want can be
built from actions that already exist. A template names the action; it does
not reach into Excel's object model itself. The shipped phrasebook holds
this as a discipline, and it is the right one to copy.

| You want to | Use |
|---|---|
| put a value somewhere | `(set! (range {r}) {e})` |
| write a formula | `(set-formula (range {r}) {f})` |
| bold, italic | `(make-bold (range {r}))`, `(make-italic (range {r}))` |
| fill colour, font colour | `(set-fill-color (range {r}) (vlacolor {e}))`, `(set-font-color (range {r}) (vlacolor {e}))` |
| font size | `(set-font-size (range {r}) {n})` |
| alignment | `(set-alignment (range {r}) xlcenter)` |
| clear values, clear formatting | `(clear-contents (range {r}))`, `(clear-formatting (range {r}))` |
| a border | `(add-border (range {r}))` |
| column width, row height | `(set-column-width (columns {c}) {w})`, `(set-row-height {n} {h})` |
| fit, hide | `(fit-column {c})`, `(hide-column {c})`, `(hide-row {n})` |
| delete or insert a row | `(delete-row {n})`, `(insert-row-at {n})` |
| go to a sheet | `(activate-sheet {s})` |
| go to a sheet, making it if need be | `(begin (vlaensuresheet {s}) (activate-sheet {s}))` |
| sum, average, largest, smallest | `(sum-of (range {r}))`, `(average-of …)`, `(largest-of …)`, `(smallest-of …)` |
| show a message, write to the log | `(msgbox {e})`, `(debug-print {e})` |
| set the status bar | `(set-status-bar {e})` |
| do several things | `(begin … …)` |

The full list is the `defmacro` forms in
[`english.vla`](../../scripts/polyglotta/english.vla), each with a line of
description. **Export Expanded Phrasebook** writes them all out, generated
ones included.

Two details.

**A cell reference in a template is wrapped**: `(range {r})`. The hole
supplies the address as text; `range` turns it into the place.

**A colour value goes through `vlacolor`**: `(vlacolor {e})` turns
`"#FF69B4"` into what Excel wants. A captured colour word uses Excel's own
constant instead, by gluing: `vb{d}` becomes `vbred`.

## 9. Macros a phrasebook carries

When the action you need does not exist, a phrasebook may define it. This
pair gives the captured word somewhere to land.

```lisp
(defmacro
    (mark-paid r)
    "show a range as paid: green fill, bold"
    (begin (set-fill-color r vbgreen)
           (make-bold r)))

(defmacro
    (mark-unpaid r)
    "show a range as unpaid: red fill"
    (set-fill-color r vbred))

(english-vla
    "mark cell|range {r:range} as {d:paid|unpaid}"
    (mark-{d} (range {r})))

(test-success
    "Mark range A2:F2 as paid."
    (mark-paid (range "a2:f2")))

(test-success
    "Mark cell A2 as unpaid."
    (mark-unpaid (range "a2")))

(test-fail
    "Mark range A2:F2 as overdue."
    "one of 'paid'/'unpaid'")
```

A `defmacro` is a named template: `(mark-paid r)` stands for the form that
follows it, with `r` replaced by whatever was passed. The line in quotes is
its description, which **What can I say?** and the author's tools show.

Macro names are shared across every loaded phrasebook. A second macro with
a name already taken is refused, naming the file that has it. Do not begin
a name with `vla-`; that prefix belongs to names Frazaro generates.

Level 3 is about macros. For this level it is enough to know that a
phrasebook can carry them, and that they are checked when the file loads.

## 10. Proofs

```lisp
(test-success
    "Mark range A2:F2 as paid."
    (mark-paid (range "a2:f2")))
```

A `test-success` is a sentence and the translation it must produce. They
are compared as written, ignoring only differences in spacing. Notice what
is being pinned: not that the sentence *works*, but that it means exactly
this.

```lisp
(test-fail
    "Mark range A2:F2 as overdue."
    "one of 'paid'/'unpaid'")
```

A `test-fail` is a sentence that must be refused, and a fragment that must
appear in the refusal. It protects two things at once: that the sentence
stays refused, and that the refusal keeps its wording. The fragment is
matched without regard to capitals or spacing.

| A proof is | |
|---|---|
| **one sentence** | The one-line forms count as one: `If total is greater than 5, show "big".` |
| **run on every load** | Not once, at authoring. Every time. |
| **all or nothing** | One failing proof refuses the file. |
| **the worked example** | The first passing sentence for a rule is what **What can I say?** shows beside it. |

**How many?** The loader does not force you to write any. The house rule
does: every rule has at least one proof, and a rule with choices has one
for each branch that reads differently. **Export Phrasebook Test Coverage**
lists the rules of a phrasebook with no proof, and those with exactly one.

**What a proof cannot tell you.** A passing proof shows that a sentence is
translated as you intended. It does not show that Excel then does the right
thing. For that, run the sentence, both ways, on a sheet you can afford to
lose.

## 11. Order and overlap

> Rules are tried in the order they were registered. The first one that
> fits the whole sentence wins. There is no second reading.

Everything in this section follows from that sentence.

### The built-in grammar goes first

Your phrasebook is loaded after the base, so the base's rules are tried
before yours. If a base rule already fits your sentence, yours never runs.

**Your proof is what tells you.** A proof compares against the translation
that actually came back. If an earlier rule claimed the sentence, the proof
fails at load, showing the translation it got instead.

### Specific before general

Within your own file, a rule with a fixed word must come before a rule with
a hole in the same place.

```lisp
(english-vla
    "put today into|in cell {r:cell}"
    (set! (range {r}) (date)))

(english-vla
    "put {e:expr} into|in cell {r:cell}"
    (set! (range {r}) {e}))
```

In this order, `Put today into cell D1.` means the date. Reversed, the
second rule's hole would take `today` first. The shipped phrasebooks have
made this mistake and recorded it, which is how it comes to be in a manual.

### A longer rule is not shadowed by a shorter one

A rule must fit the *whole* sentence, up to the period. `delete row
{n:expr}` does not swallow `Delete row 1 of table Sales.`, because after the
number it expects the end of the sentence and finds `of`. It fails there,
and the longer rule gets its turn.

### Two rules with the same shape are refused

The loader expands every choice, every option and every spelling of every
rule, and refuses a file in which two rules can present the same shape.

> pattern '…' duplicates '…' (same shape: …) - the earlier rule always wins
>   If this replacement is intentional, use this rule's own
>   <lingua>-vla-override directive instead of <lingua>-vla.

This check is possible only because patterns are finite. It is one of the
reasons they are kept that way.

### Words a value would eat

An `expr` hole reads as much as it can. These words are part of a value, so
a pattern may not rely on them directly after an `expr` or `cond` hole:

> `plus` · `minus` · `times` · `divided` · `multiplied` · `joined` ·
> `followed`

and, after a `cond` hole, `and` and `or`. A pattern such as
`raise {e:expr} plus tax` can never match, since `plus tax` is read as the
rest of the value. Choose another word: `with`, `by`, `to`, `as`, `in`.

---

# Part IV. The craft

The mechanics took a morning. This part is the trade.

## 12. Auditioning a sentence

Every sentence in the shipped phrasebook was auditioned against the person
who would type it. The checklist below is the project's own, condensed.

### Did somebody actually say it?

A sentence should come from a transcript, not from imagination. The best
source is the log behind **Copy Diagnostic Report**: the sentences your
colleagues typed that Frazaro could not read. They are, by construction,
the sentences people reach for.

The project holds itself to this as a standing decision (`SD-7`): no section
of grammar is scheduled without a real sentence that needs it.

### One verb, one mechanism

`Put X into Y` once worked for names as well as cells. It was split, on
principle: **Set** names values; **Put** places values in locations. Later,
`Remember … under … in …` stood beside `Remember … as …`, and was split
again: remembering binds a name, storing fills a container, and lookups, as
the record has it, lack the sentience to remember. `Store` took the
container.

If your new sentence gives an existing verb a second job, find another
verb.

### Where there are two readings, the sentence names one specifically

Read your sentence to someone who has not seen the rule, and ask what they
expect Excel to do. If a neighbouring operation is a plausible answer, the
word that decides between them is missing.

| Undecided | Decided |
|---|---|
| `Add border to range …` | `Add a border around …` · `Add borders to every cell in …` |
| `Sort range … by column B.` | `… with a header row.` · `… without a header row.` |
| `Capitalize each word in …` | `… after any space.` · `… after any non-letter.` |
| `Split column C by ",".` | `… as text.` · `… reading numbers.` (both arrive in 0.7.1) |

An undecided sentence is worse than a missing one. It runs, succeeds, and
does the other thing, with no refusal to warn anybody. The project's name
for the rule is `SD-19`. A sentence with no neighbour may stay short:
specificity is owed where a second reading exists, not everywhere.

### Test the side effect, not only the words

A drafted sentence, *Set border-color of range …*, said recolour. The
property it used also drew every line it coloured. The sentence was
withdrawn before it shipped. A sentence can break the rule above through
what it does as well as through what it says.

### Refuse before you guess

`Increase total by 10%.` once added 0.1, which is what the arithmetic said
and what nobody meant. The repair gave the plain share its everyday meaning
and made the blended case refuse. When two readings survive every attempt
to name one, a refusal with directions is the honest sentence.

### Read the refusals too

Type three wrong versions of your sentence and read what Frazaro says about
each. The refusal is the most-read page of the manual, and you have just
written part of it. If the refusal steers the writer towards the wrong
rule, reorder or reword.

## 13. Promises

> Every spelling shipped is a promise made.

A sentence your colleagues have put into their procedures will be run again
next month by people who have forgotten it is there. If you change what it
means, their procedure changes with it, silently.

The project's rule (`SD-4`) is that a shipped spelling keeps its meaning,
and may be retired only through a refusal that says what to write instead.
[GRAMMAR_SINCE.md](../GRAMMAR_SINCE.md) records the release in which each
sentence first worked, and the record is only ever added to.

For your own phrasebook, the practical form is this.

| When you want to | Do |
|---|---|
| improve a sentence's wording | Add the better spelling beside the old one. Two patterns may share a template. |
| narrow a vague sentence | Add the explicit sentence, and use it from now on. Leave the vague one working. |
| withdraw a sentence | Think twice. Every procedure that uses it will stop validating on the day the rule goes. Tell the people who use it what to write instead before you remove it, and leave a note in the phrasebook where the rule stood. |
| change what a sentence does | Do not. Write a new sentence. |

The project's own policy for the shipped grammar is that a retired
sentence refuses with directions, *that spelling was retired; write this
instead*, so that a retirement reads as a migration and not as an
abandonment. No shipped sentence has been retired yet.

---

# Part V. Layers and languages

## 14. Layers and overrides

```text
   the built-in rules ............ If, Repeat, Set, Show ... part of the engine
   the base phrasebook ........... english.vla, embedded in the add-in
   the edition's overlay ......... espanol.vla, in the Spanish edition
   phrasebooks you load .......... team.vla, and as many more as you like
```

Each layer is registered after the one above it and may use the macros the
earlier layers define.

### Replacing a rule on purpose

To give an existing sentence a different meaning, say so.

```lisp
(english-vla-override
    "fit all columns"
    (begin (fit-all-columns)
           (freeze-top-row)))

(test-success
    "Fit all columns."
    (begin (fit-all-columns)
           (freeze-top-row)))
```

An override must be earned. It has to match **exactly one** earlier rule.

| The override matches | Frazaro says |
|---|---|
| no earlier rule | the override '…' matches no earlier rule - use the plain (non-override) directive, or check that the rule it replaces still exists (and loads first) |
| several | the override '…' is ambiguous - it matches … earlier rules (…) and an override replaces exactly one |
| a built-in rule | the override '…' matches the built-in rule '…' - the built-in core is not overridable; write a differently-worded rule instead |

The replacement takes the old rule's place in the order, so nothing else
about the grammar moves.

Read Section 13 again before you write one. An override changes what a
shipped sentence means for everybody who loads your file.

### Adding, and replacing wholesale

**Load Phrasebook adds.** That is the ordinary case and the one to prefer.

A file with the edition's own name (`english.vla` for the English edition)
placed beside a workbook, or in a `scripts` folder beside it, **replaces**
the built-in grammar for that workbook. Frazaro asks before using it, once
per computer, in these words:

> Load grammar rules from this file? …
> This file sits next to the workbook you are working in, and Frazaro would
> use it INSTEAD of its own built-in grammar.
> A phrasebook defines what your sentences MEAN, so a different one can
> silently change what a program does. Only load one you trust.

This mechanism exists for organizations that maintain a complete grammar of
their own. If you want to add sentences, you do not want it.

## 15. Another language

> English is not the product. English is the first phrasebook.

A rule's first word names the language of its sentences: `english-vla`,
`espanol-vla`, `latin-vla`, `pirate-vla`. Any word ending in `-vla` is
accepted, and one file may hold several. The meaning side is the same VLA
whatever the language, so a sentence in any of them inherits both ways of
running, Undo, tracing and the tests.

```lisp
(espanol-vla
    "pon hoy en la celda {r:cell}"
    (set! (range {r}) (date)))

(test-success
    "Pon hoy en la celda D1."
    (set! (range "d1") (date)))
```

```text
Pon hoy en la celda D1.              espanol.vla
Metu hodiau en la chelon D1.         esperanto.vla
Pone hodie in cellula D1.            latin.vla
Mark today upon the cell D1, savvy.  pirate.vla
```

### The small words

`If`, `Repeat`, `While` and their kin are not phrasebook rules. They are
built into the engine, in English. A phrasebook gives them another spelling
with `keyword-alias`:

```lisp
(keyword-alias "si" "if")
(keyword-alias "es" "is")
(keyword-alias "repetir" "repeat")
(keyword-alias "veces" "times")
(keyword-alias "mientras" "while")

(test-success
    "Si 5 es 5, pon 1 en la barra de estado."
    (if (= 5 5) (then (set-status-bar 1))))
```

### What is honest to say about this today

The mechanism is proven. The coverage is young, and the manual would be
lying by omission if it left out where.

| | |
|---|---|
| **The Spanish edition** is loaded and checked by the project's self-tests on every run. | It holds nineteen sentence patterns and five control words so far, layered over the whole English grammar. |
| **The other demonstration phrasebooks** (Esperanto, French, German, Danish, Latin, pirate) show the seam. | They are demonstrations, and say so at the top of each file. |
| **An alias is applied to a word wherever it stands**, not only where a control word is expected. | A word that is also ordinary vocabulary in your language cannot be an alias. Spanish *para* and *en* could not be given to `for` and `in`, for that reason. |
| **The four filler words are English.** | See Section 4. |
| **Captured words bind straight into action names.** | `{d:bold\|italic}` stays in English unless you provide actions with names in your language for it to land on. |
| **Accents are left out** of the shipped phrasebooks. | A practical choice for text files that pass through VBA. It is not a claim about spelling. |

A language is not translated by translating its words. Word order, case
and agreement arrive with the second phrasebook, and the project's notes on
what they cost are frank. [Level 4](4-wizard.md) takes this up.

## 16. Words for values

A rule adds a sentence. A *function word* adds a word that can stand inside
any value.

```lisp
(english-function
    "largest of"
    largest-of)
```

After that, `largest of results` is a value wherever a value may go.

| Written | Kind | Used as |
|---|---|---|
| a word followed by `of`, as in `"largest of"` | takes one value | `largest of results` |
| a word alone | takes nothing | the bare word, in the way the built-in `today` and `now` are used |
| several words, from `0.8.0`, as in `"median of range {r:range}"` | a *phrase*: takes what its hole says | `median of range B2:B50` |

The target is a bare name, never quoted: it names an action. Point it at a
macro the phrasebook carries, as the shipped ones do, so that both ways of
running know it.

From `0.8.0` a function word can be a phrase of several words. Its words
are fixed; then comes one hole, `{x:value}` for a value, or `{r:range}`,
`{c:column}` or `{c:cell}` for a reference written out; then, if it needs
one, a closing clause of fixed words and one choice. The target is then a
template, as a rule's is:

```lisp
(english-function
    "standard deviation of {x:value} as {k:sample|population}"
    ({k}-standard-deviation-of {x}))
```

`Set spread to standard deviation of revenues as a sample.` then reads
`revenues` into `{x}` and `sample` into `{k}`. Several words with no hole
take a value after them, so a phrasebook in another language writes its
own connector, and its own articles: `espanol.vla`'s `(espanol-function
"la suma de" suma-de)` reads `Pon la suma de ventas en la celda B14.` A
phrase is tried
before any shorter reading of its first word, and once its words before
the hole have matched, it completes or is refused, saying what it
expected. Loading refuses a phrase whose later words already follow a
value somewhere (`plus`, `is`, `and` and the like), since a sentence that
reads today would read differently, and two phrases with the same words.
A program's own `To largest of amounts:` takes the phrasebook's `largest
of` inside that program; Check notes it in yellow.

`count of` is built into the engine and answers for lists as well as
ranges. Do not redefine it.

## 17. Trust

A phrasebook decides what sentences mean. Whoever supplies the phrasebook
decides what the program does. Frazaro treats phrasebooks accordingly, and
you should too.

| Situation | What Frazaro does |
|---|---|
| A phrasebook embedded in the add-in | Trusted. It was audited when the add-in was built. |
| A phrasebook you pick with **Load Phrasebook** | Your picking it is the approval. Recorded on this computer, tied to the file's path and its exact contents. |
| A phrasebook a workbook remembers, on a computer that has not approved it | Asks first, naming the file. It does not open, read or probe the file before you answer. |
| The file changes after approval | Asks again, and says that the file has changed since you last answered for it. |
| You say no | Remembered. The file is skipped quietly from then on. |
| You change your mind | Pick the file with **Load Phrasebook**. Choosing it yourself is a yes. For a whole replacement grammar, use **Forget Phrasebook Approvals** and answer the question again. |
| A path on a network or the web | The question says so, and that Frazaro has not contacted it. |

**Forget Phrasebook Approvals**, on the Frazaro tab, lists every answer
this computer has given and clears them, so that the questions are asked
afresh. It deletes nothing from your workbooks and changes no phrasebook
file.

### Raw rules

A template may contain `(raw "…")`, which places literal VBA into a
compiled program. There is no limit on what such a rule can do, so loading
a file that contains one opens a dialog first:

> '…' wants to load a rule that runs unrestricted VBA - file access, other
> applications, anything VBA itself can do, not just Excel actions.
>
> Only allow this if you trust where this phrasebook came from.

Declining refuses the whole phrasebook. A `raw` form has meaning only when
compiling, so an interpreted run has no use for it.

You will rarely need one. If you think you do, the question to ask first is
whether the action belongs in the engine, and the address in
[SUPPORT.md](../SUPPORT.md) is where to ask it.

### For counsel

A phrasebook you write in your own file is yours, under any terms you
choose, including when it overrides a base rule or uses the base's macros.
[PHRASEBOOK-TERMS.md](../../PHRASEBOOK-TERMS.md) is four sentences and a
table, written for the person who will ask.

### The one line

> Load phrasebooks only from people you would accept a macro-enabled
> workbook from.

---

# Part VI. Practice

## Tools

| Tool | Where | Tells you |
|---|---|---|
| **What can I say?** | Frazaro tab | Every sentence shape loaded now, with a worked example. Your rules included. |
| **Load Phrasebook** | Frazaro tab | Whether the file loaded, how many rules it added, and every source now in force. |
| **Export Phrasebook Test Coverage** | Frazaro tab, Auditing | Which rules have no proof, and which have exactly one. |
| **Export Expanded Phrasebook** | Frazaro tab, Auditing | The phrasebook as the loader sees it, every generated rule written out. |
| **Lint VLA File** | Frazaro tab, Auditing | Reformats a `.vla` file in house style. It refuses to write if the change would be more than spacing. |
| **Open CLI** | Frazaro tab, Utilities | Try a sentence at once, on the workbook in front of you. |
| **Copy Diagnostic Report** | Frazaro tab, Utilities | What your colleagues typed that Frazaro could not read. Your backlog. |

If you know VBA, [IMMEDIATE.md](../IMMEDIATE.md) lists commands that
explain how one sentence was read, word by word, and which file's rule
claimed it.

## Exercises

**1.** Write a rule for `Highlight cell B2.` that fills the cell yellow.
Give it a proof.

**2.** Extend it so that the writer may say `cell` or `range`, and prove
both.

**3.** Write a rule for `Set up row 1 as a title.` that makes row 1 a
header row and freezes it. The shipped sentence `Make row 1 a header row.`
uses the template `(make-header-row {r})`, and `Freeze top row.` uses
`(freeze-top-row)`. Which category should the row's hole have?

**4.** This pattern can never match. Why?

```text
charge {e:expr} plus handling to cell {r:cell}
```

**5.** A colleague wants `Clean column A.` to remove extra spaces. Audition
the sentence against Section 12. What is wrong with it, and what would you
offer instead?

**6.** Your phrasebook has shipped `Flag cell {r:cell}` for a year. The
team now wants flagged cells to be orange. What do you do?

**7.** Write a `test-fail` for exercise 2 that pins the refusal of
`Highlight column B.`

## Answers

**1.**

```lisp
(english-vla
    "highlight cell {r:cell}"
    (set-fill-color (range {r}) vbyellow))

(test-success
    "Highlight cell B2."
    (set-fill-color (range "b2") vbyellow))
```

**2.** A `range` hole accepts a single cell, so one hole serves both.

```lisp
(english-vla
    "highlight cell|range {r:range}"
    (set-fill-color (range {r}) vbyellow))

(test-success
    "Highlight cell B2."
    (set-fill-color (range "b2") vbyellow))

(test-success
    "Highlight range A1:C1."
    (set-fill-color (range "a1:c1") vbyellow))
```

**3.** `expr`, as in the shipped rule. A row is a number, and may be worked
out.

```lisp
(english-vla
    "set up row {r:expr} as title"
    (begin (make-header-row {r}) (freeze-top-row)))
    
(test-success
    "Set up row 1 as a title."
    (begin (make-header-row 1) (freeze-top-row)))
```

The `a` in the sentence is filler and is not in the pattern. Note that
`set` is a busy first word: the built-in `set {v:var} to {e:expr}` is tried
first, fails at `row`, where it expected `to`, and yours gets its turn. The
proof confirms it.

One thing the proof does not confirm: `(freeze-top-row)` always freezes
row 1, whichever row the sentence named. `Set up row 3 as a title.` would
be translated without complaint and would then do something the sentence
did not say. Either fix the row in the pattern, `set up row 1 as title`, or
use an action that freezes the row given. This is Section 12's *test the
side effect* in miniature.

**4.** `plus` directly follows an `expr` hole. The value takes
`plus handling` as the rest of itself, and the word `plus` in the pattern
is never reached. Write `charge {e:expr} with handling to cell {r:cell}`.

**5.** *Clean* has at least four plausible readings: remove extra spaces,
remove non-printing characters, clear the contents, clear the formatting.
The sentence would run and do one of them. Offer the sentence that already
exists, `Remove extra spaces from column A.`, which says what it does.

**6.** Not an override of `flag`, which would turn every flagged cell in
every existing procedure orange on its next run, without a word. Add a
sentence that names the choice, `Flag cell {r:cell} in {d:yellow|orange}`,
leave the old one meaning yellow, and let procedures move over when their
owners decide to. The captured word needs a macro of each name to land on,
`flag-yellow` and `flag-orange`, as `mark-paid` and `mark-unpaid` did in
Section 9. Excel has a constant for yellow and none for orange, so the
orange one fills with `(vlacolor "#FFA500")`.

**7.**

```lisp
(test-fail
    "Highlight column B."
    "one of 'cell'/'range'")
```

## What a Journeyman can do

- [ ] Write a rule: pattern, template, proof.
- [ ] Choose the narrowest category that is true for a hole.
- [ ] Use a captured choice to complete an action's name.
- [ ] Say why the comma before "and" is required in a list.
- [ ] Order two rules so that the specific one is tried first.
- [ ] Say what the loader refuses, what it only warns about, and what only
      a proof will catch.
- [ ] Audition a sentence, and turn one down.
- [ ] Improve a shipped sentence without breaking the procedures that use
      it.
- [ ] Explain to a colleague why Frazaro asked about a phrasebook again
      after they edited it.

---

Previous: [Level 1 · Apprentice](1-apprentice.md) ·
Next: [Level 3 · Master: A Hacker's Guide to VLA](3-master.md)
