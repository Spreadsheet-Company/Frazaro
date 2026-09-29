# Level 4 · Wizard: A Sorcerer's Guide to Computational Grammar

*(What is a language, anyway?)*

| | |
|---|---|
| **For** | Language designers, linguists, logicians, and the reviewer who wants to know why the grammar can be trusted before asking whether it is convenient. |
| **You need** | Levels 2 and 3. Some acquaintance with formal grammars helps; every term is explained where it is first used. |
| **You will be able to** | Explain why the grammar is shaped as it is, say what it can and cannot promise, and design a new section of it. |
| **Plan on** | An evening to read, and then a good deal longer, since the last part is a method and not a fact. |
| **Comes after** | [Level 3](3-master.md) |
| **Written against** | Frazaro `0.7.0`. What arrives in `0.7.1` is marked. |

---

*Grammar*, *grimoire* and *glamour* are one word, three times. A medieval
*gramaire* was a book of learning, and learning that most people could not
read was assumed to be the dangerous kind; Scots turned the same word into
*glamour*, an enchantment cast over the eyes. The sorcerer's trade was
always saying exactly the right words in exactly the right order and
having something happen.

This manual is about that trade, with the enchantment taken out. Frazaro
is a machine that does things when the right words are said, and its whole
claim to seriousness is that it can tell you, in advance and in writing,
which words those are.

## The question

Max Weinreich made famous the remark that a language is a dialect with an
army and a navy. (He credited it to a member of his audience.) The joke is
that linguistics has no sharper test. What counts as a language is settled
by institutions.

For a language a machine will obey, this manual uses a working definition
with three parts:

1. **A set of sentences.** What may be said.
2. **A meaning for each.** What saying it does.
3. **An institution that keeps the two together over time.**

Most programming languages take care over the first two and leave the
third to the market. Frazaro is unusual mainly in the third. The set is
published and enumerable (**What can I say?**). The meaning of each
sentence is one form in a middle language ([Level 3](3-master.md)). And
the institution is a short register of standing decisions, of which four
carry most of this manual:

| | Decision | In one line |
|---|---|---|
| `SD-4` | A shipped spelling keeps its meaning. | The promise. |
| `SD-7` | No grammar without a real sentence that needs it. | Where sentences come from. |
| `SD-16` | The pattern language is the only sentence grammar. | What the parser may be. |
| `SD-19` | Where Excel has neighbouring operations, the sentence names the one it performs. | What the writer is owed. |

The register is in [BETA_REARVIEW.md](../BETA_REARVIEW.md). It is
append-only, and it records its own failures. `SD-2` carries a note that it
was found not to be honoured in practice, which is left standing beside
the repair.

## Contents

| Part | Sections |
|---|---|
| **I. The shelf** | 1 Controlled languages · 2 What was taken, and what declined |
| **II. The reader, stage by stage** | 3 Characters · 4 Words · 5 Dispatch · 6 A rule · 7 The built-in grammars |
| **III. Why it is shaped this way** | 8 Decidability as a feature · 9 Two kinds of ambiguity · 10 The refusal · 11 Loud on the dangerous branch |
| **IV. Three moods** | 12 Commands, statements, questions · 13 Nouns as variables · 14 Constants · 15 Names a reader can redo by hand · 16 Three small words |
| **V. Four engines** | 17 Four languages · 18 Restriction as the feature |
| **VI. Evidence** | 19 Proofs about the prover |
| **VII. Time** | 20 Permanence |
| **VIII. Other tongues** | 21 What a phrasebook translates |
| **IX. Method** | 22 Designing a section |
| **X. Open questions** | 23 What is not settled |
| | Seminar · What a Wizard can do |

---

# Part I. The shelf

## 1. Controlled languages

A *controlled natural language* is a subset of a natural language,
restricted in vocabulary and grammar so that something can be guaranteed
about it. There are two old families.

| Family | Restricts so that | Example |
|---|---|---|
| **For people** | a reader, often reading in a second language, cannot misunderstand | ASD-STE100, Simplified Technical English, in which aircraft maintenance manuals are written |
| **For machines** | a program can take the text as a specification | Attempto Controlled English, which maps sentences to logic; Robert Kowalski's Logical English, which reads as English and runs as a logic program |

Tobias Kuhn's survey of the field classifies such languages on four
scales: **p**recision, **e**xpressiveness, **n**aturalness and
**s**implicity. They trade against each other, and a design is mostly a
choice of where to stand.

| Scale | Frazaro's position | Evidence |
|---|---|---|
| **Precision** | As high as it goes. Every sentence has exactly one reading. | The loader proves that no two rules share a shape, and refuses a phrasebook where they do. |
| **Expressiveness** | Modest, and growing by demand. | 223 sentence shapes at `0.7.0`, each traceable to a sentence somebody wanted. |
| **Naturalness** | Sentences read as English. They are not free English. | `Put today into cell D1.` A writer learns which sentences exist. |
| **Simplicity** | High. The grammar is a list. | The complete description of the pattern language is Part II of [Level 2](2-journeyman.md). |

The position is a bet, stated in the README: *a small, checked English is
better than an unlimited, guessed one.*

## 2. What was taken, and what declined

| From | Taken | Declined |
|---|---|---|
| **Definite clause grammars** (Pereira and Warren, 1980) | The shape of a rule: a production from words to meaning, written as data. | Backtracking, and nonterminals a rule-writer can define. |
| **Parsing expression grammars** (Ford, 2004) | Ordered choice. Alternatives are tried in order and the first success is final, so the grammar cannot be ambiguous. | Recursion. |
| **Montague** (*English as a Formal Language*, 1970) | The conviction that a fragment of English can be given a meaning as exact as a formal language's. | The ambition to cover English. |
| **Logical English** | The flat clause: a rule is a head, `if`, and a list of conditions joined by `and`. | Pronouns and relative clauses. It was weighed and declined, for reasons Section 8 makes plain. |
| **Lisp** | The middle language, and the habit of writing the language in itself. | Closures, `gensym`, and being nicer than the host. |
| **Datalog** | Questions that always terminate. | Nothing. It is used whole. |

`SD-16` names the result exactly:

> a Definite Clause Grammar with a cut after every clause, on purpose.

A *cut*, in Prolog, discards the alternatives not yet tried. A grammar
with a cut after every clause commits to the first rule that fits and
never reconsiders. That is a severe restriction, and Part III is about
what it buys.

---

# Part II. The reader, stage by stage

A sentence passes through five stages. Each is deterministic, and each
can refuse.

```text
   characters --> tokens --> words as read --> dispatch --> one rule --> a form
       3            3             4               5            6, 7
```

## 3. Characters

The alphabet of a bare word is `A` to `Z`, the digits, the hyphen and the
underscore. Inside double quotes, anything.

Four punctuation marks have two jobs each, and the reader tells them apart
by **what stands next to them**, which is always decidable on the spot.

| Mark | Glued | Free-standing |
|---|---|---|
| `:` | Between word characters, a range: `A1:B10` | Followed by space or the end of the line, it opens a block. |
| `.` | Between digits, a decimal: `3.14` | Otherwise, the end of the sentence. |
| `,` | Digit, comma, exactly three digits, a thousands separator: `1,000,000` | Otherwise, a separator. |
| `!` | After a sheet name, part of a reference: `Data!B2` | Otherwise, the end of the sentence. |

The project calls this *the quoting liberation*: the grammar began with
everything in quotation marks, and each of these rules freed one kind of
reference. What is left fits in half a sentence. *Quote it only if it has
a space in it.*

A character outside the alphabet is refused, with a hint for the ones
people type by habit:

> I don't understand the character '+' … write 'plus' for addition

There are no symbolic operators in the sentence language at all. A
sentence is words.

## 4. Words

Four things happen to a bare word before any rule sees it.

| What | Rule | Why |
|---|---|---|
| **Capitals are folded** | `A` to `Z` only, by code point. | A name must be the same name on every computer. A locale-aware fold is not: under Turkish rules the capital of `i` is dotted. `SD-8`. |
| **`the` and `please` are dropped** | Everywhere. | They can never be a reference. |
| **`a` and `an` are kept** | They are passed over only where a grammar word is expected. | They *can* be a reference. There is a column A. |
| **Number words become digits** | `zero` to `twenty`. | Beyond twenty, write digits. A pattern that contains one of these words is read the same way, so that `as one list` in a rule still meets `as one list` in a sentence. |

A phrasebook may also declare a **keyword alias** (Level 2), which rewrites
one word as another at this stage.

Quoted text passes through all of this untouched. It keeps its capitals,
its articles, and its `+`.

## 5. Dispatch

The first word chooses the path.

**Structural sentences** are built into the engine: `If`, `Otherwise`,
`When`, `Try`, `Repeat`, `While`, `Count`, `For each`, `To`, `Give back`,
`Define`, `Create`, `Use`, `Stop`, `Done`. Most of them own their first
word outright. `Create` is one: a sentence that begins with it is always a
declaration, and a phrasebook rule beginning with it would never be
reached. That is why the sentence for a pivot table begins with `Make`.

**Every other sentence** goes to the phrase rules, which are tried in the
order they were registered: the built-in few first, then the base
phrasebook, then each layer loaded on top.

## 6. A rule

The matcher's contract is short enough to state whole.

> Rules are tried in registration order. Within a rule, items are matched
> left to right. A word must be that word. A hole consumes one token, one
> quoted string, one list, or one phrase of a built-in grammar. There is
> no lookahead and no backtracking inside a rule. The first rule that
> reaches the end of the sentence wins. There is no second reading.

If a rule fails at its fourth item, the matcher does not retry the third
item differently. It abandons the rule and tries the next. If every rule
fails, the one that got furthest supplies the refusal.

## 7. The built-in grammars

Five holes consume more than one token. Each is parsed by a small grammar
fixed in the engine. A phrasebook can use them and cannot define one.

### `expr`, a value

```text
Expr    := Joined
Joined  := Sum     ( ("joined with" | "followed by")               Sum     )*
Sum     := Product ( ("plus" | "minus")                            Product )*
Product := Postfix ( ("times" | "multiplied by" | "divided by")    Postfix )*
Postfix := Primary ( "%" | "percent" )?
Primary := number | "text" | name
         | ordinal                           third, read as 3
         | word "of" Primary                 length of code, first of results
         | "item" Expr "of" name
         | name "for" Expr                   prices for code
         | "cell" reference | "value in cell" reference
         | "cell in column" letter "row" Expr
         | name "using" name "of" Expr ( "and" name "of" Expr )*
```

Three levels of precedence, left to right within each, and **no
parentheses**. A value too tangled to say in one breath is said in two
sentences, with a name for the part. That is a restriction on the writer.
It is also how an auditor comes to find every intermediate figure named.

Two behaviours are worth knowing.

*An operator word with no value after it is handed back to the sentence.*
That is how `Repeat 10 times:` keeps its `times`.

*`row` is greedy.* `cell in column C row 2 plus 1` is row 3. A test once
pinned the other reading, and the golden files corrected it: the greedy
reading is what a speaker means by "row two plus one". To add to the
cell's value, name the value first.

### `cond`, a condition

```text
Cond    := Simple ( ("and" | "or") Simple )*
Simple  := Expr Test Expr
         | Expr ("is empty" | "is not empty")
Test    := "is" | "equals" | "is equal to" | "is not" | "does not equal"
         | "is greater than" | "is more than" | "is less than"
         | "is at least" | "is at most"
         | "is greater than or equal to" | "is less than or equal to"
         | "contains" | "does not contain" | "starts with" | "ends with"
         | "is divisible by"
```

`and` and `or` are read strictly left to right, with no precedence
between them and no parentheses. The design chose no precedence over a
precedence that half of all writers would remember the wrong way round.
A condition that mixes the two is better written as two conditions.

### `conditions`, `clause` and `question`

These three belong to the logic sentences of Part IV, and they are held
to a stricter standard than the two above: each is **regular**. The
grammar, from the engine's source:

```text
Conditions := Condition ( Sep Condition )*
Sep        := ","? "and"
Condition  := "not"? ( TableRow | Relation | Set | Comparison | Text )
TableRow   := <table> "lists" Column ( Sep Column )*
Column     := Operand "as" <header>
Relation   := Operand <rel> Operand ( "directly" "or" "not" )?
Set        := <role> "is" <set>
Comparison := <role> "is" Op ( <role> | <number> )
Text       := <role> ( "starts with" | "ends with" | "contains" ) Operand
Op         := "at least" | "at most" | "greater than" | "less than"
Operand    := <role> | <constant>

Clause     := Operand ( "is" <set> | <rel> Operand )
              ( "if" Conditions ( ","? "otherwise" Operand "if" Conditions )* )?

Question   := "whether" <constant> Predicate
            | Unknown Predicate
            | Unknown <constant> <rel>
Predicate  := "is" <set> | <rel> ( <constant> | Unknown )
Unknown    := "who" | "what" | "which" <noun>
```

(The question grammar has further shapes, for counting and for lists.
Section 12 shows them by example.)

*Regular* means a finite automaton could recognize it: no nesting, no
parentheses, no pronouns, single terms as operands. The parser needs to
look ahead in exactly two places, and each window is the width of the
shape it is recognizing.

- **Which shape a condition is** is decided by its second token. `lists`
  makes a table row, `is` a set or a comparison, anything else a relation.
- **After a column, whether a separator continues the list or begins a
  new condition.** `, and level as Level` is another column.
  `, and level is at least min` is a comparison. The parser reads one
  operand and looks for `as`.

Neither window looks past the shape it is matching, so the whole remains
one left-to-right scan.

---

# Part III. Why it is shaped this way

## 8. Decidability as a feature

The restriction of Section 6 makes three properties *provable*, where a
more generous parser could only hope for them.

### The loader can prove that two rules do not overlap

Every choice, option and spelling in a pattern can be expanded, because
there are finitely many. The loader does so for every rule, and compares.
Two rules that can present the same shape are refused at load.

For context-free grammars in general this question has no algorithm.
Whether two such grammars can ever accept the same sentence is
undecidable, and so is whether one grammar is ambiguous. A grammar
written as a DCG can only be tested for overlap. A grammar written as
finite patterns can be checked for it.

Each built-in grammar is a single opaque placeholder to this check. That
is safe because a phrasebook cannot change what the placeholder accepts.

### A refusal can say how far the reading got

> I understood 'sort range a1:c50 by column b' - then I expected one of
> 'with'/'without' but found the end of the sentence.

A deterministic scan always knows where it stopped. A backtracking
parser, on failure, has abandoned every path it tried, and knows only
that none worked. The teaching refusal of Level 1 is a consequence of the
parser's poverty.

### One sentence, one meaning

`SD-4` promises that a shipped sentence keeps its meaning. The promise
can be kept only while no sentence has a second reading for a later rule
to disturb.

### What the decision forbids

In its own words: a grammar rule written as a clause with a body; a slot
that recurses into another rule; any rule that can succeed by more than
one path; any matcher that re-enters a rule after a slot failed further
in; and any use of the Prolog engine, or the unifier beneath it, to
*recognize* a sentence.

The last clause is there because the temptation is structural. Frazaro
ships a Prolog. The cheapest way to write a new section of grammar will
one day look like a DCG, since the engine is right there. Every section
written that way would drop out of the overlap check, the refusal and the
promise, and no test would fail to say so.

> Prolog's strength at natural language is to enumerate the parses and let
> search resolve the ambiguity. That is aimed at *natural* language.
> Frazaro parses a *controlled* one, where ambiguity is a bug the loader
> refuses, and not a reading the engine explores.

This is also the wall between Frazaro and a language model. A model
accepts anything and chooses a reading. Frazaro accepts what it can prove
it reads one way.

## 9. Two kinds of ambiguity

`SD-16` makes a sentence mean one thing to the parser. That is half of
what a writer needs.

```text
Add border to range A1:C10.
```

This has exactly one parse. It draws a line round every cell in the
range. An English speaker, reading it, sees a box.

> A sentence can have exactly one parse and still be read as its
> neighbour.

`SD-19` is the second guarantee: where Excel has neighbouring operations,
the sentence carries the word that decides between them.

| Undecided | Decided |
|---|---|
| border | `around` · `to every cell in` |
| colour | `fill-color` · `font-color` |
| clear | contents · formatting |
| sort | `with a header row` · `without a header row` |
| capitalize | `after any space` · `after any non-letter` |

The two guarantees are independent, and they fail differently. A
violation of the first is caught by the loader. A violation of the second
can be caught only by reading the sentence to someone who has not seen
the rule.

An undecided sentence is the worse failure. It runs, succeeds, and does
the other thing, and no refusal warns anybody.

A sentence that has already shipped keeps its meaning and gains an
explicit sibling. A rule with no neighbour stays short. *Specificity is
owed where a second reading exists, not everywhere.*

## 10. The refusal

In most languages the error message is an afterthought. In a controlled
language it is the page of the manual that is read most, since a writer
finds the edges of the language by walking into them.

| Part | Example | Says |
|---|---|---|
| What was understood | `I understood 'put 1'` | How far you were right. |
| What was expected | `then I expected one of 'into'/'in'` | The grammar at that point. |
| What was found | `but found 'onto'` | Your word. |
| What to write | `Did you mean: 'put {e:expr} into\|in cell {r:cell}'` | The nearest rule. |
| Where | `(line 12)` | The row. |

**The voice is first person.** *I understood*, *I don't understand*. It
was chosen after trying the alternative: a tool that says "I don't
understand" invites teaching, and one that says "parse failure" assigns
blame.

**The wording is tested.** A `test-fail` proof pins a fragment of the
message. If the message drifts, the phrasebook stops loading.

**There are two kinds of failure, and they go to different places.** In
the logic grammars, a failure of *syntax* falls through to the frame
above. A failure of *sense* is raised by name: a role that no condition
binds, a relation named with one of the grammar's own words. By then the
sentence has matched and what the writer meant is known, so "I expected a
condition" would point at the wrong place.

**No message sends a user to the VBA editor.** A check script refuses any
that does.

## 11. Loud on the dangerous branch

> When an ambiguity has a dangerous branch, resolve it loud, never silent.

Three cases, each from the record.

**The thousands separator is strict.** Digit, comma, exactly three digits.
A looser rule would read the typo `1,00` as `100`. In a finance tool a
silent error of magnitude is the worst failure available, so the typo is
refused.

**The percent sign.** `Increase total by 10%.` was at first refused,
because the arithmetic reading adds 0.1. The owner argued that the idiom
is day-one vocabulary and that the refusal spent the user's goodwill on
internal bookkeeping. The idiom was accepted: a plain share means growth.
The acceptance was the easy half. The guard was tightened at the same
time, so that any `%` blended into longer arithmetic refuses, including a
case the earlier check had let through.

> The mechanism follows the amount's shape, not a second meaning of the
> verb.

**The quoted numeral.** In the logic sentences, `"3"` and `3` looked
alike and were not. One engine keeps the mark that says a cell holds
text, and the other strips it. The question `whether "Ann" has-level "3"`
answered FALSE in one and found Ann in the other, with nothing to warn
anyone. A quoted numeral is now refused by name. Writing it bare is the
one spelling both engines read alike.

In each case the permissive reading would have worked most of the time.
The design question was what happens the rest of the time.

---

# Part IV. Three moods

## 12. Commands, statements, questions

Until `0.6.0` every Frazaro sentence was a command. English has two more
moods that matter to a business, and both now exist.

| Mood | Says | Sentence |
|---|---|---|
| **Imperative** | Do this. | `Sort range A1:C50 by column B with a header row.` |
| **Declarative** | This is so. This follows from that. | `Write in cell H2 that a bill is big if …` |
| **Interrogative** | What is so? | `Show in cell E2 which bill is a violation by …` |

The second and third are still, grammatically, commands: *write*, *show*.
What they carry is a clause. `that a person can-cover a shift if …` is a
content clause, and `who can-cover "Night"` is an embedded question. The
carrier sentence says where the clause goes, and places the address
first, so that a long body never buries it.

### A policy, as written

```text
Write in cell M2 that a person can-cover a shift if Shifts lists the shift as Shift, the cert as Needs, and the min as MinLevel, and Staff lists the person as Name, the cert as Cert, and the level as Level, and the level is at least the min, and not Leave lists the person as Name and the shift as Shift.
Write in cell M3 that a shift is listed if Shifts lists the shift as Shift.
Write in cell M4 that a shift is covered if the person can-cover the shift.
```

Each sentence writes one rule, as text, into one cell. The rules are
there for anybody to read, and they are read live: edit a rule, or a row
of a table, and every answer that depends on it changes.

### The questions

Each shape is named by its own opening words.

| Opening | Answers | |
|---|---|---|
| `who` · `what` · `which <noun>` | the rows that fit | `who can-cover "Night"` |
| two unknowns | every pairing | `which person can-cover which shift` |
| `whether` | TRUE or FALSE | `whether "Bob" can-cover "Night"` |
| `how many` | one number | `how many people can-cover "Night"` |
| `how many … each … that is` | a count for each member, zero included | `how many people can-cover each shift that is listed` |
| `which … that is … is not` | what is outside | `which shift that is listed is not covered` |
| `whether every … that is … is` | that nothing is outside | `whether every shift that is listed is covered` |
| `… alone` | that exactly one fits | `who alone can-cover "Night"` |
| `… as one list` | the answers in one cell | `who can-cover "Night" as one list` |

In full:

```text
Show in cell A4 who can-cover "Night" by applying the rules in M2:M4 to the data tables Staff, Shifts, and Leave.
```

The headings of an answer are the question's own words: *Who*, *What*, or
the noun after `which`. They are never the rules' variable names. It was
measured that a direct query takes its headings from whichever rule cell
happens to come first.

## 13. Nouns as variables

> a person can-cover a shift if … the person … the shift …

There is no `X`. The **role noun** is the variable, and the same noun is
the same variable, which is how a join is said: `the person` in the
Staff row and `the person` in the Leave row are one person.

English does this work with articles. *A person* introduces somebody and
*the person* refers back. Formal semantics has a literature on the
difference, under the name of discourse referents. Frazaro declines the
whole question. The articles are set aside before the noun is read, so
`a person` and `the person` are one word to the grammar, and the noun
alone carries the identity.

The cost is that two people need two nouns.

```text
Write in cell H3 that a person is-under a boss if the person reports-to the boss directly or not.
```

The gain is that there is no pronoun to resolve and nothing to resolve it
wrongly. *It*, *they* and *who* as a relative pronoun are where
controlled languages acquire their hardest problems, and where Logical
English in full was weighed and declined.

**A relation is one word**, with hyphens joining its parts: `can-cover`,
`reports-to`, `has-tier`. The hyphen is honest about what is happening.
`can-cover` is a name the writer has coined, not the English verb phrase,
and it does not conjugate.

**The variable is the noun with its first letter raised**, and the letter
is raised by code point, not by the operating system's idea of a capital.
`SD-8` again: a locale that raises `i` to a dotted capital would mint a
different variable on a different machine.

## 14. Constants

> A constant is typed by how it was written.

| Written | Is | |
|---|---|---|
| `"Night"` | text | In quotes. It keeps its capitals. |
| `10000` | a number | Bare. |
| `"10000"` | refused | By name. Section 11. |
| `Night` | not a constant | A bare word is folded to small letters and would never match the table's `Night`. It reads as a role noun. |

**A constant stands where a role does, and there is no `=`.** `Roster
lists the person as Name and "Ops" as Dept` is how one says that the
department is Ops. Every spelling of equality was measured in both
engines and they disagreed; the column spelling was the one on which both
agreed. So `the dept is "Ops"` refuses, and the refusal teaches the
column spelling.

## 15. Names a reader can redo by hand

Some sentences need rules the writer did not write. A negated table row
needs a projection. A question needs a rule that narrows the relation to
what was asked. `directly or not` needs a closure. Each of those rules
needs a name.

Level 3 explained why VLA has no `gensym`. The same objection applies
here with more force, since these names appear in a cell that an auditor
will read. So the names are made from **content**.

| Sentence says | Generated name |
|---|---|
| `who can-cover "Night"` | `vla-ask-can-cover` |
| `not Leave lists the person as Name and the shift as Shift` | `vla-not-leave-name-shift` |
| `reports-to … directly or not` | `vla-any-reports--to` |
| `how many people can-cover …` | `vla-count-can--cover` |
| `… otherwise "Silver" if …` | `vla-first-has--tier-1` |

A generated name must pass four tests.

1. **Deterministic.** The same sentence gives the same name, whatever
   order the sentences are written in. A name made from an ordinal fails:
   two sentences in one range would give one name to two rules.
2. **Derivable with a pencil.** A reader can work out what the name will
   be, and work back from a name to what made it.
3. **Injective.** Different content gives different names. The parts are
   joined with `-`, and `-` is legal inside a word, so the headers
   `needs-by, shift` and `needs, by-shift` once gave one name to two
   different rules. A hyphen inside a part is therefore written `--`. A
   single hyphen only ever separates parts.
4. **Reserved.** Every generator owns a prefix, and nothing a writer types
   may begin with `vla-`.

> Deterministic names are safe to use where gensym would have been only
> if their namespaces can never meet each other, or meet a writer's.

The table shows one name that does not double its hyphen. `vla-ask-` is
the oldest of the generators. Its name has a single part, the relation,
which is written as it stands, and with one part there is nothing for a
hyphen to be confused with.

A constant in a negated row goes into the *call*, never into the rule.
`not Leave lists the person as Name and "Night" as Shift` calls the same
all-variable rule with `"Night"` as an argument. Written into the rule
instead, `"Night"` and `"Day"` would have been two different rules under
one name.

## 16. Three small words

Each of these is a decision about logic that was settled by measuring
both engines, and then given a word.

**`or`, by repetition.** A second rule sentence with the same head is
"or". This creates a hazard neither engine can see: a question whose
range of rules includes one of those cells and not the other answers from
part of the relation, silently. The grammar can see it, within one
program. Every rule sentence records its cell and its head, every
question records its range, and after the whole program has been read a
question whose range splits a relation is refused.

**`otherwise`, first match.**

```text
Write in cell H2 that a customer has-tier "Gold" if Customers lists the customer as Customer and the spend as Spend, and the spend is at least 10000, otherwise "Silver" if Customers lists the customer as Customer and the spend as Spend, and the spend is at least 5000, otherwise "Bronze" if Customers lists the customer as Customer.
```

Each branch holds only where every earlier branch did not, and the choice
is made per customer. Both alternatives were measured first. Prolog's
`if` chose per *row*, so a customer with two rows was Gold and Bronze.
"Or" by repetition chose every tier that held.

**`directly or not`.** One step along a relation, or more. It writes a
closure, always right-recursive, so that it terminates whichever way the
data runs.

---

# Part V. Four engines

## 17. Four languages

Frazaro ships four worksheet functions. Each is a language of its own,
with its own guarantees, and the differences are the point.

| | `=SQL()` | `=DATALOG()` | `=PROLOG()` | `=OPTIMIZE()` |
|---|---|---|---|---|
| **Asks** | Which rows? | What follows? | What can be shown, and how? | Which choice breaks no rule? |
| **Written as** | SQL text | Facts, rules and a query, as forms | Clauses and a query, as forms | Rules, choices and constraints |
| **Evaluates** | Relational operators | Bottom-up, to a fixpoint | Top-down, with unification and backtracking | Search |
| **Terminates** | Yes, except a recursive query over cyclic data, which meets a ceiling | **Always**, for rules that only relate what is there. See below. | When the rules allow; otherwise at a ceiling | Within a work budget |
| **Answers** | A table | A set, or TRUE or FALSE | Every solution, in order | An assignment, with a status |
| **Negation** | `NOT`, `EXCEPT` | Stratified | `not`, as failure | `forbid` |
| **Order matters** | No | No | Yes | The tables' own order decides the first answer |

**SQL** is a frozen, SQLite-leaning subset: `SELECT`, `DISTINCT`, inner
joins, `WHERE`, arithmetic, the aggregates with `GROUP BY` and `HAVING`,
`ORDER BY`, `LIMIT`, the set operators, and `WITH`, recursive included.
Outer joins, self-joins and subqueries in a `FROM` are refused by name.
Tables must be real Excel Tables, since a plain range has no column
names.

**DATALOG** is a logic language without function symbols, evaluated from
the facts upward until nothing new follows. The set of facts that can be
derived from the values present is finite, so it always stops, and cyclic
data is no trouble. A variable begins with a capital. An atom may address
a table's columns by header: `(staffing (name X) (salary S))`.

The engine also offers arithmetic, `(let Z (+ X Y))`, and arithmetic can
make values that were not there. A recursive rule that counts upward
without end is the one way to write a DATALOG program that does not stop
by itself. It meets a ceiling of 10,000 rounds and is refused, with the
advice to look for a rule whose recursion never narrows. The sentences of
Part IV cannot write such a rule. They have comparisons and no
arithmetic.

**PROLOG** is the general language: unification, backtracking, lists,
arithmetic, `findall`, cut. Its power is that the question can be shaped
like the policy. The price is that a rule can run for ever, and so the
engine has ceilings, on depth, on list length and on work, and refuses by
name when one is reached.

**OPTIMIZE** takes rules, *choices* and *constraints*, and searches for
an assignment that breaks none. It reports one of five states, from
refused to proven best, and a violation names the constraint and the
rows. It is the newest and is partly built. What it does not yet do, it
refuses by name: `prefer`, `minimize`, `maximize`.

Errors come back in the cell, in words, after the function's name:
`#DATALOG! …`.

## 18. Restriction as the feature

Why is there a DATALOG, when PROLOG can express everything it can?

It is the same reason as `SD-16`, one level down. Datalog is Prolog with
things removed, and each removal is a guarantee.

| Removed | Guaranteed |
|---|---|
| Function symbols | The answer is finite. Evaluation terminates. |
| Dependence on the order of rules | A rule set means the same however it is arranged. |
| Unstratified negation | A program has one meaning, and the loader can check that it has. |

Every question written as a sentence goes to DATALOG. A person who writes
`who reports-to "Alice" directly or not` over an organization chart that
contains a loop, by mistake, gets an answer. Under Prolog the same person
would get a ceiling and a refusal, and would be right to feel that the
fault was not theirs.

The sentence layer is kept inside the subset on which both engines agree.
Where they were found to disagree, the grammar was changed until the
disagreement could not be written.

> Restriction is the feature, checked when the rules load and not
> discovered while they run.

---

# Part VI. Evidence

## 19. Proofs about the prover

A system that refuses to guess has to show that its refusals and its
answers are right. The evidence comes in layers, and each layer exists
because the one before it was found wanting.

| Layer | Shows | Cannot show |
|---|---|---|
| **A rule's proof** | This sentence means this form. | That the form does the right thing. |
| **Golden files** | The whole corpus translates and compiles to exactly what it did. | That what it did was right. A golden detects change, not error. |
| **Parity** | Both ways of running leave the same sheets. | That both are not wrong together. |
| **Pins** | A bug once fixed stays fixed. | Anything about bugs not yet met. |
| **Static checks** | Thirty properties of the source, checked without Excel: no network call, no message that recommends VBA, every rule with a proof. | Behaviour. |
| **Proofs as forms** | A logic program, and the answer it must give, written as data. | See below. |
| **An oracle of another lineage** | That an independent solver gives the same answer. | |

### Proofs as forms

In the repository today, and part of `0.7.1`, the DATALOG engine's tests
are written in VLA:

```lisp
(test-datalog "a head variable absent from the body is refused"
  (program
    (rule (foo X Y) (bar X))
    (query foo))
  (refuses datalog-unsafe-head-variable))
```

A proof is the program it is, beside the answer it must give: the rows,
in order or in any order; TRUE or FALSE; or a refusal, named by its
identifier. Moving the tests into this form made them stronger. Tests of
refusals had been passing on *any* error, a crash included, and now each
names the refusal that must fire. Tests that had counted rows now name
them.

### An oracle of another lineage

The answers in those proofs were written down by the hands that wrote
the engine. Two readings by one author converge on consistency, which is
not the same as truth.

So each proof of an answer is also exported as a program for `clingo`, an
answer-set solver from another tradition entirely. The export includes a
judge that reports whether clingo's answer agrees with the one written
down, and three controls, each a proof with its expectation wrong on
purpose, which the judge must catch.

Two details are in character.

- **Frazaro never runs clingo**, and nothing in Frazaro calls it. The
  exporter writes files. A person runs the solver. `SD-13` is untouched.
- **The translation declines what it cannot translate faithfully**, by
  name: quoted strings, decimals, sums, the text tests. A proof of a
  refusal cannot be exported at all, since no other solver raises
  Frazaro's messages. So the oracle covers a part of the proofs and says
  which part, and the number it covers is a floor that a check enforces.

The project's word for the whole habit is *receipts*.

---

# Part VII. Time

## 20. Permanence

A natural language changes and nobody is consulted. A language that runs
procedures cannot, because a procedure is run by someone who has
forgotten what is in it.

| Instrument | Does |
|---|---|
| `SD-4` | A sentence that has shipped keeps its spelling and its meaning, or is retired through a refusal that says what to write instead. |
| [GRAMMAR_SINCE.md](../GRAMMAR_SINCE.md) | Records the release in which each sentence first *worked*. It is append-only. |
| `requires-version` | Lets a phrasebook, or a `<Frazaro>` section of a document, say which release it needs. |
| Version numbers, `SD-14` | **Patch**: nothing an existing sentence does has changed. **Minor**: something new can be said. **Major**: `SD-4` was invoked. |

The ledger records when a form first *worked*, and not when it was first
spelled. The founding case is a sentence that shipped in two releases and
never worked in either: it called a helper whose name was misspelled. It
is dated to the release that fixed it. To record the earlier date would
be to tell an author that their phrasebook runs where it does not.

`1.0.0` is reserved. It will be the release in which a named person
outside the project runs their own procedure, on their own machine, with
the author unreachable.

### When the promise and safety disagree

`0.7.1` narrows a shipped meaning, and the record says so plainly.
`Replace "A" with "B" in column C.` used Excel's own Replace, which edits
the text of formulas: it turned `=A1*2` into `=B1*2`, and could turn a
value into a live formula. From `0.7.1` the sentence changes values only,
and a new sentence, `… in formulas of column C`, does the other thing on
request.

This is `SD-4` yielding to security, and `SD-19` being applied late. The
method is worth noticing: the change is dated, reasoned in the release
notes, and paired with an explicit sibling, so that nobody who wanted the
old behaviour is left without a sentence for it.

---

# Part VIII. Other tongues

## 21. What a phrasebook translates

*Frazaro* is Esperanto for a phrasebook: *frazo*, a phrase; *-ar-*, a
collection of; *-o*, a noun. The name was chosen as a statement of
intent. It is only fair to say how far the intent has been carried out.

### What moves freely

Any sentence that is a phrase rule. The pattern is text and the template
is shared, so word order inside a rule is whatever the language wants.

```text
Pon hoy en la celda D1.
Pone hodie in cellula D1.
```

The Spanish edition is loaded and proved by the self-tests on every run.

### What the second phrasebook found

The project ran a falsification test: a small second language, wired
into the tests, to see what would break. Two things did, and both are
instructive.

**Rule order.** `pon {e:expr} en la celda` was registered before `pon hoy
en la celda`, and the hole took `hoy`. The English phrasebook had made
and repaired the same mistake. It is a property of ordered choice and
will be met by every phrasebook's author once.

**The filler words are English.** *A* is an article in English and a
preposition in Spanish. `Envia un correo a {who:expr}` failed, because
`a` is stripped from every pattern and is not passed over before a hole.
The repair was made in the phrasebook, by saying *para*. The finding was
about the engine.

### What the engine still assumes

| Assumption | Status |
|---|---|
| **The alphabet of a bare word is ASCII.** A word with an accent, outside quotation marks, is a stray character. | The shipped phrasebooks write without diacritics. It is the largest assumption left, and it lives in stage 3, not in the grammar. |
| **Names with accents** in the middle language are transliterated, `á` to `a`, since VBA's identifiers are ASCII. A name with no letter left is refused, naming the original. | Decided and shipped. |
| **The four filler words** and **the number words** are English. | Kept, and marked as data for a sibling language to replace. |
| **The structural sentences**, `If …, …`, are English in their words *and their order*. | The words can be aliased. The order cannot. |
| **The list grammar** requires the comma before *and*. | Structural. It cannot be swapped through a table. |
| **A keyword alias rewrites a word wherever it stands.** | A word that is also ordinary vocabulary cannot be an alias. Spanish *en* and *para* could not be. |
| **Captured words bind to English action names.** | `{d:bold\|italic}` stays English unless the phrasebook supplies actions named in its own language. |
| **Messages** are in a catalogue, by identifier. | The catalogue is English today. The structure for another is in place. |

### Honest prospects

A language with English's word order and a Latin alphabet is a
phrasebook away, and the Spanish edition is the evidence. A verb-final
language could say every phrase rule and none of the structural
sentences in its natural order. A language with case marking would need
more from holes than they give: the noun after *in* and the noun after
*from* would differ in form, and a pattern can only list the forms.

Those are questions about a dialect, and `SD-16` says how they are to be
raised: first as a specification of the dialect, and then as a dedicated
re-argument of the decision. They are not to be answered by a grammar
section that quietly grows a recursive slot.

---

# Part IX. Method

## 22. Designing a section

A *section* is a family of sentences for one kind of work: text, pivot
tables, borders. What follows is the project's practice, in the order it
is done.

**1. Collect sentences that were said.** From the log of refused
sentences and from transcripts. A prediction of what people will say is
a hypothesis, and specimens are what hypotheses are for.

**2. Write them as a user would, before writing any rule.** Read them
aloud. Ask of each verb whether it belongs to the person who will type
it or to the developer who will implement it. `Show` once printed to a
developer's window; it is now a message box, and `Log` does the
developer's job under its own name.

**3. Find the neighbours.** For each sentence, list what else Excel could
reasonably do that the sentence might describe. For every neighbour, the
sentence needs the word that excludes it.

**4. One verb, one mechanism.** If the section gives an existing verb a
second job, find another verb.

**5. When a sentence wants options, write a second sentence.** `Repeat`
stayed simple. `Count k from 2 to last-row` carries the deliberate cases,
and declares its variable by saying it.

**6. Choose the holes.** The narrowest category that is true. Check
every word that follows an `expr` or `cond` hole against the list of
words a value would eat.

**7. Decide what is refused, and write the refusals first.** For each
near-miss a writer will plausibly type, say what Frazaro will answer.
Pin the important ones with `test-fail`.

**8. Look for the dangerous branch.** For each ambiguity, ask what the
wrong reading costs. If the cost is a silent wrong number, refuse.

**9. Test the effect and not only the words.** Run it. A sentence can
break `SD-19` through a side effect.

**10. If two engines are involved, measure both** before choosing a
spelling, and choose the spelling on which they agree.

**11. Price it at real sizes.** A feature proven only on ten rows is not
proven. State the arithmetic at the sizes people have, and bound the
work and not the time, so that the same input gives the same answer on a
slow machine.

**12. Order the rules.** Specific before general.

**13. Prove it.** A proof for every rule and for every branch that reads
differently. Regenerate the golden files and read the difference. Run
both ways.

**14. Date it.** One row in the ledger for each pattern.

**15. Tell the truth about what is left out.** Every section in the
record ends with its named limits.

A section that passes all of this is small. That is the expected result.

---

# Part X. Open questions

## 23. What is not settled

A manual that claimed everything was settled would be the one document
in the repository to do so.

| Question | Where it stands |
|---|---|
| **Relations of more than one word.** `can cover` without the hyphen. | Planned, as declared relations. Today a relation is one token. |
| **Compositional references.** *The cell to the right of the last row of Sales.* | Explicitly deferred. It would need a recursive slot, which is to say a re-argument of `SD-16`. |
| **Pronouns and relative clauses** in the logic sentences. | Declined. The question is whether a regular fragment of them exists that is worth having. |
| **A precedence for `and` and `or`.** | None, by decision. The alternative is a rule that is right and widely misremembered. |
| **Diacritics in bare words.** | Open. It is a change to the alphabet, which every rule sits on. |
| **Structural sentences in another word order.** | Open. |
| **Which phrasebook a line of emitted code came from.** | Open, and listed publicly as a security item. |
| **Whether the expression grammar should be regular too.** | Not asked yet. It is fixed in the engine, which is what the overlap check needs. |
| **How large the language should grow.** | Unanswerable in advance. `SD-7` is the brake: one real sentence at a time. |

---

# Seminar

These have no single answer. Each has a note on what a good answer
attends to.

**1.** `SD-16` says that every slot consumes "one phrase of a built-in
sub-grammar" and that this does not weaken the overlap check. Argue that
it does not. Then construct the circumstance in which it would.

*Attend to: what the check treats as opaque, and who is able to change
what lies behind it.*

**2.** Find a sentence in the shipped phrasebook that has a neighbour and
no deciding word. Propose the sibling.

*Attend to: `SD-4`. The original keeps its meaning. `Add border to range`
is the founding case, so find another.*

**3.** Design the sentences for one kind of work the phrasebook does not
yet cover: data validation, say, or conditional formatting. Follow
Section 22. Submit the refusals before the rules.

*Attend to: where your sentences came from. If you imagined them, say
so.*

**4.** A colleague proposes that Frazaro accept any sentence, by sending
the ones it cannot read to a language model that rewrites them as
sentences it can. Which of the properties of Section 8 survive? Which of
Level 0's five promises?

*Attend to: what validation would then certify. The reference,
[Level 0b101](5-reference.md), takes a position on where a model belongs.*

**5.** Translate Section 12's three rules into a language you know well,
keeping the noun-as-variable convention. What did the articles do in
your language? What did case or gender do to the repeated noun?

*Attend to: whether the same noun is still visibly the same noun.*

**6.** The generated name `vla-not-leave-name-shift` is derivable by
hand. Derive the name for `not Roster lists the person as Staff-Name and
the day as Day`. Then say what would go wrong under a scheme that did
not double the hyphen.

*Attend to: injectivity. The answer is `vla-not-roster-staff--name-day`.*

**7.** Is Frazaro's English a language, by the definition this manual
opened with? Is it a dialect of English? Does it have an army?

*Attend to: the third clause of the definition. It has a register of
standing decisions and thirty check scripts. Opinions differ on whether
that is a navy.*

## What a Wizard can do

- [ ] State the matcher's contract from memory.
- [ ] Say which three properties the contract makes provable, and why a
      more generous parser could not prove them.
- [ ] Distinguish the parser's ambiguity from the reader's, and give the
      remedy for each.
- [ ] Write the grammar of `conditions`, and say what *regular* rules out.
- [ ] Explain why a noun can serve as a variable once the articles are
      gone, and what it costs.
- [ ] Derive a generated name by hand, and state the four tests it
      passes.
- [ ] Choose between the four engines for a given question, by what each
      guarantees.
- [ ] Say what each layer of evidence cannot show.
- [ ] Design a section, refusals first.
- [ ] Say, without embarrassment, what is not settled.

---

Previous: [Level 3 · Master](3-master.md) ·
Next: [Level 0b101 · Reference: An LLM's Guide to Frazaro](5-reference.md)
