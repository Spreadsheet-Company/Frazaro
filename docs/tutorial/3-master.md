# Level 3 · Master: A Hacker's Guide to VLA (Visual Lisp for Applications)

*(Macros that make macros that make macros that make...)*

| | |
|---|---|
| **For** | Developers: whoever maintains the phrasebook's machinery, reviews what Frazaro generates, or wants the middle language for its own sake. |
| **You need** | Level 2. Reading knowledge of VBA. Some Lisp helps and is not required; the parentheses are the easy part. |
| **You will be able to** | Read and write VLA, say what VBA any form becomes, write macros, and write macros that write macros. |
| **Plan on** | A day, most of it at the console (CLI). |
| **Comes after** | [Level 2](2-journeyman.md) |
| **Written against** | Frazaro `0.7.0` |

---

The sentence in the subtitle does terminate. It stops at depth 200, with a
refusal, and Section 14 says where the number came from. It was measured.

This is the manual for the layer in the middle. Above it, people write
English. Below it, Excel does what it is told. Between the two there is one
statement of what each sentence means, and that statement is written in a
Lisp.

A word about the name. *Visual Lisp for Applications* is *Visual Basic for
Applications* with one word changed, and the joke is exact: VLA is VBA,
wearing parentheses. If you were hoping for a more beautiful language than
VBA, this manual has to disappoint you early, and then explain why the
disappointment is the design.

## Contents

| Part | Sections |
|---|---|
| **I. The middle language** | 1 The bargain · 2 Where VLA is written · 3 Reading a form |
| **II. The core forms** | 4 Statements · 5 Expressions · 6 Places, dots and keywords · 7 Declarations · 8 Procedures · 9 Errors |
| **III. Macros** | 10 Templates · 11 No gensym · 12 The prelude · 13 Shielding and fusing · 14 Data at expansion time · 15 Macros that make macros |
| **IV. Two ways to run** | 16 Interpreter and compiler · 17 Reading the VBA · 18 Worksheet functions · 19 Rows by the slab |
| **V. The console** | 20 The CLI |
| **VI. Field notes** | 21 Hazards |
| **VII. Practice** | Exercises · Answers · What a Master can do · Coda |

---

# Part I. The middle language

## 1. The bargain

> VLA is VBA wearing parentheses: a 1:1 mapping, no semantic smoothing, no
> expression-if, no closures.

Every VBA statement has exactly one spelling in VLA, and VLA has nothing
that VBA lacks. `Set` and `Let` are different in VBA, so they are different
in VLA: `obj-set!` and `set!`. `Dim` is scoped to the procedure in VBA, so
it is in VLA. Where VBA has an edge, VLA has the same edge, and the
comments say so.

The project's history calls the alternative *the trapdoor we never opened*:
the moment the middle layer's meaning diverges from VBA's, somebody is
maintaining a real compiler, in VBA, alone, for ever. Each time a feature
offered to make VBA nicer than it is, the offer was declined.

What the bargain buys:

| | |
|---|---|
| **One statement of meaning** | There are two ways to run a program. If English went straight to VBA text, the interpreter would need a second reading of English, and two readings are two meanings. |
| **One-line templates** | A sentence rule produces a form. It does not produce VBA syntax with declarations, error handling, step numbers and doubled quotes. |
| **Composition** | A hole's value is a form placed inside a form. Splicing strings of VBA invites accidents of precedence and quoting; nesting lists does not. |
| **Languages for the price of a text file** | Every phrasebook shares the middle, so a new language is a file and not a backend. |
| **Review** | VLA is text. A reviewer can read it, a formatter can check it, and a test can compare it byte for byte. |
| **Room to move** | VBA is a backend, not the language (`SD-1`). The interpreter is already the default way to run, and any future host must agree with the same recorded results rather than invent its own reading (`SD-18`). |

What keeps VLA from being a bad idea is the third layer. Nobody writes a
procedure in it for a living. People write English; VLA is where that
English is *defined*. And macros, which cost nothing at run time because
they are gone by then, carry the comfort that the core declines to.

## 2. Where VLA is written

| Place | How | Notes |
|---|---|---|
| **A row of the *Frazaro* sheet** | Begin the row with `(`. | The form is a statement like any sentence. It may run over several rows; the parentheses say where it ends. No period is wanted. Validation checks it by compiling it in a scratch procedure, so a malformed form turns red before anything runs. |
| **The console** | **Open CLI**, or `` Ctrl+Shift+` `` | Text beginning with `(` or `;` is VLA. Anything else is a sentence. See Section 20. |
| **A phrasebook** | The template of a rule, and any `defmacro` the file carries. | Level 2. |
| **A library** | A `.vla` file of procedures and macros, brought in by the sentence `Use library "tools.vla".` or the form `(include "tools.vla")`. | Read from the folder of the workbook you are working in. Libraries may include libraries, sixteen deep. |

Definitions belong in a phrasebook or a library. A row of the sheet is a
statement inside the program, so a `sub`, `function`, `defmacro`, `type` or
`enum` there is refused, with directions.

Validation proves that a row is well-formed VLA. Whether the interpreter
can reach what the form names is settled when it runs (Section 16), and a
refusal there names the row.

And two ways to see what Frazaro wrote:

| Button | Gives you |
|---|---|
| **Translate File to VLA** | The program's translation, as a `.vla` file. |
| **Show me the VBA** · **Translate File to VBA** | What the compiler makes of it. |

## 3. Reading a form

```lisp
(set! (range "d1") (date))     ; put today into cell D1
```

| Thing | Rule |
|---|---|
| **List** | `(head item item …)`. The head says what kind of form it is. |
| **Symbol** | Any run of characters that is not a space, a parenthesis, a double quote or a semicolon. `set!`, `zero?`, `err.number`, `application.screenupdating` and `=\|>` are all symbols. |
| **Number** | Written plainly: `5`, `0.08`, `-1`. |
| **String** | Between double quotes. Two escapes: `\"` and `\\`. |
| **Keyword** | A symbol beginning with a colon, `:key1`. It names the argument that follows it. |
| **Comment** | From `;` to the end of the line. There is no other kind. `#` means nothing here. |
| **Capitals** | Ignored in symbols. Only the letters A to Z are folded, on purpose, so that a name means the same on a Turkish computer as on any other. |
| **Hyphens** | `row-number` in VLA is `row_number` in VBA. |
| **Layout** | Yours. **Lint VLA File** will put any file into house style. |

Three kinds of head:

1. **A core form**, such as `set!`, `if`, `for`. There are about sixty-five,
   operators included, and Part II lists them.
2. **A macro**, such as `when` or `make-bold`. It is replaced by what it
   stands for before anything runs.
3. **Anything else** is a call: `(date)` is `date()`, `(msgbox "hi")` as a
   statement is `Call msgbox("hi")`, and `(stamp :row-number 2)` calls
   your own procedure.

---

# Part II. The core forms

The VBA shown is what the compiler emits, taken from the project's golden
files. Emitted VBA is in small letters where VLA was; the VBA editor
restores the capitals it knows. The line tags the compiler adds are left
out until Section 17, which is about them.

Where this manual writes a form of its own, in the exercises and in a few
illustrations, it says so or it is plain from the context. Everything
attributed to the prelude, the phrasebook or a golden file is copied from
it.

## 4. Statements

| VLA | VBA |
|---|---|
| `(set! total 0)` | `total = 0` |
| `(obj-set! results (range "b2:b4"))` | `Set results = range("b2:b4")` |
| `(if (> grand 40) (then …) (elseif (> grand 20) …) (else …))` | `If (grand > 40) Then` … `ElseIf (grand > 20) Then` … `Else` … `End If` |
| `(for (k 1 f-last) …)` | `For k = 1 To f_last` … `Next k` |
| `(for (k 1 f-last 2) …)` | `For k = 1 To f_last Step 2` … |
| `(for-each (r results) …)` | `For Each r In results` … `Next r` |
| `(while (> countdown 0) …)` | `Do While (countdown > 0)` … `Loop` |
| `(do-until (= fuel 0) …)` | `Do Until (fuel = 0)` … `Loop` |
| `(select region (case ("North") …) (case ("South" "East") …) (case-else …))` | `Select Case region` · `Case "North"` · `Case "South", "East"` · `Case Else` · `End Select` |
| `(with (range "a1") …)` | `With range("a1")` … `End With` |
| `(begin … …)` | The statements, one after another. It has no VBA of its own. |
| `(return (* amount 0.08))` | `tax = (amount * 0.08)` then `Exit Function`, in a function named `tax` |
| `(return)` | `Exit Sub` or `Exit Function`, whichever it is inside |
| `(exit-sub)` `(exit-function)` `(exit-for)` `(exit-do)` | The same words, spaced. |
| `(debug-print (& "found " f))` | `Debug.Print ("found " & f)` |
| `(stamp :row-number 2 :value "beta")` | `Call stamp(row_number:=2, value:="beta")` |
| `(. (columns "a") autofit)` | `Call columns("a").autofit` |
| `(raw "' any text at all")` | The text, exactly. See Section 16 before using it. |

`then`, `elseif`, `else`, `case` and `case-else` are clauses. Outside their
parent form they are refused.

An `if` has no value. Neither has `select`. They are statements, as in
VBA, and an operator in statement position is refused in these words:

> '(+ ...)' is an expression, not a statement - use its value, as in
> (set! x (+ ...))

## 5. Expressions

| VLA | VBA |
|---|---|
| `(+ a b c)` | `(a + b + c)` Arithmetic chains: `+` `-` `*` `/` `\` |
| `(- x)` | Negation, with one operand. |
| `(& "step " counter)` | `("step " & counter)` |
| `(= a b)` `(<> a b)` `(< a b)` `(> a b)` `(<= a b)` `(>= a b)` | Comparison. |
| `(and a b)` `(or a b)` `(xor a b)` `(not a)` | Logic. |
| `(mod a b)` `(is a b)` `(like a b)` `(imp a b)` `(eqv a b)` | VBA's own, by their own names. |
| `(new Collection)` | `New Collection` |
| `(array "region" "product")` | `Array("region", "product")` |
| `(interpolate "Dear {name}, you owe {sum}." :name who :sum total)` | The pieces joined with `&`: `"Dear "`, `who`, `", you owe "`, `total`, `"."` |
| `(quote (a b c))` | Data, not code. Section 14. |
| `true` `false` `empty` `nothing` `null` | The same. |
| `(len code)` `(ucase name)` `(date)` | Any other head is a call. |

`interpolate` takes a literal template and keyword pairs. Every hole must
be given, every keyword must be used, and a doubled brace is a brace. It is
how the phrasebook writes worksheet formulas without counting quotation
marks.

## 6. Places, dots and keywords

A *place* is something a value can be put into.

```lisp
(set! total 5)                        ; a name
(set! (range "b2") 5)                 ; a cell
(set! (cells row-number "f") value)   ; a cell, computed
(set! (. r font.bold) true)           ; a property, through the dot
(set! application.screenupdating false)
```

**The dot form** is `(. object member arguments…)`.

```lisp
(. (columns c) autofit)                                   ; a method
(. r sort :key1 k :order1 order :header xlyes)            ; with named arguments
(set! (. (activesheet.listobjects n) showtotals) true)    ; a property of a result
```

A member may be a path: `font.bold`, `entirecolumn.autofit`. And a symbol
may carry its own dots, which is the usual way to write a global:
`err.number`, `activewindow.freezepanes`, `(err.clear)`.

**Keywords** become VBA's named arguments: `:key1 k` is `key1:=k`. A
macro's own named pieces are lists and no keywords: Section 10 has them.

**`set!` or `obj-set!`.** VBA makes this distinction and so VLA does.
Values are assigned with `set!`; object references with `obj-set!`. The
choice is made once, by whoever writes the macro, and nobody who writes
English ever meets it. The project's name for this habit is a bucket
brigade: every hard distinction is resolved exactly once, by the one
person equipped to answer it, and then hidden from everyone downstream.

One detail worth knowing: `(set! (. r formula) f)` writes through Excel's
`Formula2`, not `Formula`, so that a formula which returns several cells
spills instead of being quietly cut down to one.

## 7. Declarations

| VLA | VBA |
|---|---|
| `(dim total)` | `Dim total As Variant` |
| `(dim total Double)` | `Dim total As Double` |
| `(dim names (array String))` | `Dim names() As String` |
| `(dim grid (array Long 3 4))` | `Dim grid(3, 4) As Long` |
| `(dim band (array Long (to 1 5)))` | `Dim band(1 To 5) As Long` |
| `(redim names 10)` · `(redim preserve names 10)` | `ReDim names(10)` · `ReDim Preserve names(10)` |
| `(const hot-pink "#FF69B4")` | `Const hot_pink = "#FF69B4"` |
| `(type point (x Double) (y Double))` | `Type point` … `End Type` |
| `(enum colour red green)` | `Enum colour` … `End Enum` |
| `(public …)` `(private …)` | Wraps a `sub`, `function`, `dim`, `const`, `type` or `enum`. |

Two of VBA's rules arrive unchanged.

**`Dim` belongs to the procedure, not to the block.** A name declared
inside a loop is declared once, for the whole procedure, and keeps its
value after the loop. Section 11 turns this into a feature.

**Declarations at the top of a file come before the first procedure.** VBA
ignores one placed later, and every use of the name then fails. VLA refuses
instead:

> a module-level (dim ...) must come before the first (sub ...) or
> (function ...)

`type` and `enum` live at the top of a file only.

## 8. Procedures

```lisp
(function tax ((amount Variant))
  "eight per cent of an amount"
  (return (* amount 0.08)))

(sub stamp ((optional row-number Variant 1) (optional value Variant "ok"))
  "write a value into column F"
  (set! (cells row-number "f") value))
```

```vb
Public Function tax(amount As Variant) As Variant
    tax = (amount * 0.08)
    Exit Function
End Function

Public Sub stamp(Optional row_number As Variant = 1, Optional value As Variant = "ok")
    cells(row_number, "f") = value
End Sub
```

**Parameters**, one form each:

| Written | Means |
|---|---|
| `amount` | `amount`, by reference, a Variant. VBA's default, kept. |
| `(amount Double)` | `amount As Double` |
| `(byval n Long)` | `ByVal n As Long` |
| `(byref total Double)` | `ByRef total As Double` |
| `(optional rate Variant 0.05)` | `Optional rate As Variant = 0.05` |
| `(paramarray rest)` | `ParamArray rest() As Variant` |

**A function's type** follows its parameters: `(function area ((w Double)
(h Double)) Double …)`. With none written, it returns a Variant.

**A docstring** may follow the signature. A string there is documentation
only when at least one more form follows it; a string alone is the body.
The documentation never reaches the emitted code.

**`return`** with a value is legal only in a function.

### Tail calls

VBA has no tail-call elimination. VLA's compiler supplies it, under
conditions narrow enough to state exactly.

```lisp
(function fact-iter ((byval n Long) (byval acc Double)) Double
  "factorial, accumulator style"
  (if (<= n 1)
      (then (return acc))
      (else (return (fact-iter (- n 1) (* acc n))))))
```

When **every** parameter is `byval`, a `(return (self …))` is compiled as
*rebind the parameters and jump to the top*. It is recursion in the
language and a loop in the metal, with a stack that does not grow.

- With a `byref` parameter, rebinding would change the caller's variable.
  Such a function recurses in the ordinary way. Correctness came first.
- `(return (+ 1 (self …)))` is not a tail call and is not treated as one.
- All the new arguments are computed before any parameter is reassigned,
  so `(self b (- a 1))` means what it says.
- **This is the compiler's trick.** The interpreter recurses in fact, and
  its depth is VBA's.

## 9. Errors

| VLA | VBA |
|---|---|
| `(on-error goto handler)` | `On Error GoTo handler` |
| `(on-error resume-next)` | `On Error Resume Next` |
| `(on-error goto 0)` | `On Error GoTo 0` |
| `(label handler)` | `handler:` |
| `(goto done)` | `GoTo done` |
| `(resume retry)` | `Resume retry` |

Here is what the English `Try:` becomes. The sentences:

```text
Try:
  Go to sheet Nowhere-Land.
  Set rescue to "unreachable".

If that fails:
  Log "the problem was " joined with the problem.
```

The translation, from the golden file, with the step-tracking and
line-tag forms left out:

```lisp
(on-error goto vla-tryf-1)
  (activate-sheet "nowhere-land")
  (set! rescue "unreachable")
(goto vla-tryd-1)
(label vla-tryf-1)
(set! vla-problem err.description)
(resume vla-tryr-1)
(label vla-tryr-1)
(on-error goto vla-fail)
  (debug-print (& "the problem was " vla-problem))
(label vla-tryd-1)
(on-error goto vla-fail)
```

It is the idiom a careful VBA programmer would write by hand: arm, try,
jump over the handler, record the problem, `Resume` to clear the error
state, restore the outer handler. `the problem` in English is the name
`vla-problem` here. Nothing is hidden, and there is no runtime to trust
beyond VBA's own.

Names that begin `vla-` are reserved for Frazaro. Do not take them.

---

# Part III. Macros

## 10. Templates

```lisp
(defmacro
    (when test & body)
    "run the body when the test is true"
    (if test (then body)))
```

| Part | |
|---|---|
| `(when test & body)` | The name and its parameters. A name after `&` collects whatever arguments remain. |
| the string | Documentation. Optional, and worth writing: the tools show it. |
| the rest | The template. Where a parameter's name appears, the argument is put. |

```lisp
(when (> total 5)
      (inc! count)
      (msgbox "big"))
```

becomes, in order,

```lisp
(if (> total 5)
    (then (inc! count) (msgbox "big")))

(if (> total 5)
    (then (set! count (+ count 1))
    (msgbox "big")))
```

A macro in VLA is **pure substitution**. A template computes nothing, runs
nothing, and asks the workbook nothing. The argument is placed, whole and
unexamined, where the parameter stood. The result is expanded again until
no macro remains. What does compute while a template expands says so by its
name: Section 14's expand-time names, and the functions a file defines.

Several forms in a template body are taken as a sequence. A macro may
stand where a statement may, or where a value may: `(zero? x)` expands to
`(= x 0)` in either position.

**Names that are taken.** A macro may not be named `quote`, or after any
of the shielding forms and expansion primitives of Sections 13 and 14, the
word `self` among them. The expander handles those itself, so such a macro
could never fire, and the refusal says what the name already is:

> defmacro 'quasiquote': 'quasiquote' is the template-shielding form and
> cannot be a macro name

**Names that are not taken, and should be treated as if they were.**
Nothing stops a macro from being named like a function of VBA or of
Excel. It will then intercept every call of that name. Section 21 has the
story of the day this mattered.

**Names are global.** Every macro in the prelude, the phrasebooks and your
file shares one table. In a phrasebook, a second macro of the same name is
refused. A function (Section 14) shares the table too: a word names a macro
or a function, never both.

### Named pieces

A parameter is read by its place, which serves a macro of two or three and
stops serving one whose caller must remember what the fifth argument was.
A list in the signature, after the parameters, declares a **named piece**:

```lisp
(defmacro (rule range (sheet "Screen") (text "=1"))
  "a formula row, on the Screen and of the plain text unless the call says"
  (formula sheet range text))
```

`(sheet "Screen")` is a piece the call may leave out, its default standing
for it; `(sheet)`, the word alone, is one the call must give. A call writes
its positional arguments first, and after them each piece it names, a list
of the word and its value, in any order:

```lisp
(rule "B2:K11")                               ; (formula "Screen" "B2:K11" "=1")
(rule "B2:K11" (sheet "Tower"))               ; (formula "Tower" "B2:K11" "=1")
(rule "B2:K11" (text "=2") (sheet "Tower"))   ; (formula "Tower" "B2:K11" "=2")
(rule "B2:K11" (sheet "Tower") (text "=2"))   ; the same row
```

The pieces are bound in the order the signature declares them, never the
order the call writes them, so one call in two orders writes the same row.
A piece's word folds as a parameter's does, and stands in the template as a
parameter's does.

A default may name what is declared before it, a positional parameter or an
earlier piece, given or left out:

```lisp
(defmacro (frame sheet (top 1) (bottom top))
  "a box of rows, its bottom its top unless the call says"
  (box sheet top bottom))

(frame "S")                      ; (box "S" 1 1)
(frame "S" (top 3))              ; (box "S" 3 3)
(frame "S" (bottom 9) (top 3))   ; (box "S" 3 9)
```

A default is a form of the macro's, as its template is. Nothing happens to
it at the definition. Each call that leaves its piece out fills it in, and
it expands with the rest of what the call wrote, so a macro in it fires
there, and a word of `define` reads what the call's form sees (Section 14).

What the reading cannot take is refused by name. At the definition: a
declaration other than `(word)` or `(word default)`; a parameter after a
piece; a rest parameter beside a piece, since an argument after the
positional ones would be either; a word declared twice; a piece spelled as
an expand-time name; a default naming its own word or a later piece's. At
the call: anything after the positional arguments that is none of the
pieces, which the refusal lists; a piece holding other than one form; a
piece given twice; a piece with no default left out; and a positional
argument written as one of the pieces, since `(rule (sheet "Tower"))` has
left out the range. The call's line is named where its author wrote it, and
a call a macro made names none:

> macro 'rule': '(color "red")' stands after its positional arguments, where
> only its pieces go: sheet, text (vla line 5)

`&` keeps its one meaning in a signature, the mark before a rest parameter,
so `(& rest)` there is refused as no piece; and `(& a b)` is the joining of
two texts everywhere else, a positional argument, a piece's value or a
default among them.

A template replaces every word that is one of its parameters, a list's head
among them. So a template that calls a macro with pieces names no parameter
after one of them: in `(defmacro (tower sheet) (rule "B2" (sheet sheet)))`
the call becomes `(rule "B2" ("Tower" "Tower"))`, and is refused. Spell the
parameter otherwise, or write the call under `quasiquote` (Section 13), which
is how a macro calls itself by a piece.

## 11. No gensym

Lisp programmers will have been waiting for this section.

A macro that needs a temporary variable has to call it something. If the
name is fixed inside the template, two uses in one procedure declare it
twice. Most Lisps answer with `gensym`, which invents a fresh name nobody
will ever see. VLA answers differently, on purpose:

> The caller names every binding.

```lisp
(defmacro
    (swap! a b tmp)
    "swap two values through a caller-named temporary"
    (begin (dim tmp) (set! tmp a) (set! a b) (set! b tmp)))
```

```lisp
(swap! left right holder)
```

The price is visible in the signature: one more argument. What it buys is
worth listing.

- **The emitted code has no invented names.** An auditor who reads the VBA
  sees `holder`, which a person chose, and not `G__1047`.
- **Expansion is deterministic.** The same source gives the same text,
  byte for byte, which is what lets a test compare it.
- **Collisions are loud.** Use `holder` twice in one procedure and VBA
  refuses to compile, naming the duplicate declaration. There is no
  hygiene mechanism to fail quietly.

The same holds for labels. `with-fast-excel` takes the name of its own
restore label, and two uses in one procedure need two names.

And VBA's procedure-wide `Dim` stops being a hazard and becomes part of
the contract: after `(when-let x (range "b2") …)`, the name `x` still holds
the value, because the expansion says so and VBA agrees.

One word looks like an exception and is not. `self` (Section 14, *The place
a verb writes*) is a word the language keeps and never one it invents: the
caller's own `bind-expand` says what it stands for, a cell the caller wrote,
which a reader finds by hand.

## 12. The prelude

The standard library is [`prelude.vla`](../../scripts/prelude.vla),
written in VLA and placed in front of every program. It is worth reading
whole; the comments are half of it.

| Family | Macros |
|---|---|
| **Control** | `when` · `unless` · `if-not` · `while-not` · `dotimes` |
| **Counting** | `inc!` · `dec!` · `add!` |
| **Predicates** | `zero?` · `positive?` · `negative?` · `even?` · `odd?` · `empty?` · `blank?` |
| **Binding** | `let-one` · `let-one-as` · `when-let` · `if-let` · `try-else` · `swap!` · `returning` |
| **Brackets** | `with-fast-excel` · `with-screen-off` · `with-no-alerts` · `with-protected-sheet` · `with-sheet` · `with-error-handler` |
| **Instruments** | `time` · `trace-form` · `check` · `check=` · `comment` |
| **Lists, at expansion time** | `identity` · `cadr` · `caddr`, and six functions (Section 14): `length` · `reverse` · `append` · `quote-map` · `quote-reduce` · `quote-filter` |

Some signatures, since the arguments carry the doctrine:

```lisp
(dotimes v n & body)                     ; count v from 1 to n
(let-one x expr & body)                  ; dim x, set it, run the body
(when-let x expr & body)                 ; ... only when x is not Empty
(if-let x expr then-form else-form)      ; one form each: wrap several in begin
(try-else tmp expr fallback)             ; the value, or a fallback when it errors
(returning tmp expr & body)              ; compute the answer, tidy up, return it
(with-fast-excel restore-label & body)   ; screen and calculation off, restored on both exits
(with-protected-sheet ws pw restore-label & body)
(with-sheet ws prior & body)             ; work on a sheet, then go back
(with-error-handler handler-label handler & body)
(time label tmp & body)                  ; prints how long the body took
(check= label expected expr)             ; raises unless they are equal
```

`blank?` is the spreadsheet's idea of empty: Empty, Null, the empty text
and text that is only spaces. `empty?` is VBA's `IsEmpty` and nothing more.
The conditions of the English phrasebook read through `blank?`.

**The brackets restore on both exits.** `with-fast-excel` turns screen
updating and calculation off, runs the body, and turns them on again
whether the body finished or failed. If it failed, the error is raised
again for the caller after the restore. Leaving alerts off for the rest of
the session is the classic automation injury, which is the whole reason
`with-no-alerts` is a bracket.

## 13. Shielding and fusing

In most Lisps a template is quoted, and the writer unquotes what should be
filled in. VLA's templates are the other way up: **everything is filled in
unless shielded.** A parameter's name is replaced wherever it appears.

Usually that is what you want. When a template must mention a symbol that
happens to be spelled like a parameter, shield it:

| Form | Does |
|---|---|
| `(quasiquote form)` | Places the form as written, substituting nothing. |
| `(unquote x)` | Inside a shield, substitutes after all. |
| `(unquote-splicing xs)` | Inside a shield, places the items of a list in line. It must stand as an element of a list. |

Shields do not nest, and each of the three is meaningful only inside a
template.

**`symbol` fuses.**

```lisp
((symbol "un" name) param)
```

With `name` bound to `hide-column`, this is `(unhide-column param)`. The
pieces may be symbols or strings, and the result is one symbol. It is how
one template defines a pair of macros whose names differ by a prefix.

In a phrasebook's *rule* the same effect is had by gluing a captured word,
`make-{d}`. `symbol` is for templates, where there are no holes, only
parameters.

## 14. Data at expansion time

A macro receives forms. Sometimes it helps to treat a form as a list and
walk it. VLA has twenty-nine primitives for that, counting the arithmetic,
the text and the addresses a walk tends to need, a lookup and a test that
tells a word from a list, a loop that writes a form for each element, a
word for a form, `define`, and one for the place a verb writes, `self`; and a
file may define functions of its own, `defexpand`, that compute beside them.
They work on **quoted data in the source text**, while the program is being expanded.
They never see a cell, a range, or anything else that exists only when the
program runs.

| Primitive | Gives |
|---|---|
| `(quote (a b c))` | The list itself, as data. |
| `(car xs)` `(cdr xs)` `(cddr xs)` | The first item; all but the first; all but the first two. |
| `(cons x xs)` `(list a b c)` | A list with `x` in front; a list of the items. |
| `(null? xs)` | `true` when the list is empty. |
| `(list? x)` | `true` when `x` is a list, the empty one among them, and `false` for a word or a text; see *A word or a list* below. |
| `(eq? a b)` `(equal? a b)` | Whether two atoms are the same; whether two structures are. |
| `(quote-if test then-form else-form)` | One form or the other. The test must already be `true` or `false`. |
| `(cond (test form) … (else form))` | The first clause whose test holds. |
| `(+expand a b)` | The sum, as a number in the source. |
| `(=expand a b)` `(<>expand a b)` `(<expand a b)` `(>expand a b)` `(<=expand a b)` `(>=expand a b)` | The comparison, as `true` or `false`. |
| `(-expand a b)` `(*expand a b)` | The difference; the product. |
| `(quotient-expand a b)` `(mod-expand a b)` | Whole numbers only: the quotient rounded toward zero and the remainder with the divisor's sign, as Excel's `QUOTIENT` and `MOD` give them. |
| `(&expand a b ...)` | The pieces joined into one text. |
| `(characters-expand text)` | The text as a list of one-character texts. |
| `(address-expand row column)` | The address: `"B3"` for row 3, column 2. |
| `(row-expand address)` `(column-expand address)` | The row and the column of an address. `column-expand` also reads a column's letters alone. |
| `(column-letters-expand column)` | A column's letters: `"AB"` for 28. |
| `(assoc-expand head entries)` | The first entry whose own head is `head`, or `()` when none is. |
| `(each (name list) form ...)` | The forms written once for each element; see *The loop* below. |
| `(each ((name ...) list) form ...)` | The same, each name bound to its part of an element that is a list of as many parts. |
| `(each-index (index name list start) form ...)` | The same, with a count beside each element, from `start` up by one. |
| `(define name form)` | A word for the form, in every form of the file after this one; see *A name* below. |
| `(parameter name form)` | A define whose form a door that builds the file may replace; see *A value a door gives* below. |
| `(formula-expand expression)` | The expression as a formula's text; see *A formula from an expression* below. |
| `(sheet-expand reference)` `(range-expand reference)` | A reference's sheet, and its cell or range, as texts. |
| `(first-cell-expand reference)` | A reference's first cell, as a text: `"B1"` for `Tally!B1:E10`. |
| `(bind-expand ((self reference)) form ...)` | The forms, with the word `self` standing for one cell on a sheet; see *The place a verb writes* below. |
| `(offset-expand reference rows columns)` `(offset-expand reference rows columns height width)` | A reference moved, as Excel's `OFFSET` moves it. |
| `(defexpand (name parameter ...) form)` | A function, whose call computes a value while the file expands; see *A function* below. |
| `(let-expand ((name value) ...) form ...)` | The forms, each name standing for its value; see *A function* below. |

**A name that ends in `expand` is computed while the macro expands.** The
plain names, `MOD`, `ADDRESS`, `ROW`, are left to the formulas the program
writes, which compute when the workbook does. So a macro can work out an
address now and write a formula that works out the rest later:

```lisp
(defmacro
    (total-below col first last)
    "put a SUM of rows first to last of a column in the row under them"
    (set! (. (range (address-expand (+expand last 1) col)) formula)
          (&expand "=SUM(" (address-expand first col) ":" (address-expand last col) ")")))
```

`(total-below 2 1 9)` expands to `(set! (. (range "B10") formula)
"=SUM(B1:B9)")`. A row runs from 1 to 1,048,576 and a column from 1 to
16,384, Excel's own limits, and a number past either is refused by name:

> address-expand: a column is a whole number from 1 to 16384, got '16385'

**`cond` folds when it can and waits when it cannot.** A test that comes
to `true` or `false` during expansion chooses its clause on the spot. A
test that depends on the running program, such as `(> x 0)`, is left in
place as an ordinary `if`. Each clause decides separately. `else` must be
last, and a `cond` in which nothing matches expands to nothing at all.

**A word or a list.** `(list? x)` is `true` when `x` is a list, the empty
one among them, and `false` for a word or a text. It is the one test of a
form's kind that refuses neither: `null?` is `false` for a word and for a
list with items alike, `car` refuses a word, and `eq?` refuses a list. So a
macro can take entries of two kinds, a word that is its own label and a pair
whose label is no word:

```lisp
(defmacro
    (entry-label entry)
    "an entry's label: a pair's text, or a word's own"
    (quote-if (list? entry)
        (car entry)
        (&expand entry)))
```

`(entry-label phase)` expands to `"phase"`, and `(entry-label ("stopped at"
stopped-at))` to `"stopped at"`.

### The loop

A walk by recursion says how to reach the next element. A loop says what
each element writes:

```lisp
(defmacro
    (cells sheet & entries)
    "Puts values in cells of one sheet. Each entry is a cell's address, or a range of cells, and the value that goes there."
    (each ((address value) entries)
        (cell sheet address value)))
```

`(cells "S" ("A1" 1) ("B1" 2))` expands to `(begin (cell "S" "A1" 1)
(cell "S" "B1" 2))`. The binding `((address value) entries)` names the
parts of each element and the list, as `let` names a value; the forms
after it are a template, substituted once for each element with `address`
bound to its first part and `value` to its second, beside the macro's own
parameters. One name, `(entry entries)`, binds the whole element instead.
An element of another size than the names is refused, the entry shown:

> each: the names (address value) read each entry as a list of 2, a part for each name, and the entry ("A1") is a list of 1

`each-index` binds a count beside the element, from a start the binding
names, and takes a list of names the same way, `(index (color name)
entries 0)`:

```lisp
(defmacro
    (labels-at sheet top column & texts)
    "labels' rows: a text a row, down from the top row in one column"
    (each-index (row text texts top)
        (cell sheet (address-expand row column) text)))
```

The start is required. A palette counts from 0 and a sheet's rows from 1,
and nothing in the loop can tell which is meant, so the binding says.

Where a loop stands decides what it writes:

| Where it stands | What it writes |
|---|---|
| Alone: a macro's whole template, a branch of `quote-if` or `cond`, a form of the file | One `begin` of its results. |
| As an element of a list: `(+ (each (c cells) c))` | Its results, spliced into the list. |
| Where a primitive reads it, `(car (each ...))`, or after `unquote` | The list of its results. |

So a loop inside a loop gives one flat run of rows, a loop may walk the
results of another, and `(formula-expand (+ (each (c cells) c)))` writes
a sum.

Its names are the author's, as every name in a template is (Section 11).
A loop refuses a name the template already holds, a parameter, the rest or
an enclosing loop's name, rather than shadow it:

> each: 'xs' is already a name in this template - give the name each binds
> a word of its own

It refuses a list written as a macro's call too. A primitive's argument is
never expanded, so the call would be walked as its own words:

> each: the list to walk is a call of the macro 'rows', which does not
> expand where a list is walked - give the list itself, quoted or as a
> parameter

A loop is one application and no chain, so the chain's limit below does not
hold it: the kit's `cells` written as a walk stops at 2,499 entries, and
written as a loop it writes ten thousand.

**`assoc-expand` finds a piece by its head**, as Scheme's `assoc` does, so
that a verb can take its pieces by name and in any order:

```lisp
(defmacro
    (piece name & pieces)
    "a named piece's value: the entry whose head is the name, its second"
    (car (cdr (assoc-expand name pieces))))

(defmacro
    (put at & pieces)
    "a cell from named pieces, its sheet and its value in either order"
    (cell (piece sheet pieces) at (piece value pieces)))
```

`(put "A2" (value 1) (sheet "Tower"))` expands to `(cell "Tower" "A2" 1)`.
A symbol is compared without case, a text exactly and a list by its
structure, and a head that no entry has gives `()`.

**The list library** in the prelude is written with these, as functions,
which *A function* below introduces:

```lisp
(defexpand
    (reverse lst)
    "a quoted list's elements in the other order"
    (reverse-onto lst (quote ())))

(defexpand
    (reverse-onto lst kept)
    "internal: lst's elements put in front of kept, one at a time"
    (quote-if (null? lst)
              kept
              (reverse-onto (cdr lst) (cons (car lst) kept))))
```

`length` answers a number: `(length (quote (a b c)))` is `3` in the
emitted code, with no arithmetic left to do at run time, and since it is
a function, `(+expand (length xs) 1)` adds to it.

### A name

An address says where a value is, not what it is. `define` gives a form a
word, so that the source says what its author thinks:

```lisp
(define stop State!B12)
(define press State!B9)

(formula-expand
    (if (and (= stop 1) (= press 1))
        1
        0))
```

The last form expands to the formula `=IF(AND(State!B12=1,State!B9=1),1,0)`.
A define writes nothing where it stands, and its word is gone once the file
has expanded: neither the workbook nor any formula ever meets it.

**A define binds the forms after its own.** The forms of a file expand one
at a time, in order, and a define one of them holds binds its word from the
next form on. So a macro may write defines, and one list can give a block
its labels and its words together:

```lisp
(defmacro
    (names sheet top & words)
    "a block's labels down column A from a row, each word a name for the cell beside it"
    (each-index (row word words top)
        (label-and-name sheet row word)))

(defmacro
    (label-and-name sheet row word)
    "one label in column A, and its word a name for the value beside it in column B"
    (cell sheet (address-expand row 1) (&expand word))
    (define word (ref sheet (address-expand row 2))))

(names "State" 1 phase since level)
```

`(names "State" 1 phase since level)` writes the three labels down column
A, and from the next form on `phase` stands for `State!B1`, `since` for
`State!B2` and `level` for `State!B3`. In the form that defines it, a word
is not yet bound, wherever it stands, and a use there is refused:

> 'stop' is used in the form of the source that defines it, where it is not
> yet a name - a define binds the forms after its own

**Where a word is replaced.** Wherever it stands as a value: an argument, a
statement, anything inside `formula-expand`. Never at a list's head, so
`(define left State!B5)` and `(left text 1)` never meet; and never inside a
quote or a text. An expand-time name's argument reads it as its form too,
so `(eq? stop (quote stop))` is `false`: there as everywhere, `stop` is
`State!B12`. A macro's parameter of the same word hides it, since the
parameter is replaced first. A define may name another, defined before it or after it,
and its form expands where the word stands, a macro's call in it firing
there. A row's own words are values too: after `(define hidden 7)`,
`(sheet "S" hidden)` writes `(sheet "S" 7)`. So a define named like a word
a row reads where a value stands, `hidden`, `shown`, `none`, `wrap` or a
rule's closing `stop`, replaces that word in every row written after it, a
row a library's macro writes among them.

**The words it refuses** are the ones that already mean something: a
reference (`B12`, `A:A`, `Sheet1!x`, and `R1C1`, which Excel refuses as a
name), a number, `true` and `false`, a keyword, a form of the language
(`if`, `begin`), an expand-time name (`car`, `each`, `define`), a form of
the formula syntax's own (`error`, `ref`, `last`), a macro's name, and a
word defined already. A word no formula could hold is a fair name,
`mouse-x` among them, since the word never reaches a formula. A ring of
words that stand for one another is refused, the ring named:

> define 'end-row': the words end-row -> start-row -> end-row stand for one
> another in a ring - a define names a form that does not come back to it

The expander knows no variable, so a define replaces its word as a value
everywhere, in a running program's variable of the same word too: give a
define a cell, never a variable.

**A value a door gives.** `(parameter name form)` is a define whose form a
door that builds the file may replace. In the file it is a define: its word
binds the forms after its own and is read wherever a define's word is, as a
value too. Alonzo's size ladder builds Life on larger boards so:

```text
(parameter board-width 320)
(parameter board-height 200)
(interior board-width board-height "Screen" ...)
```

A door gives its values as forms of the same shape, `(parameter
board-width 640)`, kept apart from the file and read before it, and the file
expands as if it had been written with them. A door's value is of its
parameter's kind, a number for a number, a text for a text, a word for a
word, a list for a list, and one for a word no parameter of the file
declares, a define's among them, is refused by name:

> parameter 'board-widht': a door gives it a value, and no (parameter
> board-widht ...) of the source declares it

A parameter's own refusals are a define's, each naming it as a parameter.

### A formula from an expression

`formula-expand` writes an expression as a formula's text in Excel's own
spelling: no space outside a text, the fewest parentheses Excel's
precedence allows, a function's name upper-case, a sheet quoted only where
it must be. It expands its argument first, so a macro or a define inside it
has done its work before the text is written:

```lisp
(formula-expand
    (if (> (sum Data!B2:B13) 1000)
        "over"
        "under"))
```

expands to the formula `=IF(SUM(Data!B2:B13)>1000,"over","under")`. A
reference is a bare symbol, `Input!B5`, or `(ref "Q1 Data" "A1")` where a
symbol cannot hold it; an error is `(error "#N/A")`; and anything a formula
does not spell, such as `tax-rate`, is refused rather than renamed.

**A function called where it is written.** `(funcall f a ...)` calls what
`f` answers with its arguments, as Lisp's `funcall` does: a `LAMBDA`, or a
call that answers one, is called where it is written, and a name is called
by its name.

| Expression | Formula text |
|---|---|
| `(funcall (lambda (x) (+ x 1)) 5)` | `LAMBDA(x,x+1)(5)` |
| `(funcall (if c SQ PLUS1) 5)` | `IF(c,SQ,PLUS1)(5)` |
| `(funcall (funcall (make 3) 4) 5)` | `MAKE(3)(4)(5)` |
| `(funcall sq 3)` | `SQ(3)` |

Anything a formula cannot call, an operator, a number or a reference, is
refused by name. So is what `IFERROR` answers, since Excel calls nothing in
place that `IFERROR` answers: bind it with `let` and call the name,
`(let f (iferror SQ PLUS1) (funcall f 3))`, which writes
`LET(f,IFERROR(SQ,PLUS1),F(3))`.

**A reference's twin.** `(last X)` is the twin of the one reference on a
sheet that X is: the same cells on the sheet whose name is X's sheet's with
`.last` after it, the twin in which the language's machine keeps each
sheet's values from the frame before.

| Expression | Formula text |
|---|---|
| `(last Input!B13)` | `Input.last!B13` |
| `(last Tally!B1:E10)` | `Tally.last!B1:E10` |
| `(last (ref "Q1 Data" "A1"))` | `'Q1 Data.last'!A1` |
| `(last stop)`, after `(define stop State!B12)` | `State.last!B12` |

A reference with no sheet is refused, since a formula does not know the
sheet it will stand on when it is written:

> last: 'B13' names no sheet, and a formula does not know the sheet it will
> stand on - write the reference with its sheet, such as Input!B5

and so are a twin's twin and a name.

**A reference's marks.** `(absolute X)`, `(absolute-row X)` and
`(absolute-column X)` write the reference X with the `$` marks each adds,
Excel's absolute reference and its two mixed ones, so that a formula filled
across a range keeps a cell fixed as Excel's F4 key fixes it:

| Expression | Formula text |
|---|---|
| `(absolute State!B12)` | `State!$B$12` |
| `(absolute-row Gate!A4)` | `Gate!A$4` |
| `(absolute-column Gate!D1)` | `Gate!$D1` |
| `(absolute Tally!B1:E10)` | `Tally!$B$1:$E$10` |
| `(absolute B6)` | `$B$6` |
| `(last (absolute stop))`, after `(define stop State!B12)` | `State.last!$B$12` |

A mark is added and never taken away, so `(absolute-row $B12)` writes
`$B$12`, and the three marks and `last` stand in either order: `(absolute
(last stop))` writes `State.last!$B$12` too. A mark needs no sheet, as `$B$6`
shows. A name, a span of sheets and a text are refused, and so is a mark a
whole column or row has no place for:

> absolute-row: 'Screen!A:A' is a whole column, and absolute-row takes a
> reference with a row to mark, such as Input!B5 or 3:3

**A reference's two parts.** `(sheet-expand reference)` and `(range-expand
reference)` give the sheet and the cell or range of one reference on a
sheet, as the two texts `ref` takes, so that a macro can write a row at a
word:

```lisp
(defmacro
    (rule name expression)
    "one formula row at a named cell: its sheet and its cell from the name, its text from the expression"
    (formula (sheet-expand name) (range-expand name) (formula-expand expression)))
```

With the block above, `(names "State" 9 press take start stop)` and
`(define sliding 1)`, the form

```lisp
(rule stop
    (if (and (= (last phase) sliding) (= press 1))
        1
        0))
```

writes `(formula "State" "B12" "=IF(AND(State.last!B1=1,State!B9=1),1,0)")`,
the row Kairos's source writes today by its addresses. Like
`formula-expand`, the two expand their argument first, and are read where
the walk meets them, a row's field or a statement, never in an expand-time
primitive's argument. Each reads what a mark or `last` writes, so
`(range-expand (absolute stop))` is `"$B$12"`.

**A reference's first cell.** `(first-cell-expand reference)` gives the
first cell of one reference on a sheet, as the text `ref` takes: `"B1"` for
`Tally!B1:E10`, and a cell's own address for a cell, its marks as written. A
range written backwards gives its top-left, `"B1"` for `Tally!E10:B1`, as a
row's range is read. A verb that reads its own last value twins its range's
first cell by it, the verb's two texts made one reference by `ref`:

```lisp
(defmacro
    (count sheet range event)
    "each cell its own last value plus an event: =sheet.last!first+event"
    (formula sheet range
        (formula-expand (+ (last (ref sheet (first-cell-expand (ref sheet range)))) event))))
```

`(count "Tally" "B1:E10" Input!B16)` writes `(formula "Tally" "B1:E10"
"=Tally.last!B1+Input!B16")`, and so does the same verb at a word, its sheet
and range handed in as `(sheet-expand tally)` and `(range-expand tally)`
after `(define tally Tally!B1:E10)`. A whole column or row is refused, since
a fill's range is never one; and like the two parts, the first cell expands
its argument first and is read where the walk meets it, never in a
primitive's argument.

### The place a verb writes

A row over a range is one formula, written as its first cell reads. Life's
rule over `B2:LG199` reads `Screen.last!A1:C3` to mean "my eight neighbors and
me, last frame", and its author does the arithmetic by hand against `B2`, a
cell the source never names. **`self` names it.** A verb binds the word, and
whatever stands inside the verb reads it:

```lisp
(defmacro
    (rule sheet range expression)
    "one formula row, self its range's first cell"
    (bind-expand ((self (ref sheet (first-cell-expand (ref sheet range)))))
        (formula sheet range (formula-expand expression))))

(defmacro
    (running event)
    "its own last value plus an event"
    (+ (last self) event))

(rule "Tally" "B1:E10" (running Input!B16))
```

The last form writes `(formula "Tally" "B1:E10" "=Tally.last!B1+Input!B16")`.
`running` takes no sheet and no range: it reads the place of the verb around
it.

**The binding.** `(bind-expand ((self reference)) form ...)` expands its forms
with `self` standing for one cell on a sheet, and then the word means what it
meant before. The reference is worked out first, under any binding outside
this one, so `(bind-expand ((self (offset-expand self 1 0))) ...)` reads the
outer `self` and moves it a row down. The forms are expanded completely inside
the binding, each in the order written: one form gives its own result, and
several give one `begin`. A binding inside another hides it for its own forms,
and the outer stands again after them. Two verbs beside each other each have
their own, in whichever order they expand. A word other than `self`, a text, a
reference with no sheet, and a range are refused:

> bind-expand: self is one cell on a sheet, and 'A!B2:C3' is a range - a fill
> binds its first cell, (first-cell-expand reference)

**Where the word is read.** Where a define's word is: wherever it stands as a
value, a macro's result among them, inside `formula-expand` and inside the
forms that read a reference, `sheet-expand`, `range-expand`,
`first-cell-expand` and `offset-expand`; never at a list's head, inside a quote
or a text. So a macro over it may be called from any verb, and a word of
`define` may name it: after `(define neighbors-last (last (offset-expand self
-1 -1 3 3)))`, `(sum neighbors-last)` reads the neighbors of whichever cell the
verb around it writes. A default that names `self` reads it where the call
stands. A macro's parameter of the same word hides it, since the parameter is
replaced first.

**A reference moved.** `(offset-expand reference rows columns)` and
`(offset-expand reference rows columns height width)` are Excel's `OFFSET` while
the macro expands, in its argument order:

| Expression | Formula text |
|---|---|
| `(offset-expand Tally!B2 1 2)` | `Tally!D3` |
| `(offset-expand Screen!B2 -1 -1 3 3)` | `Screen!A1:C3` |
| `(offset-expand Tally!B2:E10 1 1)` | `Tally!C3:F11` |
| `(offset-expand Tally!$B2 1 1)` | `Tally!$C3` |
| `(offset-expand B2 1 0)` | `B3` |
| `(last (offset-expand Stay!C2 1 1))` | `Stay.last!D3` |
| `(offset-expand stop 1 0)`, after `(define stop State!B12)` | `State!B13` |

A range with no size keeps its own, a `$` stays where it was written, and its
answer is `(ref sheet part)`, which `last`, the marks and a reference's parts
all read. Like them it expands its arguments first, and is read where the walk
meets it. A row or a column past the sheet, a whole column or row, a size under
1 and a number that is not whole are refused by name:

> offset-expand: a row is a whole number from 1 to 1048576, got '0'

So the words over `self` are macros over it, and each reads any reference:

```lisp
(defmacro
    (above at)
    "the cell above a reference"
    (offset-expand at -1 0))

(defmacro
    (around at radius)
    "the square of a radius around a cell, the cell at its center"
    (offset-expand at
        (-expand 0 radius)
        (-expand 0 radius)
        (+expand (*expand radius 2) 1)
        (+expand (*expand radius 2) 1)))

(defmacro
    (neighbors-alive)
    "the eight neighbors alive in the last frame"
    (- (sum (last (around self 1))) (last self)))

(rule "Screen" "B2:LG199" (neighbors-alive))
(rule "Screen" "A2:AV24" (last (above self)))
```

These two write `=SUM(Screen.last!A1:C3)-Screen.last!B2` and `=Screen.last!A1`,
and `(formula-expand (above stop))` writes `=State!B11`. A word over `self`
writes the sheet's name, `Screen!B1` where the formula stands on the Screen;
`(ref (range-expand (above self)))` writes the bare `B1`.

**A primitive does not read it.** An expand-time primitive takes its arguments
as data while a macro's template is filled in, before the walk meets the word,
so a primitive that reads a word's text refuses this one:

> row-expand takes its arguments as data, and self is read only where a form
> stands as an expression - write the words over it, such as (offset-expand self
> rows columns), or (range-expand self) where a row's field wants its address

`car`, `cons`, `eq?` and the rest carry the word as the data it is, so `rules`,
`on` and a loop move an expression that holds it.

**Where it is refused, and where it is not.** With no binding around it, the
word is refused where it would be read, in a formula's expression and wherever
one reference is read:

> self: nothing reads it here - self is the place a verb writes, and it stands
> only inside (bind-expand ((self reference)) form ...), which a verb writes
> around its row

Everywhere else it is the word it always was. A sentence's or a program's
variable named `self` is a variable, unless a binding stands around it, which
then replaces it as a define replaces its word: give a binding a verb's row,
never a running program. `(ref "self")` writes a defined name so spelled. A
macro, a loop's name, a named piece and a define may not take the word.

**A verb inside a verb.** A verb writes a row, so it has no place inside another
verb's expression, and `formula-expand` says so:

> formula-expand: one inside another would write the inner formula as a text -
> write the inner expression itself, and a verb meant to stand inside an
> expression as an expression macro over self

Since `self` tells an expression macro its place, a verb that writes one row can
be an expression, with `rule` around it, and such expressions compose:

```lisp
(defmacro
    (pressed button)
    "1 in the frame a button goes down"
    (if (and (= button 1) (= (last button) 0)) 1 0))

(rule "Tally" "B16" (running (pressed Input!B1)))
```

writes `(formula "Tally" "B16" "=Tally.last!B16+IF(AND(Input!B1=1,Input.last!B1=0),1,0)")`.

### A function

A macro's argument is a form, placed unexamined. Sometimes a helper should
answer a value instead: the address beside another, the numbers of a list
that pass a test, a run of rows. `defexpand` defines one:

```lisp
(defexpand
    (beside at columns)
    "the address so many columns right of another"
    (address-expand (row-expand at) (+expand (column-expand at) columns)))

(cell "S" (&expand (beside "A1" 1) ":C9") 1)
```

writes `(cell "S" "B1:C9" 1)`. A function's call computes wherever an
expand-time name computes: at the top of a file, in a macro's template, and
in another call's argument, `&expand`'s here. Its arguments are values,
each computed before its body is read, so `(beside (beside "A1" 1) 1)` is
`"C1"`. Its body is one form, read as a template is, and its value is
expanded again as any form is, so a function may answer a macro's call or a
row, and a loop in its body gives the list of its results, so a map is a
loop:

```lisp
(defexpand (square x) (*expand x x))
(defexpand (squares xs) (each (x xs) (square x)))

(cell "S" "A1" (squares (quote (1 2 3))))
```

writes `(cell "S" "A1" (1 4 9))`.

**A function may call itself.** A test decided while the file expands
chooses the branch, and a call in the branch chosen is the function's
value:

```lisp
(defexpand
    (keep-big xs)
    "the numbers of a list greater than 2"
    (quote-if (null? xs)
        (quote ())
        (quote-if (>expand (car xs) 2)
            (cons (car xs) (keep-big (cdr xs)))
            (keep-big (cdr xs)))))
```

`(keep-big (quote (1 5 2 7)))` is `(5 7)`. A call that is its function's
whole value, its body, the branch `quote-if` or `cond` chose, or a
`let-expand`'s form, replaces the call under way and loops, so a function
counting down from 4,999 needs no more room than one counting from 1. Any
other call, such as the one inside `cons` above, waits for its value, and
such calls nest at most 14 deep. One value is worked out in at most 5,000
calls, so a function that never stops is refused, never left running:

> function 'down': more than 5000 calls working out one value (a function
> that never stops calling?) (computing (down 0))

A function is handed to another by its name: with `(defexpand (apply-to f
x) (f x))`, `(apply-to square 4)` is `16`. A macro's parameter of the same
word hides a function, as it hides a define.

The prelude's list library is written so: `length`, `reverse`, `append`,
`quote-map`, `quote-reduce` and `quote-filter` are functions, and each
calls itself as its whole value, so it reads a list as long as the 5,000
calls allow. A list one computes is a list the next reads: `(append
(reverse (quote (2 1))) (quote (3 4)))` is `(1 2 3 4)`, and `(quote-map
square (quote (1 2 3)))` is `(1 4 9)`. Handed a macro, `quote-map` gives
the list of its calls, each expanded after.

**A name for a value, inside a form.** `let-expand` names values for its
forms, each binding reading the ones before it:

```lisp
(defmacro
    (fit h)
    "the fit's last row, nine below the wall's height"
    (let-expand ((last-row (+expand h 9)))
        (cell "Fit" (&expand "A10:A" last-row) last-row)))
```

`(fit 24)` writes `(cell "Fit" "A10:A33" 33)`. Its forms are a sequence
where a statement stands, and one form where a value is read. It may not
bind a name twice, a word taken already, or the name of a parameter of the
macro or function around it.

**A call handed in is computed.** A macro's argument is placed unexamined,
but where an expand-time name reads a parameter, a loop's name or a
`let-expand`'s name, a call of an expand-time name or of a function handed
in that way is computed first: with `(defmacro (plus-one x) (+expand x
1))`, `(plus-one (square 3))` writes `10`. A macro's call handed in stays a
list, since a macro writes forms and not a value, and a quote keeps a call
as data: with `(defmacro (first-word call) (car call))`, `(first-word
(quote (+expand 1 2)))` is `+expand`.

**The place it writes.** `self` passes through a function to the forms the
walk computes: with `(defexpand (ring at) (offset-expand at -1 -1 3 3))`,
`(formula-expand (sum (ring self)))` inside `(bind-expand ((self
Screen!B2)) ...)` writes `"=SUM(Screen!A1:C3)"`. An expand-time name in its
body still takes its arguments as data, so a body's `(column-expand at)`,
`self` handed in, is refused as it is anywhere.

**When a refusal comes from inside.** A refusal raised while a value is
worked out names the calls it stands in, the outermost first:

> +expand: expected a number, got a list (computing (outer (a)) > (inner
> (a)))

Past four calls the middle ones are written `...`, a call longer than 60
characters is cut, and a value of more than 60 forms is written `...`. The
stepper shows a call the walk meets as one application and its value as the
next.

**The names it refuses** are the ones a define refuses, and a word that
names a macro or a function already: a word names one thing, so a second
definition of a function or of a macro is refused by name, the first one's
line named where it has one. The prelude's are among them: a file's own
`length` or `when`, a macro or a function, is refused by name. A value a
door changes is a parameter's, never a macro defined again (*A value a door
gives*, above). A function declares no
named pieces, since its arguments are values and a call could not tell a
piece from a call of another function.

### Where it stops

| Limit | Value | When reached |
|---|---|---|
| Nested expansion | 200 deep | macro expansion too deep (recursive macro?) |
| Applications in one chain | 5,000 | macro self-expansion exceeded 5000 applications in one chain (infinite macro?) |
| Nested `include` | 16 files | include: nesting deeper than 16 files - is a file including itself? |
| Calls waiting inside one value | 14 deep | function 'f': calls nested more than 14 deep - a call that is the function's value, its body, a chosen branch or a let-expand's form, loops and nests none |
| Calls working out one value | 5,000 | function 'f': more than 5000 calls working out one value (a function that never stops calling?) |

The 200 was not chosen for its roundness. A first attempt raised it to
1,000 and the host ran out of stack somewhere between 150 and 250, so the
number went back. A recursive macro therefore handles lists of modest
length; a long table is a loop, which is one application and no chain. The
refusal is loud; a partial result is never returned.

Two properties are worth a Lisp programmer's attention, because they differ
from home.

- **There is no `lambda` at expansion time.** `quote-map` takes the *name*
  of a macro or function of one argument, and a function takes another by
  its name. A loop's forms are a template written in place, which is what a
  `lambda` would be for here.
- **A filter is a function.** `quote-if` needs its test already decided, and
  a macro's call in the test is not yet expanded at the moment the test is
  inspected; a function's call is computed there, so a filter is written as
  one, `keep-big` above. The prelude's is `quote-filter`: with `(defexpand
  (big? n) (>expand n 2))`, `(quote-filter big? (quote (1 5 2 7)))` is
  `(5 7)`. Its test is a function answering `true` or `false`; a macro's is
  refused by name.

## 15. Macros that make macros

A template may contain a `defmacro`. That is the whole trick.

```lisp
(defmacro
    (bool-antonym-family name param target property doc-pos doc-neg)
    "generate a boolean-property on/off macro pair sharing one target-and-property shape"
    (begin (defmacro
               (name param)
               doc-pos
               (set! (. target property) true))

           (defmacro
               ((symbol "un" name) param)
               doc-neg
               (set! (. target property) false))))

(bool-antonym-family hide-column c (columns c) hidden
                     "hide a column" "show a hidden column")
```

One call, two macros: `hide-column` and `unhide-column`.

**A generator is a phrasebook's tool.** It is the phrasebook's loader that
takes what a generator produces and registers the macros, rules and
proofs in it. In a plain library, or at the console, macro definitions are
collected from the top level of the text before anything is expanded, so a
`defmacro` that only appears after expansion arrives too late to be one.

The English rule that uses the pair was already as small as it could be:

```lisp
(english-vla
    "{d:hide|unhide} column {c:column}"
    ({d}-column {c}))
```

### One level further

In a phrasebook, a generator may produce rules and proofs as well as
macros. This one is shipped:

```lisp
(defmacro
    (table-property-family name property doc pattern call testsentence testcall value & extras)
    "generate a table-name/value-setting macro, its english-vla rule, and its test-success proof"
    (begin (defmacro
               (name n extras)
               doc
               (set! (. (activesheet.listobjects n) property) value))

           (english-vla
               pattern
               call)
           (test-success
               testsentence
               testcall)))

(table-property-family
    table-totals-on
    showtotals
    "show an Excel Table's total row"
    "show the total row of table {n:text}"
    (table-totals-on {n})
    "Show the total row of table Sales."
    (table-totals-on "sales")
    true)
```

The loader treats a generator's output exactly as if a person had typed
it. The rule is registered, the proof is run, the duplicate check applies.
**A generated sentence that fails its proof refuses the phrasebook like any
other.** Generation does not buy an exemption from anything.

### And one further than that

`tablespec` walks a quoted table and calls a generator once for each row:

```lisp
(defmacro
    (tablespec spec)
    "walk a quoted list of (name doc body pattern call testsentence testcall) rows ..."
    (quote-if (null? spec)
              (begin)
              (begin (tablespec-row
                         (car (car spec))
                         (cadr (car spec))
                         …)
                     (tablespec (cdr spec)))))
```

A macro that walks data and calls a macro that writes macros, rules and
proofs. That is three levels, and the subtitle's ellipsis is earned.

Its comment records that an early design called for `quasiquote`, and that
tracing the expansion by hand showed the plain shape was enough. The house
habit is to trace first.

### Keeping generated things traceable

Generated code is where an auditor loses the thread, so there are three
provisions.

| Provision | What it does |
|---|---|
| **Export Expanded Phrasebook** | Writes the phrasebook as the loader sees it, every generated macro, rule and proof spelled out. The shipped copy is [`english_expanded.vla`](../../scripts/polyglotta/english_expanded.vla). |
| `(at-row label form)` | A wrapper a generator puts round what it produces for one row of its table. A failing proof then names the row as well as the line. |
| The row label in compiled code | A macro produced under `at-row` carries its label into the VBA, beside the line tags of Section 17. |

The rule of thumb the project follows, and recommends: build the
traceability before the generator, not after the first question about
where a sentence came from.

---

# Part IV. Two ways to run

## 16. Interpreter and compiler

| | **Interpret** | **Compile** |
|---|---|---|
| **What happens** | Frazaro walks the forms and performs each on the live workbook. | Frazaro emits a VBA module, places it in the workbook, and runs it. |
| **Writes code into the workbook** | No. | Yes. |
| **Needs Excel's trust in the VBA project** | No. | Yes, and refuses in words without it. |
| **Reaches** | What has been reviewed and named. | Whatever VBA reaches. |
| **Tail calls** | Recursion. | A loop. |
| **For** | Every day. | The auditor who wants the artefact. |

### What the interpreter reaches

The interpreter does not pass an arbitrary member name to Excel. It knows
a list: the places (`range`, `cells`, `rows`, `columns`, `worksheets`,
`workbooks`), a set of properties and methods added one at a time as real
sentences needed them, a handful of VBA's functions (`len`, `left`,
`right`, `trim`, `ucase`, `lcase`, `instr`, `round`, `date`, `now`,
`time`, `isempty`, `msgbox`, `inputbox`), the common worksheet functions,
and Frazaro's own runtime helpers.

A member outside the list is refused by name, not attempted. That is a
security property before it is a limitation: what an interpreted program
can touch is a list somebody reviewed.

Other edges, each with its own refusal:

| | Interpreter |
|---|---|
| `(new …)` | `Collection` only. |
| Named arguments | On a fixed list of methods, and not mixed with positional ones. |
| Positional arguments to a member | Four at most. |
| `err` | `number`, `description`, `source`. |
| `(resume next)`, bare `(resume)` | Not supported; `(resume label)` is. |
| `paramarray` | Not supported in your own procedures. |

If you write raw VLA that reaches past these, compile it. If a *sentence*
needs to reach past them, the member belongs on the list, and the address
in [SUPPORT.md](../SUPPORT.md) is where to say so.

### Forms that belong to one side

| Form | Runs under | Because |
|---|---|---|
| `raw` | Compile | It is VBA text. There is nothing to walk. |
| `deflambda`, `lambda` | Compile | They produce a worksheet formula. Section 18. |
| `with`, `redim`, `type`, `enum` | Compile | No sentence produces them, so the interpreter has not been taught them. It refuses them by name. |
| `When the sheet changes:` | Interpret | The handler lives in the add-in's memory, not in your workbook. |
| `Optimize cell C1.` | Interpret | A compiled program runs without the engine beside it. |

The standing rule (`SD-5`) is that every backend supports or *explicitly
refuses* every core form. A form that one side silently ignored would be
the worst of the three possibilities.

### `raw`

`(raw "…")` places VBA text into the compiled module. It is the one form
with no limits, and the one a phrasebook cannot use without the consent
dialog of Level 2. In your own library it is yours to use, and the reason
to avoid it is practical: a program containing it has one way to run
instead of two.

### How the two are kept honest

1. **Golden files.** The regression corpus,
   [`instructions.txt`](../../scripts/instructions.txt), is translated to
   VLA and compiled to VBA on every change, and compared with the recorded
   result. An empty difference is the evidence that behaviour was kept. A
   difference is a question with exactly two answers: fix the code, or
   approve the new output. It is never approved by reflex.
2. **Parity.** The same corpus is run both ways against a live workbook,
   and the sheets that result are compared.
3. **Pins.** Every bug ever fixed has a test, with the incident cited
   beside it.

[TESTING.md](../TESTING.md) describes all six passes.

## 17. Reading the VBA

Press **Show me the VBA** on any program. A procedure looks like this:

```vb
Public Sub stamp(Optional row_number As Variant = 1, Optional value As Variant = "ok")
    On Error GoTo vla_fail ' vla:8
    vla_step = 52 ' vla:9
    If vlatraceon() Then ' vla:10
        Call vlatracestep(52, vla_step_text(52)) ' vla:10
    End If
    cells(row_number, "f") = value ' vla:12 src:90
    Exit Sub ' vla:13
vla_fail: ' vla:14
    Call vla_report_error ' vla:15
End Sub
```

| Mark | Means |
|---|---|
| `' vla:12` | This line came from line 12 of the VLA. |
| `src:90` | Which came from line 90 of the instructions: the sentence itself, in the program sheet. |
| `vla_step = 52` | The step counter. When a run stops, this is how the message knows the sentence and the row. |
| `vlatraceon()` | The trace buttons' switch. |
| `vla_fail:` | Every procedure ends in a handler, so that a failure is reported in Frazaro's words and never in a VBA dialog with a Debug button. |

Two tags and not one, deliberately. A single tag that sometimes meant a
VLA line and sometimes a sentence would have been shorter and would have
been wrong half the time.

A macro's own scaffolding carries no tags. The statements you passed into
it keep theirs. The map marks what the writer wrote.

## 18. Worksheet functions

`deflambda` registers an Excel `LAMBDA` under a name in the workbook.

```lisp
(sub register-bricks ()
     "registers the worksheet-function bricks"
     (deflambda sum-to
                (n)
                "the sum 1..n by REDUCE over SEQUENCE - iteration, not recursion"
                (reduce 0 (sequence n) (lambda (acc i) (+ acc i))))

     (deflambda clamp (x lo hi)
                "x pinned into the closed range lo..hi"
                (min hi (max lo x))))
```

After a compiled run, `=sum_to(100)` works in any cell of the workbook.

| Rule | |
|---|---|
| It lives inside a `sub`. | It registers the function when the program runs. |
| The body is one expression. | Statements cannot run inside a worksheet function. |
| `if` takes three arguments here. | `(if test then-value else-value)`, as in a formula. |
| `lambda` is for the inner function. | It has meaning nowhere else. |
| Parameters are plain names. | Formulas have no named arguments. |
| The predicates work. | `(positive? v)` expands before the formula is written. |
| Compile only. | |

The shipped library [`alonzo.vla`](../../scripts/alonzo.vla) holds five of
these, built on `REDUCE`, `SCAN` and `MAP`, which iterate and so have no
recursion ceiling to hit. It is named for Church.

## 19. Rows by the slab

Reading and writing cells one at a time is slow, because each touch is a
trip to Excel. `for-each-row` makes the trip twice.

```lisp
(for-each-row (row (range "a2:c5000"))
  (set! (row 3) (* (row 1) (row 2))))
```

The range is read once into memory. The body runs for each row, where
`(row 3)` is the third column of the current row. The whole range is
written back once, after the loop.

| Property | |
|---|---|
| **All or nothing** | An error in the body leaves the sheet untouched: the write-back is never reached. |
| `(exit-for)` | Leaves the loop. What was changed so far is written. |
| `(return)` | Leaves the procedure. Nothing is written. |
| **The body's reach** | The current row's own columns. No other cell. |
| **Both ways of running** | The same results, checked. |

The restriction on the body is what keeps the form auditable: there is one
fixed thing it can mean.

---

# Part V. The console

## 20. The CLI

**Open CLI** on the Frazaro tab, or `` Ctrl+Shift+` ``.

| Key | Does |
|---|---|
| `Ctrl+Enter` | Runs what is in the box. |
| `Esc` | Closes the window. |
| `Ctrl+Up`, `Ctrl+Down` | Earlier and later commands. |
| `Ctrl+Shift+Delete` | Clears the history, after asking. |

| Word | Does |
|---|---|
| `history` | Lists the last twenty commands, numbered. |
| `history 50` | Lists fifty. |
| `!12` | Brings back command 12. |
| `clear` | Empties the transcript. The history and the last three results are kept. |
| `*` `**` `***` | The last three results, newest first, as at any Lisp listener. |

```lisp
; VLA comments start with ; (not #)

(function
  fib-iter ((byval n Long) (byval a Long) (byval b Long)) Long
  "tail-recursive fibonacci, accumulator style"
  (if (<= n 0)
    (then (return a))
    (else (return (fib-iter (- n 1) b (+ a b))))))

(fib-iter 10 0 1)
```

The transcript records the command after `~`, anything printed, the
result after `=`, and how the run ended, with the time.

What to know before relying on it:

| | |
|---|---|
| **It interprets.** | Always. So Section 16's list applies. |
| **It acts on the sheet in front of you.** | There is no *Output* default here. |
| **It takes no snapshot.** | **Undo Last Run** does not reach what the console did. |
| **Each command starts afresh.** | A function or macro defined in one command is gone in the next. Put the definitions and the call in the same box. |
| **A form sees the core and the prelude.** | The phrasebook's macros travel with a *sentence's* translation. Type `Make cell A1 bold.` and they come along; type `(make-bold (range "a1"))` alone and `make-bold` is unknown. |
| **History is kept on disk**, in `%APPDATA%\Frazaro\history.txt`. | A command that mentions a password, secret, token, API key or credential is kept for the session and never written. |

---

# Part VI. Field notes

## 21. Hazards

Each of these cost somebody an afternoon. They are recorded in the source
at the place where they happened, and collected here.

**A macro named like something else.** The list library first shipped
`map` and `reduce`. Excel already has functions by those names, and
`alonzo.vla` was calling them inside `deflambda` bodies. Macros expand
before formulas are written, so the new macros quietly took those calls,
and a test that had passed for weeks failed. They are now `quote-map` and
`quote-reduce`. *Before naming a macro, search for the name.*

**An unexpanded call, inspected.** The first `append` was written
`(append-onto (reverse a) b)`. But a macro's argument arrives as written.
`append-onto` inspected the form `(reverse a)` as though it were a list,
and its first item was the symbol `reverse`. The result of appending
`(1 2)` and `(3 4)` was `((1 2) reverse 3 4)`. A call handed to a macro
that examines it at once is data, not yet a value. The repair fused the
two walks into one. A function's call handed in is computed before it is
read, so the library's `append`, a function now, is written as the first
draft meant: `(reverse-onto (reverse front) back)`.

**The same name twice.** Two uses of `let-one`, `swap!` or a `with-`
bracket in one procedure need different names for their temporaries and
labels. VBA's compile step says so if you forget. That is the design
working.

**`On Error` belongs to the procedure.** The `(on-error goto 0)` that
closes `try-else` or a `with-` bracket disarms *any* handler in force,
including one you armed outside it. Arm again afterwards, or keep the two
apart.

**An argument used twice is evaluated twice.** `check=` mentions its
operands in the comparison and again in the failure message.
`with-protected-sheet` mentions the sheet when unprotecting and when
protecting. If an argument has an effect, bind it with `let-one` first.

**A value is not an object.** `let-one`, `when-let` and `try-else` assign
with `set!`. To bind a worksheet or a range, write the `dim` and the
`obj-set!` yourself.

**A name VBA already owns.** VBA's names ignore capitals, so a variable
called `cvar` collides with the function `CVar`, and the line that
declares it is a syntax error with no explanation. The same goes for
`instr`, `date`, `time` and their kin.

**A clause-less `if` is legal.** `(if x)` compiles to an `If` with nothing
in it. VBA allows it, so VLA does. The project discovered this by writing
a test that expected a refusal.

**The fast bracket changes what a program reads.** With calculation off, a
program that writes a formula and then reads the cell reads a stale value.
This is why `with-fast-excel` is a bracket for you to place, and is not
wrapped round every program.

### For readers at home in the VBA editor

[IMMEDIATE.md](../IMMEDIATE.md) lists commands for the Immediate window,
among them `EnglishToVla`, which gives the VLA of one sentence, and
`EnglishExplain`, which says which rule claimed it and from which file.
Two more live in the same module as the compiler: `VlaExpandStep` prints a
form after each macro application in turn, from *as written* to the end, a
form that computes while the macro expands standing as written until the
macros inside it have fired, and `VlaApropos` searches the names and
descriptions of every macro loaded, with the English sentence that reaches
each one.

Nothing in Frazaro's messages will ever send a user there. It is a
workshop, and the door is marked.

---

# Part VII. Practice

## Exercises

**1.** What VBA does this become?

```lisp
(for (k 10 1 -1)
  (when (even? k) (debug-print k)))
```

**2.** Write `until`, the counterpart of the prelude's `while-not`, so
that `(until (= fuel 0) (dec! fuel))` loops until the test holds. Use a
core form.

**3.** Write a macro `with-status` that shows a message in the status bar
while its body runs and clears it afterwards, on both exits. Model it on
`with-screen-off`. The phrasebook clears the status bar by setting it to
`false`.

**4.** Why is this function not compiled as a loop?

```lisp
(function total-to ((byval n Long)) Long
  (if (<= n 0)
      (then (return 0))
      (else (return (+ n (total-to (- n 1)))))))
```

Rewrite it so that it is.

**5.** Using `symbol`, write a generator `on-off-family`, for a phrasebook,
that given a name and a property of `application` defines `<name>-on` and
`<name>-off`.

**6.** What does `(length (quote (a b c d)))` leave in the emitted code,
and when is the counting done?

**7.** A colleague's raw row, `(. (range "a1") addcomment "checked")`,
is green under **Validate Instructions**. **Interpret and Run** then stops
at that row with a message about a list of members. What happened, and
what are the two ways forward?

## Answers

**1.**

```vb
For k = 10 To 1 Step -1
    If ((k Mod 2) = 0) Then
        Debug.Print k
    End If
Next k
```

`when` becomes `if`, and `even?` becomes `(= (mod k 2) 0)`.

**2.**

```lisp
(defmacro
    (until test & body)
    "loop the body until the test holds"
    (do-until test body))
```

**3.**

```lisp
(defmacro
    (with-status message restore-label & body)
    "show a status-bar message around the body, cleared on both exits"
    (set! application.statusbar message)
    (on-error goto restore-label)
    body
    (label restore-label)
    (set! application.statusbar false)
    (if (<> err.number 0) (then (err.raise err.number err.source err.description)))
    (on-error goto 0))
```

The caller names the label. The error is raised again after the restore,
from inside the handler, where it cannot be caught locally and so reaches
the caller.

**4.** The recursive call is an operand of `+`, not the direct operand of
`return`, so it is not a tail call. With an accumulator it is:

```lisp
(function total-to ((byval n Long) (byval acc Long)) Long
  (if (<= n 0)
      (then (return acc))
      (else (return (total-to (- n 1) (+ acc n))))))
```

**5.**

```lisp
(defmacro
    (on-off-family name property)
    "generate <name>-on and <name>-off for a property of application"
    (begin (defmacro
               ((symbol name "-on"))
               "turn it on"
               (set! (. application property) true))

           (defmacro
               ((symbol name "-off"))
               "turn it off"
               (set! (. application property) false))))

(on-off-family alerts displayalerts)
```

This defines `(alerts-on)` and `(alerts-off)`.

**6.** The literal `4`. The counting is done while the program is being
expanded, by `+expand`, one step for each item. Nothing is left to compute
when the program runs.

**7.** Validation proved the row well-formed, which it is. The interpreter
then refused a member that is not on its reviewed list, by name, without
attempting it. The two ways forward: run the program with **Compile and
Run**, which reaches whatever VBA reaches; or, if a sentence ought to be
able to do this, ask for the member to be added, with the sentence that
needs it.

## What a Master can do

- [ ] Say why VLA declines to be nicer than VBA.
- [ ] Read a form and say what VBA it becomes.
- [ ] Choose between `set!` and `obj-set!` without looking it up.
- [ ] Write a function the compiler turns into a loop, and say why it may.
- [ ] Write a macro that takes its temporaries and labels from its caller.
- [ ] Explain, to a Lisp programmer, the absence of `gensym`, and to an
      auditor, its benefit.
- [ ] Walk a quoted list at expansion time, and say how long a list may be.
- [ ] Write a generator, and find what it generated.
- [ ] Say what the interpreter reaches that the compiler does not, and the
      reverse.
- [ ] Follow a line of emitted VBA back to its sentence.

## Coda: first contact

There is one more phrasebook in the repository,
[`alien.vla`](../../scripts/polyglotta/alien.vla), and it contains no
human language at all.

The reader accepts as a symbol any run of characters that is not a space,
a parenthesis, a quotation mark or a semicolon. So `=|>` is a symbol as
lawful as `set!`. Every glyph in the file is the name of a macro or of a
macro's parameter, and macros are gone before any code is emitted, so the
VBA that results contains none of them and compiles. The one name that has
to survive into VBA is spelled in leetspeak, because VBA's identifiers
cannot wear glyphs.

The docstrings are its Rosetta stone.

It could not have been done one layer up. The sentence reader accepts
letters, digits, hyphens and underscores, and refuses a stray character
loudly, by design. The joke lives at the one layer whose reader welcomes
it: Alien Lisp, and not Alien English.

It is a joke. It is also a proof, which is the project's preferred kind.

---

Previous: [Level 2 · Journeyman](2-journeyman.md) ·
Next: [Level 4 · Wizard: A Sorcerer's Guide to Computational Grammar](4-wizard.md)
