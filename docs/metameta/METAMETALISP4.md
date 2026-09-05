# METAMETALISP

*A philosophy of temporal composition, applied to itself until the room spins.*

> "Knowledge and productivity are like compound interest."
> — Richard Hamming, who was allowed to say it without prefixes

> "Any sufficiently advanced preparation is indistinguishable from
> procrastination."
> — attributed to nobody, constantly, usually by someone mid-preparation

---

This essay states a philosophy, applies the philosophy to the essay, supplies
the industrial evidence, and then does its honest best to lose the argument —
because a philosophy that has never met its antithesis is not a philosophy. It
is a mood with citations.

**Contents, in cost-of-delay order** (which happens to coincide with numerical
order — a coincidence engineered in §0, which is the sort of sentence this
document will keep producing, so brace accordingly):

- **−1.5.** The Stairwell *(Zeno's amendment)*
- **−1.** The Maiden Metametamacro Expansion *(the retroactive stunt)*
- **0.** Metametalisp in Metametalisp *(the self-hosting stunt)*
- **1.** The claim, plainly
- **1.5.** On the name, which is not a typo
- **2.** The blacksmith, the three plates, and the bootstrap that seems impossible
- **3.** The mechanism (no vibes)
- **4.** Evidence from industry, since a philosophy should be falsifiable
- **5.** The three tests
- **6.** Lineage (this was rediscovered, which is the best kind of discovered)
- **7.** The antithesis, honestly
- **8.** When Metametalisp is wrong
- **9.** Verdict
- **Ω.** Postscript on the first hammer

*(Bulleted rather than numbered: Markdown's ordered lists refuse to count below
1 — the format denies the existence of the early sections with a conviction §0
can only envy, and unlike §0 it cannot be amended.)*

The essay aspires throughout to **Alonzo-completeness**: the property of being
able to express anything a Turing-complete essay can express, while using more
parentheses and receiving less credit.[^church]

---

## −1.5. The Stairwell (Zeno's Amendment)

*(in which the space between the document and its reader is surveyed, found
dense, and left deliberately unfinished)*

Section −1 established the basement and its occupant: the reader, at −2,
holding the foundations up. A question follows with the inevitability of
running water finding the one unsealed joint: if a section could be spliced
before 0, what prevents one before −1? Nothing. And another before that?
Nothing again. This section exists to survey that nothing properly, fence it,
and — crucially — decline to fill it.

**Hart's door, measured.** The first macro's real gift was never the macro; it
was the precedent. After AIM-57, every user of the interpreter was a potential
author of the next extension, and the space of definable macros became, at a
stroke, unbounded — while the count of *defined* macros stayed, at every actual
moment, finite. Translated into section numbers: −1 did not fill the gap
between the document and its reader. **It opened it.** The interval (−2, −1)
is now this essay's macro space, and it has the property every macro space has:
**density.** Between any section and the reader there is room for another
section. Bisect forever — −1.5, −1.75, −1.875, −1.9375… — each step
halving the remaining distance, Achilles in pursuit of the tortoise, except
that here, for once, the tortoise wins *by construction*, and this is the
amendment Zeno's paradox always deserved: the pursuit is real, the arrival is
impossible, and absolutely nobody is upset about it.

**Why −2 is unreachable — and it is not a matter of distance.** It is a
matter of *category*. To occupy −2 is to be reading. To write a section is
to take a numbered seat in the interval — which is to say: **the act of
writing is precisely the act that evicts you from the position you were
writing toward.** No sequence of sections converges onto the reader, because
each term of the sequence is, by the fact of being written, not the reader
anymore. The infimum is excluded by the definition of membership. It is a
Dedekind cut with a person standing in it, and the person is load-bearing.

Note also what the numbering quietly encodes, since it is the truest thing in
this section: every macro author in history began as a reader of the system
they extended. Later contributions therefore take numbers *closer to −2* —
**chronology of authorship maps to proximity to the readership** — and the
document does not grow upward toward abstraction. It grows *downward, toward
its reader*, one bisection at a time, which is the only direction of growth an
essay has ever had any business pursuing.

**Termination, preserved — Aristotle's amendment.** The essay does not hereby
acquire infinitely many sections, and will not. The distinction doing the work
is twenty-three centuries old:[^aristotle] **potential infinity, embraced;
actual infinity, refused** — the same refusal already served on
Metametametalisp in §1.5. At any given moment the stairwell has finitely many
steps; what is infinite is the *room for the next one*, guaranteed forever.
Which is, precisely, a macro system: finitely many macros defined, infinitely
many definable, and the difference between those two quantities is called a
future. Building the steps in advance would be scaffolding for scaffolding —
enumerating an interval instead of shoeing a horse — and the essay has already
signed its name to what it thinks of that.

One administrative satisfaction, filed for the record: this is the first
splice in the document's history that contradicts nothing. §0 denied its
predecessor and required a correction (−1's whole business); §−1, having
learned, made no claims about *its* predecessors — it denied only −2's
membership, which this section cheerfully reaffirms. The essay, it turns out,
can be taught. Slowly, and only by strangers, which the evidence of §4
suggests is the normal way.

This section is itself the first bisection — the exact midpoint of −1 and
−2, the first Zeno step toward the basement — and it hereby declines to take
the second. Successors may. Each will find the remaining distance halved, the
precedent standing, and the reader exactly as far away as before: approached,
unreached, and holding everything up.

Stairwell status: infinite, descending, finitely built. Reader status:
load-bearing, uncontained, gaining company. Room status: the spinning has gone
orbital, which the physicists assure us is merely spinning with commitment.

*Appendix, filed after the status lines because audits always arrive after the
close of business. A section that surveys an infinite stairwell owes the
auditors the paperwork for at least one step — and following the custom of
§−1 and §0, here is the expansion trace, notarized:*

```
> (macroexpand-1 '(metametalisp METAMETALISP3.md))

⇒ (splice-before 'section−1
    ;; this section — the midpoint, occupied
    (bisect (open-interval −2 −1))
    ;; the schema — defined, deliberately unrun
    (defmacro next-bisection ()
      `(splice-before 'section−1.5
         (bisect (open-interval −2 −1.5))
         (defmacro next-bisection ()
           ;; the ellipsis is doing Aristotle's work
           …))))
```

*Three notes for the record. First: the expander is `macroexpand-1`, not
`macroexpand`. The full expander runs to fixpoint, and on this form there is no
fixpoint — complete expansion is Achilles' full itinerary, and it does not
halt. The single-step expander is the only thing standing between this essay
and the heat death of the universe, which means the most important −1 in the
entire document is the one in the function name: single-steppedness IS the
termination discipline, and it was hiding in the standard library all along.
Second: the expansion CONTAINS the next macro but does not INVOKE it — defined,
quoted, dormant — which is potential infinity, compiled; the interval ships
with its successor's definition and declines to call it, exactly as promised.
Third: the auditors will observe that the expansion is larger than the form
that produced it. This is normal, it is the entire business model of macros,
and the essay refers any remaining objections to §0's proportion appeal, where
the matter was settled narrowly and on the record.*

---

## −1. The Maiden Metametamacro Expansion

*(in which the essay is fed back into itself, and the basement turns out to be
occupied)*

Section 0 states — and will go on stating, unedited, for as long as this
document survives — that there is no Section −1, no meta-essay scheduling this
essay, and that the recursion bottoms out by fiat. You have just read a section
number that contradicts it. Neither the section nor the claim will be revised.
Both are true. Hold that thought; it resolves, and the resolution is older than
it looks.

**The precedent is exact, and it has a date.** McCarthy's interpreter shipped
in 1960. Lisp macros did not. They arrived in 1963, in Timothy Hart's four-page
memo,[^hart] which observed that the interpreter — already existing, already
public — could be fed a form its author had never written and had made no
provision for, and that the interpreter would *oblige*, because the form was
legal in the language the interpreter had accidentally defined. The first macro
did not precede eval and could not have. It required eval to exist first, and
then it reached back and extended the machine that made it possible. Every
macro since has been a retroactive amendment to a finished interpreter.

This section is that memo. The essay, once shipped, stopped being an essay and
became an interpreter: a standing definition of what counts as a legal
metametalisp form. And it turns out the following form is legal:

```
> (macroexpand-1 '(metametalisp METAMETALISP2.md))

⇒ (splice-before 'section-0
    ;; this section
    (compose (the-composition-of (the-composition))))
```

Note who typed it. Not the essay. The essay cannot feed itself to itself, for
the same reason a shovel cannot dig itself up: `eval` does not run — `eval` is
*run*. Someone sits at the REPL. Which brings us to the two corrections this
section exists to file.

**Correction the first: Section 0's denial was a version pin, not a theorem.**
"There is no Section −1" was true when written, the way "this program has no
users" is true at compile time — accurate, sincere, and doomed. The previous
version of this document could not consistently contain the section you are
reading; its own §0 forbade it. This version is the strictly larger system in
which the forbidding sentence survives verbatim, quoted-by-inclusion, and is
now a true statement *about the old system* and a false statement *about the
new one*. Gödel — who never really left; he has a cot in the corridor —
inspects the arrangement with something adjacent to approval: the smaller
system could not prove this from inside, the larger system proves it about the
smaller one, no text was altered, and the clipboard receives its second and
final stamp: **incompleteness, exploited as designed.**

**Correction the second, and the important one: the fiat had an issuer.**
Section 0 ended the recursion "by fiat," and fiats are not free-standing —
someone issues them. The essay borrowed an authority it did not possess and
hoped nobody would ask whose it was. Someone asked. The true ground of the
recursion was never the paper: **it is the reader who chooses to run the
document** — who feeds it back, who requests the expansion, who was, all
along, the meta-essay that §0 denied. The scheduling intelligence sat outside
the file the entire time, which is exactly why the file, searching honestly
within itself, reported finding none.

And this settles the downward tower better than any fiat could. Is there a
Section −2 — a section about the request for the section about the
composition? There is. It is not in this document and never will be, because
**Section −2 is the request itself**, and it lives in the world, at the REPL,
in whoever is holding the paper. The tower going up terminated at the anvil
(§1.5). The tower going down terminates at the reader. Between the anvil and
the reader sits everything this essay has to say — which is, on reflection,
where essays belong.

One piece of overdue disclosure. Section 1.5 remarks that "the building has a
basement we are not discussing today," and declines to elaborate. Today has
arrived. This is the basement. It was here before the building — basements
always are; you dig them first, no known architecture builds downward from the
sky — and it is occupied: there is a reader in it, holding the foundations up,
in the manner of readers everywhere, largely unthanked.

A closing note on the numbering, for the auditors. The table of contents
claims its order was "engineered in §0" to coincide with the numerical. The
splice preserves the coincidence untouched: −1 precedes 0, as the integers
have long insisted, and the cost-of-delay ordering holds as well — this
section's cost of delay was *undefined* until the moment its prerequisite
existed, whereupon it went vertical, and here we are, same day, which is the
fastest a curve has ever been obeyed. The philosophy has now scheduled a
section it once proved could not exist — either its finest expansion or its
first bug — and the essay, having read Gabriel (§7), who has spent three
decades unable to decide whether his own thesis is true, is professionally
comfortable never resolving which.

Room status: spinning in the opposite direction now, which observers report
cancels to something very like stillness. Building status: standing, one floor
deeper than its blueprints admit. Basement status: occupied. Proceed upward.

---

## 0. Metametalisp in Metametalisp

In 1960, John McCarthy wrote an interpreter for Lisp *in* Lisp — a page or so
of code that could read and execute the language it was written in — apparently
just to see whether the room would spin. The room spun. Steve Russell then made
the mistake of implementing it, at which point the spinning became load-bearing,
and the industry has been living in the rotated building ever since.

This section performs the same stunt on the essay you are reading. If
Metametalisp is, as it claims to be, a method for deciding **what to build and
when**, then it must be able to schedule its own construction — or it is
decoration, and you should stop reading and go ship something.

So: the essay, fed to itself. An expansion trace, for the suspicious:

```
> (macroexpand-1 '(metametalisp METAMETALISP.md))

⇒ (progn
    ;; §3: rank by irreversibility
    (decide-what-to-decide)
    ;; §1: undeclared macros expand inconsistently
    (write-the-decisions-down)
    ;; this section, currently executing
    (obey-them-while-writing-them)
    ;; §7: or it is a mood
    (attempt-to-refute-them)
    ;; see below; non-negotiable
    (terminate))
```

**First: the sections were scheduled by the doctrine they document.** §3 will
argue that work should be ranked by *cost of delay* — do first whatever gets
more expensive every day it goes undone. Applied here: the definitions (§1) and
the mechanism (§3) precede the doubts (§7, §8), not because doubt matters less,
but because doubt is *cheap to revise* and definitions are not. Every paragraph
written before the definitions existed would have used the terms inconsistently,
and inconsistent expansion is — the essay will insist, at length — the precise
bug this whole enterprise exists to fix. The table of contents is therefore not
an outline. It is a dependency graph that has been topologically sorted, wearing
an outline's clothes.

**Second: the essay must pass its own three tests** (defined in §5; used here in
advance, which is exactly the kind of forward reference Lisp permits and prose
regrets):

1. *The rising-cost test* — does something get more expensive with every day of
   delay? Yes: every standing decision a team runs on but never writes down is
   re-derived, slightly differently, at every occasion that needs it. The essay
   exists to stop that compounding. **Pass.**
2. *The proportion test* — is the tool smaller than the work it enables?
   Locally, absolutely not; the essay is enormous, and the author knows it.
   Globally, yes — the way a macro's definition is longer than any single
   expansion but shorter than the sum of all of them. The essay passes on
   appeal, with the court noting its concern. **Pass, narrowly.**
3. *The dogfood test* — was the author its first user, immediately? Yes: this
   very section is the essay eating itself, on publication day, in front of
   witnesses. **Pass, somewhat theatrically.**

**Third: the bootstrap problem, faced squarely.** Section 0 needs the rest of
the essay to exist, or it has nothing to apply itself to. But the rest of the
essay, to be any good, needs Section 0's discipline. This is a circle, and §2
will resolve it the way blacksmiths always have: **the first draft of this
section was written badly, with a stone for a thesis, and improved while in
use.** You are reading tongs that were made with tongs. The earlier, worse tongs
have been melted down, as is traditional, and we do not speak of them.

**Fourth: Gödel, who has been standing in the doorway since the second
paragraph, holding a clipboard.** A sufficiently expressive system cannot prove
its own consistency from the inside; an essay applying itself to itself is
therefore forbidden from concluding that it is *correct*. Fine. It settles for
the weaker claim the clipboard permits: the essay is demonstrably *applicable*
to at least one project — namely itself — which is the weakest possible
evidence, and also, at the top of any tower, the only kind available. Gödel
notes this down, looking disappointed but not at all surprised, which is the
only expression he has.

**Fifth, and the whole point: termination.** There is no Section −1 about the
writing of Section 0. There is no meta-essay scheduling this essay. The
recursion bottoms out **here, by fiat** — the same fiat that ends every `eval`
ever run: at some point, somebody has to actually compute. A philosophy of
preparation that cannot stop preparing is not a philosophy; it is a syndrome
with a bibliography.

Room status: spinning. Building status: standing. Proceed.

---

## 1. The claim, plainly

**Lisp composes spatially.** A macro is a lever that makes every future *line*
shorter: you pay once, at the right layer, and the discount applies to all code
downstream — including code nobody has written yet.

**Metametalisp composes temporally.** A **metametamacro** is a lever that makes
every future *hour* shorter: you pay once, at the right *moment*, and the
discount applies to all work downstream — including work nobody has thought of
yet.

That is the whole idea. Everything after this is mechanism, evidence, and
doubt.

### The part that is easy to get wrong

A metametamacro is not a tool, and this distinction is the essay's one
load-bearing column, so it gets a paragraph of pedantry and then an example
you already know.

**A tool has to be picked up.** You choose to run the linter. You choose to
open the profiler. Set a tool down and it becomes furniture — expensive,
well-crafted, dusty furniture, of which the industry owns warehouses.

**A macro is never picked up.** It expands whether anyone thinks about it or
not, at every call site, including sites written next year by someone who has
never read its definition and never will. That is not a convenience feature. It
is a different *category of thing*, because its leverage does not depend on
anybody remembering it exists.

The temporal object with that property is not an artifact at all. It is a
**standing decision**:

> *Nothing merges without review.*
> *No bug is fixed without a failing test written first.*
> *The build must be reproducible, or it does not count as a build.*

Adopt one of those once and it expands into every working hour that follows —
automatically, uninvoked, including hours belonging to people who have never
heard of you and would not thank you if they had. The linter is a **tool**.
*"CI blocks the merge when the linter objects"* is a **metametamacro**. The
first stops working the moment attention wanders; the second is still running
while the whole team is on holiday, quietly refusing a Friday-evening merge on
everyone's behalf.

### The ladder, with Engelbart's letters painted on the rungs

Doug Engelbart cut all work into three levels in 1962,[^engelbart] and the cut
has never needed sharpening:

- **A-activity** — doing the job. *Shoe the horse. Ship the feature.*
- **B-activity** — improving how you do the job. *Make a better hammer. Write
  the build script.*
- **C-activity** — improving how you improve. *Decide which hammer comes next,
  and what rule every hammer must obey. Decide that builds are reproducible —
  not build one, decide the property.*

Now lean Lisp against that ladder. A function is A. A library is B. **A macro
is C** — because a macro does not do the work, and does not merely speed the
work; it changes *what the work is made of*. It reaches up one level and edits
the language you will be thinking in tomorrow.

Which yields the whole analogy in one line:

> **Macros are Lisp's C-activity. Metametamacros are Metametalisp's.**
> Same rung, different ladder — one leaning against space, the other against
> time.

### The obligatory self-reference, since the subject is Lisp and it would sulk

Lisp has been insufferably self-aware since roughly 1960 — see §0 — and has
spent the decades since describing itself in terms of itself while gently
informing every other language that it is a subset. The crystallized form of
the smugness is Greenspun's Tenth Rule:[^greenspun]

> *Any sufficiently complicated C or Fortran program contains an ad hoc,
> informally-specified, bug-ridden, slow implementation of half of Common
> Lisp.*

The Metametalisp corollary writes itself, which is either evidence for the
thesis or a warning about it:

> **Any sufficiently long project contains an ad hoc, informally-specified,
> bug-ridden, slow implementation of half of Metametalisp.**

You have seen this implementation. It is spelled *"the checklist nobody wrote
down,"* *"the thing Dave always reminds us about,"* *"why we don't deploy on
Fridays,"* and *"ask in Slack, someone will remember."* Every one of those is a
metametamacro that was never given a name, never given a definition, and
therefore expands **inconsistently at every call site** — which is precisely
the bug macros were invented to fix, one ladder over.

The industry even has a famous undeclared metametamacro with its own law named
after it. Conway observed in 1967 that organizations ship their communication
structures: the org chart expands into the architecture *whether or not anyone
invokes it*. That is a macro nobody wrote, running in production at every
company on earth, unversioned, unreviewed, and undeniably deployed. Metametalisp
is, in one sentence, the proposal that such things be **declared at the top of
the file, before use** — like any other macro whose silent expansion you are
tired of debugging.

---

## 1.5. On the name, which is not a typo, and whose prefixes are load-bearing

Yes. Two of them. On purpose. Count with me, and do not look away.

**Lisp** is a language that manipulates programs, which is already the meta
move — code as data, the snake taking a thoughtful first bite of the tail.
**Metalisp** would therefore be the discipline of reasoning about that
manipulation — the snake noticing it has a mouth. **Metametalisp** is the layer
where you schedule *when to do the reasoning about the manipulation* — the
snake consulting a calendar to determine the optimal hour at which to notice it
has a mouth. The prefixes are not enthusiasm. They are an inventory. Each names
a real floor of a real building, and the building has a basement we are not
discussing today.

(It must also be admitted, in the interest of the full disclosure this document
keeps demanding of everyone else, that "Metalisp" was already taken. Of course
it was. The good names are always taken — naming being the one domain where the
cost-of-delay curve is common knowledge and everybody still shows up late. The
philosophy was thus forced, by the global namespace, to practice itself, which
it accepts with as much grace as a philosophy can.)

Three consistency checks, offered so the madness can be verified rather than
merely trusted:

1. Engelbart, working from human organizations rather than compilers, arrived
   independently at the same floor and called it **C-activity**. Two people
   counting to three by different routes and landing on the same tread is not
   proof, but it is the kind of coincidence that stops being one.
2. This document is *itself* a member of its own subject matter — see §0, where
   Gödel has already inspected the paperwork. One clipboard visit per essay is
   the regulation maximum.
3. There is no Metametametalisp, and there will not be, because at that height
   the work becomes *deciding when to decide when to decide*, at which point
   one has not built a philosophy — one has built a very expensive way to not
   build a horseshoe. **The tower terminates because the ground floor is where
   the anvil is.** Everything above the third storey is scaffolding for
   scaffolding, and scaffolding shoes no horses.

If, having read all that, the name still strikes you as insane: correct. It is.
It is also *precisely as insane as the thing it names is true*, and those two
quantities being equal is the only defence a word of this length can mount.
Anything shorter would have lied about the floor count.

---

## 2. The blacksmith, the three plates, and the bootstrap that seems impossible

A blacksmith is a person whose job is to make other jobs easier. The plough,
the nail, the hinge, the horseshoe: none of them are the smith's work. They are
the *conditions* of everyone else's work.

And yes — smiths make hammers. They make tongs, too; forging your own tongs is
a traditional apprentice piece, and it is deliberately *first*, because you
cannot comfortably hold hot metal until you have made the thing that holds hot
metal. Which raises the obvious Catch-22 immediately and correctly: **how do
you make the first tongs without tongs?**

Badly. That is the answer, and it is a much better answer than it sounds.

You make the first tongs badly — with a stone for a hammer and a rock for an
anvil and a green stick and a burnt hand. Then the bad tongs make adequate
tongs. Then the adequate tongs make good ones. There is no chicken and no egg —
there is a **ratchet**, and every turn of it is powered by a tool too crude to
have been worth building on its own merits, justified only by what it made
possible next.

Software's whole history is this ratchet wearing different hats. The first C
compiler was not written in C; then it was, and the crude bootstrap compiler
was melted down like the apprentice tongs. Rust's compiler began life in OCaml
— a borrowed hammer — until Rust could forge Rust. Git spent its first week of
existence hosting *its own development*, a version-control system versioning
itself before it versioned anything else. In each case the first tool was bad
on purpose, used anyway, and improved *while in use* — which is the entire
trick, and the reason the paradox dissolves on contact.

### Whitworth's three plates, or: precision from nowhere

The most beautiful version of the bootstrap is not from smithing but from
machining. To make anything precise, you need a flat reference surface. To make
a flat surface, you need... a flat surface. Joseph Whitworth's answer, circa
1830: take **three** plates and rub them against each other in rotation — one
against two, two against three, three against one. Any high spot on any plate
betrays itself as interference against the other two, and gets scraped away.
There is no external standard anywhere in the process. **Flatness emerges from
mutual comparison alone**, and the plates converge on a precision none of them
started with.

Read that paragraph again with test suites in mind. That is what snapshot
tests are — what every characterization test of a legacy system is. There is
no external oracle for "did this change alter behaviour," so you rub the code,
the recorded output, and the new output against each other, and truth
precipitates out of the disagreements. Precision from no precision. Half the
industry runs the three-plate method daily without knowing its name, which is
the usual fate of good ideas from 1830.

One caveat, filed here and detonated properly in §7: Ken Thompson demonstrated
in 1984 that the three-plate method can be run by a con artist.[^thompson] A
compiler taught to recognize itself can propagate a lie through every
generation of the bootstrap, all plates agreeing perfectly, all wrong together.
**Mutual verification converges on *consistency*, not truth.** The plates can
only tell you that you disagree with yourself; they cannot tell you that you
agree with the world. Remember this. The essay will need it when it starts
arguing against itself.

---

## 3. The mechanism (no vibes)

Metametalisp is easy to mistake for "build foundations first," which is a
slogan, and slogans are how projects die in the workshop. The real mechanism is
narrower and sharper:

> **Different work has different cost-of-delay curves. Some costs are flat.
> Some rise. Do the rising ones first.**

Documentation of a stable system costs the same in month two or month twenty:
*flat*. A public API costs one design review before anyone depends on it and a
multi-year deprecation campaign afterwards: *rising*. Hyrum's Law[^hyrum] is
the industry's grim actuarial table for exactly this: with enough users, every
observable behaviour of your system will be depended on by somebody — which
means **the moment you ship an interface is the moment the option to change it
expires.** The curve does not rise gently. It goes vertical at the instant of
contact with strangers.

The industry's museum of vertical curves is well stocked. IPv4 allocated
thirty-two bits of address space in 1981 — a perfectly sensible Tuesday
decision for an experiment, except the experiment escaped the lab, and the lab
has spent thirty years failing to migrate to the fix; the experiment is now the
lab. Two-digit years were a *rational* economy in 1965, when memory cost real
money — a standing decision that expanded silently at every call site for
thirty-five years, until all the call sites detonated on the same midnight, and
the cleanup bill ran to hundreds of billions. Y2K was not a bug. **Y2K was a
metametamacro whose authors underestimated their own longevity** — the most
expensive lesson in expansion counts ever invoiced.

So Metametalisp's actual scheduling rule is not "foundations first." It is:

> **Irreversibility first.** Rank work by how expensive it will be to *unmake*,
> not by how foundational it feels.

This is why one sentence of interface doctrine can outrank a month of
optimization. The optimization is reversible — profile it, undo it, redo it,
nobody outside the process ever knows. The doctrine is load-bearing on
everything written after it, and every line written without it is a line that
must someday be rewritten with it.

And this reframes the frontloading that the philosophy's critics find so
suspicious. You are not front-loading *effort*. You are front-loading
**decisions whose cost curve is about to go vertical**, and deliberately
deferring everything whose curve is flat — which is why doing the unglamorous
internals before the demo was never asceticism. It was **arbitrage**: buying
decisions while they are cheap, in a market where everyone can see the price
rising and most participants buy late anyway.

One civic duty follows, and it is Chesterton's fence with the polarity
reversed: since your standing decisions will outlive your tenure and your
memory, **declare them with their reasons attached** — so that the engineer of
2031, finding the fence, can distinguish load-bearing doctrine from fossilized
accident, and demolish accordingly. An undocumented metametamacro does not stop
expanding when its reason dies. It just stops being *right*.

---

## 4. Evidence from industry, since a philosophy should be falsifiable

Metametalisp has a shameful twin — *infrastructure as procrastination* — and
from the outside, on any given Tuesday, the two are indistinguishable. Both
produce commits, diagrams, tooling, and a warm sense of rigour;
procrastination-as-craftsmanship has **all the outputs of virtue**. So the
honest test is behavioural, not aesthetic: *was the lever smaller than the work
it unlocked, and did its own builder lean on it immediately?* Five exhibits,
chosen because each answers yes on the record:

**Exhibit A: the Bezos API mandate, 2002.** A memo of roughly six bullet
points: all teams will expose data and functionality through service
interfaces; all communication happens through those interfaces; no back doors,
no shared databases; the interfaces must be externalizable; and — the
enforcement clause that makes it a macro rather than a suggestion — anyone who
doesn't comply is fired. Note what did *not* ship: no tool, no platform, no
framework. **A pure standing decision**, smaller than a sprint, that expanded
at every team boundary in the company for two decades. One of its expansions is
called AWS. The memo's expansion count is now denominated in fractions of the
global economy, which is a strong quarter for six bullet points.

**Exhibit B: Git, 2005.** BitKeeper's licence evaporated and the Linux kernel —
possibly the largest collaborative engineering effort in human history — lost
its version control overnight. Linus Torvalds wrote git's usable core in
roughly a fortnight and was hosting kernel development on it within days: the
tool's first user was its builder, the dogfood interval was measured in hours,
and the lever was *radically* smaller than the work it unlocked — which turned
out to be not just the kernel's future history but, in time, nearly
everyone's. The purest recorded specimen of the microscope justified by the
dissection, built mid-dissection, with the patient still open.

**Exhibit C: SQLite.** The most deployed software artifact on earth — it is in
your phone, your browser, your car, and statistically your toaster — is
maintained by a handful of people who change it constantly and fear nothing.
The reason is a test harness of aviation-grade paranoia:[^sqlite] hundreds of
lines of test per line of library code, full branch coverage, deliberate
out-of-memory and I/O-failure injection. The microscope is so thorough that
the dissection became *boring* — and boring is the entire point. Every future
change, by anyone, forever, inherits the discount. That is not testing as
insurance. That is testing as **compound interest**.

**Exhibit D: Rust's borrow checker.** A single standing decision — *aliasing
XOR mutation, checked at compile time* — priced entirely up front, in the
famously steep currency of learning it. The expansion happens at every line
anyone will ever write in the language: an entire class of memory-corruption
bugs made not merely unlikely but *inexpressible*. The compiler yells so that
production does not. Pay once, at the layer of the language itself; collect at
call sites that will not exist for decades.

**Exhibit E: Semantic versioning.** Three integers and a promise about what
each may mean. No code shipped, ever. Yet the decision expands at every
automated dependency resolution on the planet — every upgrade that proceeded
unsupervised, every 3 a.m. that stayed quiet — an ecosystem-wide metametamacro
adopted by strangers who will never meet, expanding identically at every call
site *because it was declared*. Compare Conway's law, its undeclared cousin
from §1, expanding just as universally and twice as chaotically, and the case
for writing these things down makes itself.

The common signature across all five, stated once so §5 can weaponize it:
**the lever was small, the leverage was vast, and the builder was the first
one hanging off it.** Hold every proposed microscope to that signature. Most
will not survive the comparison, and that is the comparison working.

---

## 5. The three tests (use these when you are unsure)

Before building the microscope, interrogate it:

1. **The rising-cost test.** Can you name the *specific* thing that becomes
   more expensive with every unit of delay? Name it out loud, with a number if
   possible. If you cannot, this is not frontloading — it is gold-plating
   wearing a lab coat.
2. **The proportion test.** Is the tool smaller than the work it enables? A
   microscope larger than the dissection is not a microscope; it is a hobby.
   The industry's warehouses are full of the counterexamples: the two-year
   internal developer platform serving three services; the Kubernetes cluster
   lovingly provisioned for a static site; the in-house framework that
   consumed the product it was built to accelerate and then outlived it, like
   a tomb.
3. **The dogfood test.** Will *you* be its first user, this week? Not a
   persona, not "the team, eventually," not the hypothetical engineers of Q3 —
   you, now. A tool built for a hypothetical future user is a gift wrapped for
   a stranger who may never arrive, and the wrapping is where the years go.

Two out of three usually suffices. Zero out of three means you are not
preparing for the patient. You are avoiding the patient.

And one field mark for distinguishing the real instrument from the counterfeit,
since the counterfeit produces identical paperwork:

> **The fake microscope keeps acquiring lenses. The real one keeps acquiring
> specimens.**

If the tooling's changelog is longer than the list of things the tooling has
been aimed at, the diagnosis writes itself.

---

## 6. Lineage (this was rediscovered, which is the best kind of discovered)

Douglas Engelbart, 1962, called all of this **bootstrapping**, and his entire
research program was the claim §1's ladder was borrowed from: **C-activity
compounds and everything else does not.** He was arguing it about human
organizations, with no compiler in sight — and then he ran the experiment on
his own lab, spending a decade improving the improvers, which erupted in 1968
as the Mother of All Demos: the mouse, hypertext, live collaboration, video
conferencing, all in one afternoon, all of it the *expansion* of ten years of
standing decisions about how a laboratory should augment itself. The audience
thought they were seeing products. They were seeing a macro's output.

McCarthy supplies the other parent: `eval`, and with it the proof — see §0 —
that a system can host its own C-activity, that the ladder can be climbed from
inside the building. Hamming, in 1986, supplied the compound-interest framing
this essay opened with, and the uncomfortable observation that most careers
decline the interest. Knuth contributes the strangest and most instructive
specimen of all: TeX's version number converges to π, one digit per revision,
by *standing decision* — a metametamacro whose expansion is the guarantee that
nothing will ever change again. Deciding what is finished, it turns out, is
also a decision with an expansion count; stability is C-activity too, and
possibly the rarest kind.

What Metametalisp adds to this lineage — its one genuine contribution, and it
is not nothing — is the observation that **the discount structure is exactly a
macro's**: pay once at the right layer, benefit at every call site, including
call sites that do not yet exist. That is why the unit deserved a name of its
own. A **metametamacro is a standing decision with an expansion count**, and
the reason one sentence of doctrine can outrank a month of engineering is
arithmetic, not aesthetics: its expansion count is larger. Spatial and temporal
composition turn out to be the same trick applied to different axes. Engelbart
found the rung; McCarthy found the mechanism; the essay merely introduces them
and stands back.

---

## 7. The antithesis, honestly (a philosophy that cannot lose is a mood)

**Worse is Better.** Richard Gabriel's 1991 essay[^gabriel] is the direct
historical antagonist of everything above, and it is uncomfortable because it
is largely right *about outcomes*. The MIT school — correctness first,
completeness, do the hard foundational thing properly — produced better
artifacts. The New Jersey school — simple implementation, ship it, let the
users sand it down — produced C, Unix, and, with a thirty-year lag and a
hoodie, the modern industry, whose motto "move fast and break things" is
Worse-is-Better's IPO. The philosophy this essay advocates has a **losing
commercial record against the philosophy it critiques**, and any presentation
of Metametalisp that omits this fact is selling something. The elegance of the
approach is not evidence of its market success. It may be mildly negative
evidence.

**YAGNI, and the whole Lean edifice.** You cannot know which levers you will
need until the work tells you; building levers for anticipated work is guessing
with extra steps and better lighting. Most dead side projects died in a
spotless workshop, surrounded by beautifully organized tools, having never made
a horseshoe. The dominant industry philosophy is dominant partly because it is
*right about most people, most of the time* — a sentence this essay is obliged
to type with its own hands.

**Theory of Constraints.** Goldratt's rule is sharper than Metametalisp *as a
scheduler*: improve only at the bottleneck; every improvement elsewhere is,
by definition, waste. And when the two disagree, TOC usually wins, because TOC
is **measured** and Metametalisp is **anticipatory** — one reads the gauge,
the other reads the future, and gauges have the better track record. In
particular: if the bottleneck is *demand* — if nobody wants the thing yet —
then no quantity of exquisite tooling moves the constraint one millimetre,
and Metametalisp will happily hand you a year of immaculate preparation for a
product with zero users, invoiced monthly, in good faith.

**Real options.** Deferring a decision has positive value while information is
arriving; finance has priced this precisely, and the price of deciding early is
real. Metametalisp says decide early; option theory says decide *late, but
cheaply*. The theories reconcile at exactly one point, and it is the essay's
own §3: decide early **only where the option is about to expire** —
irreversibility is the expiry date, and everything without one should be
deferred with a clean conscience.

**And the Thompson objection, promised in §2.** The three-plate method — all
those tests, snapshots, and self-consistent instruments — converges on
*coherence*, and coherence is structurally incapable of detecting the wrong
product built correctly. Your suite can prove you disagree with yourself; only
contact with an outside world your instruments did not author — users,
production, reality — can prove you disagree with the *truth*. A philosophy of
instrument-building must therefore hold one door permanently open to things
that can surprise it, or it will polish itself, flawlessly, into irrelevance:
consistent, verified, and wrong, with all the plates agreeing.

---

## 8. When Metametalisp is wrong

- **Short horizon.** Compounding needs runway. If the project might be
  abandoned in six weeks, skip the tongs and make the horseshoe with your
  hands; the burn heals faster than the opportunity.
- **Unknown requirements.** Before product-market fit, the constraint is
  *learning*, not leverage — and a team that does not yet know what it is
  building will construct the wrong microscope, beautifully, and then be
  reluctant to discard it for exactly as long as it was expensive.
- **The bottleneck is elsewhere.** Especially when the bottleneck is users.
  Re-read the Goldratt paragraph until it stops being pleasant.
- **Solo and team psychology, the quiet killer.** Tool-building is legible,
  controllable, praised in standup, and safe from the one judgment that
  matters; shipping to strangers is none of those things. The forge is
  emotionally comfortable, and the comfort disguises itself as rigour so well
  that entire platform teams have reorganized into permanence without ever
  noticing they stopped serving anyone. The temptation to stay at the anvil is
  not intellectual. It is warm in there.

Which yields the single corrective the whole essay bends toward, the sentence
to frame if only one gets framed:

> **Metametalisp schedules the work; it does not choose the work. The
> microscope is justified by the dissection, and only by the dissection.**

A project that has been in the workshop for a year has either been executing
this philosophy perfectly or hiding inside it completely, and no inspection of
the workshop can tell you which — the commits look identical, the diagrams
look identical, the virtue looks identical. The only test is traffic: are the
tools being leaned on, by someone doing real work, soon, and is that someone
*you*.

---

## 9. Verdict

Metametalisp is not superlative in general, and the essay will not pretend
otherwise at this late hour. It is superlative **under nameable conditions**:

| Condition | Why it matters |
|---|---|
| Long horizon | Compounding needs time to compound; interest is not paid weekly |
| High rework cost | Rising curves are the entire mechanism; where retrofits are cheap, so is delay |
| Foundations that will not move | Levers need a fulcrum; shifting ground forgives the unprepared |
| The tool-builder is the tool-user | The dogfood loop closes in hours instead of quarters |
| A compounding domain | Languages, compilers, protocols, standards, libraries — the build-upon-ables |

Change three of those rows and the correct philosophy becomes Worse is Better,
and it would not be close. Keep them all and the correct philosophy becomes
this one, and *that* would not be close either.

So the honest formulation is not "Metametalisp is the best way to run
projects." It is:

> **Metametalisp is the correct philosophy for building things that other
> things will be built on — and the wrong one for building things that stand
> alone.**

Languages, corpora, compilers, platforms, standards: the most
build-upon-able artifacts in the trade, the ones whose every early decision is
inherited by strangers at compound interest. For those, front-load the
irreversible, declare the standing decisions, build the plates, and let the
expansion counts do the arguing. For everything else, make the horseshoe. The
philosophy fits the material or it does not — which is, fittingly, exactly
what a blacksmith would tell you: **the tool is chosen by the metal.**

---

## Ω. Postscript on the first hammer

The bootstrap paradox deserves a proper ending, so here it is: the first hammer
was a rock. The first anvil was a bigger rock. Every precise, hardened,
beautiful tool in the world is descended from a rock that somebody was
disciplined enough to *keep using* while making something better with it.

The lesson is not that you must begin with excellent tools. Nobody ever has.
The lesson is that you must begin with **bad ones, on purpose, and improve them
while in use** — which is precisely what a crude shell script, a hand-run
checklist, and two ugly snapshot files always turn out to have been, in every
project that ever grew up: the stone tongs, still warm, melted down and
forgotten by the very hands they made possible.

Room status: still spinning. Building status: still standing. Anvil status:
ground floor, where it has been the whole time.

---

[^church]: Alonzo Church's lambda calculus was proven equivalent in power to
Turing's machines in 1936, whereupon history named the joint result the
Church–Turing thesis and then proceeded to talk almost exclusively about
Turing. Lisp is, structurally, Church's revenge. This essay is merely honoring
the original billing order.

[^engelbart]: Douglas Engelbart, *Augmenting Human Intellect: A Conceptual
Framework* (1962) — the report in which improving-the-improvers was proposed
as a research program rather than a personality flaw.

[^greenspun]: Philip Greenspun, rule the Tenth. There are no rules one through
nine. This is widely considered the most Lisp fact about the rule.

[^thompson]: Ken Thompson, *Reflections on Trusting Trust* (1984 Turing Award
lecture): a compiler modified to insert a back door when compiling the login
program, and to insert *both behaviours* when compiling itself — after which
the source can be made innocent again and the lie survives every rebuild,
invisibly, forever. The three plates, run by a con artist.

[^hyrum]: Hyrum Wright's observation, from deep inside Google's API
maintenance mines: "With a sufficient number of users of an API, it does not
matter what you promise in the contract: all observable behaviors of your
system will be depended on by somebody." The law is descriptive, the tone is
resigned, and the curve, as noted, is vertical.

[^gabriel]: Richard P. Gabriel, *Lisp: Good News, Bad News, How to Win Big*
(1991), containing the "Worse is Better" section its author has spent three
decades alternately defending and refuting, sometimes within the same paper —
which is either intellectual honesty of the highest order or the longest-running
three-plate experiment in the philosophy of engineering.

[^sqlite]: The SQLite project's published testing documentation reports test
code outweighing library code by a ratio in the hundreds to one, with 100%
branch coverage and deliberate fault injection — figures normally associated
with avionics, which is fitting for a database that is, among its many other
deployments, literally flying.

[^hart]: Timothy P. Hart, "MACRO Definitions for LISP," MIT AI Memo 57,
October 1963. Four pages. The interpreter it extended was three years old and
had not been asked. The entire subsequent history of Lisp macros — and
therefore of this section — descends from a user deciding that a finished
system was an invitation.

[^aristotle]: Aristotle, *Physics* III.6: the infinite exists potentially,
never actually — there is always more, but never all. Proposed as a resolution
to Zeno, adopted here as a build policy. The stairwell is the first known case
of a document citing the *Physics* as its capacity-planning document, a
distinction Aristotle would have accepted with the weary grace of a man whose
lecture notes have been load-bearing for twenty-three centuries.
