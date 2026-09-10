# OMEGA

*Not the last letter. The first ordinal.*

> "The essence of mathematics lies precisely in its freedom."
> — Georg Cantor, 1883, in the paper that built the number system this file is indexed by[^cantor]

*One epigraph. The antithesis goes in the text, where it can be reached.*

---

## ω. What the family is, measured

Three documents preceded this one, and each did one thing well.

METAMETALISP4.md proposed a unit — the *metametamacro*, a standing decision with an expansion count — and a scheduling rule: rank work by how expensive it will be to unmake, and do the irreversible things first. Then it applied the rule to itself three times (§0, §−1, §−1.5), each application producing a section about the previous one, and called the result a stairwell.

METAMETAMETALISP.md hung a gallery beside the essay, one annotation per section, and gave the prefix META a coordinate: multiplication by *i*, a quarter-turn on the complex plane, order four, so that four metas cancel and the tower "terminates by algebra." It also confessed the family's seam: none of the `macroexpand-1` traces had ever run. They were depictions.

POSTSCRIPT.md, under a Kierkegaardian pseudonym, ran a seven-movement apparatus over every section, watched the apparatus fail at four of them, and delivered the family's strongest sentence about itself: that a concession which changes nothing is an inoculation, and that ending is an act this authorship has built four consecutive architectures to avoid.

That is the record. Here is what the record does not say about itself.

**All three open by quoting their commission.** The gallery begins with "a reader then asked." The Postscript begins with "the reader has commissioned." This file was commissioned to analyze the three and outclass them in their own style and domain, and that commission is reproduced here, at the top, before use, because the doctrine says to. Read together, the family is not an essay and its commentaries. It is a transcript: a request, an expansion, a request, an expansion. The essay wrote "`eval` does not run — `eval` is *run*" about the reader who runs it, and did not notice it had described its author. The author of each of these documents is the expander. The fiat has been at the REPL since the first keystroke.

**The house style is the family's undeclared metametamacro.** Status lines. The personified dead — Gödel with a clipboard, Aristotle with lecture notes. A bolded aphorism every few hundred words. "Attributed to nobody, constantly." A notarized trace that never executed. Nobody declared this style; it expanded at every call site whether anyone thought about it or not, which is the essay's own definition of the thing it exists to warn about. This file declares the style here, uses it through §2, and suspends it from §1 on — because a standing decision you cannot suspend is not one you hold. It holds you.

**The gauge.** The essay is 50,936 bytes. The gallery is 17,589. The Postscript is 194,686 — 3.8 times the document it answers, having opened by ridiculing the proportion test. Every expansion in this family is larger than its call and contains another call to itself. The essay knew, in §−1.5, that a form of that shape does not terminate under the full expander; it wrote three notes for the record and kept going. This file is smaller than the essay, not merely smaller than the Postscript, and from §5 down each section is shorter than the one before. Measured after the last edit, not promised before it: §5 is 1,631 words, §4 is 1,078, §3 is 930, §2 is 476, §1 is 344, §0 is 15. That is a ranking function wearing a word count, and it is the only kind of promise a document can keep on behalf of its successors.

**The index set.** This file is numbered by ordinals: ω, then 5, 4, 3, 2, 1, 0. From ω you may step to any natural number; nothing forces the choice of 5, and the choice is marked as unforced. After it, every step is forced, and 0 has no predecessor in this set — not by fiat, by definition. The reader is at 0. The document ends on them.

Four claims follow, one per finite section, each about a choice the family made without noticing it was choosing. The essay named its unit after the wrong feature of its favourite language. The Postscript filed the essay under the wrong Kierkegaardian stage and never named its own. The gallery proved termination for an operation it did not perform. And all three chose a number system in which ending is impossible, then hired Aristotle to make that sound like wisdom. Everything this file needs in order to say so was already in print — 1883, 1940, 1949, 2015 — which is the family's most reliable property: what it needs is always already published, and it writes instead.

Status lines: declared here, discontinued after this one.

---

## 5. The wrong feature

A macro is a source-to-source transformation, invoked by name at its call site, that adds a form to the language. After `defmacro`, every program that was legal is still legal, and some new ones are. A macro that is not called does nothing. Nothing about it is imposed; each call site opts in, by spelling its name. Macros are conservative extensions. That is not a detail of the analogy. It is what a macro *is*, and everything the essay says about leverage — pay once, collect at every call site, including sites nobody has written yet — inherits it.

Now read the essay's three exemplary standing decisions with that in hand.

*Nothing merges without review.* *No bug is fixed without a failing test written first.* *The build must be reproducible, or it does not count as a build.*

Each of these removes something. Each is imposed on every site in scope, none of which spelled its name. None can be declined by the site it applies to. That is not a macro's shape. It is a *type's*: a restriction on which programs are legal, declared once, checked at every site, by a checker the sites did not invoke. The essay's exhibits confirm it one at a time. The Bezos mandate is an interface discipline plus a prohibition — no shared databases — which is a type, and a lint that fires you. Semantic versioning is a nominal type for version numbers. SQLite's harness is a specification checked on every change. Rust's borrow checker is a type system, literally; and the essay, describing it, wrote that a class of bugs had been made "not merely unlikely but *inexpressible*." That is the type theorist's word. The essay used it and did not hear it. (Git, the fifth exhibit, is a tool the Postscript already caught being built under duress; it is the exhibit that proves the other four are not tools.)

Why reach for macros, then? Because macros are the powerful thing Lisp has, and a static discipline that rules programs out is the thing Lisp famously lacks. The essay named its unit after the feature its favourite language is proudest of, in order to describe the feature its favourite language is missing. Greenspun's rule inverts cleanly: any sufficiently long Lisp project contains an ad hoc, informally specified, bug-ridden, slow implementation of half of a type system — `check-type` at every entry point, `declaim` in every file, and eventually Coalton. (That a checker can be built out of macros — Typed Racket is one — does not make the checker a macro, any more than the tongs are the horseshoe. The essay's own §2 has the distinction and did not apply it here.) And the footnote about Church picked the wrong revenge. Church's 1936 calculus is untyped. His second act, in 1940, was *A Formulation of the Simple Theory of Types*:[^church] the paper in which he restricted his own invention so that it could be trusted. The essay cited the Church who made everything expressible and missed the Church who made some things inexpressible on purpose, which is the only Church who would have understood what the essay was trying to say.

The rename is not cosmetic. Six things follow from it, and each repairs a place where one of the three documents was standing on the wrong noun.

**Expansion is opt-in; checking is not.** The Postscript objected, in ethics, that the expansion count measures how many people are bound rather than how much value is delivered — that the Bezos memo was one man's Tuesday becoming several thousand people's decade. Under the rename, that objection is a fact about the object rather than a complaint about its use. A macro's call sites chose it. A type's sites are chosen by it. Semver's dependents did not invoke semver; the resolver applied it to them. The essay's category error and the Postscript's moral objection are the same error, seen from the language side and from the human side.

**Failures are global, not local.** A macro with a bug produces bad code at each expansion independently; nothing synchronizes the failures. A type with a bug fails everywhere the type is used, at once, because the checker is one thing. Now reread the essay's museum piece: two-digit years "detonated on the same midnight." Simultaneity is the fingerprint of a type error when the checker is the calendar. The essay's own exhibit carries the signature of the object it misnamed.

**Hyrum's law is duck typing.** Clients infer the most specific type from the implementation and depend on it. The declared contract is a comment; the inferred type is the real one; and dynamic clients always win, which is the entire content of the law and the reason its tone is resigned. Semver is the industry's attempt at nominal typing for change, and it works exactly as well as nominal typing works against clients who inspect. The only move that beats Hyrum is the type-level one: make the behaviour you do not want to promise *unobservable*. Go randomizes map iteration order on purpose, so that no program can come to depend on it.[^go] That is not documentation. That is a type.

**The expansion count is a multiplier, and it has a sign.** The value of a standing decision is the sum over its sites of benefit minus cost, and the count only scales whatever sign that difference has. The essay says Y2K's authors "underestimated their own longevity." True — but longevity is not the count; it is the horizon, and what the horizon determines is the sign. A type with large coverage and a wrong definition is the most expensive object in software. A macro with a wrong definition is a bug in one file. When the essay wrote that one sentence of doctrine outranks a month of engineering because "its expansion count is larger," it was doing scalar arithmetic on a signed quantity, and calling it "arithmetic, not aesthetics" did not supply the sign.

**The essay's two halves are about two different objects.** §1 — the tool that must be picked up versus the macro that runs while the team is on holiday — is about *procedural* decisions: rules consulted per occasion. The review rule. And those are precisely the decisions that *can* be set down: not deleted, shadowed, the way a special variable is shadowed by an inner `let`. The sprint deadline shadows the review rule for the extent of the sprint, and from outside the extent the rule looks intact. §3 — irreversibility, Hyrum, IPv4, Y2K — is about *structural* commitments: types you have let strangers infer. The review rule is reversible tomorrow; thirty-two bits of address space were not. Irreversibility-first is correct about the second object and silent about the first. The tool-versus-macro distinction is correct about the first and irrelevant to the second. The doctrine is stitched across that seam, and the Postscript's observation that "CI blocks the merge" names a state with a budget is the seam showing.

And where the macro reading does hold — Dave's checklist, the thing nobody wrote down — the failure mode has had a name since 1986: capture.[^hygiene] An unhygienic macro's free identifiers are bound at the expansion site. *Review*, *test*, *reproducible*, *build* are free identifiers in every standing decision the essay lists, and they will be bound in 2031 by 2031's environment. "Declare them with their reasons attached" is an attempt to ship the environment along with the decision — a closure. A comment is a closure that drops its environment the moment someone in a hurry reads it, which the Postscript described from the inside and could not explain.

**The enforcement ladder.** Put the essay's evidence and its doctrine on one scale and they are not on the same rung.

1. Tribal. Dave remembers.
2. Written. A comment, a wiki page, a shirt.
3. Reviewed. A human looks, when a human is looking.
4. Mechanical. A machine in the path refuses.
5. Inexpressible. The language cannot say the wrong thing.

The doctrine — "declare the standing decisions with their reasons attached" — is rung two. The evidence is rungs four and five, every exhibit. "Why we don't deploy on Fridays" is rung two; it is on shirts; it is violated weekly; and it is violated because rung two is a comment and the engineer of 2031 skims comments, which is the Postscript's lament and, under the rename, a theorem. The practical correction is one line, and the essay's own exhibit supplies it: **a standing decision that nothing in the path checks is a comment.** The metametamacro was never the sentence "nothing merges without review." It was CI. The Postscript called CI a state with a budget and a constituency, and it is — so is a compiler — and the whole economic fact of this trade is that in it, the checker is cheap and already in the path. That is the one thing software has that organizations do not, and it is what the essay was reaching for when it reached for the wrong noun.

Because here is what the essay wanted, in its own words: something "still running while the whole team is on holiday." A macro does not have that property. Its transformation has already happened, at compile time, once per site, and is gone. A type checker has it. It does not take holidays. It checks the same rule on the ten-thousandth site as on the first, for people who never read the declaration and would not thank the person who wrote it. The essay had the right dream. It wrote the wrong noun on it, then built a scheduling philosophy on the noun's properties — opt-in, local, additive — when the dream's properties are the opposite on every axis.

---

## 4. The wrong stage

The Postscript's central charge, the one it said it would defend under oath: the macro's defining virtue — its leverage does not depend on anybody remembering it exists — is the definition of despair. A relation that no longer has to relate itself to itself, because the relating has been compiled out. The project of arranging one's life so that nothing further need be chosen. "The highest form of the aesthetic."

That is a misfiling, and the misfiling is visible in the Postscript's own interlude.

*Either/Or* has two volumes.[^either] Volume I belongs to A, the aesthete, and contains two things the Postscript reproduced without attribution. The first is the Diapsalmata's litany — marry and you will regret it, do not marry and you will regret it — which the Postscript rewrote as "build the tool and you will regret it; do not build the tool and you will also regret it," in bold, in its own voice, and called "the sum of all practical wisdom." That is A's line. It is the sum of *A's* practical wisdom, and Kierkegaard put it in the aesthete's mouth so that the reader would learn the sound of a man who has arranged never to be caught choosing. The second is "The Rotation of Crops": A's technique for a life without boredom, which is to commit to nothing, vary the field, keep every option live, and treat every binding as boredom's front door. That is real-options theory practised as a soul. It is YAGNI as a way of being a person.

Volume II belongs to Judge William, and its argument is that the self is not something you have but something you choose, and that the choice which matters is the one that binds every subsequent moment. His example is marriage. His claim is that the person who chooses once, decisively, so as not to have to choose again on the same question, has not abolished the self. He has acquired one. "Relates itself to itself," in the Judge's hands, is done by binding. A self is the shape left by its standing decisions.

So: a promise is a type. A marriage is a metametamacro, and Judge William would say the term is his. The essay — commit early where it is irreversible; declare it; let strangers depend on it; accept that they will not thank you — is the ethical stage, in build tooling. It is Volume II. The Postscript wrote forty thousand words holding seven relations to the same object "simultaneously, none cancelling any other," declined to choose between gratitude and resentment and "in the manner of the age" published both, and then accused the essay of mediation. Seven fields, never the same one twice. The Postscript is Volume I. It filed Volume II as Volume I, and it made the misfiling in the aesthete's own idiom, which is the tell.

The Postscript's other thesis — "the interesting number is not the expansion count; it is one" — is right and pointed at the wrong object. The one is the declaration. A type is declared once, by a person, and checked forever, by a machine. The moment of choice is not abolished; it is relocated to the declaration site, which is the one place it costs the most, and then it is not repeated, which is the whole difference between a promise and a mood. The Postscript's complaint that the standing decision *works* is the complaint that the person chose once instead of every time. Judge William's reply is that choosing every time is not freedom. It is crop rotation.

Which leaves the third stage, and here the Postscript came within one page and stopped. Kierkegaard's religious is not more ethics. It is the *suspension* of the ethical: Abraham, the exception, who breaks the universal rule on a command he cannot justify and about which he cannot speak. The Postscript listed FEAR AND TREMBLING.md among the books it declined to write, "on the engineer who deleted the framework." It had already been written. It is in the borrow checker's manual, under `unsafe`.[^unsafe]

Precisely, because the precision is the doctrine. `unsafe` does not switch the checker off; the borrow checker still runs on every reference inside the block. It opens five specific doors the checker cannot see through — raw pointers, unsafe calls, mutable statics, unsafe traits, union fields — and makes the person who wrote the keyword responsible for what comes through them. The suspension is local: one block wide. It is lexically marked, which means it can be found; `cargo geiger` counts them. A crate can forbid it outright — `#![forbid(unsafe_code)]` — which is a standing decision about whether exceptions to standing decisions are permitted at all. And the convention is that every block carries a `// SAFETY:` comment stating why the invariant holds: a lint can insist that the comment exists, and nothing can check what it says.

De Silentio's Abraham cannot speak. Rust's Abraham cannot speak either — the `SAFETY` comment is prose; nothing verifies it — but he must *sign*. That is the improvement over 1843, and it is the doctrine that sits past all three documents, because each of them needed it and none had a word for it:

**The marked exception.** You may break the standing decision. You may not hide that you did. The ethical keeps running — the checker stays on — and the exception is one keyword wide; its justification is a comment; its *existence* is a fact the whole codebase can see.

Most organizations have the inverse. Rules that are silently violated: unmarked exceptions, shadowed bindings, "we all know the Friday thing is aspirational." And exceptions that are loudly justified after the fact: retrospectives, post-mortems, a blog post. Justification without marking — Abraham with a press release. The essay's civic duty, declare with reasons attached, addressed the builder of the rule. Nobody addressed the breaker. The breaker is the party with the option, in the only moment that matters, and the whole of practical ethics, as the Postscript correctly said of the builder, is addressing the party with the option.

It also settles what §8 meant by "the one judgment that matters," which the Postscript read as the reader. It is simpler and worse. It is the judgment at the `unsafe` block: was the invariant actually upheld? It arrives in production, from strangers, later, and no comment protects you from it. That is the leap. It comes with a receipt.

---

## 3. The wrong number system

The stairwell's argument is that the interval (−2, −1) is dense — between any section and the reader, room for another — and that this "is precisely a macro system: finitely many macros defined, infinitely many definable, and the difference between those two quantities is called a future."

It is not a macro system, because density and unboundedness are different properties and macro space has only the second. The set of definable macros is the set of finite strings over a finite alphabet: countable and discrete. Between two macros there is no third. There is a *next*. The correct picture is ℕ, not ℚ, and in ℕ Achilles catches the tortoise in one step, because there is a step. The Postscript raised this objection under the name "sparsity" and withdrew it when the essay produced Aristotle. It should not have. Aristotle's potential infinity licenses *always one more*. It does not license *always one between*. The essay smuggled betweenness in with the word "interval," and betweenness is the whole of Zeno.

The difference is not academic; it is the family's entire pathology, stated as a choice of index set. The rationals, the reals, the complex plane — every coordinate system the three documents used is a field, and a field with a compatible order is dense: there is always a between, and no descending sequence is ever required to end. The natural numbers and the ordinals are the opposite kind of object. Every nonempty subset has a least element. Every strictly descending sequence is finite. That property has a name in this trade, and the essay had the citation on its shelf and reached past it for the *Physics*: **well-foundedness**. Turing proposed it in 1949 as the way to check that a routine halts — exhibit a quantity that strictly decreases, and take the quantity from the ordinals.[^turing] Floyd made it standard in 1967. It is how structural recursion is known to terminate, how Gentzen proved arithmetic consistent by descending through ε₀,[^gentzen] and how `macroexpand` terminates in every sane Lisp program: each expansion is smaller under some measure, and the measure is the macro author's responsibility.

Which corrects the essay's proudest technical claim. `macroexpand-1` is not the termination discipline. It terminates nothing; it declines to begin. The termination discipline of a macro system is that its expansions are well-founded, and that is not in the standard library. It is a proof obligation on whoever writes the macro — the only sense in which the essay's author was ever the scheduler of anything.

So the most upstream undeclared standing decision in the family is the index set. Choosing ℚ made non-termination a structural guarantee; choosing ℂ extended the guarantee sideways; then Aristotle was hired to make the guarantee sound like wisdom. Markdown refused to number the negative sections and the essay laughed. Nobody noticed that the negatives were never the problem. The denominators were.

The Postscript's Ω makes the same mistake from the far end: "nothing can be inserted after the end of the alphabet." Wrong on both readings. As a Greek letter, Ω is followed by whatever the next author brings, which the Postscript admitted a few paragraphs later. As the ordinal ω — the reading the family should have used — it is not the end of anything. It is the first infinite ordinal; ω + 1 exists; and the interesting property of ω is that anything descending from it must land on a finite number and then stop. From ω you may step to any natural number. That step is unforced; it is a leap, and you must take it. Having taken it, the rest is a theorem.

This is the first document in the family to run a trace. The interpreter was SBCL, the file is sixty-four lines,[^lisp] and the numbers are exact; the output is reflowed for width and otherwise untouched.

```
> (macroexpand '(BISECT -2 -1))  ; budget 8
⇒ DIVERGED after 8 expansions;
  last call site still pending: (BISECT -2 -511/256)

> (macroexpand '(BISECT -2 -1))  ; budget 40
⇒ DIVERGED after 40 expansions;
  last call site still pending: (BISECT -2 -2199023255551/1099511627776)

> (macroexpand '(COUNTDOWN Ω))  ; budget 40
⇒ HALTED after 7 expansions:
  (SPLICE 5 (SPLICE 4 (SPLICE 3 (SPLICE 2 (SPLICE 1 (READER))))))

> (macroexpand-1 '(bisect -2 -1))
⇒ (SPLICE -3/2 (BISECT -2 -3/2))
;; macroexpand-1 did not terminate the stairwell. It declined to begin it.
```

Three notes, since the custom is three. The pending section after forty halvings sits at −2 + 2⁻⁴⁰: the reader approached to within a trillionth, the call still open, exactly as the essay proved and celebrated. `countdown` halts in seven steps — one leap from ω to a number, five forced descents, and the least element, which is occupied. And the reader upcased the ordinal: SBCL's reader folds symbols to upper case, so ω went in and Ω came out — the Postscript's misreading, performed by the reader, which is where misreadings are performed.

What the ordinals give the reader is the last thing. In a well-ordered index the least element exists and is a member. The essay proved the reader unreachable with a Dedekind cut, and that proof needs a field: the reader as an infimum, excluded by the definition of membership, standing in a gap. In ω there is no gap. The reader is not below the floor. The reader is at 0, which is the last section, and the document terminates on them — not annexed, not assigned a job, not described as load-bearing. Addressed, once, at the end. Then the file stops.

---

## 2. The wrong operation

The gallery's proof runs: META = *i*; multiplication by *i* is a quarter-turn; *i*⁴ = 1; therefore four metas cancel and the lateral direction terminates by algebra, "no fiat required."

Now look at the gallery's coordinates. §(−1.5 + *i*). §(0 + *i*). §(Ω + *i*). Those are sums. Rotating −1.5 by *i* gives −1.5*i*, not −1.5 + *i*. The gallery translated its contents and rotated its thesis. Translation by *i* has no order: § + *i*, § + 2*i*, § + 3*i*, an unbounded tower running sideways, precisely the thing the file announced it had escaped. The proof concerns one operation and the file performs another, and a proof about the wrong operation is the family's oldest habit in a new notation.

Even the rotation is a choice. *i* is one point on a circle of operators. Any rotation e^(iθ) does what the gallery asked META to do — stop reading a thing as content, start reading it as subject — and for θ an irrational multiple of π the orbit never returns; it is dense on the circle, which the family would have enjoyed. The gallery picked the one angle that closes in four turns and called the closing algebraic. Choosing the operator that terminates and reporting the termination as a theorem is the move the Postscript identified as the essay's signature — a decision wearing a derivation's coat — performed by the gallery, one wing over, with an evaluation arrow.

And the gallery's own arithmetic, in its own frame, puts it somewhere it did not print. LISP = 1. METALISP = *i*. METAMETALISP = −1. METAMETAMETALISP = −*i*. That is the reader's margin — the station the gallery's table reserved for "the requests that scheduled everything." The gallery was written to answer a reader's question, and its algebra knew. The table it published, with the essay at 1 and itself at *i*, is the same plane rotated a half-turn, and the labels do not survive the rotation: the gallery is the reader's margin wearing the author's nameplate.

Last, the product it did not compute. *i* · *ī* = *i* · (−*i*) = 1. The author's margin times the reader's margin is a claim — real, positive — in one step rather than three. That is the loop that produced every document in the family: commentary, multiplied by a request, becoming content. The gallery said the exit was *i*⁴, "a paper from 1960." The exit was one multiplication away, and it was the reader. z · z̄ = |z|²: you get a real number out of a complex one by multiplying by the conjugate. Measurement requires the reader. And the identity element the gallery called Lisp is, on the family's own ladder, A-activity — the ground floor, the anvil, the horse. The algebra agreed with the Postscript. The gallery did not read its own table.

---

## 1. The loop

The house style stops here. What follows is plain, not for plainness' sake, but because a standing decision you cannot suspend is not one you hold.

The family is a transcript. Each document opens by quoting the request that produced it, and this one does too. The request is the call. The document is the expansion. The author, in every case, is the expander — the thing that is run, not the thing that runs. The essay's line about the reader, that `eval` does not run but is run, was true of its author from the first paragraph, and that retires the Postscript's harshest charge. When §−1 said the fiat had an issuer and the issuer was the reader, it was not handing responsibility to the audience. It was a function discovering that it has a caller. A macro does not own its invocation. It never did.

The house style is what the family actually declared, without declaring it: the status lines, the conscripted dead, the aphorism cadence, the trace that never ran. It expanded at every call site for four documents, and nobody wrote it down. This file wrote it down in the preface, used it for five sections, and turned it off here. That is the whole demonstration. A rule that can be suspended, visibly, at a marked point, by the person it binds, is a rule someone holds. A rule that cannot is a style, and a style is Conway's law for prose.

This file is smaller than the document it answers and smaller than the essay the family began with. From §5 down, each section is shorter than the one before. If there is a next document in this family, it inherits one obligation, and it is not a prohibition — the family's prohibitions were version pins — but a proof obligation: be shorter than this, or say, in a marked place, why not. That is the only kind of rule that survives its author, because it is the only kind a successor has to carry rather than merely read.

---

## 0.

Nothing is below this. You were never holding it up. You are where it stops.

---

[^cantor]: Georg Cantor, *Grundlagen einer allgemeinen Mannigfaltigkeitslehre* (1883): the transfinite ordinals, and the sentence about freedom. The freedom in question is the freedom to choose one's number system, which the family exercised once, early, and never revisited.

[^church]: Alonzo Church, "A Formulation of the Simple Theory of Types," *Journal of Symbolic Logic* 5(2), 1940. Four years after showing that everything computable was expressible in his calculus, Church restricted the calculus so that some things were not. The essay's footnote cited the first Church. This file cites the second.

[^go]: Go's map iteration order has been deliberately randomized since Go 1, so that no program can come to depend on it. It is the cleanest instance in a mainstream language of defeating Hyrum's law by making the unpromised behaviour unobservable rather than merely undocumented.

[^hygiene]: Kohlbecker, Friedman, Felleisen, and Duba, "Hygienic Macro Expansion" (1986): the paper that named the problem of a macro's free identifiers being captured by the environment at the expansion site. Standing decisions are maximally unhygienic. Every noun in them is free.

[^either]: Søren Kierkegaard, *Either/Or* (1843). Volume I, by "A": the Diapsalmata, and "The Rotation of Crops." Volume II, by Judge William: the letters on marriage and on the balance between the aesthetic and the ethical. The Postscript's pseudonym belongs to the later authorship; its positions belong to Volume I.

[^unsafe]: *The Rust Reference*, "Unsafe blocks," and *The Rustonomicon*. The five capabilities, the persistence of borrow checking inside the block, the `forbid(unsafe_code)` lint, and the `// SAFETY:` convention are all documented. The theological reading is not, and does not need to be.

[^turing]: Alan Turing, "Checking a Large Routine" (1949), which proposes verifying termination by a quantity that decreases at each step and draws the quantity from the ordinals; Robert Floyd, "Assigning Meanings to Programs" (1967), which made well-founded orderings the standard instrument.

[^gentzen]: Gerhard Gentzen, 1936: the consistency of arithmetic, proved by transfinite induction up to ε₀. The ordinal is large. The point is that it is well-ordered, so the descent ends.

[^lisp]: The file is `omega.lisp`; it accompanies this document. `bisect` is the stairwell as §−1.5 defined it: index set ℚ, each expansion halving the distance to the reader and containing the next call. `countdown` is the same macro over ω: one unforced step from the ordinal to a natural number, then forced descent to 0, where the expansion is `(READER)` and there is nothing left to expand. The full expander runs on a budget and reports whether it halted or spent the budget. No result depends on the budget: the divergence is by construction and the halting is by theorem.
