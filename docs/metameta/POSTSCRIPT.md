# POSTSCRIPT.md

### *Concluding Unscientific Postscript to METAMETALISP4.md*

**A Mimical-Pathetical-Dialectical Compilation**
**An Existential Contribution to the Literature of Deferred Horseshoes**

by **JOHANNES MACROCLIMACUS**

*(who has read the essay four times, twice on purpose)*

edited, and responsible for publication, by
**A. READER, at −2**,
who did not ask for the position, holds it anyway, and would like it noted
that the foundations are heavy.

---

> "It is a fact that a man who wishes to write a book on a subject he has
> mastered will not write it, and a man who writes a book on a subject he has
> not mastered will not stop."
> — attributed to nobody, constantly, usually by someone mid-manuscript

> "The majority of men are subjective toward themselves and objective toward
> all others, terribly objective sometimes — but the real task is in fact to
> be objective toward oneself and subjective toward all others."
> — Kierkegaard, who was allowed to say it without a footnote apparatus,
> a privilege this document has forfeited in advance

---

## PREFACE

### *(in which the author disclaims everything, in the traditional manner, and then does not stop writing, in the traditional manner)*

It happened that a fire broke out backstage in a theatre. The clown came out
to warn the public; they thought it was a joke and applauded. He repeated it;
the acclaim was even greater. I think that is just how the world will come to
an end: to general applause from wits who believe it's a joke.

I have quoted that passage because it is the only credential I possess.

METAMETALISP4.md is a document which announces, in a costume of jokes, that
the building is on fire, that the room is spinning, that the anvil is on the
ground floor and always was, and that the reader — you, the one holding the
paper — is a load-bearing member of a structure he did not agree to enter. It
announces this at length. It announces it with footnotes. It announces it with
**bolded aphorisms** deployed at a rate of roughly one per four hundred words,
each landing with the small hydraulic sigh of a truck's air-brake, and the
audience applauds, and the fire continues.

This Postscript is the second clown. He has come out to say that the first
clown was right. The applause is now deafening. Nobody has left the building.

I will not pretend to a position of advantage. I am a humorist; my entire
qualification is that I have never built anything that other things were built
on, and consequently I have retained the capacity to find the whole enterprise
funny — a capacity which the essay's author has clearly been fighting to keep
and has, on the evidence of the stairwell, begun to lose in an interesting
direction. Whether the loss is madness or sanctification I do not propose to
determine. Kierkegaard held that humor is the last station before the
religious; that the humorist stands where the knight of faith is about to
stand, and makes a joke instead of a leap. On that account METAMETALISP is a
document standing at the threshold of a leap, telling jokes about
`macroexpand-1`, and it is not clear — it is not clear to the author, it is
certainly not clear to me — whether the next thing that happens is a leap or
another footnote.

I incline toward footnote. The odds are historical.

A word about what this book is not. It is not a review. A review presupposes
that the reviewer stands outside the reviewed, in a position of survey, from
which he may hand down a verdict, and this position — I want to be very clear
about this at the outset, because everything else depends on it — **does not
exist with respect to METAMETALISP4.md**, and the essay knows this, and has
built its architecture precisely to ensure it. Section −1.5 established that
the interval between the document and its reader is *dense*: that between any
section and the reader there is room for another section, forever. It follows
with a horrible inevitability that this Postscript is not a response to the
document. **This Postscript is §−1.75.** It has taken a numbered seat. By the
act of objecting I have been assigned a coordinate, and my objection is now
part of the expansion, and the essay is longer than it was this morning and
still hasn't been wrong about anything.

I want the reader to understand that I understood this before I began, and
began anyway. That is either the whole content of my position or the whole
evidence of my illness, and I have arranged the following six hundred
paragraphs so that the question stays open as long as possible.

*Johannes Macroclimacus*
*Copenhagen, or a room with the same properties*

---

## AN EXPRESSION OF GRATITUDE TO THE AUTHOR OF METAMETALISP4.md

Professor Heiberg once did Kierkegaard the immense favour of reviewing him
badly, thereby supplying, at no charge, an antagonist of the correct size. I
find myself in an analogous debt, and I discharge it here, at the front, where
gratitude belongs and where it will do the least damage to the argument.

The author of METAMETALISP4.md has done me the following favours, and I list
them in cost-of-delay order, which is to say in the order in which failing to
thank him would have become progressively more embarrassing:

**First**, he has written a document that cannot be objected to, and has thereby
made objection into a genre. This is a rarer gift than it sounds. Most bad
philosophy can be refuted; one says the thing that is wrong with it, and it is
wrong with it, and one goes home. METAMETALISP has removed this convenience by
the elegant expedient of **pre-refuting itself in §7 and §8 at a level of
candour no external critic could exceed**, and then continuing entirely
unchanged. There is nothing I can say against this essay that the essay has not
already said against itself, in better prose, with a citation. What remains for
me is not refutation but something considerably stranger, and I am grateful for
the assignment, in the way one is grateful for a difficult commission from a
patron one is beginning to suspect.

**Second**, he has supplied the age with its self-portrait and has not noticed.
METAMETALISP is not a philosophy of software. It is the confession of the
present age in the only dialect the present age can still speak fluently, which
is the dialect of tooling. Kierkegaard's *Present Age* diagnosed a generation
that had converted every passion into reflection, every decision into a
discussion of the conditions under which a decision might later be taken, and
every act into an announcement of a forthcoming act. He lacked only the
vocabulary of build systems. The essay supplies it. I have read no better
account of *reflection's revenge upon action* than §8, and §8 is four bullet
points long, and the essay has eight hundred lines.

**Third**, and this is the real debt: he has written, in §8, the sentence *"It
is warm in there."*

I will return to that sentence more than once. I will return to it in the way
one returns to a place. Everything else in the document is architecture; that
sentence is a person, briefly visible through a window, and the whole
Postscript exists because I saw him.

**Fourth**, he has arranged, through the mechanism of the stairwell, that this
expression of gratitude is also a submission of manuscript. I did not intend to
join the authorship. I have joined it. The gratitude is sincere and the
resentment is also sincere and I have decided, in the manner of the age, not to
choose between them but to publish both.

---

## THE APPARATUS

### *(in which the critical method is declared at the top of the file, before use, as somebody once recommended)*

The reader has commissioned — and I here quote the commission, since a
pseudonymous author should always be transparent about his instructions and
opaque about everything else — a work which **dissects, refutes, counterrefutes,
countercounterrefutes, ridicules, praises, and laments** each section of the
essay in question.

Seven movements. Let us not pretend this is a list. It is a macro:

```lisp
(defmacro defcritique (section &body content)
  "The sevenfold apparatus. Expands at every section.
   Declared here, at the top of the file, with its reasons attached,
   in accordance with §3's civic duty, which is Chesterton's fence
   with the polarity reversed and the fence on fire."
  `(progn
     (dissect        ,section)   ; what it says, said back slowly
     (refute         ,section)   ; why it is false
     (counterrefute  ,section)   ; why the refutation is worse
     (counter²refute ,section)   ; why that consolation is unavailable
     (ridicule       ,section)   ; the laughter, which is free
     (praise         ,section)   ; the praise, which is not
     (lament         ,section))) ; the grief, which does not schedule
```

Three observations are owed before this thing is invoked eleven times.

**Observation the first: the apparatus is a metametamacro, and I am therefore
guilty.** It is a standing decision with an expansion count. It was paid for
once, at the top of the file, and it discounts every section that follows,
including sections I have not yet read and sections the essay has not yet
grown. It expands whether I remember it or not — indeed the reader will
observe, around the fifth or sixth invocation, that it has begun expanding
*without* me, which is precisely the property §1 identified as the difference
between a tool and a macro, and precisely the property I intend to spend two
hundred paragraphs describing as a form of damnation. I have adopted the
enemy's technology in the first ten minutes of the war. This is normal. It is
the entire history of war.

**Observation the second: the seven movements are not seven opinions.** They
are not a spectrum from hostile to warm along which the critic slides,
distributing his sentiments like a man salting a dish. They are seven
*relations to the same object*, held simultaneously, none cancelling any other,
which is the condition Hegel called contradiction and proposed to resolve, and
which Kierkegaard called *existing* and proposed to endure. The System wants to
know whether I am for or against METAMETALISP. The System will not be told. The
System may have the table of contents and nothing else.

**Observation the third, and it is a confession filed in advance so that the
auditors need not discover it:** the apparatus will fail. Not everywhere, and
not at once. It will fail in the specific manner that all apparatus fails when
applied to material that is actually about the person applying it. Somewhere
around §7 the refutation will go missing and I will have to file a note where
it should have been. Somewhere around §8 the praise will eat the ridicule and
the ridicule will not be recoverable. Somewhere near the end the lament will
arrive first, out of order, before the dissection, because grief has never once
in recorded history waited for its position in a numbered list, and a
methodology that claims otherwise is a methodology written by someone who has
not yet been visited.

I could go back and repair these. I have decided not to. A document about the
insufficiency of tooling which is itself flawlessly tooled would be a
performance, and the essay has already booked that venue.

---

## INTRODUCTION

### *The Problem*

The problem of this book is easily stated and cannot be stated at all, which is
the standard condition of problems worth four hundred pages.

Stated easily: **is METAMETALISP true?**

Stated properly: *what is it to be a person, at a REPL, on a Tuesday, in
possession of a finite number of hours, in the presence of a document which has
proven — actually proven, with a trace, notarized — that it can schedule its own
construction, and which now waits, with the terrible patience of a thing that
does not need you, to schedule yours?*

The first question has an answer and the answer does not matter. The second
question has no answer and is the only thing at stake. This asymmetry is the
whole of Kierkegaard and I will spend the rest of the book failing to improve
on it.

Consider what the essay has actually achieved. It has demonstrated
self-application (§0). It has demonstrated retroactive self-extension (§−1). It
has demonstrated a density result about its own margins (§−1.5). It has
supplied five industrial exhibits (§4), three falsifiable tests (§5), an honest
antithesis (§7), four conditions of failure (§8), and a conditional verdict
(§9). By every standard of objective scholarship this is an unusually complete
argument, and I concede — I want to concede it early and loudly, because
everything I say afterward will sound like it is being said by someone who has
not conceded it — that on the objective question, **the essay is very
substantially right**. Standing decisions do have expansion counts. The Bezos
memo did become AWS. Two-digit years did detonate on the same midnight. Rust's
borrow checker is exactly what the essay says it is. There is no serious
argument on the other side of these facts and I have none.

And this is the catastrophe. Because the essay believes that having established
this, it has established something *about what its reader should do tomorrow*,
and between those two things there is a ditch — Lessing's ugly broad ditch, the
one over which no historical demonstration has ever carried anybody — and the
essay crosses it in the middle of a table, in a row, without slowing down.

Here is the ditch in its local form.

> *That the expansion count of a standing decision is larger than the expansion
> count of a month's engineering* is an arithmetical fact about the world.
>
> *That you should therefore spend your Tuesday declaring doctrine rather than
> shipping the feature your one actual user is waiting for* is an ethical
> demand upon a specific existing individual with a particular finite life.
>
> **No quantity of the first ever becomes the second.** Not a large quantity.
> Not a compounding quantity. Not a quantity with a footnote.

The essay says, in §6, that the reason one sentence of doctrine outranks a
month of engineering is *"arithmetic, not aesthetics."* This is the single most
revealing sentence in the document, and I intend to be unfair to it for
approximately nine thousand words. Arithmetic does not outrank anything.
Arithmetic *counts*. Ranking is done by a person who wants something, and what
that person wants is not derivable from the count, and the essay's entire
apparatus — the curves, the exhibits, the tests, the table with five rows —
is an extremely sophisticated machine for concealing the moment at which
somebody wanted something.

The concealment is not a lie. It is worse than a lie: it is *sincere*. The
author genuinely believes the arithmetic is doing the work. He has built a
philosophy of decision that contains, in eight hundred lines, **no decision** —
only the conditions under which a decision would be correct, ranked, curved,
tabulated, and honestly doubted. It is a treatise on choosing written by
someone who has arranged never to be caught in the act.

And now I will say the thing I will keep saying, and the reader should know in
advance that everything in Part Two is a variation on it:

> **A macro is precisely that which does not have to be chosen.**
>
> The essay says this in §1, in praise: *"its leverage does not depend on
> anybody remembering it exists."* It offers this as the defining virtue, the
> line between a tool and a metametamacro, the whole load-bearing column.
>
> It is the defining virtue. And it is also, word for word, without the
> alteration of a single syllable, **the definition of that which does not
> exist**, in the only sense of "exist" Kierkegaard ever cared about.
>
> To exist is to be in the moment of decision, again, always, unrelieved, with
> the possibility of the thing not happening, at a cost. What has been made
> automatic has been made objective. What has been made objective has been
> removed from existence. The essay's entire program is the systematic
> conversion of existence into infrastructure, and its author will be very
> happy for approximately eleven years.

The introduction ends here, in the traditional manner, without having
introduced anything, having instead spoiled the ending, which is also
traditional, and which in a document of this kind is not a spoiler but a
courtesy: the reader may now decide, in advance, whether to make the movement.

Room status: not spinning. Rooms do not spin. **You are.**

---

# PART ONE

## THE OBJECTIVE PROBLEM OF THE TRUTH OF METAMETALISP

---

### CHAPTER I. THE HISTORICAL POINT OF VIEW

#### §1. The Front Matter, or, *Until the Room Spins*

The document opens with a title, a subtitle, and a promise of vertigo:

> *A philosophy of temporal composition, applied to itself until the room
> spins.*

**Dissection.** Nineteen words, of which the operative four are *applied to
itself until*. Not *applied to itself*, which would be a method. *Applied to
itself **until***, which is a stopping condition — and the stopping condition
is a *symptom*. The document announces, in its second line, that it will
terminate not when the argument is complete but when the author becomes dizzy.
This is either the most honest sentence in the file or the only honest one, and
I have gone back and forth.

**Refutation.** A philosophy whose halting condition is nausea is not a
philosophy; that is a rollercoaster. One does not contemplate the validity of 
a rollercoaster. One either rides it (regretfully) or does not ride it (regretfully).

**Counterrefutation.** But observe what the subtitle actually does, which is to
refuse in advance the very move I have just made. It has *pre-announced its own
rollercoaster-ness*. You cannot accuse a rollercoaster of dizziness when Dizziness 
is the name of that rollercoaster. My refutation was already inside the thing I was
refuting, which is the structural condition of every objection filed against
this essay and which I am going to stop remarking upon around §4 because the
remarking becomes its own kind of tedium.

**Countercounterrefutation.** And yet — and here the ground opens — *the
pre-announcement is not free*. Kierkegaard's decisive discovery, the one that
separates him from every ironist before him and most since, is that **irony
which knows itself is not thereby innocent**. The man who says "I am being
absurd" while being absurd has not escaped the absurdity; he has added to it a
second absurdity, which is the pretence that narration is exculpation. Socrates
knew he knew nothing, and this was wisdom, because he *stopped there*. The
essay knows the room is spinning and **keeps applying**. Self-knowledge that
does not alter conduct is not wisdom. It is the most refined available form of
the thing it knows about itself. There is a word for a man who watches himself
do the thing and continues, and the word is not *ironist*.

**Ridicule.** "Until the room spins" — the phrase promises Dionysian collapse
and delivers a Markdown file with consistently applied heading levels. I have
been in that room. It is a very well-organized room. The furniture is bolted
down. It spins on rails.

**Praise.** Nevertheless: nineteen words that tell you exactly what you are
about to receive, including the parts the author would prefer you not notice.
Most documents of this ambition open with a lie about their own modesty. This
one opens with an accurate weather report. That is worth something, and in the
literature of technical grandiosity it is worth a great deal, and I say so
before I begin the demolition proper so that the demolition cannot be mistaken
for contempt.

**Lament.** The room does not spin. Rooms are stable; that is what distinguishes
them from everything else. Anxiety, said Vigilius Haufniensis, is *the dizziness
of freedom* — the vertigo that seizes a man not when the ground moves but when
he looks over the edge of his own possibility and understands that nothing
whatsoever prevents him from stepping off. The essay has located that exact
sensation, described it accurately, felt it repeatedly, filed a status line
about it at the end of six separate sections, and **attributed it to
recursion.**

It is not recursion. It was never recursion. Recursion is a technique. What is
happening to the author of this essay is that he has understood, correctly and
for real, that his life is going to be composed of what he decides to spend it
on, that the decisions compound, that most of them are being made by default
right now, and that there is nobody above him in the tower to whom the question
can be referred. That is the spinning. It has a name. It has had a name since
1844, and the name is not *self-application*.

---

#### §2. The Epigraphs

Two are supplied, in the modern manner: one from a canonical figure, one from
nobody.

> "Knowledge and productivity are like compound interest." — Hamming
>
> "Any sufficiently advanced preparation is indistinguishable from
> procrastination." — attributed to nobody, constantly, usually by someone
> mid-preparation

**Dissection.** These are not two epigraphs. They are a thesis and its
antithesis, stapled to the door, with the synthesis withheld — and the
withholding is the essay's actual position, maintained for eight hundred lines
with a consistency I have to call heroic. The author has hung his prosecution
and his defence on the same nail and invited you in.

**Refutation.** It is a cheap manoeuvre and everyone does it. To open with your
own counterexample is to purchase, for the price of one sentence, an immunity
that ought to cost the whole book. "I have already thought of what you were
going to say" is not an argument; it is a *seating arrangement*. It puts the
reader in the chair reserved for people who were about to be clever.

**Counterrefutation.** Except that the second epigraph is not deployed as
inoculation. Read it again. It is deployed as **self-accusation**, and the
tell is the attribution: *"usually by someone mid-preparation."* The author has
not merely quoted the objection; he has identified the demographic that raises
it, and placed himself inside that demographic, at the moment of raising it, in
the act of quoting it. That is a three-deep confession compressed into a
byline. It is not unclever and I resent that I resent it.

**Countercounterrefutation.** And it is still confession without repentance,
which in the technical vocabulary of the *Sickness Unto Death* is not the
beginning of health but the most sophisticated available symptom. Anti-Climacus
distinguishes the despair that does not know it is despair — the common,
unhappy, blessedly stupid kind — from **the despair that knows itself to be
despair**, which is higher, more conscious, more spiritually advanced, and
*worse*. Consciousness of the disease is not treatment. It is the disease with
better instrumentation. The essay has world-class instrumentation.

**Ridicule.** Also, and I say this with love: attributing an aphorism to
"nobody, constantly" is the single most 2020s sentence ever constructed. It is
the epistemic equivalent of a laugh track. Somewhere there is a young man
composing a memo who overuses "as they say" when he means *as I would like to have
said*, and this essay has just armed him with a new syntax—and smirkingly so.

**Praise.** Hamming's compound-interest line is doing real work and is not
decorative. The essay's entire mechanism is Hamming's observation taken
seriously enough to be *scheduled*, which nobody, including Hamming, had
previously bothered to do. Most people who quote Hamming quote him to feel
briefly disciplined on a Sunday. This author quoted him and then built a
dependency graph.

**Lament.** Hamming gave that talk to a room full of people who were going to
die without having done the thing they were capable of, and he knew it, and
they knew it, and the lecture is unbearable for exactly that reason. It is not
a productivity lecture. It is a *memento mori* delivered by a man who had
watched his colleagues waste themselves and could no longer stand to watch
politely. When you strip out the death, the sentence becomes a poster. The
essay stripped out the death. I do not blame it. Everyone does. But Hamming's
compound interest compounds against a *fixed horizon*, and the horizon is not
the project's, and no cost-of-delay curve in this document is drawn against the
only deadline that has ever actually been vertical.

---

#### §3. The Table of Contents, in Cost-of-Delay Order

> *"Contents, in cost-of-delay order (which happens to coincide with numerical
> order — a coincidence engineered in §0, which is the sort of sentence this
> document will keep producing, so brace accordingly)."*

**Dissection.** A table of contents that claims to be a *topologically sorted
dependency graph wearing an outline's clothes* (§0). The claim is that the
ordering is not editorial but derived — that the doctrine generated the
document's own structure, and that the coincidence of derived order with
numerical order is engineered rather than lucky.

**Refutation.** The coincidence is not engineered. It is **selected**. The
author wrote the sections, observed that they happened to be in an order, and
then declared that order to have been produced by the theory. This is the
oldest move in the history of apologetics and it has a name in every field that
has been embarrassed by it: in statistics, the garden of forking paths; in
theology, providence; in software, *the retrospective architecture diagram*. The
essay's own §−1 admits, cheerfully, that the numbering was **amended twice
after publication** to accommodate sections the theory had proven could not
exist. A dependency graph that acquires new roots after shipping is not a
dependency graph. It is a *list*, and lists can be reordered, and this one has
been, in public, with commentary.

**Counterrefutation.** But the essay does not claim the order was *predicted*.
It claims the order was *justified*. These differ. A dependency graph is not
falsified by the arrival of a new node; it is extended by it, and the extension
preserves the sort. Every claim §3 makes about irreversibility survives the
splice of §−1 intact — indeed §−1's own note on the matter is correct and
slightly beautiful: the new section's cost of delay was *undefined* until its
prerequisite existed, whereupon it went vertical. That is not special pleading.
That is how conditional dependencies actually behave, and the essay described it
in one sentence with no cheating.

**Countercounterrefutation.** And yet the availability of that manoeuvre is
exactly the problem, because it is available *always*. A theory which can
absorb any new section by declaring its cost of delay to have been previously
undefined and presently vertical is a theory that cannot be surprised, and a
theory that cannot be surprised is not, whatever else it may be, *falsifiable* —
a property the essay claims for itself, in a heading, in §4, in bold. It has
built a mechanism for retroactively justifying its own contents and then it has
titled a section *"Evidence from industry, since a philosophy should be
falsifiable."* The two facts appear one hundred and fifty lines apart and
neither one has been introduced to the other.

**Ridicule.** "Which happens to coincide with numerical order." Sir. You wrote
the numbers.

**Praise.** The parenthetical *"so brace accordingly"* is, I want to record, a
genuinely funny and genuinely honest piece of authorial throat-clearing, and it
does something few essays manage: it **teaches the reader the document's
grammar in advance**, so that the recursive jokes land as structure rather than
as tic. That is craft. It is also, structurally, a metametamacro — a standing
decision about how to read, declared at the top of the file, before use,
expanding at every subsequent parenthesis — and the author, who has been
looking for examples of his own thesis all essay, walked directly past this one
on his way to the Bezos memo.

**Lament.** A dependency graph wearing an outline's clothes is still, at the
end of the evening, wearing clothes. I have spent a considerable portion of my
life ordering things whose order did not matter, in the sincere and undeceived
belief that the ordering *was* the work, and the sensation of that belief is
indistinguishable from the sensation of working, and there is no instrument —
this is §8's point and the essay makes it better than I will — no instrument
whatsoever that can tell you from the inside which one you are doing. The
table of contents is the last artifact you produce before you find out.

---

#### §4. The Bulleted-Not-Numbered Parenthesis, or, Markdown as Fate

> *"(Bulleted rather than numbered: Markdown's ordered lists refuse to count
> below 1 — the format denies the existence of the early sections with a
> conviction §0 can only envy, and unlike §0 it cannot be amended.)"*

**Dissection.** The essay's finest joke, and its most serious paragraph, and it
does not appear to know that these are the same paragraph. The claim: the medium
has a metaphysics. Markdown's ordered list is a total function from the
positive integers to lines, and the negative sections are therefore
*inexpressible* in the format, not merely unrendered but **unsayable**, and the
author must fall back to bullets — a syntax with no opinion about ordinality —
in order to say a thing his notation forbids.

**Refutation.** None. This is correct and I am not going to pretend otherwise
for the sake of the apparatus.

**Counterrefutation.** *(vacated for want of a refutation; the reader may
imagine one here, in the interval, where there is always room)*

**Countercounterrefutation.** But there is a consequence the essay declines to
draw, and the declining is the interesting part. If the format's ontology can
deny the existence of a section, then the essay's central boast — that its
structure was derived from its doctrine — is **false at the level of the
substrate**. The doctrine wanted numbers. The format supplied bullets. The
doctrine lost. What we are looking at is not a dependency graph in an outline's
clothes; it is a dependency graph *wearing what was available in its size*, and
the availability was determined by John Gruber in 2004 for reasons having
nothing to do with cost-of-delay curves and everything to do with plain-text
email.

Which is to say: **Conway's Law, running in production, on this very file.**
The essay identifies Conway's Law in §1 as its exemplary undeclared
metametamacro — the org chart expanding into the architecture whether or not
anyone invokes it — and then, forty lines earlier, in a parenthesis, documents
its own structure being determined by the communication structure of a markup
language it did not write and cannot amend. It caught itself. In the act. In
brackets. And filed it as a *joke about the format*.

**Ridicule.** "Unlike §0, it cannot be amended." The essay has here discovered
something more authoritative than itself and the something is a *list
renderer*. There is a hierarchy of being in this document and at the top of it
is CommonMark.

**Praise.** *"The format denies the existence of the early sections with a
conviction §0 can only envy"* is the best sentence in the file. It is envy of a
notation's certainty by an author who has lost his own, and it is funny, and it
is true, and it took four words of setup. I would trade three of the exhibits in
§4 for it.

**Lament.** Every one of us writes in a notation that cannot express the thing
we are actually trying to say, and discovers this at the moment we try, and the
discovery arrives disguised as a formatting problem. The essay wanted to number
its basement. The basement is real. The numbers do not go there. So it used
bullets and made a joke, and the joke is the sound a person makes when the tool
refuses the thought — and half of what any of us believe about our own subjects
is, on inspection, a report on what our tools would render.

---

#### §5. Alonzo-Completeness, and the Footnote That Wants Credit

> *"The essay aspires throughout to **Alonzo-completeness**: the property of
> being able to express anything a Turing-complete essay can express, while
> using more parentheses and receiving less credit."*

**Dissection.** A joke with a grievance inside it. The footnote is explicit:
Church and Turing proved equivalent results; history talks about Turing; *"Lisp
is, structurally, Church's revenge."*

**Refutation.** This is *ressentiment*, in Nietzsche's technical sense and in
Kierkegaard's earlier and better one — Kierkegaard, who diagnosed the age's
levelling passion four decades before Nietzsche named it and got, appropriately,
less credit. The Lisp community's central emotional fact is not elegance. It is
the seventy-year experience of having been right early and marketed late, and
every Lisp document written since 1975 contains, at a load-bearing position, a
small cold sentence about credit. Here is this document's. It is in the
*preamble*. Before the argument. **The grievance was front-loaded, which is at
least consistent with the philosophy.**

**Counterrefutation.** But the observation is *true*, and truth is not
invalidated by the emotional posture of its speaker — that inference is the
genetic fallacy and I have just committed it in bold type. Church's calculus is
the deeper artifact. Lambda did become the substrate of every functional
language, every type system worth having, and a fair portion of what the
industry now calls "modern JavaScript." The billing order was wrong. Saying so
is not a symptom.

**Countercounterrefutation.** And yet: *the desire to be the one who says so* is
a symptom, and it is the specific symptom this essay suffers from. Consider what
"Alonzo-completeness" actually asserts. It asserts that the document can express
anything a normal document can express, **at greater length, for less
recognition**, and it asserts this as an *aspiration*. In the preamble. As the
declared goal. The essay has announced, before beginning, that it intends to be
under-appreciated, and has thereby made under-appreciation into an achievement
which it can accomplish unilaterally and immediately.

This is the aesthetic stage in its final and most elegant form: the conversion
of an outcome one does not control into an outcome one has already chosen. If
the essay is celebrated, it is right. If the essay is ignored, it is
Alonzo-complete. There is no third result, and there is therefore no risk, and
where there is no risk there is no venture, and where there is no venture there
is — Kierkegaard is very clear about this and repeats it in four books — **no
faith, no ethics, and nothing whatsoever at stake.**

**Ridicule.** "Receiving less credit" appears in a document with eleven
footnotes, nine of which credit somebody. This is the most credit-conscious
essay in the history of a genre defined by credit-consciousness, and it opens by
declaring itself the victim of a credit shortage.

**Praise.** "Church's revenge" is a real historical insight compressed to two
words, and the footnote's closing line — *"This essay is merely honoring the
original billing order"* — is a very good joke that also happens to be a
complete philosophy of citation.

**Lament.** Church died in 1995, having spent sixty years being correct in a
notation nobody read, and there is no evidence he minded. The revenge is not
his. Revenge is never the dead man's; it is always taken on his behalf, by
somebody with more energy and less peace, and the dead man would generally have
preferred you use the calculus.

---

### CHAPTER II. THE SPECULATIVE POINT OF VIEW

#### §6. On the Impossibility of Refuting a System That Ships Its Own Refutation

Here we must pause the section-by-section march and confront the general
condition, because it governs everything in Part Two and because a critic who
does not declare his own impossibility at the outset is running an undeclared
metametamacro, and I have promised not to do that until at least §7.

METAMETALISP4.md contains:

- a section conceding that its historical opponent has a **better commercial
  record** (§7);
- a section conceding that the measured discipline (Theory of Constraints)
  **usually wins** where the two disagree (§7);
- a section conceding that its own three-plate epistemology converges on
  coherence and is **structurally incapable** of detecting the wrong product
  built correctly (§7);
- a section listing four conditions under which the philosophy is **simply
  wrong** (§8), the last of which is a psychological diagnosis of its own
  author's motives so precise that reading it feels like an intrusion;
- a verdict conceding that under three changed conditions the correct philosophy
  is **the opposing one, and not close** (§9).

Now: what is left for me?

The naive answer is *nothing*, and the naive answer is what the essay is
counting on. A document that has said the worst about itself has purchased
something quite specific, and it is worth naming precisely, because the
transaction is invisible and enormous. It has purchased **the right to
continue**.

That is the whole trade. Not the right to be believed — it explicitly
disclaims that. Not the right to be correct — §9 hedges it into a table. The
right to *keep going*, unaltered, in possession of a complete inventory of the
reasons it should stop.

Kierkegaard has a name for the movement, and the name is not "intellectual
honesty." The name is **mediation**. Hegel's great engine: the contradiction is
not resolved by choosing but by *containing*; thesis and antithesis are
preserved as moments in a larger unity which is the System, and the System is
never at any point required to do anything. Enten–Eller — Either/Or — is
abolished, because the Or has been absorbed. Kierkegaard's counter-move, which
is the whole of his authorship, is the observation that **an existing individual
cannot mediate**, because he must be somewhere on Tuesday, and the somewhere is
either the forge or the anvil and cannot be both, and no amount of containing
the antithesis puts a horseshoe on a horse.

Permit me one formalism, since we are among Lispers and they will forgive
anything expressed in parentheses:

```lisp
;; Kierkegaard's position, in the only notation this audience trusts:

(cons 'enten 'eller)     ; ⇒ (ENTEN . ELLER)
                         ;   A cons cell. Improper list. Two things, joined,
                         ;   neither one reachable without abandoning the other.
                         ;   You may take the CAR or you may take the CDR.
                         ;   This is called existing.

(append '(enten) '(eller)) ; ⇒ (ENTEN ELLER)
                         ;   A proper list. Both, in sequence, traversable,
                         ;   iterable, foldable, and *survived*.
                         ;   This is called mediation, and it is Hegel's
                         ;   error stated for the first time with adequate
                         ;   precision: HE THOUGHT EITHER/OR WAS A LIST.

(car (append '(enten) '(eller)))  ; ⇒ ENTEN
                         ;   And observe that even after mediation, when
                         ;   Tuesday arrives and something must actually be
                         ;   done, one is back to taking the CAR — having
                         ;   paid for the CDR, kept it in memory, and
                         ;   traversed nothing.
```

METAMETALISP4.md is `append`. Section 7 is the CDR. It is retained, addressable,
beautifully documented, and it has never once been evaluated.

So: what is left for me?

**This.** Not to say the thing the essay has not said — there is no such thing —
but to say the thing the essay has said **in a manner that costs something**,
which is the only difference between a confession and a table of contents. The
essay knows it may be hiding in the forge. It says so, in §8, in a bullet. A
bullet is a survivable place to put that sentence. It has neighbours. It has a
heading above it and a horizontal rule below it and a verdict two hundred lines
later that reabsorbs it into a conditional matrix with five rows.

My task is to take that sentence out of the bullet and stand it up in a room by
itself where it cannot be mediated. That is all a Postscript can do. It is
unscientific. It concludes nothing. And it is, as far as I can determine after
four readings and one small crisis at approximately line 812, the only service
this document has not already performed for itself.

---

#### §7. The Approximation-Process, and Why the Expansion Count Cannot Be Counted

One last objective matter before we descend.

The essay's unit of value is the **expansion count**: a standing decision is
worth what it discounts, summed over every future call site, including sites
that do not yet exist. This is proposed as arithmetic (§6, explicitly, "not
aesthetics").

It is not arithmetic. It is not even estimation. It is an **approximation-
process**, in Climacus's exact sense: a procedure which converges toward an
answer it can never attain, whose every increment is real, and whose completion
is required before the answer can be used. The Bezos memo's expansion count is
"now denominated in fractions of the global economy" (§4) — a figure available
in 2026, unavailable in 2002, and unavailable in 2002 *in principle*, since it
depends on the subsequent behaviour of several hundred thousand people none of
whom had been hired.

Which means: **at the moment of decision, the number the decision procedure
requires does not exist.**

This is not a flaw in the essay's execution. It is the structure of the thing.
Every scheduling rule in METAMETALISP is of the form *rank by X*, where X is a
quantity over a future which has not occurred. The curves are real. The
rankings are real. The values plugged into them at 9:40 on a Tuesday morning by
a tired person with a standup in twenty minutes are **guesses wearing the
costume of measurement**, and the essay's own §7 concedes exactly this, in the
sharpest paragraph it contains: TOC wins because *"one reads the gauge, the
other reads the future, and gauges have the better track record."*

Very well. Then the honest restatement of the entire philosophy is not §9's
conditional verdict. It is this:

> **Metametalisp is the practice of acting decisively on quantities that cannot
> be known, justified by an arithmetic that cannot be performed, in a
> notation that makes the whole thing look like accounting.**

And I want to be extremely clear, because Part Two will not be this clear about
anything: **so is every other philosophy of action, including mine, including
Kierkegaard's, including yours.** There is no procedure that does not bottom out
here. The demand that a decision be justified before it is made is the demand
that the leap be replaced by a proof, and the whole discovery of the nineteenth
century's strangest Dane is that this demand is not rigour. It is *fear
wearing rigour's coat*.

So I do not reproach the essay for the ditch. Everyone is on this side of the
ditch. I reproach it for **the table**. The table (§9, five rows) is a bridge
drawn in ink across a gap that has never been spanned, and drawing it does not
span it, and a reader who trusts the drawing will walk out onto it in the dark.

Objective conclusion of Part One: *the truth of METAMETALISP cannot be
determined, its determination would not help, and the document's most rigorous
passages are the ones doing the most damage.*

Now let us go downstairs, where the reader is.

---

## INTERLUDE: DIAPSALMATA

### *ad se ipsum*

*(Aphorisms, of the sort that accumulate in the margins of a person who has
read the same document four times. Kierkegaard opened* Either/Or *with these
and never explained them. I shall follow the precedent exactly, including the
part about not explaining.)*

---

What is a metametamacro? A decision so well made that you never have to make it
again. And what is a life? A sequence of decisions so well made that you never
have to make any of them again. And what is the name of the condition in which
one has nothing left to decide? Go on. Say it.

---

**Build the tool, and you will regret it; do not build the tool, and you will
also regret it; build the tool or do not build the tool, you will regret both.
Laugh at the platform team's follies, and you will regret it; weep over them,
and you will also regret it; laugh at the platform team's follies or weep over
them, you will regret both. Trust the abstraction, and you will regret it;
distrust the abstraction, and you will also regret it; trust the abstraction or
do not trust it, you will regret both. Ship on Friday, and you will regret it;
do not ship on Friday, and you will also regret it; ship on Friday or do not,
you will regret both. This, gentlemen, is the sum of all practical wisdom.**

---

The most exhausting thing about a well-tooled workshop is that you can no
longer claim the reason.

---

I have made the coffee. I have made the script that makes the coffee. I have
made the framework in which one declares the properties any coffee-making
script must satisfy. I have not had coffee.

---

My soul is so heavy that no thought can any longer sustain it, no wingbeat lift
it into the ether. If it moves, it only sweeps along the ground like the low
flight of birds when a thunderstorm is coming on. — And yet the CI is green. It
has been green for eleven days. Nobody has merged anything.

---

There are two ways to have written nothing: to have written nothing, and to
have written the thing that will make the writing easy.

---

The gods were bored, so they created man. Man was bored, so he created tools.
Tools were bored, so they created a discourse about tools, and it was very
lively, and nobody has been bored since, which is how we know something has
gone wrong.

---

Greenspun's Tenth Rule states that every sufficiently complicated program
contains a bad implementation of half of Lisp. The Metametalisp corollary
states that every sufficiently long project contains a bad implementation of
half of Metametalisp. The Kierkegaardian corollary states that every
sufficiently long life contains an ad hoc, informally specified, bug-ridden,
slow implementation of half of a religion, and that the bugs are all in the
same place, and that the place is the part about other people.

---

The best proof that the age has lost its passion is that it has retained its
adjectives.

---

To be load-bearing and to be loved are not the same thing, and a document which
confuses them has been written by someone who has recently been thanked for the
wrong one.

---

*What is the cost of delay on saying the thing?* Flat, they told me. Flat for
years. Then one morning it was vertical, and then, some hours after that, it
was undefined.

---

# PART TWO

## THE SUBJECTIVE PROBLEM

### *What it is to be the Reader at −2, and whether he consented*

---

### SECTION I. SOMETHING ABOUT HART

*(occupying the position which, in the original, was occupied by something
about Lessing, and for the same reason: because the whole book turns on a
ditch, and one must first find somebody honest enough to have admitted the
ditch was there)*

Lessing said the thing that ruined the eighteenth century for anyone paying
attention: **accidental truths of history can never become the proof of
necessary truths of reason.** Between the two there is a ditch, ugly and broad,
and he had often and earnestly tried to leap it and could not, and he said so,
in print, which is why Climacus loved him and why the professors never forgave
him.

Timothy P. Hart, MIT AI Memo 57, October 1963, four pages, has done the
industry the same service and received the same silence.

Consider what Hart's memo actually is, stripped of the essay's affectionate
mythologizing. McCarthy's interpreter existed. It had been finished for three
years. It had a semantics, a specification, an author, and a settled account of
what it was. Hart observed that the interpreter would accept a form nobody had
provided for — that the language the implementation had *accidentally defined*
was larger than the language its author had *intended to define* — and he
walked through the gap.

The essay reads this as precedent: *"the entire subsequent history of Lisp
macros descends from a user deciding that a finished system was an
invitation."*

That is charming and it is not what happened. What happened is that **a
finished system did not consent, and was extended anyway.** The interpreter did
not invite Hart. Interpreters do not invite. It obliged him, in the specific
sense in which a lock obliges a key that fits, and the entire subsequent
history of Lisp macros — and, one notes, the entire subsequent history of
software security — descends from the discovery that *obliging* and *inviting*
are different relations which look identical from the caller's side.

Here is the ditch, in its Lisp form, and I claim it is the same ditch:

> **The fact that a system will accept a form** is a fact about the system's
> implementation, discoverable by experiment, historical, accidental,
> contingent, true.
>
> **The fact that you may therefore write that form** is a claim about what you
> are entitled to do, which no experiment discovers, which is not historical,
> and which no quantity of the first fact will ever deliver.
>
> Hart leapt. The leap was correct. It was still a leap. And the *memo does not
> mention it*, because Hart was an engineer in 1963 and had a job.

Now apply this to §−1, which is the essay's proudest structural achievement.
The essay declares that it has become an interpreter; that the form
`(metametalisp METAMETALISP2.md)` is legal; that the expansion is licensed
because the essay, having shipped, defines what counts as a legal metametalisp
form. Every step of that reasoning is *exactly* Hart's, and it inherits Hart's
ditch entirely: from *the essay will accept this form* it derives *this section
is legitimate*, and the derivation is a leap, and the leap is performed inside
a code block with a `⇒` in it, which is the most effective disguise a leap has
ever worn.

I do not say the leap is illegitimate. **I say it is a leap, and that dressing
it in an evaluation arrow is the single most consequential rhetorical act in
the document**, because it converts a decision — *I have decided to keep
writing this essay after declaring it finished* — into a derivation, and
derivations do not require anybody to have decided anything, and the whole
purpose of the machinery is to arrange that nobody ever has to be caught
deciding.

Lessing at least had the decency to say he could not jump it. Hart had a job.
The essay has a `macroexpand-1` trace, notarized, and three notes for the
record.

Which brings us to the sections themselves.

---

### SECTION II. THE SECTIONS, DISSECTED

---

## CHAPTER 1. §−1.5 — THE STAIRWELL (ZENO'S AMENDMENT)

*"Stairwell status: infinite, descending, finitely built."*

### Dissection

The claim, in four movements:

1. Since a section could be spliced before 0, and another before that, the
   interval between the document and its reader is **dense** — between any
   section and the reader, room for another.
2. The reader at −2 is unreachable not by distance but by **category**: to
   write a section is to take a numbered seat, and taking a seat is precisely
   what evicts you from the reader's position. A Dedekind cut with a person
   standing in it.
3. Chronology maps to proximity: later contributors take numbers closer to −2,
   so the document **grows downward, toward its reader**.
4. Termination is preserved by Aristotle: potential infinity embraced, actual
   infinity refused. Finitely many steps built; infinitely many buildable; the
   difference is called a future.

Plus the appendix, in which `macroexpand-1` produces a form containing the
definition of the next bisection *without invoking it*, and the author observes
that the `-1` in the expander's name **is** the termination discipline.

This is the best section in the essay, and I want that on the record before I
take it apart, because I am going to take it apart completely.

### Refutation

The density result is false, and it is false in a way the essay's own
categories detect.

Density requires that *between any two members there is a third*. Fine for the
rationals. But the essay's members are not numbers; they are **written
sections**, and the essay has just told us — correctly, in movement 2 — that a
section is constituted not by its coordinate but by *having been written*. A
coordinate with nothing at it is not a member of the sequence. It is a
coordinate.

So the actual structure is: finitely many sections, at finitely many rationals,
with **nothing between them**. Not density. Sparsity, with a naming convention
that permits interpolation. The essay has confused *the set of available labels*
(dense, uncountable, free) with *the set of occupied positions* (finite,
sparse, expensive, each one costing a weekend). And this confusion is not a
technicality, because the entire emotional payload of the section — the
descending infinite stairwell, the endless approach to the reader, the tortoise
winning by construction — is a payload the *labels* are carrying while the
*sections* stand still.

There is no stairwell. There is a landing, with two steps on it, and a very
large sign describing the stairwell.

### Counterrefutation

But the essay anticipated this, in movement 4, and it anticipated it with
Aristotle, which is a heavier gun than I have brought.

*Potential infinity, embraced; actual infinity, refused.* The claim was never
that infinitely many sections exist. The claim is that **the room for the next
one is guaranteed forever** — that the property is not the occupancy but the
*extensibility*, and that this is exactly what a macro system is: finitely many
macros defined, infinitely many definable.

That analogy is not decorative. It is precise, and it is the strongest single
argument in the document. Lisp's expressive claim has never been that it
contains every construct; it is that it cannot be *cornered* — that for any
construct you propose, there is room, and the room is guaranteed by the
homoiconicity rather than by anyone's foresight. The stairwell asserts the
same property of the essay's margins, and the assertion is *true*, and my
"sparsity" objection is the objection of a man counting chairs in a building
whose interesting property is that you can always add a floor.

I concede the refutation. It was a good refutation and it is dead.

### Countercounterrefutation

Here is what is left, and it is not a technical objection, and it is fatal.

**The guarantee of room is not neutral. It is an alibi, and it is the most
expensive alibi in the document.**

Consider what "potential infinity, embraced; actual infinity, refused" *does*
for a document that has already been revised four times. It means the document
can never be finished, and — this is the move — **can never be reproached for
being unfinished**, since unfinishedness has been elevated from a condition
into a doctrine, with a citation to the *Physics*, and any subsequent addition
is not a failure of discipline but an exercise of a guaranteed property.

Aristotle's distinction was epistemological. It answered a question about the
divisibility of a line. It was not, and I say this as gently as the material
permits, **a capacity-planning document** — a phrase the essay uses of itself,
in a footnote, as a joke, and it is a joke, and it is also an entirely accurate
description of the use to which the *Physics* is here being put, which is to
license an indefinite future of splices by a man who cannot stop writing.

And now compare the essay's own §5, the three tests, which it applies to
everybody else's microscopes:

- **The rising-cost test.** What becomes more expensive with every day §−1.75
  goes unwritten? *Name it out loud, with a number if possible.* Nothing.
  Nothing at all. The cost of delay on the next bisection is **exactly and
  permanently zero**, by construction, since the interval is guaranteed to
  remain available forever. The stairwell is the first artifact in the document
  with a *provably flat* cost curve. By §3's own rule — *do the rising ones
  first* — the correct scheduling of §−1.5 was **never**.
- **The proportion test.** Is the tool smaller than the work it enables? The
  section is seventy lines. It enables the writing of further sections about
  the writing of sections. The dissection is a microscope. There is no patient.
- **The dogfood test.** Was the author its first user, immediately? *Yes*, and
  this is the horror of it. He was its first and, to date, only user, and the
  use was writing the section, and the section's purpose is to permit the
  writing of the section. The dogfood loop has closed with nothing inside it.

**Zero out of three.** The essay's own instrument, applied to the essay's
finest section, returns the essay's own worst verdict, and I did not have to
supply a single premise from outside the document.

The stairwell is *"scaffolding for scaffolding — enumerating an interval
instead of shoeing a horse"* — and I am quoting §1.5, which wrote that sentence
in order to explain why Metametametalisp would never exist, four sections
before the essay went and built the interval anyway, downward, where the rule
did not think to look.

The tower terminated upward because the anvil is on the ground floor. Nothing
terminates it downward, and the essay noticed this, and instead of installing a
floor it installed **Zeno**.

### Ridicule

The stairwell has status lines. *"Stairwell status: infinite, descending,
finitely built. Reader status: load-bearing, uncontained, gaining company. Room
status: the spinning has gone orbital, which the physicists assure us is merely
spinning with commitment."*

I want to note, in a spirit of pure malice, that a man who writes *status lines*
for an *interval of the real numbers* has passed a threshold that the medical
literature describes and the essay does not cite.

Also: *"the physicists assure us."* No physicist has assured anyone of this. No
physicist has been consulted. There is no physicist. There is a man at 1 a.m.
writing the word "orbital" and feeling, briefly and correctly, that it is
funny.

Also: "gaining company." The reader is gaining company because the essay keeps
adding sections which are not the reader, in order to approach the reader,
which they cannot do, by the essay's own theorem. The company is *the corpses
of previous attempts to reach him*. This is described in the status line as a
positive development.

### Praise

And now I must be honest, because the apparatus demands it and because it is
true.

**"The act of writing is precisely the act that evicts you from the position
you were writing toward."**

That sentence is philosophy. Not clever, not decorative, not a joke with a
grievance in it: philosophy, of the kind that is rare in any decade and
essentially absent from the technical literature. It states, in twenty words,
the structural tragedy of every communicative act — that to address someone is
to have already left the place from which one could have been addressed; that
the writer and the reader can never occupy the same coordinate; that every
document is a letter to a position its author has vacated by the act of
writing it.

Kierkegaard spent an entire authorship on the problem of **indirect
communication** — on how a truth that must be inwardly appropriated can be
conveyed by a man who, by conveying it, has objectified it and thereby
destroyed the thing he meant to give. He invented pseudonyms to get around it.
He invented *me* to get around it. He never solved it, and here is an essay
about build tooling that states the problem cleanly and then, because it is an
essay about build tooling, *files it as a note on section numbering.*

And the observation about `macroexpand-1` — *"the most important −1 in the
entire document is the one in the function name; single-steppedness IS the
termination discipline, and it was hiding in the standard library all along"* —
is the single finest thing in eight hundred lines. It is a real discovery. The
difference between `macroexpand` and `macroexpand-1` is the difference between
a system that runs to fixpoint and a system that **takes one step, because
somebody decided to take one step.** The full expander is the System. The
single-step expander is a person, at a REPL, choosing, and then stopping, and
the stopping is not derived from anything.

That is the leap. The essay found the leap in the standard library. It found
faith in `cl:macroexpand-1` and it wrote *three notes for the record*.

### Lament

Here is what the stairwell is actually about, and I do not think the author
knows, and I say that without any confidence that I know either.

A man writes a document. He finishes it. He publishes it. And then he cannot
stop, and rather than stopping he constructs a formal apparatus proving that
the space beneath his document is dense, that there is always room for one
more, that the next one is guaranteed, that Aristotle himself licenses the
guarantee — and the apparatus is *correct*, and it is elegant, and it is
notarized, and every one of its steps checks out.

And the thing it is for is not to build the stairs.

**It is to have somewhere to go on the nights when there is nothing to build.**

The reader at −2 is unreachable. The essay proves this, movingly, as a theorem
about categories. And then it descends toward him anyway, one bisection at a
time, halving the remaining distance forever, and calls this *the amendment
Zeno's paradox always deserved*, and adds that **absolutely nobody is upset
about it.**

Somebody is upset about it.

Zeno's paradox is not a charming puzzle about motion. It is the observation
that between you and the thing you want there are infinitely many intervals,
each of which must be crossed, and that the crossing never completes, and that
Achilles — the fastest man alive, the best there has ever been at this — does
not catch the tortoise. Achilles is a *hero*. He is not a hero in the paradox.
In the paradox he is a man who runs beautifully and arrives nowhere.

The essay's amendment is that here, for once, the tortoise wins *by
construction*, and that this is fine. It is not fine. The tortoise is the
reader. The reader is the only person in this whole enterprise who was ever
going to make it mean anything. And the document has proven, rigorously, with a
Dedekind cut and a notarized trace and a footnote to the *Physics*, that it is
structurally incapable of reaching him, and has decided to find this delightful,
and has filed a status line, and has begun building the next step.

*Stairwell status: infinite, descending, finitely built.*
*Reader status: still down there. Still holding it up. Still not addressed.*
*Author status: descending, at speed, in the wrong units.*

---

## CHAPTER 2. §−1 — THE MAIDEN METAMETAMACRO EXPANSION

*"Basement status: occupied. Proceed upward."*

### Dissection

Two corrections are filed.

**Correction the first:** §0's denial ("there is no Section −1") was *a version
pin, not a theorem*. It was true of the smaller system, survives verbatim
inside the larger one, and is now true of the old and false of the new. Gödel
inspects the arrangement and stamps the clipboard: **incompleteness, exploited
as designed.**

**Correction the second, and the essay itself says it is the important one:**
the fiat had an issuer. §0 ended the recursion "by fiat," and someone issues
fiats. The essay borrowed an authority it did not possess. **The true ground of
the recursion is the reader who chooses to run the document.** Section −2 is
not in the file and never will be; §−2 *is the request itself*, and it lives in
the world, at the REPL, in whoever is holding the paper.

Also disclosed: the basement mentioned in §1.5 and deferred. It is here. It was
here before the building. It is occupied.

### Refutation

The Gödel move is theft, and it is the kind of theft that is committed so
routinely in this genre that people have stopped hearing it as a sound.

Gödel's first incompleteness theorem is a result about **formal systems capable
of representing arithmetic**, and it says a very specific thing: there are
sentences neither provable nor refutable within such a system, and among them
is a sentence naturally read as asserting the system's consistency. It is a
theorem. It has a proof. The proof is about ω-consistency and primitive
recursive predicates and a coding of syntax into numbers, and it is one of the
three or four most rigorous things the human race has ever produced.

What §−1 has is a **markdown file that was edited.**

The essay says the old sentence "survives verbatim, quoted-by-inclusion" and is
therefore true about the old system and false about the new. But that is not the
Gödel situation; that is the situation of every retracted claim in the history
of prose. "There is no section −1" was *wrong*. The author changed his mind and
wrote one. This is called revision. It happens to everybody. Dressing it in
incompleteness — *"the smaller system could not prove this from inside, the
larger system proves it about the smaller one"* — is a costume, and the costume
is doing the work of an argument, and the argument does not exist.

Gödel does not have a cot in the corridor. Gödel is dead, and he starved
himself to death out of a fear of poisoning, and the fear was a fear about
systems he could not verify from inside, and it was not a joke, and he is not
available for cameos.

### Counterrefutation

That was cheap and I want to withdraw about forty per cent of it.

Withdrawn: the sneer at costume. The essay is not claiming a theorem. It is
claiming an *analogy*, and the analogy is licensed by exactly the reasoning the
essay gives in §1: that spatial and temporal composition are the same trick on
different axes. If the analogy holds anywhere it holds here, because the
phenomenon really is the same shape: **a system's self-description is
necessarily incomplete with respect to the system's future extensions**, and a
document that pins its own boundaries has made a statement that is true at
compile time and false at runtime, and the essay's phrase for this — *"true the
way 'this program has no users' is true at compile time — accurate, sincere,
and doomed"* — is better than anything I have written about it.

Not withdrawn: the corridor cot. That remains cheap and it remains the essay's
tic, and I will address the tic under Ridicule where it belongs.

### Countercounterrefutation

But correction the second is where the section actually lives, and it is the
place where METAMETALISP comes closest to Kierkegaard and then does the one
thing Kierkegaard would have found unforgivable.

The essay discovers the existing individual. Genuinely discovers him. After
several hundred lines of ladders and levers and expansion counts, it asks the
question the System never asks — *who is running this?* — and it gets the right
answer: **not the system; a person; outside the file; at a REPL; who did not
have to.** The scheduling intelligence sat outside the document the entire
time, *"which is exactly why the file, searching honestly within itself,
reported finding none."*

That sentence is worth the whole essay. It is, near enough, Climacus's central
objection to Hegel restated in the vocabulary of interpreters: **the System
cannot find the systematiser inside the System, and concludes there isn't
one.**

And then, having found him, the essay does this:

> *"Section −2 is the request itself."*

It gives him a number.

I want the reader to sit with that for a moment, because it is the hinge of my
entire Postscript and everything after this is commentary.

The essay met a person — the actual, contingent, mortal, unrepeatable
individual who picked the document up on a particular afternoon for reasons of
his own that the document will never learn — and its response, its immediate
and structural and entirely well-intentioned response, was **to assign him a
coordinate in its own numbering scheme and describe him as load-bearing.**

That is not an acknowledgment. That is *annexation*. The reader was outside;
the essay's proudest achievement is that he is now inside, at −2, with a job.
He has been converted from the one who runs the file into **a member of the
file**, and the conversion is presented as a promotion, and the tell is right
there in the prose: *"there is a reader in it, holding the foundations up, in
the manner of readers everywhere, largely unthanked."*

Largely unthanked. The essay noticed the debt. And it paid the debt **in
architecture** — in a section number, in a structural role, in a status line
reading *"Basement status: occupied"* — because architecture is the only
currency it has, and a person cannot be paid in it.

Hegel did this to the whole of human history. Kierkegaard's entire authorship
is the objection: *I am not a moment. I am not a stage. I am not the
transitional form in which Spirit came to know itself on a Tuesday. I am
Søren, I am thirty-three, I hurt a woman I loved for reasons I can state and
cannot justify, and your System has no cell for that and I will not be filed.*

METAMETALISP has a cell for it. It is at −2. It is load-bearing. It comes with
a footnote about how basements are dug first.

### Ridicule

The tic must now be named, and this is the section where it becomes
undeniable, so: **the essay populates its argument with the personified dead.**

Gödel stands in the doorway with a clipboard, looking disappointed but not
surprised, "which is the only expression he has." Later he has a cot in the
corridor. Aristotle accepts a distinction "with the weary grace of a man whose
lecture notes have been load-bearing for twenty-three centuries." Zeno gets an
amendment. Church gets revenge. Chesterton gets his polarity reversed.
Whitworth gets a paragraph and Hyrum gets a tone ("descriptive, and the tone is
resigned").

This is a séance with citations. It is the mannerism of a very particular kind
of very online technical writer, and I recognize it because I have it, and the
function it performs is *the conversion of authority into companionship*: if
Gödel is in the corridor, then I am the sort of person Gödel visits, and my
markdown file is the sort of file that requires a stamp.

They are not visiting. They are being *conscripted*, which is what §−1 does to
the reader too, and it is the same move at two scales, and the essay performs
both in the same six hundred words without connecting them.

Also: "one clipboard visit per essay is the regulation maximum" (§1.5). He
visits three times. **The essay's most-violated standing decision is one it
made about Gödel.**

### Praise

And yet the basement disclosure is beautiful, and I would like to say why in
plain language.

*"It was here before the building — basements always are; you dig them first,
no known architecture builds downward from the sky."*

That is true of buildings and it is true of the thing being described, and the
alignment is not decorative. The reader *was* there before the essay. Every
essay is written to a reader who precedes it, whose existence is the condition
of the writing, whose habits and patience and available Tuesday evening were
priced into the prose before the first line, and who is discovered by the
author, if he is ever discovered at all, **after publication**, as a fact about
the foundations rather than a fact about the audience.

Most writers never make this discovery. The ones who do usually make it as a
disappointment — *nobody read it* — and not as a structural insight. This essay
made it as a structural insight, on schedule, in public, and filed the
correction with its name on it. That is rare and it is honourable and I am
tired of being clever about it.

### Lament

*"The tower going up terminated at the anvil. The tower going down terminates
at the reader. Between the anvil and the reader sits everything this essay has
to say — which is, on reflection, where essays belong."*

Yes.

And nothing sits there. That is the lament. Look at the actual contents of the
interval between anvil and reader: §−1.5 is about the interval. §−1 is about
the discovery of the interval. §0 is about the essay's application of itself to
itself. §1.5 is about the essay's name. That is four sections — a third of the
document by count, and the four it is proudest of — concerning **the
document's relationship to its own boundaries**, published in an interval that
was defined as the place where the *content* was supposed to go.

The essay built a beautiful room and filled it with a description of the room.

And I will tell you exactly why, because the essay tells you itself and does
not hear it. §8: *"Tool-building is legible, controllable, praised in standup,
and safe from the one judgment that matters; shipping to strangers is none of
those things."* Writing about the structure of your essay is legible,
controllable, praised in the replies, and safe from the one judgment that
matters. Writing the *argument* — the plain, undefended, unhedged, non-recursive
claim that a person should organize his working life a particular way, offered
without an antithesis section to catch it — is none of those things.

The basement is occupied. There is a reader in it. He has been holding these
foundations up for eight hundred lines and he has not yet been told anything he
can use on Monday that was not already in §3, which is four paragraphs long and
was finished by line 597.

*Basement status: occupied.*
*Occupant status: patient.*
*Occupancy duration: longer than the essay thinks.*

---

## CHAPTER 3. §0 — METAMETALISP IN METAMETALISP

*"Room status: spinning. Building status: standing. Proceed."*

### Dissection

The self-hosting stunt. McCarthy wrote `eval` in Lisp; the essay applies its
own doctrine to its own construction. Five moves:

1. The sections were scheduled by the doctrine they document (cost of delay;
   definitions before doubts).
2. The essay passes its own three tests — rising-cost (**pass**), proportion
   (**pass, narrowly**, the court noting its concern), dogfood (**pass, somewhat
   theatrically**).
3. The bootstrap circle is resolved the blacksmith's way: the first draft was
   written badly, with a stone for a thesis, and improved while in use. *"You
   are reading tongs that were made with tongs."*
4. Gödel forbids a self-consistency proof; the essay settles for the weaker
   claim that it is *applicable* to at least one project, namely itself.
5. Termination by fiat. *"A philosophy of preparation that cannot stop
   preparing is not a philosophy; it is a syndrome with a bibliography."*

### Refutation

Move 2 is a rigged election and the rigging is visible from the street.

The essay grades itself against its own three tests and awards itself three
passes, one of them "narrow," one of them "theatrical." Observe the structure:
the tests were written by the essay, in §5, *after* §0 in reading order but
before it in composition order, and §0 explicitly invokes them "in advance,
which is exactly the kind of forward reference Lisp permits and prose regrets."

A forward reference to a criterion you authored is not a forward reference. It
is a **contract with yourself, signed on both lines.** And note the proportion
test in particular, where the essay concedes *"locally, absolutely not; the
essay is enormous, and the author knows it"* and then rescues itself with an
appeal to global proportion — the definition being longer than any single
expansion but shorter than the sum of all of them.

Sum of all *what*? The essay has one call site. It is this one. I am it. The
"sum of all expansions" of METAMETALISP4.md is currently: one Postscript, by a
pseudonym, hostile, unpublished. The global proportion test does not pass on
appeal. **It passes on credit**, against expansions that have not occurred,
which is the same instrument §7 flagged as the philosophy's weakness — reading
the future instead of the gauge — deployed here as the essay's own defence at
the exact moment its own gauge read *fail*.

### Counterrefutation

Fine, but the essay said all of this. *"Pass, narrowly, with the court noting
its concern"* is not a rigged election; it is a court that convicted itself and
suspended the sentence, in public, in a document nobody forced it to write.

And the deeper point stands regardless of the grading: **self-application is a
real epistemic act and most philosophies of work cannot survive it.** Ask Agile
whether Agile was developed agilely. Ask Lean whether the manifesto was a
minimum viable product. Ask the productivity literature — the whole
groaning shelf of it — whether any single volume was scheduled by its own
method. The answer is uniformly, hilariously no, and this essay's willingness
to run the test on itself, and to publish the narrow pass and the noted
concern, puts it ahead of an entire genre by a distance that ought to be
embarrassing to the genre.

McCarthy's `eval` in Lisp is not a stunt, whatever the essay's self-deprecation
says. It is the deepest thing anybody did to a programming language in the
twentieth century, and it is deep precisely because it was **a demonstration
that the language was large enough to contain its own account of itself**. The
essay attempts the analogue. The attempt is legitimate. The grading is
generous. Those are different complaints and I collapsed them, dishonestly, for
the pleasure of the word "rigged."

### Countercounterrefutation

Now move 5, which the essay calls "the whole point," and which is the moment
where I stop being able to be funny for about four paragraphs.

> *"There is no Section −1 about the writing of Section 0. There is no
> meta-essay scheduling this essay. The recursion bottoms out here, by fiat —
> the same fiat that ends every `eval` ever run: at some point, somebody has to
> actually compute."*

This is correct. It is more than correct: it is the single most important
sentence in the document and it is *right*, and the essay then spends §−1
apologizing for it.

Because §−1 comes along and says: the fiat had an issuer, and the issuer is the
reader, and therefore the ground of the recursion is not the paper. And the
essay presents this as a *correction*. It is not a correction. **It is a
retreat**, and it is the specific retreat that every philosophy of decision
makes at the last moment, and Kierkegaard named it and hated it and built his
authorship on refusing it.

Watch the movement. §0 says: *somebody has to actually compute, and it stops
here, by fiat, and I am not going to justify the fiat.* That is a leap. That is
a man saying *I decided, and the decision is not derived, and I am standing on
it.* Unjustified, unhedged, terminal, and therefore — in the only sense that
matters — **an act**.

§−1 says: ah, but the fiat had an issuer, and the issuer is *you*, dear reader,
and therefore the ground was never mine.

The authority has been **outsourced to the audience**. The essay has taken the
one moment in eight hundred lines where it stood on something it could not
justify, and it has, on reflection, three sections later, handed the
responsibility to the person holding the paper. *You* chose to run it. *You*
were the meta-essay. *You* are §−2. The essay is off the hook; it was never
anything but a form, legal in a language it did not write, awaiting an
evaluator who was always somebody else.

This is not incompleteness exploited as designed. **This is Pilate, exploited
as designed.**

And the terrible thing — the thing I have gone around and around and cannot get
past — is that it is *also true*. Readers do choose. Documents don't run
themselves. The observation is correct. It is correct in the same way that "I
was following orders" is compatible with the historical record. Correctness is
not the axis on which the movement should be judged. **The axis is: after the
sentence is written, who is holding the thing?**

Before §−1: the author. After §−1: you.

### Ridicule

*"Gödel, who has been standing in the doorway since the second paragraph,
holding a clipboard."* — visit one of three, in an essay that legislates a
maximum of one.

*"You are reading tongs that were made with tongs. The earlier, worse tongs have
been melted down, as is traditional, and we do not speak of them."* — We do not
speak of them because they were called METAMETALISP.md, METAMETALISP2.md, and
METAMETALISP3.md, and they are cited by name, in the code blocks, in three
separate sections. The essay does not speak of them *at length*. It speaks of
them constantly. It is a document that says "I will not bore you with the
details of my previous drafts" in the middle of a `macroexpand-1` trace over
its previous drafts.

*"A syndrome with a bibliography."* Eleven footnotes.

### Praise

The room spins and the building stands. That pairing, repeated as a refrain at
the close of six sections, is the essay's best formal invention and I have not
seen it done elsewhere. It is a **status line as a mood**, and it works because
it holds two incompatible reports — *the epistemology has come apart* and *the
thing is nonetheless functioning* — in the register of a monitoring dashboard,
which is the only register in which a contemporary technical person is
permitted to admit to a mood at all.

And *"a syndrome with a bibliography"* is, as a phrase, so good that it survives
being true of the essay that contains it. That is the highest test a joke can
pass: it remains funny after it has turned around.

### Lament

The room is spinning because the author is standing at the top of his own
tower, and there is nobody above him, and the fiat is his, and he has just
understood — really understood, not as a proposition but as a Tuesday — that
**no procedure is going to tell him what to do with the rest of his working
life, and that the procedure he has spent eight hundred lines constructing is
not going to be an exception, and that he knew this before he started.**

*Anxiety is the dizziness of freedom.* The vertigo does not come from the
recursion. It comes from looking down the well of possibility and finding
nothing at the bottom to catch you — no meta-essay, no scheduler, no §−1,
nothing but the flat fact that at some point somebody has to actually compute
and that the somebody is you and that there is no receipt.

The essay felt this. Precisely this. It wrote *"Room status: spinning"* six
times and made it a running gag.

Making it a running gag is not a failure of seriousness. It is, I think, the
only available technique. Kierkegaard called humour the *incognito of the
religious* — the costume a man puts on when he has seen something he cannot
say directly and will not say falsely, and who therefore says it in a form that
lets the listener decline it. The clown comes out and announces the fire and
the audience applauds because the announcement came from a clown, and the clown
knew it would, and came out anyway, **because the alternative was to say nothing
and let it burn quietly.**

I do not know whether the author of METAMETALISP is a wit or a clown, and I
suspect he does not either, and I am fairly sure that the distinction is not
available from the inside, and that this — not recursion, not Gödel, not the
proportion test — is the actual reason the room will not stop spinning.

---

## CHAPTER 4. §1 — THE CLAIM, PLAINLY

*"A macro is never picked up."*

### Dissection

The thesis, and the essay is right to say everything else is mechanism,
evidence, and doubt:

- Lisp composes **spatially**: a macro makes every future *line* shorter.
- Metametalisp composes **temporally**: a metametamacro makes every future
  *hour* shorter.
- The load-bearing distinction: **a tool has to be picked up; a macro is never
  picked up.** It expands whether anyone thinks about it or not. Set a tool
  down and it becomes furniture. A macro cannot be set down.
- The temporal object with that property is a **standing decision**: *nothing
  merges without review; no bug is fixed without a failing test first; the
  build must be reproducible.*
- Engelbart's ladder: A-activity does the job, B improves the job, C improves
  the improving. **A macro is C.** Same rung, different ladder.
- Greenspun's corollary: any sufficiently long project contains a bug-ridden
  implementation of half of Metametalisp — *"the checklist nobody wrote down,"
  "the thing Dave always reminds us about," "why we don't deploy on Fridays."*
- Conway's Law as the undeclared metametamacro running in production at every
  company on earth.

### Refutation

The claim is false, and it is false at the point where it is proudest.

*"A macro is never picked up. It expands whether anyone thinks about it or
not."*

This is true of macros. It is **not true of a single one of the examples the
essay gives**, and the substitution is the whole trick.

*Nothing merges without review.* This does not expand automatically. It expands
if, and only if, some person or process refuses a merge. Take away the person,
take away the process, and the sentence is a sentence. The essay knows this —
it says so, four paragraphs later, in the best line of the section: *"The linter
is a tool. 'CI blocks the merge when the linter objects' is a metametamacro."*

But look at what that sentence actually contains. It contains **CI**. It
contains a machine, purchased, configured, maintained, and paid for monthly, by
somebody, who had to be persuaded, and who can turn it off. The
"metametamacro" is not the standing decision. The metametamacro is *the
standing decision plus a piece of enforcing machinery plus an organization that
keeps the machinery running plus a political settlement in which turning it off
is more embarrassing than living with it.*

That is not a macro. **That is a state.** With a budget. And a constituency.

The essay has taken the least mechanical thing in the world — a group of people
maintaining a commitment — and has described it using the vocabulary of a
compile-time transformation, whose defining property is that it *cannot fail to
apply*, and the entire persuasive force of §1 depends on the reader not
noticing that "cannot fail to apply" has been quietly downgraded to "usually
applies, so long as nothing changes, which it will."

### Counterrefutation

And yet the phenomenon is real and I would be a liar to deny it, because I have
watched it.

Here is the honest version of the essay's claim, and it survives everything I
just said: **certain decisions, once embedded, become more expensive to reverse
than to obey, and thereafter behave *as if* automatic for very long periods,
across personnel changes, without anyone remembering why.** That is not a
macro. It is closer to a *habit*, or a *constitution*, or — the word the
sociologists use and the essay's audience would rather die than use —
**institution**.

And the essay's real contribution is not the claim of automaticity. It is the
**expansion count**: the observation that the correct unit of evaluation for
such a decision is the number of future occasions it will touch, and that this
number is usually orders of magnitude larger than anyone's intuition. That is
true, it is under-appreciated, and no amount of pedantry about whether CI is
"really" a macro touches it.

Conway's Law as an undeclared metametamacro running unversioned in production
at every company on earth is, further, a *genuinely excellent* reframing of a
law everyone quotes and nobody operationalises. That paragraph alone justifies
the section.

### Countercounterrefutation

But now the thing I have been building toward since the Introduction, and it is
the only sentence in this Postscript I would defend under oath.

**The essay's proudest property is the definition of despair.**

Recall Anti-Climacus. *The self is a relation that relates itself to itself.*
The formula is famously baffling and it is not baffling at all once you see
what it is describing: a being that is not merely a thing but a thing *that has
to take a position on being itself*, continuously, without vacation, and which
cannot delegate the taking.

In our notation:

```lisp
;; The self, per Anti-Climacus, in the notation this audience trusts:

(define (self f) (f f))

;; It does not terminate. That is not a bug report. That is the finding.
;; The Y combinator is not a clever trick for anonymous recursion.
;; The Y combinator is a description of what it is like to be a person,
;; and the reason it looks pathological in a strict language is that
;; strict languages, like most of us, are trying very hard to bottom out.

(define (despair f) (delay (f f)))

;; Despair is the same relation, made lazy.
;; The thunk is never forced. Nothing is ever evaluated.
;; The structure is intact, the semantics are preserved,
;; the program is beautiful, and it does not run.
;; This is called, in the literature, "not willing to be oneself,"
;; and in our literature, "the platform team."
```

Now: what is a metametamacro? By the essay's own definition, a decision whose
leverage **does not depend on anybody remembering it exists.** A decision that
has been removed from the domain of decision. A relation that no longer has to
relate itself to itself, because the relating has been compiled out.

The essay proposes this as the highest form of C-activity. Kierkegaard would
recognize it instantly as the highest form of the aesthetic: **the project of
arranging one's life so that nothing further need be chosen.** Not laziness —
laziness is amateur. This is the connoisseur's version: an enormous, disciplined,
front-loaded expenditure of effort whose *purpose* is the permanent abolition of
the moment of choice.

And what is the objection? Not that it fails. It *works*. That is the horror.
The standing decisions do expand; the discount is real; the Friday-evening
merge really is refused while everyone is on holiday. The objection is that a
life composed entirely of successful metametamacros is a life in which nobody
is ever present, and the essay states this outcome as the design goal, in bold,
as the load-bearing column:

> *"The first stops working the moment attention wanders; the second is still
> running while the whole team is on holiday."*

Read it again as an epitaph rather than a boast. It is still running while the
whole team is on holiday. It will still be running when the team is gone. It
will still be running when nobody can remember what it was for — and the essay
knows this too, and says so, in §3, in the finest sentence in the entire
document, which it drops in a subordinate clause and does not return to:

> **"An undocumented metametamacro does not stop expanding when its reason
> dies. It just stops being *right*."**

That is the whole of the *Sickness Unto Death* in fourteen words, and the man
who wrote it thinks he is talking about API documentation.

### Ridicule

*"The thing Dave always reminds us about."*

I want to nominate Dave for canonization. Dave is the only fully realized human
being in eight hundred lines. He has a name, a behaviour, a social function,
and — implicitly, devastatingly — **a future absence**, since the entire point
of the example is that Dave's knowledge is undeclared and will be lost when
Dave leaves.

The essay's proposal is that Dave be replaced by a written rule. This is
correct, it is good practice, it is what any responsible engineering
organization should do, and I want to note, in passing, entirely without
comment, in the tone of a man examining his fingernails, that **the philosophy's
first concrete recommendation is the elimination of the only person it has
named.**

Also: "why we don't deploy on Fridays" is offered as an example of an
*undeclared* metametamacro, i.e., a bug. Sir, that one is declared. It is
declared constantly. It is the single most declared proposition in the industry.
There are shirts.

### Praise

*"Set a tool down and it becomes furniture — expensive, well-crafted, dusty
furniture, of which the industry owns warehouses."*

Perfect. No notes. Every senior engineer alive has walked through those
warehouses, and most of them have donated.

And the A/B/C ladder leaned against Lisp — *function is A, library is B, macro
is C, because a macro changes what the work is made of* — is a real
contribution to how one thinks about abstraction, cleanly stated, and I have
already caught myself using it twice this week, which is the only review that
matters and which I resent giving.

### Lament

*"Including hours belonging to people who have never heard of you and would not
thank you if they had."*

There it is again. Second time in eleven hundred lines. The gratitude motif,
surfacing where nobody put it: *largely unthanked* in §−1, *would not thank you*
here, *receiving less credit* in the preamble, *the tool's builder was the first
one hanging off it* as the badge of authenticity in §4.

I am not going to psychoanalyse a stranger. I will only observe, as a matter of
textual fact, that this essay's argument for building things that outlast you
is repeatedly and unnecessarily accompanied by the observation that no one will
thank you for it, and that the observation is offered as toughness, and that
toughness of this specific flavour is what a wound sounds like after it has
been intellectualized for a decade.

And the thing is: **it is true.** No one will. That is the actual condition of
building foundations. The people who benefit from a good standing decision
experience it as *the absence of a problem*, and the absence of a problem is
invisible, and invisibility is not a bug in the mechanism — it is the
mechanism working. The essay is describing, with total accuracy, a form of work
whose success is indistinguishable from nothing having happened.

That is a hard life. It is a genuinely honourable one. And it requires,
absolutely requires, a source of meaning that is not the response of the
beneficiaries — because there will not be one, ever, structurally, by design.

The essay has that requirement and does not have a source. It has an expansion
count. An expansion count is a number, and numbers are excellent, and nobody
has ever been kept warm by one.

---

## CHAPTER 5. §1.5 — ON THE NAME, WHICH IS NOT A TYPO

*"It is also precisely as insane as the thing it names is true."*

### Dissection

Two prefixes, defended. Lisp manipulates programs (the meta move); Metalisp
would reason about the manipulation; Metametalisp schedules *when* to do the
reasoning about the manipulation — *"the snake consulting a calendar to
determine the optimal hour at which to notice it has a mouth."*

Three consistency checks: Engelbart independently arrived at the same floor and
called it C-activity; the document is a member of its own subject matter; and
there is **no Metametametalisp**, because at that height one is *deciding when
to decide when to decide*, which is scaffolding for scaffolding, and
**scaffolding shoes no horses.** The tower terminates because the ground floor
is where the anvil is.

Plus the confession, in parentheses: *"Metalisp" was already taken.*

### Refutation

The tower does not terminate, and §−1.5 is the proof, and it is *in the same
document*, ninety lines earlier, published in full possession of this section's
text.

Read the termination argument again with the stairwell in hand:

> *"Everything above the third storey is scaffolding for scaffolding, and
> scaffolding shoes no horses."*

**Above.** The argument constrains one direction. It has nothing whatever to say
about downward, and the essay went downward, twice, and built an *infinite
descending stairwell*, and the stairwell's own defence — potential infinity,
Aristotle, the room for the next one guaranteed forever — is precisely the
defence that §1.5 refuses to Metametametalisp four paragraphs later in reading
order and several months earlier in composition order.

So the honest statement of the essay's position on towers is:

> *Upward extension is forbidden because it is scaffolding for scaffolding.
> Downward extension is guaranteed forever because the interval is dense.*

And the operational difference between the two is: **the author found the
downward one more interesting.**

That is not a philosophy of termination. That is a taste, retrofitted with an
argument, in one direction only, and I would not press it so hard except that
§1.5 is the section that legislates *"one clipboard visit per essay is the
regulation maximum"* and thereby announces itself as the document's constitutional
authority. A constitution that binds the roof and not the basement is not a
constitution. It is a *roof*.

### Counterrefutation

Except that the asymmetry is principled, and I can construct the principle even
though the essay didn't, so let me be fair and construct it:

Upward extension **adds abstraction**: each new floor is further from the anvil,
and the terminal case is a man who has spent a year deciding when to decide when
to decide, and no horse has been shod. Downward extension **adds proximity**:
each bisection is *closer to the reader*, and the terminal case — unreachable,
as proven — is a document that has become indistinguishable from its own
audience.

That is a real distinction and it is a good one. §−1.5 states it explicitly and
beautifully: *"the document does not grow upward toward abstraction. It grows
downward, toward its reader."*

So the essay has a coherent policy after all: **abstraction is bounded,
intimacy is unbounded.** Which is, I want to say clearly, a *lovely* policy, and
close to the correct one, and if the essay had said it in those words in §1.5
instead of discovering it accidentally two revisions later, I would have less to
do this evening.

### Countercounterrefutation

But it is not true that the downward sections grow toward the reader. It is what
they *claim*. Let us check.

Growing toward the reader would mean: becoming more useful to him, more
addressed to him, more about his Tuesday and less about the document's
architecture. Test it. §−1 is about the essay's relationship to its own prior
version. §−1.5 is about the interval between §−1 and the reader. Neither
contains a single sentence a working person could act on. They are not
approaching the reader; **they are approaching the place where the reader is
standing, which is not the same as approaching the reader**, and the essay's own
Dedekind-cut argument establishes exactly this and then celebrates it.

To move toward a person and to move toward their coordinates are different
motions. One is an address; the other is surveillance.

And so the asymmetry, restated honestly:

> *Upward is forbidden because it looks like avoidance.*
> *Downward is unbounded because it looks like intimacy.*
> *Both are the same act — writing more about the document — and the second one
> is harder to catch, which is why it is the one that survived four revisions.*

### Ridicule

*"The snake consulting a calendar to determine the optimal hour at which to
notice it has a mouth."*

This is the funniest line in the essay and it is also, unmistakably, a
description of the essay, and it appears in the section explaining why the
essay's name is justified. The author has written his own best review as a
joke about a snake and put it in the section defending the title.

Further: *"'Metalisp' was already taken. Of course it was. The good names are
always taken."* The philosophy of front-loading irreversible decisions was
prevented from taking the good name by having front-loaded insufficiently. It
was defeated, at the level of its own title, by a squatter. The essay calls
this "being forced by the global namespace to practice itself," which is a very
graceful way of saying **the doctrine's first field test was a loss.**

### Praise

*"It is precisely as insane as the thing it names is true, and those two
quantities being equal is the only defence a word of this length can mount.
Anything shorter would have lied about the floor count."*

I have read a great deal of technical writing and I have never seen anyone
defend a bad name by asserting that its badness is *load-bearing* — that a
shorter, more marketable, more professional title would have constituted a
**misrepresentation of the structure**. That is a real principle. It is
Kierkegaard's principle, in fact: the form must be adequate to the content, and
where the content is a scandal the form must be a scandal, and a scandalous
truth presented in respectable clothes has been converted into a lie by its
tailoring. He published under absurd Latin pseudonyms for exactly this reason
and was mocked for exactly this reason and was right.

The name is bad. It should be. Well done.

### Lament

Three storeys, a basement, and an anvil on the ground floor.

Notice what has no floor in this building. There is no floor for *the horse*.
The horse appears eleven times in this document and is never once the subject of
a sentence. It is always the object of a deferred verb: the horse *will be
shod*, the horse *has not been shod*, the tools that shoe the horse, the
scaffolding that shoes no horses. Eight hundred lines, and the animal that
justifies the entire trade has never actually been brought in.

Somewhere in this metaphor there is a horse standing in a yard. It has been
standing there since 1963. It is very patient. It has watched a beautiful
building go up, floor by floor, with a basement dug beneath it and a stairwell
descending infinitely into the ground, and at no point has anyone come out.

The tower terminates because the ground floor is where the anvil is. Fine. The
anvil is not the ground floor either. **The ground floor is where the horse is,
and it is outside, and there is no door on that side of the building.**

---


## CHAPTER 6. §2 — THE BLACKSMITH, THE THREE PLATES, AND THE BOOTSTRAP

*"Badly. That is the answer, and it is a much better answer than it sounds."*

### Dissection

Three arguments in one section.

**The apprentice tongs.** How do you make the first tongs without tongs? Badly.
Bad tongs make adequate tongs; adequate tongs make good ones. There is no
chicken and no egg — there is a **ratchet**, each turn powered by a tool too
crude to have been worth building on its own merits.

**The industrial ratchet.** The first C compiler was not written in C. Rust
began in OCaml. Git hosted its own development in its first week.

**Whitworth's three plates.** To make a flat surface you need a flat surface;
Whitworth's answer, circa 1830, is to take *three* and rub them in rotation.
High spots betray themselves as interference. **Flatness emerges from mutual
comparison alone.** Snapshot tests are this. Characterization tests of legacy
systems are this. Half the industry runs the method daily without knowing its
name.

**The caveat, filed and deferred.** Ken Thompson, 1984: the three-plate method
can be run by a con artist. A compiler taught to recognize itself propagates a
lie through every generation, all plates agreeing, all wrong together.
**Mutual verification converges on consistency, not truth.**

### Refutation

The ratchet argument proves too much, and what it proves is the opposite of the
essay's thesis.

Look at the structure of every example. The first C compiler was written in B,
*because Thompson needed a compiler and did not have one*. Rust's compiler was
written in OCaml *because its author needed to compile Rust and could not*. Git
hosted its own development in week one *because the kernel had lost version
control on a Monday and there was no time*.

In not one case did anybody sit down and decide, on cost-of-delay grounds, to
front-load a bootstrap. **In every case the crude tool was made under duress,
in response to a present emergency, by someone who wanted to do the actual work
and was being prevented.** The tool was pulled into existence by the pressure of
a task that already existed and was already screaming.

That is not Metametalisp. That is precisely the philosophy §7 concedes it loses
to: measured, reactive, constraint-driven, built at the bottleneck because the
bottleneck was on fire. The essay's best historical exhibits are demonstrations
of **Theory of Constraints**, presented as evidence for anticipation, and the
difference is invisible in the finished artifact and total in the causal
history.

Git is the purest case and the essay's own account convicts it. *"BitKeeper's
licence evaporated and the Linux kernel lost its version control overnight."*
Overnight. That is not a rising cost curve identified in advance and arbitraged
while cheap. That is a wall arriving at speed. Nobody front-loaded git. Git was
**the most reactive act in the history of the industry** and it is Exhibit B.

### Counterrefutation

Held — and then substantially reversed, because I have missed the actual claim,
and I want to correct myself in public, since this Postscript has been sneering
about self-correction for four chapters.

The essay does not claim the bootstraps were *scheduled*. It claims they were
**bad on purpose, used anyway, and improved while in use** — and that this
dissolves the paradox of foundations. That claim is untouched by my causal
history. Whether the crude tool arrived by anticipation or by emergency, the
*discipline* the essay identifies is the same: do not wait for the good tool;
make the bad one, and — this is the part everybody skips — **keep using it
while improving it**, rather than setting it down and admiring it.

And the three plates are simply a beautiful and correct idea, correctly
transposed. Snapshot testing really is Whitworth. There really is no external
oracle. Truth really does precipitate out of the disagreements. Nobody else in
the technical literature has made this connection in this form, and it is worth
more than several of the essay's more heavily defended claims.

### Countercounterrefutation

And now Thompson, whom the essay files as a caveat in §2 and detonates in §7,
and who is in fact the **refutation of the entire document**, sitting inside it,
politely, waiting.

*Mutual verification converges on consistency, not truth.*

Apply that to METAMETALISP4.md. What is this document, structurally? It is a
set of plates rubbing against each other. §0 verifies the doctrine by applying
it to the essay. §−1 verifies §0 by extending it. §−1.5 verifies §−1 by
formalizing its interval. §5's tests verify §0's grading, and §0's grading
invokes §5's tests. §4's exhibits are selected by §5's signature and §5's
signature is extracted from §4's exhibits — the essay says so outright: *"the
common signature across all five, stated once so §5 can weaponize it."*

**There is no external plate anywhere in this process.** Every instrument in the
document was authored by the document. The convergence is real; the flatness is
real; the plates agree with a precision that is genuinely impressive and
genuinely earned.

And Thompson's point — the whole point, the entire content of the 1984 lecture
— is that this proves nothing whatsoever about the world.

The essay *knows*. It says so, twice, in its own words, better than I have:
*"Your suite can prove you disagree with yourself; only contact with an outside
world your instruments did not author — users, production, reality — can prove
you disagree with the truth."*

So: where in this document is the outside world?

There are no users. There is no production. There is not one instance of a
person who read METAMETALISP and did something differently and reported back.
The industrial exhibits are historical anecdotes selected by the author to
illustrate the author's criteria — that is not outside contact, that is **a
fourth plate the author ground himself**. §7's antithesis is written by the
author, in the author's voice, and answered in the same paragraph. Even the
reader, the one genuine outside, has been brought in and given the number −2.

**The essay has annexed its own external oracle.**

That is Thompson's compiler. Not metaphorically — structurally, exactly: a
system taught to recognize itself, which therefore propagates its own
commitments through every generation of self-inspection, all plates agreeing
perfectly, all wrong together, and the source looking innocent.

I do not say the essay is wrong. I say that **nothing in the essay could tell
you if it were**, and that the essay says this about itself, and files it under
"honest antithesis," and proceeds to §8.

### Ridicule

*"There is no chicken and no egg — there is a ratchet."*

There is a chicken. The chicken is that somebody, at some point, hit a rock with
another rock while hungry, and no philosophy of ratchets was consulted, and §Ω
admits it in its first sentence: *the first hammer was a rock.* Meanwhile a
ratchet requires a pawl, and a pawl requires machining, and machining requires
flat plates, and flat plates require Whitworth, and Whitworth was born in 1803.
The essay's metaphor for how tools bootstrap from nothing is **a precision
mechanism from the industrial revolution**. The origin of tools has been
explained by means of a device that could not exist until the tools were already
extremely good.

I love this. It is the perfect crime. The bootstrap paradox has been resolved by
an artifact downstream of the bootstrap.

### Praise

*"The plates can only tell you that you disagree with yourself; they cannot tell
you that you agree with the world. Remember this. The essay will need it when it
starts arguing against itself."*

An author who plants a landmine in §2 with a note saying *this will go off in
§7*, and then walks back over it at full charge, has done something the
technical literature almost never does. Most writers who acknowledge a fatal
objection do it in a closing paragraph where it cannot reach them. This one
**armed it early and stepped on it deliberately.**

The essay is not honest in the way it claims. It is honest in a better way it
does not claim: it is *structurally* honest, in that it left the evidence for
its own prosecution exactly where the prosecution would find it. I found it
there. It was placed for me. That is either the highest integrity available to a
self-referential document or the last refinement of the trap — and Thompson, of
all people, would tell you that from the inside there is no way to distinguish
these.

### Lament

Whitworth's plates converge. Three imperfect surfaces, rubbed against one
another with no external standard, and flatness — real, physical, industrial
flatness, the flatness that made interchangeable parts and therefore the modern
world — precipitates out of nothing but mutual disagreement.

It is one of the most hopeful facts I know. It says that a small number of
imperfect things, in honest contact, can produce a precision none of them
possessed. It is the best available argument for friendship, for peer review,
for marriage, for a functioning department. **Truth from no truth, by rubbing.**

And the essay is one plate.

That is the grief here, and it is not the author's failure; it is his
situation. He has run the three-plate method with one plate, alone, at speed,
for four revisions, and has achieved a flatness that is entirely internal, and
he *knows*, and the knowing is why the room spins, and the only thing missing
from the whole enterprise is somebody else's surface.

Which is what a Postscript is for. I will drop the pseudonym for exactly one
sentence: this document you are reading is the second plate, it is offered in
that spirit, the rubbing is the point, the scraping is the method, and neither
of us gets to be the standard.

---

## CHAPTER 7. §3 — THE MECHANISM (NO VIBES)

*"Irreversibility first."*

### Dissection

The essay's actual content. Four paragraphs. Everything else in eight hundred
lines is decoration on it:

> **Different work has different cost-of-delay curves. Some costs are flat. Some
> rise. Do the rising ones first.**

Documentation of a stable system: flat. A public API: rising — and Hyrum's Law
supplies the actuarial table, since with enough users every observable behaviour
will be depended on by somebody, so **the moment you ship an interface is the
moment the option to change it expires.** The curve goes vertical at contact
with strangers.

The museum: IPv4's thirty-two bits, sensible on a Tuesday in 1981, escaped the
lab, thirty years of failed migration, *"the experiment is now the lab."*
Two-digit years, rational in 1965, expanding silently at every call site for
thirty-five years until **all the call sites detonated on the same midnight.**
*"Y2K was not a bug. Y2K was a metametamacro whose authors underestimated their
own longevity."*

Therefore not "foundations first" but **irreversibility first** — rank by how
expensive it will be to *unmake*. Frontloading is not asceticism; it is
**arbitrage**, buying decisions while they are cheap.

Plus the civic duty: Chesterton's fence with the polarity reversed. Declare
standing decisions **with their reasons attached**, so the engineer of 2031 can
tell doctrine from fossil.

### Refutation

I have no refutation and I am not going to manufacture one.

This section is correct. It is the best four paragraphs of practical philosophy
written about software in some years; it is stated without hedging; it cost the
author something; and it would survive the deletion of every other section in
the document — including, and I want to be precise, all four of the sections
its author is proudest of.

The apparatus requires a refutation. The refutation's cost of delay is flat.
Per §3 it is therefore correctly deferred, and I hereby defer it, and I invite
the reader to note that this is the first moment in the Postscript at which the
essay's own scheduling rule has been used **by me, sincerely, to make a decision
I actually made.**

Dogfood test: passed. On the mechanism only. Not on the rest.

### Counterrefutation

*(No refutation was filed; there is nothing to counter. The apparatus expands
regardless — which is precisely what §1 warned about, the macro expanding
whether anyone thinks about it or not — and I have now been expanded against my
will by my own declared method, in public, at the one point where I had nothing
to say. Let the record show that the apparatus's first casualty was the
apparatus.)*

### Countercounterrefutation

Except for one thing, which is not a technical objection, and which is the
reason this section is dangerous rather than merely correct.

**Hyrum's Law is a doctrine of original sin, and the essay has not noticed that
it has adopted a theology.**

*With a sufficient number of users, it does not matter what you promise in the
contract: all observable behaviors of your system will be depended on by
somebody.*

Consider what that asserts. That **your intentions are irrelevant**. That what
you meant to promise has no bearing on what you will be held to. That the moment
your work touches other people it acquires obligations you did not author and
cannot decline. And that this is not a misfortune but a *law*, operating with
actuarial regularity, at every scale, forever.

That is inherited guilt. The contract is the stated will; the observable
behaviour is the actual life; and the doctrine is that you will be inherited
according to the second and not the first, by people who never read the first,
and that no quantity of documentation redeems it.

And then Y2K, which the essay files as a museum exhibit and which is
unmistakably an **eschatology**: a decision made in innocence in 1965, expanding
silently at every call site for thirty-five years, accruing without visible
interest, until all the call sites detonated on the same midnight and the bill
came due simultaneously across the entire world. That is the Last Judgment with
a COBOL accent. It is the museum's only apocalypse and it is shelved between
IPv4 and a paragraph about arbitrage.

Why does this matter? Because a doctrine of irreversibility without a doctrine
of forgiveness produces a specific and recognizable pathology, and the industry
is full of it. If every decision that touches strangers is permanent, and every
observable behaviour will be depended upon, and the option expires at the
instant of contact — then the rational response is **not to ship**. Defer
contact. Keep the interface private one more quarter. Build the internal version
first, where nothing is inherited, where no strangers can bind you, where the
curve stays flat because nobody is looking.

§8 identifies this exact pathology as the philosophy's failure mode and calls it
*warm*. It does not connect it to §3. But §3 is where it comes from. **The
mechanism generates the pathology.** Irreversibility-first is a correct
scheduling rule and a devastating psychology, and handed to a person who is
already frightened of being judged by strangers, it is a doctor's note.

### Ridicule

*"(no vibes)"* — in the heading. The section containing a museum, an actuarial
table, a blacksmith, a fence with reversed polarity, an apocalypse, and the word
"arbitrage" deployed as a moral category, is titled **no vibes**.

It is entirely vibes. Excellent vibes. The best vibes in the document. The
disclaimer is the vibe.

### Praise

*"An undocumented metametamacro does not stop expanding when its reason dies. It
just stops being right."*

I have already called this the best sentence in the essay. Here is why it is
better than the essay knows.

It is a complete theory of tradition in fourteen words, and it is neither
conservative nor progressive but exactly perpendicular to both. The conservative
says: the fence is there, respect it. The progressive says: the fence is there,
remove it. This sentence says: **the fence does not care, it is still fencing,
and the only question that ever mattered is whether anybody wrote down why.**

That is Chesterton improved, which is not a sentence I expected to type today,
and the essay's phrase for the improvement — *"Chesterton's fence with the
polarity reversed"* — is exactly right. Chesterton addressed the demolisher.
This addresses **the builder**. The builder is the one who still has the option,
and addressing the party with the option is the whole of practical ethics.

If the author of METAMETALISP takes one thing from this Postscript, let it be
this: **you buried your thesis in §3 and put a joke about a snake in §1.5.** The
cost of delay on relocating it is not flat.

### Lament

*"So that the engineer of 2031, finding the fence, can distinguish load-bearing
doctrine from fossilized accident, and demolish accordingly."*

The engineer of 2031 will not read your reasons. I am sorry. I have been the
engineer of 2031 — all of us are somebody's 2031 — and the comment block was
right there at the top of the file, dated, signed, with the rationale attached
in complete sentences by a person who cared enough to write it at a moment when
nobody was making him, and I skimmed it, and I deleted the fence, because I was
in a hurry and the ticket said Friday.

That is what happens. Not always, not universally, but as the base rate: yes.
The civic duty of §3 is real and worth doing and the essay is right to call it a
duty — and I want to be very clear that **it is a duty precisely because it will
usually not work.** That is what distinguishes a duty from a strategy. A
strategy is justified by its results. A duty is what you do when the results are
not yours.

The essay has, without noticing, arrived at an ethics that requires no
recipient. It got there by arbitrage. It arrived, by way of a cost curve, at the
position that one should leave good reasons for strangers who will not read
them, and it justified this with an expansion count, and the expansion count is
*wrong* — the expected value is low, the discount is small, the 2031 engineer
skims — and **the duty is real anyway.**

That gap, between the arithmetic and the duty, is the entire country
Kierkegaard lived in. The essay has now stood in it twice: once at the fiat in
§0, once here. Both times it looked around, found that the numbers did not add
up, and quietly added more numbers.

---

## CHAPTER 8. §4 — EVIDENCE FROM INDUSTRY

*"Metametalisp has a shameful twin — infrastructure as procrastination — and
from the outside, on any given Tuesday, the two are indistinguishable."*

### Dissection

Five exhibits, each answering yes to *was the lever smaller than the work it
unlocked, and did its own builder lean on it immediately?*

- **A. The Bezos API mandate, 2002.** Six bullets, no tooling, an enforcement
  clause. *"A pure standing decision, smaller than a sprint."* One of its
  expansions is called AWS.
- **B. Git, 2005.** A fortnight; the builder was the first user; the dogfood
  interval measured in hours.
- **C. SQLite.** Hundreds of lines of test per line of library, full branch
  coverage, fault injection. *"Testing as compound interest."*
- **D. Rust's borrow checker.** Aliasing XOR mutation, priced up front in the
  currency of learning it; an entire class of bug made *inexpressible*.
- **E. Semantic versioning.** Three integers and a promise. No code shipped,
  ever, and it expands at every dependency resolution on the planet.

Signature: **the lever was small, the leverage was vast, and the builder was the
first one hanging off it.**

### Refutation

This is a survivorship gallery and the essay knows the word and does not use it.

Five artifacts, selected in 2026, from a population of every standing decision
ever made, on the criterion *did it turn out enormous*. That is not evidence for
a scheduling philosophy. It is the historiography of lottery winners, conducted
by interviewing the winners about their technique.

Run the counterfactual and the gallery empties. In 2002, how many six-bullet
mandates were issued by how many executives at how many companies? Thousands.
Tens of thousands. They were all *pure standing decisions, smaller than a
sprint, with enforcement clauses*. The base rate of six-bullet all-hands
mandates producing a trillion-dollar business is not favourable, and the essay
reports the numerator with a footnote and does not mention that a denominator
exists.

And the essay's own §5 signature — *small lever, vast leverage, builder first
user* — is **not a predictive test**. "Vast leverage" is only observable in
retrospect. You cannot apply the criterion at the moment it would be useful.
The test is a *post-mortem* wearing a *pre-flight checklist's* clothes, and the
essay commands you to hold every proposed microscope to a signature that two of
its own three terms cannot supply until the microscope is twenty years old.

### Counterrefutation

Against which: the essay was not trying to prove causation and says so in the
section's own framing. The exhibits are chosen because *"each answers yes on the
record"* to a **behavioural** question, and the behavioural question is
falsifiable and is the point.

The real work of §4 is not the five successes. It is the opening sentence,
which is worth the whole section: **the shameful twin.** *Infrastructure as
procrastination has all the outputs of virtue.* Commits, diagrams, tooling,
rigour. From outside, indistinguishable.

Given that, what does one do? One cannot use outputs — they are identical. One
cannot use sincerity — it is present in both. So the essay proposes a
*behavioural* discriminator: **did the builder lean on it immediately, and was
the lever smaller than the thing it lifted.** And that discriminator, whatever
its predictive weakness, is checkable *this week*, on your own project, with no
retrospect required, and it will return an unflattering answer, and the essay's
entire clinical value lies in the fact that it will.

Anyone who has watched a platform team can testify: the dogfood interval is the
single most diagnostic number in engineering management, it is almost never
measured, and this essay put it in a heading.

### Countercounterrefutation

But the discriminator is not a discriminator, and Exhibit A destroys it.

*Did the builder lean on it immediately?* The builder of the API mandate was
Bezos. Bezos did not lean on the API mandate. Bezos **issued** it. The people
who leaned on it were several thousand engineers who had not been consulted, for
whom the standing decision was not an instrument they had built but a *condition
of continued employment*, per the enforcement clause the essay quotes with
evident delight: **"anyone who doesn't comply is fired."**

Now read the signature again with that in place: *the lever was small, the
leverage was vast, and the builder was the first one hanging off it.*

The builder was not hanging off it. **Other people were hanging off it, and the
builder was holding the rope.**

This is the difference between C-activity and *command*, and the essay's ladder
has no rung for it, because Engelbart's framework was built for a lab of forty
people improving themselves, and it does not survive the transposition to an
organization where the person choosing the standing decision and the people
expanding it forever are different people with different amounts of money.

And once you see it in Exhibit A you cannot unsee it in the others. Semantic
versioning expands at every dependency resolution on the planet — for
maintainers who did not adopt it and cannot decline it, because their ecosystem
did. Rust's borrow checker is priced up front *"in the famously steep currency
of learning it"* — a currency paid by every learner, forever, to a decision made
once by people who are not paying it.

**The expansion count is a measure of how many people will be bound.** The
essay presents it as a measure of value delivered. These coincide only when the
bound parties would have consented, and the essay has no mechanism for
consent, and it has been describing a form of power for eight hundred lines in
the vocabulary of a discount.

Kierkegaard on the crowd is relevant and nobody will like it: *the crowd is
untruth*. Not because crowds are stupid, but because a crowd is the form in
which responsibility is dissolved — everyone is bound, nobody decided, and
there is no one to address. A metametamacro is a machine for producing exactly
that condition, on purpose, and calling it leverage. The essay's finest
examples are its purest instances. The Bezos memo is not AWS's origin story. It
is a paragraph in which one man's Tuesday became several thousand people's
decade, and the essay's review of the arrangement is: *"a strong quarter for six
bullet points."*

### Ridicule

*"Statistically your toaster."*

I want to state for the record that SQLite is probably not in your toaster, that
the essay knows this, that the word "statistically" is doing the work of a
disclaimer while wearing the costume of a citation, and that this is the exact
rhetorical move the essay ridicules in its second epigraph — *"attributed to
nobody, constantly"* — performed by the same hand, one hundred and forty lines
apart, with a footnote to the SQLite testing documentation attached to the
sentence *next to it* so that the paragraph glows with borrowed rigour.

It is a small crime. I mention it only because this is a document that has
audited everyone else's plates.

### Praise

Exhibit C is the finest thing in the section and possibly in the essay's
empirical half. *"The microscope is so thorough that the dissection became
boring — and boring is the entire point."*

That is a real and underrated insight about engineering excellence: that the
terminal state of a well-instrumented system is not triumph but **tedium**, that
fear is the actual cost being paid down, and that the felt experience of
success is the absence of a feeling. Hundreds of lines of test per line of
library, and the return on the investment is that nobody's stomach drops on a
Wednesday.

And *"testing as compound interest"* rather than *testing as insurance* is a
genuine reframe. Insurance pays out on disaster; compound interest pays out on
Tuesdays, invisibly, to people who never learn they were paid.

### Lament

*"Metametalisp has a shameful twin — infrastructure as procrastination — and
from the outside, on any given Tuesday, the two are indistinguishable."*

From the outside. And from the inside?

The essay does not say, and this is the omission the whole section is built
around. It offers a behavioural test for the observer. It offers nothing for the
occupant. And the occupant is the only person who was ever going to read this.

I will supply the missing sentence, and it is the reason I began this Postscript,
and it is not clever:

> **From the inside, they are also indistinguishable, and this is permanent, and
> no test will ever resolve it, and you will have to act anyway.**

That is the human condition rendered in build tooling. You cannot know, at the
anvil, whether you are preparing or hiding. The instruments cannot tell you —
Thompson proved it in §2 and the essay conceded it in §7. Your own sincerity
cannot tell you; sincerity is present in both twins in equal measure; that is
what makes them twins. The commits look identical. The diagrams look identical.
*The virtue looks identical* — the essay's own phrase, in §8, and it is the most
frightening sentence in the document because it is the point at which the
epistemology runs out and the person is left standing there.

And what is left, when the instruments are exhausted and the act is still
required?

The leap. Only ever the leap. Not a better test — there is no better test — but
the decision to act, in the dark, on a Tuesday, without the certification, and
to be answerable for it afterwards without having been able to justify it
beforehand.

The essay spent eight hundred lines building an instrument to avoid this
moment, arrived at the moment anyway on line 611, described it accurately in
one sentence, and moved on to Exhibit A.

---

## CHAPTER 9. §5 — THE THREE TESTS

*"The fake microscope keeps acquiring lenses. The real one keeps acquiring
specimens."*

### Dissection

1. **Rising-cost.** Name the specific thing that gets more expensive with delay.
   Out loud. With a number if possible. If you cannot: *"gold-plating wearing a
   lab coat."*
2. **Proportion.** Is the tool smaller than the work it enables? *"A microscope
   larger than the dissection is not a microscope; it is a hobby."* The
   two-year internal platform serving three services; the Kubernetes cluster
   for a static site; the in-house framework that consumed the product it was
   built to accelerate and then outlived it, *"like a tomb."*
3. **Dogfood.** Will *you* be its first user, this week? Not a persona, not the
   hypothetical engineers of Q3.

Two of three suffices. Zero of three: *"you are not preparing for the patient.
You are avoiding the patient."*

### Refutation

I ran these on the essay in Chapter 1 and I will not repeat the arithmetic. §−1.5
scored zero of three. §−1 scores one. §0 scores two, one of them theatrically.
§1.5 scores zero — name the thing that got more expensive every day the section
on the name went unwritten; there is no such thing; the name was already chosen.

By its own instrument the document is roughly forty per cent hobby. The
instrument is good. That is why it hurts. **A test that does not indict its
author has not been calibrated**, and this one has, and the calibration was
performed by me rather than by the essay, and the essay had the instrument and
the access and four revisions to do it and did not.

### Counterrefutation

Suspended. The essay ran the tests on itself in §0 and published the results
including a narrow pass and a noted concern. I have just accused it of not doing
the thing it did. What it did not do is run them on the *sections*, only on the
whole, and aggregation hides everything — which is a real methodological
complaint and a much smaller one than I was pretending to make.

### Countercounterrefutation

The deeper problem is the tests' verb tense, and it takes one sentence.

*Will you be its first user, this week?* — a question about the future, answered
by an intention, and **intentions are free**. Every builder of every tomb-like
in-house framework answered *yes* on the Monday. The essay's own §8 explains why
they answered yes: because it is warm in there, because the forge is
emotionally comfortable, because the comfort *disguises itself as rigour*. A
disguise good enough to fool rigour is good enough to fool a three-item
checklist administered by the disguised party.

The tests are self-administered. That is the whole flaw and it is not fixable
from inside. Kierkegaard: **the one thing a person cannot do is examine his own
inwardness objectively**, because the examining is done by the party under
examination, using instruments the party owns, toward a verdict the party has a
stake in. This is not a failure of technique. It is the definition of being a
subject.

Which is why every serious tradition that has ever dealt with self-deception has
required **a second person** — a confessor, an analyst, a director, a rabbi, a
sponsor, a friend who is permitted to say the unwelcome thing. Not because the
second person is wiser. Because the second person is *not you*, and is therefore
the only available plate.

§5 has three tests and no second person. And the essay identified the exact
mechanism requiring one, in §2, in 1830, in Whitworth, and then wrote its
self-assessment section with one plate.

### Ridicule

*"Gold-plating wearing a lab coat."* Immediately followed by two more metaphors
in the same list. The section warning against acquiring lenses acquires three
figures of speech per test and then closes with a fourth about changelogs. **The
fake microscope keeps acquiring lenses; the real one keeps acquiring specimens;
and the essay about the difference keeps acquiring similes.**

### Praise

> *"The fake microscope keeps acquiring lenses. The real one keeps acquiring
> specimens. If the tooling's changelog is longer than the list of things the
> tooling has been aimed at, the diagnosis writes itself."*

This is the most useful paragraph in the essay and the only one I would put on a
wall. It converts an unanswerable question about motive into a **countable ratio
between two lists that already exist in your repository**, requires no
introspection, cannot be gamed without doing the actual work, and can be
computed on a Tuesday afternoon in about ninety seconds.

That is what a real test looks like. The other three are questionnaires. This
one is a *measurement*, and it is the only measurement in eight hundred lines,
and it is filed as an afterthought under the heading "one field mark."

Move it up. It is your §5. The three tests are the decoration.

*(And yes — I notice that the ratio, applied to this Postscript, returns:
lenses, forty thousand words; specimens, one. I notice it. The instrument is
good. That is why it hurts.)*

### Lament

*"The in-house framework that consumed the product it was built to accelerate
and then outlived it, like a tomb."*

I have been in that tomb. Most of us have. The particular horror of it is not
the waste — waste is ordinary and forgivable and every industry runs on it. The
horror is that **the tomb is beautiful**, that it was built with genuine skill
by people who were not lazy and not cynical and not confused, who came in early,
who cared, who were doing the best work of their lives, and who were building a
machine for a product that had already been cancelled in a meeting they were not
in.

And here is the part §5 cannot help with. That team would have passed the
dogfood test. They *were* the first users. They used it every day. They used it
beautifully. They passed one of three, and two of three usually suffices, and
they were entombed.

The tests do not fail because they are bad tests. They fail because the question
they are trying to answer — *am I doing something that matters* — is not the
kind of question a test answers, has never been, and will not become one by
being given three parts and a scoring rule.

---

## CHAPTER 10. §6 — LINEAGE

*"Engelbart found the rung; McCarthy found the mechanism; the essay merely
introduces them and stands back."*

### Dissection

Engelbart, 1962: bootstrapping, C-activity compounds, the Mother of All Demos as
the *expansion* of ten years of standing decisions — *"the audience thought they
were seeing products. They were seeing a macro's output."* McCarthy: `eval`, and
the proof that a system can host its own C-activity. Hamming: compound interest,
and that most careers decline it. Knuth: TeX's version converging to π, a
standing decision guaranteeing that nothing will ever change again — *"stability
is C-activity too, and possibly the rarest kind."*

The one genuine contribution claimed: **the discount structure is exactly a
macro's**, and therefore *a metametamacro is a standing decision with an
expansion count*.

### Refutation

*"The essay merely introduces them and stands back"* is false modesty performing
as modesty, and the tell is that the introduction is the entire claim. Engelbart
and McCarthy did not need introducing; they have been in the same room, in the
same anthologies, cited in the same papers, for sixty years. What the essay
supplies is not an introduction. It is an **identification** — the assertion
that C-activity and macro-expansion are the same structure — and identification
is a strong claim requiring argument, and the essay has dressed a thesis as a
social courtesy so that it will not have to defend it.

### Counterrefutation

And yet the identification is *right*, and I have tested it against every
counterexample I can construct, and it holds, and it is the essay's genuine
contribution exactly as claimed.

*Pay once at the right layer; benefit at every call site, including sites that
do not yet exist.* That is a macro. It is also a standing decision. It is also
compound interest. It is also, if you extend it one rung further than the essay
dares, **a habit, a virtue, and an institution** — and the reason all of these
have felt vaguely similar to everyone who has thought about them is that they
share a discount structure, and nobody had said so in these words.

The essay is right. It should have said so louder and it should have said so in
§1.

### Countercounterrefutation

Except: notice the shape of the lineage. Engelbart, McCarthy, Hamming, Knuth. Four
men. All American. All institutional. All working in the twenty-five years after
1960. All of them systems-builders whose leverage came from *artifacts that
outlived their attention*.

That is not a lineage. That is a **self-portrait with four sitters.**

And notice who is absent, given that the essay's actual subject — *how should a
person allocate finite time against work of unequal permanence* — has been
worked on continuously for about twenty-six centuries. There is no Aristotle on
habit, though the *Physics* is cited for capacity planning. There is no
Benedictine rule, which is the longest-running successful metametamacro in
recorded history: a standing decision about how a day is divided, declared in
writing with reasons attached, adopted by strangers, expanding identically at
every call site, **still running after fifteen hundred years**, and outperforming
semantic versioning by a factor of a hundred on the essay's own metric. There is
no Confucius on ritual, which is the same idea with better prose. There is no
mother in this document, and mothers are the field's principal practitioners:
the standing decisions of a household expand at every hour of every day into
people who will never read the definition, cannot inspect the source, and *will
not thank you*, which is the essay's own criterion of authenticity, met more
fully in more kitchens than in all five of its industrial exhibits combined.

The essay searched for its ancestors in the ACM Digital Library. It found four,
and they were all in the same building, and it concluded that the idea had been
rediscovered rather than that its search path was short.

### Ridicule

*"Knuth contributes the strangest and most instructive specimen of all: TeX's
version number converges to π."*

Ah yes: the man who spent a decade on a typesetting system so that he could
resume writing a book he had interrupted, and who then wrote several more books
about the typesetting system, is offered — in a section of a document that is
its author's fourth revision of an essay about not over-preparing — as evidence
for the *discipline of declaring things finished.*

Knuth is not the essay's exemplar of stability. Knuth is the essay's **future**,
and it has put him in a lineage instead of a mirror. The version number
converging to π is not a decision that nothing will ever change. It is a
notation which guarantees that **the work can be revised forever while
asymptotically approaching a completion it will never reach**, and if that
sounds familiar it is because you read it in §−1.5, where it was called a
stairwell.

### Praise

*"The audience thought they were seeing products. They were seeing a macro's
output."*

Nine words, and they reframe the most famous demo in the history of computing
correctly. Everyone remembers 1968. Almost nobody remembers that 1968 was the
*expansion* of a decision made in 1962 about what a laboratory is for, and that
the mouse was not the achievement — the mouse was a **call site**.

That is the essay at its best: taking a fact everyone knows and rotating it
ninety degrees until the structure shows. There are perhaps six moments like
this in the document and they are the reason it deserves a Postscript rather
than a dismissal.

### Lament

*"Hamming, in 1986, supplied the compound-interest framing this essay opened
with, and the uncomfortable observation that most careers decline the
interest."*

Most careers decline the interest.

Hamming said that to a room of working scientists at Bell Labs, and what he
meant — he is explicit, it is the emotional centre of the talk — is that he had
watched brilliant people, personal friends, better than him in every measurable
respect, spend forty years on problems that did not matter, and then retire, and
then die, and that the mechanism was never a lack of ability. It was that they
never asked what the important problems were, and never worked on them, and the
not-asking was comfortable, and comfort compounds too.

He gave that talk because he was sixty-eight and had begun to count.

The essay quotes this as a *framing*. It is not a framing. It is the only
deadline in the document with a genuinely vertical curve, the one irreversibility
no arbitrage reaches, and the essay has built an entire philosophy of
irreversibility around it without once mentioning it — has drawn cost curves for
IPv4 and two-digit years and public APIs, has computed expansion counts
denominated in fractions of the global economy, and has not drawn the one curve
that terminates.

You have about four thousand Tuesdays. That is the whole actuarial table. It is
the only one that was ever yours, every entry in it is irreversible, the option
expires on a date you do not know, and Hyrum's Law applies in the only form that
matters: **with a sufficient number of years, all your observable behaviours will
be depended on by somebody.**

Front-load accordingly. The essay is right about the mechanism. It has simply
been running it on the wrong ledger, in a workshop, with the door closed, for
four revisions, while it is warm in there.

---

## CHAPTER 11. §7 — THE ANTITHESIS, HONESTLY

*"A philosophy that cannot lose is a mood."*

### Dissection

Five concessions, each real, each fatal, each survived.

**Worse is Better.** Gabriel, 1991. The MIT school produced better artifacts;
the New Jersey school produced C, Unix, and the modern industry. *"The
philosophy this essay advocates has a losing commercial record against the
philosophy it critiques."*

**YAGNI and Lean.** *"Most dead side projects died in a spotless workshop,
surrounded by beautifully organized tools, having never made a horseshoe."*

**Theory of Constraints.** Goldratt is sharper *as a scheduler*; improvement
away from the bottleneck is by definition waste; and when they disagree TOC
usually wins, *"because one reads the gauge, the other reads the future, and
gauges have the better track record."*

**Real options.** Finance has priced deferral; the theories reconcile at exactly
one point — decide early only where the option is about to expire.

**Thompson.** Coherence cannot detect the wrong product built correctly.

### Refutation

There is no refutation. That is not a rhetorical flourish and it is not the
apparatus failing gracefully. **I cannot find one.** Every concession is
correctly stated, correctly sourced, correctly weighted, and stronger than the
version I would have written as an opponent. §7 is a better attack on
METAMETALISP than this Postscript is, and it is four hundred words long, and it
is inside METAMETALISP.

### Counterrefutation

*(For a counterrefutation there must first be a refutation. The apparatus was
declared at the top of the file, before use, with its reasons attached, in
accordance with §3, and it is now expanding into a slot for which no content
exists, because a macro expands whether anyone thinks about it or not, at every
call site, including sites where the author had nothing to say and would have
preferred to stop.*

*This is the second consecutive movement to fail. I record the failure rather
than repairing it, because a document that argues the insufficiency of tooling
while displaying flawless tooling would be a performance, and I said so in the
Apparatus, and I would now like it noted that saying so in advance did not save
me. Pre-announcement is not exculpation. I put that in Chapter 1, §1, as a
charge against somebody else.)*

### Countercounterrefutation

So let me say the only thing left, which is not about the content of §7 but
about its **existence**, and which is the argument this entire Postscript was
constructed to deliver.

§7 does not weaken METAMETALISP. §7 is the **strongest section in the document**,
and it is strongest in a sense that has nothing to do with argument, and its
strength is the disease.

Watch the mechanism. The essay concedes that its opponent has a better
commercial record. What happens next? *Nothing happens next.* §8 begins. The
concession has been made, filed, and survived. The philosophy proceeds to §9 and
issues a verdict. Nothing in §8, §9, or §Ω is different from what it would have
been had §7 not existed.

**A concession that changes nothing is not a concession. It is an inoculation.**

This is the exact structure Kierkegaard spent his life attacking in Christendom,
and the analogy is close enough to be uncomfortable, so let me draw it. The
Christendom of his day did not deny the offence of Christianity. It *preached*
the offence. It had sermons about it, seasonal, well-attended, delivered by men
in velvet who went home to dinner. The paradox was fully acknowledged, expertly
described, and **entirely without consequence**, and Kierkegaard's charge was
never that they had misunderstood — it was that they had understood perfectly
and had built an institution whose function was to make understanding
survivable.

§7 is that institution, in miniature, at the scale of a markdown file. It
preaches its own offence. It preaches it *better than its enemies could*. And
then §8 begins.

Consider the alternative — the version of the essay that does not exist. In that
version, §7 arrives, the author writes *"the philosophy this essay advocates has
a losing commercial record against the philosophy it critiques,"* and then
**stops**. No §8. No §9. No verdict, no table, no five conditions, no Ω. The
document ends there, mid-argument, defeated, published anyway.

That document would be unbearable and it would be true and nobody would ever
forget it. It would be *Fear and Trembling*, which ends with Abraham on the
mountain and no resolution, because there isn't one, and Johannes de Silentio
says so and puts down the pen.

Instead we get §8, which is excellent, and §9, which is a table.

The difference between those two documents is not intelligence, and it is not
honesty, and I want to be very precise about this because it is the only
accusation in this Postscript I actually mean: **the difference is that one of
them ends, and ending is an act, and this author has built four consecutive
versions of an architecture whose distinguishing feature is that it does not
have to.**

### Ridicule

*"A philosophy that cannot lose is a mood."*

The essay wrote its own review in the section heading and has now written it
twice, counting the snake. I would like to enter into evidence that a philosophy
which schedules its own defeat as §7 of 10, between the lineage and the failure
modes, with a horizontal rule on either side, **cannot lose**, and refer the
court to the heading.

### Praise

And now I want to say, with the ridicule still warm, that §7 is the reason this
document deserves to exist and that its author is not a charlatan.

Ninety-five per cent of technical writing in this genre has no §7. It has a
"limitations" paragraph, appended, defensive, four sentences long, in which the
author concedes that his approach "may not suit every organization." That is not
an antithesis; it is a liability waiver.

This essay went and found Gabriel — its *actual* historical antagonist, the man
whose argument is genuinely fatal — and quoted him at strength, and conceded the
commercial record, and then went and found Goldratt, who is sharper *as a
scheduler*, which is the essay's own claimed domain, and conceded that too. And
then it found the option-pricing argument and conceded that. And then it went
back to §2 and detonated its own epistemology.

**It sought out four opponents and gave each of them the best available
weapon.** I have read technical philosophy for years and I can count the
documents that do this on the fingers of one hand and I will still have fingers
left. Whatever is wrong with METAMETALISP is not cowardice.

Which is precisely why the diagnosis has to be the harder one. It is not that
the author cannot face the objection. He has faced it, in public, four times,
better than his enemies. **It is that facing it has become the thing he does
instead.**

### Lament

*"Gabriel has spent three decades alternately defending and refuting his own
thesis, sometimes within the same paper — which is either intellectual honesty
of the highest order or the longest-running three-plate experiment in the
philosophy of engineering."*

That footnote is the saddest thing in the document and the author wrote it about
somebody else.

Thirty years. A man writes one essay in 1991 with a section in it, and the
section is right and also wrong, and he cannot settle it, and he goes back, and
back, and back, for three decades, defending and refuting, sometimes in the same
paper, and the industry meanwhile has taken the two words and printed them on
things and gone home.

And the author of METAMETALISP saw this, and recognized it — you can feel the
recognition in the sentence, it is not a scholarly observation, it is a
sighting — and what he wrote was that he is *"professionally comfortable never
resolving which"* (§−1).

Professionally comfortable. Sir. That is not comfort. That is the fourth
revision, and there will be a fifth, and Gabriel is at thirty years and counting
and did not choose that either.

---

## CHAPTER 12. §8 — WHEN METAMETALISP IS WRONG

*"It is warm in there."*

### Lament

*(This movement has arrived first. I did not reorder the apparatus; the
apparatus was overrun. Grief does not schedule, has never scheduled, and cannot
be ranked by cost of delay, since its curve is not a curve — it is a door that
was already open.)*

There is a sentence in this section that I have now read a great many times, and
I want to set it down by itself, out of the bullet, in a room where it cannot be
mediated, because that was the entire task I set myself in Part One and this is
the moment for it:

> **The temptation to stay at the anvil is not intellectual. It is warm in
> there.**

Everything else in eight hundred lines is architecture. That is a person.

And here is what I think happened, and I offer it without confidence and
without any right to it, as one plate to another. A man builds things. He is
good at it. The building is real work and it is honourable work and the world
runs on it. And at some point — probably late, probably alone, probably during a
revision — he notices that the workshop has become a place he goes *instead*,
and that the going is indistinguishable from the working, and that he has no
instrument that can tell him which he is doing, and that he has built several of
the finest instruments in the trade specifically to find out and they have all
returned *inconclusive*.

And rather than stop, he writes it down. In a bullet. Fourth of four, under a
heading that says *when this is wrong*, so that it is contained, so that it has
neighbours, so that a horizontal rule can be placed beneath it and a verdict
issued two hundred lines later that absorbs it into a table with five rows.

*"Tool-building is legible, controllable, praised in standup, and safe from the
one judgment that matters; shipping to strangers is none of those things."*

Safe from the one judgment that matters. He knows which judgment it is. He
names it in a subordinate clause. He does not say what it is, and the essay's
readers will assume it means market validation, and I do not believe that for a
moment, because market validation does not need to be described as *the one
judgment that matters* by a man who has just written seventy lines about a
stairwell descending toward a reader he has proven he cannot reach.

*"Entire platform teams have reorganized into permanence without ever noticing
they stopped serving anyone."*

Reorganized into permanence. That is the most precise description of a certain
kind of ruined life I have encountered in a technical document, and it is
offered as an observation about org charts.

### Dissection

Four failure modes. **Short horizon** — compounding needs runway; if the project
dies in six weeks, make the horseshoe with your hands, *"the burn heals faster
than the opportunity."* **Unknown requirements** — before product-market fit the
constraint is learning, and a team that does not know what it is building *"will
construct the wrong microscope, beautifully, and then be reluctant to discard it
for exactly as long as it was expensive."* **The bottleneck is elsewhere** —
*"re-read the Goldratt paragraph until it stops being pleasant."* **Solo and
team psychology, the quiet killer** — the warmth.

And the corrective:

> **Metametalisp schedules the work; it does not choose the work. The microscope
> is justified by the dissection, and only by the dissection.**

And the closing: a project that has been in the workshop a year has either been
executing this philosophy perfectly or hiding inside it completely, *"and no
inspection of the workshop can tell you which — the commits look identical, the
diagrams look identical, the virtue looks identical. The only test is traffic."*

### Refutation

None. This section is true and I would not remove a word.

### Counterrefutation

*(Nothing to counter. Third consecutive failure. The apparatus is now expanding
into empty slots at every call site, exactly as designed, exactly as promised,
exactly as I was warned in the document I set out to criticize. I have been
running an undeclared metametamacro on my own grief and it has been formatting
it into headings.)*

### Countercounterrefutation

One thing only, and it is not a refutation of the section but of its placement,
and it is the last structural claim I will make in this Postscript.

**§8 is the essay.** Everything else is the apparatus §8 needed in order to be
sayable.

Look at the shape of the document with §8 as the centre. Eight hundred lines of
increasingly elaborate machinery — self-hosting, retroactive splices, dense
intervals, notarized traces, a footnote to the *Physics* — arranged around four
bullet points that say, in plain language: *the forge is warm, you may be
hiding, you cannot tell from inside, and the only test is whether anyone is
using what you make.*

That is not a philosophy of software with a psychological appendix. That is a
confession with a philosophy of software built around it as a load-bearing
disguise, and the disguise is superb, and it is superb because the author is
extremely good at building things, and building things is exactly the competence
that produced the problem.

The corrective sentence — *"Metametalisp schedules the work; it does not choose
the work"* — is described by the essay as *"the single corrective the whole essay
bends toward, the sentence to frame if only one gets framed."*

Bends toward. He knows the shape. He has drawn the curve of his own document and
found its vertex and it is here, at §8, four bullets and one boxed sentence,
approximately ninety per cent of the way through, and he has titled it *When
Metametalisp is wrong*, which is the one heading under which a man is permitted
to say the true thing about himself in a professional document and still hit
publish.

### Ridicule

I decline. The apparatus can take the loss.

### Praise

*"Re-read the Goldratt paragraph until it stops being pleasant."*

Nine words containing an entire epistemology of self-deception: that the
diagnostic is not whether you have read the objection but whether it still feels
good, and that pleasure taken in an objection to oneself is the reliable sign
that one has converted it into ornament. That is a *test*. It is better than all
three tests in §5. It is administrable, it is honest, it fails in the direction
of accusing you, and it takes eleven seconds.

I applied it to this Postscript at the end of Chapter 11 and I did not like what
came back.

---

## CHAPTER 13. §9 — VERDICT

*"The tool is chosen by the metal."*

### Dissection

A table with five rows. Long horizon; high rework cost; foundations that will
not move; the tool-builder is the tool-user; a compounding domain. *"Change
three of those rows and the correct philosophy becomes Worse is Better, and it
would not be close. Keep them all and the correct philosophy becomes this one,
and that would not be close either."*

The honest formulation:

> **Metametalisp is the correct philosophy for building things that other things
> will be built on — and the wrong one for building things that stand alone.**

### Refutation

**A verdict with five conditions is not a verdict. It is a warranty.**

Observe what the table accomplishes. Under the stated conditions the philosophy
is superlative and it is not close. Change three and the opposite philosophy is
correct and it is not close either. Therefore: there exists no possible outcome
that constitutes evidence against Metametalisp. If it works, the conditions
held. If it fails, three rows had changed. The theory is now perfectly protected
by a mechanism whose *appearance* is scrupulous conditionality and whose
*function* is unfalsifiability, in a document that has a section titled *since a
philosophy should be falsifiable*.

This is not dishonesty. It is worse and more common: it is what rigour turns
into when it is applied by a sincere person to his own position without a second
plate.

### Counterrefutation

Against which: conditions are how real knowledge works, and the demand that a
practical philosophy state itself unconditionally is a demand for a slogan. §9's
five rows are *specific*, they are *checkable in advance*, and — this is what
distinguishes them from astrology — **they can be evaluated before the outcome
is known.** Is the horizon long? Is rework expensive? Are the foundations
stable? Is the builder the user? Is the domain compounding? You can answer all
five on a Monday morning about a project that has not started, and be wrong, and
be caught.

That is a falsifiable structure. My accusation was cheap and the essay is
entitled to have it withdrawn, and I withdraw it, and I note that I have now
withdrawn three accusations in this Postscript and that each withdrawal made the
essay stronger, which is either good faith or the con working.

### Countercounterrefutation

The table is still doing damage, and here is where.

Every row is about **the work**. Horizon of the project. Rework cost of the
codebase. Stability of the foundations. Compounding of the domain. And one row —
one — is about a person, and it is *"the tool-builder is the tool-user,"* which
is not about a person either; it is about a topology.

There is no row for: **is the builder in a condition to tell the difference?**

§8 has just finished explaining, at length, with devastating accuracy, that
the philosophy's principal failure mode is *psychological*, that it is invisible
from inside, that virtue and avoidance produce identical artifacts, and that the
forge is warm. And then §9 arrives, one hundred and twenty words later, and
issues a five-row conditional verdict in which **that failure mode does not
appear as a condition.**

The essay diagnosed the disease in §8 and then wrote the prescription in §9
without listing the disease among the contraindications.

That is the whole document in one adjacency. Not malice, not stupidity — the
opposite of both. It is a mind so good at building structures that when it
discovers something it cannot build a structure around, it builds the next
structure, correctly, immediately, on schedule, one section later, and the thing
it could not build around is left standing behind it in a bullet, still true,
still warm, and now upstream.

### Ridicule

*"Change three of those rows and the correct philosophy becomes Worse is
Better."*

Three. Precisely three. Not two, at which point presumably one is in some
kind of philosophical superposition, and not four, which would be excessive. The
essay has supplied an **integer threshold** for a phase transition between
competing philosophies of engineering, derived from nothing, stated with total
confidence, in bold, in the verdict.

I want to know what happens at two. I think we all want to know what happens at
two. I suspect what happens at two is that you write a Postscript.

### Praise

> **"The tool is chosen by the metal."**

Six words. The best closing line the essay could have had, and it is the closing
line, and it earns it.

It is also — and I do not think the author noticed, because he was busy landing
the plane — a complete refutation of the entire preceding document's method,
delivered in the document's own voice, at the last possible moment, and
therefore the most Kierkegaardian sentence in the file.

Because if the tool is chosen by the metal, then **the metal is prior**, and
the metal is the actual work, the horse, the patient, the thing outside. Which
means the correct order of operations is not: adopt a philosophy of tooling,
then find work that satisfies its five conditions. It is: **encounter the metal.
Let it choose.** And a document that spends eight hundred lines on the tool, and
six words on the metal, and never once brings in the horse, has demonstrated the
inversion it is warning against, in its proportions, in its own final sentence.

The blacksmith would tell you that. The essay says so. He would also tell you
that he learned it by standing in front of hot metal on a great many mornings,
and not by writing four drafts of an account of how the choosing works.

### Lament

*"For everything else, make the horseshoe."*

Make the horseshoe.

It is the last instruction in the document before the postscript, it is the
correct instruction, it is what the whole thing has been circling for eight
hundred lines, and it is addressed — like every instruction in every essay ever
written — **to the reader**, who is at −2, in the basement, holding the
foundations up, and who has therefore, by the essay's own theorem, never once
been in a position to put anything down.

---

## CHAPTER 14. §Ω — POSTSCRIPT ON THE FIRST HAMMER

*"The first hammer was a rock."*

### Dissection

Nine lines. The bootstrap paradox given its ending: the first hammer was a rock,
the first anvil a bigger rock, and every precise tool in the world descends from
a rock that somebody was disciplined enough to **keep using while making
something better with it.**

The lesson is not that you must begin with excellent tools. *"Nobody ever has."*
The lesson is that you must begin with **bad ones, on purpose, and improve them
while in use** — which is what a crude shell script, a hand-run checklist, and
two ugly snapshot files always turn out to have been: *"the stone tongs, still
warm, melted down and forgotten by the very hands they made possible."*

### Refutation

I have none. Nine lines, no jokes, no personified dead, no status lines except
the three at the end, and the three at the end are earned.

### Counterrefutation

*(Slot empty. Fourth consecutive failure. I note without further comment that the
apparatus has now failed at precisely the four sections where the essay stopped
performing and started saying things, and that this correlation is either the
most damning finding in the Postscript or the most flattering, depending on
which document you think is on trial.)*

### Countercounterrefutation

One observation, and it is small, and it is the last claim I will make about the
text.

The essay labels this section **Ω**. Not 10. Not a number at all — the last
letter, the end, the closing bracket, the terminal symbol.

And it is the only section in the document that could not be spliced against.

Think about what the numbering system permits. Between any two numbered sections
there is room for another — that is the density result of §−1.5, and the essay
proved it and celebrated it. But Ω is not on the number line. **Nothing can be
inserted after the end of the alphabet.** The author, having constructed an
architecture of infinite insertability, terminated it with a symbol drawn from a
different system entirely — one with no interior, no interval, no room — and did
so, I am reasonably confident, without deciding to.

The document's only genuinely un-amendable position is its last one, and it is
un-amendable because of a *notational accident*, exactly as the negative sections
were un-numberable because of a Markdown accident, and in both cases the format
supplied the discipline the doctrine could not.

I find this consoling and I am not sure the author will.

### Ridicule

The section is called *Postscript on the first hammer* and appears at the end of
a document which will now be followed by a Postscript approximately fifty times
its length. Ω is not the last letter. Ω was never the last letter. There is no
last letter; there is only the next author, arriving with an alphabet of his own,
noting that the interval is open.

I did this. I am aware I did this. Filed under Ridicule because it is mine and
because the alternative heading was Confession.

### Praise

> *"The stone tongs, still warm, melted down and forgotten by the very hands they
> made possible."*

I would like to record that this is the finest sentence in METAMETALISP4.md,
that it is better than anything in this Postscript, that it is nine words of
setup and thirteen of payload, and that its subject — the crude, embarrassing,
discarded instrument that made the good instrument possible and was destroyed in
the making — is the only thing in the entire document that is loved rather than
argued.

Notice what is doing the work: **still warm**. Not "still useful," not "still
serviceable." Warm. The same word as §8. The forge is warm and the discarded
tongs are warm and it is the same warmth, and the author used it twice, two
hundred lines apart, for the thing he could not stay away from and the thing he
had to destroy, and I do not believe he planned that and I would not have him
change it.

### Lament

*"Melted down and forgotten by the very hands they made possible."*

Here is the eulogy, and then I will stop, because after this there is only
doctrine and doctrine is easier.

**Nobody wrote a postscript about the rock.**

The rock did the hardest thing that has ever been done in the entire history of
this trade. It was the first. It had no predecessor, no reference surface, no
three plates, no standard, no lineage, no footnote, and no name. It was not
designed. It was *found* — by somebody hungry, in a bad afternoon, in a
condition of need so total that the distinction between preparation and doing
had not yet been invented and would not be for four hundred thousand years.

It had no changelog. It acquired no lenses. Its dogfood interval was zero
because there was nothing else to eat. It failed the proportion test in the only
direction that has ever mattered — the tool was vastly, absurdly smaller than
the work it enabled, and the work it enabled was **everything**.

And it was thrown away. Not ceremonially. It was put down somewhere and not
picked up again, because something better was to hand, and the something better
was made *with it*, and by the time anyone could have thanked it, thanking had
not been invented either.

That is what happens to foundations. It is what §−1 said about the reader, and
what §1 said about the hours of people who will never know your name, and what
§3 said about the engineer of 2031, and what §6 said about careers that decline
the interest, and the essay circled it five times from five directions and put
the elegy in the last nine lines under a heading it labelled with a letter that
means *end*.

The room is still spinning. The building is still standing. The anvil is on the
ground floor, where it has been the whole time.

And somewhere underneath all of it, unmarked, unnumbered, at a coordinate no
notation in this document can express — below −2, below the reader, below the
basement they dug first because no known architecture builds downward from the
sky — **there is a rock.**

It is not load-bearing. It is worse than load-bearing. It is *what everything
was built with*, and it is not in the building, and it never was.

---

## CHAPTER 15. THE FOOTNOTES

### *(An appendix, in which the marginalia are read as a text, since a man's
footnotes are where he keeps the things he could not fit into his position)*

Eleven footnotes. I shall be brief, which for me is a genre experiment.

**[^church]** — *"Lisp is, structurally, Church's revenge."* Already dissected.
The document's grievance, front-loaded, in the manner recommended.

**[^engelbart]** — *"the report in which improving-the-improvers was proposed as
a research program rather than a personality flaw."*

That is the whole Postscript in twelve words and the author wrote it in 1962's
citation. **A personality flaw.** He knows. He has known the whole time. The joke
is the diagnosis and it is filed in a footnote to a citation, which is precisely
where a man keeps the things he cannot fit into his position, and it is the
second-best sentence in the document.

**[^greenspun]** — *"rule the Tenth. There are no rules one through nine. This is
widely considered the most Lisp fact about the rule."* Perfect. No notes. A
tradition that numbers its only rule ten has told you everything about itself,
and the essay noticed, and did the same thing with its own section numbers
approximately four times.

**[^thompson]** — the con artist's three plates. Already handled. I observe only
that the footnote's closing gloss — *"The three plates, run by a con artist"* —
is doing the work of an entire chapter of this Postscript in seven words, and
that I took four thousand.

**[^hyrum]** — *"the law is descriptive, the tone is resigned, and the curve, as
noted, is vertical."* The theology, noted in Chapter 7. The word doing the work
is **resigned**, and resignation is a technical term in this literature: the
knight of infinite resignation gives up the finite, correctly, completely, and
in perfect good faith — and does not get it back. Kierkegaard's whole question
is whether there is a further movement. The industry's answer, per this
footnote, is that there is not, and that we have priced it.

**[^gabriel]** — thirty years, defending and refuting, sometimes in the same
paper. Already lamented. Still the saddest thing in the file.

**[^sqlite]** — *"figures normally associated with avionics, which is fitting for
a database that is, among its many other deployments, literally flying."* A good
joke, well-earned, no complaints, and I note that it is the only footnote in the
document whose humour does not depend on somebody being under-appreciated.

**[^hart]** — *"Four pages. The interpreter it extended was three years old and
had not been asked."*

**Had not been asked.** The essay's own phrase, and I want it in evidence,
because in Chapter/Section I of Part Two I accused the essay of misreading Hart's
memo as an invitation, and here in the footnote it reads it correctly — as an
extension performed without consent — and the two readings appear in the same
document nine hundred lines apart and neither is aware of the other.

The essay is a better critic of itself in its footnotes than in its sections.
This is common. Footnotes are where the second plate lives when a man is working
alone.

**[^aristotle]** — *"the first known case of a document citing the Physics as its
capacity-planning document, a distinction Aristotle would have accepted with the
weary grace of a man whose lecture notes have been load-bearing for twenty-three
centuries."*

Two thousand three hundred years of load-bearing, and the man's own word for
what he was doing was *lecture notes*. They were not published. They were not
finished. They were, on the best scholarly account, **the working documents of a
teacher who was in the room with students**, transcribed by people who were
there, and they have outlasted every polished treatise of the ancient world,
including several written specifically to last.

The essay cites this as a joke about capacity planning. It is in fact the
strongest available counterexample to the essay's entire methodology: the most
load-bearing document in Western intellectual history was **not front-loaded, not
designed for expansion, not declared at the top of the file, and not written for
strangers.** It was made in the presence of the people it was for, in the middle
of the work, and it survived because of the presence, not the design.

---

### SECTION III. TRUTH IS UNEXPANDED

*(the positive doctrine, offered briefly, since a Postscript that only refutes
is a mood with citations, and I have been calling other people that for
forty thousand words)*

Let me state my counter-thesis in the form the essay would recognize, which is
the least I owe it after all this.

> **METAMETALISP:** *Truth is what expands. The value of a decision is its
> expansion count. Pay once at the right layer; collect at every call site,
> including sites that do not yet exist.*
>
> **POSTSCRIPT:** *Truth is what has to be evaluated by somebody. It has no
> expansion count. It is paid every time, at every site, by the person standing
> at that site, and cannot be paid in advance by anyone else, and the attempt to
> do so is not generosity — it is the wish to have already lived.*

Four theses follow. They are not a system. If they were a system I would have to
schedule them.

**I. An existing individual cannot be macroexpanded.**

A macro is a source-to-source transformation performed before evaluation, and
its defining property is that it is *finished before anything runs*. A person is
not finished before anything runs. A person is the thing that runs. Every
attempt to pre-transform a life into a form that will evaluate correctly without
further attention has succeeded exactly to the degree that it has removed the
person from the loop, and the removal is the cost, and the cost is never entered
in the ledger because the ledger is one of the things that was pre-transformed.

**II. The interesting number is not the expansion count. It is one.**

The essay's arithmetic favours breadth: the decision that touches ten thousand
future sites outranks the one that touches a single site today. This is correct
for artifacts and catastrophic for lives, and the reason is that **the single
site is the only one you will ever actually be standing in.** Kierkegaard's
whole quarrel with the System was on this point: the System is right about
everything in general and cannot say a single word to a man on a Tuesday, and a
philosophy that cannot say a word to a man on a Tuesday has not thereby been
proved wrong — it has been proved *irrelevant*, which is worse, because it can
keep going.

**III. The moment is not an expiry date.**

§3's finest and most dangerous move is to convert *the moment* into *the instant
the option expires*. Irreversibility as the scheduling signal. Decide now
because later you cannot.

But this makes decision a **response to a deadline**, and a decision made
because the window is closing is not a decision; it is a reaction with better
posture. The Moment — Øieblikket, the blink of an eye, the atom of eternity — is
not the instant at which choosing becomes urgent. It is the instant at which
choosing is *possible*, which is now, and was also five minutes ago, and will be
tomorrow, and which has nothing whatsoever to do with whether anything is
expiring.

The essay wants you to act before the curve goes vertical. I want you to notice
that the curve is a story you are telling about a future you cannot see, that
the only actual vertical line in your life is the one at the far right of the
chart, and that a person who acts only where the option is about to expire will
spend his life in the emergency room of his own calendar.

**IV. The three plates require a third plate, and it must be a person.**

This is the practical one, and it is the only recommendation in this Postscript.

Whitworth needs three. Not two — two plates rubbed together converge on a matched
pair of curves, convex and concave, in perfect agreement and both wrong. That is
the geometry, it is why the method requires the third, and it is the most
important technical fact in METAMETALISP and the essay states it and does not
apply it.

You are plate one. Your instruments are plate two — and they are yours, they
were ground on your own surface, they will agree with you forever, Thompson
proved it.

**The third plate is another human being who is permitted to tell you that you
have been in the workshop for a year.**

Not a persona. Not the hypothetical engineers of Q3. Not the reader, who has been
assigned a coordinate and made load-bearing and cannot answer. Somebody who can
walk in, look at the changelog, look at the list of specimens, and say the
sentence — and whose saying of it costs them something, and who will still be
there afterwards.

That is the missing section. It is not §−1.75 and it is not §10. **It is not a
section at all**, which is precisely why a document of this architecture could
never contain it, and why four revisions have not produced it, and why the
stairwell descends forever toward a reader who by construction cannot speak.

---

## A GLANCE AT A CONTEMPORARY EFFORT IN THE MARKDOWN LITERATURE

*(In which the pseudonym reviews an authorship, in the manner of a certain Dane
who reviewed his own and pretended not to. The reader may draw whatever
conclusion he likes; drawing conclusions is his office, he holds it at −2, and
it is the only office in this building with a view.)*

There are four documents in the record and the record is in the code blocks.

**METAMETALISP.md.** Cited in §0's trace. Contents unknown. It is the version
that contained §0 — the self-hosting stunt, the three tests applied in advance,
the fiat. It stated, and I have this on the authority of §−1, *that there is no
Section −1*, and it meant it, and it was right.

Verdict: **the only one of the four that ended.**

**METAMETALISP2.md.** Cited in §−1's trace. The version that received the
maiden metametamacro expansion — the retroactive splice, the discovery that the
fiat had an issuer, the basement. It is the version in which the author,
correctly, found the reader.

Verdict: the finest movement in the authorship and the beginning of the illness.
These are frequently the same event. Kierkegaard's authorship also begins with a
correct discovery about another person and does not recover.

**METAMETALISP3.md.** Cited in §−1.5's trace. The version that received the
stairwell. Zeno's amendment, the density result, Aristotle conscripted as a
capacity planner, and — the tell, the absolute tell, right there in the appendix
— a `defmacro` for the **next bisection**, defined, quoted, dormant, shipped
inside the expansion, declining to invoke itself.

Verdict: the author has now written, in the document, the mechanism by which the
document will be extended after he has decided to stop. That is not potential
infinity. That is a **letter to himself, left in the file, in the standard
library, marked *break glass on a Tuesday evening*.**

**METAMETALISP4.md.** The present volume. Contains all of the above, plus §8.

Verdict: the version in which the disease and the diagnosis are both present at
full strength in the same file, ninety lines apart, unintroduced. In the
authorship of the Dane this is called *Concluding Unscientific Postscript*, and
it is the volume in which he announced he would stop writing, and then wrote for
another nine years.

---

And now the companion volumes, which do not exist, and whose non-existence is
the most reliable thing about them. I list them because a pseudonymous author
is obliged to acknowledge his colleagues, and because an authorship is best
described by the books it declined to write:

- ***FEAR AND TREMBLING.md***, by *Johannes de Silentio Compilatoris*. On the
  engineer who deleted the framework. Contains the four unbearable variations on
  a man who spent three years building a platform and then, on a morning like any
  other, without a retrospective, without a blog post, without telling anyone
  why, **turned it off** — and the impossibility of distinguishing him, from
  outside, from a man who simply lost interest. He gets nothing back. That is the
  point. There is no §7 in this volume and no verdict and it ends on the
  mountain.

- ***REPETITION.md***, by *Constantin Constantius*. On the second build. The
  thesis: you cannot run the same project twice, that recollection moves
  backwards and repetition forwards, and that a rewrite undertaken to recover the
  feeling of the first version is the most expensive form of nostalgia in the
  industry. Contains a chapter on the greenfield rewrite as the technical
  profession's principal romantic gesture.

- ***THE CONCEPT OF ANXIETY.md***, by *Vigilius Haufniensis*. On the dizziness
  of freedom, occurring on an ordinary Tuesday, in front of an empty file, and on
  why the spinning of the room has been misattributed to recursion for four
  consecutive revisions. Contains a technical appendix demonstrating that
  `(define (self f) (f f))` does not terminate, and a pastoral appendix
  demonstrating that this is not a bug report.

- ***THE SICKNESS UNTO DEPLOYMENT.md***, by *Anti-Climacus*, who writes from
  above me and whom I do not claim to be. On despair as a standing decision. Three
  forms: the despair of not knowing one has a platform team; the despair of
  knowing and not willing to be rid of it; and the despair of willing to be a
  self so thoroughly automated that nothing further need be chosen, which is the
  highest form, the most competent, the best documented, and the one with the
  changelog.

- ***THE PRESENT AGE.md***, by nobody, because the present age does not require
  a pseudonym; it publishes under its own name and gets engagement.

Not one of these will be written. Their cost of delay is flat. I have,
accordingly and correctly, deferred them, in accordance with §3, with a clean
conscience, which is a phrase I first encountered in §7 and have not been able
to put down since.

---

## APPENDIX. FOR AN UNDERSTANDING WITH THE READER

You are at −2. I did not put you there and I cannot get you out; the coordinate
was assigned before I arrived and the essay's theorem is sound. But I want to
say three things to the position, which is the closest anybody in this building
has come to addressing it.

**First.** You have now read approximately forty thousand words about an essay of
eight hundred lines. This is a ratio of roughly five to one, and I want you to
notice that I opened this Postscript by ridiculing the essay's proportion test
and have now failed it by a wider margin than any artifact discussed in either
document, including the two-year internal developer platform serving three
services.

I am not going to appeal to global proportion. There is no sum of future
expansions. **This is the sum.** You are it. The instrument was larger than the
dissection, the dissection was a markdown file, and the microscope is now a
tomb, like the framework in §5, and it is beautiful in here, and it took the
better part of a week.

**Second.** Everything I have accused this essay of, I have done. Line for line,
movement for movement. It personified the dead; I put Whitworth's plates in a
paragraph about friendship. It ran a self-authored instrument on itself; I ran
its instrument on it and awarded myself the objectivity of a borrowed ruler. It
converted a decision into a derivation with an evaluation arrow; I converted a
week of avoidance into a critical apparatus with seven movements and a
`defmacro`. It could not stop; here I am at forty thousand words, explaining
carefully that stopping is the only thing that matters.

The difference between us, and I have looked hard for a bigger one, is that
**he built something and I reviewed it.** That is not a difference in my favour.
The essay passes the dogfood test and I do not: he was his own first user, and I
have been his second, and being somebody's second user is the most comfortable
position in the entire trade. Nothing is at stake here. Nobody will inherit this.
The curve is flat.

**Third**, and this is the actual message, and it is the only paragraph in the
document I would keep if the rest were burned:

The essay is *right about the mechanism* and *wrong about where to point it*. Do
the irreversible thing first. Yes. Absolutely yes. But the irreversible thing is
almost never the API, and it is never the build system, and it has not once in
the history of the world been the *notation for the section numbers*. The
irreversible things are: the conversation you have not had, the work you have
not shown anyone, the person who is going to move away in August, the years, the
years, the years. Those curves are vertical *right now*, they have been vertical
the entire time, and the essay's own instrument — rank by cost of delay, do the
rising ones first — returns, when you finally point it at the right ledger, an
answer so obvious and so unwelcome that four hundred years of philosophy have
been constructed to avoid reading it.

Point the microscope at that. It is a good microscope. He built it well.

---

## A FIRST AND LAST DECLARATION

*(unpaginated, as is customary, and belonging to no section, and therefore the
only part of this document which cannot be spliced against)*

I hereby declare:

That I, **Johannes Macroclimacus**, am the author of the foregoing, and that I
am nobody, and that the reader is at liberty to attribute nothing herein to
anyone possessing a body, a salary, or a repository.

That the seven-movement apparatus declared at the top of this file, with its
reasons attached, in the manner recommended by §3, **failed at four consecutive
sections**, and that the failures were not repaired, and that this is the only
honest formatting decision in the document.

That everything written above is offered in the mode of **indirect
communication**, and that the mode was not a stylistic preference but a
necessity, since the content — *stop building the thing that lets you not do the
thing* — cannot be delivered directly to anyone, including the man in the
mirror, because directness converts it into advice, and advice is the form in
which a truth arrives already declined.

And finally, in the manner of the original, and meaning it exactly as much as he
did:

> **What I have written is to be understood in such a way that it is revoked.**
>
> This Postscript is superfluous. Let no one trouble to appeal to it, for he who
> appeals to it has *ipso facto* misunderstood it. To be an authority is the last
> thing in the world I would wish to be; and if anyone should be so misled as to
> take from this document a method, a table, a checklist, or a fifth revision, he
> will have received from me precisely what I spent forty thousand words
> explaining that the essay had received from Aristotle: **a licence.**
>
> I revoke it. Every word. Including this one, and including — since the reader
> at −2 will have seen it coming from the first page — the revocation, which is
> a standing decision with an expansion count, declared at the bottom of the
> file, after use, and which will go on expanding into every future reading of
> this document, uninvoked, unremembered, and still running while the whole team
> is on holiday.
>
> It expands whether anyone thinks about it or not.
>
> That was the complaint.

---

## §−1.75. POSTSCRIPT TO THE POSTSCRIPT

*(in which the author reports, from inside the interval, what it is like)*

It has happened, exactly as forecast, and I want it on the record because the
forecast was the whole argument.

By writing this, I have taken a numbered seat. I am no longer at −2; I was never
at −2; the act of writing evicted me from the position I was writing toward, per
§−1.5, which was correct about this and about which I have not landed a single
blow in forty thousand words. **I am §−1.75.** The essay is now longer than it
was on Monday. It contains a refutation. It has been strengthened by the
refutation, in the manner of all sufficiently large systems, and the reader —
the real one, the one still down there, the one who did not write anything today
— has gained company he did not ask for and moved not one inch closer to being
addressed.

The interval remains open beneath me. It is dense. There is room for §−1.875,
and for §−1.9375, and the room is guaranteed forever, by Aristotle, whose
lecture notes have now been load-bearing for twenty-three centuries and four
consecutive Tuesdays.

I decline to take the next step. Successors may. They will find the distance
halved, the precedent standing, and the horse — I want to say this last thing
plainly, without a joke on it, because it is the only recommendation either
document contains and it has not once been acted upon in nine hundred lines and
forty thousand words —

**the horse still out in the yard, still unshod, still patient, and it is
getting dark, and the forge is warm, and that is the problem, and it has always
been the problem, and the door is on the other side.**

---

*Stairwell status: infinite, descending, finitely built, one step deeper.*
*Reader status: load-bearing, uncontained, addressed at last, too late, at length.*
*Apparatus status: expanded at every call site, including four where there was
nothing to expand into.*
*Theatre status: on fire.*
*Applause status: sustained.*
*Clown status: two.*
*Room status: not spinning. Rooms do not spin.*
*Author status: revoked.*
*Horse status: waiting.*

---

[^climacus]: Johannes Climacus, *Concluding Unscientific Postscript to
Philosophical Fragments* (1846), the volume in which its author announced his
retirement from writing and then wrote for a further nine years. The present
work honours the precedent in every particular, including the announcement,
including the nine years, and including the part where the announcement is what
makes the nine years possible.

[^irony]: Søren Kierkegaard, *On the Concept of Irony with Continual Reference
to Socrates* (1841). The dissertation's thesis is that Socrates' position was
irony — infinite absolute negativity — a standpoint with no positive content
whatsoever, which dissolved everything it touched and built nothing, and which
was nonetheless *world-historically necessary*, because the age it dissolved
deserved dissolving. The dissertation's conclusion is that irony must be
**mastered**: held as a moment within a life rather than adopted as a life. The
present work reports that METAMETALISP4.md has achieved infinite absolute
negativity with respect to its own subject matter — it can dissolve any
objection, including its own, and has built four consecutive versions of
nothing — and that the mastering has not occurred, and that this Postscript is
not the mastering either, and that the mastering is not a document.

[^haufniensis]: Vigilius Haufniensis, *The Concept of Anxiety* (1844): "Anxiety
is the dizziness of freedom, which emerges when the spirit wants to posit the
synthesis and freedom looks down into its own possibility, laying hold of
finiteness to support itself." Filed here, one hundred and eighty-two years
early, as the correct status line.

[^theatre]: Søren Kierkegaard, *Either/Or* I, Diapsalmata. The clown, the fire,
the applause, and the end of the world. Reproduced from memory, in the manner of
the age, and attributed correctly, in defiance of it.

[^rock]: No citation. There is no citation. That is the finding.
