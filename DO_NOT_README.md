# DO NOT READ ME

*Why this project should not exist. A preface, in place of the book.*

> Fac cellulam A1 bold.
>
> — `scripts/polyglotta/latin.vla`, a sentence in Latin containing one
> English adjective it could not decline

---

## On the title

The title of this file is an imperative sentence, and you have already
disobeyed it. That is the first thing worth knowing about imperative
sentences: they are not commands until somebody complies, and nobody
has to. The grammar books call this *mood*. It is well named. A
sentence in the imperative mood is in a mood, and the reader may or
may not indulge it.

This repository is a machine for imperative sentences. `Make cell A1
bold.` `Put today into cell D1.` `Repeat 5 times.` Each is checked,
each means exactly one thing, and each is obeyed or refused in words.
The README next door explains how. This file explains why the whole
enterprise is impossible, in the hope that you will find that as funny
as we do, and then go read the README anyway.

## First objection: a sentence has one meaning

It does not. Not one sentence in any language spoken by people has ever
had exactly one meaning, and the spreadsheet vocabulary is a worse
offender than most. Consider what English brought to the table when it
was asked to name the parts of a grid of numbers.

A **cell** is where a monk prays, where a prisoner waits, where a body
divides, and where a telephone lives. A **range** is what cattle roam.
A **sheet** is for a bed or a ghost. A **table** is for dinner, and a
**column** held up a temple before it held up a total. **Row** is what
you do to a boat or what you have with your brother. **Bold** is what
one is in the face of danger. **Workbook** is the only word in the set
that means nothing outside a spreadsheet, and it is the only one nobody
uses.

So when a person writes `Make cell A1 bold.`, English has offered, in
good faith, a sentence that could reasonably be asking a small prison
to show some courage. The project's answer is a grammar in which every
pattern matches or does not, first match wins, and nothing backtracks.
It resolves ambiguity the way a border does: not by understanding both
sides but by admitting one. This is either a betrayal of language or
the only respectful thing anyone has done to it in years. The file
does not take a position. The file is not supposed to be read.

## Second objection: translation

The project keeps seven demonstration phrasebooks beside the English
one, in Spanish, German, French, Danish, Latin, Esperanto, and pirate.
Every one of them proves the same two things: that the seam works, and
that translation does not.

Take the pirate. `Brand cell A1 bold, arr.` A pirate, it turns out, is
bilingual by necessity, because the word *bold* binds straight into
the name of the machinery underneath and cannot be made to say *arr*.
The piracy lives in the carrier phrase. The cargo is English. This is
the condition of every loanword ever borrowed: the sentence goes native
and one word inside it never does.

Take the Latin. The author of `latin.vla` needed a word for *button*
and discovered that Rome had no interface elements. The file settles on
*bulla*, a knob, and says so in a header written entirely in Latin, in
a note that includes the phrase *anachronismi sine origine Latina,
aperte scripti, non aliter simulati*: anachronisms without a Latin
origin, openly written, not otherwise pretended. A dead language,
apologizing in itself for having missed the invention of the mouse.

Take the Danish. `Saet dags dato i cellen D1.` There is no *æ* in that
sentence because the dialect files drop their diacritics for the
comfort of the machine. The Danish has been de-Danished in order to be
read by a computer, which is the kind of favor languages have been
doing each other since the first scribe ran out of room for the vowels.

And take the alien, `alien.vla`, a phrasebook with no words in it.
Every operation is a glyph: `=|>` inscribes, `*!*` emboldens, `|~~~|`
wraps a transmission. It is the purest translation in the repository,
because there is nothing left to mistranslate, and it still fails at
exactly one point. The one name that must survive into the generated
VBA cannot wear a glyph, because VBA identifiers are made of letters,
so it is spelled in leetspeak, `x3n0-h41l0`, which the code generator
then quietly rewrites as `x3n0_h41l0` because it also cannot wear a
hyphen. Every translation has one word that would not cross, and when
you finally get it across, the customs officer changes the spelling.

## Third objection: today

`Put today into cell D1.` A tidy sentence. It contains the least tidy
word in any language, a word that points at the moment of its own
utterance and is wrong by the time it is written down. Linguists call
such words *deictic*, from the Greek for pointing, and have the good
grace to look embarrassed when they do.

The project has a rule for *today* in every one of its phrasebooks.
For a while, in all seven of the non-English ones, that rule could not
be reached. A more general rule, one that accepts any expression at
all, sat earlier in the file and caught the sentence first, and first
match wins, and nothing backtracks. In English, *today* could be said.
In Spanish, German, French, Danish, Latin, Esperanto, and pirate, it
could not, and nobody noticed for the whole life of the files, because
nobody had asked. The bug was found by wiring one of them into a test,
which is how most things about a language are found: not by speaking
it but by finally listening to someone else try.

The roadmap records this without drawing a moral. Neither will this
file.

## Fourth objection: the floor beneath the floor

Excel is already a translated language. Nobody thinks about this. In a
Spanish workbook `=SUM` is `=SUMA`; in German it is `=SUMME`; in French
it is `=SOMME`; across most of Europe the commas between arguments
become semicolons, and `TRUE` becomes something else in every country.
It is the only programming language in wide use whose keywords change
at the border, and a quarter of a billion people write it every day
without once calling it a programming language.

Frazaro puts a controlled English on top of that. The English is
compiled into a middle layer of parenthesized VBA, which is itself a
dialect of a dialect of BASIC. So a sentence like

```text
Chart the course "=B2*2" onto the cell B3, arr.
```

is four languages deep before it touches a number: pirate around the
outside, a quoted formula in Excel's own tongue, a cell address in a
notation borrowed from chess, and underneath, waiting, a lisp-shaped
VBA that will be rewritten once more before it runs. The structure is
tall now. Tall structures built by people who share a single language
have a well-documented failure mode, and this file, in keeping with
its title, will not go into it.

## Fifth objection: the name

*Frazaro* is Esperanto for *phrasebook*. Esperanto was invented so
that phrasebooks would be unnecessary. The project is named, in the
language of the cure, after the symptom, and the naming was done on
purpose, and it is the most honest thing about it.

## Sixth objection: a phrasebook is a confession

A tourist's phrasebook admits, on every page, that its owner does not
speak the language. It teaches you to ask where the railway station is
and offers nothing whatever for the reply. It is a one-way instrument,
which is why the tourist ends up pointing.

This phrasebook talks back. A refused sentence is told what was
understood, where expectation left the rails, and what to write
instead. The reply is written by the same grammar that refused you,
which means the project's one working conversation is between a
person and a set of rules, and the rules have the better vocabulary.

The Italians say *traduttore, traditore*: a translator is a traitor.
The project keeps two translators, one that interprets your sentence
in place and one that compiles it into a module you can read, and
holds them to agreement with a corpus of golden files. A disagreement
between them is a failing test rather than a diplomatic incident. It
is the first arrangement in the history of translation where both
traitors are watched by a third party who cannot speak either language
and only checks whether the two of them said the same thing.

## Why it exists anyway

Because everyone who has ever lived in a spreadsheet can describe, in
one breath, a procedure they cannot automate, and the two honest
options offered to them so far are a language that is not theirs and a
prompt box that is confidently wrong just often enough to matter. A
small English that refuses is a third option. It should not exist, in
the sense that every argument above is correct. It exists in the sense
that it runs.

The repository's single permitted reference to a certain melancholy
Dane was spent in the README, on its opening line, and this file will
not spend another. The only Dane admitted here is `dansk.vla`, which
contains twenty sentences, none of them about despair, and one about
today, which for a while could not be said.

Either you now close this file, or you open the README. The grammar
would accept both. It would refuse a sentence that tried to do neither,
and it would tell you why.

---

## Codified coda

*The remarks below were made over dinner by a poet of the wilted
carnation, who once had difficulty at customs and declines to be named,
being already sufficiently quoted. They were made in one language, on
the subject of all the others. The project, being unable to leave a
sentence unnumbered, has numbered them, and, being unable to leave a
sentence unchecked, has checked them. The results are appended.*

**§1.** To translate a thought is to move house with it. Something is
always left in the attic, and it is always the thing one meant.

**§2.** Every language is a conspiracy against the others, and every
speaker an accomplice who believes he is merely talking.

**§3.** A word means what its neighbours allow it to mean. Take it
abroad and it must make new friends, and it will make the wrong ones.

**§4.** The faithful translation and the beautiful one have never been
introduced. Each speaks of the other as one speaks of a cousin in
trade.

**§5.** One does not translate a sentence. One marries it into another
family and hopes that it writes.

**§6.** The dictionary is the only book everyone consults and nobody
believes.

**§7.** To understand a foreigner perfectly is to discover that he did
not mean anything either.

**§8.** A phrasebook is a work of optimism. It assumes the country will
answer in the phrases provided.

**§9.** Grammar is what a language remembers of its quarrels.

**§10.** Fidelity in translation, as in marriage, is admired chiefly
by those who are not party to it.

**§11.** Meaning is the one item that never appears on the customs
form and the only one ever seized.

**§12.** Ambiguity is the courtesy a language extends to its own
speakers. Precision is the discourtesy it extends to everyone else.

**§13.** I have never met a thought that survived a second language,
though I have met several that were improved by dying.

**§14.** The machine refuses a sentence it does not understand. This
places it well ahead of the rest of us, who reply.

**§15.** There are only two kinds of translation: the wrong and the
unread. The second is by far the more faithful.

**§16.** A pun is the one joke that stays at home. It cannot travel,
and it resents those that can.

**§17.** Cognates are false friends who have kept the family name.
They are recognised at once and trusted at one's peril.

**§18.** Every untranslatable word is a compliment a language pays
itself.

**§19.** The interpreter is the only person in the room who knows what
was said, and the only one forbidden to say so.

**§20.** A dead language is merely one that has stopped taking
corrections.

**§21.** An accent is what the other person has.

**§22.** A dialect is a language that lost the war. A language is a
dialect that kept the printing press.

**§23.** To speak two languages is to be misunderstood twice as often,
and to know it.

**§24.** The bilingual does not think in two languages. He thinks in
neither, and is fluent in the gap.

**§25.** A universal language was invented once, so that all men might
misunderstand one another equally. It succeeded with everyone who
learnt it, who were few, and who agreed with each other entirely,
having nothing left to say.

**§26.** Subtitles are the confession that the face was understood and
the words were not.

**§27.** Poetry is what is lost in translation. Prose is what is lost
in the original.

**§28.** Etymology is the study of how a word arrived at its present
respectability, and, like most such histories, is best not read aloud
at dinner.

**§29.** Grammatical gender is proof that a language may be certain of
things it has never examined. The moon is a woman in one country and a
man in the next, and has not been consulted in either.

**§30.** The subjunctive is the mood in which a language admits it
might be wrong. English has very nearly done away with it.

**§31.** Jargon is a dialect spoken in order to prevent translation.
It is the only dialect that has ever succeeded.

**§32.** The lawyer and the poet both write so as to mean one thing.
Only the poet is disappointed when he does.

**§33.** Literal translation is the sincerest form of
misunderstanding.

**§34.** A borrowed word is never returned, and is spelt wrong for the
remainder of its life.

**§35.** What cannot be said in a language is not thereby unthinkable.
It is merely unpopular.

**§36.** Every idiom is a private joke a language tells about its own
history. The translator is the stranger who asks to have it explained.

**§37.** A number is the only word that every language spells the same
and none pronounces alike.

**§38.** A formula is the one sentence in a workbook that means the
same in every country, provided one first changes all the words.

**§39.** Gesture is the mother tongue and the only one without a
dictionary. This is why it is understood.

**§40.** Spelling is a language's memory of how it used to sound.
English remembers a great deal, all of it inaccurately.

**§41.** A machine that translates has read every book ever written
and understood the same number as its critics.

**§42.** To insist that a word has one meaning is to insist that a
coin has one side. The coin will tolerate it. The word will not.

**§43.** A translation is finished when the translator can no longer
bear the original.

**§44.** The most exact rendering of a sentence is the sentence.
Everything after that is commentary, including the sentence.

**§45.** Make cell A1 bold. It is the only courage left to us.

*Check results.* Forty-five remarks were submitted. Forty-four were
refused, each on its own row, in words, the grammar having found in
none of them a cell, a range, a sheet, or an instruction of any kind.
The first sentence of the forty-fifth was accepted, and ran, and cell
A1 is bold to this day. Its second sentence was refused. The poet asked
whether this made him a programmer, and was told it made him half of
one, which he said was the customary proportion. He then asked that the
refusals be printed in place of the remarks. They were not. There was
not room, and they were longer.
