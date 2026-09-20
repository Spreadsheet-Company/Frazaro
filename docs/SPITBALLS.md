# SPITBALLS — fun things Frazaro could do, before anyone has asked

*Filed 2026-09-19 from a conversation, at the owner's request. Having
spent the day on elegance (Contemplation 9), the owner pivoted "entirely
away from elegance" to coolness: what non-business features would make a
developer, a superspreadsheeter or a hacker stop in their tracks and say
"whoa, that's radical"? The owner's two exemplars were already in the
repository: the Bob Ross palette in `examples/joy.txt` ("palettes as
importable header files") and the CLI on `` Ctrl+Shift+` `` ("elevating
the Immediate Window experience into a pseudo-IDE"). The brief was for
features "that aren't even on the roadmap".*

*House rules. A spitball is less than a contemplation and much less than
an item: nothing here is scheduled, priced or promised, and `SD-7` still
governs anything that touches grammar. Spitballs are numbered the way
contemplations are — "Spitball 8" is a handle, not an ID, and
`ID_REGISTRY.md` governs nothing in this file. The file is append-only:
a spitball that gets built, killed or promoted is annotated in place with
a date, never deleted or renumbered. Sizes use the roadmap's notation
(`~hours`, `~days`, `~weeks`) and are the assistant's unmeasured guesses.
Additions are welcome the way any contribution is (`CONTRIBUTING.md`).*

*Checked, not recalled. Every claim below about what exists was read from
the code at `HEAD` (`1866697`) on 2026-09-19: `joy.txt` is copy-paste
`Define` lines; `Use library` imports VLA code, not English; `say` and
`show` are both `(msgbox …)`; there is no text-to-speech and no `=VLA()`
function; programs can react to a sheet change and a button click, and to
nothing else; the CLI is one text box fed through the interpreter.
No Excel was run, so "should work" below means "reads as if it should".
The ideas were also checked against both roadmaps, `MARKETING.md` and
`CONTEMPLATIONS.md`; what is already claimed there is listed near the end
instead of being pitched again.*

*Second sitting, later the same day — which `MARKETING.md` §4.5 notes is
Talk Like a Pirate Day. The owner asked for "some more high-brow, some
retro, and a bunch of extensibility feature ideas", and highlighted one:
custom shortcuts in the CLI, such as a `beautify {range}` that applies ten
table styles, defined as "a (short-long command form) macro" with capture
variables "like in the usual grammar". The owner's stated belief, quoted
because all of section V rests on it: "the more extensible you make a
tool, the more easily the tool maps to any given user's motor skill
muscle-memory. And, of course, extensibility is at the heart of every
Lisp; VLA included." Spitballs 24 to 51 are this sitting's. Checked the
same way, at the same `HEAD`: the console sends any text that does not
start with `(` through the loaded phrasebook; `To …:` verbs die with
their program; `EnglishLoadVocabularyText` and `EnglishAddPhrase` are
`Public`; the console keeps no history; nothing in the grammar or the
interpreter can say "the selection"; and nothing in the language draws a
random number, which corrects an assumption the first sitting made (see
Spitball 34, and the note under Guardrails).*

*Third sitting, 2026-09-20. The owner called this file "a candy store"
and asked whether the orthogonality could stretch further: were there
other types of "autotelic spitballs that would tickle the minds and
fancies of spreadsheeters and hackers alike"? The first two sittings
mined computing's own culture and the tool's own organs. This one mines
other people's cultures that are already spreadsheet-shaped, found with a
three-part test: a practice is a spitball type when it has its own
constrained language for giving instructions, an artifact that is a grid,
and rules worth checking — which feed the phrasebook, the sheet and the
logic engines, in that order. The owner's verdict on the list: "I want
them all." Spitballs 52 to 93 are this sitting's, and a table of enablers
follows them. Checked at `HEAD` `006d193`, which had moved overnight:
`=SQL()` has `GROUP BY` and `HAVING`, but refuses a self-join, a cross
join and an unnamed computed column, each by name (which corrects
Spitball 39, in place); a `DATALOG` rule may name one relation twice;
English expressions have no word for a remainder, though conditions have
"is divisible by"; `To get … using …:` with `Give back …` makes an
English function; no sentence writes a text file; `VlaCompileToForms` is
`Public`; `OPTIMIZE.2` has shipped and `OPTIMIZE.3`, the first search,
has not; and the logic corpus is all business, with none of these hobbies
in it.*

---

## The two rules that sort a spitball from an Excel trick

1. **Only a deterministic system can do it.** Every run snapshots first,
   the interpreter has no randomness and no model in the loop, and a
   sentence means one thing. Anything that needs perfect replay is
   something a prompt box structurally cannot copy.
2. **Everything fun is a header file.** The joy palette's real idea is
   not fifteen colours. It is that a named set of anything — colours,
   sprites, fonts, songs, levels, rule strings — can be a small text file
   that a program pulls in by name.

---

## I. Only a deterministic system can do this

**Spitball 1 — The time scrubber.** A slider under the program. Drag it
and the workbook rewinds and replays to step N. Omniscient debugging, for
a spreadsheet.

- *Stands on:* `TakeRunSnapshot` before every run, a deterministic
  interpreter, and the effect log behind Interpret and Trace. "Go to step
  37" is: restore the snapshot, then re-run with a step budget — the hook
  `IN.14` wants in the statement loop anyway.
- *The catch:* a scrub costs a replay, so long programs scrub slowly. A
  program that reads the clock, a file or a dialog cannot be replayed
  faithfully and should be refused by name: "this program reads the
  clock, so its past can't be replayed." Undo's snapshots do not cover
  Names, and a CLI command takes no snapshot at all.
- *Size:* `~weeks`.

**Spitball 2 — Live mode.** Edit a sentence — change 0.15 to 0.2 — and
the workbook rewinds and replays on its own. Bret Victor's "Inventing on
Principle", for SOPs.

- *Stands on:* Spitball 1's machinery, triggered by a change to the
  workspace sheet.
- *The catch:* small programs only (the interpreter is slow by `IN.8`'s
  own numbers), it needs a debounce, and it must refuse outright any
  program with an external effect. Nothing should draft an email because
  a cell was edited.
- *Size:* `~days` after Spitball 1.

**Spitball 3 — The karaoke run.** Run in slow motion: the current
sentence lights up and the cell it touches flashes. Built for teaching,
and every screen recording of it is a finished social video.

- *Stands on:* the IDE already paints each row green or red (`MarkOK`,
  `MarkErr`). This paints the row that is *running*.
- *The catch:* it needs the same yield point in the statement loop that
  `IN.14` does. It is `IN.14`'s fun twin and should land with it.
- *Size:* `~days` once `IN.14` exists.

**Spitball 4 — "What did I just do?"** Take a snapshot, work by hand for
a minute, press the button. Frazaro diffs the sheets and writes the
sentences. The macro recorder, except it writes prose.

- *Stands on:* the snapshot sheets, and `G-RENDER` (`EnglishRenderForm`),
  which turns a form back into a sentence through the loaded phrasebook
  and has no user-facing job yet.
- *The catch:* a diff sees states, not actions. It yields *a* program
  that reproduces the difference, not the path the person took, and a
  sort looks like a hundred cell edits until something recognises it.
  Start with values, formulas, bold and fill. The upside of a diff: Excel
  fires no event for a formatting change, so a recorder built on events
  could never see one.
- *Size:* `~weeks`.

---

## II. The sheet is a screen

**Spitball 5 — Header files, made real.** `joy.txt` tells its reader to
copy the `Define` lines into their own procedure. `Use library "x.vla".`
imports VLA code, not English. The missing sentence is
`Use palette "joy".` — or, more generally, an English file of `Define`
lines and `To …:` procedures that a program pulls in by name. English
libraries.

- *Stands on:* `G12`'s import sentence and its splice-once rule.
- *The catch:* an included file is program text, so it inherits every
  trust rule a program has (`SEC.8`, and `SEC.9`'s lesson about files
  that sit beside a workbook). A name defined twice must refuse, not
  shadow.
- *Size:* `~days`.

**Spitball 6 — Palettes worth importing.** xkcd's colour survey: 954
names, dedicated to the public domain, so "booger green" is legally
clean. Solarized, Gruvbox, Nord and Dracula, so people can theme a
workbook like their terminal. Okabe-Ito, eight colours chosen to survive
colour blindness: a default that is kind. Game Boy green, the C64's
sixteen, CGA.

- *Stands on:* hyphenated names already work (`phthalo-blue`), and
  `Set fill-color of … to …` takes a hex string.
- *The catch:* barely one. Names with spaces or apostrophes need
  hyphenating (`robins-egg-blue`). These work today as copy-paste files,
  and better after Spitball 5.
- *Size:* `~hours` each.

**Spitball 7 — Sprites and a pixel font.**
`Write "HELLO" in big letters at B2 in phthalo-blue.` A 5×7 font is a
header file. So is a sprite, as rows of palette names. Instant visual
payoff for a demo.

- *Stands on:* fills, loops and lists.
- *The catch:* cell-by-cell fills through the interpreter are slow; the
  prelude's `with-fast-excel` helps.
- *Size:* `~days`.

**Spitball 8 — Turtle graphics.** `To draw a square:` is Logo's
`TO SQUARE`. Frazaro is Logo's grandchild and has not said so yet. A
`turtle.vla` phrasebook: "Move forward 10. Turn right 90. Pen up.", drawn
in cell fills.

- *Stands on:* phrase rules, and the `To …:` definition form that already
  exists.
- *The catch:* line drawing wants a runtime helper, and today every new
  helper is an edit in three modules (Contemplation 9, collapse 3). Keep
  headings to multiples of 45° and the arithmetic stays whole-number and
  exact.
- *Size:* `~days`. Teachers are the audience (`MARKETING.md` §3.12).

**Spitball 9 — Life, stated as rules.** Conway's Life as a few `DATALOG`
rules over a Table of live cells, one generation per button click, with
rule variants (`B3/S23`, HighLife, Seeds) as header files.

- *Stands on:* `DATALOG` has `count`, `let` arithmetic and stratified
  negation, which is what Life needs, and `When "Next" is clicked:`
  exists.
- *The catch:* unverified. Nobody has written it. `G-PROLOG`'s English
  probably cannot say "has three live neighbours" yet; the s-expression
  form should. Measure a 20 × 20 board before promising anything — the
  owner's asymptotics rule.
- *Size:* `~days`.

**Spitball 10 — Rule 110 as a worksheet function.** A `deflambda` that
maps one row of cells to the next. Rule 110 is Turing-complete, so that
is a universal computer registered as a Name, born from an s-expression.

- *Stands on:* `deflambda`, and `alonzo.vla`'s iteration-shaped bricks
  (`REDUCE`, `SCAN`, `MAP`).
- *The catch:* `deflambda` is export-only. The interpreter has no arm for
  it, so this needs Compile, or an interpreter bridge to `Names.Add`.
- *Size:* `~hours` for the brick.

**Spitball 11 — A demo disk, and a sixteen-sentence compo.** Plasma,
fire, a starfield, matrix rain: N frames, a stated seed, cell fills. Then
a competition: best effect in sixteen sentences. The demoscene cannot
resist a size limit, and `MARKETING.md` §4.9 already has the shape of a
challenge post.

- *Stands on:* fills and loops, and Spitball 3's yield point for frames.
- *The catch:* the frame rate is whatever the interpreter gives. In a
  size-coding contest that is part of the sport.
- *Size:* `~days` for the disk. The compo costs a post.

---

## III. Language toys

**Spitball 12 — Ancestor phrasebooks.** "Put 5 into cell B2" is already
HyperTalk, word for word. A `hypertalk.vla` adds `answer "Hello"` and
`add 1 to total`. A `flowmatic.vla` runs `MOVE 5 TO CELL B2.` Grace
Hopper's FLOW-MATIC was English for business procedures in the 1950s. It
is this project's origin myth, as a loadable file.

- *Stands on:* the dialect mechanism `pirate.vla` and its siblings
  already use.
- *The catch:* a phrasebook gets the verbs, not the loops.
  `repeat with i = 1 to 10` is block structure, and block structure is
  engine grammar (Contemplation 9, "The one moving the wrong way").
- *Size:* `~hours` each.

**Spitball 13 — Typology tests in costume.** `yoda.vla` is object-first:
"Into cell B2, 5 put." Klingon is object–verb–subject. Toki Pona has
about 120 words: a small checked language for a small checked grammar.
An emoji phrasebook pushes surrogate pairs through the tokenizer.
Contemplation 3 asked for exactly this falsification test — a verb-final
language, a case-marking one, an unspaced one — and these are that test
wearing a costume.

- *Stands on:* the pattern language.
- *The catch:* they will hit the engine's English assumptions.
  `espanol.vla`'s header already records two: the article `a` dropped
  from every pattern, and rule order. That is the point. If `yoda.vla`
  refuses to load, the language-neutrality claim learned something real.
- *Size:* `~hours` each, plus whatever they find.

**Spitball 14 — The quine relay.** An English program whose output is
itself in Pirate, whose output is itself in Latin, and so on back to
English — after Yusuke Endoh's 128-language relay.

- *Stands on:* the dialects share one template layer, and `G-RENDER` can
  turn forms back into sentences.
- *The catch:* there are two versions and only one is cheap. The
  renderer-assisted relay ("put this program, in Pirate, into column A")
  reads its own source, which a purist calls cheating. It is also the
  hardest `G-RENDER` round-trip test anyone could write, so it earns its
  keep as a test. A true quine, building its own text from a string, is
  a puzzle. That is why people post them.
- *Size:* `~days` for the assisted relay. The true one is a weekend for
  someone who likes that kind of thing.

**Spitball 15 — Interactive fiction.** The world is English clauses. The
player's commands — "go north", "take lamp" — go through a phrasebook, so
`SD-16` holds: no second parser. The console is `` Ctrl+Shift+` ``, the
map is cells, and the saved game is the workbook. The genre's curse is
guess-the-verb. Here refusals teach, and Known Sentences lists the verbs.

- *Stands on:* `G-PROLOG`'s clause grammar, `PROLOG` over Tables, phrase
  rules, the CLI.
- *The catch:* whether the clause grammar stretches to "The lamp is in
  the Kitchen." is untested, and the CLI is one text box with no
  scrollback, so it needs an output pane first.
- *Size:* `~weeks`, as an example workbook.

---

## IV. The console and the cell

**Spitball 16 — A Prolog toplevel in the CLI.** `?- (ancestor tom X)`
against the workbook's live Tables; press `;` for the next answer. Anyone
who has used a Prolog system knows what to do.

- *Stands on:* `PROLOG` already returns every solution, so `;` only pages
  through rows.
- *The catch:* the same output pane as Spitball 15, and a rule for which
  Tables are in scope.
- *Size:* `~days`.

**Spitball 17 — `=VLA("(* 6 7)")`.** Lisp in a cell. Greenspun's tenth
rule, on purpose. Cells come in as named arguments:
`=VLA("(+ a b)", "a", A1, "b", B1)`.

- *Stands on:* `VlaEvalExpression` is already `Public`.
- *The catch:* pure expressions only — arithmetic, text, the prelude's
  predicates. It must refuse `range`, member access and helpers by name,
  because Excel does not track a read inside a function as a dependency,
  and the answer would go stale. Contemplation 5 met the same wall with
  reflection.
- *Size:* `~days`. It pairs with Church numerals and a Y combinator as
  `alonzo.vla` bricks, under Spitball 10's catch.

**Spitball 18 — Read it aloud.** `Read cell B2 aloud.` Excel's speech
engine is built in and works offline. An accessibility win that is also a
party trick.

- *Stands on:* nothing yet. `say` and `show` are both `(msgbox …)` today.
- *The catch:* `SD-4`. `say` is a shipped spelling and keeps its meaning,
  so speech needs a new sentence, not a new `say`. Windows first; the Mac
  is unchecked.
- *Size:* `~hours` to `~days`.

**Spitball 19 — Sparklines made of text.** `▁▂▃▅▇` in one cell, or a
braille-dot line plot. As an `alonzo.vla` brick built on `UNICHAR` it is
a pure LAMBDA: no VBA in the recalculation at all.

- *Stands on:* `deflambda` and `MAP`.
- *The catch:* Spitball 10's.
- *Size:* `~hours`.

**Spitball 20 — Tracker music.** A spreadsheet already is a tracker: rows
are time, columns are channels. `Play the song in table Melody.`

- *Stands on:* nothing in the repository.
- *The catch:* it needs Windows MIDI `Declare`s, which land on `EN.8`'s
  ratchet (one site today, `frmCLI.frm`). There is no Mac path. It
  belongs in an optional module that never ships in the core add-in.
- *Size:* `~days`.

**Spitball 21 — Console toys.** `fortune` prints one line from
`sententiae.txt`'s current draft, chosen by the date: a fortune that
refuses to be random. `monkeysay` is `cowsay`, delivered by the suited
monkey from the project's social posts. And `Xyzzy.` answers "Nothing
happens." — then teaches, because `LX.8` applies to jokes too.

- *Stands on:* the CLI. The test suite already uses "Xyzzy plugh
  nonsense." as its canonical nonsense, so the culture is in place.
- *The catch:* an easter-egg refusal is still a refusal. It gets a stable
  ID, a proof, and a second sentence that says what to write instead.
- *Size:* `~hours`.

**Spitball 22 — Games with readable brains.** Turn-based games in
English: 2048, Sokoban, tic-tac-toe where the opponent is a range of
English clauses. Edit the clauses and the opponent plays differently.

- *Stands on:* button-click and sheet-change handlers exist, and
  `Application.OnKey` already powers the CLI hotkey, so
  `When the left arrow is pressed:` is one event family away.
- *The catch:* `OnKey` bindings are application-wide and survive an
  error, so a game that takes the arrow keys must give them back on every
  exit path. `OnTime` ticks in whole seconds, which is why these are
  turn-based.
- *Size:* `~weeks`.

**Spitball 23 — Brainfuck in twenty-five sentences.** The tape is column
A. A playful proof that the sentence language is Turing-complete, in a
form hackers recognise on sight.

- *Stands on:* `While`, `If` and cells.
- *The catch:* it needs a "character N of" phrase, and `mid` is not among
  the interpreter's built-in functions today (`left`, `right`, `instr`
  and `len` are).
- *Size:* `~hours`.

---

## V. Extensibility, or muscle memory

*A third sorting rule, the owner's: **a tool should bend to the hand that
uses it.** Every command line that lasted grew the same three layers:
commands that ship, names the user gives their own, and a startup file
that loads them. `alias` and `doskey`; Vim's `:command`; fish's `abbr`;
and, closest to home, AutoCAD, where a drafter's two-letter aliases live
in `acad.pgp`, new commands are Lisp functions named `C:…`, and
`acad.lsp` loads them at startup. A command line inside a GUI
application, extended by its users in Lisp, is a forty-year-old success
story. VLA's name is a pun on VBA. Whether or not it was meant to, it
also rhymes with AutoCAD's Visual LISP.*

**Spitball 24 — Commands you name yourself.** *The owner's.*
`beautify {r:range}` applies ten table styles to whatever range follows
it: a short form, a long form, and capture variables in the grammar's own
syntax. The owner asked whether that is the essence of all CLI handiwork.
It is.

- *Stands on:* more than it looks. The console already sends any text
  that does not start with `(` through the loaded phrasebook, so a rule
  in a loaded phrasebook is already a console command:
  `(english-vla "beautify {r:range}" (beautify-table (range {r})))`
  beside a `defmacro` should work today. *Unverified: read, not run.*
  `To stamp, with row-number of 1 and value of "ok":` already gives a
  program its own verbs, but they die with the program, and each console
  command is its own program.
- *What is new:* three things.
  1. *Define it at the prompt.* A phrasebook directive typed into the
     console loads into the live grammar (`EnglishLoadVocabularyText` is
     already `Public`), proofs and shadow audit included. The REPL
     extends the language it reads, which is the Lisp way.
  2. *A long form in English.*
     `(english-english "beautify {r:range}" "Set style of table {r} to TableStyleMedium2." …)`
     is an arrow from English to English, and it composes with
     `english-vla` the way the heads say it should (Contemplation 9). It
     can only say what could already be said. It adds brevity, never
     power, so it needs no consent prompt. An `english-vla` rule over a
     macro — the owner's form — stays available for the cases that need
     power.
  3. *Expansion you can see.* fish's `abbr`, not bash's `alias`: the
     console prints the long form before it runs, and the long form is
     what the history, the trace and any saved SOP keep. In a workspace
     sheet the short form expands in place. An auditor never meets a
     private word.
- *The catch:* a personal command that shadows a shipped sentence must
  refuse unless it is marked as an override — the existing `G3` rule,
  applied to a person instead of a dialect. How strict the prompt should
  be about capitals and full stops is a surface decision, and `SD-4` will
  freeze whatever ships.
- *Size:* `~days` for prompt-defined rules, `~days` more for
  `english-english`.

**Spitball 25 — Define by demonstration.** Do it by hand at the prompt,
then: `remember the last three as "tidy {r:range}", where {r} was A1:D10`.
History becomes a command. Emacs keyboard macros and Vim's `q`, with one
difference: the person names the hole.

- *Stands on:* Spitball 24's `english-english`, since the long form is
  just the sentences that ran.
- *The catch:* the console keeps no history today, so that comes first.
  Turning a literal into a slot must be explicit. Frazaro does not guess
  which `A1:D10` was meant.
- *Size:* `~days` after Spitball 24.

**Spitball 26 — Snippets: commands with holes.** A long form that leaves
slots open expands into the workspace sheet as a skeleton: type `aging`
and get ten sentences with the ranges left to fill. `LE.3` completes a
sentence; a snippet drops in a procedure.

- *Stands on:* Spitball 24.
- *The catch:* an unfilled hole must fail Check by name, never run as
  text.
- *Size:* `~days`.

**Spitball 27 — `init.vla`, the dotfile.** Personal rules, commands,
palettes and bindings, loaded at startup from the user's profile.
`.emacs` for Excel.

- *Stands on:* `EnglishLoadVocabulary`, and the hash-keyed consent that
  `SEC.2` and `SEC.9` already use.
- *The catch:* trust decides where it lives. In the person's profile,
  never beside or inside a workbook: `SEC.9` and `SEC.10` are both
  lessons about a workbook carrying its own words or its own permission
  slip. Nothing in the product uses the profile folder today. A broken
  init file must not take the add-in down with it, so there is a safe
  mode, and it says so when it skips the file.
- *Size:* `~days`.

**Spitball 28 — Bindings as sentences.**
`Bind Ctrl+Shift+B to: beautify the selection.` A key, an item on the
Frazaro menu, or a button on the sheet. Select, press, done: muscle
memory in the literal sense.

- *Stands on:* `Application.OnKey` powers the CLI hotkey, the Frazaro
  menu is built at run time from temporary command-bar buttons, and
  `Make a button … at cell …` exists.
- *The catch:* "the selection" cannot be said today. The grammar has no
  such reference and the interpreter's receivers do not include it.
  `OnKey` is application-wide, so binding over one of Excel's own
  shortcuts must be refused unless the sentence says so, `bindings` lists
  what is bound, and unloading gives everything back.
- *Size:* `~days`.

**Spitball 29 — Observer hooks.** `After every run: …`
`When a program is refused: …` Emacs hooks and git hooks, with one house
principle: a personal hook may watch but not touch. It can log, speak, or
stamp a run sheet. It cannot change the workbook, because then one
person's setup would change what another person's SOP means.

- *Stands on:* runs already pass through two doors, `InterpretProgram`
  and `VlaCliRun`.
- *The catch:* "watch, don't touch" needs the declared effect classes of
  Contemplation 9's third collapse before it can be enforced and not
  merely promised.
- *Size:* `~days`.

**Spitball 30 — Packs without a network.** A pack is a folder:
`pack.vla` (rules, macros, proofs), its header files, one example. At the
prompt: `packs`, `pack on turtle`, `pack off turtle`, `pack test turtle`,
and `pack why "Beautify A1:D10."` to see whose rule won. Homebrew, minus
the network.

- *Stands on:* phrasebook chains; proofs that refuse a broken pack at
  load; `(requires-version …)` and `(requires-capability …)` (`F.10`);
  and Explain's provenance, which already names the winning file.
- *The catch:* `SD-13`, so packs travel by file, by USB stick, by git.
  Load order is meaning (`G3`), so `packs` must show the order. `GO.3`'s
  warning stands: rules emit code, so a pack is a supply chain, and `raw`
  consent applies per pack.
- *Size:* `~weeks`.

**Spitball 31 — The command palette.** `Ctrl+Shift+P`: fuzzy-find any
Known Sentence or personal command, pick it, and fill its slots from
prompts that check shape as you type. For people who live on the mouse.
The ranking is deterministic — subsequence match, stable tie-break, no
learning — so the same keystrokes always find the same sentence, which is
what muscle memory needs.

- *Stands on:* the Known Sentences list, and the reference-shape
  validators behind typed slots.
- *The catch:* it is a UserForm, and the Mac's are weaker. A sibling of
  `LE.3`, not a substitute.
- *Size:* `~days` to `~weeks`.

**Spitball 32 — Enumerations from a Table.** `{a:account}`, where the
legal accounts are a column of a Table. `Post 500 to travel.` checks.
`Post 500 to travle.` is refused at Check, with the near miss named.
Domain vocabularies without touching the engine.

- *Stands on:* `{d:bold|italic}` is this already, with the branches typed
  by hand.
- *The catch:* `SD-16` guards this door. A finite word list is regular,
  but the grammar would then depend on workbook data: the shadow audit
  has to re-run when the Table changes, and one sentence is valid in one
  workbook and not in another. That is a parsing fork, and it wants the
  owner's long-term analysis before anyone builds it.
- *Size:* `~days` to build, longer to decide.

**Spitball 33 — Units you define.** `Define a fortnight as 14 days.`
`Set due to today plus 2 fortnights.` — and a refusal, at Check, for
adding metres to dollars. Dimensional analysis, as in F#'s units of
measure, for SOPs. Unit systems are header files.

- *Stands on:* `Define`.
- *The catch:* it lives inside the built-in `expr` sub-grammar, so it is
  engine work. Only literals and defined names carry a dimension. A
  number read from a cell has none unless the sentence gives it one.
- *Size:* `~weeks`.

---

## VI. High-brow

**Spitball 34 — Dice with a serial number.** "Anyone who considers
arithmetical methods of producing random digits is, of course, in a state
of sin" (von Neumann). Frazaro can only ever sin in the open:
`Pick a number from 1 to 6 with seed 42.` The generator is the project's
own few lines, identical on every machine and under both backends, and an
unseeded draw is refused by name, as Contemplation 5 recommended. The
high-brow edition draws from a published table, the way statisticians did
before computers — RAND's *A Million Random Digits* (1955) is the famous
one — so that a draw can be cited by page and line.

- *Stands on:* nothing. No sentence, macro or helper draws a random
  number today. Spitballs 11, 38, 44 and 50 all assume one.
- *The catch:* VBA's own `Rnd` is not a specification, and
  `INTRINSICS.md` exists because a port must match behaviour exactly.
  Check RAND's terms before shipping a digit of the book.
- *Size:* `~days`.

**Spitball 35 — Carroll's sorites.** "No ducks waltz. No officers ever
decline to waltz. All my poultry are ducks." Therefore my poultry are not
officers. Lewis Carroll's *Symbolic Logic* (1896) is a public-domain
corpus of exactly the sentences `G-PROLOG` wants to read, with
Aristotle's syllogisms as the warm-up. The owner's own `spreadsheet.lisp`
already answers "All ? are mortal." in pure LAMBDA.

- *Stands on:* `G-PROLOG`'s clause and question grammar, and `DATALOG`.
- *The catch:* Carroll reasons by contraposition, and negation as failure
  is not classical negation. Some of his conclusions will be out of reach
  without stating both polarities, and finding out which is a result
  worth writing down.
- *Size:* `~days`.

**Spitball 36 — Bisect the run.** A run ends wrong. State what should
have held — `Column D never goes below zero.` — and Frazaro finds the
first sentence that broke it, in log₂ N replays. `git bisect` for a
spreadsheet; runtime verification's "always", asked after the fact.

- *Stands on:* Spitball 1's replay, and the prelude's `check`.
- *The catch:* it is sound only for replayable programs, and the property
  must be pure. If the property is cheap and known in advance, checking
  it after every step is simpler. Bisecting earns its keep when the
  question arrives after the damage.
- *Size:* `~days` after Spitball 1.

**Spitball 37 — Literate SOPs: tangle and weave.** Knuth's WEB split one
document into a program (tangle) and a typeset essay (weave). Frazaro
already tangles: Import Program File reads the sentences out of a Word
document. Weave is missing: one document per run, each sentence set
beside the VLA it became and what it did. An executable paper, and the
audit artifact as literature.

- *Stands on:* the Word import, Show me the VBA (`IN.4`), Interpret and
  Trace.
- *The catch:* writing Word or PDF is host work, and Word automation is a
  guarded surface (`SEC.13`). Start with a "Woven" sheet.
- *Size:* `~days`.

**Spitball 38 — Instruction art.** Sol LeWitt, 1967: "The idea becomes a
machine that makes the art." His wall drawings are sentences that other
people execute. So:
`Place fifty points with seed 7. Join every point to every other.` The
sheet is the wall, and the SOP is the certificate of authenticity.

- *Stands on:* Spitball 8's line drawing and Spitball 34's dice.
- *The catch:* LeWitt's own texts belong to his estate. Write new ones in
  the form.
- *Size:* `~hours` after those two.

**Spitball 39 — A canon is a self-join.** A melody is a Table of (beat,
pitch). A canon is the same rows again, later and higher:
`SELECT beat + 8, pitch + 7 FROM melody UNION ALL SELECT beat, pitch FROM melody`.
The crab canon of Bach's *Musical Offering* — the one *Gödel, Escher,
Bach* is built around — is the same Table read backwards. A piano roll in
cell fills shows it even in silence.

- *Stands on:* `=SQL()` has arithmetic in `SELECT`, and `UNION ALL`.
- *The catch:* hearing it needs Spitball 20.
- *Size:* `~hours`.

*Corrected at the third sitting (2026-09-20):* `=SQL()` refuses a
computed column that has no name, so the query is
`SELECT beat + 8 AS beat, pitch + 7 AS pitch FROM melody UNION ALL SELECT beat, pitch FROM melody`.
"Self-join" in the title is a figure of speech: the engine refuses a true
self-join by name, and whether one Table may stand on both sides of a
`UNION` is unchecked.

**Spitball 40 — The commuting diagram.** Draw the project's category on a
sheet: languages as boxes, phrasebooks as arrows — `english → vla ←
espanol`, `vla → vba`, `vla → formula`. An arrow goes red when one of its
proofs fails. A square goes red when its two paths disagree on the paired
corpus. Contemplation 5's Milewski, made literal.

- *Stands on:* `test-success` proofs, and per-rule provenance by file.
- *The catch:* the arrows below the waist are not phrasebooks yet
  (Contemplation 9's first collapse), so they have no proofs to colour.
- *Size:* `~days`.

**Spitball 41 — Philology for procedures.** An interlinear gloss: Explain
laid out the way linguists lay out a sentence — the words, under them the
slot and its category, under that the value each one bound. And a
concordance, the keyword-in-context index of the 1950s: every sentence in
a workbook that uses "sort", aligned on the word.

- *Stands on:* `EnglishExplain` already knows the tokens, the claim, the
  rule and the file it came from.
- *The catch:* none. It is presentation.
- *Size:* `~hours` to `~days`.

**Spitball 42 — Basic English, and other vows.** C. K. Ogden's Basic
English (1930) is 850 words, and the ancestor of every controlled
language, this one included. As a lint: flag any word in an SOP that is
outside the list. Its Oulipian twin is a lipogram phrasebook with no
letter e, in which `cell` itself is forbidden. A small checked English,
checked smaller.

- *Stands on:* the tokenizer, and a word list as a header file.
- *The catch:* check the word list's status before shipping it. Modern
  controlled-language dictionaries, such as the aerospace industry's, are
  licensed.
- *Size:* `~hours`.

**Spitball 43 — Bartleby mode.** A voice pack. Every refusal opens "I
would prefer not to." (Melville, 1853) and then teaches as usual. Once
refusals are forms (Contemplation 9's second collapse), a voice is a
phrasebook overlay, and pirate refusals for 19 September come free.

- *Stands on:* nothing until then. The catalogue is VBA today.
- *The catch:* `LX.8`. The voice may change; the teaching may not. A
  proof should pin a refusal's ID, never its voiced text.
- *Size:* `~hours` after that collapse.

---

## VII. Retro

**Spitball 44 — 10 PRINT.** `10 PRINT CHR$(205.5+RND(1)); : GOTO 10` —
the Commodore 64 one-liner that fills a screen with a maze, and got a
whole book from MIT Press. Here it is one sentence:
`Fill range A1:AN25 with ╱ or ╲, with seed 64.`

- *Stands on:* Spitball 34.
- *The catch:* none beyond it.
- *Size:* `~hours`.

**Spitball 45 — Line numbers.** `10 PUT 5 INTO CELL B2` …
`20 GOTO 10`. `goto` and `label` are core forms already, because VLA is
VBA wearing parentheses. A `basic.vla` costume for everyone's first
language, with a lint that says "considered harmful" (Dijkstra, 1968) and
runs it anyway.

- *Stands on:* the core forms.
- *The catch:* a number in front of every sentence is structure, not a
  phrase, so it is engine grammar again. And a `GOTO` loop meets no step
  ceiling: `SEC.14` is an accepted risk and `IN.14` is open.
- *Size:* `~days`.

**Spitball 46 — Slash commands.** Lotus 1-2-3's `/` menu has lived in a
generation's fingers for forty years: `/re` erases a range, `/wir`
inserts a row, `/fs` saves. A command pack built with Spitball 24:
`/re A1:B5` expands, visibly, to `Clear range A1:B5.` It is the owner's
muscle-memory thesis, tested on the people with the oldest muscle memory
in the business. Excel itself honoured these keys for decades.

- *Stands on:* Spitball 24, and sentences that already exist.
- *The catch:* whether `/` survives the tokenizer as the start of a word
  is unchecked.
- *Size:* `~hours` after Spitball 24.

**Spitball 47 — The CRT pack.** Palettes: P1 phosphor green, amber, Turbo
Pascal's yellow on blue, and Windows 3.1's Hot Dog Stand. For the
console, typewriter pacing and a bell. Header files again.

- *Stands on:* Spitballs 5 and 6, and Spitball 15's output pane.
- *The catch:* none.
- *Size:* `~hours`.

**Spitball 48 — ELIZA.** Weizenbaum, 1966: keywords, captured fragments,
response templates. ELIZA's script is a phrasebook, and a `doctor.vla`
would run in the console. The first chatbot was a rulebook and said so in
its source. In the age of the prompt box, that is this project's argument
as a museum piece.

- *Stands on:* the pattern language, and the console.
- *The catch:* ELIZA needs a slot that swallows the rest of a sentence —
  "I am {x}" — and `SD-16`'s slots take one token or one quoted string. A
  rest-of-sentence slot is regular, but it is a grammar decision and the
  owner's to make. The pronoun swap (my → your) is a small table.
- *Size:* `~days`, plus the decision.

**Spitball 49 — Punch cards.** Render an SOP as a deck: one sentence per
80-column card, its Hollerith code in cell fills, the sentence printed
along the top edge. A lint comes free — "this sentence does not fit on a
card" — and it is the oldest style rule in computing. It pairs with
`flowmatic.vla` from Spitball 12.

- *Stands on:* Spitball 7's sprites.
- *The catch:* none. It is a rendering.
- *Size:* `~days`.

**Spitball 50 — Hamurabi.** 1968: ten years ruling Sumer by typing
numbers — acres, bushels, plague. It is a ledger with a narrator: the
first spreadsheet game, a decade before the spreadsheet. About sixty
readable sentences, played at the prompt. And Rogue, where the dungeon is
a function of one cell: a seed of the day that everyone can share with no
network at all, because the seed is the date.

- *Stands on:* `Ask … and put answer into …` exists, with loops and
  conditions.
- *The catch:* Spitball 34's dice, Spitball 15's output pane, and for
  Rogue, Spitball 22's keys.
- *Size:* `~days`.

**Spitball 51 — Copy as a text table.** Box-drawing characters from the
BBS era, `┌─┬─┐`, and their modern twin:
`Copy range A1:D10 as Markdown.` Developers paste tables into issues,
READMEs and chat every day.

- *Stands on:* nothing needed. It is a pure function from a range to
  text.
- *The catch:* the clipboard is host work; writing the text to a cell is
  not. Alignment wants a monospaced font.
- *Size:* `~hours`.

---

## VIII. Found phrasebooks

*Controlled languages that people invented long before computing. Each is
a costume over the same pattern language, so `SD-16` holds. The best of
them are ancestors of the core idea, not decorations.*

**Spitball 52 — Blazon.** `Azure, a bend Or.` Heraldry's language is some
eight hundred years old, has a real grammar, and compiles to pictures: a
blue shield, a gold diagonal band. The tinctures — Or, Argent, Gules,
Azure, Vert, Purpure, Sable — are a palette header file, and charges are
sprites.

- *Stands on:* alternation slots, fills, Spitballs 5 to 7.
- *The catch:* blazon nests ("a bend Or between two lions…"), and only
  the flat first layer fits one phrase pattern. How the tokenizer treats
  a comma inside a pattern is unchecked.
- *Size:* `~days` for fields, bends, fesses, pales and chiefs.

**Spitball 53 — The International Code of Signals.** Since 1857: a
phrasebook keyed by flag hoists, in which the same hoist means the same
sentence in every language on the sea. Sailors built Contemplation 3's
hourglass — an interlingua of actions — 170 years ago. Two toys:
`Spell FRAZARO in signal flags at B2.`, and a `signals.vla` in which a
hoist is a sentence.

- *Stands on:* sprites (Spitball 7), and the dialect mechanism.
- *The catch:* editions differ in who publishes them. Use a government
  edition, and check.
- *Size:* `~days`.

**Spitball 54 — Telegraph codebooks.** When every word cost money,
merchants bought books that turned five letters into a sentence.
Bentley's *Complete Phrase Code* (1906) is one. That is Spitball 24's
`english-english` arrow a century early, and the old books are in the
public domain. A pack of them is a working museum of compression by
phrasebook.

- *Stands on:* Spitball 24.
- *The catch:* none beyond it.
- *Size:* `~hours` per book, after Spitball 24.

**Spitball 55 — Chess on an 8 × 8 range.** Replay a game from its score.
`♔` and `♟` are ordinary characters, the board is a range, and the karaoke
run (Spitball 3) is the broadcast.

- *Stands on:* fills and characters.
- *The catch:* standard notation ("Nf3") names the destination and leaves
  the origin to the rules of chess, so reading it needs a move generator.
  Coordinate moves ("g1f3") need only the first enabler below. Start
  there.
- *Size:* `~days`.

**Spitball 56 — Siteswap.** Jugglers' notation from the 1980s. `531` is a
pattern and `532` is not, and the test is one line of modular arithmetic:
add each throw to its position, take the remainder by the length, and no
two may collide. The average is the number of balls.

- *Stands on:* arithmetic. As an `alonzo.vla` brick it is a pure LAMBDA.
- *The catch:* Spitball 10's for the brick. In English it wants the first
  two enablers below.
- *Size:* `~hours`.

**Spitball 57 — Change ringing.** English bell-ringers were doing group
theory in the 1600s. Every row is a permutation of the bells, a "method"
generates the rows from a line of place notation, and a touch is *true*
when no row repeats. Rows of numbers are rows of a sheet, truth is a
uniqueness query, and one bell's path down the page — the blue line — is
a fill.

- *Stands on:* `=SQL()`'s `GROUP BY` and `HAVING`, and fills.
- *The catch:* place notation is dense ("x16x16"), and the pattern
  language reads whole words.
- *Size:* `~days`.

---

## IX. People who already spreadsheet for love

*Hobbyists are the people who open a spreadsheet on a Saturday. Each of
these communities already has a constrained English, a grid, and rules.*

**Spitball 58 — The craft room.** Knitters chart in spreadsheets already.
A knitting phrasebook reads a row — "Knit 2, purl 2, yarn over, knit 2
together." — draws the chart, and checks the count: "row 5 uses 38
stitches, but row 4 left 40 on the needle." A refusal in words, for
knitters. Cross-stitch charts come free from sprites, and Jacquard's loom
is the ancestor of Spitball 49's punch cards.

- *Stands on:* the Oxford-comma list grammar, arithmetic, fills.
- *The catch:* the abbreviations knitters actually write (`k2tog`,
  `*…; rep from *`) glue a word to a number, and the pattern language
  reads whole words. Longhand works. The shorthand is engine grammar.
- *Size:* `~days`.

**Spitball 59 — The tabletop.** A character sheet is a workbook, and
house rules are English clauses, so `PROLOG` becomes the rules lawyer:
"Can Mira wield the halberd?" And exact odds by enumeration — "what is
the chance that three dice total at least 12?" — is certainty about
chance, with no randomness anywhere.

- *Stands on:* `G-PROLOG`'s clauses and questions; and `DATALOG`, where
  one rule may name the dice relation twice and `count` does the rest.
- *The catch:* `=SQL()` cannot do the dice, because it refuses a
  self-join and a cross join by name. Keep to generic rules, and away
  from any publisher's names and marks.
- *Size:* `~days`.

**Spitball 60 — The kitchen.** A recipe is an SOP. Scale it to twelve
servings, convert its units (Spitball 33), keep a baker's percentages
honest, and schedule it backwards: "the bread comes out at 18:00, so when
do I start the dough?"

- *Stands on:* arithmetic, `Ask … and put answer into …`, dates.
- *The catch:* ingredient lines are free text ("2 cloves garlic,
  minced"), and only a disciplined subset will parse.
- *Size:* `~days`.

**Spitball 61 — Kinship.** A family tree in a Table, and "second cousin
once removed" defined in English clauses. The interesting part is that
kinship terms differ by language. Danish splits *farmor* from *mormor*
where English has one grandmother, so the phrasebook localises the logic
and not just the words. A real test of the polyglot layer, and a huge
hobby.

- *Stands on:* `G-PROLOG`'s "directly or not", the corpus's Lineage
  family, and `dansk.vla`.
- *The catch:* "cousin" means "shares a grandparent but not a parent", so
  it needs negation and distinctness together. Importing a GEDCOM file
  waits on `L-FILE-HELPERS`.
- *Size:* `~days`.

**Spitball 62 — The league office.** Club secretaries run on
spreadsheets. A round robin by the circle method is pure arithmetic. A
Swiss pairing — nobody meets twice, colours alternate — is a constraint
problem stated in English.

- *Stands on:* arithmetic today; `OPTIMIZE` for the pairings.
- *The catch:* `OPTIMIZE` checks constraints but does not search yet
  (`OPTIMIZE.3`), and Contemplation 8's arithmetic says small fields
  only.
- *Size:* `~hours` for the round robin.

---

## X. Famous experiments you can run

*A seeded, replayable run is a reproducible experiment.*

**Spitball 63 — Schelling's checkerboard.** 1971: mild preferences, stark
segregation. Schelling ran it by hand, with coins on a checkerboard. Here
the board is a range and the rule is two sentences.

- *Stands on:* fills and loops.
- *The catch:* the usual model moves an unhappy agent to a random empty
  square (Spitball 34). Moving to the nearest one needs no dice. It is
  slow on a big board (`IN.8`).
- *Size:* `~days`.

**Spitball 64 — Axelrod's tournament.** 1980: strategies for the
prisoner's dilemma play a round robin, and the winner is two sentences
long: "Cooperate first. Then do what the other player did last." Here
every strategy is an English function, so the tournament's source can be
read by the people it is about.

- *Stands on:* `To get … using …:` and `Give back …` already make English
  functions.
- *The catch:* a fixed roster, since a sentence cannot call a function
  that is chosen at run time.
- *Size:* `~days`.

**Spitball 65 — Feynman's orbit table.** *Lectures on Physics*, volume I,
chapter 9: a planet's path computed by hand, one step at a time, in a
table of position, velocity and force. It is a spreadsheet, drawn in
1963. Forty sentences of arithmetic reproduce it, and a chart of two of
its columns is an orbit.

- *Stands on:* `Count … from … to …`, arithmetic, and
  `Put … into column … row …`. It should work today. *Unverified: read,
  not run.*
- *The catch:* none.
- *Size:* `~hours`.

**Spitball 66 — L-systems.** Lindenmayer, 1968: `F → F[+F]F[-F]F`,
applied to its own output, grows a plant. A rewriting rule is a pattern
and a template, so an L-system is a phrasebook that talks to itself, and
Spitball 8's turtle draws what it says.

- *Stands on:* Spitball 8.
- *The catch:* rewriting walks a text one character at a time (the first
  enabler below), and the turtle needs a stack for `[` and `]`.
- *Size:* `~days`.

**Spitball 67 — Turing patterns, and the sandpile.** Turing's 1952 paper
on how a leopard gets its spots, as two numbers per cell and a diffusion
rule. And Bak, Tang and Wiesenfeld's sandpile: drop grains on the centre
cell, topple any cell that holds four, and a fractal appears. The
sandpile is whole numbers and needs no dice.

- *Stands on:* arithmetic and fills.
- *The catch:* every generation touches every cell (`IN.8`), so small
  grids, or a `deflambda` brick.
- *Size:* `~days`.

---

## XI. Gardner's column

**Spitball 68 — The faro shuffle.** Cut the deck exactly in half and
interleave it perfectly. Eight of them put fifty-two cards back in their
original order. It is a shuffle with no randomness in it — the only kind
this project can have without Spitball 34 — and it rhymes with the
product's name. `Shuffle range A1:A52 perfectly.`

- *Stands on:* index arithmetic. It should work today as a short
  procedure. *Unverified: read, not run.*
- *The catch:* none.
- *Size:* `~hours`.

**Spitball 69 — Magic squares.** The Siamese method builds any odd square
with one rule — up and right, or down when the cell is taken — which is a
turtle walk. Check then proves that the rows, columns and diagonals
agree. Dürer's square and the Lo Shu are the fixtures.

- *Stands on:* arithmetic and conditions.
- *The catch:* wrapping round the edge wants a remainder (the second
  enabler below).
- *Size:* `~hours`.

**Spitball 70 — Number curiosities.** Kaprekar's 6174, reached from
almost any four digits by sorting and subtracting. Conway's look-and-say
sequence. Josephus's circle. Ulam's spiral, where the primes line up on
diagonals that nobody has fully explained.

- *Stands on:* `Sort range …`, lists, fills.
- *The catch:* most of them read the digits of a number (the first
  enabler below).
- *Size:* `~hours` each.

---

## XII. Codes and ciphers

*Classical, and in the open. See the guardrail on steganography.*

**Spitball 71 — Playfair is a range.** The Playfair cipher's key square
is a 5 × 5 range, and its three rules — same row, same column, rectangle
— are sentences about cells. Caesar and Vigenère are the warm-ups. A
one-time pad drawn from Spitball 34's published table is the only
unbreakable cipher, done by hand as it was in the field.

- *Stands on:* ranges and lookups.
- *The catch:* the first enabler below.
- *Size:* `~days`.

**Spitball 72 — Enigma with visible rotors.** Each rotor is a column, and
stepping is a rotation of it. Run it in karaoke mode (Spitball 3) and
watch the current pass through three rotors, the reflector, and back.

- *Stands on:* ranges; Spitball 3 for the show.
- *The catch:* the first enabler below, and patience. One letter is a
  dozen lookups.
- *Size:* `~days`.

**Spitball 73 — Check digits as sentences.**
`Check that cell B2 is a valid IBAN.` Luhn's algorithm for card numbers,
ISBN's weighted sum, IBAN's remainder by 97. Quietly the most useful
entry in this sitting.

- *Stands on:* arithmetic.
- *The catch:* the first two enablers below. An IBAN is also a number of
  thirty digits, which overflows a `Double` and has to be reduced a piece
  at a time.
- *Size:* `~days`.

---

## XIII. Clocks, calendars, the heavens

**Spitball 74 — Computus, and the Doomsday rule.** The date of Easter was
the medieval killer app of computation, and Gauss's version is ten lines
of remainders. Conway's Doomsday rule finds the weekday of any date in
the head. `Put the date of Easter in 2027 into cell B2.`

- *Stands on:* arithmetic. As bricks they are pure LAMBDAs.
- *The catch:* the second enabler below for the English version, and
  Spitball 10's catch for the bricks.
- *Size:* `~hours`.

**Spitball 75 — `ddate`.** "Today is Sweetmorn, the 44th day of
Bureaucracy." The Discordian calendar lived among the standard Linux
utilities for years, and a certain kind of person will grin.

- *Stands on:* `today`, arithmetic, a table of names.
- *The catch:* the second enabler below.
- *Size:* `~hours`.

**Spitball 76 — The sky tonight.** Moon phase, sunrise and sunset from a
latitude, the planets' rough places. All of it is arithmetic on a date,
published long ago for pocket calculators. An orrery in cells.

- *Stands on:* arithmetic and fills.
- *The catch:* English has no words for trigonometry yet. Bricks are the
  natural home.
- *Size:* `~days`.

**Spitball 77 — A life in weeks.** A grid 52 wide and 90 tall: one cell
for each week of a long life, with the weeks already lived filled in. A
memento mori that is literally a range.

- *Stands on:* dates and fills. It should work today. *Unverified: read,
  not run.*
- *The catch:* none.
- *Size:* `~hours`.

---

## XIV. Maps and mazes

**Spitball 78 — Tile-grid maps.** One cell per state or country, placed
roughly where it belongs: the data journalists' map, and already a
spreadsheet layout. The arrangement is a header file.
`Colour each state by column B.`

- *Stands on:* Spitball 5, fills, Tables.
- *The catch:* a colour ramp, from a value to a shade, needs a small
  helper.
- *Size:* `~hours` per map.

**Spitball 79 — Four colours.** `No two neighbours share a colour.` The
classic constraint problem, and in 1976 the first theorem proved by
computer. With Spitball 78 the map colours itself.

- *Stands on:* `OPTIMIZE`'s constraints.
- *The catch:* the search is not built (`OPTIMIZE.3`), and Contemplation
  8's arithmetic means small maps.
- *Size:* `~hours` once it is.

**Spitball 80 — Königsberg.** Euler, 1736: can one walk cross all seven
bridges exactly once? The answer is a parity count — how many landmasses
have an odd number of bridges — which is `count` and "is divisible by".
Graph theory's founding question, asked of a Table of bridges.

- *Stands on:* `DATALOG`'s `count`, and the corpus's Routes and
  connectivity family.
- *The catch:* none.
- *Size:* `~hours`.

**Spitball 81 — Mazes.** Generate one from a seed, solve it, paint the
path.

- *Stands on:* `DATALOG` reachability, for "is there a way out".
- *The catch:* reachability says *whether*, not *which way*. The path
  wants distances, and recursion through arithmetic is where a Datalog
  engine gets careful. Generation needs Spitball 34.
- *Size:* `~days`.

---

## XV. Out of the screen

**Spitball 82 — Drawings that leave the screen.** SVG is text. The
turtle, the L-systems and the instruction art can all write it, and then
a pen plotter or a laser cutter draws them. STL is text too:
`Save range A1:Z26 as a relief.` and a 3D printer turns the data into an
object. A fixtures Table becomes an `.ics` calendar the same way.

- *Stands on:* `joined with`.
- *The catch:* no sentence writes a text file today, and one that did
  would be an external effect under `SEC.8`. Writing the text into a cell
  is not.
- *Size:* `~days`.

**Spitball 83 — The SOP as a flowchart.** Walk a program's forms and emit
Mermaid or DOT. Paste it into GitHub and it renders. The diagram every
procedure manual wants, generated from the procedure itself.

- *Stands on:* `VlaCompileToForms` is `Public`, and Show me the VBA is
  the precedent for a generated view.
- *The catch:* it is one more walker over head symbols, and Contemplation
  9 already counted fourteen dispatches — unless it is built as rules.
- *Size:* `~days`.

---

## XVI. Sport and ritual

**Spitball 84 — Golf, and par.** Fewest sentences wins. Every sample SOP
gets a par, and Known Sentences is the rulebook.

- *Stands on:* counting rows.
- *The catch:* golf rewards the opposite of what an auditor wants. Keep
  it in the clubhouse.
- *Size:* `~hours`.

**Spitball 85 — An Advent calendar.** Twenty-four small puzzles, one a
day, each solvable in sentences or a query. They ship as one workbook,
because nothing here is fetched (`SD-13`).

- *Stands on:* everything that ships.
- *The catch:* it is content, and content is work.
- *Size:* `~weeks` of evenings.

**Spitball 86 — Core War on a column.** Dewdney, 1984: two programs fight
over one circular memory. The memory is a column. A warrior is a Table of
an opcode and two operands, so nothing needs parsing. The fight is
deterministic, so Spitball 1 is the instant replay.

- *Stands on:* Tables, loops, fills.
- *The catch:* the second enabler below, for the circular addressing, and
  speed (`IN.8`).
- *Size:* `~days`.

**Spitball 87 — Tournaments of readable brains.** Spitball 22's
opponents and Spitball 64's strategies, entered against each other. Every
entrant's source is English, every match replays exactly, and the bracket
is a sheet.

- *Stands on:* Spitballs 22 and 64.
- *The catch:* theirs.
- *Size:* `~days` after them.

---

## XVII. Civics and fairness

**Spitball 88 — The count.** A club election by instant runoff, counted
step by step where every member can watch: tally the first preferences,
eliminate the lowest, transfer, repeat. Borda and Condorcet on the same
ballots, for comparison. An auditable count is the brand in miniature.

- *Stands on:* `=SQL()`'s `GROUP BY`; Spitball 3 for the watching.
- *The catch:* each round depends on the last one's answer, so a sentence
  has to read a query's result back. *Unverified.*
- *Size:* `~days`.

**Spitball 89 — The Alabama paradox.** 1880: make the House bigger and
Alabama loses a seat. Hamilton's method does that. Jefferson's and
Huntington–Hill's do not. The whole scandal fits on one sheet.

- *Stands on:* arithmetic and `Sort range …`.
- *The catch:* none.
- *Size:* `~hours`.

**Spitball 90 — Secret Santa.** `Nobody draws themselves or their spouse.`
The friendliest constraint demo there is: seasonal, offline, and
understood by everyone. A chore rota is the same problem in an apron.

- *Stands on:* `OPTIMIZE`'s constraints.
- *The catch:* `OPTIMIZE.3` again. For six people the search is tiny.
- *Size:* `~hours` once it is.

---

## XVIII. Office satire

**Spitball 91 — Buzzword bingo.** A card generator: a word list, a grid,
and Spitball 68's faro shuffle, so that every card has a serial number
and the draw can be audited.

- *Stands on:* Spitball 68.
- *The catch:* none.
- *Size:* `~hours`.

**Spitball 92 — The jargon lint.** It flags "leverage", "circle back" and
"going forward" in an SOP, and says what the sentence would be without
them. Plain-language advocacy in a clown nose, and Spitball 42's funnier
twin.

- *Stands on:* the tokenizer, and a word list as a header file.
- *The catch:* none.
- *Size:* `~hours`.

---

## XIX. The shelf as a database

**Spitball 93 — Ask the roadmap.** The roadmap's items, IDs and "depends
on" lines as a Table, and then `DATALOG`: what blocks `IN.14`, directly
or not? This file too: which spitballs wait on which enabler. Dogfooding,
as play.

- *Stands on:* `DATALOG`'s closure, and the ID discipline that makes the
  roadmap parseable at all.
- *The catch:* a script has to turn Markdown into rows first, in the
  `tools/` house style.
- *Size:* `~days`.

---

## The enablers

*Added at the third sitting. Eight small things keep turning up under
"The catch". Each is the cheapest way to make several spitballs possible
at once, and the first two are the cheapest of all.*

| Enabler | Spitballs that wait on it |
|---|---|
| A way to say "character N of a text" | 23, 55, 56, 66, 70, 71, 72, 73 |
| A word for a remainder in an expression | 56, 69, 73, 74, 75, 86 |
| Dice with a seed (Spitball 34) | 11, 38, 44, 50, 63, 81 |
| A console output pane (Spitball 15) | 15, 16, 47, 50 |
| English header files (Spitball 5) | 6, 7, 42, 47, 52, 53, 78, 92 |
| Line drawing (Spitball 8) | 38, 66, 82 |
| `deflambda` under Interpret (Spitball 10) | 17, 19, 56, 67, 74, 76 |
| `OPTIMIZE`'s search (`OPTIMIZE.3`) | 62, 79, 90 |

---

## Already on the shelf, so not here

- **`LE.3`** — in-sheet autocomplete: "a constrained language is an
  autocompletable one." The CLI would be a natural first home for it.
- **`MARKETING.md` §3.3** — a logic-puzzle-in-a-cell series, the zebra
  puzzle first. Sudoku and N-Queens are, in the roadmap's words, "the
  honest hello-world, not the pitch".
- **Contemplation 4** — dictation through a constrained recogniser.
- **Contemplation 5** — reflection: the workbook's own formulas, names
  and sheets as tables, so that `SELECT … FROM formulas` works.

*Added at the second sitting:*

- **`GO.3`** — the community phrasebook registry and its trust model.
  Spitball 30 is its offline half, not a rival.
- **`L.11`** — `@doc:` annotations surfacing through apropos, which is the
  console's `apropos`, already claimed.
- **`EDITION-PARITY`** — the paired-corpus check between editions.
  Spitball 40 would only draw it.

*Added at the third sitting:*

- **`MARKETING.md` §3.3** already lists "a family tree" among its
  puzzle-in-a-cell posts. Spitball 61 is the hobby behind that post.
- **The logic corpus** (`pareto_logic.txt`) has Lineage and Routes and
  connectivity families, both business-shaped. Spitballs 61, 80 and 81
  stand on them.
- **`L-FILE-HELPERS`** — text and CSV import, which Spitball 61's GEDCOM
  waits on.
- **`OPTIMIZE.3`** — the first search, which Spitballs 62, 79 and 90 wait
  on.

## Guardrails that keep it cool

- **`SD-13`: nothing networked.** No multiplayer, no leaderboard, no
  fetching a palette from a URL.
- **Seeds are stated.** Contemplation 5 already refuses random sampling
  without one.
- **Refusals still teach** (`LX.8`), jokes included.
- **One sentence grammar** (`SD-16`). A game's parser is a phrasebook.
- **Every `Declare` answers to `EN.8`.**
- **The joy palette stays an easter egg.** Bob Ross Inc. holds registered
  marks on the painter's name and likeness, and says it strictly controls
  their use. Pigment names are generic. Keep the palette in `examples/`,
  and out of marketing.

*Added at the second sitting, for section V:*

- **A private word never reaches a shared SOP.** Expansion is visible,
  and the long form is what gets saved.
- **Extensions load from the person's profile, never from the workbook
  that arrived by email** (`SEC.9`, `SEC.10`).
- **A personal rule may not silently shadow a shipped sentence** (`G3`'s
  override rule).
- **Hooks watch. They do not touch.**
- **`SD-16` guards two doors in this file:** Spitball 32's enumerations
  and Spitball 48's rest-of-sentence slot. Both are the owner's to
  adjudicate, long-term first.

*Corrected at the second sitting:* "Seeds are stated" above assumed there
was a seed to state. Nothing in the language draws a random number today
(Spitball 34), and Contemplation 5 only recommends refusing an unseeded
draw. Nothing enforces it yet.

*Added at the third sitting:*

- **No steganography.** A cipher whose rules anyone can read is a toy.
  Hiding data inside a workbook's colours or formats is what an IT
  reviewer fears, and this product sells legibility.
- **A found phrasebook is a costume over one grammar** (`SD-16`). If a
  notation will not fit the pattern language, the spitball is the
  longhand, never a second parser.
- **Hobbies have owners.** Generic rules, pigments and stitches are free.
  A publisher's game, a thread maker's colour numbers and a board game's
  name are not. Word lists and datasets need a clean licence before they
  ship.
- **An export is an external effect** (`SEC.8`). Text written into a cell
  is not.

## If only three get built

1. **Spitballs 1 and 2.** One mechanism, and the story developers retell.
2. **Spitballs 5 and 6.** Cheap, and they make the owner's own concept
   true.
3. **Spitball 14, the assisted relay.** An example file that tests the
   renderer harder than any suite does, and hackers will post it for the
   project.

The afternoon-sized ones: 6, 18, 19, 21, and `yoda.vla` from 13.

*Added at the second sitting:* Spitballs 24 and 27 go to the top of this
list. Half of 24 may already work, it is the cheapest item in the file
for what it unlocks, and everything above that says "pack" is built with
it. The test of it is the owner's own hands: whether they stop typing the
long forms. More afternoon-sized ones: 39, 41, 51, and 46 once 24 exists.

*Added at the third sitting:* from this sitting the first four would be
Spitball 65 (it should run today, and the story tells itself), Spitball
68 (the same, and it names the product), Spitball 52 (the flagship of the
richest seam) and Spitball 58 (a new audience that already writes
instructions for a living room). Before any of them come the top two rows
of the enablers table: a way to read a character, and a word for a
remainder. Between them they unlock a dozen entries.

## Sources for the outside facts

- The xkcd colour survey: [the results post](https://blog.xkcd.com/2010/05/03/color-survey-results/)
  and [the 954 colours](https://xkcd.com/color/rgb/), whose `rgb.txt` is
  released under CC0 1.0.
- Bob Ross Inc.: [its own page on use of the Bob Ross intellectual property](https://experience.bobross.com/hosting-a-bob-ross-event/)
  and [its registrations as listed by Justia](https://trademark.justia.com/owners/bob-ross-inc-351808).
